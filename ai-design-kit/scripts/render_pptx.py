#!/usr/bin/env python3
"""pptx（または PowerPoint で書き出した PDF）をスライド画像にする。

    python render_pptx.py decks/q3/deck.pptx --out decks/q3/renders [--dpi 110] [--sheet]

- .pptx/.potx: LibreOffice（headless）で PDF 化し、PDF を PNG に分割する。PDF は入力と同じ場所に置く
- .pdf: そのまま PNG に分割する（LibreOffice を入れられない環境の代替運用。PDF 書き出しだけ人が行う）
- PNG 化は pdftoppm → PyMuPDF の順に使えるものを使う
- --sheet: 全スライドのサムネ一覧（contact sheet）も作る
"""

from __future__ import annotations

import argparse
import glob
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from lib.pptx_common import find_soffice  # noqa: E402


def pptx_to_pdf(src: Path, timeout: int) -> Path:
    soffice = find_soffice()
    if not soffice:
        sys.exit(
            "LibreOffice（soffice）が見つからない。導入するか、PowerPoint で PDF に書き出して"
            " その .pdf を渡すこと（代替運用）。"
        )
    out_dir = src.parent
    # 既存の LibreOffice プロファイルとのロック競合を避けるため、毎回使い捨てのプロファイルを使う
    profile = Path(tempfile.mkdtemp(prefix="decklab-lo-"))
    cmd = [
        soffice,
        f"-env:UserInstallation={profile.as_uri()}",
        "--headless",
        "--norestore",
        "--convert-to",
        "pdf",
        "--outdir",
        str(out_dir),
        str(src),
    ]
    try:
        subprocess.run(cmd, check=True, capture_output=True, timeout=timeout)
    except subprocess.TimeoutExpired:
        sys.exit(f"LibreOffice が {timeout} 秒で終わらなかった。--timeout を伸ばすか、PDF を手動で書き出すこと。")
    except subprocess.CalledProcessError as exc:
        sys.exit(f"LibreOffice の変換に失敗: {exc.stderr.decode(errors='replace')[:500]}")
    finally:
        shutil.rmtree(profile, ignore_errors=True)
    pdf = out_dir / (src.stem + ".pdf")
    if not pdf.exists():
        sys.exit(f"PDF が生成されなかった: {pdf}")
    return pdf


def pdf_to_pngs(pdf: Path, out_dir: Path, dpi: int) -> list[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    for old in glob.glob(str(out_dir / "slide-*.png")):
        os.remove(old)
    if shutil.which("pdftoppm"):
        subprocess.run(
            ["pdftoppm", "-r", str(dpi), "-png", str(pdf), str(out_dir / "slide")],
            check=True,
            capture_output=True,
        )
        files = sorted(out_dir.glob("slide-*.png"))
        # pdftoppm はページ数に応じて桁数が変わるので 2 桁にそろえる
        renamed = []
        for f in files:
            n = int(f.stem.split("-")[-1])
            target = out_dir / f"slide-{n:02d}.png"
            if f != target:
                f.rename(target)
            renamed.append(target)
        return sorted(renamed)
    try:
        import fitz  # PyMuPDF
    except ImportError:
        sys.exit("pdftoppm（poppler）も PyMuPDF も無い。どちらかを導入すること: pip install pymupdf")
    doc = fitz.open(str(pdf))
    files = []
    for i, page in enumerate(doc, start=1):
        target = out_dir / f"slide-{i:02d}.png"
        page.get_pixmap(dpi=dpi).save(str(target))
        files.append(target)
    return files


def contact_sheet(files: list[Path], target: Path, cols: int = 4, thumb_w: int = 480) -> Path | None:
    try:
        from PIL import Image, ImageDraw
    except ImportError:
        print("Pillow が無いのでサムネ一覧は作らない（pip install pillow）", file=sys.stderr)
        return None
    if not files:
        return None
    thumbs = []
    for f in files:
        im = Image.open(f).convert("RGB")
        ratio = thumb_w / im.width
        thumbs.append(im.resize((thumb_w, int(im.height * ratio))))
    th = max(t.height for t in thumbs)
    label_h, pad = 28, 16
    rows = (len(thumbs) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * (thumb_w + pad) + pad, rows * (th + label_h + pad) + pad), "white")
    draw = ImageDraw.Draw(sheet)
    for i, t in enumerate(thumbs):
        x = pad + (i % cols) * (thumb_w + pad)
        y = pad + (i // cols) * (th + label_h + pad)
        draw.text((x, y), f"{i + 1}  {files[i].name}", fill=(60, 60, 60))
        sheet.paste(t, (x, y + label_h))
        draw.rectangle([x - 1, y + label_h - 1, x + t.width, y + label_h + t.height], outline=(200, 200, 200))
    sheet.save(target)
    return target


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("src", help=".pptx / .potx / .pdf")
    ap.add_argument("--out", required=True, help="PNG の出力先ディレクトリ")
    ap.add_argument("--dpi", type=int, default=110)
    ap.add_argument("--timeout", type=int, default=240)
    ap.add_argument("--sheet", action="store_true", help="サムネ一覧 sheet.png も作る")
    args = ap.parse_args()

    src = Path(args.src).resolve()
    if not src.exists():
        sys.exit(f"入力が無い: {src}")
    if src.suffix.lower() == ".pdf":
        pdf = src
    else:
        if src.suffix.lower() == ".potx":
            tmp = Path(tempfile.mkdtemp(prefix="decklab-")) / (src.stem + ".pptx")
            shutil.copy(src, tmp)
            src = tmp
        pdf = pptx_to_pdf(src, args.timeout)
    out_dir = Path(args.out).resolve()
    files = pdf_to_pngs(pdf, out_dir, args.dpi)
    print(f"pdf: {pdf}")
    for f in files:
        print(f"png: {f}")
    if args.sheet:
        sheet = contact_sheet(files, out_dir / "sheet.png")
        if sheet:
            print(f"sheet: {sheet}")


if __name__ == "__main__":
    main()
