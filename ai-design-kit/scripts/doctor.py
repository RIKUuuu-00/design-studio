#!/usr/bin/env python3
"""Phase 0 の環境チェック。deck-lab / design-lab を使う前に 1 回実行する。

    python doctor.py [--template design-system/slides/template.pptx] [--image-test out.png]

確認項目:
- Python ライブラリ（python-pptx, Pillow, 任意で PyMuPDF）
- LibreOffice（soffice）と PDF→PNG 手段（pdftoppm / PyMuPDF）
- 日本語フォント、テンプレのテーマフォントが入っているか
- Node と Playwright（UI トラック用）
- --image-test: 画像入力の疎通確認用 PNG を作る。Claude に Read させ、書かれた合言葉を答えられれば
  ゲートウェイ経由の画像入力が通っている
"""

from __future__ import annotations

import argparse
import importlib
import random
import shutil
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

OK, NG, WARN = "OK  ", "NG  ", "WARN"
results: list[tuple[str, str, str]] = []


def add(status: str, item: str, detail: str) -> None:
    results.append((status, item, detail))


def check_python() -> None:
    for mod, pip, required in (("pptx", "python-pptx", True), ("PIL", "pillow", True), ("fitz", "pymupdf", False), ("lxml", "lxml", True)):
        try:
            m = importlib.import_module(mod)
            add(OK, f"python: {pip}", getattr(m, "__version__", "") or "導入済み")
        except ImportError:
            add(NG if required else WARN, f"python: {pip}", f"未導入 → pip install {pip}" + ("" if required else "（pdftoppm が無い場合に必要）"))


def check_office() -> None:
    from lib.pptx_common import find_soffice

    soffice = find_soffice()
    if soffice:
        try:
            ver = subprocess.run([soffice, "--version"], capture_output=True, text=True, timeout=60).stdout.strip()
        except Exception:  # noqa: BLE001
            ver = soffice
        add(OK, "LibreOffice", ver or soffice)
    else:
        add(NG, "LibreOffice", "未導入。導入不可なら PowerPoint で PDF を書き出し、render_pptx.py に .pdf を渡す代替運用にする")
    if shutil.which("pdftoppm"):
        add(OK, "PDF→PNG", "pdftoppm")
    else:
        try:
            importlib.import_module("fitz")
            add(OK, "PDF→PNG", "PyMuPDF")
        except ImportError:
            add(NG, "PDF→PNG", "pdftoppm（poppler）か PyMuPDF（pip install pymupdf）が必要")


def check_fonts(template: str | None) -> None:
    from lib.pptx_common import installed_font_families, missing_theme_fonts, open_presentation, theme_info

    fams = installed_font_families()
    if fams is None:
        add(WARN, "フォント", "fc-list が無く確認できない（Windows では設定 > フォント で目視確認）")
        return
    jp = [f for f in fams if any(k in f for k in ("gothic", "ゴシック", "mincho", "明朝", "noto sans cjk", "noto sans jp", "meiryo", "メイリオ", "yu ", "游", "hiragino", "ヒラギノ", "biz ud", "cjk jp", "noto serif jp"))]
    names = sorted({f for f in jp if not f.startswith(".")}, key=str.lower)
    add(OK if jp else NG, "日本語フォント", (f"{len(names)} 種類: " + ", ".join(names[:12]) + (" ほか" if len(names) > 12 else "")) if jp else "見つからない。テンプレと同じフォントを導入する")
    if template:
        prs = open_presentation(template)
        theme = theme_info(prs)
        missing = missing_theme_fonts(theme)
        fonts = sorted({f for k in theme["fonts"].values() for f in k.values() if f and not f.startswith("+")})
        if missing:
            add(NG, "テーマフォント", f"{', '.join(missing)} が無い。画像化の文字幅が実物とずれる。導入できない場合は check_deck.py --margin 0.1 で 1 割の余裕を残す運用にする")
        else:
            add(OK, "テーマフォント", ", ".join(fonts) or "（テーマにフォント指定なし）")


def check_node() -> None:
    node = shutil.which("node")
    if not node:
        add(WARN, "Node", "未導入（UI トラックの export.mjs / token-lint.mjs に必要）")
        return
    ver = subprocess.run([node, "--version"], capture_output=True, text=True).stdout.strip()
    add(OK, "Node", ver)
    probe = (
        "const {createRequire}=require('module');const {execSync}=require('child_process');"
        "const c=[process.cwd()+'/x.js'];try{c.push(execSync('npm root -g').toString().trim()+'/x.js')}catch(e){}"
        "for(const p of c){try{createRequire(p)('playwright');console.log('ok');process.exit(0)}catch(e){}}process.exit(1)"
    )
    r = subprocess.run([node, "-e", probe], capture_output=True, text=True)
    if r.returncode == 0:
        add(OK, "Playwright", "導入済み（ブラウザ取得が社内プロキシで止まる場合は PLAYWRIGHT_CHANNEL=chrome|msedge）")
    else:
        add(WARN, "Playwright", "未導入 → npm i -D playwright（UI トラックで必要）")


def image_test(target: Path) -> str:
    from PIL import Image, ImageDraw, ImageFont

    word = random.choice(["さくら", "みなと", "こはく", "あおば", "ひかり"]) + str(random.randint(10, 99))
    im = Image.new("RGB", (640, 240), "white")
    d = ImageDraw.Draw(im)
    font = None
    for cand in ("/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf", "C:/Windows/Fonts/meiryo.ttc", "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc"):
        if Path(cand).exists():
            font = ImageFont.truetype(cand, 56)
            break
    d.rectangle([10, 10, 630, 230], outline=(31, 79, 209), width=6)
    d.text((40, 80), f"合言葉: {word}" if font else f"WORD: {word}", fill=(20, 20, 20), font=font)
    im.save(target)
    return word


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--template")
    ap.add_argument("--image-test", help="画像入力テスト用 PNG の出力先")
    args = ap.parse_args()

    add(OK, "Python", sys.version.split()[0])
    check_python()
    check_office()
    try:
        check_fonts(args.template)
    except ImportError:
        pass
    check_node()

    width = max(len(r[1]) for r in results)
    for status, item, detail in results:
        print(f"[{status}] {item.ljust(width)}  {detail}")
    if args.image_test:
        word = image_test(Path(args.image_test))
        print(f"\n画像入力テスト: {args.image_test} を Claude に Read させ、合言葉が「{word}」と読めれば OK")
    if any(r[0] == NG for r in results):
        sys.exit(1)


if __name__ == "__main__":
    main()
