#!/usr/bin/env python3
"""pptx を構造面から検査する（画像を見なくても分かる不具合を拾う）。

    python check_deck.py decks/q3/deck.pptx [--json report.json] [--margin 0.1] [--min-font 10] [--strict]

検出するもの:
- placeholder_empty  … 空のまま残ったプレースホルダー（編集画面で「クリックして…」が出る）
- leftover_text      … 見本文・入力促し文の消し忘れ（Lorem ipsum、〇〇、ここに入力 など）
- overflow_risk      … 文字量が枠に収まらない見込み（全角=1em の概算。--margin で余裕率を指定）
- out_of_bounds      … スライド外へのはみ出し
- overlap            … 文字を持つ図形同士の重なり
- small_font         … --min-font 未満の文字
- font_off_theme     … テーマ外のフォント直指定
- color_off_theme    … テーマ外の RGB 直指定（テンプレ準拠の逸脱）
- font_missing       … テーマのフォントが検証環境に無い（画像化結果が実物とずれる）

既定では指摘があっても終了コード 0（直すための材料として使う）。--strict で error があれば 1。
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from lib.pptx_common import (  # noqa: E402
    NS,
    estimate_text_height,
    intersection_area,
    leftover_hits,
    missing_theme_fonts,
    open_presentation,
    placeholder_role,
    shape_box,
    theme_info,
)

SEVERITY = {
    "placeholder_empty": "error",
    "leftover_text": "error",
    "overflow_risk": "error",
    "out_of_bounds": "error",
    "overlap": "warn",
    "small_font": "warn",
    "font_off_theme": "warn",
    "color_off_theme": "info",
    "font_missing": "warn",
    "tight_fit": "info",
}


def finding(kind: str, slide: int | None, shape: str, message: str) -> dict:
    return {"kind": kind, "severity": SEVERITY[kind], "slide": slide, "shape": shape, "message": message}


def iter_shapes(shapes):
    for shp in shapes:
        if shp.shape_type is not None and getattr(shp, "shapes", None) is not None and not shp.is_placeholder:
            # グループ内も見る
            yield from iter_shapes(shp.shapes)
        else:
            yield shp


def has_content(shp) -> bool:
    if getattr(shp, "has_text_frame", False) and shp.text_frame.text.strip():
        return True
    return getattr(shp, "has_table", False) or getattr(shp, "has_chart", False)


def check(path: Path, margin: float, min_font: float, extra_patterns: list[str]) -> list[dict]:
    prs = open_presentation(path)
    sw, sh = int(prs.slide_width), int(prs.slide_height)
    theme = theme_info(prs)
    theme_fonts = {f for kind in theme["fonts"].values() for f in kind.values() if f}
    theme_rgb = {v for v in theme["colors"].values() if v}
    out: list[dict] = []

    for face in missing_theme_fonts(theme):
        out.append(finding("font_missing", None, "-", f"テーマのフォント「{face}」がこの環境に無い。画像の文字幅は実物と異なる"))

    for s_no, slide in enumerate(prs.slides, start=1):
        content_boxes = []
        for shp in iter_shapes(slide.shapes):
            name = shp.name
            box = shape_box(shp)
            # 空プレースホルダー
            if shp.is_placeholder and getattr(shp, "has_text_frame", False) and not shp.text_frame.text.strip():
                if placeholder_role(shp) not in ("slide_number", "date", "footer"):
                    out.append(finding("placeholder_empty", s_no, name, "空のプレースホルダーが残っている。不要なら削除する"))
            # 見本文
            if getattr(shp, "has_text_frame", False):
                hits = leftover_hits(shp.text_frame.text, extra_patterns)
                if hits:
                    out.append(finding("leftover_text", s_no, name, f"見本文・仮文字が残っている: 「{shp.text_frame.text.strip()[:40]}」"))
            if getattr(shp, "has_table", False):
                for r_i, row in enumerate(shp.table.rows):
                    for c_i, cell in enumerate(row.cells):
                        if leftover_hits(cell.text, extra_patterns):
                            out.append(finding("leftover_text", s_no, name, f"表の {r_i + 1} 行 {c_i + 1} 列に仮文字: 「{cell.text[:30]}」"))
            # はみ出し
            if box:
                x, y, w, h = box
                tol = int(min(sw, sh) * 0.005)
                if x < -tol or y < -tol or x + w > sw + tol or y + h > sh + tol:
                    out.append(finding("out_of_bounds", s_no, name, "スライドの外にはみ出している"))
            # 文字量
            est = estimate_text_height(shp, slide) if getattr(shp, "has_text_frame", False) and shp.text_frame.text.strip() else None
            if est:
                if est["ratio"] > 1.0:
                    note = "（自動縮小が掛かる設定。縮小後の文字が小さすぎないか画像で確認）" if est["autofit"] == "shrink" else ""
                    out.append(finding("overflow_risk", s_no, name, f"文字量が枠の約 {est['ratio']:.0%}。削るか枠・レイアウトを変える{note}"))
                elif est["ratio"] > 1.0 - margin:
                    out.append(finding("tight_fit", s_no, name, f"文字量が枠の約 {est['ratio']:.0%}。フォント差で溢れる余地がある"))
            # 文字サイズ・フォント・色
            if getattr(shp, "has_text_frame", False):
                for para in shp.text_frame.paragraphs:
                    for run in para.runs:
                        if run.font.size is not None and run.font.size.pt < min_font and run.text.strip():
                            out.append(finding("small_font", s_no, name, f"{run.font.size.pt:g}pt の文字がある（下限 {min_font:g}pt）"))
                            break
                    for r in para._p.findall(".//a:rPr", NS):
                        for tag in ("latin", "ea"):
                            el = r.find(f"a:{tag}", NS)
                            face = el.get("typeface") if el is not None else None
                            if face and not face.startswith("+") and face not in theme_fonts:
                                out.append(finding("font_off_theme", s_no, name, f"テーマ外のフォント「{face}」を直接指定している"))
            for clr in shp._element.findall(".//a:srgbClr", NS):
                val = (clr.get("val") or "").upper()
                if val and val not in theme_rgb and val not in ("FFFFFF", "000000", "D9D9D9"):
                    out.append(finding("color_off_theme", s_no, name, f"テーマ外の色 #{val} を直接指定している"))
                    break
            if box and has_content(shp):
                content_boxes.append((name, box))
        # 重なり
        for i in range(len(content_boxes)):
            for j in range(i + 1, len(content_boxes)):
                (na, a), (nb, b) = content_boxes[i], content_boxes[j]
                inter = intersection_area(a, b)
                smaller = min(a[2] * a[3], b[2] * b[3]) or 1
                if inter / smaller > 0.05:
                    out.append(finding("overlap", s_no, f"{na} / {nb}", f"内容のある図形が重なっている（小さい方の {inter / smaller:.0%}）"))
    return dedupe(out)


def dedupe(items: list[dict]) -> list[dict]:
    seen = set()
    result = []
    for it in items:
        key = (it["kind"], it["slide"], it["shape"], it["message"])
        if key not in seen:
            seen.add(key)
            result.append(it)
    return result


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("deck")
    ap.add_argument("--json", help="結果を JSON でも書き出す")
    ap.add_argument("--margin", type=float, default=0.1, help="枠に残す余裕率（既定 0.1 = 1割）")
    ap.add_argument("--min-font", type=float, default=10)
    ap.add_argument("--pattern", action="append", default=[], help="仮文字として扱う正規表現を追加")
    ap.add_argument("--strict", action="store_true", help="error があれば終了コード 1")
    args = ap.parse_args()

    results = check(Path(args.deck), args.margin, args.min_font, args.pattern)
    if args.json:
        Path(args.json).write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
    counts = {s: sum(1 for r in results if r["severity"] == s) for s in ("error", "warn", "info")}
    print(f"check: error {counts['error']} / warn {counts['warn']} / info {counts['info']}")
    for r in sorted(results, key=lambda r: (r["slide"] or 0, ("error", "warn", "info").index(r["severity"]))):
        where = f"slide {r['slide']}" if r["slide"] else "deck"
        print(f"  [{r['severity']}] {where} {r['shape']}: {r['kind']} - {r['message']}")
    if args.strict and counts["error"]:
        sys.exit(1)


if __name__ == "__main__":
    main()
