#!/usr/bin/env python3
"""比較キャンバス（index.html）を canvas.json から生成する。両トラック共通。

    python build_canvas.py decks/q3/canvas.json --out decks/q3/index.html
    python build_canvas.py designs/settings/canvas.json --out designs/settings/index.html

canvas.json:
{
  "track": "deck" | "ui",
  "slug": "q3-report",
  "title": "Q3 報告資料",
  "brief": "ブリーフ原文または要約",
  "assumptions": ["不足情報を推測で補った点"],
  "widths": [375, 1280],                 # ui のみ。画面幅の切替候補
  "options": [
    {
      "id": "A", "title": "結論先行", "aim": "狙い（1行）", "tradeoff": "代償（1行）",
      "outline": [{"no": 1, "title": "主張タイトル", "points": ["要点"], "layout": "Title and Content"}],  # deck
      "src": "a.html", "height": 900,     # ui（index.html からの相対パス）
      "images": ["renders/slide-01.png"], # 任意。画像化した結果を並べる
      "notes": ["自己検証で残った指摘"]
    }
  ]
}
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

TEMPLATE = Path(__file__).resolve().parent.parent / "skills" / "design-lab" / "templates" / "canvas.html"
MARKER = "/*__CANVAS_DATA__*/"


def validate(data: dict, base: Path) -> list[str]:
    errors = []
    if data.get("track") not in ("deck", "ui"):
        errors.append("track は 'deck' か 'ui'")
    if not data.get("slug"):
        errors.append("slug が無い")
    opts = data.get("options") or []
    if not opts:
        errors.append("options が空")
    ids = [o.get("id") for o in opts]
    if len(set(ids)) != len(ids) or None in ids:
        errors.append("options の id が重複または欠落している")
    for o in opts:
        if not (o.get("outline") or o.get("src") or o.get("images")):
            errors.append(f"案{o.get('id')}: outline / src / images のいずれかが必要")
        for rel in [o.get("src")] + list(o.get("images") or []):
            if rel and not (base / rel).exists():
                errors.append(f"案{o.get('id')}: ファイルが無い {rel}")
    return errors


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("canvas_json")
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    src = Path(args.canvas_json)
    data = json.loads(src.read_text(encoding="utf-8"))
    out = Path(args.out)
    errors = validate(data, out.parent)
    if errors:
        sys.exit("canvas.json の不備:\n  - " + "\n  - ".join(errors))
    payload = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    html = TEMPLATE.read_text(encoding="utf-8")
    if MARKER not in html:
        sys.exit(f"テンプレにマーカー {MARKER} が無い: {TEMPLATE}")
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(html.replace(MARKER, payload), encoding="utf-8")
    print(f"written: {out}  options: {', '.join(o['id'] for o in data['options'])}")


if __name__ == "__main__":
    main()
