#!/usr/bin/env python3
"""骨子（deck.json）を社内テンプレのレイアウトに流し込んで .pptx を作る。

    python build_deck.py decks/q3/deck.json --out decks/q3/deck.pptx [--template path]

方針: テンプレを開き、既存スライドを外し、レイアウトからスライドを起こしてプレースホルダーに流し込む。
マスター・テーマ・レイアウトはテンプレのまま残るので、PowerPoint でそのまま編集できる。
図表は PowerPoint ネイティブの表・グラフで作る。色はテーマ色（accent1 など）で指定する。

仕様は SKILL のディレクトリにある deck-spec.md を参照。
"""

from __future__ import annotations

import argparse
import copy
import json
import re
import sys
from pathlib import Path

from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.shapes.placeholder import ChartPlaceholder, TablePlaceholder
from pptx.util import Pt

sys.path.insert(0, str(Path(__file__).resolve().parent))
from lib.pptx_common import layout_placeholder_keys, open_presentation, remove_all_slides  # noqa: E402


class SpecError(Exception):
    pass


THEME_COLORS = {
    "text1": MSO_THEME_COLOR.TEXT_1, "dk1": MSO_THEME_COLOR.TEXT_1,
    "text2": MSO_THEME_COLOR.TEXT_2, "dk2": MSO_THEME_COLOR.TEXT_2,
    "bg1": MSO_THEME_COLOR.BACKGROUND_1, "lt1": MSO_THEME_COLOR.BACKGROUND_1,
    "bg2": MSO_THEME_COLOR.BACKGROUND_2, "lt2": MSO_THEME_COLOR.BACKGROUND_2,
    **{f"accent{i}": getattr(MSO_THEME_COLOR, f"ACCENT_{i}") for i in range(1, 7)},
}

CHART_TYPES = {
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "stacked_column": XL_CHART_TYPE.COLUMN_STACKED,
    "stacked_column_100": XL_CHART_TYPE.COLUMN_STACKED_100,
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "stacked_bar": XL_CHART_TYPE.BAR_STACKED,
    "stacked_bar_100": XL_CHART_TYPE.BAR_STACKED_100,
    "line": XL_CHART_TYPE.LINE_MARKERS,
    "line_plain": XL_CHART_TYPE.LINE,
    "area": XL_CHART_TYPE.AREA,
    "stacked_area": XL_CHART_TYPE.AREA_STACKED,
    "pie": XL_CHART_TYPE.PIE,
    "doughnut": XL_CHART_TYPE.DOUGHNUT,
}

SHAPES = {
    "rect": MSO_SHAPE.RECTANGLE,
    "rounded": MSO_SHAPE.ROUNDED_RECTANGLE,
    "oval": MSO_SHAPE.OVAL,
    "chevron": MSO_SHAPE.CHEVRON,
    "pentagon": MSO_SHAPE.PENTAGON,
    "arrow_right": MSO_SHAPE.RIGHT_ARROW,
    "arrow_down": MSO_SHAPE.DOWN_ARROW,
    "triangle": MSO_SHAPE.ISOSCELES_TRIANGLE,
}

NARROW_TEXT_SHAPES = {"chevron", "pentagon", "arrow_right", "arrow_down", "triangle"}

ALIGN = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT, "justify": PP_ALIGN.JUSTIFY}
VALIGN = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}
LEGEND = {"bottom": XL_LEGEND_POSITION.BOTTOM, "right": XL_LEGEND_POSITION.RIGHT, "top": XL_LEGEND_POSITION.TOP, "left": XL_LEGEND_POSITION.LEFT}

BOLD_RE = re.compile(r"(\*\*.+?\*\*)")


# ---------------------------------------------------------------- 色・テキスト

def apply_color(color_format, spec: str) -> None:
    """'accent1' / 'accent1@0.6'（明るさ -1〜1）/ '#RRGGBB' を適用する。"""
    name, _, bright = spec.partition("@")
    if name.startswith("#"):
        color_format.rgb = RGBColor.from_string(name[1:].upper())
    elif name in THEME_COLORS:
        color_format.theme_color = THEME_COLORS[name]
    else:
        raise SpecError(f"色 '{spec}' は不明。accent1〜6 / text1 / text2 / bg1 / bg2 / #RRGGBB を使う")
    if bright:
        color_format.brightness = float(bright)


def add_runs(paragraph, text: str, opts: dict) -> None:
    for part in BOLD_RE.split(text):
        if not part:
            continue
        run = paragraph.add_run()
        if part.startswith("**") and part.endswith("**"):
            run.text = part[2:-2]
            run.font.bold = True
        else:
            run.text = part
        if opts.get("bold") is not None and not (part.startswith("**") and part.endswith("**")):
            run.font.bold = opts["bold"]
        if opts.get("size"):
            run.font.size = Pt(opts["size"])
        if opts.get("color"):
            apply_color(run.font.color, opts["color"])


def flatten_items(value, level: int = 0) -> list[dict]:
    """str / list（入れ子でレベル）/ dict を段落のリストにする。"""
    items: list[dict] = []
    if isinstance(value, str):
        for line in value.split("\n"):
            items.append({"text": line, "level": level})
    elif isinstance(value, list):
        for v in value:
            if isinstance(v, list):
                items.extend(flatten_items(v, level + 1))
            elif isinstance(v, dict):
                items.append({"level": level, **v})
            else:
                items.append({"text": str(v), "level": level})
    elif isinstance(value, dict) and "text" in value:
        base = {k: v for k, v in value.items() if k != "text"}
        for it in flatten_items(value["text"], level):
            items.append({**base, **it})
    else:
        raise SpecError(f"テキストとして解釈できない値: {value!r}")
    return items


def write_text(text_frame, value, opts: dict | None = None) -> None:
    opts = dict(opts or {})
    if isinstance(value, dict):
        opts.update({k: v for k, v in value.items() if k in ("size", "bold", "color", "align", "valign")})
    items = flatten_items(value)
    text_frame.clear()  # 最初の段落は残る
    if opts.get("valign"):
        text_frame.vertical_anchor = VALIGN[opts["valign"]]
    for i, it in enumerate(items):
        para = text_frame.paragraphs[0] if i == 0 else text_frame.add_paragraph()
        para.level = int(it.get("level", 0))
        align = it.get("align", opts.get("align"))
        if align:
            para.alignment = ALIGN[align]
        run_opts = {k: it.get(k, opts.get(k)) for k in ("size", "bold", "color")}
        add_runs(para, str(it.get("text", "")), run_opts)


# ---------------------------------------------------------------- 図表

def resolve_box(prs, layout_keys: dict, item: dict) -> tuple[int, int, int, int]:
    sw, sh = int(prs.slide_width), int(prs.slide_height)
    if "box" in item:
        x, y, w, h = item["box"]
        return int(sw * x / 100), int(sh * y / 100), int(sw * w / 100), int(sh * h / 100)
    if "area" in item:
        key = item["area"]
        if key not in layout_keys:
            raise SpecError(f"area '{key}' はこのレイアウトに無い。使えるキー: {', '.join(layout_keys)}")
        return layout_keys[key]["box"]
    raise SpecError(f"位置が無い（box か area が必要）: {item}")


def add_table(shapes, box, spec: dict, placeholder=None):
    columns = spec.get("columns") or []
    rows = spec.get("rows") or []
    n_rows = len(rows) + (1 if columns else 0)
    n_cols = len(columns) if columns else max(len(r) for r in rows)
    if placeholder is not None and hasattr(placeholder, "insert_table"):
        frame = placeholder.insert_table(n_rows, n_cols)
    else:
        x, y, w, h = box
        frame = shapes.add_table(n_rows, n_cols, x, y, w, h)
    table = frame.table
    table.first_row = bool(columns)
    table.first_col = bool(spec.get("first_col", False))
    if spec.get("col_widths"):
        ratios = spec["col_widths"]
        total_w = sum(c.width for c in table.columns)
        for col, r in zip(table.columns, ratios):
            col.width = int(total_w * r / sum(ratios))
    size = spec.get("font_size", 12)
    data = ([columns] if columns else []) + rows
    aligns = spec.get("align") or []
    for r_i, row in enumerate(data):
        for c_i in range(n_cols):
            cell = table.cell(r_i, c_i)
            value = row[c_i] if c_i < len(row) else ""
            opts = {"size": size}
            if c_i < len(aligns) and aligns[c_i]:
                opts["align"] = aligns[c_i]
            write_text(cell.text_frame, "" if value is None else value if isinstance(value, (dict, list)) else str(value), opts)
    return frame


def add_chart(shapes, box, spec: dict, placeholder=None):
    ctype = spec.get("type", "column")
    if ctype not in CHART_TYPES:
        raise SpecError(f"グラフ種別 '{ctype}' は未対応。使える種別: {', '.join(CHART_TYPES)}")
    data = CategoryChartData()
    data.categories = spec["categories"]
    for s in spec["series"]:
        data.add_series(s.get("name", ""), s["values"], number_format=spec.get("number_format"))
    xl_type = CHART_TYPES[ctype]
    if placeholder is not None and hasattr(placeholder, "insert_chart"):
        frame = placeholder.insert_chart(xl_type, data)
    else:
        x, y, w, h = box
        frame = shapes.add_chart(xl_type, x, y, w, h, data)
    chart = frame.chart
    chart.font.size = Pt(spec.get("font_size", 12))
    title = spec.get("title")
    chart.has_title = bool(title)
    if title:
        chart.chart_title.text_frame.text = title
    legend = spec.get("legend", "bottom" if len(spec["series"]) > 1 or ctype in ("pie", "doughnut") else False)
    chart.has_legend = bool(legend)
    if legend:
        chart.legend.position = LEGEND[legend]
        chart.legend.include_in_layout = False
    colors = spec.get("colors") or [f"accent{i}" for i in range(1, 7)]
    plot = chart.plots[0]
    if ctype in ("pie", "doughnut"):
        for i, point in enumerate(plot.series[0].points):
            point.format.fill.solid()
            apply_color(point.format.fill.fore_color, colors[i % len(colors)])
    else:
        for i, series in enumerate(plot.series):
            fmt = series.format
            if ctype.startswith("line"):
                fmt.line.width = Pt(2.25)
                apply_color(fmt.line.color, colors[i % len(colors)])
            else:
                fmt.fill.solid()
                apply_color(fmt.fill.fore_color, colors[i % len(colors)])
        if ctype.startswith("bar") or ctype.startswith("stacked_bar"):
            # 横棒は既定で下から並ぶため、上から読める順にそろえる
            chart.category_axis.reverse_order = True
        try:
            chart.value_axis.has_major_gridlines = spec.get("gridlines", True)
            chart.value_axis.major_gridlines.format.line.color.rgb = RGBColor(0xD9, 0xD9, 0xD9)
        except (ValueError, AttributeError):
            pass
    if spec.get("data_labels", True):
        plot.has_data_labels = True
        labels = plot.data_labels
        labels.font.size = Pt(spec.get("label_size", spec.get("font_size", 12) - 1))
        if spec.get("number_format"):
            labels.number_format = spec["number_format"]
            labels.number_format_is_linked = False
        if ctype in ("column", "bar", "line", "line_plain"):
            labels.position = XL_LABEL_POSITION.OUTSIDE_END if ctype in ("column", "bar") else XL_LABEL_POSITION.ABOVE
    return frame


def add_image(shapes, box, path: Path, placeholder=None):
    if not path.exists():
        raise SpecError(f"画像が無い: {path}")
    if placeholder is not None and hasattr(placeholder, "insert_picture"):
        return placeholder.insert_picture(str(path))
    x, y, w, h = box
    pic = shapes.add_picture(str(path), x, y)
    # 枠に収まるよう縦横比を保って縮小し、中央に置く
    scale = min(w / pic.width, h / pic.height)
    pic.width, pic.height = int(pic.width * scale), int(pic.height * scale)
    pic.left, pic.top = x + (w - pic.width) // 2, y + (h - pic.height) // 2
    return pic


def add_free_shape(shapes, prs, layout_keys, item: dict, base_dir: Path):
    kind = item.get("type", "shape")
    if kind == "grid":
        return add_grid(shapes, prs, layout_keys, item, base_dir)
    if kind == "line":
        sw, sh = int(prs.slide_width), int(prs.slide_height)
        (x1, y1), (x2, y2) = item["from"], item["to"]
        line = shapes.add_connector(MSO_CONNECTOR.STRAIGHT, int(sw * x1 / 100), int(sh * y1 / 100), int(sw * x2 / 100), int(sh * y2 / 100))
        apply_color(line.line.color, item.get("color", "text2"))
        line.line.width = Pt(item.get("width", 1.5))
        return line
    box = resolve_box(prs, layout_keys, item)
    if kind == "table":
        return add_table(shapes, box, item)
    if kind == "chart":
        return add_chart(shapes, box, item)
    if kind == "image":
        return add_image(shapes, box, (base_dir / item["path"]).resolve())
    if kind == "text":
        tb = shapes.add_textbox(*box)
        tb.text_frame.word_wrap = True
        write_text(tb.text_frame, item.get("text", ""), item)
        return tb
    if kind == "shape":
        name = item.get("shape", "rect")
        if name not in SHAPES:
            raise SpecError(f"図形 '{name}' は未対応。使える図形: {', '.join(SHAPES)}")
        shp = shapes.add_shape(SHAPES[name], *box)
        shp.shadow.inherit = False
        fill = item.get("fill", "accent1")
        if fill == "none":
            shp.fill.background()
        else:
            shp.fill.solid()
            apply_color(shp.fill.fore_color, fill)
        line = item.get("line", "none")
        if line == "none":
            shp.line.fill.background()
        else:
            apply_color(shp.line.color, line)
        if item.get("text") is not None:
            # 矢羽根・矢印は内側の文字領域が狭く、折り返すと 1 文字ずつ縦に並ぶ。改行は \n で明示させる
            shp.text_frame.word_wrap = name not in NARROW_TEXT_SHAPES
            opts = {"color": item.get("text_color", "bg1"), "align": item.get("align", "center"), "valign": item.get("valign", "middle")}
            for k in ("size", "bold"):
                if k in item:
                    opts[k] = item[k]
            write_text(shp.text_frame, item["text"], opts)
        return shp
    raise SpecError(f"add の type '{kind}' は未対応（text / shape / line / table / chart / image / grid）")


def add_grid(shapes, prs, layout_keys, item: dict, base_dir: Path):
    """area / box を cols × rows に等分し、items を順に置く（カード・手順フロー用）。"""
    x, y, w, h = resolve_box(prs, layout_keys, item)
    cols, rows = int(item.get("cols", len(item["items"]))), int(item.get("rows", 1))
    gap = int(prs.slide_width * item.get("gap", 2) / 100)
    cw = (w - gap * (cols - 1)) // cols
    ch = (h - gap * (rows - 1)) // rows
    sw, sh = int(prs.slide_width), int(prs.slide_height)
    made = []
    for i, child in enumerate(item["items"]):
        r, c = divmod(i, cols)
        if r >= rows:
            raise SpecError(f"grid の items が {cols}×{rows} を超えている")
        cx, cy = x + c * (cw + gap), y + r * (ch + gap)
        child = {**child, "box": [cx / sw * 100, cy / sh * 100, cw / sw * 100, ch / sh * 100]}
        child.pop("area", None)
        made.append(add_free_shape(shapes, prs, layout_keys, child, base_dir))
    return made


# ---------------------------------------------------------------- レイアウト・プレースホルダー

def all_layouts(prs) -> list[tuple[str, object]]:
    out = []
    for master in prs.slide_masters:
        for layout in master.slide_layouts:
            out.append((f"L{len(out) + 1:02d}", layout))
    return out


def find_layout(layouts, name: str):
    for lid, layout in layouts:
        if name == lid or name == layout.name:
            return layout
    for lid, layout in layouts:
        if name.strip().lower() == layout.name.strip().lower():
            return layout
    raise SpecError(f"レイアウト '{name}' が無い。使えるレイアウト: " + ", ".join(f"{lid} {l.name}" for lid, l in layouts))


def as_kind(ph, cls):
    if isinstance(ph, cls):
        return ph
    if ph.placeholder_format.type in (PP_PLACEHOLDER.OBJECT, PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.CHART, PP_PLACEHOLDER.TABLE):
        return cls(ph._element, ph._parent)
    # 画像枠などに表・グラフを指定された場合は枠を外し、同じ位置に普通の図形として置く
    ph._element.getparent().remove(ph._element)
    return None


def copy_footer_placeholder(slide, layout, ph_type, text: str | None = None) -> bool:
    for lph in layout.placeholders:
        if lph.placeholder_format.type == ph_type:
            el = copy.deepcopy(lph._element)
            # 図形 ID はスライド内で一意でないと PowerPoint が修復を求めるため振り直す
            ids = [int(v) for v in slide.shapes._spTree.xpath("//p:cNvPr/@id")]
            el.xpath("./*[1]/p:cNvPr")[0].set("id", str(max(ids + [1]) + 1))
            slide.shapes._spTree.append(el)
            if text is not None:
                new = slide.placeholders[lph.placeholder_format.idx]
                write_text(new.text_frame, text)
            return True
    return False


def build(spec: dict, template: Path, base_dir: Path):
    prs = open_presentation(template)
    if not spec.get("keep_template_slides", False):
        remove_all_slides(prs)
    layouts = all_layouts(prs)
    report = []
    for s_no, s in enumerate(spec["slides"], start=1):
        try:
            layout = find_layout(layouts, s["layout"])
            slide = prs.slides.add_slide(layout)
            keyed = layout_placeholder_keys(layout)
            by_key = {it["key"]: it for it in keyed}
            filled_idx = set()
            for key, value in (s.get("fill") or {}).items():
                if key.startswith("#"):
                    idx = int(key[1:])
                else:
                    match = by_key.get(key) or next((it for it in keyed if it["name"] == key), None)
                    if match is None:
                        raise SpecError(f"キー '{key}' はレイアウト '{layout.name}' に無い。使えるキー: {', '.join(k for k, it in by_key.items() if not it['footer'])}")
                    idx = match["idx"]
                try:
                    ph = slide.placeholders[idx]
                except KeyError as exc:
                    raise SpecError(f"idx {idx} のプレースホルダーがスライドに無い") from exc
                box = (int(ph.left), int(ph.top), int(ph.width), int(ph.height))
                filled_idx.add(idx)
                if isinstance(value, dict) and "table" in value:
                    # 本文（OBJECT）枠にも表を入れられるよう TablePlaceholder として扱う
                    add_table(slide.shapes, box, value["table"], as_kind(ph, TablePlaceholder))
                elif isinstance(value, dict) and "chart" in value:
                    add_chart(slide.shapes, box, value["chart"], as_kind(ph, ChartPlaceholder))
                elif isinstance(value, dict) and "image" in value:
                    if hasattr(ph, "insert_picture"):
                        add_image(slide.shapes, box, (base_dir / value["image"]).resolve(), ph)
                    else:
                        # 画像枠以外は枠を外し、同じ位置に縦横比を保って置く
                        ph._element.getparent().remove(ph._element)
                        add_image(slide.shapes, box, (base_dir / value["image"]).resolve())
                else:
                    write_text(ph.text_frame, value)
            # insert_* 系は元の要素を置き換えるので、ここで改めて空き枠を集める
            removed = []
            if not s.get("keep_empty", False):
                for ph in list(slide.placeholders):
                    idx = ph.placeholder_format.idx
                    if idx in filled_idx:
                        continue
                    if ph.has_text_frame and ph.text_frame.text.strip():
                        continue
                    removed.append(ph.name)
                    ph._element.getparent().remove(ph._element)
            for item in s.get("add") or []:
                add_free_shape(slide.shapes, prs, by_key, item, base_dir)
            page_numbers = s.get("page_number", spec.get("page_numbers", True))
            if page_numbers:
                copy_footer_placeholder(slide, layout, PP_PLACEHOLDER.SLIDE_NUMBER)
            footer = s.get("footer", spec.get("footer"))
            if footer:
                copy_footer_placeholder(slide, layout, PP_PLACEHOLDER.FOOTER, footer)
            if s.get("notes"):
                slide.notes_slide.notes_text_frame.text = s["notes"]
            report.append({"no": s_no, "layout": layout.name, "filled": sorted(filled_idx), "removed_empty": removed})
        except SpecError as exc:
            raise SpecError(f"スライド {s_no}: {exc}") from exc
    if spec.get("title"):
        prs.core_properties.title = spec["title"]
    return prs, report


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("spec", help="deck.json")
    ap.add_argument("--out", required=True)
    ap.add_argument("--template", help="spec の template より優先")
    args = ap.parse_args()

    spec_path = Path(args.spec)
    spec = json.loads(spec_path.read_text(encoding="utf-8"))
    template = Path(args.template or spec.get("template") or "")
    if not template.is_file():
        # spec からの相対パスも試す
        alt = (spec_path.parent / template).resolve()
        if alt.is_file():
            template = alt
        else:
            sys.exit(f"テンプレが見つからない: {template}")
    try:
        prs, report = build(spec, template, spec_path.parent)
    except SpecError as exc:
        sys.exit(f"[spec エラー] {exc}")
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out)
    print(f"written: {out}  slides: {len(report)}")
    for r in report:
        extra = f"  空き枠を削除: {', '.join(r['removed_empty'])}" if r["removed_empty"] else ""
        print(f"  {r['no']:>2}. {r['layout']}{extra}")


if __name__ == "__main__":
    main()
