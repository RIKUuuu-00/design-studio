"""スライドに描くための共通部品（色・文字・表・グラフ・図形）。

色はすべてテーマ色（accent1〜6 / text1・2 / bg1・2）で指定し、テンプレのテーマから外れないようにする。
フォントは指定しない（テーマを継承）。見出しだけテーマの見出しフォント（+mj）を使う。
"""

from __future__ import annotations

import re
from pathlib import Path

from lxml import etree
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION, XL_TICK_MARK
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

A = "http://schemas.openxmlformats.org/drawingml/2006/main"


class SpecError(Exception):
    pass


THEME_COLORS = {
    "text1": MSO_THEME_COLOR.TEXT_1, "dk1": MSO_THEME_COLOR.TEXT_1,
    "text2": MSO_THEME_COLOR.TEXT_2, "dk2": MSO_THEME_COLOR.TEXT_2,
    "bg1": MSO_THEME_COLOR.BACKGROUND_1, "lt1": MSO_THEME_COLOR.BACKGROUND_1,
    "bg2": MSO_THEME_COLOR.BACKGROUND_2, "lt2": MSO_THEME_COLOR.BACKGROUND_2,
    **{f"accent{i}": getattr(MSO_THEME_COLOR, f"ACCENT_{i}") for i in range(1, 7)},
}
SCHEME_XML = {"text1": "tx1", "dk1": "tx1", "text2": "tx2", "dk2": "tx2", "bg1": "bg1", "lt1": "bg1", "bg2": "bg2", "lt2": "bg2",
              **{f"accent{i}": f"accent{i}" for i in range(1, 7)}}

# よく使う中間色（テーマ色の明暗で作るので、テンプレが変わっても追従する）
RULE = "text1@0.8"        # 罫線・区切り線
MUTED = "text1@0.4"       # 補足の文字
TINT = "accent1@0.9"      # 強調面のごく薄い色
QUIET = "text1@0.92"      # 背景面

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


# ---------------------------------------------------------------- 文字サイズの段階

class TypeScale:
    """スライドの高さから文字サイズの段階を決める（7.5in = 1.0）。"""

    STEPS = {"xs": 10, "sm": 12, "body": 14, "lead": 16, "h": 18, "xl": 24, "xxl": 32, "display": 44}

    def __init__(self, slide_height_emu: int, base: float | None = None, slide_width_emu: int | None = None):
        k = (slide_height_emu / 914400) / 7.5
        if base:
            k *= base / self.STEPS["body"]
        elif slide_width_emu:
            # 横長（16:9）は 1 行に入る文字が多く、同じ高さでも文字を大きくできる
            k *= max(slide_width_emu / slide_height_emu / (4 / 3), 1.0) ** 0.5
        self.k = k

    def __getitem__(self, name: str) -> float:
        return round(self.STEPS[name] * self.k * 2) / 2


# ---------------------------------------------------------------- 色

def _split(spec: str) -> tuple[str, float | None]:
    name, _, bright = spec.partition("@")
    return name, float(bright) if bright else None


def apply_color(color_format, spec: str) -> None:
    """'accent1' / 'accent1@0.6'（明るさ -1〜1）/ '#RRGGBB' を適用する。"""
    name, bright = _split(spec)
    if name.startswith("#"):
        color_format.rgb = RGBColor.from_string(name[1:].upper())
    elif name in THEME_COLORS:
        color_format.theme_color = THEME_COLORS[name]
    else:
        raise SpecError(f"色 '{spec}' は不明。accent1〜6 / text1 / text2 / bg1 / bg2 / #RRGGBB を使う")
    if bright:
        color_format.brightness = bright


def color_xml(spec: str):
    """<a:solidFill> 要素を作る（表の罫線など python-pptx が扱わない箇所用）。"""
    name, bright = _split(spec)
    fill = etree.SubElement(etree.Element(qn("a:dummy")), qn("a:solidFill"))
    if name.startswith("#"):
        clr = etree.SubElement(fill, qn("a:srgbClr"), val=name[1:].upper())
    elif name in SCHEME_XML:
        clr = etree.SubElement(fill, qn("a:schemeClr"), val=SCHEME_XML[name])
    else:
        raise SpecError(f"色 '{spec}' は不明")
    if bright:
        if bright > 0:
            etree.SubElement(clr, qn("a:lumMod"), val=str(int((1 - bright) * 100000)))
            etree.SubElement(clr, qn("a:lumOff"), val=str(int(bright * 100000)))
        else:
            etree.SubElement(clr, qn("a:lumMod"), val=str(int((1 + bright) * 100000)))
    return fill


def fill_shape(shape, spec: str | None) -> None:
    if spec in (None, "none"):
        shape.fill.background()
    else:
        shape.fill.solid()
        apply_color(shape.fill.fore_color, spec)


def line_shape(shape, spec: str | None, width: float = 1.0) -> None:
    if spec in (None, "none"):
        shape.line.fill.background()
    else:
        apply_color(shape.line.color, spec)
        shape.line.width = Pt(width)


# ---------------------------------------------------------------- 文字

def add_runs(paragraph, text: str, opts: dict) -> None:
    for part in BOLD_RE.split(text):
        if not part:
            continue
        run = paragraph.add_run()
        is_bold = part.startswith("**") and part.endswith("**")
        run.text = part[2:-2] if is_bold else part
        if is_bold:
            run.font.bold = True
        elif opts.get("bold") is not None:
            run.font.bold = opts["bold"]
        if opts.get("size"):
            run.font.size = Pt(opts["size"])
        if opts.get("color"):
            apply_color(run.font.color, opts["color"])
        if opts.get("heading"):
            rpr = run._r.get_or_add_rPr()
            for tag, face in (("a:latin", "+mj-lt"), ("a:ea", "+mj-ea")):
                el = rpr.find(qn(tag))
                if el is None:
                    el = etree.SubElement(rpr, qn(tag))
                el.set("typeface", face)


def flatten_items(value, level: int = 0) -> list[dict]:
    """str / list（入れ子でレベル）/ dict を段落のリストにする。"""
    items: list[dict] = []
    if isinstance(value, str):
        for line in value.split("\n"):
            items.append({"text": line, "level": level})
    elif isinstance(value, (int, float)):
        items.append({"text": str(value), "level": level})
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


def set_bullet(paragraph, indent_emu: int, char: str = "•") -> None:
    ppr = paragraph._p.get_or_add_pPr()
    ppr.set("marL", str(indent_emu))
    ppr.set("indent", str(-indent_emu))
    for tag in ("a:buNone", "a:buChar", "a:buAutoNum", "a:buFont"):
        for el in ppr.findall(qn(tag)):
            ppr.remove(el)
    font = etree.SubElement(ppr, qn("a:buFont"), typeface="Arial")
    bu = etree.SubElement(ppr, qn("a:buChar"), char=char)
    # 子要素の順序（buFont → buChar）は pPr 内で最後尾でよい
    del font, bu


def set_spacing(paragraph, before_pt: float | None = None, line: float | None = None) -> None:
    ppr = paragraph._p.get_or_add_pPr()
    if line:
        ln = etree.Element(qn("a:lnSpc"))
        etree.SubElement(ln, qn("a:spcPct"), val=str(int(line * 100000)))
        ppr.insert(0, ln)
    if before_pt:
        sb = etree.Element(qn("a:spcBef"))
        etree.SubElement(sb, qn("a:spcPts"), val=str(int(before_pt * 100)))
        ppr.insert(1 if line else 0, sb)


def write_text(text_frame, value, opts: dict | None = None) -> None:
    opts = dict(opts or {})
    if isinstance(value, dict):
        opts.update({k: v for k, v in value.items() if k in ("size", "bold", "color", "align", "valign", "heading")})
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
        if opts.get("bullets") and str(it.get("text", "")).strip():
            size = it.get("size", opts.get("size")) or 14
            set_bullet(para, int(Pt(size) * (1.0 + 1.2 * para.level)), "•" if para.level == 0 else "–")
        if opts.get("para_space") and i > 0:
            set_spacing(para, before_pt=opts["para_space"])
        run_opts = {k: it.get(k, opts.get(k)) for k in ("size", "bold", "color", "heading")}
        add_runs(para, str(it.get("text", "")), run_opts)


def text_box(shapes, box, value, **opts):
    """内部余白 0・折り返しありのテキストボックス。位置合わせを正確にするため余白を消す。"""
    x, y, w, h = (int(v) for v in box)
    tb = shapes.add_textbox(Emu(x), Emu(y), Emu(max(w, 1)), Emu(max(h, 1)))
    tf = tb.text_frame
    tf.word_wrap = opts.pop("wrap", True)
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = VALIGN[opts.pop("valign", "top")]
    name = opts.pop("name", None)
    write_text(tf, value, opts)
    if name:
        tb.name = name
    return tb


def strip_style(shape) -> None:
    """図形のテーマ既定スタイル（影・効果）を外す。線や面の色は明示的に指定する前提。"""
    style = shape._element.find(qn("p:style"))
    if style is not None:
        shape._element.remove(style)


def rect(shapes, box, fill: str | None = None, line: str | None = None, line_w: float = 1.0, shape: str = "rect", name: str | None = None):
    x, y, w, h = (int(v) for v in box)
    shp = shapes.add_shape(SHAPES[shape], Emu(x), Emu(y), Emu(max(w, 1)), Emu(max(h, 1)))
    strip_style(shp)
    fill_shape(shp, fill)
    line_shape(shp, line, line_w)
    if name:
        shp.name = name
    return shp


def line(shapes, x1, y1, x2, y2, color: str = RULE, width: float = 0.75, name: str | None = None):
    ln = shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Emu(int(x1)), Emu(int(y1)), Emu(int(x2)), Emu(int(y2)))
    strip_style(ln)
    apply_color(ln.line.color, color)
    ln.line.width = Pt(width)
    if name:
        ln.name = name
    return ln


# ---------------------------------------------------------------- 表

def _cell_border(cell, side: str, color: str | None, width_pt: float) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tag = {"L": "a:lnL", "R": "a:lnR", "T": "a:lnT", "B": "a:lnB"}[side]
    for el in tc_pr.findall(qn(tag)):
        tc_pr.remove(el)
    ln = etree.Element(qn(tag), w=str(int(width_pt * 12700)) if color else "0")
    if color:
        ln.append(color_xml(color))
    else:
        etree.SubElement(ln, qn("a:noFill"))
    # 罫線は tcPr の先頭に lnL, lnR, lnT, lnB の順で並べる必要がある
    order = ["a:lnL", "a:lnR", "a:lnT", "a:lnB"]
    pos = 0
    for t in order[: order.index(tag)]:
        pos += len(tc_pr.findall(qn(t)))
    tc_pr.insert(pos, ln)


def add_table(shapes, box, spec: dict, placeholder=None, ts: TypeScale | None = None):
    columns = spec.get("columns") or []
    rows = spec.get("rows") or []
    if not rows and not columns:
        raise SpecError("表に columns も rows も無い")
    n_rows = len(rows) + (1 if columns else 0)
    n_cols = len(columns) if columns else max(len(r) for r in rows)
    if placeholder is not None and hasattr(placeholder, "insert_table"):
        frame = placeholder.insert_table(n_rows, n_cols)
    else:
        x, y, w, h = box
        frame = shapes.add_table(n_rows, n_cols, Emu(x), Emu(y), Emu(w), Emu(h))
    table = frame.table
    style = spec.get("style", "clean")
    table.first_row = bool(columns)
    table.first_col = bool(spec.get("first_col", False))
    table.horz_banding = style == "template"
    if spec.get("col_widths"):
        ratios = spec["col_widths"]
        total_w = sum(c.width for c in table.columns)
        for col, r in zip(table.columns, ratios):
            col.width = int(total_w * r / sum(ratios))
    size = spec.get("font_size") or (ts["body"] if ts else 12)
    # 行の高さは文字に合わせて詰める（枠の高さいっぱいに引き伸ばさない）
    row_h = int(Pt(size) * 2.3)
    for r in table.rows:
        r.height = Emu(row_h)
    data = ([columns] if columns else []) + rows
    aligns = spec.get("align") or []
    hl_rows = {int(i) + (1 if columns else 0) for i in spec.get("highlight_rows", [])}
    hl_cols = {int(i) for i in spec.get("highlight_cols", [])}
    for r_i, row in enumerate(data):
        is_head = bool(columns) and r_i == 0
        for c_i in range(n_cols):
            cell = table.cell(r_i, c_i)
            value = row[c_i] if c_i < len(row) else ""
            opts = {"size": size}
            if c_i < len(aligns) and aligns[c_i]:
                opts["align"] = aligns[c_i]
            highlighted = r_i in hl_rows or (c_i in hl_cols and not is_head)
            if style == "clean":
                cell.margin_left = cell.margin_right = Emu(int(Pt(size) * 0.6))
                cell.margin_top = cell.margin_bottom = Emu(int(Pt(size) * 0.3))
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                if highlighted:
                    cell.fill.solid()
                    apply_color(cell.fill.fore_color, TINT)
                else:
                    cell.fill.background()
                opts["color"] = MUTED if is_head else "text1"
                if is_head or highlighted or (table.first_col and c_i == 0):
                    opts["bold"] = True
                for side in ("L", "R", "T"):
                    _cell_border(cell, side, None, 0)
                if is_head:
                    _cell_border(cell, "B", "text1", 1.25)
                else:
                    _cell_border(cell, "B", RULE, 0.5)
            elif highlighted:
                opts["bold"] = True
            text = "" if value is None else value if isinstance(value, (dict, list)) else str(value)
            write_text(cell.text_frame, text, opts)
    return frame


# ---------------------------------------------------------------- グラフ

def add_chart(shapes, box, spec: dict, placeholder=None, ts: TypeScale | None = None):
    ctype = spec.get("type", "column")
    if ctype not in CHART_TYPES:
        raise SpecError(f"グラフ種別 '{ctype}' は未対応。使える種別: {', '.join(CHART_TYPES)}")
    categories = spec["categories"]
    data = CategoryChartData()
    data.categories = categories
    for s in spec["series"]:
        data.add_series(s.get("name", ""), s["values"], number_format=spec.get("number_format"))
    xl_type = CHART_TYPES[ctype]
    if placeholder is not None and hasattr(placeholder, "insert_chart"):
        frame = placeholder.insert_chart(xl_type, data)
    else:
        x, y, w, h = box
        frame = shapes.add_chart(xl_type, Emu(x), Emu(y), Emu(w), Emu(h), data)
    chart = frame.chart
    fsize = spec.get("font_size") or (ts["sm"] if ts else 12)
    chart.font.size = Pt(fsize)
    apply_color(chart.font.color, "text1@0.25")
    title = spec.get("title")
    chart.has_title = bool(title)
    if title:
        chart.chart_title.text_frame.text = title
    single = len(spec["series"]) == 1
    legend = spec.get("legend", "bottom" if not single or ctype in ("pie", "doughnut") else False)
    chart.has_legend = bool(legend)
    if legend:
        chart.legend.position = LEGEND[legend]
        chart.legend.include_in_layout = False
    colors = spec.get("colors") or [f"accent{i}" for i in range(1, 7)]
    plot = chart.plots[0]
    # 強調: 指定したカテゴリだけ accent1、他は灰色（単系列の棒・円）
    hl = spec.get("highlight")
    hl_idx = None
    if hl is not None:
        hl_list = hl if isinstance(hl, list) else [hl]
        hl_idx = {categories.index(h) if isinstance(h, str) else int(h) for h in hl_list}
    is_bar = ctype in ("column", "bar", "stacked_column", "stacked_bar", "stacked_column_100", "stacked_bar_100")
    if ctype in ("pie", "doughnut"):
        for i, point in enumerate(plot.series[0].points):
            point.format.fill.solid()
            c = colors[i % len(colors)] if hl_idx is None else ("accent1" if i in hl_idx else f"text1@{0.55 + 0.1 * (i % 3)}")
            apply_color(point.format.fill.fore_color, c)
            point.format.line.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    else:
        for i, series in enumerate(plot.series):
            fmt = series.format
            if ctype.startswith("line"):
                fmt.line.width = Pt(2.5)
                apply_color(fmt.line.color, colors[i % len(colors)])
                series.smooth = False
            else:
                fmt.fill.solid()
                apply_color(fmt.fill.fore_color, colors[i % len(colors)])
                if single and hl_idx is not None:
                    for p_i, point in enumerate(series.points):
                        point.format.fill.solid()
                        apply_color(point.format.fill.fore_color, "accent1" if p_i in hl_idx else "text1@0.7")
        if is_bar:
            plot.gap_width = spec.get("gap_width", 60)
            if ctype.startswith("stacked"):
                plot.overlap = 100
        if ctype.startswith("bar") or ctype.startswith("stacked_bar"):
            # 横棒は既定で下から並ぶため、上から読める順にそろえる
            chart.category_axis.reverse_order = True
        cat = chart.category_axis
        cat.major_tick_mark = XL_TICK_MARK.NONE
        cat.format.line.color.rgb = RGBColor(0xBF, 0xBF, 0xBF)
        cat.has_major_gridlines = False
        val = chart.value_axis
        val.major_tick_mark = XL_TICK_MARK.NONE
        # データラベルを出す単系列の棒グラフは値軸を消す（数字の二重表示をなくす）
        hide_axis = spec.get("value_axis", not (spec.get("data_labels", True) and is_bar and single))
        val.visible = bool(hide_axis)
        val.has_major_gridlines = bool(spec.get("gridlines", hide_axis))
        if val.has_major_gridlines:
            val.major_gridlines.format.line.color.rgb = RGBColor(0xE3, 0xE3, 0xE3)
        val.format.line.fill.background()
        if spec.get("number_format"):
            val.tick_labels.number_format = spec["number_format"]
            val.tick_labels.number_format_is_linked = False
    if spec.get("data_labels", True):
        plot.has_data_labels = True
        labels = plot.data_labels
        labels.font.size = Pt(spec.get("label_size", fsize))
        apply_color(labels.font.color, "text1@0.15")
        if spec.get("number_format"):
            labels.number_format = spec["number_format"]
            labels.number_format_is_linked = False
        if ctype in ("column", "bar"):
            labels.position = XL_LABEL_POSITION.OUTSIDE_END
        elif ctype.startswith("line"):
            labels.position = XL_LABEL_POSITION.ABOVE
        elif ctype.startswith("stacked"):
            labels.position = XL_LABEL_POSITION.CENTER
            apply_color(labels.font.color, "bg1")
    return frame


def add_image(shapes, box, path: Path, placeholder=None):
    if not path.exists():
        raise SpecError(f"画像が無い: {path}")
    if placeholder is not None and hasattr(placeholder, "insert_picture"):
        return placeholder.insert_picture(str(path))
    x, y, w, h = box
    pic = shapes.add_picture(str(path), Emu(x), Emu(y))
    # 枠に収まるよう縦横比を保って縮小し、中央に置く
    scale = min(w / pic.width, h / pic.height)
    pic.width, pic.height = int(pic.width * scale), int(pic.height * scale)
    pic.left, pic.top = x + (w - pic.width) // 2, y + (h - pic.height) // 2
    return pic
