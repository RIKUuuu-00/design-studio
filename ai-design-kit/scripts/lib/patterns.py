"""スライドの定番構図（パターン）。コンテンツ領域の中に、余白・文字サイズ・色の規則どおりに組む。

Claude が図形を 1 個ずつ置くと位置・間隔・文字サイズがぶれる。構図をここで固定し、
Claude は「どのパターンに何を入れるか」だけを決める。

パターン一覧（PATTERNS）:
  kpi            大きな数値を 2〜4 個並べる
  cards          見出し＋本文の塊を並べる（列数指定、複数行可）
  steps          手順（番号の丸を線でつなぐ）
  timeline       時系列（軸線と点、日付・見出し・本文）
  compare        2〜4 案の比較（推奨案を強調）
  matrix         2×2 マトリクス（軸ラベル付き）
  chart_takeaway グラフ＋示唆パネル
  statement      1 メッセージを大きく見せる
  agenda         目次（現在の章を強調）
"""

from __future__ import annotations

from .draw import MUTED, QUIET, RULE, TINT, SpecError, TypeScale, add_chart, line, rect, text_box, write_text
from .textfit import apply_scale, fit_scale


class Ctx:
    def __init__(self, prs, slide, box, spec: dict):
        self.prs = prs
        self.slide = slide
        self.shapes = slide.shapes
        self.x, self.y, self.w, self.h = (int(v) for v in box)
        self.ts = TypeScale(int(prs.slide_height), spec.get("base_size"), int(prs.slide_width))
        self.u = int(int(prs.slide_height) * 0.02)  # 余白の基本単位（7.5in で約 0.15in）
        self.made_text = []
        self.notes: list[str] = []

    def text(self, box, value, **opts):
        tb = text_box(self.shapes, box, value, **opts)
        self.made_text.append(tb)
        return tb

    def pt(self, size: float) -> int:
        return int(size * 12700)

    def lines_h(self, size: float, n: float, spacing: float = 1.45) -> int:
        return int(self.pt(size) * spacing * n)

    def est_h(self, value, size: float, width: int, spacing: float = 1.45, para_space: float = 0.0, bullets: bool = False) -> int:
        """実フォントで折り返しを計算した、テキストの高さの見積もり。"""
        if not value:
            return 0
        from .draw import flatten_items  # noqa: PLC0415
        from .textfit import Measurer, theme_fonts, wrap  # noqa: PLC0415

        _, minor = theme_fonts(self.slide)
        m = Measurer(minor)
        items = flatten_items(value)
        size_emu = self.pt(size)
        lines = 0
        for it in items:
            indent = (1.0 + 1.2 * int(it.get("level", 0))) if bullets else 0
            avail = max(width / size_emu - indent, 1)
            lines += len(wrap(str(it.get("text", "")), avail, m))
        # フォント差（PowerPoint と検証環境）で溢れないよう 1 割の余裕を持たせる
        return int((lines * size_emu * spacing + max(len(items) - 1, 0) * self.pt(para_space)) * 1.1)

    def balance(self, block_h: int, weight: float = 0.35) -> int:
        """内容の塊を領域内でやや上寄せに置くための開始 y。"""
        return self.y + max(int((self.h - block_h) * weight), 0)

    def fit_all(self, min_scale: float = 0.8) -> None:
        for tb in self.made_text:
            s = fit_scale(tb, self.slide, min_scale)
            if s is None:
                self.notes.append(f"「{tb.text_frame.text[:16]}…」が枠に収まらない（文字を減らす）")
                apply_scale(tb, self.slide, min_scale)
            elif s < 1.0:
                apply_scale(tb, self.slide, s)


def _cols(ctx: Ctx, n: int, gap: int | None = None) -> list[tuple[int, int]]:
    gap = ctx.u * 2 if gap is None else gap
    cw = (ctx.w - gap * (n - 1)) // n
    return [(ctx.x + i * (cw + gap), cw) for i in range(n)]


def _need(spec: dict, key: str, kind: str):
    if key not in spec or not spec[key]:
        raise SpecError(f"{kind} パターンに '{key}' が無い")
    return spec[key]


# ---------------------------------------------------------------- kpi

def kpi(ctx: Ctx, spec: dict) -> None:
    items = _need(spec, "items", "kpi")
    if not 1 <= len(items) <= 4:
        raise SpecError("kpi の items は 1〜4 個")
    ts = ctx.ts
    hl = spec.get("highlight")
    wide = ctx.prs.slide_width / ctx.prs.slide_height > 1.5
    v_size = spec.get("value_size") or (ts["display"] * (1.25 if wide else 1) if len(items) <= 3 else ts["xxl"] + 8)
    block_h = ctx.lines_h(v_size, 1, 1.25) + ctx.u * 2 + ctx.lines_h(ts["lead"], 2) + ctx.lines_h(ts["sm"], 2)
    caption = spec.get("caption")
    cap_h = ctx.lines_h(ts["lead"], 2) + ctx.u * 3 if caption else 0
    top = ctx.balance(block_h + cap_h, 0.4)
    cols = _cols(ctx, len(items), ctx.u * 3)
    for i, (cx, cw) in enumerate(cols):
        it = items[i]
        strong = hl is None or i == hl
        y = top
        ctx.text((cx, y, cw, ctx.lines_h(v_size, 1, 1.25)), it["value"], size=v_size, bold=True, heading=True,
                 color="accent1" if strong else "text1@0.35", wrap=False)
        y += ctx.lines_h(v_size, 1, 1.25) + ctx.u // 2
        line(ctx.shapes, cx, y, cx + min(cw, ctx.pt(v_size) * 1.6), y, "accent1" if strong else RULE, 2.0)
        y += ctx.u
        label_h = ctx.est_h(it.get("label", ""), ts["lead"], cw, 1.4) or ctx.lines_h(ts["lead"], 1)
        ctx.text((cx, y, cw, label_h), it.get("label", ""), size=ts["lead"], bold=True, color="text1")
        y += label_h + ctx.u // 2
        if it.get("note"):
            ctx.text((cx, y, cw, ctx.lines_h(ts["sm"], 2)), it["note"], size=ts["sm"], color=MUTED)
    if caption:
        cy = top + block_h + ctx.u
        line(ctx.shapes, ctx.x, cy, ctx.x + ctx.w, cy, RULE, 0.75)
        ctx.text((ctx.x, cy + ctx.u * 1.5, ctx.w, cap_h - ctx.u * 1.5), caption, size=ts["lead"], color="text1")


# ---------------------------------------------------------------- cards

def cards(ctx: Ctx, spec: dict) -> None:
    items = _need(spec, "items", "cards")
    n = len(items)
    cols_n = int(spec.get("cols", n if n <= 4 else 3))
    rows_n = (n + cols_n - 1) // cols_n
    ts = ctx.ts
    style = spec.get("style", "rule")  # rule: 上罫線のみ / tint: 薄い面 / outline: 枠線
    hl = spec.get("highlight")
    numbered = spec.get("numbered", False)
    gap_y = ctx.u * 2
    ch = (ctx.h - gap_y * (rows_n - 1)) // rows_n
    cols = _cols(ctx, cols_n, ctx.u * 2)
    pad = ctx.u * 1.5 if style in ("tint", "outline") else 0
    top = ctx.y
    if spec.get("compact", True):
        # 本文が短いのに縦に長いカードは間延びするので、内容に合わせて高さを決め、領域内でやや上寄せに置く
        iw = cols[0][1] - 2 * pad
        need = max(
            ctx.est_h(it.get("title", ""), ts["lead"], iw, 1.35) + ctx.u // 2
            + ctx.est_h(it.get("body"), ts["body"], iw, 1.45, ts["body"] * 0.35, isinstance(it.get("body"), list))
            for it in items
        ) + int(ctx.u * 2.4) + int(2 * pad)
        ch = min(ch, need)
        top = ctx.balance(ch * rows_n + gap_y * (rows_n - 1), 0.3)
    for i, it in enumerate(items):
        r, c = divmod(i, cols_n)
        cx, cw = cols[c]
        cy = top + r * (ch + gap_y)
        strong = hl is not None and i == hl
        if style == "tint" or strong:
            rect(ctx.shapes, (cx, cy, cw, ch), fill=TINT if strong or style == "tint" else None)
        elif style == "outline":
            rect(ctx.shapes, (cx, cy, cw, ch), fill=None, line=RULE, line_w=0.75)
        if style == "rule" or strong:
            line(ctx.shapes, cx, cy, cx + cw, cy, "accent1" if (strong or hl is None) else "text1@0.6", 2.5)
        ix, iw = int(cx + pad), int(cw - 2 * pad)
        y = int(cy + ctx.u * 1.2 + (pad if style != "rule" else 0))
        head = it.get("title", "")
        if numbered:
            head = f"{i + 1}. {head}"
        head_h = ctx.est_h(head, ts["lead"], iw, 1.35)
        ctx.text((ix, y, iw, head_h), head, size=ts["lead"], bold=True, heading=True, color="accent1" if strong else "text1")
        y += head_h + ctx.u // 2
        body = it.get("body")
        if body:
            ctx.text((ix, y, iw, int(cy + ch - y - pad)), body, size=ts["body"], color="text1",
                     bullets=isinstance(body, list), para_space=ts["body"] * 0.35)


# ---------------------------------------------------------------- steps

def steps(ctx: Ctx, spec: dict) -> None:
    items = _need(spec, "items", "steps")
    ts = ctx.ts
    current = spec.get("current")
    cols = _cols(ctx, len(items), ctx.u * 2)
    d = min(int(ctx.pt(ts["xl"]) * 1.7), cols[0][1] // 3)
    cw0 = cols[0][1]
    text_h = max(ctx.est_h(it.get("title", ""), ts["lead"], cw0, 1.35) + ctx.u // 2
                 + ctx.est_h(it.get("body"), ts["body"], cw0, 1.45, ts["body"] * 0.3, isinstance(it.get("body"), list)) for it in items)
    y0 = ctx.balance(d + int(ctx.u * 1.5) + text_h, 0.3)
    for i, (cx, cw) in enumerate(cols):
        done = current is None or i <= current
        color = "accent1" if done else "text1@0.7"
        if i < len(cols) - 1:
            nx = cols[i + 1][0]
            line(ctx.shapes, cx + d + ctx.u // 2, y0 + d // 2, nx - ctx.u // 2, y0 + d // 2, color if (current is None or i < current) else RULE, 1.5)
        circle = rect(ctx.shapes, (cx, y0, d, d), fill=color, shape="oval")
        tf = circle.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        write_text(tf, str(i + 1), {"size": ts["h"], "bold": True, "color": "bg1", "align": "center", "valign": "middle", "heading": True})
        it = items[i]
        y = y0 + d + ctx.u * 1.5
        head_h = ctx.lines_h(ts["lead"], 2, 1.35)
        ctx.text((cx, y, cw, head_h), it.get("title", ""), size=ts["lead"], bold=True, heading=True, color="text1" if done else MUTED)
        y += head_h + ctx.u // 2
        if it.get("body"):
            ctx.text((cx, y, cw, ctx.y + ctx.h - y), it["body"], size=ts["body"], color="text1" if done else MUTED,
                     bullets=isinstance(it["body"], list), para_space=ts["body"] * 0.3)


# ---------------------------------------------------------------- timeline

def timeline(ctx: Ctx, spec: dict) -> None:
    items = _need(spec, "items", "timeline")
    ts = ctx.ts
    current = spec.get("current")
    cols = _cols(ctx, len(items), ctx.u * 2)
    date_h = ctx.lines_h(ts["body"], 1)
    cw0 = cols[0][1]
    text_h = max(ctx.est_h(it.get("title", ""), ts["lead"], cw0, 1.35) + ctx.u // 2
                 + ctx.est_h(it.get("body"), ts["body"], cw0, 1.45, ts["body"] * 0.3, isinstance(it.get("body"), list)) for it in items)
    base_y = ctx.balance(date_h + ctx.u * 4 + text_h, 0.3)
    axis_y = base_y + date_h + ctx.u * 2
    line(ctx.shapes, ctx.x, axis_y, ctx.x + ctx.w, axis_y, "text1@0.6", 1.5)
    dot = int(ctx.pt(ts["body"]) * 0.9)
    for i, (cx, cw) in enumerate(cols):
        it = items[i]
        past = current is None or i <= current
        is_cur = current is not None and i == current
        ctx.text((cx, base_y + ctx.u // 2, cw, date_h), it.get("date", ""), size=ts["body"], bold=True,
                 color="accent1" if past else MUTED, wrap=False)
        size = int(dot * (1.5 if is_cur else 1))
        rect(ctx.shapes, (cx, axis_y - size // 2, size, size), fill="accent1" if past else "bg1",
             line="accent1" if past else "text1@0.5", line_w=1.5, shape="oval")
        y = axis_y + ctx.u * 2
        head_h = ctx.lines_h(ts["lead"], 2, 1.35)
        ctx.text((cx, y, cw, head_h), it.get("title", ""), size=ts["lead"], bold=True, heading=True, color="text1" if past else MUTED)
        y += head_h + ctx.u // 2
        if it.get("body"):
            ctx.text((cx, y, cw, ctx.y + ctx.h - y), it["body"], size=ts["body"], color="text1" if past else MUTED,
                     bullets=isinstance(it["body"], list), para_space=ts["body"] * 0.3)


# ---------------------------------------------------------------- compare

def compare(ctx: Ctx, spec: dict) -> None:
    columns = _need(spec, "columns", "compare")
    ts = ctx.ts
    hl = spec.get("highlight")
    cols = _cols(ctx, len(columns), ctx.u * 3)
    pad = ctx.u * 1.5
    cw0 = cols[0][1]
    tag_h = ctx.lines_h(ts["sm"], 1) if any(c.get("tag") for c in columns) else 0
    head_h = max(ctx.est_h(c.get("title", ""), ts["h"], cw0, 1.3) for c in columns)
    items_h = max(ctx.est_h(c.get("items"), ts["body"], cw0, 1.45, ts["body"] * 0.4, True) for c in columns)
    v_h = max((ctx.est_h(c.get("verdict"), ts["lead"], cw0, 1.4) for c in columns), default=0)
    block_h = pad + tag_h + head_h + ctx.u * 1.5 + items_h + (ctx.u * 2 + v_h if v_h else 0) + pad
    block_h = min(block_h, ctx.h)
    top = ctx.balance(block_h, 0.25)
    for i, (cx, cw) in enumerate(cols):
        col = columns[i]
        strong = hl is not None and i == hl
        if strong:
            rect(ctx.shapes, (cx - pad, top, cw + 2 * pad, block_h), fill=TINT)
        y = top + pad
        if tag_h:
            ctx.text((cx, y, cw, tag_h), col.get("tag", ""), size=ts["sm"], bold=True, color="accent1" if strong else MUTED)
            y += tag_h
        ctx.text((cx, y, cw, head_h), col.get("title", ""), size=ts["h"], bold=True, heading=True, color="accent1" if strong else "text1",
                 valign="bottom")
        y += head_h + ctx.u // 2
        line(ctx.shapes, cx, y, cx + cw, y, "accent1" if strong else "text1@0.5", 2.0 if strong else 1.0)
        y += ctx.u
        if col.get("items"):
            ctx.text((cx, y, cw, items_h), col["items"], size=ts["body"], color="text1", bullets=True, para_space=ts["body"] * 0.4)
        y += items_h + ctx.u
        if col.get("verdict"):
            line(ctx.shapes, cx, y, cx + cw, y, RULE, 0.75)
            ctx.text((cx, y + ctx.u, cw, v_h), col["verdict"], size=ts["lead"], bold=True, color="accent1" if strong else "text1")


# ---------------------------------------------------------------- matrix

def matrix(ctx: Ctx, spec: dict) -> None:
    quads = _need(spec, "quadrants", "matrix")
    if len(quads) != 4:
        raise SpecError("matrix の quadrants は 4 個（左上・右上・左下・右下の順）")
    ts = ctx.ts
    hl = spec.get("highlight")
    label_w = ctx.lines_h(ts["body"], 1) + ctx.u
    label_h = ctx.lines_h(ts["body"], 1) + ctx.u
    gx, gy = ctx.x + label_w, ctx.y
    gw, gh = ctx.w - label_w, ctx.h - label_h
    gap = ctx.u // 2
    qw, qh = (gw - gap) // 2, (gh - gap) // 2
    for i, q in enumerate(quads):
        r, c = divmod(i, 2)
        qx, qy = gx + c * (qw + gap), gy + r * (qh + gap)
        strong = hl is not None and i == hl
        rect(ctx.shapes, (qx, qy, qw, qh), fill=TINT if strong else QUIET)
        pad = ctx.u * 1.5
        ctx.text((qx + pad, qy + pad, qw - 2 * pad, ctx.lines_h(ts["lead"], 1, 1.35)), q.get("title", ""), size=ts["lead"], bold=True,
                 heading=True, color="accent1" if strong else "text1")
        if q.get("items"):
            ty = qy + pad + ctx.lines_h(ts["lead"], 1, 1.35) + ctx.u // 2
            ctx.text((qx + pad, ty, qw - 2 * pad, qy + qh - ty - pad), q["items"], size=ts["body"], color="text1", bullets=True,
                     para_space=ts["body"] * 0.3)
    # 軸
    ax_y = gy + gh + ctx.u // 2
    line(ctx.shapes, gx, ax_y, gx + gw, ax_y, "text1@0.4", 1.25)
    ctx.text((gx, ax_y + ctx.u // 3, gw, ctx.lines_h(ts["body"], 1)), spec.get("x_label", ""), size=ts["body"], color=MUTED, align="center")
    ax_x = gx - ctx.u // 2
    line(ctx.shapes, ax_x, gy, ax_x, gy + gh, "text1@0.4", 1.25)
    ylab = ctx.text((ctx.x - gh // 2 + label_w // 2, gy + gh // 2 - ctx.lines_h(ts["body"], 0.5), gh, ctx.lines_h(ts["body"], 1)),
                    spec.get("y_label", ""), size=ts["body"], color=MUTED, align="center")
    ylab.rotation = 270
    ctx.made_text.remove(ylab)


# ---------------------------------------------------------------- chart_takeaway

def chart_takeaway(ctx: Ctx, spec: dict) -> None:
    chart = _need(spec, "chart", "chart_takeaway")
    ts = ctx.ts
    ratio = spec.get("ratio", 0.62)
    cw = int(ctx.w * ratio)
    add_chart(ctx.shapes, (ctx.x, ctx.y, cw, ctx.h), chart, ts=ts)
    px = ctx.x + cw + ctx.u * 2
    pw = ctx.x + ctx.w - px
    take = spec.get("takeaway", {})
    rect(ctx.shapes, (px, ctx.y, pw, ctx.h), fill=TINT)
    pad = ctx.u * 1.5
    y = ctx.y + pad
    head = take.get("title", "示唆")
    ctx.text((px + pad, y, pw - 2 * pad, ctx.lines_h(ts["lead"], 1, 1.35)), head, size=ts["lead"], bold=True, heading=True, color="accent1")
    y += ctx.lines_h(ts["lead"], 1, 1.35) + ctx.u
    if take.get("points"):
        ctx.text((px + pad, y, pw - 2 * pad, ctx.y + ctx.h - y - pad), take["points"], size=ts["lead"], color="text1", bullets=True,
                 para_space=ts["lead"] * 0.6)


# ---------------------------------------------------------------- statement

def statement(ctx: Ctx, spec: dict) -> None:
    text = _need(spec, "text", "statement")
    ts = ctx.ts
    size = spec.get("size") or ts["xxl"]
    main_h = ctx.lines_h(size, 3, 1.35)
    sub = spec.get("sub")
    sub_h = ctx.lines_h(ts["lead"], 3) if sub else 0
    total = ctx.u * 3 + main_h + (ctx.u * 2 + sub_h if sub else 0)
    y = ctx.y + max((ctx.h - total) // 2, 0)
    width = int(ctx.w * spec.get("width", 0.82))
    line(ctx.shapes, ctx.x, y, ctx.x + ctx.pt(size) * 1.5, y, "accent1", 3.0)
    y += ctx.u * 3
    ctx.text((ctx.x, y, width, main_h), text, size=size, bold=True, heading=True, color="text1")
    if sub:
        ctx.text((ctx.x, y + main_h + ctx.u * 2, width, sub_h), sub, size=ts["lead"], color=MUTED)


# ---------------------------------------------------------------- agenda

def agenda(ctx: Ctx, spec: dict) -> None:
    items = _need(spec, "items", "agenda")
    ts = ctx.ts
    current = spec.get("current")
    row_h = min(ctx.h // len(items), ctx.lines_h(ts["xl"], 1, 2.8))
    num_w = ctx.pt(ts["xl"]) * 2.2
    top = ctx.balance(row_h * len(items), 0.3)
    for i, it in enumerate(items):
        y = top + i * row_h
        active = current is None or i == current
        title = it if isinstance(it, str) else it.get("title", "")
        note = None if isinstance(it, str) else it.get("note")
        ctx.text((ctx.x, y, num_w, row_h), f"{i + 1:02d}", size=ts["xl"], bold=True, heading=True,
                 color="accent1" if active else "text1@0.6", valign="middle", wrap=False)
        tw = ctx.w - num_w
        ctx.text((ctx.x + num_w, y, tw * (0.6 if note else 1), row_h), title, size=ts["h"], bold=active and current is not None,
                 heading=True, color="text1" if active else MUTED, valign="middle")
        if note:
            ctx.text((ctx.x + num_w + int(tw * 0.62), y, int(tw * 0.38), row_h), note, size=ts["body"], color=MUTED, valign="middle", align="right")
        line(ctx.shapes, ctx.x, y + row_h, ctx.x + ctx.w, y + row_h, "accent1" if (current is not None and i == current) else RULE,
             1.5 if (current is not None and i == current) else 0.75)


PATTERNS = {
    "kpi": kpi,
    "cards": cards,
    "steps": steps,
    "timeline": timeline,
    "compare": compare,
    "matrix": matrix,
    "chart_takeaway": chart_takeaway,
    "statement": statement,
    "agenda": agenda,
}


def draw_pattern(prs, slide, box, spec: dict) -> list[str]:
    kind = spec.get("type")
    if kind not in PATTERNS:
        raise SpecError(f"パターン '{kind}' は無い。使えるパターン: {', '.join(PATTERNS)}")
    ctx = Ctx(prs, slide, box, spec)
    PATTERNS[kind](ctx, spec)
    ctx.fit_all(spec.get("min_scale", 0.8))
    return ctx.notes
