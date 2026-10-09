# deck.json の書き方

`build_deck.py` の入力。テンプレのレイアウトからスライドを起こし、プレースホルダーに流し込む。
完全な例は リポジトリの `tests/fixtures/sample-deck.json`。

## 全体

```json
{
  "title": "資料タイトル（ファイルのプロパティに入る）",
  "template": "design-system/slides/template.pptx",
  "page_numbers": true,
  "footer": "株式会社〇〇 / 社外秘",
  "slides": [ ... ]
}
```

| キー | 既定 | 説明 |
| --- | --- | --- |
| `template` | 必須 | 実行ディレクトリか deck.json からの相対パス |
| `page_numbers` | `true` | レイアウトにスライド番号枠があれば表示する |
| `footer` | なし | レイアウトにフッター枠があれば入れる |
| `keep_template_slides` | `false` | テンプレ内のサンプルスライドを残す（通常は使わない） |
| `fit` | `"shrink"` | 本文プレースホルダーが溢れたら最大 15% まで文字を縮める。`"none"` で縮めず報告だけ。タイトルは常に縮めない |
| `base_size` | 自動 | パターン・表・グラフの本文サイズ（pt）。既定はスライドの大きさから決まる（16:9 で約 15pt） |

## スライド

```json
{
  "layout": "Title and Content",
  "fill": { "title": "...", "body": [...] },
  "add": [ ... ],
  "notes": "話す内容・出典",
  "keep_empty": false,
  "page_number": true
}
```

- `layout`: layouts.md のレイアウト名、または ID（`L02`）
- `fill`: キー → 値。キーは layouts.md の「流し込めるキー」（`title` `subtitle` `body` `body2` `picture` …）。`#13` で idx 指定、プレースホルダー名でも可
- 使わなかったプレースホルダーは自動で消える（`keep_empty: true` で残す）
- `notes`: 発表者ノート。出典や推測値の根拠もここに書く

## fill の値

| 値 | 結果 |
| --- | --- |
| `"文字列"` | 段落。`\n` で改段。`**太字**` が使える |
| `["要点1", ["補足a", "補足b"], "要点2"]` | 箇条書き。入れ子の配列は 1 段下げ |
| `{"text": "...", "size": 14, "bold": true, "color": "accent1", "align": "center"}` | 書式付き（必要なときだけ） |
| `{"table": {...}}` | ネイティブの表（本文枠・表枠に入る） |
| `{"chart": {...}}` | ネイティブのグラフ（本文枠・グラフ枠に入る） |
| `{"image": "path/to.png"}` | 画像。画像枠なら枠に合わせてトリミング、それ以外は縦横比を保って収める |
| `{"pattern": "kpi", ...}` | 本文枠の位置にパターンを組む（枠は外れる）。中身は下の「パターン」と同じ |

### 表

```json
{"table": {
  "columns": ["項目", "初年度", "2年目以降"],
  "rows": [["ライセンス", "480万円", "480万円"], ["**合計**", "**780万円**", "**480万円**"]],
  "col_widths": [2, 1, 1],
  "align": ["left", "right", "right"],
  "font_size": 14,
  "first_col": false,
  "highlight_rows": [1],
  "style": "clean"
}}
```

- `style`: `clean`（既定。罫線だけの読みやすい表）/ `template`（テンプレの表スタイル）
- `highlight_rows` / `highlight_cols`: 強調する行・列（0 始まり、見出し行は数えない）。薄い強調色＋太字
- 数値列は `align: "right"`。単位は列見出しかセルにそろえて書く
- 行が 8 を超えるなら分割するか、要点だけ残す

### グラフ

```json
{"chart": {
  "type": "column",
  "categories": ["4月", "5月", "6月"],
  "series": [{"name": "売上", "values": [120, 135, 150]}],
  "title": null,
  "number_format": "#,##0\"万円\"",
  "data_labels": true,
  "legend": false,
  "colors": ["accent1"],
  "highlight": "6月",
  "value_axis": false,
  "font_size": 12
}}
```

- `highlight`: 強調するカテゴリ（名前か 0 始まりの番号、配列も可）。単系列の棒・円で、そこだけ accent1、他は灰色になる。**タイトルの主張と同じ要素を強調する**
- `value_axis`: 値軸を出すか。既定はデータラベル付きの単系列棒グラフでは消す（数字の二重表示を避ける）

- `type`: `column` `stacked_column` `stacked_column_100` `bar` `stacked_bar` `stacked_bar_100` `line` `line_plain` `area` `stacked_area` `pie` `doughnut`
- 色は既定でテーマの accent1→6 の順。強調したい系列だけ色を変え、他は `text2@0.6` などで沈める
- タイトルはスライドの主張タイトルと重複するので通常は付けない
- `bar`（横棒）は上から categories の順に並ぶ

## パターン（推奨）

本文スライドは「タイトルのみ」系レイアウトに `add` でパターンを 1 つ置く。位置は既定で `area: "content"`
（タイトルの下からフッターの上まで）。余白・文字サイズ・色・行間はパターンが決め、文字は実フォントで測って収める。

| type | 主なキー |
| --- | --- |
| `kpi` | `items: [{value, label, note}]`（1〜4 個）, `highlight`（番号）, `caption`（だから何か） |
| `cards` | `items: [{title, body}]`（body は文字列か配列）, `cols`, `highlight`, `numbered`（手順のときだけ）, `style`（`rule` / `tint` / `outline`） |
| `steps` | `items: [{title, body}]`（3〜5 個）, `current`（今の段階。以降は灰色） |
| `timeline` | `items: [{date, title, body}]`（4〜6 個）, `current` |
| `compare` | `columns: [{tag, title, items, verdict}]`（2〜4 列）, `highlight`（推奨案） |
| `matrix` | `quadrants: [{title, items}]`（左上・右上・左下・右下の順で 4 個）, `x_label`, `y_label`, `highlight` |
| `chart_takeaway` | `chart: {グラフと同じ}`, `takeaway: {title, points}`, `ratio`（グラフの幅、既定 0.62） |
| `statement` | `text`（1 文。`\n` で改行位置を決める）, `sub`, `size` |
| `agenda` | `items: [文字列 か {title, note}]`, `current` |

```json
{"layout": "タイトルのみ", "fill": {"title": "日報の入力負荷が提出率と商談化率を押し下げている"},
 "add": [{"type": "kpi", "highlight": 0, "items": [
   {"value": "25分", "label": "日報1件の平均入力時間", "note": "2026年上期 営業120名の実測"},
   {"value": "62%", "label": "日報の提出率", "note": "目標90%に対し28pt不足"}],
   "caption": "入力負荷を下げれば、提出率と記録の質が上がる"}]}
```

完全な例はリポジトリの `tests/fixtures/showcase-deck.json`（全パターンを使った 12 枚）。

## add（部品の手置き）

パターンに無い図解だけ部品で組む。位置は `box: [x, y, w, h]`（スライドに対する %）か、`area`（`content` かこのレイアウトのプレースホルダーのキー）。

| type | 主なキー |
| --- | --- |
| `text` | `text`, `size`, `bold`, `color`, `align`, `valign` |
| `shape` | `shape`（`rect` `rounded` `oval` `chevron` `pentagon` `arrow_right` `arrow_down` `triangle`）, `fill`, `line`, `text`, `text_color`, `size`, `bold` |
| `line` | `from: [x, y]`, `to: [x, y]`（%）, `color`, `width` |
| `table` / `chart` / `image` | fill と同じ中身をトップレベルに書く（`{"type": "chart", "area": "body", "categories": ...}`） |
| `grid` | `cols`, `rows`, `gap`（%）, `items`（中身は上の type。位置は自動） |

```json
{"type": "grid", "area": "body", "cols": 3, "items": [
  {"type": "shape", "shape": "chevron", "fill": "accent1", "text": "Step 1\n定型化", "bold": true},
  {"type": "shape", "shape": "chevron", "fill": "accent1@0.3", "text": "Step 2\n音声入力", "bold": true},
  {"type": "shape", "shape": "chevron", "fill": "accent1@0.5", "text": "Step 3\nAI要約", "bold": true}
]}
```

- 矢羽根・矢印の中の文字は折り返さない。改行は `\n` で明示する
- 図解は `Title Only` 系のレイアウトに載せ、タイトルは必ずプレースホルダーに入れる

## 色

`accent1`〜`accent6`, `text1`, `text2`, `bg1`, `bg2`（テーマ色）。`@` で明るさを調整（`accent1@0.4` は明るく、`accent1@-0.25` は暗く）。
`#RRGGBB` も書けるが、テンプレ準拠から外れるので使わない（check_deck が color_off_theme として報告する）。
