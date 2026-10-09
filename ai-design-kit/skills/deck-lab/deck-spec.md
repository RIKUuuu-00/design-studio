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

### 表

```json
{"table": {
  "columns": ["項目", "初年度", "2年目以降"],
  "rows": [["ライセンス", "480万円", "480万円"], ["**合計**", "**780万円**", "**480万円**"]],
  "col_widths": [2, 1, 1],
  "align": ["left", "right", "right"],
  "font_size": 14,
  "first_col": false
}}
```

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
  "font_size": 12
}}
```

- `type`: `column` `stacked_column` `stacked_column_100` `bar` `stacked_bar` `stacked_bar_100` `line` `line_plain` `area` `stacked_area` `pie` `doughnut`
- 色は既定でテーマの accent1→6 の順。強調したい系列だけ色を変え、他は `text2@0.6` などで沈める
- タイトルはスライドの主張タイトルと重複するので通常は付けない
- `bar`（横棒）は上から categories の順に並ぶ

## add（テンプレに無い図解）

位置は `box: [x, y, w, h]`（スライドに対する %）か、`area: "body"`（このレイアウトのプレースホルダーの位置を借りる）。

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
