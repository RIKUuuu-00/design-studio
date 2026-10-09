---
name: design-lab
description: UIの画面・モックアップ・ワイヤーフレーム・LPを、実装前に複数案で比較したいときに使う。ブリーフから構造の異なる案を素のHTMLで生成し、ローカル比較キャンバスを作り、スクショで自己検証してから提示する。選択案の実装は design-handoff を使う。パワポ資料は deck-lab を使う。
argument-hint: "[<slug>] [案数=3] [画面幅=375,1280] [検証周回=2]"
---

# design-lab

画面の方向性を実装前に決めるため、構造の異なる案を並べて比べさせる。
人が触るのは「ブリーフを投げる」「キャンバスで選ぶ」の 2 点だけ。生成・検証は止まらずに進める。

## 道具

スクリプトは `${CLAUDE_PLUGIN_ROOT}/scripts/` にある。

| 用途 | コマンド |
| --- | --- |
| スクショ（PNG/PDF） | `node ${CLAUDE_PLUGIN_ROOT}/scripts/export.mjs designs/<slug> --widths 375,1280` |
| トークン逸脱の検出 | `node ${CLAUDE_PLUGIN_ROOT}/scripts/token-lint.mjs designs/<slug> --tokens design-system/tokens.css` |
| 比較キャンバス生成 | `python3 ${CLAUDE_PLUGIN_ROOT}/scripts/build_canvas.py designs/<slug>/canvas.json --out designs/<slug>/index.html` |
| 配信 | `bash ${CLAUDE_PLUGIN_ROOT}/scripts/serve.sh .`（バックグラウンドで実行） |

Playwright MCP（`playwright`）も使える。ホバー・クリック後の状態を確かめたいときに使う。

## 優先順位

**ブリーフの指示 ＞ design-system/ ＞ 既存コードの見た目 ＞ 自分の判断**。自分の好みでスタイルを決めない。

## デザインの考え方

**frontend-design** に従う。公式の frontend-design プラグインが有効ならその Skill を、無ければ同梱の
`${CLAUDE_PLUGIN_ROOT}/references/frontend-design/frontend-design.md` を手順 1 の前に読む。
特に「主題から考える」「計画 → 計画の見直し → 作る → 自己批評」「生成っぽさの一覧」「見せ場は 1 か所」を守る。
デザインシステムがある場合、frontend-design の自由度は **トークンの範囲内**（構図・情報の見せ方・文字の扱い）で使う。

## 手順

### 1. 基準を読む

- `design-system/tokens.css` と `components.md` を読む
- 無ければ frontend-design の「計画」（色 4〜6 色・書体・構図・原則）をブリーフの主題から立て、`${CLAUDE_PLUGIN_ROOT}/starter/design-system/` を `design-system/` に複製したうえで tokens.css / tokens.json の値を計画に合わせて書き換える。「仮のデザインシステムを作った。ブランドに合わせて差し替えが必要」と返答に明記する

### 2. 制約を集める

- 対象画面の既存コードがあれば読み、表示するデータ項目・既存部品・遷移を抜き出す
- 不足は推測で補い、仮定としてキャンバス冒頭に載せる。質問で止めない

### 3. 案を設計する（既定 3 案）

**情報密度・レイアウト構造・操作導線のうち 2 軸以上** を案ごとに変える。色違い・角丸違いは別案にしない。

| 軸 | 振れ幅の例 |
| --- | --- |
| 情報密度 | 要約カード中心 ↔ 表で全件一覧 |
| レイアウト構造 | 1 カラム縦積み ↔ サイドバー＋詳細 ↔ タブ切替 |
| 操作導線 | 一覧から詳細へ遷移 ↔ その場で展開・編集 ↔ ウィザード |

各案に **狙い** と **代償** を 1 行ずつ付ける。

案ごとに frontend-design の「計画」を短く書き（構図の ASCII ワイヤーフレーム、揃え、見せ場）、
「同じような依頼なら毎回これになるか？」で見直してから作る。見直しで変えた点は meta.json の `notes` に残す。

### 4. アートボードを作る

`designs/<slug>/<id>.html`（id は `a` `b` `c`…）。1 案 = 1 つの自己完結 HTML。

- 素の HTML + CSS。React 等のフレームワークは使わない（どのスタックにも移植でき、スクショが安定する）
- `<link rel="stylesheet" href="../../design-system/tokens.css">` を読み込み、**値は tokens.css の変数だけ** を使う
- 部品は components.md の見本に合わせる。無い部品が要るなら、トークンだけで組み、返答で「components.md への追加候補」として挙げる
- 375px と 1280px の両方で崩れないレスポンシブにする。375px で横スクロールを出さない
- ダミーデータは現実的な日本語・桁数で書く（「テキスト」「XXX」は使わない）
- 外部 CDN・画像 URL は使わない（社内環境で表示できないことがある）。画像は CSS の面で代替する

`designs/<slug>/meta.json` に案ごとの `{id, title, aim, tradeoff, status}` を書く（status: `draft` / `selected` / `implemented`）。

### 5. キャンバスを作る

`designs/<slug>/canvas.json`（`track: "ui"`, `widths: [375, 1280]`, 各案 `src: "a.html"`。形式は build_canvas.py の docstring）を書いて `build_canvas.py` を実行する。

### 6. 自己検証（最大 2 周）

1. `export.mjs` で各案を 375px と 1280px で撮る。出力の `[横はみ出し]` `[コンソールエラー]` は先に直す
2. `token-lint.mjs` を実行し、逸脱を直す
3. **design-reviewer サブエージェント** に `shots/*.png`、`checklists/review.md`、`design-system/tokens.css`、`components.md` のパスだけを渡す（検出モード）。
   続けて frontend-design の「生成っぽさの一覧」と照らした自己批評を行い、当てはまる箇所は直す（装飾を 1 つ外しても伝わるなら外す）
4. 指摘を直して撮り直す。2 周回して残った指摘は canvas.json の各案の `notes` に書き、キャンバスを再生成する（合否で止めない）

### 7. 返す

`serve.sh` をバックグラウンドで起動し、`http://127.0.0.1:<port>/designs/<slug>/index.html` を返す。返答に含めるもの:

- 各案の狙い・代償の要約と、**どれを推すかとその理由**
- 置いた仮定、残った指摘
- 使い方: 「案を選び、コメントやテキストの直接編集をして **Copy as prompt** を押し、ここに貼る」

## 再実行

- キャンバスから貼られたプロンプト（採用案・コメント・テキスト修正）→ その案の修正として扱う。テキスト修正はそのまま反映する
- 「案Bをベースにあと 2 案」→ 既存案を残したまま `d` `e` を追加し、キャンバスを再生成する
- 「案Bで実装して」→ design-handoff を使う

## やらないこと

- tokens.css にない値の直書き（色・余白・角丸・文字サイズ・フォント）
- 本番コードへの書き込み（それは design-handoff の仕事）
