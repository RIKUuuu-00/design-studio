# design-studio

「ローカル版デザイン生成 構想書（パワポ優先）」の実装。ゲートウェイ認証の Claude Code でも、公式 `/design` 相当の **複数案を比較 → 選択 → 仕上げ** のループを回すためのプラグイン `ai-design-kit` を配布する。Artifacts は使わない。

| トラック | Skill | 成果物 |
| --- | --- | --- |
| パワポ（優先） | `deck-lab` | 社内テンプレ準拠の `.pptx`（PowerPoint でそのまま編集できる） |
| UI 画面 | `design-lab` → `design-handoff` | 比較キャンバス上の HTML アートボード → 既存スタックでの実装 |
| 共通 | `design-reviewer`（サブエージェント） | 画像だけを見た指摘 |

## 導入

```text
/plugin marketplace add RIKUuuu-00/design-studio
/plugin install ai-design-kit@design-studio
```

Claude Code 2.1.275 以降なら 1 行でも入る: `/plugin install ai-design-kit --marketplace RIKUuuu-00/design-studio`

### 必要なもの

| 用途 | 必要なもの | 無い場合 |
| --- | --- | --- |
| 資料（必須） | Python 3.10+、`pip install -r ai-design-kit/scripts/requirements.txt` | — |
| 資料の画像化 | LibreOffice、pdftoppm（poppler）または PyMuPDF | PowerPoint で PDF を書き出し、`.pdf` を渡す代替運用 |
| 資料の検証精度 | テンプレと同じ日本語フォント | 文字量判定で 1 割の余裕を残す（既定） |
| UI | Node 18+、Playwright | 社内プロキシでブラウザを取れなければ `PLAYWRIGHT_CHANNEL=chrome` か `msedge` |

環境チェック（Phase 0 の確認項目）は `deck-lab register` の中で `doctor.py` が自動で行う。LibreOffice・フォント・Playwright に加え、テスト画像を Claude 自身に読ませて **ゲートウェイ経由の画像入力が通るか** も確認する。手動で行う場合はこのリポジトリを clone して次を実行する:

```bash
python3 ai-design-kit/scripts/doctor.py --template <project>/design-system/slides/template.pptx --image-test /tmp/img.png
```

## 使い方（パワポ）

1. 利用側プロジェクトに `design-system/slides/template.pptx`（社内テンプレ）を置く
2. `/ai-design-kit:deck-lab register` … レイアウトカタログ（`layouts.md`）とサムネを作り、テンプレの構造上の問題を一覧で返す
3. `/ai-design-kit:deck-lab` に素材（議事録・メモ・数値）とブリーフ（目的・読み手・枚数）を渡す
4. 返ってきた比較キャンバスで骨子案を選び、コメントを書いて **Copy as prompt** → Claude Code に貼る
5. Claude が流し込み → 構造チェック → 画像化 → design-reviewer の指摘で修正（最大 2 周）を回し、`decks/<slug>/deck.pptx` を返す
6. PowerPoint で仕上げる

## 使い方（UI）

1. `/ai-design-kit:design-lab` にブリーフを渡す（`design-system/tokens.css` が無ければ初期値を複製する）
2. 比較キャンバスで案を選ぶ。テキストは画面上で直接書き換えられる → **Copy as prompt** → 貼る
3. `/ai-design-kit:design-handoff <slug> <案ID>` で既存スタックに実装し、スクショで突き合わせる

利用側プロジェクトの `CLAUDE.md` には `ai-design-kit/starter/CLAUDE.md.snippet` を貼る。

## 構成

```text
.claude-plugin/marketplace.json      # 配布元（このリポジトリ自体がマーケットプレイス）
ai-design-kit/
├── .claude-plugin/plugin.json
├── .mcp.json                        # Playwright MCP（バージョン固定）
├── skills/
│   ├── deck-lab/                    # SKILL.md / register.md / deck-spec.md / checklists/
│   ├── design-lab/                  # SKILL.md / templates/canvas.html（両トラック共通）/ checklists/
│   └── design-handoff/
├── agents/design-reviewer.md
├── scripts/
│   ├── doctor.py                    # Phase 0 環境チェック
│   ├── register_template.py         # テンプレ → layouts.md / layouts.json / thumbs/
│   ├── build_deck.py                # deck.json → .pptx（テンプレのレイアウトに流し込み）
│   ├── check_deck.py                # 空き枠・仮文字・文字量超過・はみ出し・重なり・テーマ逸脱
│   ├── render_pptx.py / render-pptx.sh   # pptx|pdf → PNG（+ サムネ一覧）
│   ├── build_canvas.py              # canvas.json → 比較キャンバス
│   ├── serve.sh                     # キャンバスのローカル配信
│   ├── export.mjs                   # アートボード → PNG/PDF（横はみ出し・JS エラーも報告）
│   └── token-lint.mjs               # トークン外の色・余白・フォント・未定義トークン
└── starter/                         # design-system の初期値、CLAUDE.md の追記例
tests/run_e2e.sh                     # 全スクリプトを使い捨てプロジェクトで通す
```

## 構想書からの変更点と理由

| 項目 | 構想書 | 実装 | 理由 |
| --- | --- | --- | --- |
| pptx Skill のラップ | anthropics/skills の pptx Skill に社内カタログを足す | **依存せず python-pptx で自前実装** | pptx Skill のライセンスは Proprietary で、複製・派生物の作成・サービス外での保持を禁じている。社内配布物に同梱・改変できない |
| 流し込みの実装 | テンプレのレイアウトを複製 | `deck.json`（宣言的な仕様）→ `build_deck.py` | Claude が毎回 XML を手で編集するより再現性が高く、修正が `deck.json` の差分で済む |
| 画像化スクリプト | `render-pptx.sh` | 実体は Python（`.sh` は薄いラッパー） | 社内 PC は Windows が多く、bash と pdftoppm を前提にできない |
| 視覚検証のスクショ | Playwright MCP | 一括撮影は `export.mjs`、操作確認は MCP | 案×幅の一括撮影は MCP を 1 枚ずつ呼ぶより速く、トークンも少ない |
| 環境確認 | 手作業の確認リスト | `doctor.py` | Phase 0 の項目（LibreOffice・フォント・画像入力）を 1 コマンドで確認できる |

## 既知の制約（v0.1）

- テンプレ内の **サンプルスライドの複製** は未対応。流し込めるのはレイアウトのプレースホルダーだけ。図解の見本スライドを多用するテンプレでは効果が落ちる
- 文字量の判定は概算（全角 = 1em）。最終判断は画像化 + レビュー
- LibreOffice の描画は PowerPoint と完全には一致しない（特にグラフ・自動縮小）。最終確認は PowerPoint で行う
- キャンバスは同一オリジン（`serve.sh` 経由）で開いたときだけ、画面内のテキスト編集を取得できる
- 社内テンプレでの検証はまだしていない（テストは python-pptx 既定テンプレの擬似テンプレで実施）

## テスト

```bash
bash tests/run_e2e.sh            # テンプレ登録 → 流し込み → チェック → 画像化 → キャンバス → UI 書き出し・lint
```
