---
name: deck-lab
description: 提案書・報告書・社内説明資料などのパワポ（.pptx）を、完成度の高い状態で作るときに使う。社内テンプレがあればそれに準拠し（テンプレモード）、無ければデザイン方針からテンプレを作る（デザインモード）。論理構成の異なる骨子案を比較キャンバスで選ばせ、定番構図（パターン）で組み、実フォントでの文字量チェック・画像化・レビューと採点を合格するまで回す。テンプレ登録・差し替え時（引数 register）は register.md に従う。
argument-hint: "[register | <slug>] [案数] [検証周回]"
---

# deck-lab

資料を「骨子比較 → 選択 → 組版 → 検証と採点」で作り、**そのまま出せる完成度** まで仕上げる。
PowerPoint でそのまま編集できる pptx（テーマ・レイアウト・ネイティブの表とグラフ）を返す。

人が判断するのは「骨子の選択」と「最終仕上げ」の 2 点だけ。それ以外は止まらずに進める。

## 道具

スクリプトは `${CLAUDE_PLUGIN_ROOT}/scripts/` にある。Python は `python3`（Windows は `python`）。

| 用途 | コマンド |
| --- | --- |
| 環境チェック（初回） | `python3 ${CLAUDE_PLUGIN_ROOT}/scripts/doctor.py --template design-system/slides/template.pptx` |
| テンプレ生成（デザインモード） | `python3 ${CLAUDE_PLUGIN_ROOT}/scripts/make_theme.py design-system/slides/theme.json --out design-system/slides/template.pptx` |
| テンプレ登録 | `python3 ${CLAUDE_PLUGIN_ROOT}/scripts/register_template.py design-system/slides/template.pptx` |
| 比較キャンバス生成 | `python3 ${CLAUDE_PLUGIN_ROOT}/scripts/build_canvas.py decks/<slug>/canvas.json --out decks/<slug>/index.html` |
| 配信 | `bash ${CLAUDE_PLUGIN_ROOT}/scripts/serve.sh .`（バックグラウンドで実行） |
| 組版 | `python3 ${CLAUDE_PLUGIN_ROOT}/scripts/build_deck.py decks/<slug>/deck.json --out decks/<slug>/deck.pptx --report decks/<slug>/build.json` |
| 構造チェック | `python3 ${CLAUDE_PLUGIN_ROOT}/scripts/check_deck.py decks/<slug>/deck.pptx --json decks/<slug>/check.json` |
| 画像化 | `python3 ${CLAUDE_PLUGIN_ROOT}/scripts/render_pptx.py decks/<slug>/deck.pptx --out decks/<slug>/renders --sheet` |

| 資料 | 内容 |
| --- | --- |
| [design-guide.md](design-guide.md) | 見た目の判断基準（frontend-design の考え方を資料向けに読み替えたもの）。**組版の前に必ず読む** |
| [deck-spec.md](deck-spec.md) | deck.json の書き方（パターン・表・グラフの指定） |
| [checklists/slide-review.md](checklists/slide-review.md) | 崩れの検出観点 |
| [checklists/deck-critique.md](checklists/deck-critique.md) | 完成度の採点基準 |
| `${CLAUDE_PLUGIN_ROOT}/references/frontend-design/frontend-design.md` | frontend-design 本体（公式プラグインが有効ならそちらを優先） |

## 引数

- `register` → [register.md](register.md) を実行して終わる
- `<slug>` → `decks/<slug>/` を作業場所にする。省略時はブリーフから英小文字とハイフンで決める
- 案数: 既定 2（最大 3）。検証周回: 既定 3

## 手順

### 0. モードを決め、前提をそろえる

- `design-system/slides/template.pptx` がある → **テンプレモード**。見た目はテンプレが正。完成度は構図（パターン）と文字組で上げる
- 無い → **デザインモード**。frontend-design 本体と design-guide.md を読み、手順 0-D でテンプレを作る
- `layouts.md` が無ければ [register.md](register.md) を実行する。「用途: TODO」が残っていれば `thumbs/` を見て埋める
- 初めてのプロジェクトなら `doctor.py` を実行し、NG を利用者に伝える（LibreOffice が無ければ手順 5 は代替運用）

#### 0-D. デザインモードのテンプレ作成

1. design-guide.md の「計画」を `decks/<slug>/design-plan.md` に書く（色 4〜6 色・書体・構図・見せ場）
2. 計画を見直す（「同じような依頼なら毎回これになるか？」）。直した点を 1 行残す
3. `design-system/slides/theme.json` に落とす（形式は make_theme.py の docstring。例は `${CLAUDE_PLUGIN_ROOT}/starter/themes/`）。フォントは doctor.py が列挙した導入済みのものから選び、**利用者の PC にも入っているか** を返答で確認事項にする
4. `make_theme.py` → `register_template.py` を実行する。`[配色]` の警告が出たら色を直して作り直す

### 1. 入力を整理する

- ブリーフから **目的（何を決めてほしいか）・読み手・枚数感・使われ方（説明用か配布用か）** を抜き出す
- 素材（議事録・メモ・既存資料・数値）を読む。既存 pptx は python-pptx でテキストを抜く
- 不足は推測で補い、**仮定として明記** する。質問で止めない（目的が全く読み取れない場合だけ聞く）

### 2. 骨子案を作る

案の違いは **論理構成** で出す。構成の型の例: 結論先行（ピラミッド）／課題起点（現状→原因→打ち手）／選択肢比較（案A/B/C→推奨）／時系列（経緯→現在→今後）／読み手の疑問順（Q&A）。

各案に必ず付けるもの:

- **狙い** と **代償** を 1 行ずつ
- スライドごとの **主張タイトル**（1 スライド 1 メッセージ。「〜は〜だ」「〜で〜する」の文）
- スライドごとの **要点**（2〜4 個）と **見せ方**（レイアウト名、またはパターン名。design-guide.md の「パターンの選び方」）

主張タイトルは layouts.md の `title` の目安（全角/行 × 行）に収める。

書き出すもの: `decks/<slug>/outlines.md`（人が読む用）と `decks/<slug>/canvas.json`（`track: "deck"`、各案の `outline` に `{no, title, points, layout}`。layout 欄には見せ方を書く）。
キャンバスを生成し、`serve.sh` をバックグラウンドで起動して URL（`http://127.0.0.1:<port>/decks/<slug>/index.html`）を返す。
各案の狙い・代償の要約と、**どれを推すかとその理由** を添えて、選択を待つ。

### 3. 選択を受け取る

- 「案B」などの指示、またはキャンバスの **Copy as prompt** の文面を受け取る。テキスト修正はそのまま反映し、コメントは骨子に反映する
- 「任せる」なら推した案で進める
- `outlines.md` に採用案と反映した修正を記録する

### 4. 組版する

design-guide.md を読んでから `decks/<slug>/deck.json` を書き、`build_deck.py` で pptx にする。

- 本文スライドは原則「タイトルのみ」系レイアウト＋ **パターン**（`add: [{"type": "kpi", ...}]`）で組む。箇条書きだけのスライドは全体の 3 割以下にする
- 表紙・章扉・statement・kpi の数値など大きな文字は、文節の切れ目で `\n` を入れる
- 図表は PowerPoint ネイティブ。グラフは `highlight` でタイトルの主張と同じ要素だけ色を付ける
- 色はテーマ色のみ、フォントは指定しない（テーマ継承）。テンプレモードでは独自の装飾図形を足さない
- 数値・固有名詞は素材から写す。素材に無い数値を作らない。推測値には「（仮）」を付けて報告に挙げる
- `build.json`（--report）の注意（縮小した・溢れた・最終行が 1〜2 文字）を確認する

### 5. 検証と採点（合格まで。最大 3 周、引数で変更可）

各周で次を行う:

1. `check_deck.py` … `error` と `widow` を **画像を見る前に** 0 にする
2. `render_pptx.py --sheet` で画像化する
3. **design-reviewer サブエージェント** を 2 回呼ぶ（生成の経緯は渡さない。甘くなるため）
   - 検出モード: スライド画像 + `checklists/slide-review.md` + `layouts.md` + `check.json` の残り
   - 採点モード（critique）: `renders/sheet.png` と各スライド画像 + `checklists/deck-critique.md` + `design-guide.md`
4. 「要修正」と、4 点未満の項目に付いた修正案を `deck.json`（必要なら theme.json）に反映して再ビルドする

**合格: 要修正 0 件、かつ採点が全項目 4 以上。** 3 周で合格しなければ、残った指摘と点数を報告に回す（止めない）。

LibreOffice が使えない環境（代替運用）: 利用者に「PowerPoint で `decks/<slug>/deck.pptx` を開き、`deck.pdf` として PDF 保存して」と依頼し、`render_pptx.py decks/<slug>/deck.pdf --out decks/<slug>/renders --sheet` から続ける。
画像を読めない（ゲートウェイが画像入力を通さない）場合は、3 を飛ばし `check_deck.py` だけで直し、その旨を報告に明記する。

### 6. 返す

`decks/<slug>/` に `deck.pptx`（納品物）、`deck.pdf`、`renders/sheet.png`、`outlines.md`、`deck.json`、（デザインモードなら）`design-plan.md` がそろっていることを確認して返す。返答に含めるもの:

- 何枚・どの骨子で作ったか、サムネ一覧のパス、**最終の採点**
- 残った指摘（スライド番号つき）と、推測で置いた数値・表現
- **PowerPoint で人が仕上げるべき点**（写真の差し替え、ロゴ、最終の言い回し）

## 再実行

- 「3枚目を表にして」「結論をもっと強く」→ `deck.json` を直して手順 4〜5（骨子は変えない）
- 「骨子からやり直し」「C案も見たい」→ 既存案を残したまま追加し、キャンバスを再生成する
- 「色を変えたい」（デザインモード）→ theme.json を直して make_theme → register → build
- テンプレを差し替えた → `register` を実行し直す（用途の記入は引き継がれる）

## やらないこと

- 社内テンプレのマスター・レイアウト・テーマを書き換えない（直すべき点は register の構造チェックとして人に返す）
- フォント名や RGB を deck.json に直接書かない
- パターンで組めるものを図形 1 個ずつの手置きで組まない（位置・間隔・文字サイズがぶれる）
