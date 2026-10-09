---
name: deck-lab
description: 提案書・報告書・社内説明資料などのパワポ（.pptx）を、登録済みの社内テンプレで作るときに使う。素材から論理構成の異なる骨子案を2〜3つ作って比較キャンバスで選ばせ、選ばれた骨子をテンプレのレイアウトに流し込み、画像化して自己検証する。テンプレ未登録・テンプレ差し替え時（引数 register）は register.md に従う。
argument-hint: "[register | <slug>] [案数] [検証周回]"
---

# deck-lab

社内テンプレ準拠の資料を「骨子比較 → 選択 → 流し込み → 自己検証」で作る。
**自由に描かない。** テンプレのレイアウトとテーマを正とし、PowerPoint でそのまま編集できる pptx を返す。

人が判断するのは「骨子の選択」と「最終仕上げ」の 2 点だけ。それ以外は止まらずに進める。

## 道具

スクリプトは `${CLAUDE_PLUGIN_ROOT}/scripts/` にある。Python は `python3`（Windows は `python`）。

| 用途 | コマンド |
| --- | --- |
| 環境チェック（初回） | `python3 ${CLAUDE_PLUGIN_ROOT}/scripts/doctor.py --template design-system/slides/template.pptx` |
| テンプレ登録 | `python3 ${CLAUDE_PLUGIN_ROOT}/scripts/register_template.py design-system/slides/template.pptx` |
| 比較キャンバス生成 | `python3 ${CLAUDE_PLUGIN_ROOT}/scripts/build_canvas.py decks/<slug>/canvas.json --out decks/<slug>/index.html` |
| 配信 | `bash ${CLAUDE_PLUGIN_ROOT}/scripts/serve.sh .`（バックグラウンドで実行） |
| 流し込み | `python3 ${CLAUDE_PLUGIN_ROOT}/scripts/build_deck.py decks/<slug>/deck.json --out decks/<slug>/deck.pptx` |
| 構造チェック | `python3 ${CLAUDE_PLUGIN_ROOT}/scripts/check_deck.py decks/<slug>/deck.pptx --json decks/<slug>/check.json` |
| 画像化 | `python3 ${CLAUDE_PLUGIN_ROOT}/scripts/render_pptx.py decks/<slug>/deck.pptx --out decks/<slug>/renders --sheet` |

`deck.json` の書き方は [deck-spec.md](deck-spec.md)、レビュー観点は [checklists/slide-review.md](checklists/slide-review.md)。

## 引数

- `register` → [register.md](register.md) を実行して終わる
- `<slug>` → `decks/<slug>/` を作業場所にする。省略時はブリーフから英小文字とハイフンで決める
- 案数: 既定 2（最大 3）。検証周回: 既定 2

## 手順

### 0. 前提をそろえる

1. `design-system/slides/layouts.md` を読む。無ければ [register.md](register.md) を実行してから続ける
2. layouts.md に「用途: TODO」が残っていれば、`thumbs/` の画像を見て埋めてから続ける
3. このプロジェクトで初めて使うなら `doctor.py` を実行し、NG を利用者に伝える（LibreOffice が無ければ手順 5 は代替運用）

### 1. 入力を整理する

- ブリーフから **目的（何を決めてほしいか）・読み手・枚数感・使われ方（説明用か配布用か）** を抜き出す
- 素材（議事録・メモ・既存資料・数値）を読む。既存 pptx は python-pptx でテキストを抜く
- 不足は推測で補い、**仮定として明記** する。質問で止めない（目的が全く読み取れない場合だけ聞く）

### 2. 骨子案を作る

案の違いは **論理構成** で出す。見た目の案は作らない（テンプレ準拠なら見た目の差は小さい）。
構成の型の例: 結論先行（ピラミッド）／課題起点（現状→原因→打ち手）／選択肢比較（案A/B/C→推奨）／時系列（経緯→現在→今後）／読み手の疑問順（Q&A）。

各案に必ず付けるもの:

- **狙い** と **代償** を 1 行ずつ（例: 狙い「冒頭で判断を求められる」／代償「背景の納得感は弱い」）
- スライドごとの **主張タイトル**（1 スライド 1 メッセージ。名詞止めでなく「〜で〜する」「〜は〜だ」の文）
- スライドごとの **要点**（2〜4 個）と **使うレイアウト**（layouts.md の用途を見て選ぶ）と、図表があれば種類

主張タイトルは layouts.md の `title` の目安（全角/行 × 行）に収める。収まらない主張は分けるか短くする。

書き出すもの:

- `decks/<slug>/outlines.md` … 全案の骨子（人が読む用）
- `decks/<slug>/canvas.json` … `track: "deck"`。各案の `outline` に `{no, title, points, layout}`（形式は `build_canvas.py` の docstring）

キャンバスを生成し、`serve.sh` をバックグラウンドで起動して URL（`http://127.0.0.1:<port>/decks/<slug>/index.html`）を返す。
返答には各案の狙い・代償の要約と、**どれを推すかとその理由** を添えて、選択を待つ。

### 3. 選択を受け取る

- 「案B」「Bで」などの指示、またはキャンバスの **Copy as prompt** から貼られた文面（採用案・コメント・テキスト修正）を受け取る
- 貼られたテキスト修正はそのまま反映する。コメントは骨子に反映する
- 「任せる」なら推した案で進める
- `outlines.md` に採用案と反映した修正を記録する

### 4. 流し込む

採用した骨子を `decks/<slug>/deck.json` に書き、`build_deck.py` で pptx にする。

- レイアウトは名前か ID（`L02`）で指定し、`fill` のキーは layouts.md の「流し込めるキー」だけを使う
- 図表は PowerPoint ネイティブの表・グラフで作る（`{"table": ...}` / `{"chart": ...}`）。画像化したグラフは使わない
- テンプレに無い図解（手順フロー、カード並び）だけ `add` で組む。色はテーマ色（`accent1` など）、フォントは指定しない（テーマを継承）
- 数値・固有名詞は素材から写す。素材に無い数値を作らない。推測値には「（仮）」を付けて報告に挙げる
- spec エラーが出たら、メッセージの「使えるキー」「使えるレイアウト」に従って直す

### 5. 自己検証（最大 2 周、引数で変更可）

1. `check_deck.py` を実行する。`error` は画像を見る前に直す（空き枠・仮文字・文字量超過・はみ出し）
2. `render_pptx.py --sheet` で画像化する
3. **design-reviewer サブエージェント** に渡す。渡すのは次だけ（生成の経緯は渡さない。甘くなるため）:
   - 画像のパス（`renders/slide-XX.png`）
   - `checklists/slide-review.md` のパス
   - `design-system/slides/layouts.md` のパス
   - `check.json` の残りの指摘
4. 指摘を `deck.json` に反映して再ビルドする。2 周回して残った指摘は報告に回す（合否で止めない）

LibreOffice が使えない環境（代替運用）: 利用者に「PowerPoint で `decks/<slug>/deck.pptx` を開き、`deck.pdf` として PDF 保存して」と依頼し、`render_pptx.py decks/<slug>/deck.pdf --out decks/<slug>/renders --sheet` から続ける。

画像を読めない（ゲートウェイが画像入力を通さない）場合は、手順 5-3 を飛ばし `check_deck.py` の結果だけで直し、その旨を報告に明記する。

### 6. 返す

`decks/<slug>/` に次がそろっていることを確認して返す:

- `deck.pptx`（納品物）、`deck.pdf`、`renders/sheet.png`（サムネ一覧）
- `outlines.md`（採用した骨子と修正履歴）、`deck.json`（再生成用）

返答に含めるもの:

- 何枚・どの骨子で作ったか、サムネ一覧のパス
- 残った指摘（スライド番号つき）と、推測で置いた数値・表現
- **PowerPoint で人が仕上げるべき点**（例: 写真の差し替え、社名ロゴ、最終の言い回し）

## 再実行

- 「3枚目を表にして」「結論をもっと強く」→ `deck.json` を直して手順 4〜5 をやり直す（骨子は変えない）
- 「骨子からやり直し」「C案も見たい」→ 既存案を残したまま追加し、キャンバスを再生成する
- テンプレを差し替えた → `register` を実行し直す（用途の記入は引き継がれる）

## やらないこと

- テンプレのマスター・レイアウト・テーマを書き換えない（直すべき点は register の構造チェックとして人に返す）
- フォント名や RGB を直接指定しない
- python-pptx / pptxgenjs で一からデザインを描かない
