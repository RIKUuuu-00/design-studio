#!/usr/bin/env bash
# ai-design-kit のスクリプト一式を、使い捨てプロジェクトで端から端まで通す。
#   tests/run_e2e.sh [作業ディレクトリ]
# 前提: python-pptx, Pillow, LibreOffice, pdftoppm（または PyMuPDF）, 日本語フォント, Node, Playwright
set -euo pipefail
repo="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
kit="$repo/ai-design-kit"
work="${1:-$(mktemp -d)}"
py="$(command -v python3 || command -v python)"
fail() { echo "FAIL: $*" >&2; exit 1; }
step() { echo; echo "== $*"; }

rm -rf "$work/project"
mkdir -p "$work/project"
cd "$work/project"
cp -r "$kit/starter/design-system" design-system

step "doctor"
"$py" "$kit/scripts/doctor.py" || echo "(doctor の NG は環境依存。続行する)"

step "テンプレ登録"
"$py" "$repo/tests/fixtures/make_template.py" design-system/slides/template.pptx >/dev/null
"$py" "$kit/scripts/register_template.py" design-system/slides/template.pptx
test -s design-system/slides/layouts.json || fail "layouts.json が無い"
grep -q "^### L02 Title and Content" design-system/slides/layouts.md || fail "layouts.md にレイアウト詳細が無い"
test -s design-system/slides/thumbs/sheet.png || fail "サムネ一覧が無い"

step "用途の引き継ぎ"
sed -i 's/^- \*\*用途\*\*: TODO.*$/- **用途**: テスト用途/' design-system/slides/layouts.md
"$py" "$kit/scripts/register_template.py" design-system/slides/template.pptx --no-render >/dev/null
[ "$(grep -c '^- \*\*用途\*\*: テスト用途' design-system/slides/layouts.md)" -ge 10 ] || fail "再登録で用途が消えた"

step "流し込み"
mkdir -p decks/sample
sed 's#"template": "../slides/template.pptx"#"template": "design-system/slides/template.pptx"#' "$repo/tests/fixtures/sample-deck.json" > decks/sample/deck.json
"$py" "$kit/scripts/build_deck.py" decks/sample/deck.json --out decks/sample/deck.pptx
"$py" - <<'EOF'
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
prs = Presentation("decks/sample/deck.pptx")
assert len(prs.slides) == 6, len(prs.slides)
s3 = prs.slides[2]
assert any(getattr(sh, "has_chart", False) and sh.has_chart for sh in s3.shapes), "スライド3にネイティブのグラフが無い"
s5 = prs.slides[5]
assert any(getattr(sh, "has_table", False) and sh.has_table for sh in s5.shapes), "スライド5にネイティブの表が無い"
for i, s in enumerate(prs.slides, 1):
    for sh in s.placeholders:
        if sh.has_text_frame and not sh.text_frame.text.strip() and sh.placeholder_format.type not in (PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.DATE):
            raise AssertionError(f"スライド{i}に空プレースホルダー {sh.name}")
import re, zipfile
z = zipfile.ZipFile("decks/sample/deck.pptx")
for n in z.namelist():
    if re.match(r"ppt/slides/slide\d+\.xml$", n):
        ids = re.findall(r'<p:cNvPr id="(\d+)"', z.read(n).decode())
        assert len(ids) == len(set(ids)), f"{n} に重複した図形 ID がある（PowerPoint が修復を求める）"
print("pptx の構造: OK")
EOF

step "spec エラーの検出"
echo '{"template":"design-system/slides/template.pptx","slides":[{"layout":"Title Only","fill":{"bodyy":"x"}}]}' > decks/sample/bad.json
if "$py" "$kit/scripts/build_deck.py" decks/sample/bad.json --out decks/sample/bad.pptx 2>err.txt; then fail "不正なキーで失敗しなかった"; fi
grep -q "使えるキー" err.txt || fail "エラーに使えるキーが出ていない"

step "構造チェック"
"$py" "$kit/scripts/check_deck.py" decks/sample/deck.pptx --json decks/sample/check.json
"$py" "$kit/scripts/check_deck.py" design-system/slides/template.pptx --json tmpl-check.json >/dev/null
grep -q leftover_text tmpl-check.json || fail "見本文の消し忘れを検出できていない"

step "画像化"
bash "$kit/scripts/render-pptx.sh" decks/sample/deck.pptx decks/sample/renders --sheet
[ "$(ls decks/sample/renders/slide-*.png | wc -l)" -eq 6 ] || fail "スライド画像が 6 枚でない"
test -s decks/sample/deck.pdf || fail "PDF が無い"

step "PDF 入力（代替運用）"
"$py" "$kit/scripts/render_pptx.py" decks/sample/deck.pdf --out decks/sample/renders-pdf >/dev/null
[ "$(ls decks/sample/renders-pdf/slide-*.png | wc -l)" -eq 6 ] || fail "PDF からの画像化に失敗"

step "文字組（禁則・最終行の検出）"
"$py" - "$kit/scripts" <<'EOF'
import sys
sys.path.insert(0, sys.argv[1])
from lib.textfit import Measurer, wrap
m = Measurer("sans-serif")
lines = wrap("入力時間を半分以下にできる。", 6, m)
assert not any(l and l[0] in "、。" for l in lines), f"行頭に句読点: {lines}"
assert wrap("CRM連携", 2.5, m)[0] == "CRM", wrap("CRM連携", 2.5, m)
print("禁則: OK")
EOF
grep -q '"widow"' decks/sample/check.json || fail "最終行 1 文字（る）を検出できていない"

step "デザインモード（テーマ生成 → 全パターン）"
mkdir -p decks/showcase design-mode
"$py" "$kit/scripts/make_theme.py" "$repo/tests/fixtures/theme-harbor.json" --out design-mode/template.pptx | tee theme.log
if grep -q '^\[配色\]' theme.log; then fail "既定テーマで配色の警告が出た"; fi
"$py" "$kit/scripts/register_template.py" design-mode/template.pptx --no-render >/dev/null
grep -q "タイトルのみ" design-mode/layouts.md || fail "生成テンプレのレイアウトがカタログに無い"
sed 's#"template": "template.pptx"#"template": "design-mode/template.pptx"#' "$repo/tests/fixtures/showcase-deck.json" > decks/showcase/deck.json
"$py" "$kit/scripts/build_deck.py" decks/showcase/deck.json --out decks/showcase/deck.pptx --report decks/showcase/build.json
"$py" "$kit/scripts/check_deck.py" decks/showcase/deck.pptx --json decks/showcase/check.json
"$py" - <<'EOF'
import json
items = json.load(open("decks/showcase/check.json"))
bad = [i for i in items if i["severity"] in ("error", "warn")]
assert not bad, "showcase に指摘が残っている: " + json.dumps(bad, ensure_ascii=False)
print("showcase: 指摘 0")
EOF
"$py" "$kit/scripts/render_pptx.py" decks/showcase/deck.pptx --out decks/showcase/renders --sheet >/dev/null
[ "$(ls decks/showcase/renders/slide-*.png | wc -l)" -eq 12 ] || fail "showcase の画像が 12 枚でない"
"$py" - <<'EOF'
import re, zipfile
z = zipfile.ZipFile("decks/showcase/deck.pptx")
for n in z.namelist():
    if re.match(r"ppt/slides/slide\d+\.xml$", n):
        ids = re.findall(r'<p:cNvPr id="(\d+)"', z.read(n).decode())
        assert len(ids) == len(set(ids)), f"{n} に重複した図形 ID がある"
print("showcase の図形 ID: OK")
EOF

step "骨子比較キャンバス"
cat > decks/sample/canvas.json <<'EOF'
{"track":"deck","slug":"sample","title":"営業日報のデジタル化","brief":"役員向け、10枚程度","assumptions":["予算は年1,000万円以内"],
 "options":[
  {"id":"A","title":"結論先行","aim":"冒頭で判断を求める","tradeoff":"背景の納得感は弱い","outline":[{"no":1,"title":"商談化率を3pt上げる","points":["入力5分以内"],"layout":"Title and Content"}]},
  {"id":"B","title":"課題起点","aim":"現場の課題から積み上げる","tradeoff":"結論までが長い","outline":[{"no":1,"title":"日報に25分かかっている","points":["自由記述が6割"],"layout":"Title and Content"}],"images":["renders/slide-01.png"]}
 ]}
EOF
"$py" "$kit/scripts/build_canvas.py" decks/sample/canvas.json --out decks/sample/index.html

step "UI: アートボード"
mkdir -p designs
cp -r "$repo/tests/fixtures/designs/sample" designs/sample
"$py" "$kit/scripts/build_canvas.py" designs/sample/canvas.json --out designs/sample/index.html
node "$kit/scripts/token-lint.mjs" designs/sample --tokens design-system/tokens.css | tee lint.txt
grep -q "b.html.*\[color\]" lint.txt || fail "b.html の色の直書きを検出できていない"
grep -q "b.html.*\[undefined-token\]" lint.txt || fail "未定義トークンを検出できていない"
if grep -q "a.html" lint.txt; then fail "a.html に誤検出がある"; fi
node "$kit/scripts/export.mjs" designs/sample --widths 375,1280
[ "$(ls designs/sample/shots/*.png | wc -l)" -eq 4 ] || fail "スクショが 4 枚でない"

step "キャンバスの動作"
node "$repo/tests/canvas_smoke.mjs" "$work/project"

echo
echo "ALL PASSED  ($work/project)"
