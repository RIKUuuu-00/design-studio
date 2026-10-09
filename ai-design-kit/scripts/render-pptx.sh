#!/usr/bin/env bash
# LibreOffice → PDF → PNG。実体は render_pptx.py（Windows でも動くよう Python で実装）。
#   render-pptx.sh decks/<slug>/deck.pptx decks/<slug>/renders [--sheet]
set -euo pipefail
here="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
src="${1:?使い方: render-pptx.sh <deck.pptx|deck.pdf> <出力ディレクトリ> [オプション]}"
out="${2:?出力ディレクトリを指定すること}"
shift 2
py="$(command -v python3 || command -v python)"
exec "$py" "$here/render_pptx.py" "$src" --out "$out" "$@"
