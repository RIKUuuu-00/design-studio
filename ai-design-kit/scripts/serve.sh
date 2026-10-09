#!/usr/bin/env bash
# 比較キャンバスをローカル配信する（127.0.0.1 のみ）。
#   serve.sh [配信ルート=.] [開始ポート=4173]
# 配信ルートはプロジェクト直下にする（designs/ と decks/ の両方を同じオリジンで開くため）。
# 同一オリジンで開くと、キャンバス上での画面内テキスト編集を取得できる（file:// では不可）。
# Claude Code からはバックグラウンド実行する。
set -euo pipefail
root="${1:-.}"
port="${2:-4173}"
py="$(command -v python3 || command -v python)"
# 空きポートを探す
for _ in $(seq 1 30); do
  if ! "$py" - "$port" <<'EOF' 2>/dev/null
import socket, sys
s = socket.socket()
try:
    s.bind(("127.0.0.1", int(sys.argv[1])))
except OSError:
    sys.exit(1)
finally:
    s.close()
EOF
  then
    port=$((port + 1))
  else
    break
  fi
done
echo "serving ${root} at http://127.0.0.1:${port}/"
exec "$py" -m http.server "$port" --bind 127.0.0.1 --directory "$root"
