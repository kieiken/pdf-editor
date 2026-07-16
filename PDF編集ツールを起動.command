#!/bin/bash
# ============================================================
#  PDF・Word 結合 / ページ編集ツール  起動スクリプト
#  ダブルクリックで起動します（ローカルサーバー + ブラウザ）
#  ファイルは外部に送信されません。すべてこの端末内で処理します。
# ============================================================
cd "$(dirname "$0")" || exit 1

# python3 を探す（Finder からの起動は PATH が狭いため複数候補を確認）
PY=""
for c in /opt/homebrew/bin/python3 /usr/local/bin/python3 /usr/bin/python3 python3; do
  if command -v "$c" >/dev/null 2>&1; then PY="$c"; break; fi
done
if [ -z "$PY" ]; then
  echo "python3 が見つかりませんでした。"
  echo "ターミナルで  xcode-select --install  を実行してから再度お試しください。"
  read -n 1 -s -r -p "何かキーを押すと閉じます…"
  exit 1
fi

# 空きポートを探す（8800〜8899）
PORT=""
for p in $(seq 8800 8899); do
  if ! lsof -iTCP:"$p" -sTCP:LISTEN >/dev/null 2>&1; then PORT="$p"; break; fi
done
[ -z "$PORT" ] && PORT=8810
URL="http://localhost:$PORT/"

echo ""
echo "  ============================================"
echo "   PDF・Word 結合 / 編集ツールを起動します"
echo "   $URL"
echo "  ============================================"
echo ""

# このフォルダをローカル配信（バックグラウンド）
"$PY" -m http.server "$PORT" >/dev/null 2>&1 &
SRV=$!
sleep 1
open "$URL"

echo "  ブラウザが開きました。"
echo "  ※ このウィンドウは開いたままにしてください（閉じると停止します）。"
echo "  終了するには Control + C を押すか、このウィンドウを閉じてください。"
echo ""

trap 'kill $SRV 2>/dev/null; exit 0' INT TERM
wait $SRV
