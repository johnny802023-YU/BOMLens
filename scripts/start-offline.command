#!/bin/zsh
set -e

cd "$(dirname "$0")/.."
clear
echo "BOMLens 離線模式"
echo "資料只在本機處理；請勿把終端機視窗關閉。"
echo

if ! command -v node >/dev/null 2>&1; then
  echo "找不到 Node.js，請先請 IT 協助安裝 Node.js 22 或更新版本。"
  read -r "?按 Enter 關閉..."
  exit 1
fi

if [ ! -d node_modules ]; then
  echo "缺少程式套件。請先在可連線環境執行：npm ci"
  read -r "?按 Enter 關閉..."
  exit 1
fi

npm run build
npm run offline:serve &
server_pid=$!
trap 'kill "$server_pid" 2>/dev/null || true' EXIT INT TERM

for attempt in {1..30}; do
  if curl --silent --fail http://127.0.0.1:3784/ >/dev/null 2>&1; then
    open http://127.0.0.1:3784/
    wait "$server_pid"
    exit $?
  fi
  sleep 1
done

echo "啟動逾時，請把這個視窗的內容交給 IT 人員。"
wait "$server_pid"
