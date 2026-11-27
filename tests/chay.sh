#!/bin/sh
# Chạy toàn bộ thử nghiệm: sh tests/chay.sh  (cần: npm i, pip playwright python-docx openpyxl pillow, Chrome)
cd "$(dirname "$0")"
python3 fixtures.py >/dev/null
python3 ui_local.py
node mock-gas.js >/tmp/mock-gas.log 2>&1 & P=$!; sleep 1
python3 ui_server.py
kill $P
