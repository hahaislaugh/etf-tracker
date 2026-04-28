#!/bin/bash
# 每日自動更新：下載持股資料 → 產生 HTML → push 到 GitHub（觸發 Pages 部署）
set -e

cd "$(dirname "$0")"
LOG="update.log"

echo "====== $(TZ=Asia/Taipei date '+%Y/%m/%d %H:%M:%S') 開始更新 ======" >> "$LOG"

# 1. 下載資料並重新產生 HTML
/usr/bin/python3 update_data.py >> "$LOG" 2>&1

# 2. Git push（推送後 GitHub Actions 自動部署 Pages）
git add index.html 復華/ 群益/ 群益982/ 統一/ 野村/ >> "$LOG" 2>&1
if git diff --cached --quiet; then
    echo "  無變動，略過 commit" >> "$LOG"
else
    git commit -m "chore: 自動更新持股資料 $(TZ=Asia/Taipei date '+%Y/%m/%d %H:%M')" >> "$LOG" 2>&1
    git push >> "$LOG" 2>&1
    echo "  ✓ Push 完成" >> "$LOG"
fi

echo "====== 完成 ======" >> "$LOG"
