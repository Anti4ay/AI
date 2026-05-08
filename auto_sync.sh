#!/bin/bash
SOURCE="/Users/nikitachukreev/Documents/ИИ"
REPO="/Users/nikitachukreev/Documents/Система/AI"
TARGET="$REPO/ИИ"

cd "$REPO" || exit 1

git pull origin main --rebase 2>&1

rsync -av --delete "$SOURCE/" "$TARGET/" 2>&1

if [[ -n $(git status --porcelain) ]]; then
  git add -A
  git commit -m "sync ИИ: $(date '+%Y-%m-%d %H:%M:%S')"
  git push origin main 2>&1
fi
