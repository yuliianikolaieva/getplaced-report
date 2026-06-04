#!/bin/bash
# Публікація звіту Get Placed (Glovo vs Bolt) на GitHub Pages.
# Копіює актуальний HTML у index.html, комітить і пушить у repo getplaced-report.
set -e

SRC="/Users/yuliia.nikolaieva/Downloads/Cursor/05_MS % getplaced/Аналіз_Glovo_vs_Bolt_correct.html"
DIR="$(cd "$(dirname "$0")" && pwd)"

if [ ! -f "$SRC" ]; then
  echo "Не знайдено вихідний файл: $SRC"
  exit 1
fi

echo "Копіюю свіжий звіт у index.html..."
cp "$SRC" "$DIR/index.html"

cd "$DIR"
git add index.html
if git diff --cached --quiet; then
  echo "Змін немає — нічого публікувати."
  exit 0
fi

git commit -m "Update Get Placed report ($(date '+%Y-%m-%d %H:%M'))"
git push origin main

echo "Готово. Live: https://yuliianikolaieva.github.io/getplaced-report/"
