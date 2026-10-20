#!/usr/bin/env bash
# מעתיק את נתוני השוק העדכניים מ-corp-market-data אל corp-data/03_Market/data (לשאילתת Power Query במאסטר החי).
# שימוש: bash corp/sync_market_to_corp.sh <corp-market-data dir> <work dir>
# יציאה: 0 = עודכן / לא היה שינוי; 2 = אין נתוני שוק של היום (לא עודכן)
set -euo pipefail
SRC=${1:-.}; WORK=${2:-/tmp/corp-data-sync}
YEAR=$(TZ=Asia/Jerusalem date +%Y); TODAY=$(TZ=Asia/Jerusalem date +%F)
cd "$SRC"; git pull -q --rebase || true
LAT=data/latest.csv; SP=data/spreads/$YEAR.csv
[ -f "$LAT" ] && [ -f "$SP" ] || { echo "missing source files"; exit 2; }
LAST=$(tail -n +2 "$SP" | cut -d, -f1 | sort | tail -1)
echo "spreads last date: $LAST (today $TODAY)"
REMOTE=$(git remote get-url origin | sed 's#corp-market-data#corp-data#')
rm -rf "$WORK"; git clone -q --depth 1 "$REMOTE" "$WORK"
mkdir -p "$WORK/03_Market/data"
cp "$LAT" "$WORK/03_Market/data/daily_latest.csv"
cp "$SP" "$WORK/03_Market/data/spreads_$YEAR.csv"
cd "$WORK"; git add 03_Market/data
if git diff --cached --quiet; then echo "no change"; else
  git -c user.name="market-bot" -c user.email="corp-bot@users.noreply.github.com" commit -qm "03_Market/data: נתוני שוק $LAST"
  for i in 1 2 3; do git push -q && break || { git pull -q --rebase; sleep 3; }; done
  echo "pushed market data $LAST"
fi
[ "$LAST" = "$TODAY" ] || exit 2
