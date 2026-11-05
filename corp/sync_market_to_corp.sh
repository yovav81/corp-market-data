#!/usr/bin/env bash
# מעתיק את נתוני השוק העדכניים מ-corp-market-data אל corp-data/03_Market/data (לשאילתת Power Query במאסטר החי)
# ובונה מחדש את קובץ הממשלתי corp-data/00_Gov/ממשלתי.xlsx.
# שימוש: bash corp/sync_market_to_corp.sh <corp-market-data dir> <work dir>
# יציאה: 0 = הקובץ מעודכן ליום המסחר הצפוי; 2 = נתוני יום המסחר הצפוי עדיין לא במאגר (נבנה מהנתונים האחרונים שיש)
#
# יום המסחר הצפוי: בערב (מ-12:00) — היום; בבוקר (עד 12:00) — יום המסחר הקודם (ריצת גיבוי של הבוקר).
# הריצה המתוזמנת של market-data ב-GitHub מתעכבת לפעמים בשעות (תור של GitHub) ומסתיימת אחרי 23:10 — לכן בערב,
# אם נתוני היום חסרים, הסקריפט מפעיל את הריצה בעצמו (workflow_dispatch) ומחכה לה (WAIT_MAX שניות, ברירת מחדל 35 דקות).
# בבוקר אין dispatch (fetch_market מתארך לפי השעה — משיכה בבוקר הייתה נרשמת על היום החדש); רק בנייה ממה שיש.
set -euo pipefail
SRC=${1:-.}; WORK=${2:-/tmp/corp-data-sync}
REPO=yovav81/corp-market-data
YEAR=$(TZ=Asia/Jerusalem date +%Y); TODAY=$(TZ=Asia/Jerusalem date +%F); HOUR=$(TZ=Asia/Jerusalem date +%-H)
prev_trading_day() {   # יום המסחר הקודם (הבורסה: שני–שישי)
  local d=$1 dow
  while :; do d=$(date -u -d "$d -1 day" +%F); dow=$(date -u -d "$d" +%u); [ "$dow" -ge 1 ] && [ "$dow" -le 5 ] && { echo "$d"; return; }; done
}
if [ "$HOUR" -ge 12 ]; then EXPECT=$TODAY; MODE=evening; else EXPECT=$(prev_trading_day "$TODAY"); MODE=morning; fi
cd "$SRC"; git pull -q --rebase || true
LAT=data/latest.csv; SP=data/spreads/$YEAR.csv
[ -f "$LAT" ] && [ -f "$SP" ] || { echo "missing source files"; exit 2; }
last_date() { tail -n +2 "$SP" | cut -d, -f1 | sort | tail -1; }
LAST=$(last_date)
echo "spreads last date: $LAST (expected $EXPECT, $MODE, today $TODAY)"

# --- ערב: נתוני היום חסרים → להפעיל את ריצת הנתונים ולחכות לה ---
if [ "$LAST" != "$EXPECT" ] && [ "$MODE" = evening ] && [ "${NO_DISPATCH:-0}" != 1 ]; then
  if gh api -X POST "repos/$REPO/actions/workflows/market-data.yml/dispatches" -f ref=main >/dev/null 2>&1; then
    echo "dispatched market-data ($(date -u +%H:%MZ)); waiting for $EXPECT data"
    DEADLINE=$(( $(date +%s) + ${WAIT_MAX:-2100} ))
    while [ "$(date +%s)" -lt "$DEADLINE" ]; do
      sleep 60; git pull -q --rebase || true; LAST=$(last_date)
      if [ "$LAST" = "$EXPECT" ]; then echo "today's data arrived ($(date -u +%H:%MZ))"; break; fi
      # הריצה הסתיימה בלי לשמור יום חדש (יום ללא מסחר / כשל) — אין טעם להמשיך לחכות
      ST=$(gh api "repos/$REPO/actions/workflows/market-data.yml/runs?per_page=1" --jq '.workflow_runs[0] | "\(.status) \(.conclusion)"' 2>/dev/null || echo "")
      case "$ST" in "completed failure"*|"completed cancelled"*) echo "market-data run ended: $ST"; break;; esac
    done
    [ "$LAST" = "$EXPECT" ] || echo "waited; latest data still $LAST"
  else
    echo "dispatch failed (gh not available or no permission) — building from existing data"
  fi
fi

REMOTE=$(git remote get-url origin | sed 's#corp-market-data#corp-data#')
rm -rf "$WORK"; git clone -q --depth 1 "$REMOTE" "$WORK"
mkdir -p "$WORK/03_Market/data"
cp "$LAT" "$WORK/03_Market/data/daily_latest.csv"
cp "$SP" "$WORK/03_Market/data/spreads_$YEAR.csv"
# כל קבצי הממשלתי (data/gov/*.csv) → 03_Market/data/gov_<שם>.csv — מקור ה-Power Query של קובץ הממשלתי v2
for f in data/gov/*.csv; do [ -f "$f" ] && cp "$f" "$WORK/03_Market/data/gov_$(basename "$f")"; done
# קובץ הממשלתי v2 — נבנה רק כשגרסת הבנייה (gov/VERSION) שונה מזו שב-corp-data/00_Gov/_build/VERSION או כשהקובץ חסר.
# הנתונים זורמים ב-Power Query מקבצי gov_*.csv; הקלטים הידניים של יובב עוברים (--settings). הגרסה הקודמת → 00_Gov/_archive.
GOV="$WORK/00_Gov/ממשלתי.xlsx"; mkdir -p "$WORK/00_Gov/_build" "$WORK/00_Gov/_archive"
python3 -c "import pandas, numpy, openpyxl" 2>/dev/null || pip install -q --break-system-packages pandas numpy openpyxl 2>/dev/null || true
if [ -f data/gov/bonds_latest.csv ]; then
  if [ ! -f "$GOV" ] || ! cmp -s gov/VERSION "$WORK/00_Gov/_build/VERSION"; then
    echo "בניית קובץ הממשלתי v$(cat gov/VERSION) (קודם: $(cat "$WORK/00_Gov/_build/VERSION" 2>/dev/null || echo 'אין'))"
    [ -f "$GOV" ] && cp "$GOV" "$WORK/00_Gov/_archive/ממשלתי_לפני_v$(cat gov/VERSION)_$(TZ=Asia/Jerusalem date +%F).xlsx"
    if python3 gov/build_gov2.py data "$GOV.new" --settings "$GOV"; then
      mv -f "$GOV.new" "$GOV"
      cp gov/VERSION gov/build_gov2.py gov/build_gov.py gov/pq_gov.py gov/check_gov2.py gov/settings_default.json gov/gov_data.py gov/gov_decide.py gov/gov_history.py gov/gov_cpi.py gov/gov_foreign.py "$WORK/00_Gov/_build/"
    else echo "אזהרה: בניית קובץ הממשלתי v2 נכשלה — הקובץ הקודם נשאר"; rm -f "$GOV.new"; fi
  else echo "קובץ הממשלתי: גרסה $(cat gov/VERSION) כבר מותקנת — הנתונים מתעדכנים דרך Power Query"; fi
fi
cd "$WORK"; git add 03_Market/data 00_Gov
if git diff --cached --quiet; then echo "no change"; else
  git -c user.name="market-bot" -c user.email="corp-bot@users.noreply.github.com" commit -qm "נתוני שוק וקובץ ממשלתי $LAST"
  for i in 1 2 3; do git push -q && break || { git pull -q --rebase; sleep 3; }; done
  echo "pushed market data $LAST"
fi
[ "$LAST" = "$EXPECT" ] || exit 2
