# gov/ — מקור הממשלתי של מערכת CORP וקובץ הממשלתי v2

## מה יש כאן
| קובץ | תפקיד |
|---|---|
| `gov_data.py` | רץ בכל ריצת `market-data`: `bonds_latest.csv` (42 סדרות + `key`), `curves_latest.csv` (12 טווחים: שקלי/ריאלי/BEI/פורוורד/roll-down), `curves_history.csv`, `us_latest.csv` (FRED), `boi_curves.csv`, `boi_expectations.csv`; וקורא למודולים למטה. |
| `gov_history.py` | `bond_history.csv` — לכל סדרה ויום מסחר: פער מול העקום (נ"ב), מח"מ, מחזור. `vol.csv` — תנודתיות לכל טווח: הרצפה מההגדרות עד 60 ימי מסחר, אח"כ יומי. |
| `gov_cpi.py` | `cpi_path.csv` — תוואי המדדים (ידוע/צפוי) + `season_pct` כדי שהאקסל יבנה תוואי לכל תחזית שנתית. רץ אחרי `short/cpi_expected.py`. |
| `gov_foreign.py` | `foreign_latest.csv` — ארה"ב (FRED), גוש האירו (ECB, עקום AAA), בריטניה (BoE) + ריבית שנה למטבע; כשל → הערך האחרון עם `stale=1`. |
| `boi_policy_manual.csv` | **ידני עם מקור**: ריבית ב"י, ההחלטה הבאה, תחזית חטיבת המחקר. מועתק ל-`boi_policy.csv` עם `il_1y`. |
| `settings_default.json` | כל ספי ההחלטה (מפתח, ערך, תווית, הסבר, קבוצה). משמש גם את האקסל (01_הגדרות, שם מוגדר `p_<key>`) וגם את הפייתון. |
| `gov_decide.py` | כללי ההחלטה בפייתון (מראה של הנוסחאות): משקל צמוד, מח"מ לאפיק, סינתטי/בולט, דירוג סדרות, חו"ל. כותב `decisions_latest.csv` (עם "למה") ו-`decisions_history.csv` (שורה ליום — מקור "אתמול" בדשבורד). |
| `pq_gov.py` | ערכת Power Query: שאילתה לכל `gov_*.csv` ב-`C:\CORP\03_Market\data` (DataMashup על בסיס `corp/pq_kit`). |
| `build_gov2.py` | בונה את `ממשלתי.xlsx`: 13 גיליוני עבודה + 12 גיליוני נתונים. `--settings <קובץ קודם>` מעביר קלטים ידניים (v1 או v2). `--no-pq` לבדיקות. |
| `check_gov2.py` | חישוב ב-LibreOffice, 0 שגיאות, זהות אקסל↔פייתון (עקומים, החלטות, מבנים, מובילות, חו"ל). `python gov/check_gov2.py <xlsx> data` |
| `VERSION` | גרסת הבנייה. `corp/sync_market_to_corp.sh` בונה את הקובץ מחדש רק כשהיא שונה מ-`corp-data/00_Gov/_build/VERSION`. |
| `build_gov.py` | v1 (עיצוב משותף; נשאר כתלות). |

## זרימה יומית
1. `market-data.yml` (GitHub Actions, ב׳–ו׳): גלובס → `data/` → `gov_data.py` → `short/*` → `gov_cpi.py` + `gov_decide.py`.
2. משימת הענן 23:10 (`sync_market_to_corp.sh`): אם נתוני היום חסרים — dispatch והמתנה; מעתיקה את כל `data/gov/*.csv` ל-`corp-data/03_Market/data/gov_*.csv`; בונה את הקובץ רק בשינוי גרסה.
3. בדיקת בוקר 06:41 (ג׳–ש׳): אותו סקריפט במצב בוקר (יום המסחר הקודם; בלי dispatch).
4. המחשב של יובב מסנכרן את corp-data כל 30 דקות; פתיחת הקובץ מרעננת את Power Query.

## שינוי פרמטר / כלל
- פרמטר: לשנות ב-01_הגדרות (באקסל, נשמר דרך `--settings`) או את ברירת המחדל ב-`settings_default.json`.
- כלל: לשנות **גם** בנוסחה (`build_gov2.py`) **וגם** בפייתון (`gov_decide.py`), להעלות את `VERSION`, ולהריץ `pytest -q tests` + `check_gov2.py`.

## בדיקות
`python -m pytest -q tests` (LibreOffice נדרש לבדיקות הבנייה; נדלגות בלעדיו). ב-CI: `.github/workflows/tests.yml`.
