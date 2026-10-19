# CORP daily run — runbook for Claude's scheduled task (18:00 Israel time)

Goal: every evening, pick up the reports the CORP issuers published on Maya since the last run,
file them in the issuer folders on Yovav's PC, update the issuer Excel models (as drafts),
rebuild the master, and send one Hebrew Telegram message. **Days with nothing new: send nothing.**

Standing rules (from the CORP project — `claude/CORP-principles.md`, `claude/CORP-process.md`):
- Models are built only from the company's own reports. Any non-report datum is marked with a URL.
  Never invent or estimate a number; derived figures are flagged as derived.
- Models stay **drafts** until Yovav approves. Never delete his files; before overwriting a model,
  copy the previous version to `<issuer>\_archive\<file>_<YYYY-MM-DD>.xlsx`.
- All user-facing text in Hebrew.

PC: `C:\CORP` (issuers under `C:\CORP\02_Issuers\<corp_folder>`; master `C:\CORP\00_Master_vNN.xlsx`;
template `C:\CORP\01_Template`; per-issuer build scripts in `<issuer>\_build`).
Repo: `yovav81/corp-market-data` (bot code in `maya/`, master builder in `corp/`,
spreads in `data/spreads/<year>.csv`). Only REST `gh api` works from the session (no GraphQL,
no branch deletes, no artifact downloads) — results travel through `maya-out/*` branches and
`git clone -b <branch>`; branches are deleted by dispatching `maya-cleanup`.

## 1. Fetch
1. `git clone --depth 1 https://github.com/yovav81/corp-market-data /home/claude/cmd`
2. Dispatch: `gh api -X POST repos/yovav81/corp-market-data/actions/workflows/maya-watch.yml/dispatches -f ref=main`
   Find the run (`.../workflows/maya-watch.yml/runs?per_page=1`, created after the dispatch) and poll
   `.../actions/runs/<id>` every 20 s until `completed` (normally 3–6 min, give up after 30 min →
   Telegram "בדיקת מאיה נכשלה היום" and stop).
3. List branches `gh api repos/yovav81/corp-market-data/branches?per_page=100` → every `maya-out/watch-*`
   (today's plus any left from a day the PC was unreachable). For each: `git clone --depth 1 -b <branch> ...`
   and read `summary.json`. `issuers[].kept` = new relevant reports; `dir` = folder with the files.
4. If no branch has kept reports → dispatch `maya-cleanup` with those branches, **send nothing**, stop.

## 2. Deliver to the PC
1. Check the PC: `device_list_dir("C:\\CORP\\02_Issuers")`. If unreachable (retry once): send Telegram
   "<b>CORP אג״ח</b> | התקבלו N דיווחים חדשים (…); המחשב לא זמין — הקבצים ממתינים ויועברו בריצה הבאה",
   **do not** clean up the branches, stop.
2. Per issuer with kept reports:
   - `device_list_dir(<issuer>\\דוחות, recursive=true)` → save JSON →
     `python maya/prep_existing.py listing.json <branch>/<dir> <issuer root>` → stage those candidates,
     `md5sum` them → `hashes.txt` ("md5  name").
   - `python maya/sort_to_corp.py <branch>/<dir> <staging> --existing hashes.txt` (CORP naming + folders).
   - Drop `*_שער-מגנא.htm` and `*_טקסט.txt` from staging (low value); mark them in the manifest.
   - `device_commit_files` in batches (≤45 files, ≤90 MB per call) from `/mnt/user-data/outputs/...`.
   - Append the manifest rows to `<issuer>\_log\maya_manifest.csv` and a row to `<issuer>\_log\updates.csv`
     (stage the existing CSVs first; append, never rewrite history).

## 3. Update the issuer model (draft)
Stage the model (`<corp_folder>_*.xlsx`, newest) and archive a copy first. By report type:

| form / reason | what to update | how |
|---|---|---|
| ת125 / ד125 דירוג | `סדרות` rating cols (R/S…), `דירוג` sheet, `פרופיל` B44 events | read the rating PDF; per series rating + outlook + date; source note |
| ת126 מצבת התחייבויות | `סדרות` J (ע.נ. במחזור) per ISIN, notes; `חוב ונזילות` schedule if the model reads it | parse the htm table (cp1255) |
| ת077 מצבת בעלי עניין | `פרופיל` holdings block (row 55+) and B11 controller % | parse htm/pdf; exact % with report ref |
| ת020/ת022/ת128/ת026, הרחבת סדרה | `סדרות` nominal/terms notes (AG), `קובננטים` if new covenants | quote the document; mark drafts as drafts |
| ת930 דוח רבעוני/שנתי | full quarter: P&L, CF, BS, debt, covenants (`רבעונים`, `מאזן`, `רווח והפסד`, `תזרים`, `קובננטים`, `סדרות`) | follow the issuer's `_build` scripts as the pattern (fin.json → fill → finish); G City / AFI also have MAGNA iXBRL (see invest-automation `components/magna_ixbrl.py`) |
| מצגת / אחר / משפטי / מימון | `פרופיל` B44 events + `מקורות` | one line + source |

After any change: `python /mnt/skills/public/xlsx/scripts/recalc.py <model> 180` must report 0 errors, and
the `בקרות` sheet must have no critical failures. **If either fails, or a quarterly report cannot be
completed with confidence, do not overwrite the model** — keep the files filed and report
"נדרשת בדיקה" with the reason in the Telegram message. Register every document in `מקורות` and log
in `יומן`.

## 4. Master
If any model changed: stage all issuer models (newest per folder) + `C:\CORP\00_Master` contents;
`python corp/build_master.py <dir_with_models> data/spreads/<year>.csv 00_Master_vNN.xlsx corp/sectors.csv`
(NN = highest existing + 1), recalc (0 errors), commit to `C:\CORP\00_Master_vNN.xlsx`.

## 5. Telegram (one message per run, only if something new)
Dispatch `corp-notify` with input `text` (HTML). Format:

```
<b>CORP אג״ח | דיווחים חדשים — 19.10.2026</b>

<b>ביג</b>
• 18.10 דוח רבעון 3/2026 — <a href="https://maya.tase.co.il/he/reports/NNN">מאיה</a>
• 18.10 מעלות: אישור דירוג ilAA- יציב
✅ האקסל עודכן (טיוטה): רבעון 3/2026, דירוג; בקרות תקינות
— או — ⚠️ נדרשת בדיקה: <סיבה>

📊 המאסטר עודכן: v19
```
Keep it short; company names in Hebrew; no file paths.

## 6. Clean up
Dispatch `maya-cleanup` with input `branches` = the branches you delivered (space-separated).
Leave undelivered branches in place.
