# CORP daily run — runbook for Claude's scheduled task (17:57 Israel time, cloud only)

Goal: every evening, pick up the reports the CORP issuers published on Maya since the last run, file them,
update the issuer Excel models (as drafts), rebuild the master, and send one Hebrew Telegram message.
**Days with nothing new (and no pending merge): send nothing.** Yovav's PC is NOT needed: the models live in the
private repo `yovav81/corp-data`, and his PC syncs with it on its own (`_sync/corp_sync.ps1`, see its README).

Standing rules (from the CORP project — `claude/CORP-principles.md`, `claude/CORP-process.md`):
- Models are built only from the company's own reports. Any non-report datum is marked with a URL.
  Never invent or estimate a number; derived figures are flagged as derived.
- Models stay **drafts** until Yovav approves. Never delete his files; before overwriting a model,
  copy the previous version to `<issuer>/_archive/<file>_<YYYY-MM-DD>.xlsx`.
- **Never overwrite a change Yovav made by hand**: always start from the latest `corp-data` main, and if the push
  races with one of his, rebase and redo your edit on top of his version (step 6).
- All user-facing text in Hebrew.

Repos (clone both; if a clone or `gh api` call is refused, attach the repo with the add_repo tool —
owner `yovav81`, access `push` — and retry once):
- `yovav81/corp-market-data` → `/home/claude/cmd`: bot code `maya/`, master builder `corp/build_live.py` (+ `build_master.py`, `short_sheet.py`),
  sectors `corp/sectors.csv`, merge tool `corp/merge_xlsx.py`, spreads `data/spreads/<year>.csv`, workflows.
- `yovav81/corp-data` → `/home/claude/cd` (full clone, not shallow): the CORP folder tree —
  `00_Master_LIVE.xlsx` (the single master — no `00_Master_vNN` versions since 29.10.2026), `00_Master/` (build scripts, `פרמטרים_רוחביים.xlsx`), `01_Template/`, `_archive/` (old versions — ignore), `02_Issuers/<corp_folder>/` (model `*.xlsx`, `_log/`, `_archive/`, `_build/`),
  `_sync/` (PC sync tool; conflict records in `_sync/conflicts/`). No report PDFs here.
Only REST `gh api` works for workflows (no GraphQL, no branch deletes, no artifact downloads) — bot results
travel through `maya-out/*` branches of corp-market-data; those are deleted by dispatching `maya-cleanup`.

```
git clone --depth 1 https://github.com/yovav81/corp-market-data /home/claude/cmd
git clone https://github.com/yovav81/corp-data /home/claude/cd
git -C /home/claude/cd config user.name "CORP bot"; git -C /home/claude/cd config user.email corp-bot@users.noreply.github.com
git -C /home/claude/cd config core.quotepath false
```

## 1. Fetch new Maya reports
1. Dispatch: `gh api -X POST repos/yovav81/corp-market-data/actions/workflows/maya-watch.yml/dispatches -f ref=main`
   Find the run (`.../workflows/maya-watch.yml/runs?per_page=1`, created after the dispatch) and poll
   `.../actions/runs/<id>` every 20 s until `completed` (normally 3–6 min; give up after 30 min →
   Telegram "בדיקת מאיה נכשלה היום" and continue with step 2 only).
2. List branches `gh api repos/yovav81/corp-market-data/branches?per_page=100` → every `maya-out/watch-*`
   (today's plus any leftovers). For each: `git clone --depth 1 -b <branch> https://github.com/yovav81/corp-market-data <dir>`
   and read `summary.json`. `issuers[].kept` = new relevant reports; `dir` = folder with the files.

## 2. Merge Yovav's conflicting edits (every run, even with no new reports)
For each `cd/_sync/conflicts/*.json` with `"resolved": null` (the PC wrote it when a file changed both on the PC and
in the cloud; his version is at `local_copy`, the cloud version stayed at `path`):
1. `git -C cd cat-file blob <base_blob> > base.xlsx` (if `base_blob` is null, use the commit before the first
   cloud change of `path` after `time`: `git -C cd log --format=%H -- <path>`).
2. `python cmd/corp/merge_xlsx.py base.xlsx <cd/local_copy> <cd/path> merged.xlsx --report rep.json`
3. Recalc (`python /mnt/skills/public/xlsx/scripts/recalc.py merged.xlsx 180`, 0 errors) and check `בקרות`.
   OK → archive the current `path` to `_archive/`, replace `path` with merged.xlsx, log in the model's `יומן`.
   Conflicting cells keep the cloud value; list them. Not OK → leave both files, report "נדרשת בדיקה".
4. Set `"resolved"` in the json to `{"time": ..., "applied": N, "conflicts": [cells...], "result": "merged"|"review"}`.
   Do not delete the `_מקומי` copy (Yovav moves it to DELETE himself; the PC sync then removes it from the cloud).
If a model is a non-xlsx file (csv/py/md), merge by hand the same way (his lines win where the cloud did not touch them).

## 3. File the new reports
For each issuer with kept reports (`corp_folder` from summary.json):
1. `python cmd/maya/index_hashes.py cd/02_Issuers/<corp_folder>/_log/reports_index.csv > hashes.txt`
   (the PC writes that index: every file already in its דוחות folder, with md5).
2. `python cmd/maya/sort_to_corp.py <branch>/<dir> stg/<corp_folder> --existing hashes.txt` (CORP naming + folders).
3. Drop `*_שער-מגנא.htm` and `*_טקסט.txt` from `stg/<corp_folder>/דוחות` (low value; mark "לא נשמר" in the manifest).
4. Move `stg/<corp_folder>/דוחות` to `deliver_root/02_Issuers/<corp_folder>/דוחות`.
5. Append the manifest rows to `cd/02_Issuers/<corp_folder>/_log/maya_manifest.csv` and one row to
   `_log/updates.csv` (append, never rewrite history; keep the existing header/encoding: utf-8-sig).
Then push all issuers' files at once:
`python cmd/maya/push_deliver.py deliver_root https://github.com/yovav81/corp-data <YYYY-MM-DD>-<watch run id>`
The PC copies `deliver/*` into its דוחות folders when it is on and deletes the branches — never delete them here.
Keep `deliver_root` for step 4: the models are updated from these files.

## 4. Update the issuer model (draft) in `cd/02_Issuers/<corp_folder>/`
Take the newest `<corp_folder>_*.xlsx` (not `_מקומי`), archive a copy first. By report type:

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
"נדרשת בדיקה" with the reason. Register every document in `מקורות` and log in `יומן`.

## 5. Master
If any model changed, rebuild the single live master **in place** (never create `00_Master_vNN.xlsx`):
`python3 cd/00_Master/_build/build_live.py /home/claude/cd /home/claude/cmd --check`
It rebuilds `cd/00_Master_LIVE.xlsx` from all models in `cd/02_Issuers` + `cmd/data` (spreads, hist, short, factors),
adds the `בדיקות_לדוח_הבא` sheet, keeps Yovav's Watchlist / parameters / index overrides from the existing file,
and re-injects the three Power Query queries. `--check` recalcs a copy in LibreOffice and must report 0 errors
(needs `pandas openpyxl numpy` and `soffice`). Do not open or edit `00_Master_LIVE.xlsx` with openpyxl afterwards —
that would strip the queries; any change goes through the build scripts.

## 6. Commit and push corp-data
`git -C cd add -A && git -C cd commit -m "<Hebrew summary>"` then `git -C cd push origin main`.
If the push is rejected (Yovav's PC pushed meanwhile — it syncs on every save and every 30 min):
1. `git -C cd fetch origin`; mine = `git -C cd diff --name-only origin/main...HEAD`, his = `git -C cd diff --name-only HEAD...origin/main`.
2. No common file → `git -C cd rebase origin/main` and push.
3. Common file(s) → copy all of *your* changed files to a temp dir, `git -C cd reset --hard origin/main`, copy back
   the files he did not touch, and **redo** your edit for each common file on top of his new version (re-run the
   step-4 update; for CSV logs re-append your rows). Commit, push. Never discard his commit.

## 7. Telegram (one message per run, only if something new or merged)
Dispatch `corp-notify` with input `text` (HTML). Link the updated model in GitHub:
`https://github.com/yovav81/corp-data/blob/main/<url-encoded path>`. Format:

```
<b>CORP אג״ח | דיווחים חדשים — 19.10.2026</b>

<b>ביג</b>
• 18.10 דוח רבעון 3/2026 — <a href="https://maya.tase.co.il/he/reports/NNN">מאיה</a>
• 18.10 מעלות: אישור דירוג ilAA- יציב
✅ <a href="https://github.com/yovav81/corp-data/blob/main/...">האקסל</a> עודכן (טיוטה): רבעון 3/2026, דירוג; בקרות תקינות
— או — ⚠️ נדרשת בדיקה: <סיבה>

🔀 מוזגו השינויים הידניים שלך ב-<קובץ>: N תאים; בהתנגשות (נשאר ערך הענן): A1, B7
📊 המאסטר עודכן: v20
💻 הקבצים יגיעו למחשב בסנכרון הבא שלו
```
Keep it short; company names in Hebrew; no local paths.

## 8. Clean up
Dispatch `maya-cleanup` (corp-market-data) with input `branches` = the `maya-out/watch-*` branches you filed
(space-separated). Leave any branch you could not process.
