# -*- coding: utf-8 -*-
"""monthly_reports.py — דוחות האחזקות החודשיים של כל מנהלי הקרנות (מאיה, טופס ק203 "דוח חודשי").

    python funds/monthly_reports.py data [--days 45] [--force] [--managers 10083,10031]

לכל מנהל קרנות (הרשימה ממאיה: /api/v1/funds/metadata/mutual → fundsManagers) נלקח הדוח החודשי האחרון
(אירוע 5618 = דוח חודשי), קובץ ה-TXT שלו מורד מ-mayafiles, מנורמל ב-rashut.py (אותו קוד שמשרת את קובץ הקרן)
ונכתב:
    data/funds/state.json                      — לכל מנהל: דיווח אחרון שנקלט (id, חודש, תאריך פרסום)
    data/funds/managers.csv                    — רשימת המנהלים (key, value)
    data/funds/<YYYY-MM>/holdings.csv.gz       — כל האחזקות של כל הקרנות לחודש (מנורמל, 26 עמודות)
    data/funds/<YYYY-MM>/funds.csv             — שורה לכל קרן: נתונים מדווחים (שווי נקי, חשיפות, מח״מ) + סיווג
    data/funds/<YYYY-MM>/sources.csv           — איזה דיווח/קובץ של איזה מנהל נכנס
    data/funds/latest_holdings.csv.gz, latest_funds.csv, latest.json  — החודש השלם האחרון (כל המנהלים שדיווחו)
    data/funds/index_members.csv               — חברות בת״א 35/90 מקרנות הסל העוקבות (מתעדכן כשיש)
הקבצים הגולמיים (TXT, ~16MB למנהל) לא נשמרים — ניתנים להורדה חוזרת ממאיה.
חודש נחשב "שלם" כשכל המנהלים הפעילים דיווחו עליו, או כשעברו 75 יום מסוף החודש.
"""
import argparse, csv, datetime as dt, gzip, io, json, os, pathlib, re, sys, time
import pandas as pd
from playwright.sync_api import sync_playwright

HERE = pathlib.Path(__file__).parent
sys.path.insert(0, str(HERE)); sys.path.insert(0, str(HERE.parent / "maya"))
import rashut                                   # noqa: E402
from fetch import Maya, B                       # noqa: E402

EVENT_MONTHLY = 5618
HEB_MONTHS = {"ינואר": 1, "פברואר": 2, "מרץ": 3, "מרס": 3, "אפריל": 4, "מאי": 5, "יוני": 6, "יולי": 7, "אוגוסט": 8,
              "ספטמבר": 9, "אוקטובר": 10, "נובמבר": 11, "דצמבר": 12}


def load_json(p, default):
    return json.loads(p.read_text(encoding="utf-8")) if p.exists() else default


def month_from_title(title):
    m = re.search(r"(" + "|".join(HEB_MONTHS) + r")\s*(20\d\d)", title or "")
    if m:
        return f"{m.group(2)}-{HEB_MONTHS[m.group(1)]:02d}"
    m = re.search(r"(\d{1,2})[/.](20\d\d)", title or "")
    return f"{m.group(2)}-{int(m.group(1)):02d}" if m else None


def list_monthly(m, frm, to):
    out, off = [], 0
    while True:
        page = m.api("/api/v1/reports/mutual-funds", {"pageNumber": 1, "limit": 50, "offset": off, "eventsIds": [EVENT_MONTHLY],
                                                      "fromDate": frm.isoformat(), "toDate": to.isoformat()})
        page = [r for r in page if (r.get("formId") == "ק203") or ("דוח חודשי" in (r.get("title") or ""))]
        out.extend(page)
        if len(page) < 50 or off > 2000:
            return out
        off += 50


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("data", nargs="?", default="data")
    ap.add_argument("--days", type=int, default=45, help="חלון חיפוש אחורה לדיווחים (ימים)")
    ap.add_argument("--force", action="store_true", help="לקלוט מחדש גם דיווחים שכבר נקלטו")
    ap.add_argument("--managers", default="", help="רק מנהלים אלה (מזהי מאיה מופרדים בפסיק)")
    a = ap.parse_args()
    data = pathlib.Path(a.data)
    fdir = data / "funds"; fdir.mkdir(parents=True, exist_ok=True)
    state_p = fdir / "state.json"
    state = load_json(state_p, {"managers": {}, "months": {}})
    today = dt.date.today()
    frm = today - dt.timedelta(days=a.days)
    mkt = rashut.Market(str(data.parent) if data.name == "data" else str(data), None)

    with sync_playwright() as p:
        m = Maya(p)
        meta = m.api("/api/v1/funds/metadata/mutual")
        managers = {int(x["key"]): x["value"] for x in meta.get("fundsManagers", [])}
        with open(fdir / "managers.csv", "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f); w.writerow(["manager_id", "name"]); [w.writerow([k, v]) for k, v in sorted(managers.items())]
        print(f"{len(managers)} מנהלי קרנות; חלון {frm}..{today}", flush=True)
        reps = list_monthly(m, frm, today)
        print(f"{len(reps)} דיווחי 'דוח חודשי' בחלון", flush=True)
        # הדיווח האחרון לכל מנהל ולכל חודש (מנהל יכול לפרסם תיקון → לוקחים את המאוחר)
        latest = {}
        for r in reps:
            mid = r.get("reporterId") or (r.get("companies") or [{}])[0].get("companyId")
            mon = month_from_title(r.get("title"))
            if not mid or not mon:
                continue
            key = (int(mid), mon)
            if key not in latest or r["publishDate"] > latest[key]["publishDate"]:
                latest[key] = r
        only = {int(x) for x in a.managers.split(",") if x.strip().isdigit()}
        new_rows = []
        for (mid, mon), r in sorted(latest.items(), key=lambda kv: (kv[0][1], kv[0][0] != 10083, kv[0][0])):   # מיטב ראשון: קרנות הסל העוקבות שלו נותנות את חברות המדדים
            if only and mid not in only:
                continue
            st = state["managers"].setdefault(str(mid), {})
            if not a.force and st.get(mon, {}).get("report_id") == r["id"]:
                continue
            att = [x for x in (r.get("attachments") or []) if (x.get("fileType") or "").startswith("txt") or (x.get("url") or "").endswith(".txt")]
            if not att:
                print(f"  {managers.get(mid, mid)} {mon}: דיווח {r['id']} ללא קובץ TXT — מדלג", flush=True)
                continue
            url = att[0]["url"]
            print(f"  {managers.get(mid, mid)} {mon}: דיווח {r['id']} → {url}", flush=True)
            body = m.file(url)
            if not body:
                print("    הורדה נכשלה", flush=True); continue
            mdir = fdir / mon; mdir.mkdir(exist_ok=True)
            tmp = mdir / f"_raw_{mid}.txt"
            tmp.write_bytes(body)
            try:
                h, fdf, members = rashut.normalize(str(tmp), str(data.parent) if data.name == "data" else str(data), None,
                                                   str(fdir / "index_members.csv"), mkt=mkt)
            except Exception as e:
                print(f"    נרמול נכשל: {e}", flush=True); tmp.unlink(missing_ok=True); continue
            tmp.unlink(missing_ok=True)
            h.insert(0, "manager_id", mid); fdf.insert(0, "manager_id", mid)
            fdf["מנהל_מאיה"] = managers.get(mid, "")
            h = h.astype(object).where(h.notna(), ""); fdf = fdf.astype(object).where(fdf.notna(), "")
            # צבירה לקובץ החודש: מחליפים את השורות של המנהל הזה
            hp, fp = mdir / "holdings.csv.gz", mdir / "funds.csv"
            if hp.exists():
                old = pd.read_csv(hp, dtype=str, keep_default_na=False)
                old = old[old["manager_id"] != str(mid)]
                h = pd.concat([old, h.astype(str)], ignore_index=True)
            if fp.exists():
                oldf = pd.read_csv(fp, dtype=str, keep_default_na=False)
                oldf = oldf[oldf["manager_id"] != str(mid)]
                fdf = pd.concat([oldf, fdf.astype(str)], ignore_index=True)
            with gzip.open(hp, "wt", encoding="utf-8", newline="") as g:
                h.to_csv(g, index=False)
            fdf.to_csv(fp, index=False, encoding="utf-8-sig")
            st[mon] = {"report_id": r["id"], "published": r["publishDate"][:10], "file": url, "funds": int(fdf[fdf["manager_id"].astype(str) == str(mid)].shape[0]),
                       "rows": int(h[h["manager_id"].astype(str) == str(mid)].shape[0])}
            new_rows.append([mon, mid, managers.get(mid, ""), r["id"], r["publishDate"][:10], url, st[mon]["funds"], st[mon]["rows"]])
            time.sleep(2)
        m.b.close()

    # sources.csv לכל חודש שנגע + בחירת "החודש השלם האחרון"
    for mon in sorted({row[0] for row in new_rows}):
        sp = fdir / mon / "sources.csv"
        rows = [["month", "manager_id", "manager", "report_id", "published", "file", "funds", "rows"]]
        for mid, st in state["managers"].items():
            if mon in st:
                rows.append([mon, mid, managers.get(int(mid), ""), st[mon]["report_id"], st[mon]["published"], st[mon]["file"], st[mon]["funds"], st[mon]["rows"]])
        with open(sp, "w", encoding="utf-8-sig", newline="") as f:
            csv.writer(f).writerows(rows)
    months = sorted({mon for st in state["managers"].values() for mon in st})
    complete = None
    for mon in months:
        reported = {mid for mid, st in state["managers"].items() if mon in st}
        y, mo = int(mon[:4]), int(mon[5:])
        month_end = (dt.date(y + (mo == 12), mo % 12 + 1, 1) - dt.timedelta(days=1))
        if len(reported) >= len(managers) * 0.8 or (today - month_end).days >= 75:
            complete = mon
        state["months"][mon] = {"managers_reported": len(reported), "managers_total": len(managers)}
    if complete:
        import shutil
        shutil.copy(fdir / complete / "holdings.csv.gz", fdir / "latest_holdings.csv.gz")
        shutil.copy(fdir / complete / "funds.csv", fdir / "latest_funds.csv")
        (fdir / "latest.json").write_text(json.dumps({"month": complete, **state["months"][complete], "updated": today.isoformat()}, ensure_ascii=False), encoding="utf-8")
    state_p.write_text(json.dumps(state, ensure_ascii=False, indent=1), encoding="utf-8")
    print("חדש:", len(new_rows), "| חודש שלם אחרון:", complete, flush=True)
    for row in new_rows:
        print("  ", row[:5], row[6], "קרנות", row[7], "שורות")


if __name__ == "__main__":
    main()
