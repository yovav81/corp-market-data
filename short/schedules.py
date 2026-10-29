"""
schedules.py — לוחות סילוקין (תזרים פירעונות) לאג"ח קצרות מתוך דף "נתוני אג"ח" של ביזפורטל.

לכל סדרה קונצרנית שמח"מ גלובס שלה מתחת לסף (ברירת מחדל 1.3, כדי שסדרות שיורדות מתחת ל-1 יהיו מוכנות)
נמשך הדף https://www.bizportal.co.il/bonds/quote/bondsdata/<ני"ע> ונקראים:
  * פרטי אג"ח: סוג הפדיון, סוג הצמדה, שיעור ריבית, סוג ריבית, תדירות ריבית, תשלום ריבית/קרן הבא, תשלום קרן אחרון, תאריך אקס
  * תזרים פירעונות: מועד תשלום, תשלום קרן, תשלום ריבית (סכומים כספיים — מנורמלים אחר כך ל-100 ע.נ. נותרת)

פלט:
  data/short/bond_terms.csv   שורה לסדרה (תנאים + fetched_at + status)
  data/short/schedules.csv    שורה לכל תשלום עתידי (security_id, pay_date, principal_amt, interest_amt, principal_pct, interest_pct)

לוח שנמשך לפני פחות מ-REFRESH_DAYS ימים ושמועד התשלום הקרוב שלו לא עבר — לא נמשך שוב.
הרצה:  python short/schedules.py data            (רגיל)
        python short/schedules.py data --ids 1158732,1141050 --force --debug debug/short
"""
from __future__ import annotations
import argparse, datetime as dt, io, json, os, re, sys, time, urllib.request
from pathlib import Path
import pandas as pd

BASE = "https://www.bizportal.co.il/bonds/quote/bondsdata/{sid}"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36",
      "Accept-Language": "he-IL,he;q=0.9,en;q=0.8"}
DUR_MAX = float(os.environ.get("SHORT_DUR_MAX", "1.3"))
REFRESH_DAYS = int(os.environ.get("SHORT_REFRESH_DAYS", "7"))
TERM_COLS = ["security_id", "name", "redemption_type", "linkage", "coupon", "coupon_type", "coupon_freq", "next_interest",
             "next_ex", "next_principal", "last_principal", "collateral", "fetched_at", "status", "n_flows", "source_url"]
SCHED_COLS = ["security_id", "pay_date", "principal_amt", "interest_amt", "principal_pct", "interest_pct"]

LABELS = {  # תווית בדף -> שדה
    "סוג הפדיון": "redemption_type", "סוג הצמדה": "linkage", "שיעור ריבית": "coupon", "סוג ריבית": "coupon_type",
    "תדירות ריבית": "coupon_freq", "תשלום ריבית הבא": "next_interest", "תאריך אקס": "next_ex",
    "תשלום קרן הבא": "next_principal", "תשלום קרן אחרון": "last_principal", "סוג שעבוד": "collateral",
}


def get(url: str, timeout=60) -> str:
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read().decode("utf-8", errors="replace")


def _num(s):
    if s is None: return None
    s = str(s).replace(",", "").replace("%", "").replace("₪", "").strip()
    if s in ("", "--", "-", "nan", "None"): return None
    try: return float(s)
    except ValueError: return None


def _date(s):
    if not s: return None
    m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", str(s))
    return dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1))) if m else None


def parse(html: str) -> tuple[dict, list[dict]]:
    """מחזיר (תנאים, שורות תזרים). קריאה סובלנית: תוויות לפי טקסט, טבלת התזרים לפי כותרות."""
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html, "lxml")
    terms = {}
    # 1) תוויות: כל אלמנט שהטקסט שלו הוא בדיוק התווית — הערך הוא הטקסט של האח/התא הבא
    for lab, key in LABELS.items():
        for el in soup.find_all(string=re.compile(r"^\s*" + re.escape(lab) + r"\s*:?\s*$")):
            node = el.parent
            val = None
            # תא טבלה: התא הבא בשורה
            if node.name in ("td", "th"):
                sib = node.find_next_sibling(["td", "th"])
                val = sib.get_text(" ", strip=True) if sib else None
            if not val:
                sib = node.find_next_sibling()
                val = sib.get_text(" ", strip=True) if sib else None
            if not val:
                nxt = el.find_next(string=True)
                val = str(nxt).strip() if nxt else None
            if val:
                terms[key] = val; break
    text = soup.get_text(" ", strip=True)
    if "linkage" not in terms:
        m = re.search(r"סוג הצמדה\s*:?\s*(\S+)", text)
        if m: terms["linkage"] = m.group(1)
    # 2) טבלת התזרים
    rows = []
    try:
        tables = pd.read_html(io.StringIO(html))
    except Exception:
        tables = []
    for t in tables:
        cols = [str(c) for c in t.columns]
        if any("מועד" in c and "תשלום" in c for c in cols) and any("קרן" in c for c in cols):
            cmap = {}
            for c in cols:
                if "מועד" in c: cmap["pay_date"] = c
                elif "תשלום קרן" in c or c.strip() == "קרן": cmap["principal_amt"] = c
                elif "תשלום ריבית" in c: cmap["interest_amt"] = c
                elif "% קרן" in c or "%קרן" in c: cmap["principal_pct"] = c
                elif "% ריבית" in c or "%ריבית" in c: cmap["interest_pct"] = c
            for _, r in t.iterrows():
                d = _date(r[cmap["pay_date"]]) if "pay_date" in cmap else None
                if not d: continue
                rows.append({"pay_date": d,
                             "principal_amt": _num(r.get(cmap.get("principal_amt"))) if "principal_amt" in cmap else None,
                             "interest_amt": _num(r.get(cmap.get("interest_amt"))) if "interest_amt" in cmap else None,
                             "principal_pct": _num(r.get(cmap.get("principal_pct"))) if "principal_pct" in cmap else None,
                             "interest_pct": _num(r.get(cmap.get("interest_pct"))) if "interest_pct" in cmap else None})
            if rows: break
    # ניקוי ערכים
    out = {}
    for k, v in terms.items():
        if k in ("coupon",): out[k] = _num(v)
        elif k in ("next_interest", "next_ex", "next_principal", "last_principal"): out[k] = _date(v)
        else: out[k] = v
    return out, rows


def load_prev(d: Path):
    tp, sp = d / "bond_terms.csv", d / "schedules.csv"
    terms = pd.read_csv(tp, encoding="utf-8-sig") if tp.exists() else pd.DataFrame(columns=TERM_COLS)
    sched = pd.read_csv(sp, encoding="utf-8-sig") if sp.exists() else pd.DataFrame(columns=SCHED_COLS)
    return terms, sched


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("data", nargs="?", default="data")
    ap.add_argument("--ids", help="רשימת ני\"ע מופרדת בפסיקים (במקום הבחירה לפי מח\"מ)")
    ap.add_argument("--force", action="store_true", help="למשוך מחדש גם לוחות טריים")
    ap.add_argument("--debug", help="תיקייה לשמירת ה-HTML הגולמי של דפים שנכשלו/נדגמו")
    ap.add_argument("--sleep", type=float, default=0.7)
    a = ap.parse_args()
    data = Path(a.data); outd = data / "short"; outd.mkdir(parents=True, exist_ok=True)
    mk = pd.read_csv(data / "latest.csv", encoding="utf-8-sig")
    mk = mk[mk["source"].astype(str).str.startswith("corp")].copy()
    mk["security_id"] = pd.to_numeric(mk["security_id"], errors="coerce")
    mk = mk[mk["security_id"].notna()]
    if a.ids:
        want = mk[mk["security_id"].isin([int(x) for x in a.ids.split(",")])]
        if want.empty:
            want = pd.DataFrame({"security_id": [int(x) for x in a.ids.split(",")], "name": ""})
    else:
        want = mk[pd.to_numeric(mk["duration"], errors="coerce") < DUR_MAX]
    prev_t, prev_s = load_prev(outd)
    prev_t["security_id"] = pd.to_numeric(prev_t["security_id"], errors="coerce")
    today = dt.date.today(); now = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    keep_t = []; keep_s = []; fetched = 0; failed = 0
    debug = Path(a.debug) if a.debug else None
    if debug: debug.mkdir(parents=True, exist_ok=True)
    for _, r in want.iterrows():
        sid = int(r["security_id"]); name = r.get("name", "")
        old = prev_t[prev_t["security_id"] == sid]
        fresh = False
        if not old.empty and not a.force:
            o = old.iloc[0]
            try:
                age = (today - dt.datetime.strptime(str(o["fetched_at"])[:10], "%Y-%m-%d").date()).days
            except Exception:
                age = 999
            nxt = _date(o.get("next_principal")) or _date(o.get("next_interest"))
            fresh = o.get("status") == "ok" and age < REFRESH_DAYS and (nxt is None or nxt >= today)
        if fresh:
            keep_t.append(old.iloc[0].to_dict())
            keep_s += prev_s[pd.to_numeric(prev_s["security_id"], errors="coerce") == sid].to_dict("records")
            continue
        url = BASE.format(sid=sid); status = "ok"; terms, rows = {}, []
        try:
            html = get(url); fetched += 1
            terms, rows = parse(html)
            if not rows:
                status = "no_schedule"
                if debug: (debug / f"{sid}.html").write_text(html, encoding="utf-8")
            elif debug and fetched <= 3:
                (debug / f"{sid}.html").write_text(html, encoding="utf-8")
        except Exception as e:
            status = f"error: {str(e)[:80]}"
        if status != "ok": failed += 1
        rec = {"security_id": sid, "name": name, **{k: terms.get(k) for k in TERM_COLS if k in LABELS.values()},
               "fetched_at": now, "status": status, "n_flows": len(rows), "source_url": url}
        keep_t.append({k: rec.get(k) for k in TERM_COLS})
        for x in rows:
            keep_s.append({"security_id": sid, **x})
        print(f"{sid} {name}: {status}, {len(rows)} תשלומים, הצמדה={terms.get('linkage')} ריבית={terms.get('coupon')}")
        time.sleep(a.sleep)
    t = pd.DataFrame(keep_t, columns=TERM_COLS); s = pd.DataFrame(keep_s, columns=SCHED_COLS)
    t.to_csv(outd / "bond_terms.csv", index=False, encoding="utf-8-sig")
    s.to_csv(outd / "schedules.csv", index=False, encoding="utf-8-sig")
    print(f"סדרות: {len(t)} (נמשכו {fetched}, נכשלו {failed}); תשלומים: {len(s)}")
    if fetched and failed == fetched:
        print("::error::כל המשיכות נכשלו — ייתכן שמבנה הדף בביזפורטל השתנה"); sys.exit(1)


if __name__ == "__main__":
    main()
