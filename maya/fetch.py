"""Maya (TASE) report fetcher for the CORP bond models.

Usage:
    python maya/fetch.py --company "ביג" --out maya_out [--from 2023-10-01] [--mode download|catalog]

* --company: a Maya company id (e.g. 1327) or a name; names are resolved with Maya's own autocomplete.
* Lists every report of the company from --from (default: 3 years back) to today,
  writes catalog.csv (all reports + the keep/skip decision and the reason),
  and in download mode saves every attachment of the kept reports.
* All API calls run inside a real browser page (same origin as the Maya site), slowly.
"""
import argparse, csv, datetime as dt, json, re, sys, time, pathlib
from playwright.sync_api import sync_playwright

B = "https://maya.tase.co.il"
FILES = "https://mayafiles.tase.co.il/"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
      "Chrome/130.0.0.0 Safari/537.36")
MAX_FILE = 95 * 1024 * 1024      # GitHub rejects files over 100 MB

# ---- relevance rules (catalog.csv records which rule decided) -------------------------
KEEP_FORMS = {
    "ת128": "שטר נאמנות", "ת020": "תוצאות הנפקה", "ת022": "דוח הצעת מדף", "ת026": "טיוטת דוח הצעה",
    "ת125": "דירוג", "ת126": "מצבת התחייבויות", "ת077": "מצבת בעלי עניין", "ת012": "תשקיף", "ת011": "תשקיף",
}
SKIP_FORMS = {
    "ת076": "שינוי החזקות", "ת079": "חדל להיות בעל עניין", "ת087": "מרשם / מצבת הון",
    "ת081": "בעל עניין", "ת091": "נושא משרה", "ת093": "נושא משרה", "ת097": "נושא משרה",
    "ת304": "נושא משרה", "ת306": "נושא משרה",
}
KEEP_KW = [
    ("דוח תקופתי|דוח רבעוני|דוח שנתי|דוחות כספיים|דו\"ח תקופתי|דו\"ח רבעוני|דוח חצי שנתי|רבעון", "דוח כספי"),
    ("מצבת התחייבויות", "מצבת התחייבויות"),
    ("תשקיף|הצעת מדף|דוח הצעה|דו\"ח הצעה|שטר נאמנות|שטרי נאמנות|הרחבת סדרה|הרחבת סדרת|הצעה פרטית|תוצאות (?:ה)?(?:מכרז|הנפקה)", "הנפקה / שטר"),
    ("דירוג|מעלות|מידרוג|S&P", "דירוג"),
    ("מצגת", "מצגת"),
    ("שומ[הת]|הערכת שווי|הערכות שווי", "שומה"),
    ("מחזיקי (?:אגרות|אג\"ח|האג\"ח)|אסיפת מחזיקי|דוח נאמן|דו\"ח נאמן", "מחזיקי אג\"ח"),
    ("פדיון מוקדם|פירעון מוקדם|רכישה עצמית|רכישת אגרות|הצעת רכש|הצעת חליפין|רכש חליפין", "פעולות בחוב"),
    ("הסדר|חדלות|הקפאת הליכים|שינוי שליטה|העברת שליטה|בעל השליטה|עסקה מהותית|אירוע מהותי", "אירוע מהותי"),
    ("מימון|הלוואה|אשראי|שעבוד|קובננט|אמות מידה", "מימון"),
    ("מכירת|רכישת|מימוש נכס|השקעה ב", "עסקאות"),
    ("דיבידנד|חלוקה", "חלוקה"),
    ("החזקות בעלי עניין.*ליום|מצבת החזקות", "מצבת בעלי עניין"),
]
SKIP_KW = [
    ("פתיחת מסחר|אישור מנכ\"?ל הבורסה|פרטים מזהים|שינוי החזקות|חדל להיות|מינוי|סיום כהונה|"
     "תוצאות אסיפה כללית|זימון אסיפה כללית|אסיפה כללית (?:שנתית|מיוחדת)|הקצאת אופציות|מימוש אופציות|"
     "שינוי שם|שינוי כתובת|תשלום ריבית|תשלום קרן|בעל זכות חתימה", "שוטף"),
]

def decide(rep):
    form, title = rep.get("formId") or "", rep.get("title") or ""
    for rx, why in SKIP_KW:
        if re.search(rx, title) and not re.search("מחזיקי (?:אגרות|אג\"ח)", title):
            return False, "דלג: " + why
    if not form:
        return False, "דלג: הודעת בורסה"
    if form in SKIP_FORMS:
        return False, "דלג: " + SKIP_FORMS[form]
    if form in KEEP_FORMS:
        return True, KEEP_FORMS[form]
    for rx, why in KEEP_KW:
        if re.search(rx, title):
            return True, why
    return False, "דלג: לא תואם"

def clean(s, n=70):
    s = re.sub(r'[\\/:*?"<>|\r\n\t]+', " ", s or "").strip()
    s = re.sub(r"\s+", "-", s)
    return s[:n].strip("-.")

class Maya:
    def __init__(self, p):
        self.b = p.chromium.launch(headless=True)
        self.ctx = self.b.new_context(locale="he-IL", timezone_id="Asia/Jerusalem", user_agent=UA)
        self.page = self.ctx.new_page()
        self.warm()

    def warm(self, url=B + "/he/reports/companies"):
        self.page.goto(url, wait_until="networkidle", timeout=90000)
        self.page.wait_for_timeout(3000)

    def api(self, path, body=None, tries=4):
        for i in range(tries):
            res = self.page.evaluate(
                """async ([url, body]) => {
                     const o = body ? {method:'POST', headers:{'Content-Type':'application/json','Accept':'application/json'}, body: JSON.stringify(body)}
                                    : {headers:{'Accept':'application/json'}};
                     const r = await fetch(url, o); const t = await r.text();
                     return {status: r.status, text: t};
                   }""", [B + path, body])
            time.sleep(0.8)
            if res["status"] == 200:
                try:
                    return json.loads(res["text"])
                except Exception:
                    pass
            print(f"  api {path} -> {res['status']} ({res['text'][:80]!r}); retry {i+1}", flush=True)
            time.sleep(10 * (i + 1)); self.warm()
        raise RuntimeError(f"API failed: {path} {body}")

    def file(self, url, tries=3):
        for i in range(tries):
            try:
                r = self.ctx.request.get(FILES + url, timeout=180000)
                if r.ok:
                    return r.body()
                print(f"  file {url} -> {r.status}", flush=True)
            except Exception as e:
                print(f"  file {url} error {e}", flush=True)
            time.sleep(8 * (i + 1))
        return None

def resolve(m, company):
    if company.strip().isdigit() and len(company.strip()) <= 5:
        return int(company), None
    from urllib.parse import quote
    res = m.api(f"/api/v1/companies/autocomplete?search={quote(company.strip())}&take=10")
    comps = [r for r in res if r.get("type") in ("COMPANY", "CANDIDATE")]
    if not comps:
        raise SystemExit(f"no Maya company matches {company!r}: {res}")
    exact = [r for r in comps if company.strip() in (r.get("value") or "")]
    pick = (exact or comps)[0]
    return int(pick["key"]), {"picked": pick, "candidates": comps}

def month_windows(start, end):
    d = dt.date(start.year, start.month, 1)
    while d <= end:
        nxt = dt.date(d.year + (d.month == 12), d.month % 12 + 1, 1)
        yield max(d, start), min(nxt - dt.timedelta(days=1), end)
        d = nxt

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--company", required=True)
    ap.add_argument("--out", default="maya_out")
    ap.add_argument("--from", dest="frm", default="")
    ap.add_argument("--mode", default="download", choices=["download", "catalog"])
    ap.add_argument("--all", action="store_true", help="download every report, not only the relevant ones")
    a = ap.parse_args()
    today = dt.date.today()
    start = dt.date.fromisoformat(a.frm) if a.frm else today.replace(year=today.year - 3, day=1)
    out = pathlib.Path(a.out)
    with sync_playwright() as p:
        m = Maya(p)
        cid, how = resolve(m, a.company)
        details = m.api(f"/api/v1/companies/{cid}/details")
        name = (details.get("longName") or details.get("name") or str(cid)) if isinstance(details, dict) else str(cid)
        root = out / f"{cid}_{clean(name, 40)}"
        root.mkdir(parents=True, exist_ok=True)
        (root / "company.json").write_text(json.dumps({"companyId": cid, "resolve": how, "details": details},
                                                      ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"company {cid} {name}; from {start} to {today}", flush=True)
        reps, seen = [], set()
        for f, t in month_windows(start, today):
            off = 0
            while True:
                page = m.api("/api/v1/reports/companies", {"pageNumber": 1, "limit": 30, "companyId": cid, "offset": off,
                                                           "fromDate": f.isoformat(), "toDate": t.isoformat()})
                for r in page:
                    if r["id"] not in seen:
                        seen.add(r["id"]); reps.append(r)
                if len(page) < 30:
                    break
                off += 30
            print(f"  {f:%Y-%m}: {len(reps)} reports so far", flush=True)
        reps.sort(key=lambda r: r["publishDate"])
        rows, nbytes = [], 0
        for r in reps:
            keep, why = decide(r)
            keep = keep or a.all
            atts = [x for x in (r.get("attachments") or []) if x.get("url") and x.get("fileType") != "edit"]
            saved = []
            if keep and a.mode == "download":
                date = r["publishDate"][:10]
                for k, x in enumerate(atts):
                    ext = x["url"].rsplit(".", 1)[-1].lower()
                    fn = f"{date}_{r['id']}_{(r.get('formId') or 'NA')}_{clean(r['title'], 60)}" + (f"__{k}" if len(atts) > 1 else "") + f".{ext}"
                    data = m.file(x["url"])
                    if data is None:
                        saved.append("FAILED"); continue
                    if len(data) > MAX_FILE:
                        saved.append(f"TOO_BIG:{len(data)}"); continue
                    (root / fn).write_bytes(data); nbytes += len(data); saved.append(fn)
                    time.sleep(0.4)
            rows.append({"report_id": r["id"], "date": r["publishDate"][:19], "form": r.get("formId") or "",
                         "title": r.get("title") or "", "keep": "כן" if keep else "לא", "reason": why,
                         "attachments": " | ".join(f"{x.get('fileType')}:{x.get('fileName') or ''}:{x['url']}" for x in atts),
                         "saved": " | ".join(saved),
                         "maya_url": f"{B}/he/reports/{r['id']}"})
        with open(root / "catalog.csv", "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.DictWriter(fh, fieldnames=list(rows[0].keys()) if rows else ["report_id"])
            w.writeheader(); w.writerows(rows)
        kept = sum(1 for x in rows if x["keep"] == "כן")
        print(f"done: {len(rows)} reports, {kept} kept, {nbytes/1e6:.1f} MB downloaded -> {root}", flush=True)
        m.b.close()

if __name__ == "__main__":
    main()
