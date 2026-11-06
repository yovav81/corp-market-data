"""Probe 2 for fund reports: capture the Maya API calls behind the mutual-fund pages (fund list, manager list,
fund reports list) and test listing a fund manager's reports by companyId, so the monthly ק203 pipeline
can discover every manager's latest monthly report."""
import json, sys, time, pathlib
from playwright.sync_api import sync_playwright

OUT = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else "probe_out"); OUT.mkdir(parents=True, exist_ok=True)
log = []
def note(**k): log.append(k); print(json.dumps(k, ensure_ascii=False)[:400], flush=True)
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36")
B = "https://maya.tase.co.il"
PAGES = [B + "/he/funds/mutual-funds", B + "/he/funds/mutual-funds/5123138/reports", B + "/he/funds/mutual-funds/5123138"]
n = [0]
with sync_playwright() as p:
    b = p.chromium.launch(headless=True)
    ctx = b.new_context(locale="he-IL", timezone_id="Asia/Jerusalem", user_agent=UA)
    page = ctx.new_page()
    def on_resp(r):
        u = r.url
        if "maya.tase.co.il/api" not in u: return
        rec = dict(kind="resp", status=r.status, method=r.request.method, url=u)
        if r.request.method != "GET": rec["post"] = (r.request.post_data or "")[:3000]
        try:
            body = r.text(); n[0] += 1; fn = f"r_{n[0]:03d}.json"
            (OUT / fn).write_text(body[:600000], encoding="utf-8"); rec["saved"] = fn; rec["len"] = len(body); rec["head"] = body[:200]
        except Exception as e: rec["err"] = str(e)[:200]
        note(**rec)
    page.on("response", on_resp)
    for i, u in enumerate(PAGES):
        try:
            page.goto(u, wait_until="networkidle", timeout=90000); page.wait_for_timeout(4000)
            for _ in range(2):
                page.mouse.wheel(0, 3000); page.wait_for_timeout(1500)
            note(kind="page", url=u, title=page.title())
        except Exception as e: note(kind="page_err", url=u, err=str(e)[:300])
    # direct API experiments from the page context (same origin)
    def api(path, body=None):
        return page.evaluate("""async ([url, body]) => {
             const o = body ? {method:'POST', headers:{'Content-Type':'application/json','Accept':'application/json'}, body: JSON.stringify(body)}
                            : {headers:{'Accept':'application/json'}};
             const r = await fetch(url, o); const t = await r.text(); return {status: r.status, text: t}; }""", [B + path, body])
    tests = [
        ("/api/v1/reports/companies", {"pageNumber": 1, "limit": 30, "companyId": 10083, "offset": 0, "fromDate": "2026-08-01", "toDate": "2026-11-06"}),
        ("/api/v1/reports/mutual-funds", {"pageNumber": 1, "limit": 30, "fundId": 5123138, "offset": 0}),
        ("/api/v1/reports/mutual-funds", {"pageNumber": 1, "limit": 30, "companyId": 10083, "offset": 0, "fromDate": "2026-08-01", "toDate": "2026-11-06"}),
        ("/api/v1/companies/autocomplete?search=%D7%A7%D7%A8%D7%A0%D7%95%D7%AA%20%D7%A0%D7%90%D7%9E%D7%A0%D7%95%D7%AA&take=50", None),
        ("/api/v1/funds/mutual/autocomplete?search=%D7%9E%D7%99%D7%98%D7%91&take=20", None),
        ("/api/v1/funds/managers", None),
        ("/api/v1/funds/mutual-funds/managers", None),
        ("/api/v1/companies/10083", None),
        ("/api/v1/companies/10083/reports?offset=0&limit=30", None),
    ]
    for path, body in tests:
        try:
            res = api(path, body); n[0] += 1; fn = f"t_{n[0]:03d}.json"
            (OUT / fn).write_text(res["text"][:600000], encoding="utf-8")
            note(kind="test", path=path, body=body, status=res["status"], len=len(res["text"]), saved=fn, head=res["text"][:300])
            time.sleep(1)
        except Exception as e: note(kind="test_err", path=path, err=str(e)[:300])
    b.close()
(OUT / "log.json").write_text(json.dumps(log, ensure_ascii=False, indent=1), encoding="utf-8")
print("done", len(log))
