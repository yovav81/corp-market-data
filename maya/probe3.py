"""Probe 3: company-name search endpoint + report-list paging/filter limits."""
import json, sys, pathlib, requests
from playwright.sync_api import sync_playwright
OUT = pathlib.Path(sys.argv[1]); OUT.mkdir(parents=True, exist_ok=True)
log = []
def note(**k): log.append(k); print(json.dumps(k, ensure_ascii=False)[:300], flush=True)
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
      "Chrome/130.0.0.0 Safari/537.36")
B = "https://maya.tase.co.il"
n = [0]
with sync_playwright() as p:
    b = p.chromium.launch(headless=True)
    ctx = b.new_context(locale="he-IL", timezone_id="Asia/Jerusalem", user_agent=UA, viewport={"width": 1400, "height": 1000})
    page = ctx.new_page()
    def on_resp(r):
        u = r.url
        if "/api/" not in u or "tase.co.il" not in u: return
        rec = dict(kind="resp", status=r.status, method=r.request.method, url=u)
        if r.request.method != "GET": rec["post"] = (r.request.post_data or "")[:3000]
        try:
            body = r.text(); n[0] += 1; fn = f"r_{n[0]:03d}.json"
            (OUT / fn).write_text(body[:300000], encoding="utf-8"); rec["saved"] = fn; rec["len"] = len(body)
        except Exception as e: rec["err"] = str(e)[:200]
        note(**rec)
    page.on("response", on_resp)
    for u in [B + "/he/reports/companies"]:
        page.goto(u, wait_until="networkidle", timeout=60000); page.wait_for_timeout(3000)
        try: page.get_by_text("אל תציגו שוב").click(timeout=3000)
        except Exception: pass
        (OUT / "companies_page.html").write_text(page.content(), encoding="utf-8")
        page.screenshot(path=str(OUT / "companies_page.png"))
        inputs = page.locator("input")
        note(kind="inputs", count=inputs.count(),
             ph=[inputs.nth(i).get_attribute("placeholder") for i in range(min(inputs.count(), 15))],
             aria=[inputs.nth(i).get_attribute("aria-label") for i in range(min(inputs.count(), 15))])
        typed = False
        for sel in ["input[placeholder*='חברה']", "[placeholder*='חברה']", "input[aria-label*='חברה']",
                    "text=בחר חברה"]:
            try:
                el = page.locator(sel).first
                el.click(timeout=5000); page.keyboard.type("ביג", delay=150); page.wait_for_timeout(4000)
                page.screenshot(path=str(OUT / "typed.png")); typed = True
                note(kind="typed", sel=sel); break
            except Exception as e:
                note(kind="type_fail", sel=sel, err=str(e)[:150])
        # header search box
        try:
            page.goto(B + "/he", wait_until="networkidle", timeout=60000); page.wait_for_timeout(2000)
            page.locator("button[aria-label*='חיפוש'], [class*='search'] button, button:has(svg)").first.click(timeout=5000)
            page.wait_for_timeout(1500); page.keyboard.type("לייטסטון", delay=150); page.wait_for_timeout(4000)
            page.screenshot(path=str(OUT / "header_search.png")); note(kind="header_search_ok")
        except Exception as e: note(kind="header_search_fail", err=str(e)[:200])
    b.close()
# paging / filters on the company report list, plain requests
S = requests.Session(); S.headers.update({"User-Agent": UA, "Accept": "application/json", "Content-Type": "application/json",
                                          "Referer": B + "/he/companies/1327/reports", "Origin": B})
tests = [
    {"pageNumber": 1, "limit": 100, "companyId": 1327, "offset": 0},
    {"pageNumber": 1, "limit": 30, "companyId": 1327, "offset": 0},
    {"pageNumber": 1, "limit": 20, "companyId": 1327, "offset": 1000},
    {"pageNumber": 1, "limit": 20, "companyId": 1327, "offset": 0, "fromDate": "2023-10-01", "toDate": "2023-12-31"},
    {"pageNumber": 1, "limit": 20, "companyId": 1327, "offset": 0, "formIds": ["ת930"]},
]
for i, body in enumerate(tests):
    try:
        r = S.post(B + "/api/v1/reports/companies", json=body, timeout=60)
        txt = r.text; (OUT / f"t_{i}.json").write_text(txt[:300000], encoding="utf-8")
        cnt = None; first = last = None
        try:
            j = r.json()
            if isinstance(j, list) and j: cnt = len(j); first = j[0]["publishDate"]; last = j[-1]["publishDate"]
        except Exception: pass
        note(kind="list", body=body, status=r.status_code, n=cnt, first=first, last=last, head=txt[:150])
    except Exception as e: note(kind="list_err", body=body, err=str(e)[:200])
for rid in (1715903, 1726427, 1741167):
    try:
        r = S.get(B + f"/api/v1/reports/{rid}", timeout=60); j = r.json()
        note(kind="rep", id=rid, reporter=j.get("reporterId"), name=[c["name"] for c in j.get("companies", [])], form=j.get("formId"))
    except Exception as e: note(kind="rep_err", id=rid, err=str(e)[:200])
(OUT / "log.json").write_text(json.dumps(log, ensure_ascii=False, indent=1), encoding="utf-8")
