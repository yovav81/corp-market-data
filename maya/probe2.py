"""Probe 2: discover company/report-list APIs of the new Maya site."""
import json, sys, time, pathlib, requests
from playwright.sync_api import sync_playwright
OUT = pathlib.Path(sys.argv[1]); OUT.mkdir(parents=True, exist_ok=True)
log = []
def note(**k): log.append(k); print(json.dumps(k, ensure_ascii=False)[:300], flush=True)
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
      "Chrome/130.0.0.0 Safari/537.36")
B = "https://maya.tase.co.il"
PAGES = [B + "/he/companies/1327", B + "/he/companies/1327/reports", B + "/he/reports"]
API = [
    "/api/v1/reports/1716282/siblings?offset=0&limit=50",
    "/api/v1/reports/1716282/siblings?offset=0&limit=500",
    "/api/v1/companies/1327",
    "/api/v1/companies/1327/reports?offset=0&limit=50",
    "/api/v1/companies?search=%D7%91%D7%99%D7%92",
    "/api/v1/companies/search?term=%D7%91%D7%99%D7%92",
    "/api/v1/companies/autocomplete?term=%D7%91%D7%99%D7%92",
    "/api/v1/autocomplete?term=%D7%91%D7%99%D7%92",
    "/api/v1/forms",
    "/api/v1/reports/forms",
]
n = [0]
with sync_playwright() as p:
    b = p.chromium.launch(headless=True)
    ctx = b.new_context(locale="he-IL", timezone_id="Asia/Jerusalem", user_agent=UA)
    page = ctx.new_page()
    def on_resp(r):
        u = r.url
        if "maya.tase.co.il/api" not in u and "mayaapi" not in u: return
        rec = dict(kind="resp", status=r.status, method=r.request.method, url=u)
        if r.request.method != "GET": rec["post"] = (r.request.post_data or "")[:3000]
        try:
            body = r.text(); n[0] += 1; fn = f"r_{n[0]:03d}.json"
            (OUT / fn).write_text(body[:400000], encoding="utf-8"); rec["saved"] = fn; rec["len"] = len(body)
        except Exception as e: rec["err"] = str(e)[:200]
        note(**rec)
    page.on("response", on_resp)
    for i, u in enumerate(PAGES):
        try:
            page.goto(u, wait_until="networkidle", timeout=60000); page.wait_for_timeout(3000)
            for _ in range(3):
                page.mouse.wheel(0, 4000); page.wait_for_timeout(1500)
            note(kind="page", url=u, title=page.title())
            (OUT / f"page_{i}.html").write_text(page.content(), encoding="utf-8")
            page.screenshot(path=str(OUT / f"page_{i}.png"), full_page=True)
        except Exception as e: note(kind="page_err", url=u, err=str(e)[:300])
    # try typing a company into the search form on /he/reports
    try:
        inp = page.locator("input").first
        inp.click(); inp.type("ביג", delay=120); page.wait_for_timeout(4000)
        page.screenshot(path=str(OUT / "search_typed.png"))
    except Exception as e: note(kind="search_err", err=str(e)[:300])
    for a in API:
        try:
            r = ctx.request.get(B + a, timeout=60000); body = r.text(); n[0] += 1; fn = f"a_{n[0]:03d}.json"
            (OUT / fn).write_text(body[:400000], encoding="utf-8")
            note(kind="ctx_api", url=a, status=r.status, len=len(body), saved=fn, head=body[:120])
        except Exception as e: note(kind="ctx_api_err", url=a, err=str(e)[:200])
    b.close()
for a in API[:3]:
    try:
        r = requests.get(B + a, headers={"User-Agent": UA, "Accept": "application/json"}, timeout=60)
        note(kind="plain_api", url=a, status=r.status_code, len=len(r.content), head=r.text[:120])
    except Exception as e: note(kind="plain_api_err", url=a, err=str(e)[:200])
(OUT / "log.json").write_text(json.dumps(log, ensure_ascii=False, indent=1), encoding="utf-8")
