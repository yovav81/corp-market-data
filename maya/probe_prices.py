"""Probe: Maya mutual-fund historical-data page — capture the API behind the daily buy/redemption price table."""
import json, sys, time, pathlib
from playwright.sync_api import sync_playwright
OUT = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else "probe_out"); OUT.mkdir(parents=True, exist_ok=True)
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36"
B = "https://maya.tase.co.il"
log = []; n = [0]
def note(**k): log.append(k); print(json.dumps(k, ensure_ascii=False)[:500], flush=True)
with sync_playwright() as p:
    b = p.chromium.launch(headless=True)
    ctx = b.new_context(locale="he-IL", timezone_id="Asia/Jerusalem", user_agent=UA)
    page = ctx.new_page()
    def on_resp(r):
        u = r.url
        if "tase.co.il" not in u or "/api" not in u: return
        rec = dict(kind="resp", status=r.status, method=r.request.method, url=u)
        if r.request.method != "GET": rec["post"] = (r.request.post_data or "")[:1500]
        try:
            body = r.text(); n[0] += 1; fn = f"r_{n[0]:03d}.json"
            (OUT / fn).write_text(body[:400000], encoding="utf-8"); rec.update(saved=fn, len=len(body), head=body[:300])
        except Exception as e: rec["err"] = str(e)[:200]
        note(**rec)
    page.on("response", on_resp)
    for u in (B + "/he/funds/mutual-funds/5123138/historical-data", B + "/he/funds/mutual-funds/5123138"):
        try:
            page.goto(u, wait_until="networkidle", timeout=90000); page.wait_for_timeout(5000)
            page.mouse.wheel(0, 2000); page.wait_for_timeout(2000)
            note(kind="page", url=u, title=page.title(), text=page.inner_text("body")[:1500])
        except Exception as e: note(kind="page_err", url=u, err=str(e)[:300])
    b.close()
(OUT / "log.json").write_text(json.dumps(log, ensure_ascii=False, indent=1), encoding="utf-8")
