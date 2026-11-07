"""Probe 3: drive the historical-data page filter (period / date range / paging) and record the POST bodies the UI sends."""
import json, sys, time, pathlib
from playwright.sync_api import sync_playwright
OUT = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else "probe_out"); OUT.mkdir(parents=True, exist_ok=True)
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36"
log = []
def note(**k): log.append(k); print(json.dumps(k, ensure_ascii=False)[:700], flush=True)
with sync_playwright() as p:
    b = p.chromium.launch(headless=True)
    ctx = b.new_context(locale="he-IL", timezone_id="Asia/Jerusalem", user_agent=UA, viewport={"width": 1400, "height": 1000})
    page = ctx.new_page()
    def on_req(r):
        if "/history" in r.url and r.method == "POST":
            note(kind="req", url=r.url, post=r.post_data)
    def on_resp(r):
        if "/history" in r.url and r.request.method == "POST":
            try: d = r.json(); note(kind="resp", n=len(d), first=d[0]["tradeDate"] if d else None, last=d[-1]["tradeDate"] if d else None)
            except Exception as e: note(kind="resp_err", err=str(e)[:100])
    page.on("request", on_req); page.on("response", on_resp)
    page.goto("https://maya.tase.co.il/he/funds/mutual-funds/5123138/historical-data", wait_until="networkidle", timeout=90000)
    page.wait_for_timeout(3000)
    note(kind="html_filter", html=page.evaluate("() => { const e=[...document.querySelectorAll('*')].find(x=>x.children.length<6 && /סינון/.test(x.textContent||'') && (x.textContent||'').length<400); return e? e.parentElement.outerHTML.slice(0,4000):'' }"))
    for label in ("שנה", "חצי שנה", "רבעון", "חודש"):
        try:
            page.get_by_text(label, exact=True).first.click(timeout=5000); page.wait_for_timeout(2500)
            note(kind="clicked", label=label)
        except Exception as e: note(kind="click_err", label=label, err=str(e)[:150])
    # paging
    try:
        note(kind="pager", html=page.evaluate("() => { const e=document.querySelector('[class*=pagin],[class*=Pagin],nav'); return e? e.outerHTML.slice(0,1500):'none' }"))
    except Exception as e: note(kind="pager_err", err=str(e)[:100])
    page.screenshot(path=str(OUT / "page.png"), full_page=True)
    b.close()
(OUT / "log.json").write_text(json.dumps(log, ensure_ascii=False, indent=1), encoding="utf-8")
