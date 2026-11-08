"""Probe: Maya mutual-fund LIST api — capture the calls behind /he/funds/mutual-funds (filters, classification tree,
list rows with classification / exposure profile / AUM / fee), and try pulling the full list."""
import json, sys, time, pathlib
from playwright.sync_api import sync_playwright
OUT = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else "probe_out"); OUT.mkdir(parents=True, exist_ok=True)
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36"
B = "https://maya.tase.co.il"; log = []; n = [0]
def note(**k): log.append(k); print(json.dumps(k, ensure_ascii=False)[:600], flush=True)
with sync_playwright() as p:
    b = p.chromium.launch(headless=True)
    ctx = b.new_context(locale="he-IL", timezone_id="Asia/Jerusalem", user_agent=UA, viewport={"width": 1400, "height": 1000})
    page = ctx.new_page()
    def on_resp(r):
        u = r.url
        if "tase.co.il/api" not in u: return
        rec = dict(kind="resp", status=r.status, method=r.request.method, url=u)
        if r.request.method != "GET": rec["post"] = (r.request.post_data or "")[:2000]
        try:
            body = r.text(); n[0] += 1; fn = f"r_{n[0]:03d}.json"; (OUT / fn).write_text(body[:2000000], encoding="utf-8")
            rec.update(saved=fn, len=len(body), head=body[:400])
        except Exception as e: rec["err"] = str(e)[:200]
        note(**rec)
    page.on("response", on_resp)
    page.goto(B + "/he/funds/mutual-funds", wait_until="networkidle", timeout=90000); page.wait_for_timeout(4000)
    for _ in range(3): page.mouse.wheel(0, 3000); page.wait_for_timeout(1500)
    # try "show more"/next page buttons
    for sel in ("text=הצג עוד", "text=טען עוד", ".table_num >> text=2", "button:has-text('הבא')"):
        try: page.locator(sel).first.click(timeout=3000); page.wait_for_timeout(2500); note(kind="clicked", sel=sel)
        except Exception as e: note(kind="click_err", sel=sel, err=str(e)[:80])
    def api(path, body=None):
        return page.evaluate("""async ([url, body]) => { const o = body ? {method:'POST', headers:{'Content-Type':'application/json','Accept':'application/json'}, body: JSON.stringify(body)} : {headers:{'Accept':'application/json'}};
            const r = await fetch(url, o); const t = await r.text(); return {status:r.status, text:t}; }""", [B + path, body])
    for path, body in [("/api/v1/funds/metadata/mutual", None), ("/api/v1/funds/mutual/5123138", None),
                       ("/api/v1/funds/mutual/search", {"pageNumber": 1, "pageSize": 50}), ("/api/v1/funds/mutual", {"pageNumber": 1, "pageSize": 50}),
                       ("/api/v1/funds/mutual/list", {"pageNumber": 1, "pageSize": 50})]:
        try:
            res = api(path, body); n[0] += 1; fn = f"t_{n[0]:03d}.json"; (OUT / fn).write_text(res["text"][:2000000], encoding="utf-8")
            note(kind="test", path=path, body=body, status=res["status"], len=len(res["text"]), saved=fn, head=res["text"][:400]); time.sleep(1)
        except Exception as e: note(kind="test_err", path=path, err=str(e)[:200])
    b.close()
(OUT / "log.json").write_text(json.dumps(log, ensure_ascii=False, indent=1), encoding="utf-8")
