"""Probe: can GitHub runners reach Maya (TASE)? Records every tase.co.il network call
made by the Maya web app, saves JSON bodies, and tries direct file downloads."""
import json, os, sys, time, pathlib, traceback
import requests
from playwright.sync_api import sync_playwright

OUT = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else "probe_out")
OUT.mkdir(parents=True, exist_ok=True)
log = []

def note(**k):
    k["t"] = time.strftime("%H:%M:%S"); log.append(k); print(json.dumps(k, ensure_ascii=False)[:400], flush=True)

PAGES = [
    "https://maya.tase.co.il/he",
    "https://maya.tase.co.il/he/reports/1716282",   # Big interested parties 31.12.2025
    "https://maya.tase.co.il/he/reports/1706324",   # G City / Orion
]
FILES = [  # direct file URLs (classic mayafiles layout)
    "https://mayafiles.tase.co.il/rpdf/1603001-1604000/P1603854-00.pdf",
    "https://mayafiles.tase.co.il/rpdf/1698001-1699000/P1698351-00.pdf",
]
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36")

with sync_playwright() as p:
    b = p.chromium.launch(headless=True)
    ctx = b.new_context(locale="he-IL", timezone_id="Asia/Jerusalem", user_agent=UA,
                        accept_downloads=True)
    page = ctx.new_page()
    n = [0]
    def on_resp(r):
        u = r.url
        if "tase.co.il" not in u: return
        ct = r.headers.get("content-type", "")
        rec = dict(kind="resp", status=r.status, method=r.request.method, url=u, ct=ct)
        if r.request.method == "POST":
            rec["post"] = (r.request.post_data or "")[:2000]
        if "json" in ct:
            try:
                body = r.text(); n[0] += 1
                fn = OUT / f"json_{n[0]:03d}.json"; fn.write_text(body[:200000], encoding="utf-8")
                rec["saved"] = fn.name; rec["len"] = len(body)
            except Exception as e:
                rec["err"] = str(e)
        note(**rec)
    page.on("response", on_resp)
    for i, u in enumerate(PAGES):
        try:
            r = page.goto(u, wait_until="networkidle", timeout=60000)
            page.wait_for_timeout(4000)
            note(kind="page", url=u, status=r.status if r else None, title=page.title())
            (OUT / f"page_{i}.html").write_text(page.content(), encoding="utf-8")
            page.screenshot(path=str(OUT / f"page_{i}.png"), full_page=False)
        except Exception as e:
            note(kind="page_err", url=u, err=str(e)[:500])
    # file downloads through the browser context (shares cookies)
    for u in FILES:
        try:
            r = ctx.request.get(u, timeout=60000)
            body = r.body()
            note(kind="ctx_file", url=u, status=r.status, ct=r.headers.get("content-type"), len=len(body),
                 head=body[:8].decode("latin1", "replace"))
            if r.ok and body[:4] == b"%PDF":
                (OUT / u.rsplit("/", 1)[1]).write_bytes(body)
        except Exception as e:
            note(kind="ctx_file_err", url=u, err=str(e)[:300])
    cookies = ctx.cookies()
    note(kind="cookies", names=[c["name"] for c in cookies])
    b.close()

# plain requests (no browser) — tells us if a light script would work
for u in FILES + ["https://maya.tase.co.il/he/reports/1716282"]:
    try:
        r = requests.get(u, headers={"User-Agent": UA, "Accept-Language": "he-IL"}, timeout=60)
        note(kind="plain", url=u, status=r.status_code, ct=r.headers.get("content-type"), len=len(r.content),
             head=r.content[:8].decode("latin1", "replace"))
    except Exception as e:
        note(kind="plain_err", url=u, err=str(e)[:300])

(OUT / "log.json").write_text(json.dumps(log, ensure_ascii=False, indent=1), encoding="utf-8")
print("done", len(log))
