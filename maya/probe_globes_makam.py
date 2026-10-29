"""probe_globes_makam.py <out_dir> — בדיקה: מה ה-handler של גלובס (/portal/Handlers/GTOQuotes2022.ashx) מחזיר למק"מ במצבים שונים."""
import json, sys
from pathlib import Path

out = Path(sys.argv[1]); out.mkdir(parents=True, exist_ok=True)
MODES = ["unlinkedShortGovernmentBonds", "shortTermTreasuryBills", "madad800", "unlinkedFixedRateGovernmentBonds", "cpiLinkedGovernmentBonds"]
GROUPS = ["", "Goverment"]
from playwright.sync_api import sync_playwright
res = {}
with sync_playwright() as p:
    b = p.chromium.launch(headless=True)
    ctx = b.new_context(locale="he-IL", timezone_id="Asia/Jerusalem")
    page = ctx.new_page()
    page.goto("https://www.globes.co.il/portal/quotes/unlinked-short-government-bonds", wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(5000)
    for mode in MODES:
        for g in GROUPS:
            key = f"{mode}|{g}"
            try:
                r = ctx.request.post("https://www.globes.co.il/portal/Handlers/GTOQuotes2022.ashx",
                                     form={"sortby": "HebName", "direction": "asc", "showAll": "true", "searchValue": "", "mode": mode,
                                           "groupType": g, "isMobile": "", "gls": "", "feeder": "0", "test": ""},
                                     headers={"X-Requested-With": "XMLHttpRequest", "Referer": "https://www.globes.co.il/portal/quotes/unlinked-short-government-bonds"})
                txt = r.text()
                res[key] = {"status": r.status, "len": len(txt), "head": txt[:1500]}
                (out / f"{mode}_{g or 'none'}.txt").write_text(txt, encoding="utf-8")
            except Exception as e:
                res[key] = {"error": str(e)}
    # גם: איזה בקשות הדף עצמו שלח
    reqs = []
    page.on("request", lambda rq: reqs.append({"url": rq.url, "method": rq.method, "post": (rq.post_data or "")[:300]}) if "GTOQuotes" in rq.url or "Handlers" in rq.url else None)
    page.reload(wait_until="domcontentloaded"); page.wait_for_timeout(15000)
    res["_page_requests"] = reqs
    b.close()
json.dump(res, open(out / "makam_probe.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps({k: (v.get("status"), v.get("len"), (v.get("head") or "")[:120]) if isinstance(v, dict) and "status" in v else v for k, v in res.items()}, ensure_ascii=False)[:4000])
