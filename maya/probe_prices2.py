"""Probe 2: parameters of POST /api/v1/funds/mutual/{id}/history (page size, date range) and speed."""
import json, sys, time, pathlib
from playwright.sync_api import sync_playwright
sys.path.insert(0, str(pathlib.Path(__file__).parent))
from fetch import Maya
OUT = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else "probe_out"); OUT.mkdir(parents=True, exist_ok=True)
res = []
def raw(m, path, body):
    r = m.page.evaluate("""async ([url, body]) => { const r = await fetch(url, {method:'POST', headers:{'Content-Type':'application/json','Accept':'application/json'}, body: JSON.stringify(body)});
        const t = await r.text(); return {status:r.status, text:t}; }""", ["https://maya.tase.co.il" + path, body])
    return r
with sync_playwright() as p:
    m = Maya(p)
    m.page.goto("https://maya.tase.co.il/he/funds/mutual-funds/5123138/historical-data", wait_until="networkidle", timeout=90000)
    P = "/api/v1/funds/mutual/5123138/history"
    tests = [{"pageSize": 20, "pageNumber": 1}, {"pageSize": 500, "pageNumber": 1}, {"pageSize": 100, "pageNumber": 2},
             {"pageSize": 100, "pageNumber": 1, "fromDate": "2026-01-01", "toDate": "2026-03-31"},
             {"pageSize": 100, "pageNumber": 1, "dateFrom": "2026-01-01", "dateTo": "2026-03-31"},
             {"pageSize": 100, "pageNumber": 1, "period": 5}, {"pageSize": 100, "pageNumber": 1, "period": 1},
             {"pageSize": 100, "pageNumber": 1, "fromDate": "2025-01-01", "toDate": "2025-01-31", "period": 0}]
    for b in tests:
        t0 = time.time(); r = raw(m, P, b)
        try: d = json.loads(r["text"]); n = len(d); first = d[0]["tradeDate"][:10] if n else None; last = d[-1]["tradeDate"][:10] if n else None
        except Exception: n = first = last = None
        rec = dict(body=b, status=r["status"], n=n, first=first, last=last, sec=round(time.time() - t0, 2), head=r["text"][:150])
        res.append(rec); print(json.dumps(rec, ensure_ascii=False), flush=True); time.sleep(0.7)
    # other funds: speed + ids of different kinds
    for fid in (5128590, 5113271, 5124078):
        r = raw(m, f"/api/v1/funds/mutual/{fid}/history", {"pageSize": 20, "pageNumber": 1})
        rec = dict(fund=fid, status=r["status"], head=r["text"][:200]); res.append(rec); print(json.dumps(rec, ensure_ascii=False), flush=True)
    m.b.close()
(OUT / "log.json").write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
