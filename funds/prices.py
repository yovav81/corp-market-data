# -*- coding: utf-8 -*-
"""prices.py — מחירי קרנות נאמנות יומיים (מחיר פדיון/קנייה) ממאיה → data/funds/prices.csv.gz

    python funds/prices.py data [--from 2025-12-20] [--funds 5123138,5128590] [--all-classes] [--limit N]

מקור: POST maya.tase.co.il/api/v1/funds/mutual/<מס' קרן>/history  (pageSize ≤ 30, period=4 = טווח תאריכים fromDate/toDate).
יקום: כל הקרנות ב-data/funds/latest_funds.csv שאינן מניות (אג״ח/מעורב/כספית) + קרנות ב-funds/extra_price_funds.txt (+ --funds).
עדכון מצטבר: לכל קרן מורידים מהתאריך האחרון ששמור פחות 7 ימים; קרן חדשה — מ---from (ברירת מחדל: תחילת הדצמבר הקודמת, לבסיס YTD).
פלט: data/funds/prices.csv.gz  (fund,date,buy,sell,nav_m)  + data/funds/prices_meta.json
"""
import argparse, datetime as dt, gzip, json, pathlib, sys, time
import pandas as pd
from playwright.sync_api import sync_playwright

HERE = pathlib.Path(__file__).parent
sys.path.insert(0, str(HERE.parent / "maya"))
from fetch import Maya                      # noqa: E402

JS = """
async ([ids, fromIso, toIso]) => {
  const out = {};
  async function one(id) {
    const rows = []; let status = 200;
    for (let pg = 1; pg <= 40; pg++) {
      let res = null;
      for (let t = 0; t < 3; t++) {
        const r = await fetch('/api/v1/funds/mutual/' + id + '/history', {method: 'POST',
          headers: {'Content-Type': 'application/json', 'Accept': 'application/json'},
          body: JSON.stringify({pageSize: 30, pageNumber: pg, period: 4, fromDate: fromIso, toDate: toIso})});
        status = r.status;
        if (r.status === 200) { try { res = await r.json(); break; } catch (e) {} }
        if (r.status === 404 || r.status === 400) break;
        await new Promise(s => setTimeout(s, 1500 * (t + 1)));
      }
      if (res === null) return {status, rows};
      for (const x of res) rows.push([x.tradeDate.slice(0, 10), x.purchasePrice, x.sellPrice, x.assetValue]);
      if (res.length < 30) break;
    }
    return {status, rows};
  }
  await Promise.all(ids.map(async id => { out[id] = await one(id); }));
  return out;
}
"""


def iso(d):
    return (d - dt.timedelta(days=1)).strftime("%Y-%m-%dT21:00:00.000Z")   # חצות ישראל (חורף) — עודף של יום מסונן ע״י dedupe


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("data", nargs="?", default="data")
    ap.add_argument("--from", dest="frm", default="")
    ap.add_argument("--funds", default="")
    ap.add_argument("--all-classes", action="store_true", help="כולל קרנות מניות")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--batch", type=int, default=6)
    a = ap.parse_args()
    data = pathlib.Path(a.data); fdir = data / "funds"
    today = dt.date.today()
    start = dt.date.fromisoformat(a.frm) if a.frm else dt.date(today.year - 1, 12, 15)

    ff = pd.read_csv(fdir / "latest_funds.csv", dtype=str, encoding="utf-8-sig")
    ff["מס_קרן"] = ff["מס_קרן"].str.zfill(8)
    uni = ff if a.all_classes else ff[~ff["סיווג"].fillna("").str.startswith("מניות")]
    ids = set(uni["מס_קרן"])
    extra = HERE / "extra_price_funds.txt"
    if extra.exists():
        ids |= {x.strip().zfill(8) for x in extra.read_text().split() if x.strip().isdigit()}
    ids |= {x.strip().zfill(8) for x in a.funds.split(",") if x.strip().isdigit()}
    ids = sorted(ids)
    if a.limit:
        ids = ids[:a.limit]

    pp = fdir / "prices.csv.gz"
    old = pd.read_csv(pp, dtype={"fund": str}) if pp.exists() else pd.DataFrame(columns=["fund", "date", "buy", "sell", "nav_m"])
    last = old.groupby("fund")["date"].max().to_dict() if len(old) else {}
    print(f"{len(ids)} קרנות; קיימות {len(last)} עם מחירים; בסיס חדשות {start}", flush=True)

    new, bad, empty = [], [], []
    t0 = time.time()
    with sync_playwright() as p:
        m = Maya(p)
        m.page.goto("https://maya.tase.co.il/he/funds/mutual-funds/5123138/historical-data", wait_until="networkidle", timeout=90000)
        groups = {}
        for fid in ids:
            lo = dt.date.fromisoformat(last[fid]) - dt.timedelta(days=7) if fid in last else start
            groups.setdefault(lo, []).append(fid)       # קרנות עם אותו תאריך התחלה בקבוצה אחת
        todo = [(lo, fid) for lo, fs in sorted(groups.items()) for fid in fs]
        for i in range(0, len(todo), a.batch):
            chunk = todo[i:i + a.batch]
            res = {}
            # קבוצת fromDate אחת לכל קריאה (ההתחלה המוקדמת ביותר בצ'אנק)
            lo = min(c[0] for c in chunk)
            for attempt in range(3):
                try:
                    res = m.page.evaluate(JS, [[str(int(c[1])) for c in chunk], iso(lo), iso(today + dt.timedelta(days=1))])
                    break
                except Exception as e:
                    print("  שגיאת דפדפן:", str(e)[:120], flush=True); time.sleep(8 * (attempt + 1)); m.warm()
            for _, fid in chunk:
                r = res.get(str(int(fid)))
                if not r or r["status"] not in (200,):
                    bad.append((fid, r and r["status"])); continue
                if not r["rows"]:
                    empty.append(fid); continue
                for d_, buy, sell, nav in r["rows"]:
                    new.append((fid, d_, buy, sell, nav))
            if (i // a.batch) % 25 == 0:
                print(f"  {i + len(chunk)}/{len(todo)} קרנות | {len(new)} שורות חדשות | {int(time.time() - t0)} שנ׳", flush=True)
            time.sleep(0.3)
        m.b.close()

    add = pd.DataFrame(new, columns=["fund", "date", "buy", "sell", "nav_m"])
    allp = pd.concat([old, add], ignore_index=True).drop_duplicates(["fund", "date"], keep="last").sort_values(["fund", "date"])
    with gzip.open(pp, "wt", encoding="utf-8", newline="") as g:
        allp.to_csv(g, index=False)
    meta = {"updated": today.isoformat(), "funds_with_prices": int(allp["fund"].nunique()), "rows": int(len(allp)),
            "last_date": allp["date"].max(), "first_date": allp["date"].min(), "requested": len(ids), "errors": len(bad), "empty": len(empty)}
    (fdir / "prices_meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=1), encoding="utf-8")
    print(json.dumps(meta, ensure_ascii=False), flush=True)
    if bad:
        print("שגיאות (קרן, סטטוס):", bad[:30], flush=True)
    if empty:
        print("ללא מחירים בטווח:", len(empty), empty[:20], flush=True)


if __name__ == "__main__":
    main()
