# -*- coding: utf-8 -*-
"""maya_funds.py — רשימת כל קרנות הנאמנות ממאיה עם הסיווג הרשמי, פרופיל חשיפה, דמי ניהול, שווי ותשואות מאיה.

    python funds/maya_funds.py data            → data/funds/maya_funds.csv

מקור: POST maya.tase.co.il/api/v1/funds/mutual {pageSize:30, pageNumber:n} (כל הקרנות, ~90 עמודים).
עמודות: מס_קרן, שם, manager_id, מנהל, סיווג_ראשי, סיווג_משני, סיווג_תת, חשיפה_מניות_סמל, חשיפה_מניות, חשיפה_מטח_סמל, חשיפה_מטח,
        דמי_ניהול, דמי_נאמן, שווי_מיליוני_ש״ח, מחיר_פדיון, תאריך_מחיר, תשואה_יום, תשואה_חודש, תשואה_מתחילת_שנה, תשואה_12ח, סטיית_תקן_מאיה,
        מחקה, ממונפת, מועדים_קבועים, כשרה, חוץ, נכס_בסיס, isin
"""
import datetime as dt, json, pathlib, sys, time
import pandas as pd
from playwright.sync_api import sync_playwright

HERE = pathlib.Path(__file__).parent
sys.path.insert(0, str(HERE.parent / "maya"))
from fetch import Maya                      # noqa: E402

JS = """
async ([pages]) => {
  const out = [];
  for (const pg of pages) {
    let res = null;
    for (let t = 0; t < 3; t++) {
      const r = await fetch('/api/v1/funds/mutual', {method: 'POST', headers: {'Content-Type': 'application/json', 'Accept': 'application/json'},
                            body: JSON.stringify({pageSize: 30, pageNumber: pg})});
      if (r.status === 200) { try { res = await r.json(); break; } catch (e) {} }
      await new Promise(s => setTimeout(s, 1500 * (t + 1)));
    }
    out.push({pg, rows: res});
  }
  return out;
}
"""
FEAT = {"מחקה": "קרן מחקה", "ממונפת": "ממונפת", "מועדים_קבועים": "מועדים קבועים", "כשרה": "קרן כשרה", "חוץ": "בלתי מוגבלת בני"}


def row(f):
    cl = f.get("classification") or {}; ep = f.get("exposureProfile") or {}
    feats = {x.get("title", ""): bool(x.get("enabled")) for x in (f.get("features") or [])}
    def feat(sub): return int(any(v for k, v in feats.items() if sub in k))
    return dict(מס_קרן=str(f.get("fundId")).zfill(8), שם=f.get("name"), manager_id=f.get("managerId"), מנהל=f.get("managerName"),
                סיווג_ראשי=cl.get("major"), סיווג_משני=cl.get("main"), סיווג_תת=cl.get("secondary"),
                חשיפה_מניות_סמל=ep.get("shareSymbol"), חשיפה_מניות=ep.get("sharesDesc"), חשיפה_מטח_סמל=ep.get("fxSymbol"), חשיפה_מטח=ep.get("fxDesc"),
                דמי_ניהול=f.get("managementFee"), דמי_נאמן=f.get("trusteeFee"), **{"שווי_מיליוני_ש״ח": f.get("assetValue")}, מחיר_פדיון=f.get("redemptionPrice"),
                תאריך_מחיר=f.get("ratesAsOf"), תשואה_יום=f.get("dayYield"), תשואה_חודש=f.get("monthYield"), תשואה_מתחילת_שנה=f.get("yearYield"),
                תשואה_12ח=f.get("last12MonthYield"), סטיית_תקן_מאיה=f.get("standardDeviation"),
                **{k: feat(v) for k, v in FEAT.items()}, נכס_בסיס=f.get("underlyingAssetsText"), isin=f.get("isin"))


def main():
    data = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else "data"); fdir = data / "funds"; fdir.mkdir(parents=True, exist_ok=True)
    rows, t0 = [], time.time()
    with sync_playwright() as p:
        m = Maya(p)
        m.page.goto("https://maya.tase.co.il/he/funds/mutual-funds", wait_until="networkidle", timeout=90000)
        pg = 1
        while pg < 400:
            res = m.page.evaluate(JS, [list(range(pg, pg + 10))])
            got = 0
            for x in res:
                if x["rows"] is None:
                    print(f"  עמוד {x['pg']}: נכשל", flush=True); continue
                rows.extend(row(f) for f in x["rows"]); got += len(x["rows"])
                if len(x["rows"]) < 30:
                    pg = 10 ** 6; break
            if pg >= 10 ** 6 or got == 0:
                break
            pg += 10
            if (pg // 10) % 3 == 0:
                print(f"  {pg} עמודים, {len(rows)} קרנות, {int(time.time() - t0)} שנ׳", flush=True)
            time.sleep(0.3)
        m.b.close()
    df = pd.DataFrame(rows).drop_duplicates("מס_קרן")
    df.to_csv(fdir / "maya_funds.csv", index=False, encoding="utf-8-sig")
    (fdir / "maya_funds_meta.json").write_text(json.dumps({"updated": dt.date.today().isoformat(), "funds": len(df),
                                                            "price_date": str(df["תאריך_מחיר"].max())}, ensure_ascii=False), encoding="utf-8")
    print(f"{len(df)} קרנות; סיווג ראשי:", df["סיווג_ראשי"].value_counts().head(8).to_dict(), flush=True)


if __name__ == "__main__":
    main()
