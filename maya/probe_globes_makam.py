"""probe_globes_makam.py <out_dir> — בדיקה חד-פעמית: מה גלובס מחזיר בטבלת המק"מ (למה fetch_market לא קולט תשואות).
שומר לכל וריאנט של הכתובת: כותרות וטבלאות שנמצאו, קטע מהטקסט, HTML וצילום מסך."""
import json, sys, time
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from fetch_market import EXTRACT_JS, extract_table  # noqa: E402

out = Path(sys.argv[1]); out.mkdir(parents=True, exist_ok=True)
URLS = {"showall": "https://www.globes.co.il/portal/quotes/unlinked-short-government-bonds?showAll=true",
        "plain": "https://www.globes.co.il/portal/quotes/unlinked-short-government-bonds",
        "plain_slash": "https://www.globes.co.il/portal/quotes/unlinked-short-government-bonds/"}
from playwright.sync_api import sync_playwright
res = {}
with sync_playwright() as p:
    b = p.chromium.launch(headless=True)
    ctx = b.new_context(locale="he-IL", timezone_id="Asia/Jerusalem", viewport={"width": 1600, "height": 1000})
    page = ctx.new_page()
    for key, url in URLS.items():
        r = {"url": url}
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=60000)
            snaps = []
            for i in range(8):
                time.sleep(8)
                tables = page.evaluate(EXTRACT_JS)
                snaps.append([{"headers": t["headers"], "n": len(t["rows"]), "rows": t["rows"][:3]} for t in tables])
            r["tables_over_time"] = snaps
            try:
                df = extract_table(page)
                r["extract_table"] = None if df is None else {"cols": list(df.columns), "n": len(df), "head": df.head(5).to_dict("records")}
            except Exception as e:
                r["extract_table"] = f"err {e}"
            all_tables = page.evaluate("() => [...document.querySelectorAll('table')].map(t => ({rows: t.querySelectorAll('tr').length, text: t.innerText.slice(0,400)}))")
            r["all_tables"] = all_tables
            r["text"] = page.inner_text("body")[:3000]
            (out / f"{key}.html").write_text(page.content(), encoding="utf-8")
            page.screenshot(path=str(out / f"{key}.png"), full_page=True)
        except Exception as e:
            r["error"] = str(e)
        res[key] = r
    b.close()
json.dump(res, open(out / "makam_probe.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps({k: {"tables": [[t["headers"][:4], t["n"]] for t in v.get("tables_over_time", [[]])[-1]], "extract": (v.get("extract_table") or {}).get("n") if isinstance(v.get("extract_table"), dict) else v.get("extract_table")} for k, v in res.items()}, ensure_ascii=False))
