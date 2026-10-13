"""
fetch_market.py - נתוני שוק יומיים לאג"ח (מח"מ, תשואה, דירוג) מטבלאות האג"ח של גלובס.

למה דפדפן: המספרים בטבלה נטענים רק אחרי שהדף נפתח, ולכן Power Query לא רואה אותם.
הסקריפט פותח את הדף בכרום מוסתר, ממתין שהמספרים יופיעו וקורא את הטבלה כפי שהיא מוצגת.

רץ בענן (GitHub Actions, ראו README) או במחשב הבית. שומר:
    data/latest.csv                      העותק האחרון — אקסל העל קורא אותו
    data/daily/<YYYY>/<YYYY-MM-DD>.csv   צילום של כל יום מסחר
    data/history/<YYYY>.csv              היסטוריה מצטברת (שורה לכל סדרה בכל יום)

הרצה:
    python fetch_market.py --out data
    python fetch_market.py --out data --skip-if-exists   # לא מוריד שוב אם היום כבר נשמר
    python fetch_market.py --only corp_cpi --show        # טבלה אחת בחלון גלוי (לאבחון)
    python fetch_market.py --out data --xlsx             # גם קובץ Excel יומי

התקנה חד-פעמית במחשב:  pip install playwright pandas openpyxl  ואז  playwright install chromium
"""
import argparse
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
from playwright.sync_api import sync_playwright

TZ = ZoneInfo("Asia/Jerusalem")
BASE = "https://www.globes.co.il/portal/quotes/"
PAGES = {
    "corp_cpi": "cpi-linked-corporate-bonds",
    "corp_unlinked": "unlinked-corporate-bonds",
    "corp_fx": "currency-linked-corporate-bonds",
    "gov_unlinked_fixed": "unlinked-fixed-rate-government-bonds",
    "gov_cpi": "cpi-linked-government-bonds",
}

# שמות עמודות אחידים לפי מילות מפתח בכותרת (אחרי הסרת מרכאות, רווחים וסימנים)
COLUMN_MAP = [
    ("שםאגח", "name"), ("שם", "name"), ("סימול", "security_id"), ("עסקהאחרונה", "last_trade_time"),
    ("שערבסיס", "base_price"), ("שעראחרון", "last_price"), ("שינוייומיבאג", "chg_agorot"),
    ("שינוייומיב", "chg_pct"), ("שינוייומי", "chg_agorot"), ("נפחמסחר", "volume_k_nis"), ("מחמ", "duration"),
    ("תשואהברוטו", "ytm_gross"), ("תשואהנטו", "ytm_net"), ("ערךמתואם", "adjusted_value"),
    ("תאריךפדיון", "maturity"), ("דירוגמעלות", "rating_maalot"), ("דירוגמדרוג", "rating_midroog"),
    ("תשואה", "ytm_gross"),
]
NUMERIC = ["base_price", "last_price", "chg_agorot", "chg_pct", "volume_k_nis", "duration", "ytm_gross", "ytm_net",
           "adjusted_value"]
FIRST = ["date", "snapshot", "source", "security_id", "name", "last_trade_time", "last_price", "chg_pct",
         "volume_k_nis", "duration", "ytm_gross", "ytm_net", "adjusted_value", "maturity", "rating_maalot",
         "rating_midroog"]

EXTRACT_JS = """
() => {
  const out = [];
  for (const t of document.querySelectorAll('table')) {
    const rows = [...t.querySelectorAll('tr')];
    if (rows.length < 3) continue;
    const cells = r => [...r.querySelectorAll('th,td')].map(c => c.innerText.replace(/\\s+/g, ' ').trim());
    const head = rows.find(r => /מח.?מ/.test(r.innerText));
    if (!head) continue;
    const idx = rows.indexOf(head);
    out.push({headers: cells(head), rows: rows.slice(idx + 1).map(cells)});
  }
  return out;
}
"""


def norm(s):
    return re.sub(r"[\s\"'״׳%()\-–.,/]", "", s or "")


def map_col(header, used):
    n = norm(header)
    for key, out in COLUMN_MAP:
        if key in n and out not in used:
            return out
    return header.strip() or "col"


def extract_table(page):
    tables = page.evaluate(EXTRACT_JS)
    if not tables:
        return None
    best = max(tables, key=lambda t: len(t["rows"]))
    used, cols = set(), []
    for h in best["headers"]:
        c = map_col(h, used)
        used.add(c)
        cols.append(c)
    rows = [r for r in best["rows"] if len(r) >= 3 and any(x for x in r)]
    df = pd.DataFrame([r[: len(cols)] + [""] * (len(cols) - len(r)) for r in rows], columns=cols)
    return df.drop(columns=[c for c in ("לתיק", "למעקב") if c in df.columns])


def count_filled(df):
    if df is None or "ytm_gross" not in df.columns:
        return 0
    return int(df["ytm_gross"].astype(str).str.contains(r"\d").sum())


def fetch(page, url, wait_s):
    page.goto(url, wait_until="domcontentloaded", timeout=60000)
    deadline = time.time() + wait_s
    df = None
    while time.time() < deadline:
        time.sleep(3)
        try:
            df = extract_table(page)
        except Exception:
            df = None
        if df is not None and len(df) and count_filled(df) > 0:
            time.sleep(4)
            return extract_table(page)
    return df


def to_num(x):
    s = str(x).replace(",", "").replace("%", "").strip()
    return float(s) if re.fullmatch(r"-?\d+(\.\d+)?", s) else None


def to_date(x):
    s = str(x).strip()
    m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{2,4})", s)
    if not m:
        return s or None
    d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    y = y + 2000 if y < 100 else y
    return f"{y:04d}-{mo:02d}-{d:02d}"


def signature(df):
    keys = []
    for _, r in df.iterrows():
        vals = []
        for c in ("source", "security_id", "last_price", "ytm_gross"):
            v = r.get(c, "")
            try:
                v = f"{float(v):.4f}"
            except (TypeError, ValueError):
                v = "" if pd.isna(v) else str(v)
            vals.append(v)
        keys.append("|".join(vals))
    return "\n".join(sorted(keys))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="data")
    ap.add_argument("--only", choices=list(PAGES))
    ap.add_argument("--show", action="store_true")
    ap.add_argument("--url", help="כתובת מלאה לבדיקה בודדת")
    ap.add_argument("--skip-if-exists", action="store_true")
    ap.add_argument("--xlsx", action="store_true")
    ap.add_argument("--wait", type=int, default=75)
    ap.add_argument("--debug-dir", default="debug")
    args = ap.parse_args()

    now = datetime.now(TZ)
    today = now.strftime("%Y-%m-%d")
    out = Path(args.out)
    daily = out / "daily" / today[:4] / f"{today}.csv"
    if args.skip_if_exists and daily.exists():
        print(f"כבר נשמר היום: {daily}")
        return 0

    targets = {"custom": args.url} if args.url else {
        k: BASE + v + "?showAll=true" for k, v in PAGES.items() if not args.only or k == args.only}
    dbg = Path(args.debug_dir)
    frames, problems = [], []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not args.show)
        ctx = browser.new_context(locale="he-IL", timezone_id="Asia/Jerusalem",
                                  viewport={"width": 1600, "height": 1000})
        page = ctx.new_page()
        for name, url in targets.items():
            df = None
            for attempt in (1, 2):
                print(f"[{name}] ניסיון {attempt}: {url}")
                try:
                    df = fetch(page, url, args.wait)
                except Exception as e:
                    print("   שגיאה:", e)
                    df = None
                if df is not None and count_filled(df) > 0:
                    break
            if df is None or df.empty or count_filled(df) == 0:
                problems.append(f"{name}: לא התקבלה טבלה עם תשואות")
                dbg.mkdir(parents=True, exist_ok=True)
                try:
                    (dbg / f"{name}.html").write_text(page.content(), encoding="utf-8")
                    page.screenshot(path=str(dbg / f"{name}.png"), full_page=False)
                except Exception:
                    pass
                continue
            print(f"   {len(df)} שורות, {count_filled(df)} עם תשואה")
            df.insert(0, "source", name)
            frames.append(df)
        browser.close()

    if not frames:
        print("אף טבלה לא נקלטה:", "; ".join(problems))
        return 1

    allf = pd.concat(frames, ignore_index=True)
    for c in NUMERIC:
        if c in allf.columns:
            allf[c] = allf[c].map(to_num)
    if "maturity" in allf.columns:
        allf["maturity"] = allf["maturity"].map(to_date)
    if "security_id" in allf.columns:
        allf["security_id"] = allf["security_id"].astype(str).str.extract(r"(\d+)", expand=False)
    allf.insert(0, "snapshot", now.strftime("%Y-%m-%d %H:%M"))
    allf.insert(0, "date", today)
    allf = allf[[c for c in FIRST if c in allf.columns] + [c for c in allf.columns if c not in FIRST]]

    latest = out / "latest.csv"
    if latest.exists():
        try:
            old = pd.read_csv(latest, dtype=str, encoding="utf-8-sig")
            if signature(old) == signature(allf):
                print("אין שינוי מול הקובץ האחרון (כנראה יום ללא מסחר) — לא נשמר דבר")
                return 1 if problems else 0
        except Exception as e:
            print("אזהרה: לא נקרא latest.csv:", e)

    daily.parent.mkdir(parents=True, exist_ok=True)
    allf.to_csv(daily, index=False, encoding="utf-8-sig")
    allf.to_csv(latest, index=False, encoding="utf-8-sig")
    hist = out / "history" / f"{today[:4]}.csv"
    hist.parent.mkdir(parents=True, exist_ok=True)
    if hist.exists():
        h = pd.concat([pd.read_csv(hist, dtype=str, encoding="utf-8-sig"), allf.astype(str)], ignore_index=True)
        h = h.drop_duplicates(subset=["date", "source", "security_id"], keep="last")
    else:
        h = allf
    h.to_csv(hist, index=False, encoding="utf-8-sig")
    if args.xlsx:
        allf.to_excel(daily.with_suffix(".xlsx"), index=False)
    print(f"נשמר: {daily} ({len(allf)} שורות)")
    for m in problems:
        print("אזהרה:", m)
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
