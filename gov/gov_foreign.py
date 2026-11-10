"""gov_foreign.py — עקומי ממשלתי בחו"ל לקובץ הממשלתי (data/gov/foreign_latest.csv).

שווקים: US (FRED: DGS1..DGS30), EUR (ECB Data Portal, עקום AAA של גוש האירו, dataflow YC, יומי), GBP (Bank of England, עקום ספוט נומינלי יומי).
פלט: market, tenor, yield, rate_1y, date, source, as_of, stale — לטווחים TENORS; rate_1y = הטווח 1 של אותו שוק (לחישוב גידור בקירוב CIP).
כשל במשיכה → שורות השוק מהקובץ הקודם עם stale=1; אין קודם → השוק מושמט (אזהרה).
"""
from __future__ import annotations
import csv, datetime as dt, io, os, re, sys, zipfile
from pathlib import Path
import numpy as np, pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gov_data  # noqa: E402

TENORS = [1, 2, 3, 5, 7, 10, 15, 20, 30]
COLS = ["market", "tenor", "yield", "rate_1y", "date", "source", "as_of", "stale"]
US_SERIES = {1: "DGS1", 2: "DGS2", 3: "DGS3", 5: "DGS5", 7: "DGS7", 10: "DGS10", 20: "DGS20", 30: "DGS30"}
ECB_URL = "https://data-api.ecb.europa.eu/service/data/YC/B.U2.EUR.4F.G_N_A.SV_C_YM.{keys}?lastNObservations=3&format=csvdata"   # בקשה אחת לכל הטווחים
TREASURY_URL = ("https://home.treasury.gov/resource-center/data-chart-center/interest-rates/daily-treasury-rates.csv/all/{year}"
                "?type=daily_treasury_yield_curve&field_tdr_date_value={year}&page&_format=csv")                          # אוצר ארה"ב — עקום UST יומי רשמי
BOE_URL = "https://www.bankofengland.co.uk/-/media/boe/files/statistics/yield-curves/latest-yield-curve-data.zip"
SOURCES = {"US": "US Treasury — daily par yield curve (FRED fallback)", "EUR": "ECB Data Portal — AAA euro area spot", "GBP": "Bank of England — nominal spot"}


# ---------- פרסרים (נבדקים בלי רשת) ----------
def parse_ecb_csv(text: str) -> tuple[str, float]:
    rows = [r for r in csv.DictReader(io.StringIO(text)) if r.get("OBS_VALUE") not in (None, "")]
    if not rows: raise ValueError("ECB: אין תצפיות")
    last = max(rows, key=lambda r: r["TIME_PERIOD"])
    return last["TIME_PERIOD"], float(last["OBS_VALUE"])


def parse_ecb_multi_csv(text: str) -> tuple[str, dict[int, float]]:
    """CSV של ECB עם כמה מפתחות (SR_1Y+SR_2Y+...): לכל טווח — התצפית האחרונה. מחזיר (התאריך המאוחר, {טווח: תשואה})."""
    rows = [r for r in csv.DictReader(io.StringIO(text)) if r.get("OBS_VALUE") not in (None, "")]
    if not rows: raise ValueError("ECB: אין תצפיות")
    out: dict[int, tuple[str, float]] = {}
    for r in rows:
        m = re.search(r"SR_(\d+)Y", r.get("KEY", "") + r.get("DATA_TYPE_FM", ""))
        if not m: continue
        t = int(m.group(1))
        if t not in out or r["TIME_PERIOD"] > out[t][0]: out[t] = (r["TIME_PERIOD"], float(r["OBS_VALUE"]))
    return max(v[0] for v in out.values()), {t: v[1] for t, v in sorted(out.items())}


def parse_treasury_csv(text: str) -> tuple[str, dict[int, float]]:
    """CSV של האוצר האמריקאי (daily_treasury_yield_curve): עמודות '1 Yr'...'30 Yr', תאריך mm/dd/yyyy. מחזיר (התאריך האחרון, {שנים: תשואה})."""
    rows = list(csv.DictReader(io.StringIO(text)))
    if not rows: raise ValueError("Treasury: אין שורות")
    def d(r): return dt.datetime.strptime(r["Date"].strip(), "%m/%d/%Y").date()
    last = max(rows, key=d); out = {}
    for k, v in last.items():
        m = re.match(r"\s*(\d+)\s*Yr", k or "")
        if m and v not in (None, "", "N/A"):
            try: out[int(m.group(1))] = float(v)
            except ValueError: pass
    if not out: raise ValueError("Treasury: אין תשואות")
    return d(last).isoformat(), out


def parse_boe_spot(zip_bytes: bytes) -> tuple[str, dict[float, float]]:
    import openpyxl
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
        name = next(n for n in z.namelist() if "Nominal" in n and n.lower().endswith(".xlsx"))
        wb = openpyxl.load_workbook(io.BytesIO(z.read(name)), data_only=True, read_only=True)
    ws = next(ws for ws in wb.worksheets if "spot curve" in ws.title.lower())
    rows = list(ws.iter_rows(values_only=True))
    hdr_i = next(i for i, r in enumerate(rows) if sum(isinstance(v, (int, float)) for v in r[1:]) >= 5)
    tenors = {j: float(v) for j, v in enumerate(rows[hdr_i]) if j >= 1 and isinstance(v, (int, float))}
    last = None
    for r in rows[hdr_i + 1:]:
        if r and isinstance(r[0], (dt.datetime, dt.date)) and any(isinstance(v, (int, float)) for v in r[1:]):
            last = r
    if last is None: raise ValueError("BoE: אין שורת נתונים")
    d = last[0].date() if isinstance(last[0], dt.datetime) else last[0]
    return d.isoformat(), {t: float(last[j]) for j, t in tenors.items() if isinstance(last[j], (int, float))}


# ---------- משיכות ----------
def fetch_us() -> tuple[str, dict[float, float]]:
    """UST: קודם האוצר האמריקאי (עקום פארי יומי, כולל 3/7/20 שנים); אם נכשל — FRED."""
    try:
        y = dt.date.today().year
        try: return parse_treasury_csv(gov_data.get(TREASURY_URL.format(year=y), timeout=60, retries=2))
        except Exception:
            return parse_treasury_csv(gov_data.get(TREASURY_URL.format(year=y - 1), timeout=60, retries=1))   # תחילת שנה — הקובץ של השנה עדיין ריק
    except Exception as e_tr:
        fred = gov_data.fetch_fred()
        ok = fred[fred["value"].notna()]
        vals = {t: float(ok.loc[ok["series"] == s, "value"].iloc[0]) for t, s in US_SERIES.items() if (ok["series"] == s).any()}
        if not vals: raise RuntimeError(f"Treasury: {str(e_tr)[:50]}; FRED: אין נתונים")
        return str(ok["date"].max()), vals


def fetch_ecb(tenors=TENORS) -> tuple[str, dict[float, float]]:
    keys = "+".join(f"SR_{t}Y" for t in tenors)
    d, vals = parse_ecb_multi_csv(gov_data.get(ECB_URL.format(keys=keys), timeout=90, retries=3))
    missing = [t for t in tenors if t not in vals]
    if missing: raise RuntimeError(f"ECB: חסרים טווחים {missing}")
    return d, {float(t): v for t, v in vals.items()}


def fetch_boe() -> tuple[str, dict[float, float]]:
    import urllib.request
    req = urllib.request.Request(BOE_URL, headers=gov_data.UA)
    with urllib.request.urlopen(req, timeout=90) as r: return parse_boe_spot(r.read())


def _on_grid(vals: dict[float, float]) -> dict[int, float]:
    xs = sorted(vals); ys = [vals[x] for x in xs]
    return {t: round(float(np.interp(t, xs, ys)), 4) for t in TENORS}


def build_foreign(data: Path | str, prev: pd.DataFrame | None) -> pd.DataFrame:
    today = dt.date.today().isoformat(); out = []
    for market, fetch in (("US", fetch_us), ("EUR", lambda: fetch_ecb(TENORS)), ("GBP", fetch_boe)):
        try:
            d, vals = fetch(); grid = _on_grid(vals)
            r1 = grid[1]
            out += [{"market": market, "tenor": t, "yield": grid[t], "rate_1y": r1, "date": d, "source": SOURCES[market], "as_of": today, "stale": 0} for t in TENORS]
        except Exception as e:
            print(f"אזהרה: {market} — {str(e)[:80]}")
            if prev is not None and (prev["market"] == market).any():
                p = prev[prev["market"] == market].copy(); p["stale"] = 1; p["as_of"] = today
                out += p[COLS].to_dict("records")
            else:
                print(f"אזהרה: {market} הושמט — אין נתון קודם")
    return pd.DataFrame(out, columns=COLS)


if __name__ == "__main__":
    data = Path(sys.argv[1] if len(sys.argv) > 1 else "data"); f = data / "gov" / "foreign_latest.csv"
    prev = pd.read_csv(f, encoding="utf-8-sig") if f.exists() else None
    df = build_foreign(data, prev); df.to_csv(f, index=False, encoding="utf-8-sig")
    print(df.groupby("market")[["date", "stale"]].first())
