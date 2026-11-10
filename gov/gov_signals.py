"""gov_signals.py — אותות לאופטימיזציה של הקובץ הממשלתי (data/gov/signals_latest.csv) ונתוני הרקע שלהם.

קבצים:
  curves_monthly.csv   בנק ישראל (SDMX ZCM): עקומי אפס שקלי/ריאלי חודשיים מ-START — period, tenor, nominal, real, bei
  us_monthly.csv       אוצר ארה"ב: ממוצע חודשי של 5 ו-10 שנים — period, us_5y, us_10y
  economists.csv       gov/economists_manual.csv כפי שהוא (תחזיות בתי ההשקעות: ריבית ואינפלציה בעוד שנה)
  signals_latest.csv   key, value, note, as_of — אחוזוני רמה/תלילות מול היסטוריה של WINDOW_YEARS, פער ישראל–ארה"ב, קונצנזוס ריבית ופיזור
כשל במשיכה → הקובץ הקודם נשאר; אות שאין לו נתונים = ריק (האופטימיזציה מתעלמת ממנו).
"""
from __future__ import annotations
import csv, datetime as dt, io, os, re, sys
from pathlib import Path
import numpy as np, pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gov_data, gov_foreign  # noqa: E402

START = "2008-01"; WINDOW_YEARS = 10; MIN_MONTHS = 12
ZCM_URL = "https://edge.boi.org.il/FusionEdgeServer/sdmx/v2/data/dataflow/BOI.STATISTICS/ZCM/1.0?startperiod={start}&format=csv"
COLS = ["key", "value", "note", "as_of"]


# ---------- פרסרים ----------
def parse_zcm_monthly(text: str) -> pd.DataFrame:
    rows = []
    for r in csv.DictReader(io.StringIO(text)):
        m = re.match(r"ZC_TSB_Z(N|R)D_(\d+)Y_MA", r.get("SERIES_CODE") or "")
        if not m or r.get("OBS_VALUE") in (None, ""): continue
        rows.append({"period": str(r["TIME_PERIOD"])[:7], "tenor": int(m.group(2)), "kind": "nominal" if m.group(1) == "N" else "real", "value": float(r["OBS_VALUE"])})
    if not rows: return pd.DataFrame(columns=["period", "tenor", "nominal", "real", "bei"])
    df = pd.DataFrame(rows).pivot_table(index=["period", "tenor"], columns="kind", values="value").reset_index()
    for c in ("nominal", "real"):
        if c not in df: df[c] = np.nan
    df["bei"] = (((1 + df["nominal"] / 100) / (1 + df["real"] / 100) - 1) * 100).round(4)
    df.columns.name = None
    return df[["period", "tenor", "nominal", "real", "bei"]].sort_values(["period", "tenor"]).reset_index(drop=True)


def parse_treasury_monthly(text: str) -> pd.DataFrame:
    rows = []
    for r in csv.DictReader(io.StringIO(text)):
        try: d = dt.datetime.strptime((r.get("Date") or "").strip(), "%m/%d/%Y")
        except ValueError: continue
        def f(k):
            try: return float(r.get(k))
            except (TypeError, ValueError): return np.nan
        rows.append({"period": d.strftime("%Y-%m"), "us_5y": f("5 Yr"), "us_10y": f("10 Yr")})
    if not rows: return pd.DataFrame(columns=["period", "us_5y", "us_10y"])
    return pd.DataFrame(rows).groupby("period", as_index=False)[["us_5y", "us_10y"]].mean().round(4).sort_values("period").reset_index(drop=True)


def merge_monthly(old: pd.DataFrame | None, new: pd.DataFrame, keys=("period",)) -> pd.DataFrame:
    """החדש גובר על הישן באותם מפתחות; השאר נשמר."""
    if old is None or not len(old): return new.sort_values(list(keys)).reset_index(drop=True)
    if not len(new): return old
    k = list(keys); o = old.merge(new[k], on=k, how="left", indicator=True); o = o[o["_merge"] == "left_only"].drop(columns="_merge")
    return pd.concat([o, new], ignore_index=True).sort_values(k).reset_index(drop=True)


def pct_rank(hist, value) -> float | None:
    """אחוז התצפיות בהיסטוריה שנמוכות מהערך (0–100). אין היסטוריה/ערך → None."""
    h = pd.to_numeric(pd.Series(list(hist), dtype="float64"), errors="coerce").dropna()
    if not len(h) or value is None or (isinstance(value, float) and np.isnan(value)): return None
    return round(float((h < float(value)).mean() * 100), 1)


# ---------- משיכות ----------
def fetch_curves_monthly(start: str = START) -> pd.DataFrame:
    return parse_zcm_monthly(gov_data.get(ZCM_URL.format(start=start), timeout=120, retries=2))


def fetch_us_monthly(prev: pd.DataFrame | None, first_year: int = 2015) -> pd.DataFrame:
    """שנים שחסרות בקובץ הקודם + השנה הנוכחית (תמיד). שנה שנכשלה מדולגת."""
    y_now = dt.date.today().year
    have = set() if prev is None or not len(prev) else {int(p[:4]) for p in prev["period"]}
    out = prev
    for y in [y for y in range(first_year, y_now) if y not in have] + [y_now]:
        try: out = merge_monthly(out, parse_treasury_monthly(gov_data.get(gov_foreign.TREASURY_URL.format(year=y), timeout=60, retries=1)))
        except Exception as e: print(f"אזהרה: אוצר ארה״ב {y} — {str(e)[:60]}")
    return out if out is not None else pd.DataFrame(columns=["period", "us_5y", "us_10y"])


# ---------- אותות ----------
def _num(v):
    try:
        x = float(v); return np.nan if np.isnan(x) else x
    except (TypeError, ValueError): return np.nan


def build_signals(cm: pd.DataFrame, us: pd.DataFrame, econ: pd.DataFrame, policy: pd.DataFrame, curve: pd.DataFrame, us10_now: float | None = None,
                  window_years: int = WINDOW_YEARS) -> pd.DataFrame:
    today = dt.date.today().isoformat(); rows = []
    def add(key, value, note=""): rows.append({"key": key, "value": (np.nan if value is None else value), "note": note, "as_of": today})
    c = curve.set_index("tenor")
    def cv(t, col): return float(c.loc[t, col]) if t in c.index and col in c else np.nan
    real10, nom10, real2 = cv(10, "real"), cv(10, "nominal"), cv(2, "real")
    w = pd.DataFrame()
    if len(cm):
        last = sorted(cm["period"].unique())[-window_years * 12:]; w = cm[cm["period"].isin(last)]
    def hist(t, col): return w.loc[w["tenor"] == t, col].dropna() if len(w) else pd.Series(dtype=float)
    n = int(w["period"].nunique()) if len(w) else 0; ok = n >= MIN_MONTHS
    h_r10, h_n10 = hist(10, "real"), hist(10, "nominal")
    add("hist_months", n, f"חודשים בחלון ההיסטוריה (עד {window_years} שנים, בנק ישראל ZCM)")
    add("hist_from", np.nan if not len(w) else float(str(w["period"].min()).replace("-", ".")), "תחילת החלון (שנה.חודש)")
    add("real10_now", real10, "ריאלי 10 שנים היום (העקום שלנו) %")
    add("real10_pct", pct_rank(h_r10, real10) if ok else None, "אחוזון הריאלי ל-10 שנים מול ההיסטוריה (גבוה = תשואה גבוהה היסטורית)")
    add("real10_med", round(float(h_r10.median()), 3) if len(h_r10) else None, "חציון היסטורי %")
    add("real10_min", round(float(h_r10.min()), 3) if len(h_r10) else None, "מינימום היסטורי %")
    add("real10_max", round(float(h_r10.max()), 3) if len(h_r10) else None, "מקסימום היסטורי %")
    add("nom10_now", nom10, "שקלי 10 שנים היום %")
    add("nom10_pct", pct_rank(h_n10, nom10) if ok else None, "אחוזון השקלי ל-10 שנים מול ההיסטוריה")
    slope_now = (real10 - real2) * 100 if not (np.isnan(real10) or np.isnan(real2)) else np.nan
    hs = pd.Series(dtype=float)
    if len(w):
        p = w.pivot_table(index="period", columns="tenor", values="real")
        if 10 in p and 2 in p: hs = ((p[10] - p[2]) * 100).dropna()
    add("real_slope_now", round(slope_now, 1) if not np.isnan(slope_now) else None, "תלילות ריאלית 2–10 היום (נ״ב)")
    add("real_slope_pct", pct_rank(hs, slope_now) if ok else None, "אחוזון התלילות הריאלית 2–10")
    # ישראל מול ארה"ב, 10 שנים שקלי
    if us10_now is None and len(us): us10_now = float(us.sort_values("period")["us_10y"].dropna().iloc[-1]) if us["us_10y"].notna().any() else None
    spread_now = (nom10 - us10_now) * 100 if us10_now is not None and not np.isnan(nom10) else np.nan
    hsp = pd.Series(dtype=float)
    if len(w) and len(us):
        il = w[w["tenor"] == 10][["period", "nominal"]].merge(us[["period", "us_10y"]], on="period")
        hsp = ((il["nominal"] - il["us_10y"]) * 100).dropna()
    add("us10_now", us10_now, "אג״ח ארה״ב 10 שנים %")
    add("ilus10_now", round(spread_now, 1) if not np.isnan(spread_now) else None, "פער ישראל–ארה״ב 10 שנים (נ״ב)")
    add("ilus10_pct", pct_rank(hsp, spread_now) if len(hsp) >= 12 else None, "אחוזון הפער מול ההיסטוריה (נמוך = ישראל יקרה מול ארה״ב)")
    add("ilus10_med", round(float(hsp.median()), 1) if len(hsp) else None, "חציון היסטורי של הפער (נ״ב)")
    # ריבית: בנק ישראל + כלכלנים
    pol = {} if not len(policy) else dict(zip(policy["item"], policy["value"]))
    boi_fc, boi_infl = _num(pol.get("forecast_end_next_year")), _num(pol.get("forecast_infl_4q"))
    er = pd.to_numeric(econ["rate_12m"], errors="coerce").dropna().tolist() if len(econ) and "rate_12m" in econ else []
    ei = pd.to_numeric(econ["infl_12m"], errors="coerce").dropna().tolist() if len(econ) and "infl_12m" in econ else []
    rates = er + ([boi_fc] if not np.isnan(boi_fc) else []); infl = ei + ([boi_infl] if not np.isnan(boi_infl) else [])
    add("rate_now", _num(pol.get("rate_current")), "ריבית בנק ישראל הנוכחית %")
    add("boi_rate_fc", boi_fc, "תחזית חטיבת המחקר — ריבית בעוד כשנה %")
    add("boi_infl_fc", boi_infl, "תחזית חטיבת המחקר — אינפלציה 4 רבעונים %")
    add("rate_fc_n", len(rates), "מספר התחזיות (חטיבת המחקר + בתי השקעות)")
    add("rate_fc_mean", round(float(np.mean(rates)), 4) if rates else None, "ממוצע תחזיות הריבית בעוד שנה %")
    add("rate_fc_min", min(rates) if rates else None, "התחזית הנמוכה %")
    add("rate_fc_max", max(rates) if rates else None, "התחזית הגבוהה %")
    add("infl_fc_mean", round(float(np.mean(infl)), 4) if infl else None, "ממוצע תחזיות האינפלציה 12 ח׳ %")
    return pd.DataFrame(rows, columns=COLS)


def main(data="data") -> pd.DataFrame:
    g = Path(data) / "gov"; here = Path(__file__).parent
    def rd(p, **k): return pd.read_csv(p, encoding="utf-8-sig", **k) if Path(p).exists() else pd.DataFrame()
    cm = rd(g / "curves_monthly.csv", dtype={"period": str})
    try:
        new = fetch_curves_monthly()
        if len(new): cm = merge_monthly(cm, new, keys=("period", "tenor")); cm.to_csv(g / "curves_monthly.csv", index=False, encoding="utf-8-sig")
    except Exception as e: print("אזהרה: היסטוריה חודשית (בנק ישראל) —", str(e)[:80])
    us = rd(g / "us_monthly.csv", dtype={"period": str})
    try:
        us = fetch_us_monthly(us if len(us) else None)
        if len(us): us.to_csv(g / "us_monthly.csv", index=False, encoding="utf-8-sig")
    except Exception as e: print("אזהרה: ארה״ב חודשי —", str(e)[:80])
    econ = rd(here / "economists_manual.csv", dtype=str)
    if len(econ): econ.to_csv(g / "economists.csv", index=False, encoding="utf-8-sig")
    policy = rd(g / "boi_policy.csv", dtype=str); curve = rd(g / "curves_latest.csv"); fo = rd(g / "foreign_latest.csv")
    us10 = None
    if len(fo) and ((fo["market"] == "US") & (fo["tenor"] == 10) & (fo["stale"] == 0)).any():
        us10 = float(fo[(fo["market"] == "US") & (fo["tenor"] == 10)]["yield"].iloc[0])
    sig = build_signals(cm, us, econ, policy, curve, us10_now=us10); sig.to_csv(g / "signals_latest.csv", index=False, encoding="utf-8-sig")
    return sig


if __name__ == "__main__":
    s = main(sys.argv[1] if len(sys.argv) > 1 else "data")
    print(s[["key", "value"]].to_string(index=False))
