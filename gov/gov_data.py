"""
gov_data.py — נתוני אג"ח ממשלת ישראל מתוך מאגר המידע שלנו (גלובס יומי + היסטוריה רבעונית), בנק ישראל ו-FRED.
זהו מקור הנתונים של קובץ הממשלתי (gov/build_gov.py) ושל כל מי שצריך נתוני ממשלתי: עקומים, BEI, ציפיות.

פלט (data/gov/):
  bonds_latest.csv    שורה לכל אג"ח ממשלתית / מק"מ ליום המסחר האחרון: מחיר, תשואה, מח"מ (גלובס), ועוד חישובים שלנו —
                      שנים לפדיון, ריבית צבורה, מחיר נקי, מח"מ מתואם, קמירות (אג"ח בולט, קופון שנתי, ריבית שנתית אפקטיבית, ACT/365)
  curves_latest.csv   עקום שקלי (מק"מ + שקלית קבועה), עקום צמוד, BEI, פורוורדים ו-roll-down לפי טווחים קבועים (אינטרפולציה ליניארית על שנים לפדיון)
  curves_history.csv  אותם עקומים ב-1/2/5/10 שנים לכל יום מסחר שיש לנו (גלובס מ-13.10.2026) ולכל תאריך בקובץ ההיסטוריה הרבעוני (מגמה, 1/2023–10/2026)
  us_latest.csv       ארה"ב (FRED): ריבית פד, UST 3M/1/2/5/10/30, TIPS 5/10, Breakeven 5/10 — ערך אחרון + שינוי 1M/3M
  boi_curves.csv      בנק ישראל: עקומי אפס שקלי/ריאלי 1–20 שנה (חודשי, SDMX ZCM) להשוואה מול העקום שלנו
  boi_expectations.csv ציפיות לאינפלציה לשנה לפי מקור (חזאים, שוק ההון, בנקים, חוזים) + פורוורדים (SDMX ZCM)
הרצה: python gov/gov_data.py data
"""
from __future__ import annotations
import csv, datetime as dt, io, json, os, re, sys, urllib.request
from pathlib import Path
import numpy as np, pandas as pd

UA = {"User-Agent": "Mozilla/5.0 corp-market-data gov"}
TENORS = [0.25, 0.5, 1, 2, 3, 4, 5, 7, 10, 15, 20, 30]
HORIZON = 1.0      # אופק roll-down בשנים
EX_DAYS = 7        # חלון אקס-קופון (ימים לפני התשלום)
NOMINAL_SRC = ("gov_makam", "gov_unlinked_fixed")
REAL_SRC = ("gov_cpi",)
FRED = [("DFF", "ריבית הפד (אפקטיבית)"), ("DGS3MO", "UST 3 חודשים"), ("DGS1", "UST שנה"), ("DGS2", "UST 2 שנים"), ("DGS5", "UST 5 שנים"),
        ("DGS10", "UST 10 שנים"), ("DGS30", "UST 30 שנים"), ("DFII5", "TIPS 5 שנים (ריאלי)"), ("DFII10", "TIPS 10 שנים (ריאלי)"),
        ("T5YIE", "Breakeven 5 שנים"), ("T10YIE", "Breakeven 10 שנים")]


def get(url, timeout=60):
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read().decode("utf-8-sig", errors="replace")


# ---------- זיהוי מועד פדיון משם הנייר (לממשלתיות אין מס' ני"ע בגלובס; ההיסטוריה הרבעונית בלי פדיון) ----------
def _eom(y, m):
    return (dt.date(y + m // 12, m % 12 + 1, 1) - dt.timedelta(days=1))


def maturity_from_name(name: str, ref: dt.date) -> dt.date | None:
    s = str(name)
    m = re.search(r"גלובל\s*(\d{2})/(\d{2})", s)
    if m: return _eom(2000 + int(m.group(2)), int(m.group(1)))
    if "מ.ק.מ" in s or "מקמ" in s or "מק\"מ" in s:
        d = re.sub(r"\D", "", s)
        if len(d) == 3: mo, yd = int(d[0]), int(d[2])
        elif len(d) == 4: mo, yd = int(d[:2]), int(d[3])
        else: return None
        base = ref.year - ref.year % 10
        cands = [base - 10 + yd, base + yd, base + 10 + yd]
        y = min(cands, key=lambda yy: abs((dt.date(yy, mo, 1) - ref).days) if 1 <= mo <= 12 else 9e9)
        return dt.date(y, mo, 1) if 1 <= mo <= 12 else None
    m = re.search(r"(\d{2})(\d{2})\s*$", s)
    if m:
        mo, yy = int(m.group(1)), int(m.group(2))
        if 1 <= mo <= 12: return _eom(2000 + yy, mo)
    return None


def gov_type(name: str, source: str | None = None) -> str:
    s = str(name)
    if "מ.ק.מ" in s or "מקמ" in s or "מק\"מ" in s or source == "gov_makam": return "מק\"מ"
    if "צמוד" in s or source == "gov_cpi": return "מדד"
    if "משתנה" in s or source == "gov_float": return "משתנה"
    if "גלובל" in s: return "גלובל"
    return "שקלי"


# ---------- אנליטיקה לאג"ח בולט עם קופון שנתי ----------
def bond_analytics(val: dt.date, maturity: dt.date, coupon: float, ytm: float, kind: str):
    """מחזיר dict: ttm, accrued (ל-100 ע.נ.), pv (מחיר מלוכלך תיאורטי ב-100 ע.נ. ריאלי), macaulay, modified, convexity.
    ריבית שנתית אפקטיבית, ACT/365. מק"מ = אפס קופון."""
    y = ytm / 100.0
    ttm = (maturity - val).days / 365.0
    if ttm <= 0: return None
    if kind == "מק\"מ" or not coupon:
        flows = [(maturity, 100.0)]
        prev = None
    else:
        flows, d = [], maturity
        while d > val:
            flows.append((d, coupon + (100.0 if d == maturity else 0.0)))
            try: d = d.replace(year=d.year - 1)
            except ValueError: d = d.replace(year=d.year - 1, day=28)
        prev = d
        flows.reverse()
        # אקס-קופון: תשלום בתוך EX_DAYS ימים כבר לא שייך לקונה (גלובס מציג ערך מתואם 100) — מושמט, והריבית הצבורה מתאפסת
        if flows and (flows[0][0] - val).days <= EX_DAYS and flows[0][0] != maturity:
            prev = flows[0][0]; flows = flows[1:]
    ts = [((d - val).days / 365.0, cf) for d, cf in flows]
    pv = sum(cf / (1 + y) ** t for t, cf in ts)
    mac = sum(t * cf / (1 + y) ** t for t, cf in ts) / pv
    conv = sum(t * (t + 1) * cf / (1 + y) ** t for t, cf in ts) / (pv * (1 + y) ** 2)
    acc = 0.0
    if prev is not None and coupon:
        nxt = flows[0][0]
        acc = max(0.0, coupon * (val - prev).days / max(1, (nxt - prev).days))
    return {"ttm": ttm, "accrued": acc, "pv": pv, "macaulay": mac, "modified": mac / (1 + y), "convexity": conv}


# ---------- עקומים ----------
def interp(x, xs, ys):
    xs = np.asarray(xs, float); ys = np.asarray(ys, float)
    if len(xs) == 0: return np.nan
    o = np.argsort(xs); xs, ys = xs[o], ys[o]
    return float(np.interp(x, xs, ys))     # מחוץ לטווח — ערך הקצה (כמו במודל של יובב)


def build_curve(nom: pd.DataFrame, real: pd.DataFrame, tenors=TENORS):
    rows = []; prev = None
    for t in tenors:
        n = interp(t, nom["ttm"], nom["ytm"]) if len(nom) else np.nan
        r = interp(t, real["ttm"], real["ytm"]) if len(real) else np.nan
        bei = (1 + n / 100) / (1 + r / 100) - 1 if not (np.isnan(n) or np.isnan(r)) else np.nan
        row = {"tenor": t, "nominal": round(n, 4), "real": round(r, 4), "bei": round(bei * 100, 4),
               "bei_cum": round(((1 + bei) ** t - 1) * 100, 4) if not np.isnan(bei) else np.nan}
        if prev is None:
            row.update(fwd_nominal=row["nominal"], fwd_real=row["real"])
        else:
            t0 = prev["tenor"]
            fn = ((1 + n / 100) ** t / (1 + prev["nominal"] / 100) ** t0) ** (1 / (t - t0)) - 1
            fr = ((1 + r / 100) ** t / (1 + prev["real"] / 100) ** t0) ** (1 / (t - t0)) - 1
            row.update(fwd_nominal=round(fn * 100, 4), fwd_real=round(fr * 100, 4))
        fb = (1 + row["fwd_nominal"] / 100) / (1 + row["fwd_real"] / 100) - 1
        row["fwd_bei"] = round(fb * 100, 4)
        rows.append(row); prev = row
    # roll-down: תשואה בטווח פחות תשואת העקום (על טבלת הטווחים, כמו במודל) בטווח הקצר יותר באופק
    ts = [r["tenor"] for r in rows]
    for row in rows:
        tr = max(row["tenor"] - HORIZON, tenors[0])
        row["roll_nominal_bp"] = round((row["nominal"] - interp(tr, ts, [r["nominal"] for r in rows])) * 100, 1)
        row["roll_real_bp"] = round((row["real"] - interp(tr, ts, [r["real"] for r in rows])) * 100, 1)
    df = pd.DataFrame(rows)
    n2 = df.loc[df.tenor == 2, "nominal"].iloc[0]; r2 = df.loc[df.tenor == 2, "real"].iloc[0]
    df["slope_nominal_vs2y_bp"] = ((df["nominal"] - n2) * 100).round(1); df["slope_real_vs2y_bp"] = ((df["real"] - r2) * 100).round(1)
    return df


def curve_inputs(df: pd.DataFrame, val: dt.date):
    """df: שורות גלובס/היסטוריה עם name, source, ytm_gross, maturity (או בלי — נגזר מהשם). מחזיר (nom, real) עם ttm, ytm."""
    out = []
    for _, r in df.iterrows():
        y = pd.to_numeric(r.get("ytm_gross"), errors="coerce")
        if pd.isna(y) or y > 50 or y < -5: continue
        mat = None
        if isinstance(r.get("maturity"), str) and r["maturity"]:
            try: mat = pd.to_datetime(r["maturity"]).date()
            except Exception: mat = None
        if mat is None: mat = maturity_from_name(r["name"], val)
        if mat is None:
            d = pd.to_numeric(r.get("duration"), errors="coerce")
            if pd.isna(d) or d <= 0: continue
            ttm = float(d)
        else:
            ttm = (mat - val).days / 365.0
        if ttm <= 0.02: continue
        out.append({"name": r["name"], "kind": gov_type(r["name"], r.get("source")), "ttm": ttm, "ytm": float(y)})
    o = pd.DataFrame(out, columns=["name", "kind", "ttm", "ytm"])
    return o[o.kind.isin(["מק\"מ", "שקלי"])], o[o.kind == "מדד"]


# ---------- מקורות חיצוניים ----------
def fetch_fred():
    rows = []
    for sid, name in FRED:
        try:
            txt = get(f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={sid}&cosd={(dt.date.today() - dt.timedelta(days=130)).isoformat()}")
            s = pd.read_csv(io.StringIO(txt)); s.columns = ["date", "v"]; s["v"] = pd.to_numeric(s["v"], errors="coerce"); s = s.dropna()
            if s.empty: continue
            last = s.iloc[-1]; d0 = pd.to_datetime(last["date"])
            def back(days):
                b = s[pd.to_datetime(s["date"]) <= d0 - pd.Timedelta(days=days)]
                return float(b.iloc[-1]["v"]) if len(b) else np.nan
            rows.append({"series": sid, "name": name, "date": last["date"], "value": float(last["v"]),
                         "chg_1m_bp": round((float(last["v"]) - back(30)) * 100, 0), "chg_3m_bp": round((float(last["v"]) - back(91)) * 100, 0),
                         "source": "FRED (St. Louis Fed)"})
        except Exception as e:
            rows.append({"series": sid, "name": name, "date": None, "value": None, "chg_1m_bp": None, "chg_3m_bp": None, "source": f"שגיאה: {str(e)[:60]}"})
    return pd.DataFrame(rows)


def fetch_boi():
    start = (dt.date.today() - dt.timedelta(days=120)).isoformat()
    txt = get(f"https://edge.boi.org.il/FusionEdgeServer/sdmx/v2/data/dataflow/BOI.STATISTICS/ZCM/1.0?startperiod={start}&format=csv")
    rows = list(csv.DictReader(io.StringIO(txt)))
    last = {}
    for r in rows:
        c = r.get("SERIES_CODE"); v = r.get("OBS_VALUE")
        if not c or v in (None, ""): continue
        if c not in last or r["TIME_PERIOD"] > last[c]["TIME_PERIOD"]: last[c] = r
    curves = []
    for c, r in last.items():
        m = re.match(r"ZC_TSB_Z(N|R)D_(\d+)Y_MA", c)
        if m: curves.append({"tenor": int(m.group(2)), "kind": "nominal" if m.group(1) == "N" else "real", "value": float(r["OBS_VALUE"]), "period": r["TIME_PERIOD"]})
    cv = pd.DataFrame(curves)
    if len(cv):
        cv = cv.pivot_table(index=["tenor", "period"], columns="kind", values="value").reset_index().sort_values("tenor")
        cv["bei"] = (((1 + cv["nominal"] / 100) / (1 + cv["real"] / 100) - 1) * 100).round(3)
    names = {"EXP_HAZ_MAD_AVG_01Y_MA": "ממוצע החזאים — 12 חודשים", "EXP_EXP_01Y_BI_MA": "שוק ההון — שנה", "EXP_EXP_IRB_1Y_MA": "ריביות הבנקים — שנה",
             "EXP_EXP_ISSWIT1_01Y_MA": "חוזי אינפלציה — שנה", "EXP_EXB_FRW_ZD_01Y_02Y_MA": "שוק ההון — פורוורד שנה 2", "EXP_EXB_FRW_ZD_02Y_03Y_MA": "שוק ההון — פורוורד שנה 3",
             "EXP_EXB_FRW_ZD_03Y_05Y_MA": "שוק ההון — פורוורד 3–5", "EXP_EXB_ZD_5Y_MA": "שוק ההון — 5 שנים", "EXP_EXB_FRW_ZD_05Y_10Y_MA": "שוק ההון — פורוורד 5–10",
             "EXP_EXB_ZD_10Y_MA": "שוק ההון — 10 שנים", "INF_MIN_TRGT_D": "יעד אינפלציה — תחתון", "INF_MAX_TRGT_D": "יעד אינפלציה — עליון"}
    ex = pd.DataFrame([{"series": c, "name": names[c], "period": r["TIME_PERIOD"], "value": float(r["OBS_VALUE"])} for c, r in last.items() if c in names])
    return cv, ex


# ---------- ראשי ----------
def main(data="data"):
    d = Path(data); out = d / "gov"; out.mkdir(parents=True, exist_ok=True)
    mk = pd.read_csv(d / "latest.csv", encoding="utf-8-sig")
    g = mk[mk["source"].astype(str).str.startswith("gov")].copy()
    date = pd.to_datetime(mk["date"].max()).date(); val = date + dt.timedelta(days=1)   # סליקה T+1
    static = pd.read_csv(d / "hist" / "bond_static.csv", encoding="utf-8-sig")
    sid_by_name = {str(n).strip(): int(s) for n, s in zip(static["name"], static["security_id"]) if pd.notna(s)}
    rows = []
    for _, r in g.iterrows():
        name = str(r["name"]).strip(); kind = gov_type(name, r["source"])
        y = pd.to_numeric(r.get("ytm_gross"), errors="coerce"); c = pd.to_numeric(r.get("coupon"), errors="coerce")
        mat = None
        try: mat = pd.to_datetime(r["maturity"]).date() if isinstance(r.get("maturity"), str) and r["maturity"] else None
        except Exception: pass
        if mat is None: mat = maturity_from_name(name, val)
        a = bond_analytics(val, mat, 0.0 if pd.isna(c) else float(c), float(y), kind) if (mat and not pd.isna(y)) else None
        price = pd.to_numeric(r.get("last_price"), errors="coerce"); adj = pd.to_numeric(r.get("adjusted_value"), errors="coerce")
        L = None
        if a and kind == "מדד" and not pd.isna(adj): L = adj / (100 * (1 + a["accrued"] / 100))
        rows.append({"date": date.isoformat(), "security_id": sid_by_name.get(name), "name": name, "kind": kind, "source": r["source"],
                     "maturity": mat.isoformat() if mat else None, "coupon": None if pd.isna(c) else float(c),
                     "last_price": None if pd.isna(price) else float(price), "adjusted_value": None if pd.isna(adj) else float(adj),
                     "ytm_gross": None if pd.isna(y) else float(y), "ytm_net": pd.to_numeric(r.get("ytm_net"), errors="coerce"),
                     "duration_globes": pd.to_numeric(r.get("duration"), errors="coerce"), "volume_k_nis": pd.to_numeric(r.get("volume_k_nis"), errors="coerce"),
                     "ttm": round(a["ttm"], 4) if a else None, "accrued": round(a["accrued"] * (L or 1), 4) if a else None,
                     "clean_price": round(float(price) - a["accrued"] * (L or 1), 4) if (a and not pd.isna(price)) else None,
                     "macaulay_calc": round(a["macaulay"], 4) if a else None, "modified": round(a["modified"], 4) if a else None,
                     "convexity": round(a["convexity"], 4) if a else None, "L": round(L, 5) if L else None,
                     "pv_check": round(a["pv"] * (L or 1), 2) if a else None,
                     "in_curve": "" if kind in ("מק\"מ", "שקלי", "מדד") and a else "לא"})
    bonds = pd.DataFrame(rows).sort_values(["kind", "ttm"])
    bonds.to_csv(out / "bonds_latest.csv", index=False, encoding="utf-8-sig")
    # עקום נוכחי
    nom, real = curve_inputs(g, val)
    cur = build_curve(nom, real); cur.insert(0, "date", date.isoformat())
    cur.to_csv(out / "curves_latest.csv", index=False, encoding="utf-8-sig")
    # היסטוריה: גלובס יומי + רבעוני (מגמה)
    hist = []
    for f in sorted((d / "history").glob("*.csv")):
        h = pd.read_csv(f, encoding="utf-8-sig"); h = h[h["source"].astype(str).str.startswith("gov")]
        for dte, grp in h.groupby("date"):
            v = pd.to_datetime(dte).date() + dt.timedelta(days=1)
            n, rr = curve_inputs(grp, v)
            if len(n) < 3 or len(rr) < 3: continue
            c = build_curve(n, rr, [1, 2, 5, 10])
            hist.append({"date": dte, "src": "גלובס", **{f"nom_{t}": c.loc[c.tenor == t, "nominal"].iloc[0] for t in (1, 2, 5, 10)},
                         **{f"real_{t}": c.loc[c.tenor == t, "real"].iloc[0] for t in (1, 2, 5, 10)}, **{f"bei_{t}": c.loc[c.tenor == t, "bei"].iloc[0] for t in (1, 2, 5, 10)}})
    try:
        q = pd.read_csv(d / "hist" / "quarterly_spreads.csv", encoding="utf-8-sig")
        gov_ids = set(static.loc[static["is_gov"] == 1, "security_id"].astype(int))
        q = q[q["security_id"].isin(gov_ids)].copy(); q["source"] = q["name"].map(lambda n: {"מק\"מ": "gov_makam", "מדד": "gov_cpi", "שקלי": "gov_unlinked_fixed"}.get(gov_type(n), "gov_other"))
        daily_dates = {h["date"] for h in hist}
        for dte, grp in q.groupby("date"):
            if dte in daily_dates: continue
            v = pd.to_datetime(dte).date() + dt.timedelta(days=1)
            n, rr = curve_inputs(grp.assign(maturity=None), v)
            if len(n) < 3 or len(rr) < 3: continue
            c = build_curve(n, rr, [1, 2, 5, 10])
            hist.append({"date": dte, "src": "מגמה (רבעוני)", **{f"nom_{t}": c.loc[c.tenor == t, "nominal"].iloc[0] for t in (1, 2, 5, 10)},
                         **{f"real_{t}": c.loc[c.tenor == t, "real"].iloc[0] for t in (1, 2, 5, 10)}, **{f"bei_{t}": c.loc[c.tenor == t, "bei"].iloc[0] for t in (1, 2, 5, 10)}})
    except Exception as e:
        print("אזהרה: היסטוריה רבעונית —", e)
    pd.DataFrame(hist).sort_values("date").to_csv(out / "curves_history.csv", index=False, encoding="utf-8-sig")
    # ארה"ב ובנק ישראל
    try: fetch_fred().to_csv(out / "us_latest.csv", index=False, encoding="utf-8-sig")
    except Exception as e: print("אזהרה: FRED —", e)
    try:
        cv, ex = fetch_boi(); cv.to_csv(out / "boi_curves.csv", index=False, encoding="utf-8-sig"); ex.to_csv(out / "boi_expectations.csv", index=False, encoding="utf-8-sig")
    except Exception as e: print("אזהרה: בנק ישראל —", e)
    b1 = cur.loc[cur.tenor == 1].iloc[0]
    print(f"{date}: {len(bonds)} ממשלתיות ({bonds.kind.value_counts().to_dict()}); עקום 1y: שקלי {b1.nominal:.2f} ריאלי {b1.real:.2f} BEI {b1.bei:.2f}; "
          f"היסטוריה {len(hist)} תאריכים")
    chk = bonds.dropna(subset=["macaulay_calc", "duration_globes"])
    print("מח״מ מחושב מול גלובס — הפרש חציוני:", round((chk["macaulay_calc"] - chk["duration_globes"]).abs().median(), 3),
          "מקס':", round((chk["macaulay_calc"] - chk["duration_globes"]).abs().max(), 3))


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "data")
