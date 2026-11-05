"""gov_history.py — היסטוריה לכל סדרה ממשלתית ותנודתיות העקום (מקור הממשלתי, data/gov).

build_bond_history(data) -> DataFrame: שורה לכל סדרה ויום מסחר ב-data/history/*.csv (source gov_*):
    date, security_id, key, name, kind, ttm, ytm, duration, volume_k_nis, curve_ytm, spread_bp
    curve_ytm = אינטרפולציה של עקום אותו היום (על רשת הטווחים TENORS, כמו בגיליון 03) לפי ttm; מק"מ ושקלי מול העקום השקלי, צמוד מול הריאלי.
    key = מס' ני"ע כטקסט, או השם כשאין מס' ני"ע (שתי צמודות חדשות חסרות במאגר הסטטי).
build_vol(curves_history, floor_bp) -> DataFrame: tenor, nominal_vol_bp, real_vol_bp, bei_vol_bp, n_daily, method
    סטיית תקן שנתית (נ"ב) של שינויי התשואה היומיים × sqrt(250); n_daily < 60 → max(יומי, רבעוני×sqrt(4), רצפה) ו-method="quarterly+floor".
"""
from __future__ import annotations
import datetime as dt, os, sys
from pathlib import Path
import numpy as np, pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gov_data  # noqa: E402

TENORS = gov_data.TENORS
HIST_TENORS = (1, 2, 5, 10)
MIN_DAILY = 60
COLS = ["date", "security_id", "key", "name", "kind", "ttm", "ytm", "duration", "volume_k_nis", "curve_ytm", "spread_bp"]


def bond_key(sid, name: str) -> str:
    s = pd.to_numeric(sid, errors="coerce")
    return str(int(s)) if pd.notna(s) else str(name).strip()


def build_bond_history(data: Path | str) -> pd.DataFrame:
    d = Path(data); rows = []
    for f in sorted((d / "history").glob("*.csv")):
        h = pd.read_csv(f, encoding="utf-8-sig")
        h = h[h["source"].astype(str).str.startswith("gov")]
        for dte, grp in h.groupby("date"):
            val = pd.to_datetime(dte).date() + dt.timedelta(days=1)
            nom, real = gov_data.curve_inputs(grp, val)
            if len(nom) < 3 or len(real) < 3: continue
            cur = gov_data.build_curve(nom, real)
            for _, r in grp.iterrows():
                name = str(r["name"]).strip(); kind = gov_data.gov_type(name, r.get("source"))
                y = pd.to_numeric(r.get("ytm_gross"), errors="coerce")
                mat = None
                if isinstance(r.get("maturity"), str) and r["maturity"]:
                    try: mat = pd.to_datetime(r["maturity"]).date()
                    except Exception: mat = None
                if mat is None: mat = gov_data.maturity_from_name(name, val)
                ttm = (mat - val).days / 365.0 if mat else np.nan
                col = "real" if kind == "מדד" else "nominal"
                cy = gov_data.interp(ttm, cur["tenor"], cur[col]) if (kind in ("מק\"מ", "שקלי", "מדד") and not np.isnan(ttm)) else np.nan
                rows.append({"date": dte, "security_id": pd.to_numeric(r.get("security_id"), errors="coerce"), "key": bond_key(r.get("security_id"), name),
                             "name": name, "kind": kind, "ttm": round(ttm, 4) if not np.isnan(ttm) else np.nan,
                             "ytm": float(y) if pd.notna(y) else np.nan, "duration": pd.to_numeric(r.get("duration"), errors="coerce"),
                             "volume_k_nis": pd.to_numeric(r.get("volume_k_nis"), errors="coerce"),
                             "curve_ytm": round(cy, 4) if not np.isnan(cy) else np.nan,
                             "spread_bp": round((float(y) - cy) * 100, 1) if (pd.notna(y) and not np.isnan(cy)) else np.nan})
    df = pd.DataFrame(rows, columns=COLS)
    if len(df): df["security_id"] = df["security_id"].astype("Int64")
    return df.sort_values(["date", "kind", "ttm"]).reset_index(drop=True)


def _ann_std(series: pd.Series, periods_per_year: float) -> float:
    s = pd.to_numeric(series, errors="coerce").dropna()
    if len(s) < 3: return np.nan
    return float(s.diff().dropna().std(ddof=1) * np.sqrt(periods_per_year) * 100)


def build_vol(curves_history: pd.DataFrame, floor_bp: dict[str, float]) -> pd.DataFrame:
    """תנודתיות שנתית (נ"ב) לכל טווח. n_daily ≥ 60 → סטיית התקן היומית × sqrt(250) ("daily");
    אחרת max(אומדן יומי, רצפה) ("floor") — האומדן הרבעוני (16 נקודות, 2023–2026) רועש מדי לפי טווח ולכן מדווח רק לעיון (quarterly_vol_bp)."""
    h = curves_history.copy(); h["date"] = pd.to_datetime(h["date"]); h = h.sort_values("date")
    daily = h[h["src"].astype(str) == "גלובס"]; quarterly = h[h["src"].astype(str) != "גלובס"]
    n_daily = int(len(daily)); base = {}
    floors = {"nominal": floor_bp.get("nominal", 0.0), "real": floor_bp.get("real", 0.0), "bei": min(floor_bp.get("nominal", 0.0), floor_bp.get("real", 0.0))}
    for t in HIST_TENORS:
        rec = {"n_daily": n_daily}
        for kind, fl in floors.items():
            col = {"nominal": "nom", "real": "real", "bei": "bei"}[kind] + f"_{t}"
            dv = _ann_std(daily[col], 250) if n_daily >= 3 else np.nan
            qv = _ann_std(quarterly[col], 4) if len(quarterly) >= 3 else np.nan
            if n_daily >= MIN_DAILY and not np.isnan(dv): v, method = dv, "daily"
            else: v, method = max([x for x in (dv, fl) if not np.isnan(x)] or [np.nan]), "floor"
            rec[f"{kind}_vol_bp"] = round(v, 1) if not np.isnan(v) else np.nan
            rec[f"{kind}_q_bp"] = round(qv, 1) if not np.isnan(qv) else np.nan
        rec["method"] = method; base[t] = rec
    xs = list(HIST_TENORS); out = []
    for t in TENORS:
        row = {"tenor": t}
        for k in ("nominal_vol_bp", "real_vol_bp", "bei_vol_bp"):
            row[k] = round(gov_data.interp(t, xs, [base[x][k] for x in xs]), 1)
        row["n_daily"] = n_daily; row["method"] = base[xs[0]]["method"]
        for k in ("nominal_q_bp", "real_q_bp", "bei_q_bp"):
            vals = [base[x][k] for x in xs]
            row[k.replace("_q_bp", "_quarterly_bp")] = round(gov_data.interp(t, xs, vals), 1) if not any(np.isnan(v) for v in vals) else np.nan
        out.append(row)
    return pd.DataFrame(out, columns=["tenor", "nominal_vol_bp", "real_vol_bp", "bei_vol_bp", "n_daily", "method", "nominal_quarterly_bp", "real_quarterly_bp", "bei_quarterly_bp"])


if __name__ == "__main__":
    data = Path(sys.argv[1] if len(sys.argv) > 1 else "data")
    bh = build_bond_history(data); bh.to_csv(data / "gov" / "bond_history.csv", index=False, encoding="utf-8-sig")
    ch = pd.read_csv(data / "gov" / "curves_history.csv", encoding="utf-8-sig")
    vol = build_vol(ch, {"nominal": 60.0, "real": 50.0}); vol.to_csv(data / "gov" / "vol.csv", index=False, encoding="utf-8-sig")
    print(f"bond_history: {len(bh)} שורות, {bh['date'].nunique()} ימים; vol: {vol.iloc[0]['method']}, n_daily={vol.iloc[0]['n_daily']}")
