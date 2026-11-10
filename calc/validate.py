"""validate.py — אימות שיטת החישוב מול נתוני גלובס: לסדרות עם מח"מ ותשואה בגלובס ועם לוח מפורש במודל (validation_schedules.csv),
משווה את התשואה והמח"מ המחושבים מהלוח לאלה של גלובס בכל ימי data/history. הרצה: python calc/validate.py data"""
from __future__ import annotations
import datetime as dt
import sys
from pathlib import Path
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from calc.bond_calc import load_schedules, price_calc


def run(data="data", sched_file=None) -> pd.DataFrame:
    data = Path(data); here = Path(__file__).resolve().parent
    sched = load_schedules(sched_file or here / "validation_schedules.csv")
    h = pd.concat([pd.read_csv(f, dtype=str, encoding="utf-8-sig") for f in sorted((data / "history").glob("*.csv"))], ignore_index=True)
    for c in ("last_price", "duration", "ytm_gross"): h[c] = pd.to_numeric(h[c], errors="coerce")
    rows = []
    for sid, sc in sched.items():
        for r in h[(h.security_id == sid) & h.ytm_gross.notna() & h.duration.notna() & h.last_price.notna()].itertuples():
            res = price_calc(dt.date.fromisoformat(r.date), r.last_price, sc)
            if res: rows.append({"security_id": sid, "name": r.name, "date": r.date, "price": r.last_price, "ytm_globes": r.ytm_gross, "ytm_calc": res[0],
                                 "gap_bp": (res[0] - r.ytm_gross) * 100, "dur_globes": r.duration, "dur_calc": res[1], "dur_gap": res[1] - r.duration})
    return pd.DataFrame(rows)


def main(data="data", sched_file=None):
    v = run(data, sched_file)
    if v.empty: print("אין נתוני אימות"); return 1
    last = v[v.date == v.date.max()]
    s = v.groupby(["security_id", "name"]).agg(n=("date", "count"), gap_bp_mean=("gap_bp", "mean"), gap_bp_maxabs=("gap_bp", lambda x: x.abs().max()),
                                                  dur_gap_mean=("dur_gap", "mean"), dur_gap_maxabs=("dur_gap", lambda x: x.abs().max())).round(2)
    print(s.to_string()); print()
    print("סיכום ל-%d סדרות, %d נקודות: פער תשואה ממוצע %.1f נ״ב, מקס׳ מוחלט %.1f נ״ב; פער מח״מ ממוצע %.3f, מקס׳ מוחלט %.3f" %
          (v.security_id.nunique(), len(v), v.gap_bp.mean(), v.gap_bp.abs().max(), v.dur_gap.mean(), v.dur_gap.abs().max()))
    print("היום האחרון (%s):" % v.date.max()); print(last[["name", "price", "ytm_globes", "ytm_calc", "gap_bp", "dur_globes", "dur_calc"]].round(3).to_string(index=False))
    return 0


if __name__ == "__main__":
    sys.exit(main(*sys.argv[1:]))
