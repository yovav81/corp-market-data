"""bond_calc.py — שכבת חישוב: תשואה, מח"מ והפרש ממשלתי לסדרות קונצרניות ששמגלובס אין להן מח"מ/תשואה (corp_unlinked, "חישוב").

הלוח: calc/schedules.csv (מועתק ממודל המנפיק; ראו gen_schedules.py). סדרה בלי לוח — לא מחושבת.
שיטה (כמו בבלוק החישוב של דיפלומט ובאימות מול גלובס, calc/validate.py): המחיר של גלובס (% מהע.נ. הנותרת) הוא השווי הנוכחי של התשלומים
העתידיים (ללא ריבית צבורה), מועד הערך = יום המסחר, תשואה שנתית אפקטיבית ACT/365, מח"מ = מקוילי. ההפרש — spreads.py: אינטרפולציה
ליניארית לפי מח"מ על העקום הממשלתי השקלי הקבוע של אותו יום.
"""
from __future__ import annotations
import datetime as dt
import sys
from pathlib import Path
import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from short.engine import xirr, macaulay

SCHED = Path(__file__).with_name("schedules.csv")


def load_schedules(path=SCHED) -> dict[str, list[tuple[dt.date, float, float]]]:
    df = pd.read_csv(path, dtype={"security_id": str}, encoding="utf-8-sig")
    out = {}
    for sid, g in df.groupby("security_id"):
        rows = sorted((dt.date.fromisoformat(r.pay_date), float(r.principal_pct), float(r.coupon_pct)) for r in g.itertuples())
        if abs(sum(r[1] for r in rows) - 100) > 0.06:
            raise ValueError(f"{sid}: סך הקרן בלוח {sum(r[1] for r in rows)} ולא 100")
        out[sid] = rows
    return out


def flows_per_100(day: dt.date, rows) -> list[tuple[dt.date, float]]:
    """תזרים עתידי (תאריך > day) ל-100 ע.נ. נותרת בתאריך day; ריבית על היתרה שלפני התשלום."""
    bal, raw, out_now = 100.0, [], None
    for d, p, c in rows:
        if d > day:
            if out_now is None: out_now = bal
            raw.append((d, bal * c / 100.0 + p))
        bal -= p
    if not raw or not out_now or out_now <= 1e-9: return []
    return [(d, v * 100.0 / out_now) for d, v in raw]


def price_calc(day: dt.date, price: float, rows) -> tuple[float, float] | None:
    """(תשואה באחוזים, מח"מ מקוילי בשנים) או None."""
    fl = flows_per_100(day, rows)
    if not fl: return None
    y = xirr([(day, -price)] + fl)
    if y is None: return None
    dur = macaulay(day, fl, y)
    return (y * 100.0, dur) if dur is not None else None


def calc_rows(day: str, raw: pd.DataFrame, sched: dict, gx: np.ndarray, gy: np.ndarray, gov: str,
              max_ytm: float, min_dur: float) -> pd.DataFrame:
    """שורות spreads לסדרות עם לוח ושגלובס לא נתן להן מח"מ או תשואה. raw — שורות ההיסטוריה של היום ושל קבוצת ההצמדה (לפני הסינון)."""
    d0 = dt.date.fromisoformat(day); out = []
    for r in raw.itertuples():
        rows = sched.get(str(r.security_id))
        if rows is None: continue
        if not (pd.isna(r.duration) or pd.isna(r.ytm_gross)): continue
        price = pd.to_numeric(r.last_price, errors="coerce")
        if pd.isna(price) or price <= 0: continue
        res = price_calc(d0, float(price), rows)
        if res is None: continue
        ytm, dur = res
        if not (0 < ytm <= max_ytm) or dur < min_dur: continue
        g = float(np.interp(dur, gx, gy))
        out.append({"date": day, "source": r.source, "security_id": r.security_id, "name": r.name, "duration": round(dur, 2),
                    "ytm_gross": ytm, "gov_ytm": g, "spread": ytm - g, "extrap": int(dur < gx[0] or dur > gx[-1]), "curve": gov,
                    "rating_maalot": r.rating_maalot, "rating_midroog": r.rating_midroog, "calc": 1})
    return pd.DataFrame(out)
