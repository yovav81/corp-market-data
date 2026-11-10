"""בדיקות שכבת החישוב (calc/): לוח תשלומים -> תשואה, מח"מ והפרש ממשלתי; גלובס לא נדרס."""
import datetime as dt
import sys
from pathlib import Path
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from calc.bond_calc import load_schedules, price_calc, flows_per_100, calc_rows


def test_schedules_sum_to_100():
    for sid, rows in load_schedules().items():
        assert abs(sum(r[1] for r in rows) - 100) < 0.06, sid


def test_par_bond_yield_equals_coupon():
    # אג"ח בולט, ריבית 4% שנתית, מחיר 100 בתאריך תשלום ריבית -> תשואה ~ 4%
    rows = [(dt.date(2027, 12, 31), 0.0, 4.0), (dt.date(2028, 12, 31), 100.0, 4.0)]
    y, d = price_calc(dt.date(2026, 12, 31), 100.0, rows)
    assert abs(y - 4.0) < 0.02 and 1.9 < d < 2.0


def test_partial_principal_normalised_to_remaining():
    rows = [(dt.date(2026, 6, 30), 50.0, 0.0), (dt.date(2027, 6, 30), 50.0, 5.0)]
    fl = flows_per_100(dt.date(2026, 11, 1), rows)
    assert len(fl) == 1 and abs(fl[0][1] - 105.0) < 1e-6   # (50 + 50*5%) / 50 * 100


def test_diplomat_case_study():
    h = pd.read_csv(ROOT / "data/history/2026.csv", dtype=str, encoding="utf-8-sig")
    r = h[(h.security_id == "1248780") & (h.date == "2026-11-10")].iloc[0]
    assert pd.isna(pd.to_numeric(r.ytm_gross, errors="coerce"))
    y, d = price_calc(dt.date(2026, 11, 10), float(r.last_price), load_schedules()["1248780"])
    assert round(y, 2) == 4.66 and round(d, 2) == 4.99


def test_globes_values_never_overwritten():
    sched = load_schedules()
    raw = pd.DataFrame([{"source": "corp_unlinked", "security_id": "1248780", "name": "x", "last_price": 99.57, "duration": 4.0, "ytm_gross": 5.0,
                         "rating_maalot": None, "rating_midroog": None}])
    gx, gy = np.array([1.0, 8.0]), np.array([3.0, 4.0])
    assert calc_rows("2026-11-10", raw, sched, gx, gy, "gov_unlinked_fixed", 50.0, 0.3).empty
    raw["ytm_gross"] = np.nan
    k = calc_rows("2026-11-10", raw, sched, gx, gy, "gov_unlinked_fixed", 50.0, 0.3)
    assert len(k) == 1 and k.iloc[0]["calc"] == 1 and abs(k.iloc[0]["spread"] - (k.iloc[0]["ytm_gross"] - k.iloc[0]["gov_ytm"])) < 1e-9


def test_validation_against_globes():
    sys.path.insert(0, str(ROOT / "calc"))
    import validate
    v = validate.run(str(ROOT / "data"))
    assert v.security_id.nunique() >= 9 and v.gap_bp.abs().max() < 5 and v.dur_gap.abs().max() < 0.02
