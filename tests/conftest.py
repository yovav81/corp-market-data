"""Fixtures משותפים לבדיקות gov/. מוסיף את gov/ ל-sys.path כדי שהמודולים ייובאו כמו בהרצת סקריפט (import gov_data)."""
import sys, datetime as dt
from pathlib import Path
import pandas as pd, pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "gov"))
sys.path.insert(0, str(ROOT / "corp"))

HIST_COLS = ["date", "snapshot", "source", "security_id", "name", "last_trade_time", "last_price", "chg_pct", "chg_agorot",
             "volume_k_nis", "duration", "ytm_gross", "ytm_net", "coupon", "adjusted_value", "maturity", "rating_maalot", "rating_midroog"]

# 9 סדרות: 4 שקלי, 3 צמוד, 2 מק"מ — מועדי פדיון אמיתיים
BONDS = [
    ("gov_unlinked_fixed", 1139344, "ממשל שקלית 0327", "2027-03-31", 1.0),
    ("gov_unlinked_fixed", 1160985, "ממשל שקלית 0330", "2030-03-31", 2.25),
    ("gov_unlinked_fixed", 1202332, "ממשל שקלית 0335", "2035-03-30", 1.5),
    ("gov_unlinked_fixed", 1140193, "ממשל שקלית 0347", "2047-03-31", 3.75),
    ("gov_cpi", 1157023, "ממשל צמודה 0529", "2029-05-31", 0.5),
    ("gov_cpi", 1204379, "ממשל צמודה 1033", "2033-10-31", 0.75),
    ("gov_cpi", 1134865, "ממשל צמודה 0545", "2045-05-31", 1.0),
    ("gov_makam", 8270415, "מ.ק.מ. 417", "2027-04-07", 0.0),
    ("gov_makam", 8271017, "מ.ק.מ. 1017", "2027-10-06", 0.0),
]
DATES = ["2026-10-13", "2026-10-14", "2026-10-15"]


def line_yield(source: str, ttm: float) -> float:
    """תשואה על קו ישר לפי שנים לפדיון — שקלי/מק"מ: 3.0 + 0.12·ttm ; צמוד: 1.0 + 0.06·ttm."""
    return round((1.0 + 0.06 * ttm) if source == "gov_cpi" else (3.0 + 0.12 * ttm), 4)


def history_rows(perturb: dict | None = None):
    rows = []
    for d in DATES:
        val = dt.date.fromisoformat(d) + dt.timedelta(days=1)
        for src, sid, name, mat, cpn in BONDS:
            ttm = (dt.date.fromisoformat(mat) - val).days / 365.0
            y = line_yield(src, ttm) + (perturb or {}).get(name, 0.0)
            rows.append({"date": d, "snapshot": f"{d} 18:30", "source": src, "security_id": sid, "name": name, "last_trade_time": "17:24",
                         "last_price": 100.0, "chg_pct": 0.0, "chg_agorot": 0.0, "volume_k_nis": 1000 + sid % 1000, "duration": round(ttm * 0.9, 2),
                         "ytm_gross": y, "ytm_net": y * 0.85, "coupon": cpn, "adjusted_value": 100.0, "maturity": mat, "rating_maalot": "", "rating_midroog": ""})
    return pd.DataFrame(rows, columns=HIST_COLS)


def curves_history_rows(n_daily: int, start="2026-01-05"):
    """n_daily שורות יומיות (src=גלובס) + 4 נקודות רבעוניות. שינוי יומי מתחלף ±0.02 בשקלי, ±0.01 בריאלי."""
    rows = []; d0 = dt.date.fromisoformat(start); nom = {1: 3.1, 2: 3.4, 5: 3.9, 10: 4.2}; real = {1: 1.2, 2: 1.4, 5: 1.8, 10: 2.2}
    for i in range(n_daily):
        s = 1 if i % 2 == 0 else -1
        row = {"date": (d0 + dt.timedelta(days=i)).isoformat(), "src": "גלובס"}
        for t in (1, 2, 5, 10):
            n = nom[t] + 0.02 * s; r = real[t] + 0.01 * s
            row[f"nom_{t}"] = round(n, 4); row[f"real_{t}"] = round(r, 4); row[f"bei_{t}"] = round(((1 + n / 100) / (1 + r / 100) - 1) * 100, 4)
        rows.append(row)
    for j, (dq, bump) in enumerate(zip(["2023-01-05", "2023-03-05", "2023-06-05", "2023-09-05"], [0.0, 0.6, -0.3, 0.2])):
        row = {"date": dq, "src": "מגמה (רבעוני)"}
        for t in (1, 2, 5, 10):
            n = nom[t] + bump; r = real[t] + bump / 2
            row[f"nom_{t}"] = round(n, 4); row[f"real_{t}"] = round(r, 4); row[f"bei_{t}"] = round(((1 + n / 100) / (1 + r / 100) - 1) * 100, 4)
        rows.append(row)
    return pd.DataFrame(rows).sort_values("date")


@pytest.fixture
def tmp_data(tmp_path):
    d = tmp_path / "data"
    (d / "history").mkdir(parents=True); (d / "hist").mkdir(); (d / "gov").mkdir(); (d / "short").mkdir()
    history_rows().to_csv(d / "history" / "2026.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame([{"security_id": sid, "name": name, "linkage": "cpi" if src == "gov_cpi" else "ils", "linked_asset": "מדד" if src == "gov_cpi" else "שקלי",
                   "rating_low": "", "rating_maalot": "", "rating_midroog": "", "listed_capital": 1e9, "is_gov": 1} for src, sid, name, _, _ in BONDS]
                 ).to_csv(d / "hist" / "bond_static.csv", index=False, encoding="utf-8-sig")
    curves_history_rows(70).to_csv(d / "gov" / "curves_history.csv", index=False, encoding="utf-8-sig")
    # latest.csv = היום האחרון של ההיסטוריה
    h = history_rows(); h[h["date"] == DATES[-1]].to_csv(d / "latest.csv", index=False, encoding="utf-8-sig")
    return d


def write_cpi_fixture(d, n_known=4, n_exp=18):
    """data/short/cpi_expected.csv מלאכותי: n_known חודשים ידועים ו-n_exp צפויים עם עונתיות סביב 0.15 לחודש."""
    rows = []; idx = 100.0
    for i in range(n_known):
        pct = [0.3, -0.2, 0.4, 0.5][i % 4]; idx = round(idx * (1 + pct / 100), 2)
        rows.append({"ym": f"2026-{i+1:02d}", "pub_date": f"2026-{i+2:02d}-15", "kind": "ידוע", "index": idx, "monthly_pct": pct, "cum_factor": "", "source": "הלמ״ס", "as_of": "2026-11-05"})
    cum = 1.0
    for j in range(n_exp):
        m = n_known + j; y, mo = 2026 + m // 12, m % 12 + 1
        pct = round(0.15 + 0.25 * ((j % 4) - 1.5), 3); cum *= 1 + pct / 100
        rows.append({"ym": f"{y}-{mo:02d}", "pub_date": "", "kind": "צפוי", "index": "", "monthly_pct": pct, "cum_factor": round(cum, 6), "source": "BEI+עונתיות", "as_of": "2026-11-05"})
    (d / "short").mkdir(exist_ok=True)
    pd.DataFrame(rows).to_csv(d / "short" / "cpi_expected.csv", index=False, encoding="utf-8-sig")
