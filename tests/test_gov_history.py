import pandas as pd
from conftest import history_rows, curves_history_rows, DATES, BONDS
import gov_history


def test_bond_history_rows(tmp_data):
    df = gov_history.build_bond_history(tmp_data)
    assert len(df) == len(DATES) * len(BONDS)
    assert list(df.columns) == ["date", "security_id", "key", "name", "kind", "ttm", "ytm", "duration", "volume_k_nis", "curve_ytm", "spread_bp"]
    core = df[df["kind"].isin(["שקלי", "מדד"])]
    assert core["spread_bp"].notna().all()
    # כל הניירות על קו ישר → העקום (על רשת הטווחים) עובר דרכם → פער ≈ 0 בתוך הרשת; מעבר לנייר האחרון הרשת שטוחה (ערך קצה) ולכן שם מותר פער קטן
    inner = core[(core["ttm"] >= 1) & (core["ttm"] <= 15)]
    assert inner["spread_bp"].abs().max() < 2.0   # הנייר הקצר ביותר באפיק מקבל ערך קצה מתחתיו — עד ~1.5 נ"ב
    assert core["spread_bp"].abs().max() < 10


def test_bond_history_key_from_static_when_id_missing(tmp_data):
    h = history_rows(); h["security_id"] = None            # כמו בגלובס: בלי מס' ני"ע
    h.to_csv(tmp_data / "history" / "2026.csv", index=False, encoding="utf-8-sig")
    df = gov_history.build_bond_history(tmp_data)
    assert (df.loc[df["name"] == "ממשל שקלית 0330", "key"] == "1160985").all()     # מהמאגר הסטטי
    assert df["security_id"].notna().all()


def test_bond_history_perturbed_bond_is_cheap(tmp_data):
    history_rows(perturb={"ממשל שקלית 0335": 0.20}).to_csv(tmp_data / "history" / "2026.csv", index=False, encoding="utf-8-sig")
    df = gov_history.build_bond_history(tmp_data)
    b = df[df["name"] == "ממשל שקלית 0335"]
    assert (b["spread_bp"] > 3).all()   # הרשת נמשכת אחרי הנייר, לכן פחות מ-20 נ"ב


def test_vol_daily_method():
    hist = curves_history_rows(70)
    vol = gov_history.build_vol(hist, {"nominal": 60.0, "real": 50.0})
    assert list(vol.columns)[:6] == ["tenor", "nominal_vol_bp", "real_vol_bp", "bei_vol_bp", "n_daily", "method"]
    r2 = vol[vol["tenor"] == 2].iloc[0]
    assert r2["method"] == "daily" and r2["n_daily"] == 70
    assert r2["nominal_vol_bp"] > 0 and r2["real_vol_bp"] > 0
    assert set(vol["tenor"]) == {0.25, 0.5, 1, 2, 3, 4, 5, 7, 10, 15, 20, 30}


def test_vol_floor():
    hist = curves_history_rows(5)
    vol = gov_history.build_vol(hist, {"nominal": 60.0, "real": 50.0})
    r2 = vol[vol["tenor"] == 2].iloc[0]
    assert r2["method"] == "floor"
    assert r2["nominal_vol_bp"] == 60.0 and r2["real_vol_bp"] == 50.0      # עד 60 ימים — הרצפה בלבד
    assert r2["nominal_quarterly_bp"] > 0 and r2["nominal_daily_bp"] > 0    # האומדנים מדווחים לעיון בלבד
