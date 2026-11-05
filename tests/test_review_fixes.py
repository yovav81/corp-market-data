"""בדיקות לממצאי הביקורת (fix pass): כל בדיקה נכתבה לפני התיקון."""
import json, shutil
import numpy as np, pandas as pd, pytest, openpyxl
from conftest import ROOT, write_cpi_fixture, history_rows, BONDS
import gov_decide, gov_foreign, gov_history, build_gov2, check_gov2

S = gov_decide.load_settings()
needs_soffice = pytest.mark.skipif(shutil.which("soffice") is None, reason="LibreOffice לא מותקן")


def _curve():
    rows = []
    for t in gov_decide.TENORS:
        n, r = 3.0 + 0.15 * t, 1.0 + 0.08 * t
        rows.append({"tenor": t, "nominal": n, "real": r, "bei": ((1 + n / 100) / (1 + r / 100) - 1) * 100, "roll_nominal_bp": 10.0, "roll_real_bp": 8.0, "fwd_nominal": n, "fwd_real": r})
    return pd.DataFrame(rows)


# (3) חו״ל: מעט שווקים טריים → "נתון ישן", וגיל הנתון
def test_foreign_verdict_stale_when_too_few_fresh_markets():
    cur = _curve()
    fo = pd.DataFrame([{"market": m, "tenor": 5, "yield": 6.0, "rate_1y": 4.0, "date": "2026-11-04", "source": "x", "as_of": "2026-11-05", "stale": st}
                       for m, st in (("US", 1), ("EUR", 1), ("GBP", 0))])
    v = gov_decide.foreign_verdict(fo, cur, S, asof="2026-11-05")
    assert (v["verdict"] == "נתון ישן").all()                      # רק שוק טרי אחד < 2
    fo["stale"] = 0; fo.loc[fo["market"] == "US", "date"] = "2026-10-01"   # ישן לפי תאריך (> 7 ימים) גם בלי דגל
    v = gov_decide.foreign_verdict(fo, cur, S, asof="2026-11-05")
    assert v.loc[v["market"] == "US", "stale"].iloc[0] == 1 and (v["verdict"] == "חו״ל מגודר").all()   # EUR+GBP טריים ועוברים


# (4) תחזית שער באחוזים (לא פורמט %)
def test_fx_forecast_units():
    keys = {x["key"]: x for x in json.loads((ROOT / "gov" / "settings_default.json").read_text(encoding="utf-8"))}
    for k in ("fx_usd", "fx_eur", "fx_gbp"): assert keys[k]["fmt"] == "0.0", k


# (5) אופק קבוע = שנה: אין פרמטר horizon_years
def test_horizon_locked():
    keys = {x["key"] for x in json.loads((ROOT / "gov" / "settings_default.json").read_text(encoding="utf-8"))}
    assert "horizon_years" not in keys
    assert "horizon_years" not in (ROOT / "gov" / "build_gov2.py").read_text(encoding="utf-8").replace("p_horizon_years", "")   # אין שימוש בפרמטר


# (6) רשימת הניירות לסריקה מוגבלת ל-NLIST גם בפייתון (זהות), ואזהרה כשחורגים
def test_structure_list_cap_mirrors_excel():
    rows = []
    for i in range(45):
        ttm = 0.5 + i * 0.6; y = 3.0 + 0.1 * ttm; d = (1 - (1 + y / 100) ** -ttm) / (y / 100)
        rows.append({"security_id": 1000 + i, "key": str(1000 + i), "name": f"נ{i}", "kind": "שקלי", "ttm": ttm, "ytm_gross": y, "modified": d, "convexity": d * d, "volume_k_nis": 5000, "in_curve": ""})
    b = pd.DataFrame(rows)
    seg = gov_decide._segment(b, "nominal")
    assert len(seg) == build_gov2.NLIST and seg["modified"].max() < b["modified"].max()
    assert gov_decide.segment_overflow(b, "nominal") == 45 - build_gov2.NLIST


# (7) מפתח כטקסט בכל מקום
def test_key_read_as_text(tmp_data, monkeypatch):
    write_cpi_fixture(tmp_data)
    import gov_data
    from test_gov_data_outputs import _patch_network
    _patch_network(monkeypatch); gov_data.main(str(tmp_data))
    ctx = build_gov2.Ctx(tmp_data, {}, r"C:\x")
    assert ctx.frames["נתוני_אגח"]["key"].map(type).eq(str).all() and ctx.frames["נתוני_סדרות_היסט"]["key"].map(type).eq(str).all()
    m = build_gov2.pq_schema(ctx)
    assert ("key", "text") in m["GovBonds"] and ("key", "text") in m["GovBondHist"]


@needs_soffice
def test_holding_typed_as_number_is_found(tmp_data, monkeypatch):
    write_cpi_fixture(tmp_data)
    import gov_data
    from test_gov_data_outputs import _patch_network
    _patch_network(monkeypatch); gov_data.main(str(tmp_data))
    out = tmp_data / "g.xlsx"; build_gov2.build(tmp_data, out, None, pq=False)
    wb = openpyxl.load_workbook(out); ws = wb["09_תיק_ותרחישים"]; ws["A5"] = 1160985; ws["B5"] = 1_000_000     # מספר, לא טקסט
    wb.save(out); rc = check_gov2.recalc(out); wb2 = openpyxl.load_workbook(rc, data_only=True)
    assert check_gov2.named_value(wb2, "cur_mv") > 0 and check_gov2.named_value(wb2, "n_not_found") == 0


# (8) חלון היסטוריה מתגלגל
def test_bond_history_window(tmp_data):
    import datetime as dt
    base = dt.date(2025, 1, 1); frames = []
    for i in range(420):
        d = base + dt.timedelta(days=i)
        if d.weekday() >= 5: continue
        h = history_rows(); h["date"] = d.isoformat(); frames.append(h)
    pd.concat(frames).to_csv(tmp_data / "history" / "2026.csv", index=False, encoding="utf-8-sig")
    df = gov_history.build_bond_history(tmp_data, window_days=250)
    assert df["date"].nunique() == 250 and df["date"].min() > "2025-01-01"
    df_all = gov_history.build_bond_history(tmp_data, window_days=None)
    assert df_all["date"].nunique() > 250


# (9) Z באקסל מתעלם מימים בלי תשואה (כמו dropna בפייתון)
@needs_soffice
def test_series_history_ignores_blank_spreads(tmp_data, monkeypatch):
    write_cpi_fixture(tmp_data)
    import gov_data
    from test_gov_data_outputs import _patch_network
    _patch_network(monkeypatch); gov_data.main(str(tmp_data))
    bh = pd.read_csv(tmp_data / "gov" / "bond_history.csv", encoding="utf-8-sig", dtype={"key": str})
    one = bh[bh["key"] == "1160985"].head(1)
    extra = pd.concat([one] * 5, ignore_index=True); extra["date"] = ["2026-10-20", "2026-10-21", "2026-10-22", "2026-10-23", "2026-10-26"]; extra["spread_bp"] = np.nan; extra["ytm"] = np.nan
    pd.concat([bh, extra]).to_csv(tmp_data / "gov" / "bond_history.csv", index=False, encoding="utf-8-sig")
    out = tmp_data / "g.xlsx"; build_gov2.build(tmp_data, out, None, pq=False); rc = check_gov2.recalc(out)
    wb2 = openpyxl.load_workbook(rc, data_only=True); ws = wb2["07_סדרות"]
    row = next(r for r in range(34, 154) if ws[f"B{r}"].value == "1160985")
    assert ws[f"H{row}"].value == 3            # רק 3 ימים עם פער (לא 8)


# (10) העברת פרמטרים מ-v2: רק מה שהמשתמש שינה; ברירת מחדל חדשה מגיעה
def test_carry_over_v2_keeps_only_user_changes(tmp_data, monkeypatch):
    write_cpi_fixture(tmp_data)
    import gov_data
    from test_gov_data_outputs import _patch_network
    _patch_network(monkeypatch); gov_data.main(str(tmp_data))
    a = tmp_data / "a.xlsx"; build_gov2.build(tmp_data, a, None, pq=False)
    wb = openpyxl.load_workbook(a)
    sh, cell = wb.defined_names["p_tilt_max"].attr_text.split("!"); wb[sh.strip("'")][cell.replace("$", "")].value = 0.2   # המשתמש שינה
    wb.save(a)
    co = build_gov2.carry_over(a)
    assert co["params"] == {"tilt_max": 0.2}                         # bb_tol וכל השאר לא "קפאו"


# (2) decide עם הגדרות החוברת (כולל הוצאה מהעקום ותחזית לפי טווח)
def test_decide_with_workbook_overrides(tmp_data, monkeypatch):
    write_cpi_fixture(tmp_data)
    import gov_data
    from test_gov_data_outputs import _patch_network
    _patch_network(monkeypatch); gov_data.main(str(tmp_data))
    base = gov_decide.decide(tmp_data)
    s = gov_decide.load_settings(overrides={"linked_min": 0.5, "linked_max": 0.5, "exclude": {"1160985"}, "infl_source": "תחזית מנהל השקעות", "infl_user": {t: 4.0 for t in gov_decide.TENORS}})
    res = gov_decide.decide(tmp_data, s)
    assert res["linked_weight"] == 0.5 and "1160985" not in res["top5_nominal"]
    assert "תחזית מנהל השקעות" in pd.read_csv(tmp_data / "gov" / "decisions_latest.csv", encoding="utf-8-sig").set_index("key").loc["linked_weight", "why"]
    assert callable(gov_decide.overrides_from_workbook)
