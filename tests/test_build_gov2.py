import shutil
from pathlib import Path
import openpyxl, pandas as pd, pytest
from conftest import ROOT, write_cpi_fixture
import build_gov2, check_gov2

needs_soffice = pytest.mark.skipif(shutil.which("soffice") is None, reason="LibreOffice לא מותקן")
REAL = ROOT / "data"


@pytest.fixture(scope="module")
def real_build(tmp_path_factory):
    """בנייה מנתוני המאגר האמיתיים (data/) בלי Power Query, ואז recalc ב-LibreOffice."""
    d = tmp_path_factory.mktemp("real"); out = d / "gov2.xlsx"
    import gov_history, gov_cpi
    data = d / "data"; shutil.copytree(REAL, data)
    gov_history.build_bond_history(data).to_csv(data / "gov" / "bond_history.csv", index=False, encoding="utf-8-sig")
    gov_history.build_vol(pd.read_csv(data / "gov" / "curves_history.csv", encoding="utf-8-sig"), {"nominal": 60, "real": 50}).to_csv(data / "gov" / "vol.csv", index=False, encoding="utf-8-sig")
    gov_cpi.build_cpi_path(data).to_csv(data / "gov" / "cpi_path.csv", index=False, encoding="utf-8-sig")
    # חו״ל סינתטי (אין רשת בבדיקות) + החזקה אחת ו-"כלול בעקום?" אחד — דרך קובץ v2 קודם מדומה
    import gov_foreign
    pd.DataFrame([{"market": m, "tenor": t, "yield": y0 + 0.08 * t, "rate_1y": y0 + 0.08, "date": "2026-11-04", "source": "בדיקה", "as_of": "2026-11-05", "stale": 0}
                  for m, y0 in (("US", 4.3), ("EUR", 2.0), ("GBP", 3.9)) for t in gov_foreign.TENORS]).to_csv(data / "gov" / "foreign_latest.csv", index=False, encoding="utf-8-sig")
    import gov_decide; gov_decide.decide(data)
    res = build_gov2.build(data, out, None, pq=False)
    rc = check_gov2.recalc(out) if shutil.which("soffice") else None
    return {"out": out, "recalc": rc, "data": data, "res": res}


def test_build_sheets_and_names(tmp_data, monkeypatch):
    write_cpi_fixture(tmp_data)
    import gov_data, gov_foreign
    from test_gov_data_outputs import _patch_network
    _patch_network(monkeypatch); gov_data.main(str(tmp_data))
    out = tmp_data / "g.xlsx"; res = build_gov2.build(tmp_data, out, None, pq=False)
    assert res["sheets"] == build_gov2.ORDER and len(res["sheets"]) == 25
    wb = openpyxl.load_workbook(out)
    dn = wb.defined_names["p_linked_neutral"].attr_text
    sh, cell = dn.split("!"); assert wb[sh.strip("'")][cell.replace("$", "")].value == 0.5
    assert "d_asof" in wb.defined_names and "cur_nom" in wb.defined_names and "b_key" in wb.defined_names
    assert wb["נתוני_אגח"].tables and list(wb["נתוני_אגח"].tables)[0] == "GovBonds"


@needs_soffice
def test_missing_yield_row(tmp_data, monkeypatch):
    write_cpi_fixture(tmp_data)
    import gov_data
    from test_gov_data_outputs import _patch_network
    _patch_network(monkeypatch); gov_data.main(str(tmp_data))
    b = pd.read_csv(tmp_data / "gov" / "bonds_latest.csv", encoding="utf-8-sig")
    b.loc[b["name"] == "ממשל שקלית 0335", ["ytm_gross", "modified", "convexity"]] = None
    b.to_csv(tmp_data / "gov" / "bonds_latest.csv", index=False, encoding="utf-8-sig")
    out = tmp_data / "g.xlsx"; build_gov2.build(tmp_data, out, None, pq=False)
    rc = check_gov2.recalc(out)
    assert sum(check_gov2.count_errors(rc).values()) == 0
    wb = openpyxl.load_workbook(rc, data_only=True); ws = wb["02_קלט_אגח"]
    row = next(r for r in range(6, 126) if ws[f"D{r}"].value == "ממשל שקלית 0335")
    assert ws[f"{build_gov2.CC['status']}{row}"].value == "אין תשואה"
    assert ws[f"{build_gov2.CC['xn']}{row}"].value in (None, "")          # לא נכנס לעקום


@needs_soffice
def test_recalc_zero_errors(real_build):
    errs = check_gov2.count_errors(real_build["recalc"])
    assert sum(errs.values()) == 0, errs


@needs_soffice
def test_parity_curves(real_build):
    assert check_gov2.parity_curves(real_build["recalc"], real_build["data"]) <= 0.5


# ---------- Task 8: 04 שקלי_צמוד + 05 מחמ ----------
@needs_soffice
def test_parity_linked_and_duration(real_build):
    import gov_decide
    res = check_gov2.parity_decisions(real_build["recalc"], real_build["data"], gov_decide.load_settings())
    assert set(res) >= {"linked_weight", "dur_nominal", "dur_real", "dur_weighted"}
    for k, (xl, py, ok) in res.items(): assert ok, (k, xl, py)
    wb = openpyxl.load_workbook(real_build["recalc"], data_only=True)
    why = check_gov2.named_value(wb, "m_linked_why"); assert isinstance(why, str) and "משקל צמוד" in why
    assert isinstance(check_gov2.named_value(wb, "m_dur_nominal_why"), str)


@needs_soffice
def test_sim_override(real_build, tmp_path):
    wb = openpyxl.load_workbook(real_build["out"])
    for nm, v in (("p_sim_linked", 0.6), ("p_sim_dur_nominal", 3.3)):
        sh, cell = wb.defined_names[nm].attr_text.split("!"); wb[sh.strip("'")][cell.replace("$", "")].value = v
    p = tmp_path / "sim.xlsx"; wb.save(p); rc = check_gov2.recalc(p)
    wb2 = openpyxl.load_workbook(rc, data_only=True)
    assert check_gov2.named_value(wb2, "u_linked") == pytest.approx(0.6)
    assert check_gov2.named_value(wb2, "m_linked") == pytest.approx(check_gov2.named_value(openpyxl.load_workbook(real_build["recalc"], data_only=True), "m_linked"))
    assert check_gov2.named_value(wb2, "u_dur_nominal") == pytest.approx(3.3)
    assert sum(check_gov2.count_errors(rc).values()) == 0


# ---------- Task 9: 06 סינטטי_בולט + 07 סדרות ----------
@needs_soffice
def test_pairs_table_weights(real_build):
    wb = openpyxl.load_workbook(real_build["recalc"], data_only=True); ws = wb["06_סינטטי_בולט"]
    checked = 0
    for kind in ("nominal", "real"):
        blk = build_gov2.STRUCT_BLOCKS[kind]; target = check_gov2.named_value(wb, f"u_dur_{kind}")
        for r in range(blk["pairs_r0"], blk["pairs_r0"] + build_gov2.NPAIRS):
            if ws[f"{blk['c']['valid']}{r}"].value == "כן":
                wi, wj = ws[f"{blk['c']['wi']}{r}"].value, ws[f"{blk['c']['wj']}{r}"].value
                di, dj = ws[f"{blk['c']['di']}{r}"].value, ws[f"{blk['c']['dj']}{r}"].value
                assert abs(wi + wj - 1) < 1e-9 and abs(wi * di + wj * dj - target) <= 0.01; checked += 1
    assert checked > 10


@needs_soffice
def test_parity_structures(real_build):
    res = check_gov2.parity_structures(real_build["recalc"], real_build["data"])
    for k, (xl, py, ok) in res.items(): assert ok, (k, xl, py)


@needs_soffice
def test_parity_top5(real_build):
    res = check_gov2.parity_top5(real_build["recalc"], real_build["data"])
    for k, (xl, py, ok) in res.items(): assert ok, (k, xl, py)


@needs_soffice
def test_scan_no_long_bond_excel(real_build, tmp_path):
    wb = openpyxl.load_workbook(real_build["out"])
    sh, cell = wb.defined_names["p_sim_dur_real"].attr_text.split("!"); wb[sh.strip("'")][cell.replace("$", "")].value = 25   # הצמודה הארוכה ביותר במח״מ ~21
    p = tmp_path / "nolong.xlsx"; wb.save(p); rc = check_gov2.recalc(p)
    wb2 = openpyxl.load_workbook(rc, data_only=True)
    assert str(check_gov2.named_value(wb2, "m_struct_real")).startswith("אין זוג מתאים")
    assert sum(check_gov2.count_errors(rc).values()) == 0



# ---------- Task 10: 08 חו״ל, 09 תיק ותרחישים, 00 דשבורד ----------
@needs_soffice
def test_holding_not_found(real_build, tmp_path):
    wb = openpyxl.load_workbook(real_build["out"]); ws = wb["09_תיק_ותרחישים"]
    ws["A5"] = "1194802"; ws["B5"] = 1_000_000          # ממשל שקלית 0229
    ws["A6"] = "999"; ws["B6"] = 500_000                  # לא קיים
    p = tmp_path / "hold.xlsx"; wb.save(p); rc = check_gov2.recalc(p)
    assert sum(check_gov2.count_errors(rc).values()) == 0
    wb2 = openpyxl.load_workbook(rc, data_only=True)
    assert wb2["09_תיק_ותרחישים"]["D6"].value == "לא נמצא" and check_gov2.named_value(wb2, "n_not_found") == 1
    assert check_gov2.named_value(wb2, "cur_w_linked") == pytest.approx(0.0)        # ההחזקה היחידה שנמצאה היא שקלית
    assert abs(check_gov2.named_value(wb2, "cur_dur") - check_gov2.named_value(wb2, "cur_dur")) < 1e-9 and check_gov2.named_value(wb2, "cur_dur") > 1
    assert check_gov2.named_value(wb2, "cur_mv") == pytest.approx(1_000_000 * wb2["02_קלט_אגח"][f"{build_gov2.BC['last_price']}{[r for r in range(6, 126) if wb2['02_קלט_אגח'][f'C{r}'].value == '1194802'][0]}"].value / 100, rel=1e-6)


@needs_soffice
def test_scenarios_sign(real_build, tmp_path):
    wb = openpyxl.load_workbook(real_build["out"]); ws = wb["09_תיק_ותרחישים"]
    ws["A5"] = "1194802"; ws["B5"] = 1_000_000; ws["A6"] = "1220722"; ws["B6"] = 1_000_000
    p = tmp_path / "scen.xlsx"; wb.save(p); rc = check_gov2.recalc(p)
    wb2 = openpyxl.load_workbook(rc, data_only=True)
    sc = check_gov2.scenario_table(wb2)
    for port in ("current", "model"):
        assert sc[("parallel_up_2", port)] < 0 < sc[("parallel_dn_2", port)]
        assert sc[("parallel_up_1", port)] < 0 < sc[("parallel_dn_1", port)]
    assert sc[("infl_up", "current")] > 0 > sc[("infl_dn", "current")]       # יש צמוד בתיק הנוכחי


@needs_soffice
def test_parity_foreign(real_build):
    res = check_gov2.parity_foreign(real_build["recalc"], real_build["data"])
    assert set(res) == {"foreign_5y", "foreign_10y"}
    for k, (xl, py, ok) in res.items(): assert ok, (k, xl, py)


@needs_soffice
def test_dashboard_keys(real_build):
    wb = openpyxl.load_workbook(real_build["recalc"], data_only=True); ws = wb["00_דשבורד"]
    rows = [r for r in range(1, 13) if ws[f"A{r}"].value in ("שקלי / צמוד", "מח״מ", "סינתטי / בולט", "סדרות מובילות", "חו״ל", "סיכון")]
    assert len(rows) == 6
    for r in rows: assert ws[f"B{r}"].value not in (None, ""), ws[f"A{r}"].value
    assert sum(check_gov2.count_errors(real_build["recalc"]).values()) == 0
