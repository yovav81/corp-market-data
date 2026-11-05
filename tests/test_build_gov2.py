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
