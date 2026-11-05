import io, zipfile
import openpyxl
from openpyxl.worksheet.table import Table
from conftest import ROOT
import pq_gov, build_live

KIT = ROOT / "corp" / "pq_kit"
BASE = r"C:\CORP\03_Market\data"


def test_gov_queries_cover_all_gov_files():
    assert set(pq_gov.GOV_QUERIES) >= {"GovBonds", "GovCurves", "GovHist", "GovBondHist", "GovCpi", "GovVol", "GovForeign", "GovBoiCurves", "GovBoiExp", "GovPolicy", "GovDecisions", "GovUS"}
    assert pq_gov.GOV_QUERIES["GovDecisions"] == "gov_decisions_history.csv"
    assert all(v.startswith("gov_") and v.endswith(".csv") for v in pq_gov.GOV_QUERIES.values())


def test_section_m_has_one_shared_per_query():
    m = pq_gov.section_m(BASE, pq_gov.GOV_QUERIES, {"GovVol": [("tenor", "number"), ("method", "text")]})
    assert m.startswith("section Section1;")
    for q, f in pq_gov.GOV_QUERIES.items():
        assert f"shared {q} = let" in m and f"{BASE}\\{f}" in m
    assert '{"tenor", type number}' in m and '{"method", type text}' in m


def test_roundtrip_queries():
    m = pq_gov.section_m(BASE, pq_gov.GOV_QUERIES, {})
    mashup, props, conns = pq_gov.build_mashup(m, {q: f"נתוני_{q}" for q in pq_gov.GOV_QUERIES}, KIT)
    d = pq_gov.decode_mashup(mashup)
    assert d["queries"] == list(pq_gov.GOV_QUERIES)
    assert d["section"] == m
    assert d["connections"] == 0           # connections.xml נפרד
    assert mashup[:2] == b"\xff\xfe" and b"D\x00a\x00t\x00a\x00M\x00a\x00s\x00h\x00u\x00p\x00" in mashup
    assert b"datastoreItem" in props


def test_connections_count():
    m = pq_gov.section_m(BASE, pq_gov.GOV_QUERIES, {})
    _, _, conns = pq_gov.build_mashup(m, {q: f"נתוני_{q}" for q in pq_gov.GOV_QUERIES}, KIT)
    s = conns.decode("utf-8")
    assert s.count("<connection ") == len(pq_gov.GOV_QUERIES) and s.count('refreshOnLoad="1"') == len(pq_gov.GOV_QUERIES)
    assert 'Location=GovBonds;' in s and 'command="SELECT * FROM [GovBonds]"' in s


def test_inject_into_workbook(tmp_path):
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    sheet_to_query = {}
    for i, q in enumerate(pq_gov.GOV_QUERIES, 1):
        ws = wb.create_sheet(f"נתוני_{q}"); ws.append(["col_a", "col_b"]); ws.append([1, "x"])
        ws.add_table(Table(displayName=q, ref="A1:B2")); sheet_to_query[ws.title] = (q, i)
    src = tmp_path / "in.xlsx"; wb.save(src)
    m = pq_gov.section_m(BASE, pq_gov.GOV_QUERIES, {})
    mashup, props, conns = pq_gov.build_mashup(m, {q: f"נתוני_{q}" for q in pq_gov.GOV_QUERIES}, KIT)
    out = tmp_path / "out.xlsx"
    build_live.inject_pq(str(src), str(out), mashup, props, conns, queries=sheet_to_query)
    names = zipfile.ZipFile(out).namelist()
    assert "xl/connections.xml" in names and "customXml/item1.xml" in names
    assert sum(n.startswith("xl/queryTables/queryTable") for n in names) == len(pq_gov.GOV_QUERIES)
