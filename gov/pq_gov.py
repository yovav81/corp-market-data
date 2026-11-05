"""pq_gov.py — ערכת Power Query לקובץ הממשלתי v2: שאילתה לכל קובץ data/gov → C:\\CORP\\03_Market\\data\\gov_*.csv.

build_mashup(section, targets, kit_dir) -> (DataMashup.xml bytes [utf-16], itemProps1.xml bytes, connections.xml bytes)
    הקוד M מוחלף בתוך חבילת ה-QDEFF של corp/pq_kit (Version | PackageParts(zip) | Permissions | Metadata | PermissionBindings):
    - Formulas/Section1.m ← section ; Config/Package.xml ו-[Content_Types].xml נשארים.
    - Permissions — כמו בערכה. PermissionBindings — ריק (אופציונלי לפי MS-QDEFF).
    - Metadata — נבנה מחדש: פריט AllFormulas + פריט לכל שאילתה (QueryID חדש, FillTarget = שם השאילתה, RecoveryTargetSheet = גיליון היעד) + פריט לכל צעד.
decode_mashup(xml_bytes) -> {"section", "queries", "connections"} — לבדיקות.
section_m(base_dir, queries, schemas) -> קוד M: Csv.Document + PromoteHeaders (+ TransformColumnTypes לפי schemas).
"""
from __future__ import annotations
import base64, io, re, struct, uuid, zipfile
from pathlib import Path

GOV_QUERIES = {
    "GovBonds": "gov_bonds_latest.csv", "GovCurves": "gov_curves_latest.csv", "GovHist": "gov_curves_history.csv",
    "GovBondHist": "gov_bond_history.csv", "GovCpi": "gov_cpi_path.csv", "GovVol": "gov_vol.csv", "GovForeign": "gov_foreign_latest.csv",
    "GovUS": "gov_us_latest.csv", "GovBoiCurves": "gov_boi_curves.csv", "GovBoiExp": "gov_boi_expectations.csv",
    "GovPolicy": "gov_boi_policy.csv", "GovDecisions": "gov_decisions_history.csv",
}
STEPS = ("Src", "H", "Typed")
MASHUP_NS = "http://schemas.microsoft.com/DataMashup"
EMPTY_ZIP = b"PK\x05\x06" + b"\x00" * 18


def section_m(base_dir: str, queries: dict[str, str], schemas: dict[str, list[tuple[str, str]]]) -> str:
    base = base_dir.rstrip("\\")
    parts = ["section Section1;"]
    for q, f in queries.items():
        types = ", ".join(f'{{"{c}", type {t}}}' for c, t in schemas.get(q, []))
        typed = f"Table.TransformColumnTypes(H, {{{types}}})" if types else "H"
        parts.append(f'''
shared {q} = let
    Src = Csv.Document(File.Contents("{base}\\{f}"), [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    H = Table.PromoteHeaders(Src, [PromoteAllScalars = true]),
    Typed = {typed}
in
    Typed;''')
    return "\n".join(parts) + "\n"


def _read_kit(kit_dir: Path):
    raw = (kit_dir / "DataMashup.xml").read_bytes(); txt = raw.decode("utf-16")
    b = base64.b64decode(re.search(r"<DataMashup[^>]*>(.*?)</DataMashup>", txt, re.S).group(1))
    off = 0; ver, plen = struct.unpack_from("<II", b, off); off += 8; pkg = b[off:off + plen]; off += plen
    perm_len = struct.unpack_from("<I", b, off)[0]; off += 4; perm = b[off:off + perm_len]; off += perm_len
    return ver, pkg, perm


def _package(pkg_in: bytes, section: str) -> bytes:
    zin = zipfile.ZipFile(io.BytesIO(pkg_in)); out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for n in zin.namelist():
            z.writestr(n, section.encode("utf-8") if n == "Formulas/Section1.m" else zin.read(n))
    return out.getvalue()


def _entry(t, v): return f'<Entry Type="{t}" Value="{v}" />'


def _metadata(targets: dict[str, str]) -> bytes:
    items = ['<Item><ItemLocation><ItemType>AllFormulas</ItemType><ItemPath /></ItemLocation><StableEntries>'
             '<Entry Type="Relationships" Value="sAAAAAA==" /></StableEntries></Item>']
    for q, sheet in targets.items():
        ent = [_entry("IsPrivate", "l0"), _entry("QueryID", "s" + str(uuid.uuid4())), _entry("FillEnabled", "l1"), _entry("FillObjectType", "sTable"),
               _entry("FillToDataModelEnabled", "l0"), _entry("BufferNextRefresh", "l1"), _entry("ResultType", "sTable"), _entry("NameUpdatedAfterFill", "l0"),
               _entry("NavigationStepName", "sניווט"), _entry("FillTarget", "s" + q), _entry("FilledCompleteResultToWorksheet", "l1"), _entry("FillStatus", "sComplete"),
               _entry("RecoveryTargetSheet", "s" + sheet), _entry("RecoveryTargetColumn", "l1"), _entry("RecoveryTargetRow", "l1"), _entry("AddedToDataModel", "l0")]
        items.append(f'<Item><ItemLocation><ItemType>Formula</ItemType><ItemPath>Section1/{q}</ItemPath></ItemLocation><StableEntries>{"".join(ent)}</StableEntries></Item>')
        for s in STEPS:
            items.append(f'<Item><ItemLocation><ItemType>Formula</ItemType><ItemPath>Section1/{q}/{s}</ItemPath></ItemLocation><StableEntries /></Item>')
    xml = ('\ufeff<?xml version="1.0" encoding="utf-8"?><LocalPackageMetadataFile xmlns:xsd="http://www.w3.org/2001/XMLSchema" '
           'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><Items>' + "".join(items) + "</Items></LocalPackageMetadataFile>").encode("utf-8")
    return struct.pack("<II", 0, len(xml)) + xml + struct.pack("<I", len(EMPTY_ZIP)) + EMPTY_ZIP


def build_mashup(section: str, targets: dict[str, str], kit_dir: Path) -> tuple[bytes, bytes, bytes]:
    """targets: שם שאילתה → שם גיליון היעד (לפריט RecoveryTargetSheet)."""
    kit_dir = Path(kit_dir); ver, pkg_in, perm = _read_kit(kit_dir)
    pkg = _package(pkg_in, section); meta = _metadata(targets)
    body = struct.pack("<II", ver, len(pkg)) + pkg + struct.pack("<I", len(perm)) + perm + struct.pack("<I", len(meta)) + meta + struct.pack("<I", 0)
    xml = (f'<?xml version="1.0" encoding="utf-16"?><DataMashup sqmid="{uuid.uuid4()}" xmlns="{MASHUP_NS}">'
           f'{base64.b64encode(body).decode("ascii")}</DataMashup>')
    mashup = b"\xff\xfe" + xml.encode("utf-16-le")
    props = (kit_dir / "itemProps1.xml").read_text(encoding="utf-8")
    props = re.sub(r'ds:itemID="\{[0-9A-Fa-f-]+\}"', f'ds:itemID="{{{str(uuid.uuid4()).upper()}}}"', props).encode("utf-8")
    conns = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
             'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr16" xmlns:xr16="http://schemas.microsoft.com/office/spreadsheetml/2017/revision16">']
    for i, q in enumerate(targets, 1):
        conns.append(f'<connection id="{i}" xr16:uid="{{{str(uuid.uuid4()).upper()}}}" keepAlive="1" name="שאילתה - {q}" '
                     f'description="חיבור לשאילתה \'{q}\' בחוברת העבודה." type="5" refreshedVersion="8" refreshOnLoad="1" saveData="1">'
                     f'<dbPr connection="Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={q};Extended Properties=&quot;&quot;" '
                     f'command="SELECT * FROM [{q}]"/></connection>')
    conns.append("</connections>")
    return mashup, props, "".join(conns).encode("utf-8")


def decode_mashup(mashup: bytes) -> dict:
    txt = mashup.decode("utf-16")
    b = base64.b64decode(re.search(r"<DataMashup[^>]*>(.*?)</DataMashup>", txt, re.S).group(1))
    ver, plen = struct.unpack_from("<II", b, 0); pkg = b[8:8 + plen]
    section = zipfile.ZipFile(io.BytesIO(pkg)).read("Formulas/Section1.m").decode("utf-8")
    return {"section": section, "queries": re.findall(r"^shared (\w+) = let", section, re.M), "connections": txt.count("<connection ")}


def connection_ids(targets: dict[str, str]) -> dict[str, tuple[str, int]]:
    """מפת גיליון → (שם שאילתה, מזהה חיבור) ל-build_live.inject_pq."""
    return {sheet: (q, i) for i, (q, sheet) in enumerate(targets.items(), 1)}
