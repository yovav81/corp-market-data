#!/usr/bin/env python3
"""build_gov2.py — קובץ הממשלתי v2 (ממשלתי.xlsx): כלי תומך־החלטה לתיק ממשלתי טהור. אפיון: claude/CORP-gov-design-spec.md.

שימוש: python gov/build_gov2.py <data> <פלט.xlsx> [--settings קובץ_קודם.xlsx] [--no-pq] [--base "C:\\CORP\\03_Market\\data"]
  --settings  מעביר קלטים ידניים מקובץ קודם (v1 או v2): פרמטרים, תחזית לפי טווח, החזקות, "כלול בעקום?", סימולציות, הערות.
  --no-pq     בלי הזרקת Power Query (לבדיקות ב-LibreOffice).

עקרונות: נוסחאות קלאסיות בלבד; כל פרמטר ב-01_הגדרות עם שם מוגדר p_<key>; גיליוני נתונים (נתוני_*) = טבלאות Power Query על
C:\\CORP\\03_Market\\data\\gov_*.csv עם ערכים סטטיים מהבנייה; קלטים ידניים בטבלאות נפרדות (09) עם מפתח key; כל התשואות באחוזים.
"""
from __future__ import annotations
import datetime as dt, json, os, re, sys
from pathlib import Path
import numpy as np, pandas as pd, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.chart import LineChart, Reference

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE)); sys.path.insert(0, str(HERE.parent / "corp"))
from build_gov import put, hdr, rtl, arr, lo_formula, hi_formula, NAVY, HDR, SUB, INP, YEL, GRN, RED, WH, BASE, BOLD, BLUE, NOTE, TITLE, BOX, thin  # noqa: E402
import pq_gov, gov_decide  # noqa: E402

VERSION = (HERE / "VERSION").read_text().strip() if (HERE / "VERSION").exists() else "2.0.0"
TENORS = [0.25, 0.5, 1, 2, 3, 4, 5, 7, 10, 15, 20, 30]
NB = 120; R0 = 6; R1 = R0 + NB - 1                       # שורות הניירות ב-02
GREEN = Font(name="Arial", size=10, color="008000"); GREY = Font(name="Arial", size=9, color="808080"); GFILL = PatternFill("solid", fgColor="F2F2F2")
DATA_FILL = PatternFill("solid", fgColor="EDEDED")
SOURCES = ["תחזית מנהל השקעות", "ממוצע החזאים (ב״י)", "ציפיות שוק ההון (ב״י)", "חוזי אינפלציה (ב״י)", "שקול 12 ח׳ (תוואי מדדים)"]
GROUP_ORDER = ["כללי", "שקלי/צמוד", "מח״מ", "סינתטי/בולט", "סדרות", "חו״ל", "תיק ותרחישים", "אופטימיזציה", "סימולציה", "תרחישי אופטימיזציה"]
# גיליוני נתונים: שם גיליון → (שאילתה, קובץ data/gov, מקס' שורות לטווחי הנוסחאות)
DATA_SHEETS = [("נתוני_אגח", "GovBonds", "bonds_latest.csv", 500), ("נתוני_עקומים", "GovCurves", "curves_latest.csv", 20),
               ("נתוני_היסטוריה", "GovHist", "curves_history.csv", 3000), ("נתוני_סדרות_היסט", "GovBondHist", "bond_history.csv", 20000),
               ("נתוני_מדדים", "GovCpi", "cpi_path.csv", 60), ("נתוני_תנודתיות", "GovVol", "vol.csv", 20), ("נתוני_חול", "GovForeign", "foreign_latest.csv", 60),
               ("נתוני_ארהב", "GovUS", "us_latest.csv", 20), ("נתוני_בי_עקומים", "GovBoiCurves", "boi_curves.csv", 60),
               ("נתוני_בי_ציפיות", "GovBoiExp", "boi_expectations.csv", 60), ("נתוני_בי_ריבית", "GovPolicy", "boi_policy.csv", 20),
               ("נתוני_החלטות", "GovDecisions", "decisions_history.csv", 2000), ("נתוני_אותות", "GovSignals", "signals_latest.csv", 60),
               ("נתוני_כלכלנים", "GovEcon", "economists.csv", 40)]
ORDER = ["00_דשבורד", "13_אופטימיזציה", "01_הגדרות", "02_קלט_אגח", "03_עקומים", "04_שקלי_צמוד", "05_מחמ", "06_סינטטי_בולט", "07_סדרות", "08_חול", "09_תיק_ותרחישים",
         "10_היסטוריה", "11_בנק_ישראל", "12_מדריך"] + [d[0] for d in DATA_SHEETS]
BOND_COLS = ["date", "security_id", "key", "name", "kind", "source", "maturity", "coupon", "last_price", "adjusted_value", "ytm_gross", "ytm_net", "duration_globes",
             "volume_k_nis", "ttm", "accrued", "clean_price", "macaulay_calc", "modified", "convexity", "L", "pv_check", "in_curve"]
BOND_HEB = ["תאריך נתונים", "מס׳ ני״ע", "מפתח", "שם אג״ח", "אפיק", "טבלת גלובס", "תאריך פדיון", "קופון שנתי %", "מחיר", "ערך מתואם", "תשואה ברוטו %", "תשואה נטו %",
            "מח״מ גלובס", "מחזור (אלפי ₪)", "שנים לפדיון", "ריבית צבורה", "מחיר נקי", "מח״מ Macaulay", "מח״מ מתואם", "קמירות", "מקדם הצמדה L", "מחיר תיאורטי", "כלול בעקום? (מאגר)"]
CALC_COLS = [("incl", "כלול (אפקטיבי)"), ("bucket", "Bucket"), ("lo", "lo"), ("frac", "frac"), ("ycurve", "YTM עקום %"), ("spread", "פער מול העקום (נ״ב)"),
             ("bei", "BEI בטווח %"), ("roll", "Roll-down (נ״ב)"), ("exp12", "תשואה צפויה 12 ח׳ %"), ("nomeq", "שקול נומינלי 12 ח׳ %"), ("hold", "כמות נקובה מוחזקת"),
             ("mv", "שווי שוק"), ("w", "משקל בתיק"), ("wdur", "תרומת מח״מ"), ("wytm", "תרומת תשואה"), ("wconv", "תרומת קמירות"), ("xn", "x שקלי"), ("xr", "x צמוד"), ("dn", "D שקלי"), ("dr", "D צמוד"), ("wdt", "w·D·ttm"), ("wdreal", "w·D צמוד"), ("wreal", "w צמוד"), ("wnomeq", "w·שקול"), ("status", "סטטוס")]
BC = {h: L(i + 1) for i, h in enumerate(BOND_COLS)}
CC = {k: L(len(BOND_COLS) + i + 1) for i, (k, _) in enumerate(CALC_COLS)}
Q02 = lambda c: f"'02_קלט_אגח'!${c}${R0}:${c}${R1}"          # טווח עמודה ב-02
Q03 = lambda c: f"'03_עקומים'!${c}$6:${c}$17"


class Ctx:
    def __init__(self, data: Path, prev: dict, base_dir: str):
        self.data = Path(data); self.g = self.data / "gov"; self.prev = prev; self.base_dir = base_dir
        self.settings = json.load(open(HERE / "settings_default.json", encoding="utf-8"))
        self.frames = {}
        for sheet, q, f, _ in DATA_SHEETS:
            p = self.g / f
            self.frames[sheet] = pd.read_csv(p, encoding="utf-8-sig", dtype={"key": str}) if p.exists() else pd.DataFrame()   # key תמיד טקסט
        b = self.frames["נתוני_אגח"]
        if len(b.columns):
            if "key" not in b.columns:   # מאגר ישן בלי עמודת key — נגזר כמו ב-gov_data
                sid = pd.to_numeric(b["security_id"], errors="coerce")
                b.insert(2, "key", [str(int(x)) if pd.notna(x) else str(n).strip() for x, n in zip(sid, b["name"])])
            b = b.reindex(columns=BOND_COLS + [c for c in b.columns if c not in BOND_COLS])   # סדר העמודות של 02 קבוע
        self.frames["נתוני_אגח"] = b; self.bonds = b
        self.date = pd.to_datetime(self.bonds["date"].iloc[0]).to_pydatetime() if len(self.bonds) else dt.datetime.today()
        self.wb = openpyxl.Workbook(); self.wb.remove(self.wb.active)
        self.names = {}

    def name(self, nm: str, ref: str):
        self.wb.defined_names[nm] = DefinedName(nm, attr_text=ref); self.names[nm] = ref

    def sheet(self, title: str):
        ws = self.wb.create_sheet(title); rtl(ws); return ws

    def p(self, key: str) -> str:
        return f"p_{key}"


# ---------------------------------------------------------------- נתוני_* (Power Query)
DATE_COLS = {"date", "maturity", "pub_date", "as_of"}
TEXT_COLS = {"key", "name", "kind", "source", "market", "method", "item", "value", "series", "period", "ym", "src", "in_curve", "note", "house"}


def pq_schema(ctx: "Ctx") -> dict[str, list[tuple[str, str]]]:
    """טיפוסי העמודות לשאילתות: תאריכים, מספרים, ו-key/טקסטים כ-text (מפתח חייב להישאר טקסט גם כשכל הערכים מספריים)."""
    out = {}
    for sheet, q, f, _ in DATA_SHEETS:
        df = ctx.frames[sheet]
        if not len(df.columns): continue
        cols = []
        for c in df.columns:
            if c in DATE_COLS: cols.append((c, "date"))
            elif c in TEXT_COLS: cols.append((c, "text"))
            elif pd.api.types.is_numeric_dtype(df[c]): cols.append((c, "number"))
        out[q] = cols
    return out


def sheet_data(ctx: Ctx, title: str, qname: str, df: pd.DataFrame, nmax: int):
    ws = ctx.sheet(title); ws.sheet_properties.tabColor = "808080"
    df = df.copy()
    for c in df.columns:
        if c in DATE_COLS:
            d = pd.to_datetime(df[c], errors="coerce", format="%Y-%m-%d")
            df[c] = [None if pd.isna(x) else x.to_pydatetime() for x in d]
    cols = list(df.columns) if len(df.columns) else ["(אין נתונים)"]
    for j, c in enumerate(cols, 1):
        cell = ws.cell(1, j, c); cell.font = BOLD; cell.fill = DATA_FILL
    for i, row in enumerate(df.itertuples(index=False), 2):
        for j, v in enumerate(row, 1):
            if isinstance(v, float) and np.isnan(v): v = None
            elif isinstance(v, (np.integer,)): v = int(v)
            elif isinstance(v, (np.floating,)): v = float(v)
            cell = ws.cell(i, j, v)
            if isinstance(v, dt.datetime): cell.number_format = "dd/mm/yyyy"
    nrows = max(2, len(df) + 1)
    ws.add_table(Table(displayName=qname, ref=f"A1:{L(len(cols))}{nrows}", tableStyleInfo=TableStyleInfo(name="TableStyleLight1", showRowStripes=False)))
    ws["A1"].comment = None
    for j in range(1, len(cols) + 1): ws.column_dimensions[L(j)].width = 12
    return ws


def dcol(ctx: Ctx, sheet: str, col: str, nmax: int | None = None) -> str:
    """טווח עמודה בגיליון נתונים לפי שם העמודה ב-CSV (שורות 2..nmax)."""
    df = ctx.frames[sheet]; cols = list(df.columns)
    if col not in cols: raise KeyError(f"{sheet}: אין עמודה {col} (יש {cols})")
    n = nmax or next(d[3] for d in DATA_SHEETS if d[0] == sheet)
    c = L(cols.index(col) + 1)
    return f"'{sheet}'!${c}$2:${c}${n + 1}"


# ---------------------------------------------------------------- 01_הגדרות
def sheet_settings(ctx: Ctx):
    ws = ctx.sheet("01_הגדרות")
    put(ws, "A1", "הגדרות — כל ספי ההחלטה והקלטים הידניים של קובץ הממשלתי", TITLE)
    put(ws, "A2", '="תאריך הנתונים: "&TEXT(d_asof,"dd/mm/yyyy")&"   |   גרסת בנייה ' + VERSION + '"', NOTE)
    put(ws, "C2", "תאריך הנתונים (מהמאגר):", GREY); put(ws, "D2", "=MAX('נתוני_אגח'!$A$2:$A$500)", GREY, None, "dd/mm/yyyy"); ctx.name("d_asof", "'01_הגדרות'!$D$2")
    hdr(ws, 4, ["פרמטר", "ערך", "הסבר", "מפתח", "ברירת מחדל (גרסה)"]); r = 5
    prev = ctx.prev.get("params", {})
    by_group = {}
    for x in ctx.settings: by_group.setdefault(x["group"], []).append(x)
    for grp in GROUP_ORDER:
        if grp not in by_group: continue
        c = ws.cell(r, 1, grp); c.font = BOLD; c.fill = SUB
        for j in range(2, 6): ws.cell(r, j).fill = SUB
        r += 1
        for x in by_group[grp]:
            v = prev.get(x["key"], x["value"])
            if v is None or (isinstance(v, float) and np.isnan(v)): v = ""
            put(ws, f"A{r}", x["label"], BOLD); put(ws, f"B{r}", v, BLUE, INP, x.get("fmt") or None, BOX); put(ws, f"C{r}", x["note"], NOTE); put(ws, f"D{r}", x["key"], GREY)
            put(ws, f"E{r}", "" if x["value"] in (None, "") else x["value"], GREY, None, x.get("fmt") or None)
            ctx.name(ctx.p(x["key"]), f"'01_הגדרות'!$B${r}")
            r += 1
    dv = DataValidation(type="list", formula1='"' + ",".join(SOURCES) + '"', allow_blank=True); ws.add_data_validation(dv)
    src_row = int(re.search(r"\$(\d+)$", ctx.names["p_infl_source"]).group(1)); dv.add(f"B{src_row}")
    r += 1
    put(ws, f"A{r}", "תחזית אינפלציה לפי טווח (מנהל השקעות) ומשקלי ההחלטה", BOLD); r += 1
    hdr(ws, r, ["טווח בשנים", "תחזית אינפלציה שנתית %", "משקל החלטה (0/1)", "BEI מהעקום % (לעיון)"]); r += 1
    t0 = r; pu, pw = ctx.prev.get("infl_user", {}), ctx.prev.get("dw", {})
    for i, t in enumerate(TENORS):
        e = pu.get(t, 2.0); w = pw.get(t, 1 if t in (1, 2, 3, 5) else 0)
        put(ws, f"A{r}", t, BASE, None, "0.00", BOX); put(ws, f"B{r}", e, BLUE, INP, "0.00", BOX); put(ws, f"C{r}", w, BLUE, INP, "0", BOX)
        put(ws, f"D{r}", f"='03_עקומים'!D{6 + i}", GREEN, None, "0.00", BOX); r += 1
    ctx.name("tenors01", f"'01_הגדרות'!$A${t0}:$A${r - 1}"); ctx.name("infl_user", f"'01_הגדרות'!$B${t0}:$B${r - 1}"); ctx.name("dw", f"'01_הגדרות'!$C${t0}:$C${r - 1}")
    ctx.name("p_infl_user_1y", f"'01_הגדרות'!$B${t0 + 2}")
    r += 1
    put(ws, f"A{r}", "כללי צבע", BOLD); r += 1
    for a, b in [("כחול על רקע צהוב", "קלט ידני"), ("שחור", "נוסחה"), ("ירוק", "קישור מגיליון אחר"), ("אפור", "נתון מהמאגר (Power Query)")]:
        put(ws, f"A{r}", a); put(ws, f"B{r}", b); r += 1
    r += 1; put(ws, f"A{r}", "מקורות", BOLD); r += 1
    for s in ["גלובס (טבלאות הממשלתי כולל מק״מ ומשתנה) דרך corp-market-data → data/gov; היסטוריה יומית מ-13.10.2026 + רבעונית (מגמה) 1/2023–10/2026.",
              "בנק ישראל (SDMX ZCM): עקומי אפס, ציפיות; תחזית ריבית חטיבת המחקר — ידני עם מקור (gov/boi_policy_manual.csv).",
              "FRED (ארה״ב), ECB Data Portal (עקום AAA גוש האירו), Bank of England (עקום ספוט) — יומי; כשל משיכה → הערך האחרון מסומן 'ישן'.",
              "תוואי המדדים: הלמ״ס (ידוע) + עונתיות (short/cpi_expected.py). הקובץ הזה הוא מקור הממשלתי לכל מערכת CORP."]:
        put(ws, f"A{r}", s, NOTE); r += 1
    for col, w in zip("ABCDE", (40, 22, 80, 22, 14)): ws.column_dimensions[col].width = w
    ws.column_dimensions.group("E", "E", outline_level=1, hidden=True)
    ws.freeze_panes = "A5"


# ---------------------------------------------------------------- 02_קלט_אגח
def sheet_input(ctx: Ctx):
    ws = ctx.sheet("02_קלט_אגח"); b = ctx.bonds
    put(ws, "A1", "קלט אג״ח ממשלתיות — מהמאגר (Power Query: GovBonds) + נוסחאות לכל נייר", TITLE)
    put(ws, "A3", "עמודות A–W: נתונים (מגיליון נתוני_אגח, מתרעננים בפתיחה). מימין: נוסחאות. 'כלול בעקום?' והחזקות — בטבלאות הידניות בגיליון 09 (לפי מפתח).", NOTE)
    hdr(ws, 5, BOND_HEB + [h for _, h in CALC_COLS], 1, 44)
    nd = len(BOND_COLS)
    for k in range(NB):
        r = R0 + k; dr = k + 2
        for h in BOND_COLS:
            c = BC[h]; src = f"'נתוני_אגח'!{c}{dr}"
            f = f'=IF({src}="","",{src})'
            cell = ws[f"{c}{r}"]; cell.value = f; cell.font = BASE; cell.fill = GFILL; cell.border = BOX
            if h in ("date", "maturity"): cell.number_format = "dd/mm/yyyy"
            elif h in ("coupon", "last_price", "adjusted_value", "ytm_gross", "ytm_net", "duration_globes", "ttm", "accrued", "clean_price", "macaulay_calc", "modified", "convexity", "pv_check"): cell.number_format = "0.00"
            elif h == "volume_k_nis": cell.number_format = "#,##0"
            elif h == "security_id": cell.number_format = "0"
        key, kind, ttm, ytm, mod, conv, vol, price, inc = (f"{BC[h]}{r}" for h in ("key", "kind", "ttm", "ytm_gross", "modified", "convexity", "volume_k_nis", "last_price", "in_curve"))
        C = {k2: f"{CC[k2]}{r}" for k2, _ in CALC_COLS}
        put(ws, C["incl"], f'=IF({key}="","",IFERROR(INDEX(ex_val,MATCH({key},ex_key,0)),IF({inc}="לא","לא","כן")))', BASE, None, None, BOX)
        put(ws, C["bucket"], f'=IF({ttm}="","",IF({ttm}<=1,"0-1",IF({ttm}<=2,"1-2",IF({ttm}<=5,"2-5",IF({ttm}<=10,"5-10","10+")))))', BASE, None, None, BOX)
        # אינטרפולציה על רשת הטווחים של 03 (MATCH מקורב על רשימה עולה) — lo = אינדקס הטווח התחתון, frac בתוך [0,1]
        put(ws, C["lo"], f'=IF({ttm}="","",IF({ttm}<=INDEX(tenors,1),1,MIN(11,MATCH({ttm},tenors,1))))', GREY, None, "0")
        lo = C["lo"]
        put(ws, C["frac"], f'=IF({lo}="","",MAX(0,MIN(1,({ttm}-INDEX(tenors,{lo}))/(INDEX(tenors,{lo}+1)-INDEX(tenors,{lo})))))', GREY, None, "0.000")
        fr = C["frac"]
        def interp(rng):  # ערך העקום בטווח ttm
            return f"(INDEX({rng},{lo})+(INDEX({rng},{lo}+1)-INDEX({rng},{lo}))*{fr})"
        isnom = f'OR({kind}="שקלי",{kind}="מק""מ")'
        put(ws, C["ycurve"], f'=IF(OR({lo}="",{ytm}=""),"",IF({isnom},{interp("cur_nom")},IF({kind}="מדד",{interp("cur_real")},"")))', BASE, None, "0.00", BOX)
        put(ws, C["spread"], f'=IF(OR({C["ycurve"]}="",{ytm}=""),"",({ytm}-{C["ycurve"]})*100)', BASE, None, "0.0", BOX)
        put(ws, C["bei"], f'=IF({lo}="","",{interp("cur_bei")})', BASE, None, "0.00", BOX)
        put(ws, C["roll"], f'=IF(OR({lo}="",{C["ycurve"]}=""),"",IF({isnom},{interp("cur_roll_n")},{interp("cur_roll_r")}))', BASE, None, "0.0", BOX)
        put(ws, C["exp12"], f'=IF(OR({C["roll"]}="",{mod}=""),"",{ytm}+{mod}*{C["roll"]}/100)', BASE, None, "0.00", BOX)
        put(ws, C["nomeq"], f'=IF({C["exp12"]}="","",IF({kind}="מדד",((1+{ytm}/100)*(1+infl_12m/100)-1)*100+{mod}*{C["roll"]}/100,{C["exp12"]}))', BASE, None, "0.00", BOX)
        put(ws, C["hold"], f'=IF({key}="","",IFERROR(INDEX(h_amt,MATCH({key},h_key,0)),0))', BASE, None, "#,##0", BOX)
        put(ws, C["mv"], f'=IFERROR({C["hold"]}*{price}/100,0)', BASE, None, "#,##0", BOX)
        put(ws, C["w"], f'=IFERROR({C["mv"]}/SUM({Q02(CC["mv"])}),0)', BASE, None, "0.0%", BOX)
        put(ws, C["wdur"], f'=IFERROR({C["w"]}*{mod},0)', BASE, None, "0.000", BOX)
        put(ws, C["wytm"], f'=IFERROR({C["w"]}*{ytm},0)', BASE, None, "0.000", BOX)
        put(ws, C["wconv"], f'=IFERROR({C["w"]}*{conv},0)', BASE, None, "0.00", BOX)
        put(ws, C["xn"], f'=IF(AND({C["incl"]}="כן",{isnom},ISNUMBER({ttm}),ISNUMBER({ytm})),{ttm},"")', GREY, None, "0.00")
        put(ws, C["xr"], f'=IF(AND({C["incl"]}="כן",{kind}="מדד",ISNUMBER({ttm}),ISNUMBER({ytm})),{ttm},"")', GREY, None, "0.00")
        put(ws, C["dn"], f'=IF(AND({C["incl"]}="כן",{isnom},ISNUMBER({mod}),ISNUMBER({ytm}),ISNUMBER({ttm})),{mod},"")', GREY, None, "0.000")
        put(ws, C["dr"], f'=IF(AND({C["incl"]}="כן",{kind}="מדד",ISNUMBER({mod}),ISNUMBER({ytm}),ISNUMBER({ttm})),{mod},"")', GREY, None, "0.000")
        put(ws, C["wdt"], f'=IFERROR({C["w"]}*{mod}*{ttm},0)', GREY, None, "0.000"); put(ws, C["wdreal"], f'=IF({kind}="מדד",{C["wdur"]},0)', GREY, None, "0.000")
        put(ws, C["wreal"], f'=IF({kind}="מדד",{C["w"]},0)', GREY, None, "0.000"); put(ws, C["wnomeq"], f'=IFERROR({C["w"]}*{C["nomeq"]},0)', GREY, None, "0.000")
        put(ws, C["status"], f'=IF({key}="","",IF({ytm}="","אין תשואה",IF({mod}="","אין מח״מ","")))', BASE, None, None, BOX)
    widths = [10, 10, 10, 20, 7, 12, 11, 8, 9, 9, 9, 9, 8, 10, 8, 8, 8, 8, 8, 8, 8, 8, 8] + [8, 7, 4, 5, 9, 9, 8, 9, 10, 10, 12, 11, 8, 8, 8, 8, 6, 6, 6, 6, 6, 6, 6, 6, 10]
    for i, w in enumerate(widths): ws.column_dimensions[L(i + 1)].width = w
    ws.column_dimensions.group(CC["lo"], CC["frac"], outline_level=1, hidden=True); ws.column_dimensions.group(CC["xn"], CC["wnomeq"], outline_level=1, hidden=True)
    ws.freeze_panes = f"E{R0}"; ws.auto_filter.ref = f"A5:{CC['status']}{R1}"
    for nm, col in [("b_key", BC["key"]), ("b_name", BC["name"]), ("b_kind", BC["kind"]), ("b_ttm", BC["ttm"]), ("b_ytm", BC["ytm_gross"]), ("b_mod", BC["modified"]),
                    ("b_conv", BC["convexity"]), ("b_vol", BC["volume_k_nis"]), ("b_price", BC["last_price"]), ("b_sid", BC["security_id"]), ("b_incl", CC["incl"]),
                    ("b_ycurve", CC["ycurve"]), ("b_spread", CC["spread"]), ("b_roll", CC["roll"]), ("b_exp12", CC["exp12"]), ("b_nomeq", CC["nomeq"]), ("b_hold", CC["hold"]),
                    ("b_mv", CC["mv"]), ("b_w", CC["w"]), ("b_xn", CC["xn"]), ("b_xr", CC["xr"]), ("b_dn", CC["dn"]), ("b_dr", CC["dr"]), ("b_bucket", CC["bucket"]),
                    ("b_wdur", CC["wdur"]), ("b_wconv", CC["wconv"]), ("b_wytm", CC["wytm"]), ("b_wdt", CC["wdt"]), ("b_wdreal", CC["wdreal"]), ("b_wreal", CC["wreal"]), ("b_wnomeq", CC["wnomeq"])]:
        ctx.name(nm, Q02(col))


# ---------------------------------------------------------------- 03_עקומים
def sheet_curves(ctx: Ctx):
    ws = ctx.sheet("03_עקומים")
    put(ws, "A1", "עקומי אג״ח: שקלי (מק״מ + שקלית), צמוד ריאלי, BEI, פורוורדים, roll-down, תנודתיות ואחוזונים", TITLE)
    put(ws, "A3", "העקום מחושב כאן מהניירות המסומנים 'כן' ב-02 (אינטרפולציה ליניארית על שנים לפדיון; מחוץ לטווח — ערך הקצה). עמודות P–S: העקום מהפייתון (נתוני_עקומים) לבדיקה. אחוזונים מול ההיסטוריה (10).", NOTE)
    ch = ["טווח בשנים", "שקלי %", "ריאלי %", "BEI %", "BEI מצטבר %", "פורוורד שקלי %", "פורוורד ריאלי %", "פורוורד BEI %", "שיפוע שקלי מול 2Y (נ״ב)", "שיפוע ריאלי מול 2Y (נ״ב)",
          "Roll-down שקלי (נ״ב)", "Roll-down ריאלי (נ״ב)", "תנודתיות שקלי (נ״ב/שנה)", "תנודתיות ריאלי", "תנודתיות BEI", "פייתון שקלי %", "פייתון ריאלי %", "הפרש שקלי (נ״ב)", "הפרש ריאלי (נ״ב)",
          "אחוזון BEI", "אחוזון ריאלי", "אחוזון שקלי", "lo שקלי", "hi שקלי", "lo צמוד", "hi צמוד", "x roll", "lo roll", "hi roll"]
    hdr(ws, 5, ch, 1, 36)
    XN, XR, YY = "b_xn", "b_xr", "b_ytm"
    def interp(x, xs, ys, lo, hi):
        ylo = f"INDEX({ys},MATCH({lo},{xs},0))"; yhi = f"INDEX({ys},MATCH({hi},{xs},0))"
        return f'IFERROR(IF({lo}=-1,{yhi},IF({hi}=9999,{ylo},IF({lo}={hi},{ylo},{ylo}+({x}-{lo})*({yhi}-{ylo})/({hi}-{lo})))),"")'
    hist_col = {1: "1", 2: "2", 5: "5", 10: "10"}
    for i, t in enumerate(TENORS):
        r = 6 + i
        put(ws, f"A{r}", t, BASE, None, "0.00", BOX)
        arr(ws, f"W{r}", "=" + lo_formula(f"$A{r}", XN)); arr(ws, f"X{r}", "=" + hi_formula(f"$A{r}", XN))
        arr(ws, f"Y{r}", "=" + lo_formula(f"$A{r}", XR)); arr(ws, f"Z{r}", "=" + hi_formula(f"$A{r}", XR))
        put(ws, f"B{r}", "=" + interp(f"$A{r}", XN, YY, f"W{r}", f"X{r}"), BASE, None, "0.00", BOX)
        put(ws, f"C{r}", "=" + interp(f"$A{r}", XR, YY, f"Y{r}", f"Z{r}"), BASE, None, "0.00", BOX)
        put(ws, f"D{r}", f'=IFERROR(((1+B{r}/100)/(1+C{r}/100)-1)*100,"")', BASE, None, "0.00", BOX)
        put(ws, f"E{r}", f'=IFERROR(((1+D{r}/100)^A{r}-1)*100,"")', BASE, None, "0.00", BOX)
        if i == 0:
            put(ws, f"F{r}", f'=IFERROR(B{r},"")', BASE, None, "0.00", BOX); put(ws, f"G{r}", f'=IFERROR(C{r},"")', BASE, None, "0.00", BOX)
        else:
            put(ws, f"F{r}", f'=IFERROR((((1+B{r}/100)^A{r}/(1+B{r-1}/100)^A{r-1})^(1/(A{r}-A{r-1}))-1)*100,"")', BASE, None, "0.00", BOX)
            put(ws, f"G{r}", f'=IFERROR((((1+C{r}/100)^A{r}/(1+C{r-1}/100)^A{r-1})^(1/(A{r}-A{r-1}))-1)*100,"")', BASE, None, "0.00", BOX)
        put(ws, f"H{r}", f'=IFERROR(((1+F{r}/100)/(1+G{r}/100)-1)*100,"")', BASE, None, "0.00", BOX)
        put(ws, f"I{r}", f'=IFERROR((B{r}-INDEX($B$6:$B$17,MATCH(2,$A$6:$A$17,0)))*100,"")', BASE, None, "0", BOX)
        put(ws, f"J{r}", f'=IFERROR((C{r}-INDEX($C$6:$C$17,MATCH(2,$A$6:$A$17,0)))*100,"")', BASE, None, "0", BOX)
        put(ws, f"AA{r}", f"=MAX($A{r}-1,MIN($A$6:$A$17))", GREY, None, "0.00")
        arr(ws, f"AB{r}", "=" + lo_formula(f"$AA{r}", "$A$6:$A$17")); arr(ws, f"AC{r}", "=" + hi_formula(f"$AA{r}", "$A$6:$A$17"))
        put(ws, f"K{r}", f'=IFERROR((B{r}-' + interp(f"$AA{r}", "$A$6:$A$17", "$B$6:$B$17", f"AB{r}", f"AC{r}") + ')*100,"")', BASE, None, "0.0", BOX)
        put(ws, f"L{r}", f'=IFERROR((C{r}-' + interp(f"$AA{r}", "$A$6:$A$17", "$C$6:$C$17", f"AB{r}", f"AC{r}") + ')*100,"")', BASE, None, "0.0", BOX)
        for col, vc in (("M", "nominal_vol_bp"), ("N", "real_vol_bp"), ("O", "bei_vol_bp")):
            put(ws, f"{col}{r}", f'=IFERROR(INDEX({dcol(ctx, "נתוני_תנודתיות", vc)},MATCH($A{r},{dcol(ctx, "נתוני_תנודתיות", "tenor")},0)),IF("{col}"="N",p_vol_floor_real_bp,p_vol_floor_nominal_bp))', BASE, None, "0", BOX)
        put(ws, f"P{r}", f'=IFERROR(INDEX({dcol(ctx, "נתוני_עקומים", "nominal")},MATCH($A{r},{dcol(ctx, "נתוני_עקומים", "tenor")},0)),"")', GREY, None, "0.00", BOX)
        put(ws, f"Q{r}", f'=IFERROR(INDEX({dcol(ctx, "נתוני_עקומים", "real")},MATCH($A{r},{dcol(ctx, "נתוני_עקומים", "tenor")},0)),"")', GREY, None, "0.00", BOX)
        put(ws, f"R{r}", f'=IFERROR((B{r}-P{r})*100,"")', GREY, None, "0.0", BOX); put(ws, f"S{r}", f'=IFERROR((C{r}-Q{r})*100,"")', GREY, None, "0.0", BOX)
        if t in hist_col and len(ctx.frames["נתוני_היסטוריה"].columns):
            hc = hist_col[t]
            def pct(col, x):   # אחוזון = חלק ההיסטוריה שמתחת לערך היום (כמו בפייתון)
                rng = dcol(ctx, "נתוני_היסטוריה", col)
                return f'=IF(OR({x}="",COUNT({rng})<5),"",COUNTIF({rng},"<"&{x})/COUNT({rng})*100)'
            put(ws, f"T{r}", pct("bei_" + hc, f"D{r}"), BASE, None, "0", BOX)
            put(ws, f"U{r}", pct("real_" + hc, f"C{r}"), BASE, None, "0", BOX)
            put(ws, f"V{r}", pct("nom_" + hc, f"B{r}"), BASE, None, "0", BOX)
        else:
            for col in "TUV": put(ws, f"{col}{r}", "", BASE, None, None, BOX)
        for col in ("W", "X", "Y", "Z", "AB", "AC"): ws[f"{col}{r}"].font = GREY
    for i, w in enumerate([8, 9, 9, 9, 9, 10, 10, 10, 10, 10, 10, 10, 10, 9, 9, 9, 9, 9, 9, 8, 8, 8, 6, 6, 6, 6, 6, 6, 6]): ws.column_dimensions[L(i + 1)].width = w
    ws.column_dimensions.group("W", "AC", outline_level=1, hidden=True)
    put(ws, "A19", "עזר (W–AC, מכווץ): נקודות העקום הסמוכות לכל טווח; AA = הטווח באופק ה-roll-down.", NOTE)
    for nm, col in [("tenors", "A"), ("cur_nom", "B"), ("cur_real", "C"), ("cur_bei", "D"), ("cur_fwd_n", "F"), ("cur_fwd_r", "G"), ("cur_roll_n", "K"), ("cur_roll_r", "L"),
                    ("cur_vol_n", "M"), ("cur_vol_r", "N"), ("cur_vol_bei", "O"), ("cur_pct_bei", "T"), ("cur_fwd_bei", "H")]:
        ctx.name(nm, Q03(col))
    lc = LineChart(); lc.title = "עקום שקלי וריאלי"; lc.y_axis.title = "%"; lc.x_axis.title = "שנים"; lc.height = 8; lc.width = 18
    lc.add_data(Reference(ws, min_col=2, min_row=5, max_row=17), titles_from_data=True); lc.add_data(Reference(ws, min_col=3, min_row=5, max_row=17), titles_from_data=True)
    lc.set_categories(Reference(ws, min_col=1, min_row=6, max_row=17)); ws.add_chart(lc, "A21")
    lb = LineChart(); lb.title = "BEI לפי טווח"; lb.height = 8; lb.width = 18
    lb.add_data(Reference(ws, min_col=4, min_row=5, max_row=17), titles_from_data=True); lb.set_categories(Reference(ws, min_col=1, min_row=6, max_row=17)); ws.add_chart(lb, "H21")


# ---------------------------------------------------------------- 04_שקלי_צמוד
def sheet_split(ctx: Ctx):
    ws = ctx.sheet("04_שקלי_צמוד")
    put(ws, "A1", "שקלי מול צמוד — BEI מול תחזית, תוואי מדדים 12 חודשים, והחלטת המודל על משקל הצמוד", TITLE)
    put(ws, "A3", "פער = תחזית − BEI (נ״ב). הפער המכריע = ממוצע משוקלל לפי משקלי ההחלטה (01). משקל צמוד = נייטרלי + הטיה פרופורציונית לפער (סולם, הטיה מקסימלית), תיקון אחוזון, חסום למינימום/מקסימום.", NOTE)
    hdr(ws, 5, ["טווח בשנים", "BEI מהעקום %", "תחזית אינפלציה %", "פער תחזית−BEI (נ״ב)", "פורוורד BEI %", "משקל החלטה", "פער × משקל"])
    exp_s = ctx.names.get("p_infl_source", "p_infl_source")
    for i, t in enumerate(TENORS):
        r = 6 + i
        put(ws, f"A{r}", f"=INDEX(tenors,{i + 1})", GREEN, None, "0.00", BOX)
        put(ws, f"B{r}", f"=INDEX(cur_bei,{i + 1})", GREEN, None, "0.00", BOX)
        put(ws, f"C{r}", f'=IF(p_infl_source="{SOURCES[0]}",INDEX(infl_user,{i + 1}),infl_fc)', BASE, None, "0.00", BOX)
        put(ws, f"D{r}", f'=IF(OR(B{r}="",C{r}=""),"",(C{r}-B{r})*100)', BASE, None, "0", BOX)
        put(ws, f"E{r}", f"=INDEX(cur_fwd_bei,{i + 1})", GREEN, None, "0.00", BOX)
        put(ws, f"F{r}", f"=INDEX(dw,{i + 1})", GREEN, None, "0", BOX)
        put(ws, f"G{r}", f'=IF(D{r}="",0,D{r}*F{r})', BASE, None, "0.0", BOX)
    put(ws, "A19", "פער מכריע (נ״ב)", BOLD); put(ws, "B19", '=IF(OR(SUM(F6:F17)=0,infl_fc=""),"",SUM(G6:G17)/SUM(F6:F17))', BOLD, None, "0", BOX); ctx.name("gap_bp", "'04_שקלי_צמוד'!$B$19")
    put(ws, "C19", "אחוזון BEI 2Y מול ההיסטוריה", BOLD); put(ws, "D19", "=INDEX(cur_pct_bei,MATCH(2,tenors,0))", GREEN, None, "0", BOX); ctx.name("bei_pct2", "'04_שקלי_צמוד'!$D$19")
    # מקור התחזית
    put(ws, "A21", "מקור תחזית האינפלציה (01)", BOLD); put(ws, "B21", "=p_infl_source", GREEN, None, None, BOX)
    hdr(ws, 22, ["מקור", "ערך %", "הערה"])
    ser = dcol(ctx, "נתוני_בי_ציפיות", "series") if len(ctx.frames["נתוני_בי_ציפיות"].columns) else None
    val = dcol(ctx, "נתוני_בי_ציפיות", "value") if ser else None
    def boi(code): return f'IFERROR(INDEX({val},MATCH("{code}",{ser},0)),"")' if ser else '""'
    rows = [(SOURCES[0], "=p_infl_user_1y", "תחזית לפי טווח — גיליון 01 (טווח שנה)"), (SOURCES[1], "=" + boi("EXP_HAZ_MAD_AVG_01Y_MA"), "בנק ישראל — ממוצע החזאים 12 ח׳"),
            (SOURCES[2], "=" + boi("EXP_EXP_01Y_BI_MA"), "בנק ישראל — ציפיות משוק ההון לשנה"), (SOURCES[3], "=" + boi("EXP_EXP_ISSWIT1_01Y_MA"), "בנק ישראל — חוזי אינפלציה לשנה"),
            (SOURCES[4], "=cpi12_data", "תוואי המדדים: 12 החודשים הצפויים (נתוני_מדדים)")]
    for i, (a, b, c) in enumerate(rows):
        r = 23 + i; put(ws, f"A{r}", a, BASE, None, None, BOX); put(ws, f"B{r}", b, GREEN, None, "0.00", BOX); put(ws, f"C{r}", c, NOTE)
    put(ws, "A28", "תחזית שנתית בשימוש %", BOLD); put(ws, "B28", '=IFERROR(INDEX($B$23:$B$27,MATCH(p_infl_source,$A$23:$A$27,0)),"")', BOLD, None, "0.00", BOX); ctx.name("infl_fc", "'04_שקלי_צמוד'!$B$28")
    # תוואי 12 חודשים
    put(ws, "A30", "תוואי המדדים ל-12 החודשים הצפויים: חודש צפוי = תחזית שנתית/12 + עונתיות (נתוני_מדדים)", BOLD)
    hdr(ws, 31, ["חודש", "עונתיות (נק׳)", "צפוי לפי התחזית %", "מהמאגר (BEI) %", "מקדם מצטבר"])
    cpi = ctx.frames["נתוני_מדדים"]
    if len(cpi.columns):
        kind, ym, season, mpct = dcol(ctx, "נתוני_מדדים", "kind"), dcol(ctx, "נתוני_מדדים", "ym"), dcol(ctx, "נתוני_מדדים", "season_pct"), dcol(ctx, "נתוני_מדדים", "monthly_pct")
        put(ws, "G30", f'=IFERROR(MATCH("צפוי",{kind},0),"")', GREY); put(ws, "F30", "אינדקס החודש הצפוי הראשון:", GREY)
        for i in range(12):
            r = 32 + i; k = f"$G$30+{i}"
            put(ws, f"A{r}", f'=IF($G$30="","",INDEX({ym},{k}))', GREEN, None, None, BOX)
            put(ws, f"B{r}", f'=IF($G$30="","",INDEX({season},{k}))', GREEN, None, "0.000", BOX)
            put(ws, f"C{r}", f'=IF(OR(B{r}="",infl_fc=""),"",infl_fc/12+B{r})', BASE, None, "0.000", BOX)
            put(ws, f"D{r}", f'=IF($G$30="","",INDEX({mpct},{k}))', GREEN, None, "0.000", BOX)
            put(ws, f"E{r}", f'=IF(C{r}="","",{"1" if i == 0 else f"E{r - 1}"}*(1+C{r}/100))', BASE, None, "0.0000", BOX)
        put(ws, "A44", "אינפלציה 12 ח׳ לפי התחזית %", BOLD); put(ws, "B44", '=IF(E43="","",(E43-1)*100)', BOLD, None, "0.00", BOX)
        put(ws, "C44", "אינפלציה 12 ח׳ מהמאגר (BEI×עונתיות) %", BOLD); put(ws, "D44", '=IF(D32="","",(PRODUCT(1+D32:D43/100)-1)*100)', BOLD, None, "0.00", BOX)
        arr(ws, "D44", '=IF(D32="","",(PRODUCT(1+D32:D43/100)-1)*100)')
    else:
        put(ws, "A44", "אינפלציה 12 ח׳ לפי התחזית %", BOLD); put(ws, "B44", "=infl_fc", BOLD, None, "0.00", BOX); put(ws, "D44", "", BASE)
    ctx.names.pop("infl_12m", None); ctx.name("infl_12m", "'04_שקלי_צמוד'!$B$44"); ctx.name("cpi12_data", "'04_שקלי_צמוד'!$D$44")
    # החלטת המודל
    put(ws, "A46", "החלטת המודל — משקל צמוד", BOLD)
    hdr(ws, 47, ["רכיב", "ערך", "הסבר"])
    items = [("הטיה גולמית", '=IF(gap_bp="",0,IF(ABS(gap_bp)<=p_decision_margin_bp,0,MAX(-1,MIN(1,gap_bp/p_tilt_scale_bp))*p_tilt_max))', "חסום(פער/סולם) × הטיה מקסימלית; בתוך מרווח ההחלטה — 0", "0.0%"),
             ("תיקון אחוזון", '=IF(OR(p_pct_adj_on<>1,bei_pct2=""),0,IF(bei_pct2>p_pct_hi,-p_pct_adj,IF(bei_pct2<p_pct_lo,p_pct_adj,0)))', "BEI באחוזון גבוה → לשקלי; נמוך → לצמוד", "0.0%"),
             ("משקל צמוד — כלל ה-BEI (עוגן)", "=MAX(p_linked_min,MIN(p_linked_max,p_linked_neutral+B48+B49))", "נייטרלי + הטיה + תיקון, בתוך המינימום/מקסימום; האופטימיזציה (13) בוחנת סביבו", "0%"),
             ("משקל צמוד — סימולציה", '=IF(p_sim_linked="","",p_sim_linked)', "ריק = לפי המודל (01 → סימולציה)", "0%"),
             ("משקל צמוד בשימוש", '=IF(p_sim_linked="",m_linked,p_sim_linked)', "סימולציה אם מולאה; אחרת המודל הסופי (אופטימיזציה, 13)", "0%")]
    for i, (a, b, c, f) in enumerate(items):
        r = 48 + i; put(ws, f"A{r}", a, BOLD if i in (2, 4) else BASE, None, None, BOX); put(ws, f"B{r}", b, BOLD if i in (2, 4) else BASE, GRN if i == 2 else None, f, BOX); put(ws, f"C{r}", c, NOTE)
    ctx.name("r_linked", "'04_שקלי_צמוד'!$B$50"); ctx.name("u_linked", "'04_שקלי_צמוד'!$B$52")
    put(ws, "A53", "למה", BOLD)
    put(ws, "B53", '=IF(gap_bp="","אין תחזית אינפלציה — נייטרלי","מקור "&p_infl_source&" "&TEXT(infl_fc,"0.00")&"%; פער תחזית−BEI "&TEXT(gap_bp,"+0;-0")&" נ״ב"&IF(ABS(gap_bp)<=p_decision_margin_bp," (בתוך מרווח ההחלטה → נייטרלי)"," → הטיה "&TEXT(B48,"+0%;-0%"))&IF(B49<>0,"; אחוזון BEI "&TEXT(bei_pct2,"0")&" → "&TEXT(B49,"+0%;-0%"),"")&" → משקל צמוד "&TEXT(B50,"0%"))', BASE)
    ctx.name("r_linked_why", "'04_שקלי_צמוד'!$B$53")
    # 12 חודשים: צמוד מול שקלי
    put(ws, "A55", "12 חודשים קדימה — צמוד (ריאלי + תוואי המדדים + roll) מול שקלי באותו טווח", BOLD)
    hdr(ws, 56, ["טווח", "שקלי: תשואה 12 ח׳ %", "צמוד: שקול נומינלי 12 ח׳ %", "יתרון צמוד (נ״ב)"])
    for i, t in enumerate((1, 2, 5)):
        r = 57 + i; k = TENORS.index(t) + 1
        d_n = f"(1-(1+INDEX(cur_nom,{k})/100)^-INDEX(tenors,{k}))/(INDEX(cur_nom,{k})/100)"; d_r = f"(1-(1+INDEX(cur_real,{k})/100)^-INDEX(tenors,{k}))/(INDEX(cur_real,{k})/100)"
        put(ws, f"A{r}", t, BASE, None, "0", BOX)
        put(ws, f"B{r}", f'=IFERROR(INDEX(cur_nom,{k})+{d_n}*INDEX(cur_roll_n,{k})/100,"")', BASE, None, "0.00", BOX)
        put(ws, f"C{r}", f'=IFERROR(((1+INDEX(cur_real,{k})/100)*(1+infl_12m/100)-1)*100+{d_r}*INDEX(cur_roll_r,{k})/100,"")', BASE, None, "0.00", BOX)
        put(ws, f"D{r}", f'=IF(OR(B{r}="",C{r}=""),"",(C{r}-B{r})*100)', BASE, None, "0", BOX)
    ctx.name("adv_linked_1y", "'04_שקלי_צמוד'!$D$57")
    for col, w in zip("ABCDEFG", (26, 16, 20, 20, 16, 12, 12)): ws.column_dimensions[col].width = w
    ws.freeze_panes = "A6"


# ---------------------------------------------------------------- 05_מחמ
def _dur_table(ctx: Ctx, ws, r0: int, kind: str):
    """טבלת מח״מ לאפיק: שורות r0..r0+11. מחזיר מילון שמות עמודות."""
    ycol, rcol, vcol = ("cur_nom", "cur_roll_n", "cur_vol_n") if kind == "nominal" else ("cur_real", "cur_roll_r", "cur_vol_r")
    hdr(ws, r0 - 1, ["טווח", "תשואה %", "Roll-down (נ״ב)", "מח״מ (פארי)", "תשואה 12 ח׳ %", "עודף מול שנה (נ״ב)", "Breakeven (נ״ב)", "תנודתיות (נ״ב/שנה)", "סיכון = מח״מ×תנודתיות", "ציון", "מותר?", "ציון מותר"])
    y1 = f"INDEX({ycol},MATCH(1,tenors,0))"
    for i in range(12):
        r = r0 + i; k = i + 1
        put(ws, f"A{r}", f"=INDEX(tenors,{k})", GREEN, None, "0.00", BOX); put(ws, f"B{r}", f"=INDEX({ycol},{k})", GREEN, None, "0.00", BOX)
        put(ws, f"C{r}", f"=INDEX({rcol},{k})", GREEN, None, "0.0", BOX)
        put(ws, f"D{r}", f'=IF(B{r}="","",IF(ABS(B{r})<0.000001,A{r},(1-(1+B{r}/100)^-A{r})/(B{r}/100)))', BASE, None, "0.000", BOX)
        put(ws, f"E{r}", f'=IF(D{r}="","",B{r}+D{r}*C{r}/100)', BASE, None, "0.00", BOX)
        put(ws, f"F{r}", f'=IF(E{r}="","",(E{r}-{y1})*100)', BASE, None, "0.0", BOX)
        put(ws, f"G{r}", f'=IF(OR(F{r}="",D{r}=0),"",F{r}/D{r})', BASE, None, "0.0", BOX)
        put(ws, f"H{r}", f"=INDEX({vcol},{k})", GREEN, None, "0", BOX)
        put(ws, f"I{r}", f'=IF(D{r}="","",D{r}*H{r})', BASE, None, "0", BOX)
        put(ws, f"J{r}", f'=IF(OR(I{r}="",I{r}=0),"",(F{r}-p_dur_friction_bp)/I{r})', BASE, None, "0.000", BOX)
        put(ws, f"K{r}", f'=IF(D{r}="","",IF(AND(D{r}>=p_dur_min,D{r}<=p_dur_max,F{r}>=p_dur_min_excess_bp),"כן","לא"))', BASE, None, None, BOX)
        put(ws, f"L{r}", f'=IF(K{r}="כן",J{r},-9999)', GREY, None, "0.000")
    return {"tenor": f"$A${r0}:$A${r0 + 11}", "dur": f"$D${r0}:$D${r0 + 11}", "exp": f"$E${r0}:$E${r0 + 11}", "exc": f"$F${r0}:$F${r0 + 11}",
            "be": f"$G${r0}:$G${r0 + 11}", "score": f"$J${r0}:$J${r0 + 11}", "ok": f"$K${r0}:$K${r0 + 11}", "sc_ok": f"$L${r0}:$L${r0 + 11}"}


def _dur_pick(ctx: Ctx, ws, r: int, T: dict, kind: str):
    """בלוק ההחלטה לאפיק בשורות r..r+9. מגדיר שמות m_dur_<kind>_tenor, m_dur_<kind>, u_dur_<kind>, m_dur_<kind>_why."""
    heb = "שקלי" if kind == "nominal" else "צמוד"; sim = f"p_sim_dur_{kind}"
    hcol = "dur_nominal_tenor" if kind == "nominal" else "dur_real_tenor"
    if "rule_" + hcol in ctx.frames["נתוני_החלטות"].columns: hcol = "rule_" + hcol     # מ-v2.1 ההמלצה הקודמת של הכלל נשמרת בנפרד
    hist_ok = len(ctx.frames["נתוני_החלטות"].columns) and hcol in ctx.frames["נתוני_החלטות"].columns
    prev_f = f'IFERROR(INDEX({dcol(ctx, "נתוני_החלטות", hcol)},MATCH(d_asof-1,{dcol(ctx, "נתוני_החלטות", "date")},1)),"")' if hist_ok else '""'
    hdr(ws, r, ["רכיב", "ערך", "הסבר"])
    rows = [("הטווח עם הציון המקסימלי", f'=IF(MAX({T["sc_ok"]})=-9999,"",INDEX({T["tenor"]},MATCH(MAX({T["sc_ok"]}),{T["sc_ok"]},0)))', "בין הטווחים המותרים (מח״מ מינימום/מקסימום, עודף מינימלי)", "0.00"),
            ("טווח ההמלצה הקודמת (אתמול)", "=" + prev_f, "מנתוני_החלטות (הפייתון הלילי)", "0.00"),
            ("אחרי סף יציבות", f'=IF(B{r + 1}="","",IF(B{r + 2}="",B{r + 1},IF(IFERROR(INDEX({T["ok"]},MATCH(B{r + 2},{T["tenor"]},0)),"לא")<>"כן",B{r + 1},IF(INDEX({T["score"]},MATCH(B{r + 1},{T["tenor"]},0))<INDEX({T["score"]},MATCH(B{r + 2},{T["tenor"]},0))*(1+p_dur_stability),B{r + 2},B{r + 1}))))',
             "ההמלצה מתחלפת רק אם הציון החדש גבוה בסף היציבות מציון ההמלצה הקודמת", "0.00"),
            ("אות מדיניות (נ״ב)", "=policy_gap_bp", "פורוורד שנה-בעוד-שנה מול תחזית ריבית ב״י; שלילי = השוק מגלם הפחתות מעבר לב״י", "0"),
            ("טווח מומלץ — המודל", f'=IF(B{r + 3}="","",IF(AND(p_policy_adj_on=1,B{r + 4}<>"",N(B{r + 4})<=-p_policy_gap_bp),IF(COUNTIFS({T["tenor"]},"<"&B{r + 3},{T["ok"]},"כן")>0,MAX(IF(({T["tenor"]}<B{r + 3})*({T["ok"]}="כן"),{T["tenor"]})),B{r + 3}),B{r + 3}))', "צעד טווח אחד למטה כשהאות המדיניות פעיל", "0.00"),
            (f"מח״מ מומלץ {heb} — כלל הציון", f'=IF(B{r + 5}="","",INDEX({T["dur"]},MATCH(B{r + 5},{T["tenor"]},0)))', "מח״מ (פארי) של הטווח המומלץ לפי הכלל; ההחלטה הסופית באופטימיזציה (13)", "0.00"),
            (f"מח״מ {heb} — סימולציה", f'=IF({sim}="","",{sim})', "ריק = לפי המודל", "0.00"),
            (f"מח״מ {heb} בשימוש", f'=IF({sim}="",m_dur_{kind},{sim})', "סימולציה אם מולאה; אחרת המודל הסופי (אופטימיזציה, 13)", "0.00")]
    for i, (a, b, c, f) in enumerate(rows):
        rr = r + 1 + i
        if i == 4: arr(ws, f"B{rr}", b)
        else: put(ws, f"B{rr}", b, BOLD if i in (5, 7) else BASE, GRN if i == 5 else None, f, BOX)
        ws[f"B{rr}"].number_format = f; ws[f"B{rr}"].border = BOX
        put(ws, f"A{rr}", a, BOLD if i in (5, 7) else BASE, None, None, BOX); put(ws, f"C{rr}", c, NOTE)
    ctx.name(f"r_dur_{kind}_tenor", f"'05_מחמ'!$B${r + 5}"); ctx.name(f"r_dur_{kind}", f"'05_מחמ'!$B${r + 6}"); ctx.name(f"u_dur_{kind}", f"'05_מחמ'!$B${r + 8}")
    put(ws, f"A{r + 9}", "למה", BOLD)
    put(ws, f"B{r + 9}", f'=IF(B{r + 1}="","אין טווח שעומד במגבלות","ציון מקסימלי "&TEXT(INDEX({T["score"]},MATCH(B{r + 1},{T["tenor"]},0)),"0.00")&" בטווח "&TEXT(B{r + 1},"0.##")&IF(B{r + 3}<>B{r + 1},"; נשאר ב-"&TEXT(B{r + 3},"0.##")&" (סף יציבות)","")&IF(B{r + 5}<>B{r + 3},"; אות מדיניות → צעד למטה ל-"&TEXT(B{r + 5},"0.##"),"")&IF(B{r + 4}="","; אין תחזית ריבית ב״י",""))', BASE)
    ctx.name(f"r_dur_{kind}_why", f"'05_מחמ'!$B${r + 9}")
    # סימולציה: מדדים במח״מ הידני (אינטרפולציה על עמודת המח״מ)
    put(ws, f"E{r}", "סימולציה — מדדים במח״מ שבשימוש", BOLD)
    hdr(ws, r + 1, ["", "מח״מ", "תשואה 12 ח׳ %", "Breakeven (נ״ב)", "P&L +תרחיש %", "P&L −תרחיש %"], 5)
    for j, (lab, ref) in enumerate((("כלל הציון", f"B{r + 6}"), ("בשימוש", f"B{r + 8}"))):
        rr = r + 2 + j
        put(ws, f"E{rr}", lab, BOLD, None, None, BOX); put(ws, f"F{rr}", f'=IF({ref}="","",{ref})', BASE, None, "0.00", BOX)
        put(ws, f"K{rr}", f'=IF(F{rr}="","",MAX(1,MIN(11,IFERROR(MATCH(F{rr},{T["dur"]},1),1))))', GREY, None, "0")
        put(ws, f"L{rr}", f'=IF(K{rr}="","",MAX(0,MIN(1,(F{rr}-INDEX({T["dur"]},K{rr}))/(INDEX({T["dur"]},K{rr}+1)-INDEX({T["dur"]},K{rr})))))', GREY, None, "0.000")
        def at_dur(col): return f'(INDEX({col},K{rr})+(INDEX({col},K{rr}+1)-INDEX({col},K{rr}))*L{rr})'
        put(ws, f"G{rr}", f'=IF(F{rr}="","",{at_dur(T["exp"])})', BASE, None, "0.00", BOX)
        put(ws, f"H{rr}", f'=IF(F{rr}="","",{at_dur(T["be"])})', BASE, None, "0.0", BOX)
        put(ws, f"I{rr}", f'=IF(F{rr}="","",-F{rr}*p_bb_shock_bp/100)', BASE, None, "0.00", BOX); put(ws, f"J{rr}", f'=IF(F{rr}="","",F{rr}*p_bb_shock_bp/100)', BASE, None, "0.00", BOX)


def sheet_duration(ctx: Ctx):
    ws = ctx.sheet("05_מחמ")
    put(ws, "A1", "מח״מ — איזה טווח כדאי להחזיק: תשואה צפויה 12 ח׳ מול סיכון, לשקלי ולצמוד בנפרד ואז משוקלל", TITLE)
    put(ws, "A3", "ציון = (עודף מול שנה − חיכוך) / (מח״מ × תנודתיות). Breakeven = כמה נ״ב של עליית תשואות מאפסות את ההארכה. מח״מ מומלץ = הטווח עם הציון הגבוה ביותר בתוך המגבלות, עם סף יציבות ואות מדיניות (01).", NOTE)
    pol = ctx.frames["נתוני_בי_ריבית"]
    if len(pol.columns):
        item, val = dcol(ctx, "נתוני_בי_ריבית", "item"), dcol(ctx, "נתוני_בי_ריבית", "value")
        put(ws, "N3", "תחזית ריבית ב״י לסוף השנה הבאה %", GREY); put(ws, "N4", f'=IFERROR(IF(INDEX({val},MATCH("forecast_end_next_year",{item},0))="","",VALUE(INDEX({val},MATCH("forecast_end_next_year",{item},0)))),"")', GREY, None, "0.00")
        put(ws, "O3", "פורוורד שנה-בעוד-שנה %", GREY); put(ws, "O4", "=INDEX(cur_fwd_n,MATCH(2,tenors,0))", GREY, None, "0.00")
        put(ws, "P3", "אות מדיניות (נ״ב)", GREY); put(ws, "P4", '=IF(OR(N4="",O4=""),"",(O4-N4)*100)', GREY, None, "0")
    else:
        put(ws, "P4", "", GREY)
    ctx.name("policy_gap_bp", "'05_מחמ'!$P$4")
    put(ws, "A5", "שקלי", BOLD); Tn = _dur_table(ctx, ws, 7, "nominal")
    put(ws, "A20", "צמוד", BOLD); Tr = _dur_table(ctx, ws, 22, "real")
    put(ws, "A35", "החלטת המודל — שקלי", BOLD); _dur_pick(ctx, ws, 36, Tn, "nominal")
    put(ws, "A47", "החלטת המודל — צמוד", BOLD); _dur_pick(ctx, ws, 48, Tr, "real")
    put(ws, "A59", "מח״מ משוקלל", BOLD); hdr(ws, 60, ["רכיב", "ערך", "הסבר"])
    put(ws, "A61", "מח״מ משוקלל — המודל", BOLD, None, None, BOX); put(ws, "B61", '=IF(OR(m_dur_nominal="",m_dur_real=""),"",m_linked*m_dur_real+(1-m_linked)*m_dur_nominal)', BOLD, GRN, "0.00", BOX)
    put(ws, "C61", "משקל צמוד (המודל) × מח״מ צמוד + משקל שקלי × מח״מ שקלי", NOTE)
    put(ws, "A62", "מח״מ משוקלל — בשימוש", BOLD, None, None, BOX); put(ws, "B62", '=IF(OR(u_dur_nominal="",u_dur_real=""),"",u_linked*u_dur_real+(1-u_linked)*u_dur_nominal)', BOLD, None, "0.00", BOX)
    put(ws, "C62", "לפי המשקל והמח״מים בשימוש (סימולציה אם מולאה)", NOTE)
    ctx.name("m_dur_weighted", "'05_מחמ'!$B$61"); ctx.name("u_dur_weighted", "'05_מחמ'!$B$62")
    put(ws, "A63", "למה", BOLD); put(ws, "B63", '=IF(B61="","",TEXT(m_linked,"0%")&" × צמוד "&TEXT(m_dur_real,"0.00")&" + "&TEXT(1-m_linked,"0%")&" × שקלי "&TEXT(m_dur_nominal,"0.00"))', BASE)
    ctx.name("m_dur_weighted_why", "'05_מחמ'!$B$63")
    for col, w in zip("ABCDEFGHIJKL", (30, 14, 14, 12, 14, 16, 14, 16, 18, 10, 8, 10)): ws.column_dimensions[col].width = w
    ws.column_dimensions.group("L", "L", outline_level=1, hidden=True)
    ws.freeze_panes = "A5"


# ---------------------------------------------------------------- 06_סינטטי_בולט
NLIST = gov_decide.NLIST                     # ניירות לאפיק ברשימה (ממוינים לפי מח״מ) — זהה לפייתון
PAIRS = [(i, j) for i in range(1, NLIST + 1) for j in range(i + 1, NLIST + 1)]
NPAIRS = len(PAIRS)                          # 435
STRUCT_BLOCKS: dict = {}


def _struct_block(ctx: Ctx, ws, r0: int, kind: str) -> int:
    """בלוק לאפיק: רשימת ניירות, טבלת זוגות, בולטים, המודל והסימולציה. מחזיר את השורה הבאה הפנויה."""
    heb = "שקלי" if kind == "nominal" else "צמוד"; dcolname = "b_dn" if kind == "nominal" else "b_dr"
    tgt = f"u_dur_{kind}"
    put(ws, f"A{r0}", f"אפיק {heb}", Font(name="Arial", size=12, bold=True)); put(ws, f"C{r0}", "מח״מ יעד (בשימוש, 05):", BOLD); put(ws, f"D{r0}", f"={tgt}", GREEN, GRN, "0.00", BOX)
    put(ws, f"F{r0}", f'=IF(COUNT({dcolname})>{NLIST},"אזהרה: "&COUNT({dcolname})&" ניירות באפיק — הרשימה מכילה רק את {NLIST} הקצרים במח״מ","")', Font(name="Arial", size=10, bold=True, color="C00000"))
    ctx.name(f"warn_list_{kind}", f"'06_סינטטי_בולט'!$F${r0}")
    ctx.name(f"t_{kind}", f"'06_סינטטי_בולט'!$D${r0}")
    # --- רשימת ניירות (ממוינת לפי מח״מ) ---
    lr = r0 + 2
    hdr(ws, lr, ["#", "מח״מ", "מפתח", "שם", "שנים לפדיון", "תשואה %", "קמירות", "מחזור", "Roll-down (נ״ב)", "תשואה 12 ח׳ %"], 1, 30)
    for i in range(NLIST):
        r = lr + 1 + i
        put(ws, f"A{r}", i + 1, GREY, None, "0"); put(ws, f"B{r}", f'=IFERROR(SMALL({dcolname},{i + 1}),"")', BASE, None, "0.000", BOX)
        put(ws, f"C{r}", f'=IF(B{r}="","",INDEX(b_key,MATCH(B{r},{dcolname},0)))', BASE, None, None, BOX)
        for col, nm, fmt in (("D", "b_name", None), ("E", "b_ttm", "0.00"), ("F", "b_ytm", "0.00"), ("G", "b_conv", "0.0"), ("H", "b_vol", "#,##0"), ("I", "b_roll", "0.0"), ("J", "b_exp12", "0.00")):
            put(ws, f"{col}{r}", f'=IF(C{r}="","",INDEX({nm},MATCH(B{r},{dcolname},0)))', BASE, None, fmt, BOX)
    LST = {k: f"$'{c}'${lr + 1}:${c}${lr + NLIST}".replace("'", "") for k, c in (("d", "B"), ("key", "C"), ("ttm", "E"), ("ytm", "F"), ("conv", "G"), ("vol", "H"), ("roll", "I"), ("exp", "J"))}
    # --- זוגות ---
    pr = lr + NLIST + 2
    put(ws, f"A{pr}", "סריקת כל הזוגות (קצר מהיעד + ארוך מהיעד)", BOLD)
    cols = ["i", "j", "מח״מ קצר", "מח״מ ארוך", "תקין?", "משקל קצר", "משקל ארוך", "תשואה %", "קמירות", "מחזור מינ׳", "תשואה 12 ח׳ %", "P&L +מקבילי", "P&L −מקבילי", "P&L התללה", "P&L השתטחות", "ציון", "נזיל?", "ציון תקין"]
    hdr(ws, pr + 1, cols, 1, 30)
    c = {k: L(n + 1) for n, k in enumerate(["i", "j", "di", "dj", "valid", "wi", "wj", "ytm", "conv", "liq", "exp", "up", "dn", "steep", "flat", "score", "ok", "sc_ok"])}
    p0 = pr + 2
    def leg(idx, w):   # P&L לרגל: (−D·dy + ½C·dy²)·100 ; dy בשבר עשרוני
        D = f"INDEX({LST['d']},{idx})"; C_ = f"INDEX({LST['conv']},{idx})"; T_ = f"INDEX({LST['ttm']},{idx})"
        dy = "(p_bb_shock_bp/10000)"; tw = f"(p_bb_twist_bp/10000*({T_}-$D${r0})/p_bb_twist_span)"
        return {"up": f"{w}*(-{D}*{dy}+0.5*{C_}*{dy}^2)*100", "dn": f"{w}*({D}*{dy}+0.5*{C_}*{dy}^2)*100",
                "steep": f"{w}*(-{D}*{tw}+0.5*{C_}*{tw}^2)*100", "flat": f"{w}*({D}*{tw}+0.5*{C_}*{tw}^2)*100"}
    for n, (i, j) in enumerate(PAIRS):
        r = p0 + n
        put(ws, f"{c['i']}{r}", i, GREY, None, "0"); put(ws, f"{c['j']}{r}", j, GREY, None, "0")
        put(ws, f"{c['di']}{r}", f'=INDEX({LST["d"]},{i})', GREY, None, "0.000"); put(ws, f"{c['dj']}{r}", f'=INDEX({LST["d"]},{j})', GREY, None, "0.000")
        di, dj = f"{c['di']}{r}", f"{c['dj']}{r}"
        put(ws, f"{c['valid']}{r}", f'=IF(AND(ISNUMBER({di}),ISNUMBER({dj}),ISNUMBER($D${r0})),IF(AND({di}<$D${r0},{dj}>$D${r0}),"כן","לא"),"לא")', GREY)
        v = f"{c['valid']}{r}"
        put(ws, f"{c['wi']}{r}", f'=IF({v}="כן",({dj}-$D${r0})/({dj}-{di}),"")', BASE, None, "0.000"); put(ws, f"{c['wj']}{r}", f'=IF({v}="כן",1-{c["wi"]}{r},"")', BASE, None, "0.000")
        wi, wj = f"{c['wi']}{r}", f"{c['wj']}{r}"
        put(ws, f"{c['ytm']}{r}", f'=IF({v}="כן",{wi}*INDEX({LST["ytm"]},{i})+{wj}*INDEX({LST["ytm"]},{j}),"")', BASE, None, "0.00")
        put(ws, f"{c['conv']}{r}", f'=IF({v}="כן",{wi}*INDEX({LST["conv"]},{i})+{wj}*INDEX({LST["conv"]},{j}),"")', BASE, None, "0.0")
        put(ws, f"{c['liq']}{r}", f'=IF({v}="כן",MIN(INDEX({LST["vol"]},{i}),INDEX({LST["vol"]},{j})),"")', BASE, None, "#,##0")
        put(ws, f"{c['exp']}{r}", f'=IF({v}="כן",{wi}*INDEX({LST["exp"]},{i})+{wj}*INDEX({LST["exp"]},{j}),"")', BASE, None, "0.00")
        li, lj = leg(i, wi), leg(j, wj)
        for k in ("up", "dn", "steep", "flat"):
            put(ws, f"{c[k]}{r}", f'=IF({v}="כן",{li[k]}+{lj[k]},"")', BASE, None, "0.00")
        put(ws, f"{c['score']}{r}", f'=IF({v}="כן",{c["exp"]}{r}+({c["up"]}{r}+{c["dn"]}{r}+{c["steep"]}{r}+{c["flat"]}{r})/5,"")', BASE, None, "0.000")
        put(ws, f"{c['ok']}{r}", f'=IF({v}="כן",IF({c["liq"]}{r}>=p_bb_min_volume_k,"כן","לא"),"לא")', GREY)
        put(ws, f"{c['sc_ok']}{r}", f'=IF({c["ok"]}{r}="כן",{c["score"]}{r},-9999)', GREY, None, "0.000")
    PR = {k: f"${c[k]}${p0}:${c[k]}${p0 + NPAIRS - 1}" for k in c}
    # --- בולטים ---
    br = p0 + NPAIRS + 1
    put(ws, f"A{br}", "בולטים סביב היעד (|מח״מ − יעד| ≤ סטייה מותרת)", BOLD)
    hdr(ws, br + 1, ["#", "מח״מ", "מפתח", "מתאים?", "תשואה %", "קמירות", "מחזור", "תשואה 12 ח׳ %", "P&L +מקבילי", "P&L −מקבילי", "P&L התללה", "P&L השתטחות", "ציון", "נזיל?", "ציון תקין"], 1, 30)
    b0 = br + 2
    bc = {k: L(n + 1) for n, k in enumerate(["i", "d", "key", "fit", "ytm", "conv", "liq", "exp", "up", "dn", "steep", "flat", "score", "ok", "sc_ok"])}
    for i in range(1, NLIST + 1):
        r = b0 + i - 1
        put(ws, f"A{r}", i, GREY, None, "0"); put(ws, f"B{r}", f'=INDEX({LST["d"]},{i})', GREY, None, "0.000"); put(ws, f"C{r}", f'=INDEX({LST["key"]},{i})', BASE)
        put(ws, f"D{r}", f'=IF(AND(ISNUMBER(B{r}),ISNUMBER($D${r0})),IF(ABS(B{r}-$D${r0})<=p_bb_tol,"כן","לא"),"לא")', GREY)
        fit = f"D{r}"
        for k, src, fmt in (("ytm", "ytm", "0.00"), ("conv", "conv", "0.0"), ("liq", "vol", "#,##0"), ("exp", "exp", "0.00")):
            put(ws, f"{bc[k]}{r}", f'=IF({fit}="כן",INDEX({LST[src]},{i}),"")', BASE, None, fmt)
        lg = leg(i, "1")
        for k in ("up", "dn", "steep", "flat"):
            put(ws, f"{bc[k]}{r}", f'=IF({fit}="כן",{lg[k]},"")', BASE, None, "0.00")
        put(ws, f"{bc['score']}{r}", f'=IF({fit}="כן",{bc["exp"]}{r}+({bc["up"]}{r}+{bc["dn"]}{r}+{bc["steep"]}{r}+{bc["flat"]}{r})/5,"")', BASE, None, "0.000")
        put(ws, f"{bc['ok']}{r}", f'=IF({fit}="כן",IF({bc["liq"]}{r}>=p_bb_min_volume_k,"כן","לא"),"לא")', GREY)
        put(ws, f"{bc['sc_ok']}{r}", f'=IF({bc["ok"]}{r}="כן",{bc["score"]}{r},-9999)', GREY, None, "0.000")
    BR = {k: f"${bc[k]}${b0}:${bc[k]}${b0 + NLIST - 1}" for k in bc}
    # --- המודל ---
    mr = b0 + NLIST + 1
    put(ws, f"A{mr}", f"החלטת המודל — {heb}", BOLD); hdr(ws, mr + 1, ["רכיב", "ערך", "ערך 2", "הסבר"])
    bp = f"MATCH(MAX({PR['sc_ok']}),{PR['sc_ok']},0)"; bb = f"MATCH(MAX({BR['sc_ok']}),{BR['sc_ok']},0)"
    has_p = f"MAX({PR['sc_ok']})>-9999"; has_b = f"MAX({BR['sc_ok']})>-9999"
    rows = [("ציון הזוג הטוב ביותר", f'=IF({has_p},MAX({PR["sc_ok"]}),"")', "", "ממוצע תשואה 12 ח׳ על 5 תרחישים (בסיס, ±מקבילי, התללה, השתטחות); רק זוגות נזילים", "0.000"),
            ("הזוג: מפתח קצר / משקל", f'=IF({has_p},INDEX({LST["key"]},INDEX({PR["i"]},{bp})),"")', f'=IF({has_p},INDEX({PR["wi"]},{bp}),"")', "", "0.000"),
            ("הזוג: מפתח ארוך / משקל", f'=IF({has_p},INDEX({LST["key"]},INDEX({PR["j"]},{bp})),"")', f'=IF({has_p},INDEX({PR["wj"]},{bp}),"")', "", "0.000"),
            ("ציון הבולט הטוב ביותר", f'=IF({has_b},MAX({BR["sc_ok"]}),"")', "", "", "0.000"),
            ("הבולט: מפתח", f'=IF({has_b},INDEX({BR["key"]},{bb}),"")', "", "", None),
            ("פער סינתטי−בולט (נ״ב)", f'=IF(OR(B{mr + 2}="",B{mr + 5}=""),"",(B{mr + 2}-B{mr + 5})*100)', "", "חיובי = יתרון לסינתטי", "0"),
            ("פסק דין", f'=IF(AND(B{mr + 2}="",B{mr + 5}=""),"אין זוג מתאים",IF(B{mr + 2}="","אין זוג מתאים",IF(B{mr + 5}="","סינתטי",IF(B{mr + 7}>p_bb_friction_bp,"סינתטי",IF(B{mr + 7}<-p_bb_friction_bp,"בולט","דומה")))))', "", "סינתטי רק אם היתרון גדול מהחיכוך; 'דומה' → בולט (פשוט יותר)", None)]
    for n, (a, b, c2, d, f) in enumerate(rows):
        r = mr + 2 + n
        put(ws, f"A{r}", a, BOLD if n == 6 else BASE, None, None, BOX); put(ws, f"B{r}", b, BOLD if n == 6 else BASE, GRN if n == 6 else None, f, BOX)
        put(ws, f"C{r}", c2, BASE, None, f, BOX); put(ws, f"D{r}", d, NOTE)
    verdict = f"B{mr + 8}"
    # רכיבי המבנה שנבחר (לתיק המודל): עד 2 מפתחות ומשקלים
    put(ws, f"A{mr + 9}", "מבנה נבחר — מפתח 1 / משקל", BASE, None, None, BOX)
    put(ws, f"B{mr + 9}", f'=IF({verdict}="סינתטי",B{mr + 3},IF(B{mr + 6}<>"",B{mr + 6},""))', BASE, None, None, BOX)
    put(ws, f"C{mr + 9}", f'=IF({verdict}="סינתטי",C{mr + 3},IF(B{mr + 6}<>"",1,""))', BASE, None, "0.000", BOX)
    put(ws, f"A{mr + 10}", "מבנה נבחר — מפתח 2 / משקל", BASE, None, None, BOX)
    put(ws, f"B{mr + 10}", f'=IF({verdict}="סינתטי",B{mr + 4},"")', BASE, None, None, BOX); put(ws, f"C{mr + 10}", f'=IF({verdict}="סינתטי",C{mr + 4},"")', BASE, None, "0.000", BOX)
    put(ws, f"A{mr + 11}", "תיאור", BASE, None, None, BOX)
    put(ws, f"B{mr + 11}", f'={verdict}&IF(B{mr + 9}<>""," "&B{mr + 9}&"×"&TEXT(C{mr + 9},"0%")&IF(B{mr + 10}<>"","+"&B{mr + 10}&"×"&TEXT(C{mr + 10},"0%"),""),"")', BOLD, None, None, BOX)
    put(ws, f"A{mr + 12}", "למה", BASE, None, None, BOX)
    put(ws, f"B{mr + 12}", f'=IF(B{mr + 2}="",IF(B{mr + 5}="","אין נייר ארוך וקצר סביב היעד ואין בולט נזיל","אין נייר ארוך/קצר מהיעד; הבולט "&B{mr + 6}&" משמש כמבנה"),IF(B{mr + 5}="","אין בולט בטווח הסטייה; הזוג הטוב ביותר נבחר","ממוצע 5 תרחישים: סינתטי "&TEXT(B{mr + 2},"0.00")&"% מול בולט "&TEXT(B{mr + 5},"0.00")&"% (פער "&TEXT(B{mr + 7},"+0;-0")&" נ״ב, חיכוך "&TEXT(p_bb_friction_bp,"0")&")"))', BASE)
    ctx.name(f"m_struct_{kind}", f"'06_סינטטי_בולט'!$B${mr + 11}"); ctx.name(f"m_struct_{kind}_why", f"'06_סינטטי_בולט'!$B${mr + 12}"); ctx.name(f"m_struct_{kind}_verdict", f"'06_סינטטי_בולט'!$B${mr + 8}")
    for n in (1, 2):
        ctx.name(f"s_{kind}_id{n}", f"'06_סינטטי_בולט'!$B${mr + 8 + n}"); ctx.name(f"s_{kind}_w{n}", f"'06_סינטטי_בולט'!$C${mr + 8 + n}")
    # --- סימולציה ---
    sr = mr + 14
    put(ws, f"A{sr}", f"סימולציה — {heb} (מפתחות ידניים ב-01 → סימולציה)", BOLD)
    hdr(ws, sr + 1, ["רכיב", "מפתח", "מח״מ", "משקל", "תשואה %", "קמירות", "מחזור", "תשואה 12 ח׳ %", "P&L +מקבילי", "P&L −מקבילי", "P&L התללה", "P&L השתטחות", "ציון"], 1, 30)
    def sim_leg(r, w, keycell):
        kk = f"MATCH(TEXT({keycell},\"0\"),b_key,0)"; D = f"INDEX(b_mod,{kk})"; C_ = f"INDEX(b_conv,{kk})"; T_ = f"INDEX(b_ttm,{kk})"
        dy = "(p_bb_shock_bp/10000)"; tw = f"(p_bb_twist_bp/10000*({T_}-$D${r0})/p_bb_twist_span)"
        put(ws, f"C{r}", f'=IF({keycell}="","",IFERROR({D},"לא נמצא"))', BASE, None, "0.000", BOX)
        put(ws, f"E{r}", f'=IF(OR({keycell}="",{w}=""),"",IFERROR({w}*INDEX(b_ytm,{kk}),""))', BASE, None, "0.00", BOX)
        put(ws, f"F{r}", f'=IF(OR({keycell}="",{w}=""),"",IFERROR({w}*{C_},""))', BASE, None, "0.0", BOX)
        put(ws, f"G{r}", f'=IF({keycell}="","",IFERROR(INDEX(b_vol,{kk}),""))', BASE, None, "#,##0", BOX)
        put(ws, f"H{r}", f'=IF(OR({keycell}="",{w}=""),"",IFERROR({w}*INDEX(b_exp12,{kk}),""))', BASE, None, "0.00", BOX)
        put(ws, f"I{r}", f'=IF(OR({keycell}="",{w}=""),"",IFERROR({w}*(-{D}*{dy}+0.5*{C_}*{dy}^2)*100,""))', BASE, None, "0.00", BOX)
        put(ws, f"J{r}", f'=IF(OR({keycell}="",{w}=""),"",IFERROR({w}*({D}*{dy}+0.5*{C_}*{dy}^2)*100,""))', BASE, None, "0.00", BOX)
        put(ws, f"K{r}", f'=IF(OR({keycell}="",{w}=""),"",IFERROR({w}*(-{D}*{tw}+0.5*{C_}*{tw}^2)*100,""))', BASE, None, "0.00", BOX)
        put(ws, f"L{r}", f'=IF(OR({keycell}="",{w}=""),"",IFERROR({w}*({D}*{tw}+0.5*{C_}*{tw}^2)*100,""))', BASE, None, "0.00", BOX)
    s1, s2, s3 = sr + 2, sr + 3, sr + 4
    put(ws, f"A{s1}", "קצר", BASE, None, None, BOX); put(ws, f"B{s1}", f'=IF(p_sim_short_{kind}="","",p_sim_short_{kind})', BLUE, None, "0", BOX)
    put(ws, f"A{s2}", "ארוך", BASE, None, None, BOX); put(ws, f"B{s2}", f'=IF(p_sim_long_{kind}="","",p_sim_long_{kind})', BLUE, None, "0", BOX)
    put(ws, f"A{s3}", "בולט", BASE, None, None, BOX); put(ws, f"B{s3}", f'=IF(p_sim_bullet_{kind}="","",p_sim_bullet_{kind})', BLUE, None, "0", BOX)
    put(ws, f"D{s1}", f'=IF(OR(C{s1}="",C{s2}="",NOT(ISNUMBER(C{s1})),NOT(ISNUMBER(C{s2})),C{s1}=C{s2}),"",(C{s2}-$D${r0})/(C{s2}-C{s1}))', BASE, None, "0.000", BOX)
    put(ws, f"D{s2}", f'=IF(D{s1}="","",1-D{s1})', BASE, None, "0.000", BOX); put(ws, f"D{s3}", f'=IF(C{s3}="","",1)', BASE, None, "0.000", BOX)
    sim_leg(s1, f"D{s1}", f"B{s1}"); sim_leg(s2, f"D{s2}", f"B{s2}"); sim_leg(s3, f"D{s3}", f"B{s3}")
    put(ws, f"A{s3 + 1}", "סינתטי (קצר+ארוך)", BOLD, None, None, BOX)
    for col in "EFHIJKL": put(ws, f"{col}{s3 + 1}", f'=IF(OR({col}{s1}="",{col}{s2}=""),"",{col}{s1}+{col}{s2})', BOLD, None, "0.00", BOX)
    put(ws, f"G{s3 + 1}", f'=IF(OR(G{s1}="",G{s2}=""),"",MIN(G{s1},G{s2}))', BOLD, None, "#,##0", BOX)
    put(ws, f"M{s3 + 1}", f'=IF(H{s3 + 1}="","",H{s3 + 1}+(I{s3 + 1}+J{s3 + 1}+K{s3 + 1}+L{s3 + 1})/5)', BOLD, None, "0.000", BOX)
    put(ws, f"M{s3}", f'=IF(H{s3}="","",H{s3}+(I{s3}+J{s3}+K{s3}+L{s3})/5)', BOLD, None, "0.000", BOX)
    put(ws, f"A{s3 + 2}", "פסק דין — סימולציה", BOLD, None, None, BOX)
    put(ws, f"B{s3 + 2}", f'=IF(AND(M{s3 + 1}="",M{s3}=""),"",IF(M{s3 + 1}="","בולט",IF(M{s3}="","סינתטי",IF((M{s3 + 1}-M{s3})*100>p_bb_friction_bp,"סינתטי עדיף",IF((M{s3 + 1}-M{s3})*100<-p_bb_friction_bp,"בולט עדיף","דומה")))))', BOLD, None, None, BOX)
    put(ws, f"C{s3 + 2}", f'=IF(D{s1}="","",IF(ABS(D{s1}*C{s1}+D{s2}*C{s2}-$D${r0})<=0.01,"מח״מ הסינתטי = היעד","סטיית מח״מ"))', NOTE)
    ctx.name(f"sim_struct_{kind}", f"'06_סינטטי_בולט'!$B${s3 + 2}")
    STRUCT_BLOCKS[kind] = {"r0": r0, "list_r0": lr + 1, "pairs_r0": p0, "bullets_r0": b0, "model_r0": mr, "c": c, "bc": bc}
    return s3 + 4


def sheet_structure(ctx: Ctx):
    ws = ctx.sheet("06_סינטטי_בולט")
    put(ws, "A1", "סינתטי מול בולט — סריקת כל הזוגות סביב מח״מ היעד מול כל בולט קרוב, ב-5 תרחישים", TITLE)
    put(ws, "A3", "לכל זוג (קצר<יעד<ארוך): משקלים שמביאים למח״מ היעד; תשואה, קמירות, roll-down, נזילות (מחזור מינימלי) ו-P&L במקבילי ±, התללה והשתטחות (Δתשואה(טווח) = ±נ״ב × (שנים לפדיון − יעד)/טווח הנרמול). ציון = ממוצע התשואה ל-12 ח׳ על 5 התרחישים.", NOTE)
    nxt = _struct_block(ctx, ws, 5, "nominal")
    _struct_block(ctx, ws, nxt + 1, "real")
    for col, w in zip("ABCDEFGHIJKLMNOPQR", (28, 12, 14, 24, 12, 10, 10, 10, 10, 11, 12, 11, 11, 11, 11, 10, 7, 10)): ws.column_dimensions[col].width = w
    ws.freeze_panes = "A5"


# ---------------------------------------------------------------- 07_סדרות
def sheet_series(ctx: Ctx):
    ws = ctx.sheet("07_סדרות")
    put(ws, "A1", "סדרות — זול/יקר מול העקום ביחס להיסטוריה של הסדרה (Z), roll-down, נזילות; 5 מובילות לכל אפיק; זוגות החלפה", TITLE)
    put(ws, "A3", "ציון = 0.4×Z/3 + 0.3×(2×אחוזון roll−1) + 0.3×(2×אחוזון נזילות−1); בלי היסטוריה (פחות תצפיות מהמינימום) המשקלים מנורמלים על roll ונזילות. אחוזונים בתוך האפיק. החלפה: אותו אפיק, |Δמח״מ| ≤ מותר, פיקאפ ≥ מינימום.", NOTE)
    hist = ctx.frames["נתוני_סדרות_היסט"]; has_hist = len(hist.columns) > 0
    HK = dcol(ctx, "נתוני_סדרות_היסט", "key") if has_hist else None; HS = dcol(ctx, "נתוני_סדרות_היסט", "spread_bp") if has_hist else None
    # --- מובילות ---
    put(ws, "A5", "5 הסדרות המובילות", BOLD)
    hdr(ws, 6, ["דירוג", "שקלי — מפתח", "שם", "ציון", "", "צמוד — מפתח", "שם", "ציון"])
    T0 = 34; T1 = T0 + NB - 1                      # טבלת הסדרות המלאה (אחרי 16 שורות ההחלפה)
    SEG, KEY, NAME, SC, RK = (f"$A${T0}:$A${T1}", f"$B${T0}:$B${T1}", f"$C${T0}:$C${T1}", f"$O${T0}:$O${T1}", f"$P${T0}:$P${T1}")
    for i in range(5):
        r = 7 + i; put(ws, f"A{r}", i + 1, BASE, None, "0", BOX)
        for seg, kc, nc, sc in (("N", "B", "C", "D"), ("R", "F", "G", "H")):
            arr(ws, f"{kc}{r}", f'=IFERROR(INDEX({KEY},MATCH(1,({SEG}="{seg}")*({RK}={i + 1}),0)),"")')
            put(ws, f"{nc}{r}", f'=IF({kc}{r}="","",INDEX(b_name,MATCH({kc}{r},b_key,0)))', BASE, None, None, BOX)
            put(ws, f"{sc}{r}", f'=IF({kc}{r}="","",INDEX({SC},MATCH({kc}{r},{KEY},0)))', BASE, None, "0.000", BOX)
            ws[f"{kc}{r}"].border = BOX
    ctx.name("top5_nominal", f"'07_סדרות'!$B$7:$B$11"); ctx.name("top5_real", f"'07_סדרות'!$F$7:$F$11")
    # --- החלפות ---
    put(ws, "A13", "זוגות החלפה (לכל אחת מ-8 המובילות באפיק — הנייר הקרוב במח״מ שמולו הפיקאפ הגבוה ביותר)", BOLD)
    hdr(ws, 14, ["אפיק", "לקנות (מפתח)", "שם", "פיקאפ (נ״ב)", "למכור (מפתח)", "שם", "Δמח״מ", "מוצג?"])
    DUR, SPR = f"$F${T0}:$F${T1}", f"$G${T0}:$G${T1}"
    for n in range(16):
        r = 15 + n; seg = "N" if n < 8 else "R"; rank = n % 8 + 1
        put(ws, f"A{r}", "שקלי" if seg == "N" else "צמוד", BASE, None, None, BOX)
        arr(ws, f"B{r}", f'=IFERROR(INDEX({KEY},MATCH(1,({SEG}="{seg}")*({RK}={rank}),0)),"")')
        put(ws, f"C{r}", f'=IF(B{r}="","",INDEX(b_name,MATCH(B{r},b_key,0)))', BASE, None, None, BOX)
        bd = f"INDEX({DUR},MATCH(B{r},{KEY},0))"; bs = f"INDEX({SPR},MATCH(B{r},{KEY},0))"
        cond = f'({SEG}="{seg}")*({KEY}<>B{r})*(IFERROR(ABS({DUR}-{bd}),999)<=p_switch_dur_tol)'
        arr(ws, f"D{r}", f'=IF(B{r}="","",IFERROR(IF(MAX(IF({cond},IFERROR({bs}-{SPR},-9999),-9999))=-9999,"",MAX(IF({cond},IFERROR({bs}-{SPR},-9999),-9999))),""))')
        arr(ws, f"E{r}", f'=IF(OR(B{r}="",D{r}="",D{r}=-9999),"",IFERROR(INDEX({KEY},MATCH(1,{cond}*(IFERROR(ABS(({bs}-{SPR})-D{r}),9)<0.0001),0)),""))')
        put(ws, f"F{r}", f'=IF(E{r}="","",INDEX(b_name,MATCH(E{r},b_key,0)))', BASE, None, None, BOX)
        put(ws, f"G{r}", f'=IF(E{r}="","",ABS(INDEX({DUR},MATCH(E{r},{KEY},0))-{bd}))', BASE, None, "0.00", BOX)
        put(ws, f"H{r}", f'=IF(D{r}="","",IF(D{r}>=p_switch_pickup_bp,"כן","לא"))', BASE, None, None, BOX)
        for col in "BDE": ws[f"{col}{r}"].border = BOX; ws[f"{col}{r}"].number_format = "0.0" if col == "D" else "@"
    # --- הטבלה המלאה ---
    put(ws, f"A{T0 - 2}", "כל הסדרות", BOLD)
    hdr(ws, T0 - 1, ["אפיק (N/R)", "מפתח", "שם", "סוג", "שנים לפדיון", "מח״מ", "פער מול העקום (נ״ב)", "תצפיות היסטוריה", "ממוצע היסטורי", "סטיית תקן", "Z", "אחוזון roll", "אחוזון נזילות", "Roll-down (נ״ב)", "ציון", "דירוג באפיק", "מחזור", "שורה"], 1, 40)
    for k in range(NB):
        r = T0 + k; src = R0 + k
        key, kind, ttm, mod, spread, roll, vol, incl, name = (f"'02_קלט_אגח'!{c}{src}" for c in (BC["key"], BC["kind"], BC["ttm"], BC["modified"], CC["spread"], CC["roll"], BC["volume_k_nis"], CC["incl"], BC["name"]))
        put(ws, f"A{r}", f'=IF(OR({key}="",{incl}<>"כן",{spread}="",{mod}=""),"",IF(OR({kind}="שקלי",{kind}="מק""מ"),"N",IF({kind}="מדד","R","")))', GREY)
        put(ws, f"B{r}", f'=IF({key}="","",{key})', BASE, None, "@", BOX); put(ws, f"C{r}", f'=IF({key}="","",{name})', BASE, None, None, BOX)
        put(ws, f"D{r}", f'=IF({key}="","",{kind})', BASE, None, None, BOX); put(ws, f"E{r}", f'=IF({key}="","",{ttm})', BASE, None, "0.00", BOX)
        put(ws, f"F{r}", f'=IF({key}="","",{mod})', BASE, None, "0.000", BOX); put(ws, f"G{r}", f'=IF(A{r}="","",{spread})', BASE, None, "0.0", BOX)
        if has_hist:
            put(ws, f"H{r}", f'=IF(A{r}="","",SUMPRODUCT(({HK}=B{r})*ISNUMBER({HS})))', BASE, None, "0", BOX)
            put(ws, f"I{r}", f'=IF(OR(A{r}="",H{r}=0),"",AVERAGEIFS({HS},{HK},B{r}))', BASE, None, "0.0", BOX)
            arr(ws, f"J{r}", f'=IF(OR(A{r}="",H{r}<2),"",STDEV(IF(({HK}=B{r})*ISNUMBER({HS}),{HS})))')
        else:
            put(ws, f"H{r}", f'=IF(A{r}="","",0)', BASE, None, "0", BOX); put(ws, f"I{r}", "", BASE); put(ws, f"J{r}", "", BASE)
        put(ws, f"K{r}", f'=IF(A{r}="","",IF(OR(H{r}<p_ser_min_obs,J{r}="",J{r}=0),"אין היסטוריה",(G{r}-I{r})/J{r}))', BASE, None, "0.00", BOX)
        put(ws, f"N{r}", f'=IF(A{r}="","",{roll})', BASE, None, "0.0", BOX); put(ws, f"Q{r}", f'=IF(A{r}="","",{vol})', BASE, None, "#,##0", BOX)
        put(ws, f"L{r}", f'=IF(A{r}="","",(COUNTIFS($A${T0}:$A${T1},A{r},$N${T0}:$N${T1},"<"&N{r})+1)/COUNTIF($A${T0}:$A${T1},A{r}))', BASE, None, "0.00", BOX)
        put(ws, f"M{r}", f'=IF(A{r}="","",(COUNTIFS($A${T0}:$A${T1},A{r},$Q${T0}:$Q${T1},"<"&Q{r})+1)/COUNTIF($A${T0}:$A${T1},A{r}))', BASE, None, "0.00", BOX)
        put(ws, f"O{r}", f'=IF(A{r}="","",IF(ISNUMBER(K{r}),(p_ser_w_z*MAX(-3,MIN(3,K{r}))/3+p_ser_w_roll*(2*L{r}-1)+p_ser_w_liq*(2*M{r}-1))/(p_ser_w_z+p_ser_w_roll+p_ser_w_liq),(p_ser_w_roll*(2*L{r}-1)+p_ser_w_liq*(2*M{r}-1))/(p_ser_w_roll+p_ser_w_liq)))', BASE, None, "0.000", BOX)
        put(ws, f"R{r}", f"=ROW()", GREY, None, "0")
        put(ws, f"P{r}", f'=IF(A{r}="","",COUNTIFS($A${T0}:$A${T1},A{r},$O${T0}:$O${T1},">"&O{r})+COUNTIFS($A${T0}:$A${T1},A{r},$O${T0}:$O${T1},O{r},$R${T0}:$R${T1},"<"&R{r})+1)', BASE, None, "0", BOX)
        for col in "HIJ": ws[f"{col}{r}"].border = BOX
    for col, w in zip("ABCDEFGHIJKLMNOPQ", (9, 11, 22, 7, 9, 9, 12, 9, 10, 9, 9, 9, 9, 10, 9, 8, 10)): ws.column_dimensions[col].width = w
    ws.freeze_panes = f"A{T0}"; ws.auto_filter.ref = f"A{T0 - 1}:Q{T1}"
    ctx.name("ser_seg", f"'07_סדרות'!$A${T0}:$A${T1}"); ctx.name("ser_key", f"'07_סדרות'!$B${T0}:$B${T1}"); ctx.name("ser_rank", f"'07_סדרות'!$P${T0}:$P${T1}"); ctx.name("ser_score", f"'07_סדרות'!$O${T0}:$O${T1}")


# ---------------------------------------------------------------- 08_חול
F_TENORS = [1, 2, 3, 5, 7, 10, 15, 20, 30]; MARKETS = [("US", "ארה״ב", "p_fx_usd"), ("EUR", "גוש האירו (AAA)", "p_fx_eur"), ("GBP", "בריטניה", "p_fx_gbp")]


def sheet_foreign(ctx: Ctx):
    ws = ctx.sheet("08_חול")
    put(ws, "A1", "ממשלתי חו״ל — ארה״ב, גוש האירו, בריטניה: תשואה מגודרת לשקל מול שקלי באותו טווח, ולא מגודרת לפי תחזית שער", TITLE)
    put(ws, "A3", "מגודר ≈ תשואה − (ריבית שנה במטבע − ריבית שנה שקלית) [קירוב CIP, בלי בסיס]. פיקאפ = מגודר − שקלי ישראלי באותו טווח. פסק דין לטווח: 'חו״ל מגודר' אם הפיקאפ ≥ הסף בלפחות N שווקים לא-ישנים. לא מגודר = תשואה + תחזית שינוי שער (01).", NOTE)
    fo = ctx.frames["נתוני_חול"]; has = len(fo.columns) > 0
    hdr(ws, 5, ["שוק", "טווח", "תשואה %", "ריבית שנה במטבע %", "תאריך", "ישן?", "שקלי ישראל %", "ריבית שנה שקלית %", "מגודר לשקל %", "פיקאפ מגודר (נ״ב)", "תחזית שער % לשנה", "לא מגודר %", "עובר סף?"], 1, 36)
    r = 6
    if has:
        FM, FT, FY, FR, FD, FS = (dcol(ctx, "נתוני_חול", c) for c in ("market", "tenor", "yield", "rate_1y", "date", "stale"))
    for m, heb, fx in MARKETS:
        for t in F_TENORS:
            put(ws, f"A{r}", heb, BASE, None, None, BOX); put(ws, f"B{r}", t, BASE, None, "0", BOX)
            if has:
                mt = f'MATCH(1,({FM}="{m}")*({FT}={t}),0)'
                arr(ws, f"C{r}", f'=IFERROR(INDEX({FY},{mt}),"")'); arr(ws, f"D{r}", f'=IFERROR(INDEX({FR},{mt}),"")')
                arr(ws, f"E{r}", f'=IFERROR(INDEX({FD},{mt}),"")'); arr(ws, f"F{r}", f'=IFERROR(IF(OR(INDEX({FS},{mt})=1,AND(ISNUMBER(E{r}),d_asof-E{r}>{gov_decide.STALE_DAYS})),"ישן",""),"")')
                for col, fmt in (("C", "0.00"), ("D", "0.00"), ("E", "dd/mm/yyyy"), ("F", None)): ws[f"{col}{r}"].number_format = fmt or "General"; ws[f"{col}{r}"].border = BOX; ws[f"{col}{r}"].font = GREY
            else:
                for col in "CDEF": put(ws, f"{col}{r}", "", GREY, None, None, BOX)
            put(ws, f"G{r}", f"=INDEX(cur_nom,MATCH(B{r},tenors,0))", GREEN, None, "0.00", BOX); put(ws, f"H{r}", "=INDEX(cur_nom,MATCH(1,tenors,0))", GREEN, None, "0.00", BOX)
            put(ws, f"I{r}", f'=IF(OR(C{r}="",D{r}=""),"",C{r}-(D{r}-H{r}))', BASE, None, "0.00", BOX)
            put(ws, f"J{r}", f'=IF(I{r}="","",(I{r}-G{r})*100)', BASE, None, "0", BOX)
            put(ws, f"K{r}", f'=IF({fx}="","",{fx})', BLUE, None, "0.0", BOX); put(ws, f"L{r}", f'=IF(OR(C{r}="",K{r}=""),"",C{r}+K{r})', BASE, None, "0.00", BOX)
            put(ws, f"M{r}", f'=IF(J{r}="","",IF(AND(F{r}<>"ישן",J{r}>=p_fx_pickup_bp),"כן","לא"))', BASE, None, None, BOX)
            r += 1
    FR0, FR1 = 6, r - 1
    r += 1; put(ws, f"A{r}", "פסק דין לפי טווח", BOLD); r += 1
    hdr(ws, r, ["טווח", "שווקים שעוברים את הסף", "פסק דין", "למה", "שווקים טריים"]); r += 1; vr0 = r
    for t in F_TENORS:
        put(ws, f"A{r}", t, BASE, None, "0", BOX)
        put(ws, f"B{r}", f'=COUNTIFS($B${FR0}:$B${FR1},A{r},$M${FR0}:$M${FR1},"כן")', BASE, None, "0", BOX)
        put(ws, f"E{r}", f'=SUMPRODUCT(($B${FR0}:$B${FR1}=A{r})*ISNUMBER($J${FR0}:$J${FR1})*($F${FR0}:$F${FR1}<>"ישן"))', GREY, None, "0")
        put(ws, f"C{r}", f'=IF(SUMPRODUCT(($B${FR0}:$B${FR1}=A{r})*ISNUMBER($J${FR0}:$J${FR1}))=0,"אין נתון",IF(E{r}<p_fx_min_markets,"נתון ישן",IF(B{r}>=p_fx_min_markets,"חו״ל מגודר","מקומי")))', BOLD, None, None, BOX)
        why = []
        for i, (m, heb, fx) in enumerate(MARKETS):
            rr = FR0 + i * len(F_TENORS) + F_TENORS.index(t)
            why.append(f'IF(J{rr}="","","{m} "&TEXT(J{rr},"+0;-0")&" נ״ב"&IF(F{rr}="ישן"," (ישן)","")&"; ")')
        put(ws, f"D{r}", "=" + "&".join(why), BASE)
        ctx.name(f"f_verdict_{t}y", f"'08_חול'!$C${r}"); ctx.name(f"f_why_{t}y", f"'08_חול'!$D${r}")
        r += 1
    # הקשר ארה״ב
    r += 1; put(ws, f"A{r}", "ארה״ב — ריבית הפד, UST, TIPS, Breakeven (FRED; נתוני_ארהב)", BOLD); r += 1
    hdr(ws, r, ["סדרה", "שם", "תאריך", "ערך %", "שינוי 1M (נ״ב)", "שינוי 3M (נ״ב)"]); r += 1
    us = ctx.frames["נתוני_ארהב"]
    if len(us.columns):
        for j in range(len(us)):
            for k, col in enumerate(("series", "name", "date", "value", "chg_1m_bp", "chg_3m_bp")):
                put(ws, f"{L(k + 1)}{r}", f"='נתוני_ארהב'!{L(list(us.columns).index(col) + 1)}{j + 2}", GREY, None, "dd/mm/yyyy" if col == "date" else ("0.00" if col == "value" else None), BOX)
            r += 1
        r += 1; put(ws, f"A{r}", "פער ישראל−ארה״ב 10Y (נ״ב)", BOLD)
        put(ws, f"B{r}", f'=IFERROR((INDEX(cur_nom,MATCH(10,tenors,0))-INDEX({dcol(ctx, "נתוני_ארהב", "value")},MATCH("DGS10",{dcol(ctx, "נתוני_ארהב", "series")},0)))*100,"")', BOLD, None, "0", BOX)
        ctx.name("il_us_10y_bp", f"'08_חול'!$B${r}")
    else:
        put(ws, f"A{r}", "אין נתוני ארה״ב במאגר", NOTE); ctx.name("il_us_10y_bp", f"'08_חול'!$B${r + 1}")
    for col, w in zip("ABCDEFGHIJKLM", (16, 8, 10, 14, 11, 7, 12, 12, 12, 12, 12, 11, 9)): ws.column_dimensions[col].width = w
    ws.freeze_panes = "A6"


# ---------------------------------------------------------------- 09_תיק_ותרחישים
def _num_helpers(ws, r: int):
    """עמודות עזר מספריות (L..R) לסכומי SUMPRODUCT — ריק → 0 (בלי N() שאינו תמיד מערכי באקסל)."""
    for col, src in (("L", "E"), ("M", "F"), ("N", "G"), ("O", "H"), ("P", "I"), ("Q", "J"), ("R", "K")):
        put(ws, f"{col}{r}", f'=IF(OR(B{r}="",{src}{r}="",NOT(ISNUMBER({src}{r}))),0,{src}{r})', GREY, None, "0.000")


def sheet_portfolio(ctx: Ctx):
    ws = ctx.sheet("09_תיק_ותרחישים")
    put(ws, "A1", "תיק ותרחישים — החזקות, תיק המודל, תיק סימולציה, P&L בתרחישים ומול מדד ייחוס", TITLE)
    put(ws, "A3", "החזקות: מפתח = מס׳ ני״ע (או שם לסדרה בלי מספר) וכמות נקובה. 'כלול בעקום?' = 'לא' מוציא נייר מהעקום (03). תיק המודל נבנה ממשקל הצמוד (04), המח״מים (05) והמבנים (06) שבשימוש.", NOTE)
    hdr(ws, 4, ["מפתח (מס׳ ני״ע)", "כמות נקובה מוחזקת", "שם", "סטטוס", "מפתח (טקסט)", "מפתח", "כלול בעקום? (כן/לא)", "מפתח (טקסט)"])
    hold = ctx.prev.get("holdings", {}); excl = ctx.prev.get("exclude", {})
    for i in range(60):
        r = 5 + i; k = list(hold)[i] if i < len(hold) else None; e = list(excl)[i] if i < len(excl) else None
        put(ws, f"A{r}", k, BLUE, INP, "@", BOX); put(ws, f"B{r}", hold.get(k) if k else None, BLUE, INP, "#,##0", BOX)
        put(ws, f"E{r}", f'=IF(A{r}="","",TEXT(A{r},"0"))', GREY)                       # מפתח מנורמל — מספר או טקסט → טקסט
        put(ws, f"C{r}", f'=IF(A{r}="","",IFERROR(INDEX(b_name,MATCH(E{r},b_key,0)),"לא נמצא"))', BASE, None, None, BOX)
        put(ws, f"D{r}", f'=IF(A{r}="","",IF(C{r}="לא נמצא","לא נמצא",""))', BASE, None, None, BOX)
        put(ws, f"F{r}", e, BLUE, INP, "@", BOX); put(ws, f"G{r}", excl.get(e) if e else None, BLUE, INP, None, BOX)
        put(ws, f"H{r}", f'=IF(F{r}="","",TEXT(F{r},"0"))', GREY)
    dv = DataValidation(type="list", formula1='"כן,לא"', allow_blank=True); ws.add_data_validation(dv); dv.add("G5:G64")
    ctx.name("h_key", "'09_תיק_ותרחישים'!$E$5:$E$64"); ctx.name("h_amt", "'09_תיק_ותרחישים'!$B$5:$B$64")
    ctx.name("ex_key", "'09_תיק_ותרחישים'!$H$5:$H$64"); ctx.name("ex_val", "'09_תיק_ותרחישים'!$G$5:$G$64")
    ws.column_dimensions.group("E", "E", outline_level=1, hidden=True); ws.column_dimensions.group("H", "H", outline_level=1, hidden=True)
    put(ws, "J4", "החזקות שלא נמצאו", BOLD); put(ws, "J5", '=COUNTIF($D$5:$D$64,"לא נמצא")', BOLD, RED, "0", BOX); ctx.name("n_not_found", "'09_תיק_ותרחישים'!$J$5")
    # --- תיק המודל ---
    put(ws, "A67", "תיק המודל (משקל צמוד × מח״מ × מבנה; מק״מ כרצפת נזילות; משתנה ידני)", BOLD)
    hdr(ws, 68, ["רכיב", "מפתח", "שם", "אפיק", "משקל", "מח״מ", "תשואה %", "קמירות", "שנים לפדיון", "צמוד?", "שקול 12 ח׳ %"])
    nom_share = "(1-u_linked)"; struct_share = f"MAX(0,{nom_share}-p_makam_min-p_float_weight)"
    arr(ws, "N68", '=IFERROR(INDEX(b_key,MATCH(MIN(IF((b_kind="מק""מ")*ISNUMBER(b_ttm),b_ttm)),IF(b_kind="מק""מ",b_ttm),0)),"")'); put(ws, "M68", "מק״מ הקצר ביותר:", GREY)
    arr(ws, "N69", '=IFERROR(INDEX(b_key,MATCH(MIN(IF((b_kind="משתנה")*ISNUMBER(b_ttm),b_ttm)),IF(b_kind="משתנה",b_ttm),0)),"")'); put(ws, "M69", "משתנה הקצר ביותר:", GREY)
    comps = [("שקלי — רגל 1", "=IF(s_nominal_id1=\"\",\"\",s_nominal_id1)", f"=IF(s_nominal_id1=\"\",\"\",s_nominal_w1*{struct_share})"),
             ("שקלי — רגל 2", "=IF(s_nominal_id2=\"\",\"\",s_nominal_id2)", f"=IF(s_nominal_id2=\"\",\"\",s_nominal_w2*{struct_share})"),
             ("צמוד — רגל 1", "=IF(s_real_id1=\"\",\"\",s_real_id1)", "=IF(s_real_id1=\"\",\"\",s_real_w1*u_linked)"),
             ("צמוד — רגל 2", "=IF(s_real_id2=\"\",\"\",s_real_id2)", "=IF(s_real_id2=\"\",\"\",s_real_w2*u_linked)"),
             ("מק״מ (נזילות)", '=IF(OR(p_makam_min=0,$N$68=""),"",$N$68)', '=IF(OR(p_makam_min=0,$N$68=""),"",p_makam_min)'),
             ("משתנה (ידני)", '=IF(OR(p_float_weight=0,$N$69=""),"",$N$69)', '=IF(OR(p_float_weight=0,$N$69=""),"",p_float_weight)')]
    for i, (a, kf, wf) in enumerate(comps):
        r = 69 + i
        put(ws, f"A{r}", a, BASE, None, None, BOX); put(ws, f"B{r}", kf, BASE, None, "@", BOX)
        kk = f'MATCH(TEXT(B{r},"0"),b_key,0)'
        put(ws, f"C{r}", f'=IF(B{r}="","",IFERROR(INDEX(b_name,{kk}),"לא נמצא"))', BASE, None, None, BOX); put(ws, f"D{r}", f'=IF(B{r}="","",IFERROR(INDEX(b_kind,{kk}),""))', BASE, None, None, BOX)
        put(ws, f"E{r}", wf, BASE, GRN, "0.0%", BOX)
        for col, nm, fmt in (("F", "b_mod", "0.000"), ("G", "b_ytm", "0.00"), ("H", "b_conv", "0.0"), ("I", "b_ttm", "0.00"), ("K", "b_nomeq", "0.00")):
            put(ws, f"{col}{r}", f'=IF(B{r}="","",IFERROR(INDEX({nm},{kk}),""))', BASE, None, fmt, BOX)
        put(ws, f"J{r}", f'=IF(B{r}="","",IF(D{r}="מדד",1,0))', GREY, None, "0", BOX)
        _num_helpers(ws, r)
    MP = {c: f"${c}$69:${c}$74" for c in "BEFGHIJKLMNOPQR"}
    # --- תיק סימולציה ---
    put(ws, "A77", "תיק סימולציה (ידני: מפתח ומשקל)", BOLD)
    hdr(ws, 78, ["#", "מפתח", "שם", "אפיק", "משקל", "מח״מ", "תשואה %", "קמירות", "שנים לפדיון", "צמוד?", "שקול 12 ח׳ %"])
    simp = ctx.prev.get("sim_portfolio", [])
    for i in range(10):
        r = 79 + i; k, w = (simp[i] if i < len(simp) else (None, None))
        put(ws, f"A{r}", i + 1, GREY, None, "0"); put(ws, f"B{r}", k, BLUE, INP, "@", BOX); put(ws, f"E{r}", w, BLUE, INP, "0.0%", BOX)
        kk = f'MATCH(TEXT(B{r},"0"),b_key,0)'
        put(ws, f"C{r}", f'=IF(B{r}="","",IFERROR(INDEX(b_name,{kk}),"לא נמצא"))', BASE, None, None, BOX); put(ws, f"D{r}", f'=IF(B{r}="","",IFERROR(INDEX(b_kind,{kk}),""))', BASE, None, None, BOX)
        for col, nm, fmt in (("F", "b_mod", "0.000"), ("G", "b_ytm", "0.00"), ("H", "b_conv", "0.0"), ("I", "b_ttm", "0.00"), ("K", "b_nomeq", "0.00")):
            put(ws, f"{col}{r}", f'=IF(B{r}="","",IFERROR(INDEX({nm},{kk}),""))', BASE, None, fmt, BOX)
        put(ws, f"J{r}", f'=IF(B{r}="","",IF(D{r}="מדד",1,0))', GREY, None, "0", BOX)
        _num_helpers(ws, r)
    SP = {c: f"${c}$79:${c}$88" for c in "BEFGHIJKLMNOPQR"}
    ctx.name("simp_key", "'09_תיק_ותרחישים'!$B$79:$B$88"); ctx.name("simp_w", "'09_תיק_ותרחישים'!$E$79:$E$88")
    # --- סיכום 3 תיקים ---
    put(ws, "A91", "סיכום — תיק נוכחי / המודל / סימולציה", BOLD)
    hdr(ws, 92, ["מדד", "נוכחי", "המודל", "סימולציה", "הסבר"])
    def sp(T, expr):  # SUMPRODUCT על טבלה קטנה עם N() לטיפול בריקים
        return expr
    small = {"w": lambda T: f"SUM({T['L']})", "wd": lambda T: f"SUMPRODUCT({T['L']},{T['M']})", "wy": lambda T: f"SUMPRODUCT({T['L']},{T['N']})",
             "wc": lambda T: f"SUMPRODUCT({T['L']},{T['O']})", "wdt": lambda T: f"SUMPRODUCT({T['L']},{T['M']},{T['P']})", "wdreal": lambda T: f"SUMPRODUCT({T['L']},{T['M']},{T['Q']})",
             "wreal": lambda T: f"SUMPRODUCT({T['L']},{T['Q']})", "wne": lambda T: f"SUMPRODUCT({T['L']},{T['R']})"}
    rows = [("שווי שוק", "=SUM(b_mv)", "", "", "לפי המחיר ליום הנתונים", "#,##0"),
            ("סכום משקלים", "=SUM(b_w)", "=" + small["w"](MP), "=" + small["w"](SP), "אמור להיות 100%", "0.0%"),
            ("מח״מ", "=SUM(b_wdur)", "=" + small["wd"](MP), "=" + small["wd"](SP), "משוקלל", "0.00"),
            ("% צמוד", "=SUM(b_wreal)", "=" + small["wreal"](MP), "=" + small["wreal"](SP), "", "0.0%"),
            ("תשואה %", "=SUM(b_wytm)", "=" + small["wy"](MP), "=" + small["wy"](SP), "ברוטו, משוקללת", "0.00"),
            ("שקול נומינלי 12 ח׳ %", "=SUM(b_wnomeq)", "=" + small["wne"](MP), "=" + small["wne"](SP), "צמודים לפי תוואי המדדים", "0.00"),
            ("קמירות", "=SUM(b_wconv)", "=" + small["wc"](MP), "=" + small["wc"](SP), "", "0.0"),
            ("w·D·ttm (עזר)", "=SUM(b_wdt)", "=" + small["wdt"](MP), "=" + small["wdt"](SP), "לתרחישי תלילות", "0.000"),
            ("w·D צמוד (עזר)", "=SUM(b_wdreal)", "=" + small["wdreal"](MP), "=" + small["wdreal"](SP), "לתרחישי BEI/ריאלי", "0.000")]
    for i, (a, b, c, d, e, f) in enumerate(rows):
        r = 93 + i
        put(ws, f"A{r}", a, BOLD if i < 7 else GREY, None, None, BOX)
        for col, v in (("B", b), ("C", c), ("D", d)): put(ws, f"{col}{r}", v if v else "", BASE if i < 7 else GREY, None, f, BOX)
        put(ws, f"E{r}", e, NOTE)
    ctx.name("cur_mv", "'09_תיק_ותרחישים'!$B$93"); ctx.name("cur_dur", "'09_תיק_ותרחישים'!$B$95"); ctx.name("cur_w_linked", "'09_תיק_ותרחישים'!$B$96"); ctx.name("cur_ytm", "'09_תיק_ותרחישים'!$B$97")
    ctx.name("mod_dur", "'09_תיק_ותרחישים'!$C$95"); ctx.name("mod_w_linked", "'09_תיק_ותרחישים'!$C$96"); ctx.name("sim_dur", "'09_תיק_ותרחישים'!$D$95")
    # --- תרחישים ---
    put(ws, "A104", "תרחישים — P&L מיידי (% מהתיק); מקבילי עם קמירות, תלילות סביב מח״מ התיק (לפי מח״מ בלבד), BEI = שקלי/ריאלי בכיוונים מנוגדים, אינפלציה על הצמודים בלבד", BOLD)
    hdr(ws, 105, ["תרחיש", "גודל", "נוכחי %", "המודל %", "סימולציה %", "מפתח"])
    # עזר: שורות הסיכום לכל תיק
    S = {"current": "B", "model": "C", "sim": "D"}; W, WD, WC, WDT, WDR, WR = 94, 95, 99, 100, 101, 96
    def par(col, sign): return f'=IF({col}{W}=0,"",{sign}*(-{col}{WD}*$B{{r}}/10000)*100+0.5*{col}{WC}*($B{{r}}/10000)^2*100)'
    def twist(col, sign): return f'=IF({col}{W}=0,"",-{sign}*p_bb_twist_bp/p_bb_twist_span/100*({col}{WDT}-{col}{WD}*{col}{WD}/{col}{W}))'    # pivot = מח״מ התיק/סכום המשקלים
    def bei(col, sign): return f'=IF({col}{W}=0,"",-({col}{WD}-{col}{WDR})*({sign}*$B{{r}}/2)/100+{col}{WDR}*({sign}*$B{{r}}/2)/100)'
    def real(col, sign): return f'=IF({col}{W}=0,"",-{col}{WDR}*{sign}*$B{{r}}/100)'
    def infl(col, sign): return f'=IF({col}{W}=0,"",{sign}*$B{{r}}*{col}{WR})'
    scen = [("parallel_up_1", "מקבילי +", "=p_scen_parallel_1", par, 1), ("parallel_dn_1", "מקבילי −", "=p_scen_parallel_1", par, -1),
            ("parallel_up_2", "מקבילי +", "=p_scen_parallel_2", par, 1), ("parallel_dn_2", "מקבילי −", "=p_scen_parallel_2", par, -1),
            ("steep", "התללה", "=p_bb_twist_bp", twist, 1), ("flat", "השתטחות", "=p_bb_twist_bp", twist, -1),
            ("bei_up", "BEI +", "=p_scen_bei_bp", bei, 1), ("bei_dn", "BEI −", "=p_scen_bei_bp", bei, -1),
            ("real_up", "ריאלי +", "=p_scen_real_bp", real, 1), ("real_dn", "ריאלי −", "=p_scen_real_bp", real, -1),
            ("infl_up", "אינפלציה בפועל +", "=p_scen_infl_pct", infl, 1), ("infl_dn", "אינפלציה בפועל −", "=p_scen_infl_pct", infl, -1)]
    for i, (key, lab, size, fn, sign) in enumerate(scen):
        r = 106 + i
        put(ws, f"A{r}", lab, BASE, None, None, BOX); put(ws, f"B{r}", size, GREEN, None, "0", BOX); put(ws, f"F{r}", key, GREY)
        for j, port in enumerate(("current", "model", "sim")):
            put(ws, f"{L(3 + j)}{r}", fn(S[port], sign).replace("{r}", str(r)), BASE, None, "0.00", BOX)
            ctx.name(f"scen_{key}_{port}", f"'09_תיק_ותרחישים'!${L(3 + j)}${r}")
    # --- מדד ייחוס ---
    put(ws, "A120", "מול מדד ייחוס (ידני ב-01)", BOLD); hdr(ws, 121, ["מדד", "מדד הייחוס", "התיק הנוכחי", "פוזיציה אקטיבית"])
    put(ws, "A122", "מח״מ", BASE, None, None, BOX); put(ws, "B122", '=IF(p_bench_dur="","לא הוזן",p_bench_dur)', GREEN, None, "0.00", BOX); put(ws, "C122", "=cur_dur", GREEN, None, "0.00", BOX)
    put(ws, "D122", '=IF(p_bench_dur="","",cur_dur-p_bench_dur)', BOLD, None, "+0.00;-0.00", BOX)
    put(ws, "A123", "% צמוד", BASE, None, None, BOX); put(ws, "B123", '=IF(p_bench_linked="","לא הוזן",p_bench_linked)', GREEN, None, "0%", BOX); put(ws, "C123", "=cur_w_linked", GREEN, None, "0%", BOX)
    put(ws, "D123", '=IF(p_bench_linked="","",cur_w_linked-p_bench_linked)', BOLD, None, "+0%;-0%", BOX)
    ctx.name("bench_dur_diff", "'09_תיק_ותרחישים'!$D$122")
    for col, w in zip("ABCDEFGHIJK", (26, 16, 24, 12, 10, 10, 10, 9, 10, 7, 12)): ws.column_dimensions[col].width = w
    ws.column_dimensions.group("L", "R", outline_level=1, hidden=True)
    ws.freeze_panes = "A5"


# ---------------------------------------------------------------- 10_היסטוריה, 11_בנק_ישראל
def sheet_history(ctx: Ctx):
    ws = ctx.sheet("10_היסטוריה"); h = ctx.frames["נתוני_היסטוריה"]
    put(ws, "A1", "היסטוריה — העקום ב-1/2/5/10 שנים: גלובס יומי (מ-13.10.2026) + נקודות רבעוניות (מגמה, 1/2023–10/2026); היום מול ההיסטוריה", TITLE)
    put(ws, "A3", "אחוזון = חלק ההיסטוריה שמתחת לערך היום. הנתונים בגיליון נתוני_היסטוריה (Power Query).", NOTE)
    hdr(ws, 5, ["סדרה", "היום", "מינימום", "ממוצע", "מקסימום", "אחוזון היום", "תצפיות"])
    r = 6
    if len(h.columns):
        D = lambda c: dcol(ctx, "נתוני_היסטוריה", c)
        items = [("BEI 1Y", "bei_1", "INDEX(cur_bei,MATCH(1,tenors,0))"), ("BEI 2Y", "bei_2", "INDEX(cur_bei,MATCH(2,tenors,0))"), ("BEI 5Y", "bei_5", "INDEX(cur_bei,MATCH(5,tenors,0))"), ("BEI 10Y", "bei_10", "INDEX(cur_bei,MATCH(10,tenors,0))"),
                 ("ריאלי 2Y", "real_2", "INDEX(cur_real,MATCH(2,tenors,0))"), ("ריאלי 5Y", "real_5", "INDEX(cur_real,MATCH(5,tenors,0))"), ("ריאלי 10Y", "real_10", "INDEX(cur_real,MATCH(10,tenors,0))"),
                 ("שקלי 2Y", "nom_2", "INDEX(cur_nom,MATCH(2,tenors,0))"), ("שקלי 5Y", "nom_5", "INDEX(cur_nom,MATCH(5,tenors,0))"), ("שקלי 10Y", "nom_10", "INDEX(cur_nom,MATCH(10,tenors,0))")]
        for lab, col, today in items:
            put(ws, f"A{r}", lab, BOLD, None, None, BOX); put(ws, f"B{r}", f"={today}", GREEN, None, "0.00", BOX)
            put(ws, f"C{r}", f'=IFERROR(MIN({D(col)}),"")', BASE, None, "0.00", BOX); put(ws, f"D{r}", f'=IFERROR(AVERAGE({D(col)}),"")', BASE, None, "0.00", BOX); put(ws, f"E{r}", f'=IFERROR(MAX({D(col)}),"")', BASE, None, "0.00", BOX)
            put(ws, f"F{r}", f'=IF(COUNT({D(col)})<5,"",COUNTIF({D(col)},"<"&B{r})/COUNT({D(col)})*100)', BASE, None, "0", BOX); put(ws, f"G{r}", f"=COUNT({D(col)})", BASE, None, "0", BOX)
            r += 1
        put(ws, f"A{r}", "שיפוע שקלי 10Y−2Y (נ״ב)", BOLD, None, None, BOX); put(ws, f"B{r}", "=(INDEX(cur_nom,MATCH(10,tenors,0))-INDEX(cur_nom,MATCH(2,tenors,0)))*100", GREEN, None, "0", BOX)
        arr(ws, f"F{r}", f'=IF(COUNT({D("nom_10")})<5,"",SUMPRODUCT(((N(+{D("nom_10")})-N(+{D("nom_2")}))*100<B{r})*ISNUMBER({D("nom_10")}))/COUNT({D("nom_10")})*100)')
        ws[f"F{r}"].number_format = "0"; ws[f"F{r}"].border = BOX; put(ws, f"G{r}", f"=COUNT({D('nom_10')})", BASE, None, "0", BOX)
        ctx.name("slope_pct", f"'10_היסטוריה'!$F${r}")
        n = len(h)
        wsd = ctx.wb["נתוני_היסטוריה"]; cols = list(h.columns)
        def chart(title, names, anchor):
            lc = LineChart(); lc.title = title; lc.height = 9; lc.width = 22
            for nm in names:
                ci = cols.index(nm) + 1
                lc.add_data(Reference(wsd, min_col=ci, min_row=1, max_row=n + 1), titles_from_data=True)
            lc.set_categories(Reference(wsd, min_col=cols.index("date") + 1, min_row=2, max_row=n + 1)); ws.add_chart(lc, anchor)
        chart("BEI 2Y / 5Y / 10Y", ["bei_2", "bei_5", "bei_10"], "A20"); chart("שקלי 2Y / 10Y וריאלי 2Y / 10Y", ["nom_2", "nom_10", "real_2", "real_10"], "A40")
    else:
        put(ws, "A6", "אין היסטוריה במאגר", NOTE); ctx.name("slope_pct", "'10_היסטוריה'!$F$6")
    for col, w in zip("ABCDEFG", (24, 10, 10, 10, 10, 12, 9)): ws.column_dimensions[col].width = w


def sheet_boi(ctx: Ctx):
    ws = ctx.sheet("11_בנק_ישראל")
    put(ws, "A1", "בנק ישראל — עקומי אפס מול העקום שלנו, ציפיות אינפלציה, ריבית ותחזית חטיבת המחקר מול הפורוורדים", TITLE)
    put(ws, "A3", "עקומי האפס (SDMX ZCM, ממוצע חודשי) מול עקום התשואות לפדיון שלנו — הפרש סביר של 10–30 נ״ב (שיטה ותאריך).", NOTE)
    hdr(ws, 5, ["טווח", "תקופה", "אפס שקלי %", "אפס ריאלי %", "BEI אפס %", "שלנו שקלי %", "שלנו ריאלי %", "שלנו BEI %", "הפרש שקלי (נ״ב)", "הפרש ריאלי (נ״ב)"])
    bc = ctx.frames["נתוני_בי_עקומים"]; r = 6
    if len(bc.columns):
        T, P, N_, R_, B_ = (dcol(ctx, "נתוני_בי_עקומים", c) for c in ("tenor", "period", "nominal", "real", "bei"))
        for t in (1, 2, 3, 4, 5, 7, 10, 15, 20):
            put(ws, f"A{r}", t, BASE, None, "0", BOX)
            for col, src, fmt in (("B", P, None), ("C", N_, "0.00"), ("D", R_, "0.00"), ("E", B_, "0.00")):
                put(ws, f"{col}{r}", f'=IFERROR(INDEX({src},MATCH(A{r},{T},0)),"")', GREY, None, fmt, BOX)
            put(ws, f"F{r}", f'=IFERROR(INDEX(cur_nom,MATCH(A{r},tenors,0)),"")', GREEN, None, "0.00", BOX); put(ws, f"G{r}", f'=IFERROR(INDEX(cur_real,MATCH(A{r},tenors,0)),"")', GREEN, None, "0.00", BOX)
            put(ws, f"H{r}", f'=IFERROR(INDEX(cur_bei,MATCH(A{r},tenors,0)),"")', GREEN, None, "0.00", BOX)
            put(ws, f"I{r}", f'=IF(OR(C{r}="",F{r}=""),"",(F{r}-C{r})*100)', BASE, None, "0", BOX); put(ws, f"J{r}", f'=IF(OR(D{r}="",G{r}=""),"",(G{r}-D{r})*100)', BASE, None, "0", BOX)
            r += 1
    r += 1; put(ws, f"A{r}", "ציפיות אינפלציה — בנק ישראל (ערך אחרון)", BOLD); r += 1; hdr(ws, r, ["סדרה", "תקופה", "ערך %"]); r += 1
    be = ctx.frames["נתוני_בי_ציפיות"]
    if len(be.columns):
        for j in range(len(be)):
            for k, col in enumerate(("name", "period", "value")):
                put(ws, f"{L(k + 1)}{r}", f"='נתוני_בי_ציפיות'!{L(list(be.columns).index(col) + 1)}{j + 2}", GREY, None, "0.00" if col == "value" else None, BOX)
            r += 1
    r += 1; put(ws, f"A{r}", "ריבית בנק ישראל ותחזית חטיבת המחקר (ידני עם מקור — gov/boi_policy_manual.csv) מול הפורוורדים שלנו", BOLD); r += 1
    hdr(ws, r, ["פריט", "ערך", "נכון ל", "מקור"]); r += 1
    pol = ctx.frames["נתוני_בי_ריבית"]
    if len(pol.columns):
        I_, V_, A_, S_ = (dcol(ctx, "נתוני_בי_ריבית", c) for c in ("item", "value", "as_of", "source"))
        for item, heb in (("rate_current", "ריבית ב״י נוכחית %"), ("next_decision", "ההחלטה הבאה"), ("forecast_end_next_year", "תחזית ריבית לסוף השנה הבאה %"), ("forecast_end_year_after", "תחזית ריבית לסוף השנה שאחריה %"), ("researcher_date", "מועד התחזית"), ("il_1y", "שקלי שנה שלנו %")):
            put(ws, f"A{r}", heb, BOLD, None, None, BOX)
            put(ws, f"B{r}", f'=IFERROR(INDEX({V_},MATCH("{item}",{I_},0)),"")', GREY, None, None, BOX); put(ws, f"C{r}", f'=IFERROR(INDEX({A_},MATCH("{item}",{I_},0)),"")', GREY, None, "dd/mm/yyyy", BOX)
            put(ws, f"D{r}", f'=IFERROR(INDEX({S_},MATCH("{item}",{I_},0)),"")', GREY, None, None, BOX); r += 1
    r += 1; put(ws, f"A{r}", "פורוורד שנה-בעוד-שנה %", BOLD, None, None, BOX); put(ws, f"B{r}", "=INDEX(cur_fwd_n,MATCH(2,tenors,0))", GREEN, None, "0.00", BOX)
    put(ws, f"A{r + 1}", "פורוורד שנה-בעוד-שנתיים %", BOLD, None, None, BOX); put(ws, f"B{r + 1}", "=INDEX(cur_fwd_n,MATCH(3,tenors,0))", GREEN, None, "0.00", BOX)
    put(ws, f"A{r + 2}", "אות מדיניות (נ״ב): פורוורד 1y1y − תחזית ב״י", BOLD, None, None, BOX); put(ws, f"B{r + 2}", '=IF(policy_gap_bp="","אין תחזית",policy_gap_bp)', BOLD, None, "0", BOX)
    put(ws, f"C{r + 2}", '=IF(policy_gap_bp="","",IF(policy_gap_bp<0,"השוק מגלם ריבית נמוכה מתחזית ב״י (הפחתות מעבר לב״י)","השוק מגלם ריבית גבוהה מתחזית ב״י"))', NOTE)
    for col, w in zip("ABCDEFGHIJ", (34, 12, 12, 40, 10, 12, 12, 12, 12, 12)): ws.column_dimensions[col].width = w


# ---------------------------------------------------------------- 13_אופטימיזציה
OPT = {"res": 5, "cmp": 27, "front": 48, "sc": 60, "ln": 75, "lr": 90, "cand": 105}     # שורות הכותרת של הבלוקים בגיליון 13
FLOORS = [0, 0.5, 1, 1.5, 2, 2.5, 3]


def sheet_optimize(ctx: Ctx):
    """הפוזיציה המומלצת: סריקת תיקים שלמים (משקל צמוד × רגל שקלית × רגל צמודה) על פני 9 תרחישים עם הסתברויות. מקביל ל-gov_decide.optimize."""
    G = gov_decide; NS, NL, GRID, MIX = G.NSC, G.NLEG, G.GRID, G.MIXES; NC = 5 * NL * NL
    ws = ctx.sheet("13_אופטימיזציה"); ws.sheet_properties.tabColor = "00B050"
    put(ws, "A1", "אופטימיזציה — הפוזיציה הטובה ביותר לתיק ממשלתי טהור, בהתחשב בכל הנתונים", TITLE)
    put(ws, "A2", "כל תיק = משקל צמוד × רגל שקלית × רגל צמודה (טווח בודד או חצי-חצי בין שני טווחים). לכל תיק מחושבת תשואת 12 חודשים בכל תרחיש. נבחר: עומד בתקציב הסיכון → תשואה צפויה מרבית → בין הקרובים (סף אדישות) הבטוח ביותר; התיק של אתמול נשאר אם הוא בתוך הסף.", NOTE)
    put(ws, "A3", "ההסתברויות = הסתברויות הבסיס שלך (01) + תיקוני האותות: תמחור (אחוזון הריאלי הארוך), ישראל–ארה״ב, ריבית (שוק מול בנק ישראל וכלכלנים), סיכון פיסקלי, פיזור התחזיות. כל העוצמות והתרחישים ב-01.", NOTE)
    rS, rN, rR, rC = OPT["sc"], OPT["ln"], OPT["lr"], OPT["cand"]
    c0, c1 = rC + 1, rC + NC                                                        # שורות המועמדים
    scol = [L(7 + i) for i in range(NS)]                                            # G..O — עמודות התרחישים בטבלאות הרגליים והמועמדים
    def crng(col): return f"${col}${c0}:${col}${c1}"
    # ---------- אותות (מימין לטבלת התרחישים)
    sg = ctx.frames["נתוני_אותות"]; has_sig = len(sg.columns) and "val" in sg.columns
    def sig(key):
        if not has_sig: return '""'
        k, v = dcol(ctx, "נתוני_אותות", "key"), dcol(ctx, "נתוני_אותות", "val"); ix = f'INDEX({v},MATCH("{key}",{k},0))'
        return f'IFERROR(IF({ix}="","",{ix}),"")'
    y1 = "INDEX(cur_nom,MATCH(1,tenors,0))"
    hdr(ws, rS, ["אות", "ערך", "עוצמה (−1..1)", "תיקון (נק׳ אחוז)", "פירוש"], 12)
    sig_rows = [("אחוזון הריאלי ל-10 שנים מול ההיסטוריה", "=" + sig("real10_pct"), '=IF(M{r}="","",(M{r}-50)/50)', '=IF(N{r}="",0,N{r}*p_opt_k_val)', "גבוה = תשואה ריאלית גבוהה היסטורית → יותר הסתברות לירידת תשואות (תרחישים 2, 7) על חשבון עלייה (3, 6)"),
                ("אחוזון פער ישראל–ארה״ב 10 שנים", "=" + sig("ilus10_pct"), '=IF(M{r}="","",(M{r}-50)/50)', '=IF(N{r}="",0,N{r}*p_opt_k_us)', "נמוך = ישראל יקרה מול ארה״ב → יותר הסתברות לפרמיה פיסקלית (6) על חשבון ירידת פרמיה (7)"),
                ("ריבית: מגולמת בשוק פחות התחזית (נ״ב)", '=IF(OR(M{a}="",M{b}=""),"",(M{c}-M{b})*100)', '=IF(M{r}="","",MAX(-1,MIN(1,M{r}/50)))', '=IF(N{r}="",0,N{r}*p_opt_k_policy)', "חיובי = השוק מגלם ריבית גבוהה מהתחזית → יותר הסתברות להורדות ריבית (5) על חשבון עלייה מקבילה (3)"),
                ("סיכון פיסקלי/פוליטי (01)", '=IF(p_fiscal_risk="","",p_fiscal_risk)', '=IF(M{r}="","",M{r}-2)', '=IF(N{r}="",0,N{r}*p_opt_k_fiscal)', "3 = גבוה → יותר הסתברות לפרמיה פיסקלית (6) על חשבון ירידת פרמיה (7); 1 = להפך"),
                ("פיזור תחזיות הריבית (נ״ב)", '=IF(OR({mn}="",{mx}=""),"",({mx}-{mn})*100)', '=IF(M{r}="","",MIN(1,M{r}/50))', '=IF(N{r}="",0,N{r}*p_opt_k_disp)', "פיזור רחב → פחות 'ללא שינוי' (1), יותר לשני הכיוונים (3, 5)")]
    ra, rb, rc, ri = rS + 7, rS + 8, rS + 9, rS + 10                               # ריבית נוכחית, תחזית בשימוש, מגולמת, אינפלציית בסיס
    for i, (lab, val, strength, adj, note) in enumerate(sig_rows):
        r = rS + 1 + i; f = dict(r=r, a=ra, b=rb, c=rc, mn=sig("rate_fc_min"), mx=sig("rate_fc_max"))
        put(ws, f"L{r}", lab, BOLD, None, None, BOX); put(ws, f"M{r}", val.format(**f) if "{" in val else val, GREEN, None, "0.0", BOX)
        put(ws, f"N{r}", strength.format(**f), BASE, None, "0.00", BOX); put(ws, f"O{r}", adj.format(**f), BASE, None, "+0.0%;-0.0%;0.0%", BOX); put(ws, f"P{r}", note, NOTE)
    dV, dU, dP, dF, dD = (f"$O${rS + k}" for k in range(1, 6))                      # תיקוני: תמחור, ארה"ב, ריבית, פיסקלי, פיזור
    for r, lab, f, fmt in ((ra, "ריבית בנק ישראל הנוכחית %", "=" + sig("rate_now"), "0.00"), (rb, "תחזית ריבית בעוד שנה — בשימוש %", f'=IF(p_user_rate_fc="",{sig("rate_fc_mean")},p_user_rate_fc)', "0.00"),
                           (rc, "ריבית מגולמת בשוק בעוד שנה %", f'=IF(M{ra}="","",2*{y1}-M{ra})', "0.00"), (ri, "אינפלציית הבסיס לתרחישים %", f'=IF(infl_fc="",INDEX(cur_bei,MATCH(1,tenors,0)),infl_fc)', "0.00")):
        put(ws, f"L{r}", lab, BASE, None, None, BOX); put(ws, f"M{r}", f, GREEN if r != rc else BASE, None, fmt, BOX)
    put(ws, f"P{rb}", "ריק ב-01 = ממוצע חטיבת המחקר ובתי ההשקעות (נתוני_אותות, נתוני_כלכלנים)", NOTE); put(ws, f"P{rc}", "2 × תשואת שנה − הריבית הנוכחית (מסלול ליניארי)", NOTE)
    put(ws, f"P{ri}", "מקור התחזית שנבחר ב-01 (04); אין תחזית → הגלומה לשנה", NOTE)
    for k, (lab, key, fmt) in enumerate((("ריאלי 10 שנים: היום %", "real10_now", "0.00"), ("חציון היסטורי %", "real10_med", "0.00"), ("טווח היסטורי: מינימום %", "real10_min", "0.00"), ("מקסימום %", "real10_max", "0.00"),
                                         ("חודשי היסטוריה בחלון", "hist_months", "0"), ("פער ישראל–ארה״ב 10 שנים (נ״ב)", "ilus10_now", "0"), ("חציון היסטורי של הפער (נ״ב)", "ilus10_med", "0"), ("מספר תחזיות הריבית", "rate_fc_n", "0"))):
        r = rS + 1 + k; put(ws, f"Q{r}", lab, GREY, None, None, BOX); put(ws, f"R{r}", "=" + sig(key), GREY, None, fmt, BOX)
    put(ws, f"Q{rS}", "רקע לאותות (נתוני_אותות)", BOLD)
    infl0 = f"$M${ri}"
    # ---------- תרחישים והסתברויות
    put(ws, f"A{rS - 1}", "תרחישים והסתברויות (הגדרות התרחישים והסתברויות הבסיס — 01, קבוצת 'תרחישי אופטימיזציה')", BOLD)
    hdr(ws, rS, ["תרחיש", "ריאלי קצר — 2 שנים (נ״ב)", "ריאלי בינוני — 5 שנים", "ריאלי ארוך — 10 שנים", "ציפיות אינפלציה (נ״ב)", "אינפלציה בפועל מול התחזית (נק׳)", "הסתברות בסיס", "תיקון האותות", "אחרי רצפה 1%", "הסתברות סופית"])
    adj = {1: f"-{dD}", 2: f"{dV}/2", 3: f"-{dV}/2-{dP}+{dD}/2", 4: "0", 5: f"{dP}+{dD}/2", 6: f"-{dV}/2-{dU}+{dF}", 7: f"{dV}/2+{dU}-{dF}", 8: "0", 9: "0"}
    for i in range(1, NS + 1):
        r = rS + i; put(ws, f"A{r}", f"{i}. {G.SC_NAMES[i - 1]}", BOLD, None, None, BOX)
        for col, key, fmt in (("B", "rs", "0"), ("C", "rb", "0"), ("D", "rl", "0"), ("E", "bei", "0"), ("F", "infl", "0.0"), ("G", None, "0.0%")):
            put(ws, f"{col}{r}", f"=p_sc{i}_{key}" if key else f"=p_prob_{i}", GREEN, None, fmt, BOX)
        put(ws, f"H{r}", f"=IF(p_opt_signals_on=1,{adj[i]},0)", BASE, None, "+0.0%;-0.0%;0.0%", BOX)
        put(ws, f"I{r}", f"=MAX({G.PROB_FLOOR},G{r}+H{r})", BASE, None, "0.0%", BOX); put(ws, f"J{r}", f"=I{r}/SUM($I${rS + 1}:$I${rS + NS})", BOLD, GRN, "0.0%", BOX)
    rT = rS + NS + 1
    put(ws, f"A{rT}", "סה״כ", BOLD); put(ws, f"G{rT}", f"=SUM(G{rS + 1}:G{rS + NS})", BOLD, None, "0.0%"); put(ws, f"J{rT}", f"=SUM(J{rS + 1}:J{rS + NS})", BOLD, None, "0.0%")
    ctx.name("o_probs", f"'13_אופטימיזציה'!$J${rS + 1}:$J${rS + NS}")
    rP = rN - 2                                                                     # ההסתברויות בשורה (לסכימה מול עמודות התרחישים)
    put(ws, f"A{rP}", "הסתברות סופית לפי תרחיש", GREY)
    for i in range(NS): put(ws, f"{scol[i]}{rP}", f"=$J${rS + 1 + i}", GREY, None, "0.0%")
    prow = f"${scol[0]}${rP}:${scol[-1]}${rP}"
    # ---------- טבלאות הרגליים
    def legs(r0: int, kind: str):
        ycol, rcol = ("cur_nom", "cur_roll_n") if kind == "nominal" else ("cur_real", "cur_roll_r")
        put(ws, f"A{r0 - 1}", ("רגליים שקליות" if kind == "nominal" else "רגליים צמודות (כולל אינפלציה בפועל)") + " — תשואת 12 חודשים % לפי תרחיש: תשואה + roll − מח״מ בסוף האופק × שינוי התשואה + קמירות", BOLD)
        hdr(ws, r0, ["רגל (טווח בשנים)", "טווח", "תשואה %", "Roll (נ״ב)", "מח״מ", "מח״מ בסוף האופק"] + [f"תרחיש {i}" for i in range(1, NS + 1)])
        for k, tn in enumerate(GRID):
            r = r0 + 1 + k
            put(ws, f"A{r}", f"{tn:g}", BOLD, None, None, BOX); put(ws, f"B{r}", tn, BASE, None, "0", BOX)
            put(ws, f"C{r}", f"=INDEX({ycol},MATCH(B{r},tenors,0))", GREEN, None, "0.00", BOX); put(ws, f"D{r}", f"=INDEX({rcol},MATCH(B{r},tenors,0))", GREEN, None, "0.0", BOX)
            put(ws, f"E{r}", f"=IF(ABS(C{r})<0.000001,B{r},(1-(1+C{r}/100)^-B{r})/(C{r}/100))", BASE, None, "0.00", BOX)
            put(ws, f"F{r}", f"=IF(B{r}<=1,0,IF(ABS(C{r})<0.000001,B{r}-1,(1-(1+C{r}/100)^-(B{r}-1))/(C{r}/100)))", BASE, None, "0.00", BOX)
            for i in range(1, NS + 1):
                q = rS + i; S_, B_, L_ = f"$B${q}", f"$C${q}", f"$D${q}"
                dy = f"IF($B{r}<=2,{S_},IF($B{r}<=5,{S_}+({B_}-{S_})*($B{r}-2)/3,IF($B{r}<=10,{B_}+({L_}-{B_})*($B{r}-5)/5,{L_})))" + (f"+$E${q}" if kind == "nominal" else "")
                f = f"=$C{r}+$E{r}*$D{r}/100-$F{r}*({dy})/100+0.5*$F{r}*($F{r}+1)*({dy})^2/1000000" + (f"+{infl0}+$F${q}" if kind == "real" else "")
                put(ws, f"{scol[i - 1]}{r}", f, BASE, None, "0.00", BOX)
        for k, (a, b) in enumerate(MIX):
            r = r0 + 1 + len(GRID) + k; ra_, rb_ = r0 + 1 + GRID.index(a), r0 + 1 + GRID.index(b)
            put(ws, f"A{r}", f"{a:g}+{b:g}", BOLD, None, None, BOX); put(ws, f"B{r}", "", BASE, None, None, BOX)
            for col in ["E"] + scol: put(ws, f"{col}{r}", f"=({col}{ra_}+{col}{rb_})/2", BASE, None, "0.00", BOX)
    legs(rN, "nominal"); legs(rR, "real")
    labN, labR = f"$A${rN + 1}:$A${rN + NL}", f"$A${rR + 1}:$A${rR + NL}"
    # ---------- משבצות משקל הצמוד + אתמול
    put(ws, f"Q{rN - 1}", "משקלי הצמוד הנבחנים (עוגן כלל ה-BEI ± הטווח)", BOLD)
    for k in range(5):
        put(ws, f"Q{rN + 1 + k}", f"משבצת {k + 1}", BASE, None, None, BOX); put(ws, f"R{rN + 1 + k}", f"=MAX(p_linked_min,MIN(p_linked_max,r_linked+({k - 2})*p_opt_w_band/2))", BASE, None, "0.0%", BOX)
    dh = ctx.frames["נתוני_החלטות"]; hist_ok = len(dh.columns) and all(c in dh.columns for c in ("linked_weight", "opt_leg_n", "opt_leg_r"))
    def yest(col): return f'=IFERROR(INDEX({dcol(ctx, "נתוני_החלטות", col)},MATCH(d_asof-1,{dcol(ctx, "נתוני_החלטות", "date")},1)),"")' if hist_ok else '=""'
    rY = rN + 8
    put(ws, f"Q{rY}", "התיק של אתמול (נתוני_החלטות)", BOLD)
    for k, (lab, col, fmt) in enumerate((("משקל צמוד", "linked_weight", "0.0%"), ("רגל שקלית (מס׳)", "opt_leg_n", "0"), ("רגל צמודה (מס׳)", "opt_leg_r", "0"))):
        put(ws, f"Q{rY + 1 + k}", lab, BASE, None, None, BOX); put(ws, f"R{rY + 1 + k}", yest(col), GREEN, None, fmt, BOX)
    pw, pn, pr = f"$R${rY + 1}", f"$R${rY + 2}", f"$R${rY + 3}"
    # ---------- המועמדים
    put(ws, f"A{rC - 1}", f"כל התיקים הנבחנים ({NC}): 5 משקלים × {NL} רגליים שקליות × {NL} רגליים צמודות", BOLD)
    hdr(ws, rC, ["#", "משקל צמוד", "רגל שקלית", "רגל צמודה", "מח״מ שקלי", "מח״מ צמוד"] + [f"תרחיש {i}" for i in range(1, NS + 1)] +
        ["תשואה צפויה %", "בתרחיש הגרוע %", "מח״מ מותר", "עומד ברצפה", "צפויה אם עומד", "בתוך הסף", "ציון בחירה", "אתמול"] + [f"רצפה {f:g}%" for f in FLOORS])
    cP, cQ, cR, cS, cT, cU, cV, cW = "P", "Q", "R", "S", "T", "U", "V", "W"; cF = [L(24 + k) for k in range(len(FLOORS))]      # X..AD
    hE, hN, hD, hPrev, hPick = "$T$6", "$T$7", "$T$8", "$T$9", "$T$10"              # עזר: emax, מספר העומדים, מספר במח״מ מותר, אינדקס אתמול, האינדקס הנבחר
    n = 0
    for k in range(5):
        for i in range(1, NL + 1):
            for j in range(1, NL + 1):
                r = c0 + n; n += 1; rn, rr = rN + i, rR + j
                ws.cell(r, 1, n); put(ws, f"B{r}", f"=$R${rN + 1 + k}", BASE, None, "0.0%"); ws.cell(r, 3, i); ws.cell(r, 4, j)
                put(ws, f"E{r}", f"=$E${rn}", BASE, None, "0.00"); put(ws, f"F{r}", f"=$E${rr}", BASE, None, "0.00")
                for col in scol: put(ws, f"{col}{r}", f"=$B{r}*{col}${rr}+(1-$B{r})*{col}${rn}", BASE, None, "0.00")
                put(ws, f"{cP}{r}", f"=SUMPRODUCT({prow},{scol[0]}{r}:{scol[-1]}{r})", BASE, None, "0.00"); put(ws, f"{cQ}{r}", f"=MIN({scol[0]}{r}:{scol[-1]}{r})", BASE, None, "0.00")
                put(ws, f"{cR}{r}", f"=IF(AND(E{r}>=p_dur_min,E{r}<=p_dur_max,F{r}>=p_dur_min,F{r}<=p_dur_max),1,0)", GREY)
                put(ws, f"{cS}{r}", f"=IF(AND({cR}{r}=1,{cQ}{r}>=p_risk_floor_pct),1,0)", GREY); put(ws, f"{cT}{r}", f"=IF({cS}{r}=1,{cP}{r},-999)", GREY, None, "0.00")
                put(ws, f"{cU}{r}", f"=IF(AND({cS}{r}=1,{cP}{r}>={hE}-p_opt_tol_bp/100),1,0)", GREY)
                put(ws, f"{cV}{r}", f"=IF({hN}>0,IF({cU}{r}=1,{cQ}{r},-999),IF({cR}{r}=1,{cQ}{r},-999))", GREY, None, "0.00")
                put(ws, f"{cW}{r}", f'=IF({pw}="",0,IF(AND(ABS(B{r}-{pw})<0.0005,C{r}={pn},D{r}={pr}),1,0))', GREY)
                for kf, col in enumerate(cF): put(ws, f"{col}{r}", f"=IF(AND({cR}{r}=1,{cQ}{r}>={FLOORS[kf]}),{cP}{r},-999)", GREY, None, "0.00")
    # ---------- עזר לבחירה
    put(ws, "S5", "עזר לבחירה", GREY)
    for cell, lab, f in (("6", "תשואה צפויה מרבית בין העומדים ברצפה", f"=MAX({crng(cT)})"), ("7", "מספר התיקים שעומדים ברצפה", f"=SUM({crng(cS)})"), ("8", "מספר התיקים במח״מ מותר", f"=SUM({crng(cR)})"),
                         ("9", "אינדקס התיק של אתמול", f'=IFERROR(MATCH(1,{crng(cW)},0),"")'),
                         ("10", "האינדקס הנבחר", f'=IF({hD}=0,"",IF(AND({hPrev}<>"",{hN}>0,IFERROR(INDEX({crng(cU)},{hPrev}),0)=1),{hPrev},MATCH(MAX({crng(cV)}),{crng(cV)},0)))')):
        put(ws, f"S{cell}", lab, GREY); put(ws, f"T{cell}", f, GREY, None, "0.00")
    pick = hPick
    def at(col): return f'IF({pick}="","",INDEX({crng(col)},{pick}))'
    # ---------- התוצאה
    rr0 = OPT["res"]; hdr(ws, rr0, ["הפוזיציה המומלצת", "ערך", "הסבר"])
    res = [("משקל צמוד", "=" + at("B"), "מתוך המשבצות סביב עוגן כלל ה-BEI (04)", "0%"),
           ("רגל שקלית — טווח בשנים", f'=IF({pick}="","",INDEX({labN},INDEX({crng("C")},{pick})))', "טווח בודד, או חצי-חצי בין שני טווחים (למשל 4+7)", None),
           ("מח״מ שקלי", "=" + at("E"), "מזין את 06 (סינתטי/בולט) לבחירת הסדרות", "0.00"),
           ("רגל צמודה — טווח בשנים", f'=IF({pick}="","",INDEX({labR},INDEX({crng("D")},{pick})))', "", None),
           ("מח״מ צמוד", "=" + at("F"), "", "0.00"),
           ("מח״מ משוקלל", f'=IF(B{rr0 + 1}="","",B{rr0 + 1}*B{rr0 + 5}+(1-B{rr0 + 1})*B{rr0 + 3})', "משקל צמוד × מח״מ צמוד + משקל שקלי × מח״מ שקלי", "0.00"),
           ("תשואה צפויה ל-12 חודשים %", "=" + at(cP), "ממוצע התרחישים, משוקלל בהסתברויות הסופיות", "0.00"),
           ("תשואה בתרחיש הגרוע %", "=" + at(cQ), "מול תקציב הסיכון (01): רצפה של התשואה בתרחיש הגרוע", "0.00"),
           ("סטטוס", f'=IF({hN}>0,"תקין",IF({hD}>0,"אין תיק שעומד בתקציב הסיכון — נבחר הבטוח ביותר","אין תיק במגבלות המח״מ"))', "", None),
           ("נשאר התיק של אתמול?", f'=IF(AND({pick}<>"",{hPrev}<>"",{pick}={hPrev},{hN}>0),"כן","לא")', "התיק של אתמול נשאר כל עוד הוא בתוך סף האדישות מהטוב ביותר", None)]
    for i, (a, b, c, f) in enumerate(res):
        r = rr0 + 1 + i; put(ws, f"A{r}", a, BOLD, None, None, BOX); put(ws, f"B{r}", b, BOLD, GRN if i in (0, 2, 4, 5) else None, f, BOX); put(ws, f"C{r}", c, NOTE)
    rw = rr0 + 11
    put(ws, f"A{rw}", "למה", BOLD)
    put(ws, f"B{rw}", f'=IF({pick}="","אין תיק במגבלות המח״מ",{hN}&" מתוך {NC} תיקים עומדים ברצפה של "&TEXT(p_risk_floor_pct,"0.0#")&"% בתרחיש הגרוע; נבחר: תשואה צפויה "&TEXT(B{rr0 + 7},"0.00")&"%, בתרחיש הגרוע "&TEXT(B{rr0 + 8},"0.00")&"%"&IF(B{rr0 + 10}="כן","; נשאר התיק של אתמול (בתוך סף האדישות)","")&IF(B{rr0 + 9}="תקין","","; "&B{rr0 + 9}))', BASE)
    put(ws, f"A{rw + 1}", "האותות", BOLD)
    put(ws, f"B{rw + 1}", f'=IF(p_opt_signals_on<>1,"האותות כבויים — הסתברויות הבסיס בלבד","תמחור "&TEXT({dV},"+0.0%;-0.0%;0%")&" | ישראל–ארה״ב "&TEXT({dU},"+0.0%;-0.0%;0%")&" | ריבית "&TEXT({dP},"+0.0%;-0.0%;0%")&" | פיסקלי "&TEXT({dF},"+0.0%;-0.0%;0%")&" | פיזור "&TEXT({dD},"+0.0%;-0.0%;0%"))', BASE)
    for nm, row in (("o_linked", 1), ("o_lab_nominal", 2), ("o_dur_nominal", 3), ("o_lab_real", 4), ("o_dur_real", 5), ("o_dur_weighted", 6), ("o_exp", 7), ("o_worst", 8), ("o_status", 9), ("o_kept", 10), ("o_why", 11), ("o_signals", 12)):
        ctx.name(nm, f"'13_אופטימיזציה'!$B${rr0 + row}")
    ctx.name("o_pick", f"'13_אופטימיזציה'!{hPick}")
    # ---------- המודל הסופי (מזין את כל שאר הקובץ)
    rm = rr0 + 14; hdr(ws, rm, ["המודל הסופי — מזין את 05, 06, 09 והדשבורד", "ערך", "הסבר"])
    on = f'AND(p_opt_on=1,B{rr0 + 1}<>"")'
    fin = [("משקל צמוד — המודל", f"=IF({on},B{rr0 + 1},r_linked)", "אופטימיזציה פעילה (01) → התוצאה למעלה; אחרת כלל ה-BEI (04)", "0%", "m_linked"),
           ("מח״מ שקלי — המודל", f"=IF({on},B{rr0 + 3},r_dur_nominal)", "אחרת כלל הציון (05)", "0.00", "m_dur_nominal"),
           ("מח״מ צמוד — המודל", f"=IF({on},B{rr0 + 5},r_dur_real)", "אחרת כלל הציון (05)", "0.00", "m_dur_real"),
           ("למה — משקל", f'=IF({on},"אופטימיזציה: "&B{rw}&" | עוגן: "&r_linked_why,r_linked_why)', "", None, "m_linked_why"),
           ("למה — מח״מ שקלי", f'=IF({on},"אופטימיזציה: רגל "&B{rr0 + 2}&" שנים (כלל הציון: "&IF(r_dur_nominal="","—",TEXT(r_dur_nominal,"0.0"))&")",r_dur_nominal_why)', "", None, "m_dur_nominal_why"),
           ("למה — מח״מ צמוד", f'=IF({on},"אופטימיזציה: רגל "&B{rr0 + 4}&" שנים (כלל הציון: "&IF(r_dur_real="","—",TEXT(r_dur_real,"0.0"))&")",r_dur_real_why)', "", None, "m_dur_real_why")]
    for i, (a, b, c, f, nm) in enumerate(fin):
        r = rm + 1 + i; put(ws, f"A{r}", a, BOLD if i < 3 else BASE, None, None, BOX); put(ws, f"B{r}", b, BOLD if i < 3 else BASE, GRN if i < 3 else None, f, BOX); put(ws, f"C{r}", c, NOTE)
        ctx.name(nm, f"'13_אופטימיזציה'!$B${r}")
    # ---------- השוואת תיקים לפי תרחיש
    rc0 = OPT["cmp"]; put(ws, f"A{rc0 - 1}", "תשואת 12 חודשים לפי תרחיש: התיק המומלץ, תיק הכללים הנפרדים, והתיק בשימוש (הסימולציה שלך ב-01 אם מולאה)", BOLD)
    hdr(ws, rc0, ["תרחיש", "הסתברות", "התיק המומלץ %", "תיק הכללים %", "התיק בשימוש %"])
    durN, durR = f"$E${rN + 1}:$E${rN + len(GRID)}", f"$E${rR + 1}:$E${rR + len(GRID)}"            # רגליים בודדות בלבד — אינטרפולציה לפי מח״מ
    rp = rc0 + NS + 3                                                                # שורות הפרמטרים של התיקים
    ports = {"D": ("r_linked", "r_dur_nominal", "r_dur_real"), "E": ("u_linked", "u_dur_nominal", "u_dur_real")}
    for col, (w_, dn_, dr_) in ports.items():
        put(ws, f"{col}{rp}", f'=IF({w_}="","",{w_})', GREEN, None, "0%", BOX); put(ws, f"{col}{rp + 1}", f'=IF({dn_}="","",{dn_})', GREEN, None, "0.00", BOX); put(ws, f"{col}{rp + 2}", f'=IF({dr_}="","",{dr_})', GREEN, None, "0.00", BOX)
        for k, (drow, rng) in enumerate(((rp + 1, durN), (rp + 2, durR))):           # k=0 שקלי, k=1 צמוד: אינדקס תחתון ושבר
            put(ws, f"{col}{rp + 3 + 2 * k}", f'=IF({col}{drow}="","",MAX(1,MIN({len(GRID) - 1},IFERROR(MATCH({col}{drow},{rng},1),1))))', GREY, None, "0")
            put(ws, f"{col}{rp + 4 + 2 * k}", f'=IF({col}{drow}="","",MAX(0,MIN(1,({col}{drow}-INDEX({rng},{col}{rp + 3 + 2 * k}))/(INDEX({rng},{col}{rp + 3 + 2 * k}+1)-INDEX({rng},{col}{rp + 3 + 2 * k})))))', GREY, None, "0.000")
    put(ws, f"C{rp}", "=o_linked", GREEN, None, "0%", BOX); put(ws, f"C{rp + 1}", "=o_dur_nominal", GREEN, None, "0.00", BOX); put(ws, f"C{rp + 2}", "=o_dur_real", GREEN, None, "0.00", BOX)
    for k, lab in enumerate(("משקל צמוד", "מח״מ שקלי", "מח״מ צמוד", "עזר: אינדקס שקלי", "עזר: שבר שקלי", "עזר: אינדקס צמוד", "עזר: שבר צמוד")):
        put(ws, f"A{rp + k}", lab, BOLD if k < 3 else GREY, None, None, BOX if k < 3 else None)
    for i in range(1, NS + 1):
        r = rc0 + i; sc = scol[i - 1]
        put(ws, f"A{r}", f"=A{rS + i}", GREEN, None, None, BOX); put(ws, f"B{r}", f"=J{rS + i}", GREEN, None, "0.0%", BOX)
        put(ws, f"C{r}", f'=IF({pick}="","",INDEX({crng(sc)},{pick}))', BASE, None, "0.00", BOX)
        for col in ports:
            def ip(r0, kx):   # ערך הרגל בתרחיש באינטרפולציה: kx=0 שקלי, 1 צמוד
                rng = f"${sc}${r0 + 1}:${sc}${r0 + len(GRID)}"; ix, fr = f"{col}${rp + 3 + 2 * kx}", f"{col}${rp + 4 + 2 * kx}"
                return f"(INDEX({rng},{ix})+(INDEX({rng},{ix}+1)-INDEX({rng},{ix}))*{fr})"
            put(ws, f"{col}{r}", f'=IF(OR({col}${rp}="",{col}${rp + 1}="",{col}${rp + 2}=""),"",{col}${rp}*{ip(rR, 1)}+(1-{col}${rp})*{ip(rN, 0)})', BASE, None, "0.00", BOX)
    r = rc0 + NS + 1; put(ws, f"A{r}", "תשואה צפויה (משוקלל הסתברויות)", BOLD, SUB, None, BOX); put(ws, f"A{r + 1}", "בתרחיש הגרוע", BOLD, SUB, None, BOX)
    for col in "CDE":
        rng = f"{col}{rc0 + 1}:{col}{rc0 + NS}"
        put(ws, f"{col}{r}", f'=IF({col}{rc0 + 1}="","",SUMPRODUCT($B${rc0 + 1}:$B${rc0 + NS},{rng}))', BOLD, SUB, "0.00", BOX); put(ws, f"{col}{r + 1}", f'=IF({col}{rc0 + 1}="","",MIN({rng}))', BOLD, SUB, "0.00", BOX)
    put(ws, f"F{rc0 + 1}", "תיק הכללים והתיק בשימוש מוערכים באינטרפולציה לפי מח״מ על הרגליים הבודדות (העקום); מבנה סינתטי בפועל (06) עשוי לסטות מעט.", NOTE)
    for nm, col in (("o_cmp_exp_rule", "D"), ("o_cmp_exp_use", "E"), ("o_cmp_exp_opt", "C")): ctx.name(nm, f"'13_אופטימיזציה'!${col}${r}")
    ctx.name("o_cmp_worst_use", f"'13_אופטימיזציה'!$E${r + 1}")
    # ---------- חזית: הטוב ביותר לכל רצפה
    rf = OPT["front"]; put(ws, f"A{rf - 1}", "מחיר תקציב הסיכון — התיק עם התשואה הצפויה הגבוהה ביותר לכל רצפה (בלי סף האדישות)", BOLD)
    hdr(ws, rf, ["רצפה: תשואה מינימלית בתרחיש הגרוע %", "משקל צמוד", "רגל שקלית", "רגל צמודה", "מח״מ משוקלל", "תשואה צפויה %", "בתרחיש הגרוע %"])
    for k, fl in enumerate(FLOORS):
        r = rf + 1 + k; ix = f"$I{r}"; col = cF[k]
        put(ws, f"I{r}", f'=IF(MAX({crng(col)})=-999,"",MATCH(MAX({crng(col)}),{crng(col)},0))', GREY)
        def g(c): return f'=IF({ix}="","—",INDEX({crng(c)},{ix}))'
        put(ws, f"A{r}", fl, BOLD, None, "0.0", BOX); put(ws, f"B{r}", g("B"), BASE, None, "0%", BOX)
        put(ws, f"C{r}", f'=IF({ix}="","—",INDEX({labN},INDEX({crng("C")},{ix})))', BASE, None, None, BOX); put(ws, f"D{r}", f'=IF({ix}="","—",INDEX({labR},INDEX({crng("D")},{ix})))', BASE, None, None, BOX)
        put(ws, f"E{r}", f'=IF({ix}="","—",INDEX({crng("B")},{ix})*INDEX({crng("F")},{ix})+(1-INDEX({crng("B")},{ix}))*INDEX({crng("E")},{ix}))', BASE, None, "0.00", BOX)
        put(ws, f"F{r}", g(cP), BASE, None, "0.00", BOX); put(ws, f"G{r}", g(cQ), BASE, None, "0.00", BOX)
    for col, w in zip("ABCDEFGHIJKLMNOP", (40, 14, 16, 16, 16, 16, 12, 12, 12, 12, 4, 40, 12, 12, 12, 60)): ws.column_dimensions[col].width = w
    ws.column_dimensions["Q"].width = 30; ws.column_dimensions["S"].width = 34
    ws.freeze_panes = "A5"


# ---------------------------------------------------------------- 00_דשבורד, 12_מדריך
def sheet_dashboard(ctx: Ctx):
    ws = ctx.sheet("00_דשבורד")
    put(ws, "A1", "קובץ הממשלתי — דשבורד החלטה (תיק ממשלתי טהור)", TITLE)
    put(ws, "A2", '="נתוני יום המסחר: "&TEXT(d_asof,"dd/mm/yyyy")&"   |   הנתונים מתרעננים בפתיחת הקובץ (Power Query מ-C:\\CORP\\03_Market\\data)   |   גרסה ' + VERSION + '"', NOTE)
    put(ws, "A3", '=IF(n_not_found>0,"אזהרה: "&n_not_found&" החזקות לא נמצאו בנתוני היום (09). ","")&warn_list_nominal&" "&warn_list_real', Font(name="Arial", size=10, bold=True, color="C00000"))
    hdr(ws, 5, ["שאלה", "פסק דין — המודל", "למה", "אתמול (המודל עם ההגדרות שלך, מהחישוב הלילי)", "מול התיק הנוכחי"], 1, 30)
    hist_ok = len(ctx.frames["נתוני_החלטות"].columns) > 0
    def y(col):
        if not hist_ok or col not in ctx.frames["נתוני_החלטות"].columns: return '""'
        return f'IFERROR(INDEX({dcol(ctx, "נתוני_החלטות", col)},MATCH(d_asof-1,{dcol(ctx, "נתוני_החלטות", "date")},1)),"")'
    rows = [("הפוזיציה המומלצת", '=IF(m_dur_weighted="","",TEXT(m_linked,"0%")&" צמוד במח״מ "&TEXT(m_dur_real,"0.0")&" + "&TEXT(1-m_linked,"0%")&" שקלי במח״מ "&TEXT(m_dur_nominal,"0.0")&" | משוקלל "&TEXT(m_dur_weighted,"0.0")&IF(o_exp="",""," | צפוי "&TEXT(o_exp,"0.0")&"% לשנה, בתרחיש הגרוע "&TEXT(o_worst,"0.0")&"%"))',
             '=IF(p_opt_on=1,"אופטימיזציה (13): "&o_why&" | אותות: "&o_signals,"האופטימיזציה כבויה (01) — לפי הכללים הנפרדים")',
             f'=IF({y("opt_exp")}="","","צפוי "&TEXT({y("opt_exp")},"0.0")&"% | גרוע "&TEXT({y("opt_worst")},"0.0")&"%")',
             '=IF(cur_mv=0,"אין תיק","התיק הנוכחי: "&TEXT(cur_w_linked,"0%")&" צמוד, מח״מ "&TEXT(cur_dur,"0.0"))'),
            ("שקלי / צמוד", '=IF(m_linked="","",TEXT(m_linked,"0%")&" צמוד / "&TEXT(1-m_linked,"0%")&" שקלי")', "=m_linked_why", f'=IF({y("linked_weight")}="","",TEXT({y("linked_weight")},"0%")&" צמוד")',
             '=IF(cur_mv=0,"אין תיק",TEXT(cur_w_linked,"0%")&" צמוד בתיק → "&IF(cur_w_linked<m_linked-0.05,"להגדיל צמוד",IF(cur_w_linked>m_linked+0.05,"להקטין צמוד","קרוב למודל")))'),
            ("מח״מ", '=IF(m_dur_weighted="","",,"שקלי "&TEXT(m_dur_nominal,"0.0")&" | צמוד "&TEXT(m_dur_real,"0.0")&" | משוקלל "&TEXT(m_dur_weighted,"0.0"))', '=m_dur_nominal_why&" | צמוד: "&m_dur_real_why',
             f'=IF({y("dur_weighted")}="","","משוקלל "&TEXT({y("dur_weighted")},"0.0"))', '=IF(cur_mv=0,"אין תיק","מח״מ תיק "&TEXT(cur_dur,"0.0")&" → "&IF(cur_dur<m_dur_weighted-0.25,"להאריך",IF(cur_dur>m_dur_weighted+0.25,"לקצר","קרוב למודל")))'),
            ("סינתטי / בולט", '="שקלי: "&m_struct_nominal&" | צמוד: "&m_struct_real', '=m_struct_nominal_why&" | צמוד: "&m_struct_real_why',
             f'=IF({y("struct_nominal")}="","","שקלי: "&{y("struct_nominal")}&" | צמוד: "&{y("struct_real")})', '="תיק המודל בגיליון 09"'),
            ("סדרות מובילות", '="שקלי: "&TEXTJOIN_N&" | צמוד: "&TEXTJOIN_R', '="ציון = Z/3×0.4 + roll×0.3 + נזילות×0.3; זוגות החלפה ב-07"', f'=IF({y("top5_nominal")}="","",{y("top5_nominal")})', '="לבדוק מול ההחזקות ב-02 (פער מול העקום, Z)"'),
            ("חו״ל", '="5Y: "&f_verdict_5y&" | 10Y: "&f_verdict_10y', '="5Y: "&f_why_5y&" | 10Y: "&f_why_10y', f'=IF({y("foreign_5y")}="","","5Y: "&{y("foreign_5y")}&" | 10Y: "&{y("foreign_10y")})', '="לפי פיקאפ מגודר ≥ "&p_fx_pickup_bp&" נ״ב ב-"&p_fx_min_markets&" שווקים"'),
            ("סיכון", '="P&L במקבילי +"&p_scen_parallel_2&" נ״ב: תיק נוכחי "&IF(scen_parallel_up_2_current="","—",TEXT(scen_parallel_up_2_current,"0.00")&"%")&" | המודל "&IF(scen_parallel_up_2_model="","—",TEXT(scen_parallel_up_2_model,"0.00")&"%")',
             '="תרחישים מלאים (מקבילי, תלילות, BEI, ריאלי, אינפלציה) בגיליון 09"', '=""', '=IF(bench_dur_diff="","אין מדד ייחוס","פוזיציה אקטיבית במח״מ מול מדד הייחוס: "&TEXT(bench_dur_diff,"+0.00;-0.00"))')]
    # TEXTJOIN אינו זמין בכל גרסה → שרשור ידני של 5 המובילות
    def join5(rng):
        return "&".join([f'IF(INDEX({rng},{i})="","",INDEX({rng},{i})&IF({i}<5,"; ",""))' for i in range(1, 6)])
    for i, (a, b, c, d, e) in enumerate(rows):
        r = 6 + i
        b = b.replace("TEXTJOIN_N", join5("top5_nominal")).replace("TEXTJOIN_R", join5("top5_real")).replace(",,", ",")
        put(ws, f"A{r}", a, BOLD, SUB, None, BOX); put(ws, f"B{r}", b, BOLD, GRN, None, BOX); put(ws, f"C{r}", c, BASE, None, None, BOX); put(ws, f"D{r}", d, BASE, None, None, BOX); put(ws, f"E{r}", e, BASE, None, None, BOX)
        for col in "BCDE": ws[f"{col}{r}"].alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 48
    put(ws, "A15", "התיק — נוכחי / המודל / סימולציה", BOLD); hdr(ws, 16, ["מדד", "נוכחי", "המודל", "סימולציה"])
    for i, (lab, src_r, fmt) in enumerate((("מח״מ", 95, "0.00"), ("% צמוד", 96, "0%"), ("תשואה %", 97, "0.00"), ("שקול נומינלי 12 ח׳ %", 98, "0.00"), ("קמירות", 99, "0.0"), ("P&L מקבילי +", 108, "0.00"), ("P&L מקבילי −", 109, "0.00"))):
        r = 17 + i
        put(ws, f"A{r}", lab if i < 5 else lab + " (תרחיש ב׳) %", BOLD, None, None, BOX)
        for j, col in enumerate("BCD"):
            put(ws, f"{col}{r}", f"='09_תיק_ותרחישים'!{col if i < 5 else L(3 + j)}{src_r}", GREEN, None, fmt, BOX)
    put(ws, "A26", "צ׳ק ליסט לוועדה", BOLD); hdr(ws, 27, ["שאלה", "איפה במודל", "כלל"])
    for i, (a, b, c) in enumerate([("הפוזיציה כולה", "13: תיקים שלמים על פני 9 תרחישים עם הסתברויות; תקציב סיכון; אותות תמחור, ריבית, פיסקלי וכלכלנים", "התיק עם התשואה הצפויה הגבוהה ביותר שעומד בתקציב הסיכון; בין קרובים — הבטוח ביותר"),
                                    ("שקלי מול צמוד", "04: פער תחזית−BEI לפי טווח, אחוזון BEI, 12 חודשים לפי תוואי המדדים", "צמוד כשהתחזית גבוהה מה-BEI מעבר למרווח ההחלטה; הטיה פרופורציונית, בתוך גבולות"),
                                    ("מח״מ", "05: תשואה 12 ח׳, breakeven, סיכון, אות מדיניות", "להאריך רק כשהעודף מול שנה מצדיק את הסיכון (ציון), ולא נגד תחזית ב״י"),
                                    ("סינתטי מול בולט", "06: סריקת כל הזוגות ב-5 תרחישים", "סינתטי רק אם היתרון גדול מהחיכוך ונזיל"),
                                    ("סדרות", "07: Z מול היסטוריה, roll-down, נזילות, החלפות", "לקנות זול מול ההיסטוריה שלו באותו מח״מ"),
                                    ("חו״ל", "08: פיקאפ מגודר לפי טווח", "חו״ל מגודר רק כשהפיקאפ עובר סף בכמה שווקים"),
                                    ("סיכון", "09: תרחישים על 3 תיקים; מול מדד ייחוס", "לא לאשר החלטת תשואה בלי לראות את ההפסד בתרחיש עליית תשואות")]):
        r = 28 + i; put(ws, f"A{r}", a, BOLD, None, None, BOX); put(ws, f"B{r}", b, BASE, None, None, BOX); put(ws, f"C{r}", c, BASE, None, None, BOX)
    for col, w in zip("ABCDE", (18, 44, 60, 26, 40)): ws.column_dimensions[col].width = w
    ws.freeze_panes = "A6"


def sheet_guide(ctx: Ctx):
    ws = ctx.sheet("12_מדריך")
    put(ws, "A1", "מדריך — קובץ הממשלתי v2 (מקור הממשלתי של מערכת CORP)", TITLE)
    lines = ["מה הקובץ עונה: כמה צמוד וכמה שקלי (04); איזה מח״מ בשקלי, בצמוד ומשוקלל (05); סינתטי או בולט (06); אילו סדרות (07); חו״ל מול מקומי (08); תיק המודל, תרחישים ומדד ייחוס (09). הכול מסוכם ב-00_דשבורד.",
             "", "שלבי עבודה:",
             "1. פתח את הקובץ — גיליוני נתוני_* מתרעננים אוטומטית מ-C:\\CORP\\03_Market\\data\\gov_*.csv (Power Query; אם אקסל שואל — אפשר תוכן חיצוני). תאריך הנתונים מופיע בראש הדשבורד.",
             "2. 01_הגדרות — כל ספי ההחלטה (תאים צהובים), מקור תחזית האינפלציה, תחזית לפי טווח ומשקלי ההחלטה, מדד ייחוס, תחזיות שער. כל כלל במודל מפנה לתא כאן.",
             "3. 09_תיק_ותרחישים — הזן החזקות (מפתח = מס׳ ני״ע, כמות נקובה) ו-'כלול בעקום?' = לא לניירות שלא רוצים בעקום.",
             "4. סימולציה — בקבוצת 'סימולציה' ב-01: משקל צמוד, מח״מ שקלי/צמוד, ניירות קצר/ארוך/בולט. ריק = לפי המודל. 'בשימוש' = הסימולציה אם מולאה, אחרת המודל — והוא מזין את תיק המודל והתרחישים.",
             "5. הדשבורד: פסק דין, 'למה', הערך של אתמול (מהחישוב הלילי בפייתון) ומה לעשות מול התיק הנוכחי.",
             "", "הכללים במילים:"]
    for x in ctx.settings:
        if x["group"] in ("סימולציה",): continue
        lines.append(f"• {x['label']} = {x['value']}" + (f" — {x['note']}" if x["note"] else ""))
    lines += ["", "מה אוטומטי ומה ידני: אוטומטי — גלובס (ממשלתי כולל מק״מ ומשתנה), היסטוריית הסדרות, תנודתיות, תוואי המדדים, ב״י (עקומי אפס, ציפיות), FRED/ECB/BoE, החלטות המודל (אתמול). ידני עם מקור — תחזית ריבית חטיבת המחקר (gov/boi_policy_manual.csv), תחזיות שער, מדד ייחוס, תחזית אינפלציה של מנהל ההשקעות.",
              "מגבלות: הגידור לפי פער ריביות (לא פורוורד אמיתי); תנודתיות = רצפה מההגדרות עד שיש 60 ימי מסחר במאגר (ואז יומי); אחוזוני BEI/ריאלי מול היסטוריה מ-2023 (רבעוני) + יומי מ-13.10.2026.",
              "אם Power Query לא עובד (גרסה ישנה / מחשב אחר): הקובץ מציג את הערכים מהבנייה האחרונה (תאריך הנתונים בראש); בנייה מחדש: python gov/build_gov2.py data <פלט> --settings <הקובץ הנוכחי> (הקלטים הידניים עוברים).",
              "פייתון מקביל: gov/gov_decide.py מחשב את אותם כללים כל לילה (data/gov/decisions_*.csv) — זה המקור ל-'אתמול' ולבדיקת הזהות (gov/check_gov2.py)."]
    for i, t in enumerate(lines):
        put(ws, f"A{3 + i}", t, BOLD if t.endswith(":") else BASE)
    ws.column_dimensions["A"].width = 160


def names_stub(ctx: Ctx):
    if "infl_12m" not in ctx.names: ctx.name("infl_12m", "p_infl_user_1y")


# ---------------------------------------------------------------- carry-over
def carry_over(prev_path: Path | str | None) -> dict:
    """קורא קלטים ידניים מקובץ קודם (v1 או v2). מחזיר {params, infl_user, dw, holdings, exclude, notes}."""
    out = {"params": {}, "infl_user": {}, "dw": {}, "holdings": {}, "exclude": {}, "notes": {}, "sim_portfolio": []}
    if not prev_path or not Path(prev_path).exists(): return out
    wb = openpyxl.load_workbook(prev_path, data_only=True)
    if "01_הגדרות" in wb.sheetnames:
        ws = wb["01_הגדרות"]; is_v2 = ws["D4"].value == "מפתח"
        if is_v2:
            def same(a, b):
                if a in (None, "") and b in (None, ""): return True
                try: return abs(float(a) - float(b)) < 1e-12
                except (TypeError, ValueError): return str(a) == str(b)
            valid = {x["key"] for x in json.load(open(HERE / "settings_default.json", encoding="utf-8"))}
            for r in range(5, ws.max_row + 1):
                k = ws.cell(r, 4).value; v = ws.cell(r, 2).value; d = ws.cell(r, 5).value
                if not (isinstance(k, str) and k in valid): continue
                if ws.cell(4, 5).value is None:          # קובץ v2 ישן בלי עמודת ברירת מחדל — מעבירים הכול
                    if v not in (None, ""): out["params"][k] = v
                elif not same(v, d): out["params"][k] = "" if v is None else v   # רק מה שהמשתמש שינה; "" = המשתמש ניקה
            # בלוק תחזית לפי טווח: שורת כותרת "טווח בשנים"
            for r in range(5, ws.max_row + 1):
                if ws.cell(r, 1).value == "טווח בשנים":
                    for i in range(12):
                        t = ws.cell(r + 1 + i, 1).value
                        if isinstance(t, (int, float)):
                            if isinstance(ws.cell(r + 1 + i, 2).value, (int, float)): out["infl_user"][float(t)] = ws.cell(r + 1 + i, 2).value
                            if isinstance(ws.cell(r + 1 + i, 3).value, (int, float)): out["dw"][float(t)] = ws.cell(r + 1 + i, 3).value
                    break
        else:  # v1: תוויות בעמודה A, תחזית לפי טווח ב-D/E/F
            lab = {ws.cell(r, 1).value: ws.cell(r, 2).value for r in range(4, 20)}
            mp = {"משקל צמוד מינימום": "linked_min", "משקל צמוד מקסימום": "linked_max"}
            for a, k in mp.items():
                if isinstance(lab.get(a), (int, float)): out["params"][k] = lab[a]
            if isinstance(lab.get("מרווח החלטה מול BEI"), (int, float)): out["params"]["decision_margin_bp"] = round(lab["מרווח החלטה מול BEI"] * 10000)
            for i in range(12):
                t, e, w = ws.cell(5 + i, 4).value, ws.cell(5 + i, 5).value, ws.cell(5 + i, 6).value
                if isinstance(t, (int, float)):
                    if isinstance(e, (int, float)): out["infl_user"][float(t)] = round(e * 100, 3)      # v1 בעשרוני
                    if isinstance(w, (int, float)): out["dw"][float(t)] = w
    if "02_קלט_אגח" in wb.sheetnames and ws["D4"].value != "מפתח":   # v1: החזקות/כלול/הערות לפי שם → מפתח
        p2 = wb["02_קלט_אגח"]; hm = {p2.cell(5, c).value: c for c in range(1, p2.max_column + 1)}
        for r in range(6, p2.max_row + 1):
            sid, nm = p2.cell(r, 2).value, p2.cell(r, 3).value
            if not nm: continue
            key = str(int(sid)) if isinstance(sid, (int, float)) else str(nm).strip()
            c = hm.get("כמות נקובה מוחזקת")
            if c and isinstance(p2.cell(r, c).value, (int, float)) and p2.cell(r, c).value: out["holdings"][key] = p2.cell(r, c).value
            c = hm.get("כלול בעקום?")
            if c and p2.cell(r, c).value == "לא": out["exclude"][key] = "לא"
            c = hm.get("הערות")
            if c and p2.cell(r, c).value not in (None, ""): out["notes"][key] = p2.cell(r, c).value
    if "09_תיק_ותרחישים" in wb.sheetnames:
        p9 = wb["09_תיק_ותרחישים"]; out["sim_portfolio"] = []
        for r in range(79, 89):
            k, w = p9.cell(r, 2).value, p9.cell(r, 5).value
            if k not in (None, "") and isinstance(w, (int, float)): out["sim_portfolio"].append((str(k).strip(), w))
        for r in range(5, 65):
            k, a = p9.cell(r, 1).value, p9.cell(r, 2).value
            if k not in (None, "") and isinstance(a, (int, float)) and a: out["holdings"][str(k).strip()] = a
            ek, ev = p9.cell(r, 6).value, p9.cell(r, 7).value
            if ek not in (None, "") and ev == "לא": out["exclude"][str(ek).strip()] = "לא"
    return out


# ---------------------------------------------------------------- בנייה
def build(data: Path | str, out: Path | str, prev: Path | str | None = None, base_dir: str = r"C:\CORP\03_Market\data", pq: bool = True, write_version: bool = False) -> dict:
    ctx = Ctx(Path(data), carry_over(prev), base_dir)
    for sheet, q, f, nmax in DATA_SHEETS: sheet_data(ctx, sheet, q, ctx.frames[sheet], nmax)
    sheet_settings(ctx); sheet_input(ctx); sheet_curves(ctx); sheet_portfolio(ctx); sheet_split(ctx); sheet_duration(ctx); sheet_structure(ctx); sheet_series(ctx)
    sheet_foreign(ctx); sheet_history(ctx); sheet_boi(ctx); sheet_optimize(ctx); sheet_dashboard(ctx); sheet_guide(ctx); names_stub(ctx)
    ctx.wb._sheets = [ctx.wb[t] for t in ORDER]
    ctx.wb.active = 0
    out = Path(out); tmp = out.with_suffix(".tmp.xlsx") if pq else out
    ctx.wb.save(tmp)
    n_formulas = sum(1 for ws in ctx.wb.worksheets for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("="))
    if pq:
        import build_live
        targets = {q: sheet for sheet, q, f, _ in DATA_SHEETS}
        schemas = pq_schema(ctx)
        section = pq_gov.section_m(base_dir, {q: f for _, q, f, _ in DATA_SHEETS}, schemas)
        mashup, props, conns = pq_gov.build_mashup(section, targets, HERE.parent / "corp" / "pq_kit")
        build_live.inject_pq(str(tmp), str(out), mashup, props, conns, queries=pq_gov.connection_ids(targets))
        tmp.unlink()
    if write_version:
        (out.parent / "_build").mkdir(exist_ok=True); (out.parent / "_build" / "VERSION").write_text(VERSION)
    return {"sheets": ctx.wb.sheetnames, "formulas": n_formulas, "date": ctx.date.date().isoformat(), "version": VERSION}


if __name__ == "__main__":
    a = sys.argv[1:]
    data, out = a[0], a[1]
    prev = a[a.index("--settings") + 1] if "--settings" in a else None
    base = a[a.index("--base") + 1] if "--base" in a else r"C:\CORP\03_Market\data"
    res = build(data, out, prev, base, pq="--no-pq" not in a, write_version="--version" in a)
    print(f"נבנה {out}: {len(res['sheets'])} גיליונות, {res['formulas']:,} נוסחאות, נתוני {res['date']}, גרסה {res['version']}")
