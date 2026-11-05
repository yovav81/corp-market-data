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
import pq_gov  # noqa: E402

VERSION = (HERE / "VERSION").read_text().strip() if (HERE / "VERSION").exists() else "2.0.0"
TENORS = [0.25, 0.5, 1, 2, 3, 4, 5, 7, 10, 15, 20, 30]
NB = 120; R0 = 6; R1 = R0 + NB - 1                       # שורות הניירות ב-02
GREEN = Font(name="Arial", size=10, color="008000"); GREY = Font(name="Arial", size=9, color="808080"); GFILL = PatternFill("solid", fgColor="F2F2F2")
DATA_FILL = PatternFill("solid", fgColor="EDEDED")
SOURCES = ["תחזית מנהל השקעות", "ממוצע החזאים (ב״י)", "ציפיות שוק ההון (ב״י)", "חוזי אינפלציה (ב״י)", "שקול 12 ח׳ (תוואי מדדים)"]
GROUP_ORDER = ["כללי", "שקלי/צמוד", "מח״מ", "סינתטי/בולט", "סדרות", "חו״ל", "תיק ותרחישים", "סימולציה"]
# גיליוני נתונים: שם גיליון → (שאילתה, קובץ data/gov, מקס' שורות לטווחי הנוסחאות)
DATA_SHEETS = [("נתוני_אגח", "GovBonds", "bonds_latest.csv", 500), ("נתוני_עקומים", "GovCurves", "curves_latest.csv", 20),
               ("נתוני_היסטוריה", "GovHist", "curves_history.csv", 3000), ("נתוני_סדרות_היסט", "GovBondHist", "bond_history.csv", 20000),
               ("נתוני_מדדים", "GovCpi", "cpi_path.csv", 60), ("נתוני_תנודתיות", "GovVol", "vol.csv", 20), ("נתוני_חול", "GovForeign", "foreign_latest.csv", 60),
               ("נתוני_ארהב", "GovUS", "us_latest.csv", 20), ("נתוני_בי_עקומים", "GovBoiCurves", "boi_curves.csv", 60),
               ("נתוני_בי_ציפיות", "GovBoiExp", "boi_expectations.csv", 60), ("נתוני_בי_ריבית", "GovPolicy", "boi_policy.csv", 20),
               ("נתוני_החלטות", "GovDecisions", "decisions_history.csv", 2000)]
ORDER = ["00_דשבורד", "01_הגדרות", "02_קלט_אגח", "03_עקומים", "04_שקלי_צמוד", "05_מחמ", "06_סינטטי_בולט", "07_סדרות", "08_חול", "09_תיק_ותרחישים",
         "10_היסטוריה", "11_בנק_ישראל", "12_מדריך"] + [d[0] for d in DATA_SHEETS]
BOND_COLS = ["date", "security_id", "key", "name", "kind", "source", "maturity", "coupon", "last_price", "adjusted_value", "ytm_gross", "ytm_net", "duration_globes",
             "volume_k_nis", "ttm", "accrued", "clean_price", "macaulay_calc", "modified", "convexity", "L", "pv_check", "in_curve"]
BOND_HEB = ["תאריך נתונים", "מס׳ ני״ע", "מפתח", "שם אג״ח", "אפיק", "טבלת גלובס", "תאריך פדיון", "קופון שנתי %", "מחיר", "ערך מתואם", "תשואה ברוטו %", "תשואה נטו %",
            "מח״מ גלובס", "מחזור (אלפי ₪)", "שנים לפדיון", "ריבית צבורה", "מחיר נקי", "מח״מ Macaulay", "מח״מ מתואם", "קמירות", "מקדם הצמדה L", "מחיר תיאורטי", "כלול בעקום? (מאגר)"]
CALC_COLS = [("incl", "כלול (אפקטיבי)"), ("bucket", "Bucket"), ("lo", "lo"), ("frac", "frac"), ("ycurve", "YTM עקום %"), ("spread", "פער מול העקום (נ״ב)"),
             ("bei", "BEI בטווח %"), ("roll", "Roll-down (נ״ב)"), ("exp12", "תשואה צפויה 12 ח׳ %"), ("nomeq", "שקול נומינלי 12 ח׳ %"), ("hold", "כמות נקובה מוחזקת"),
             ("mv", "שווי שוק"), ("w", "משקל בתיק"), ("wdur", "תרומת מח״מ"), ("wytm", "תרומת תשואה"), ("wconv", "תרומת קמירות"), ("xn", "x שקלי"), ("xr", "x צמוד"), ("status", "סטטוס")]
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
            self.frames[sheet] = pd.read_csv(p, encoding="utf-8-sig") if p.exists() else pd.DataFrame()
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
    hdr(ws, 4, ["פרמטר", "ערך", "הסבר", "מפתח"]); r = 5
    prev = ctx.prev.get("params", {})
    by_group = {}
    for x in ctx.settings: by_group.setdefault(x["group"], []).append(x)
    for grp in GROUP_ORDER:
        if grp not in by_group: continue
        c = ws.cell(r, 1, grp); c.font = BOLD; c.fill = SUB
        for j in range(2, 5): ws.cell(r, j).fill = SUB
        r += 1
        for x in by_group[grp]:
            v = prev.get(x["key"], x["value"])
            if v is None or (isinstance(v, float) and np.isnan(v)): v = ""
            put(ws, f"A{r}", x["label"], BOLD); put(ws, f"B{r}", v, BLUE, INP, x.get("fmt") or None, BOX); put(ws, f"C{r}", x["note"], NOTE); put(ws, f"D{r}", x["key"], GREY)
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
    for col, w in zip("ABCD", (40, 22, 80, 22)): ws.column_dimensions[col].width = w
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
        put(ws, C["exp12"], f'=IF(OR({C["roll"]}="",{mod}=""),"",{ytm}*p_horizon_years+{mod}*{C["roll"]}/100)', BASE, None, "0.00", BOX)
        put(ws, C["nomeq"], f'=IF({C["exp12"]}="","",IF({kind}="מדד",((1+{ytm}/100)*(1+infl_12m/100)-1)*100+{mod}*{C["roll"]}/100,{C["exp12"]}))', BASE, None, "0.00", BOX)
        put(ws, C["hold"], f'=IF({key}="","",IFERROR(INDEX(h_amt,MATCH({key},h_key,0)),0))', BASE, None, "#,##0", BOX)
        put(ws, C["mv"], f'=IFERROR({C["hold"]}*{price}/100,0)', BASE, None, "#,##0", BOX)
        put(ws, C["w"], f'=IFERROR({C["mv"]}/SUM({Q02(CC["mv"])}),0)', BASE, None, "0.0%", BOX)
        put(ws, C["wdur"], f'=IFERROR({C["w"]}*{mod},0)', BASE, None, "0.000", BOX)
        put(ws, C["wytm"], f'=IFERROR({C["w"]}*{ytm},0)', BASE, None, "0.000", BOX)
        put(ws, C["wconv"], f'=IFERROR({C["w"]}*{conv},0)', BASE, None, "0.00", BOX)
        put(ws, C["xn"], f'=IF(AND({C["incl"]}="כן",{isnom},ISNUMBER({ttm}),ISNUMBER({ytm})),{ttm},"")', GREY, None, "0.00")
        put(ws, C["xr"], f'=IF(AND({C["incl"]}="כן",{kind}="מדד",ISNUMBER({ttm}),ISNUMBER({ytm})),{ttm},"")', GREY, None, "0.00")
        put(ws, C["status"], f'=IF({key}="","",IF({ytm}="","אין תשואה",IF({mod}="","אין מח״מ","")))', BASE, None, None, BOX)
    widths = [10, 10, 10, 20, 7, 12, 11, 8, 9, 9, 9, 9, 8, 10, 8, 8, 8, 8, 8, 8, 8, 8, 8] + [8, 7, 4, 5, 9, 9, 8, 9, 10, 10, 12, 11, 8, 8, 8, 8, 6, 6, 10]
    for i, w in enumerate(widths): ws.column_dimensions[L(i + 1)].width = w
    ws.column_dimensions.group(CC["lo"], CC["frac"], outline_level=1, hidden=True); ws.column_dimensions.group(CC["xn"], CC["xr"], outline_level=1, hidden=True)
    ws.freeze_panes = f"E{R0}"; ws.auto_filter.ref = f"A5:{CC['status']}{R1}"
    for nm, col in [("b_key", BC["key"]), ("b_name", BC["name"]), ("b_kind", BC["kind"]), ("b_ttm", BC["ttm"]), ("b_ytm", BC["ytm_gross"]), ("b_mod", BC["modified"]),
                    ("b_conv", BC["convexity"]), ("b_vol", BC["volume_k_nis"]), ("b_price", BC["last_price"]), ("b_sid", BC["security_id"]), ("b_incl", CC["incl"]),
                    ("b_ycurve", CC["ycurve"]), ("b_spread", CC["spread"]), ("b_roll", CC["roll"]), ("b_exp12", CC["exp12"]), ("b_nomeq", CC["nomeq"]), ("b_hold", CC["hold"]),
                    ("b_mv", CC["mv"]), ("b_w", CC["w"]), ("b_xn", CC["xn"]), ("b_xr", CC["xr"]), ("b_bucket", CC["bucket"])]:
        ctx.name(nm, Q02(col))
    ctx.name("d_asof", f"MAX('נתוני_אגח'!$A$2:$A$500)")


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
        put(ws, f"AA{r}", f"=MAX($A{r}-p_horizon_years,MIN($A$6:$A$17))", GREY, None, "0.00")
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
             ("משקל צמוד — המודל", "=MAX(p_linked_min,MIN(p_linked_max,p_linked_neutral+B48+B49))", "נייטרלי + הטיה + תיקון, בתוך המינימום/מקסימום", "0%"),
             ("משקל צמוד — סימולציה", '=IF(p_sim_linked="","",p_sim_linked)', "ריק = לפי המודל (01 → סימולציה)", "0%"),
             ("משקל צמוד בשימוש", '=IF(p_sim_linked="",B50,p_sim_linked)', "מזין את המח״מ המשוקלל ואת תיק המודל", "0%")]
    for i, (a, b, c, f) in enumerate(items):
        r = 48 + i; put(ws, f"A{r}", a, BOLD if i in (2, 4) else BASE, None, None, BOX); put(ws, f"B{r}", b, BOLD if i in (2, 4) else BASE, GRN if i == 2 else None, f, BOX); put(ws, f"C{r}", c, NOTE)
    ctx.name("m_linked", "'04_שקלי_צמוד'!$B$50"); ctx.name("u_linked", "'04_שקלי_צמוד'!$B$52")
    put(ws, "A53", "למה", BOLD)
    put(ws, "B53", '=IF(gap_bp="","אין תחזית אינפלציה — נייטרלי","מקור "&p_infl_source&" "&TEXT(infl_fc,"0.00")&"%; פער תחזית−BEI "&TEXT(gap_bp,"+0;-0")&" נ״ב"&IF(ABS(gap_bp)<=p_decision_margin_bp," (בתוך מרווח ההחלטה → נייטרלי)"," → הטיה "&TEXT(B48,"+0%;-0%"))&IF(B49<>0,"; אחוזון BEI "&TEXT(bei_pct2,"0")&" → "&TEXT(B49,"+0%;-0%"),"")&" → משקל צמוד "&TEXT(B50,"0%"))', BASE)
    ctx.name("m_linked_why", "'04_שקלי_צמוד'!$B$53")
    # 12 חודשים: צמוד מול שקלי
    put(ws, "A55", "12 חודשים קדימה — צמוד (ריאלי + תוואי המדדים + roll) מול שקלי באותו טווח", BOLD)
    hdr(ws, 56, ["טווח", "שקלי: תשואה 12 ח׳ %", "צמוד: שקול נומינלי 12 ח׳ %", "יתרון צמוד (נ״ב)"])
    for i, t in enumerate((1, 2, 5)):
        r = 57 + i; k = TENORS.index(t) + 1
        d_n = f"(1-(1+INDEX(cur_nom,{k})/100)^-INDEX(tenors,{k}))/(INDEX(cur_nom,{k})/100)"; d_r = f"(1-(1+INDEX(cur_real,{k})/100)^-INDEX(tenors,{k}))/(INDEX(cur_real,{k})/100)"
        put(ws, f"A{r}", t, BASE, None, "0", BOX)
        put(ws, f"B{r}", f'=IFERROR(INDEX(cur_nom,{k})*p_horizon_years+{d_n}*INDEX(cur_roll_n,{k})/100,"")', BASE, None, "0.00", BOX)
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
        put(ws, f"E{r}", f'=IF(D{r}="","",B{r}*p_horizon_years+D{r}*C{r}/100)', BASE, None, "0.00", BOX)
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
    hist_ok = len(ctx.frames["נתוני_החלטות"].columns) and hcol in ctx.frames["נתוני_החלטות"].columns
    prev_f = f'IFERROR(INDEX({dcol(ctx, "נתוני_החלטות", hcol)},MATCH(d_asof-1,{dcol(ctx, "נתוני_החלטות", "date")},1)),"")' if hist_ok else '""'
    hdr(ws, r, ["רכיב", "ערך", "הסבר"])
    rows = [("הטווח עם הציון המקסימלי", f'=IF(MAX({T["sc_ok"]})=-9999,"",INDEX({T["tenor"]},MATCH(MAX({T["sc_ok"]}),{T["sc_ok"]},0)))', "בין הטווחים המותרים (מח״מ מינימום/מקסימום, עודף מינימלי)", "0.00"),
            ("טווח ההמלצה הקודמת (אתמול)", "=" + prev_f, "מנתוני_החלטות (הפייתון הלילי)", "0.00"),
            ("אחרי סף יציבות", f'=IF(B{r + 1}="","",IF(B{r + 2}="",B{r + 1},IF(IFERROR(INDEX({T["ok"]},MATCH(B{r + 2},{T["tenor"]},0)),"לא")<>"כן",B{r + 1},IF(INDEX({T["score"]},MATCH(B{r + 1},{T["tenor"]},0))<INDEX({T["score"]},MATCH(B{r + 2},{T["tenor"]},0))*(1+p_dur_stability),B{r + 2},B{r + 1}))))',
             "ההמלצה מתחלפת רק אם הציון החדש גבוה בסף היציבות מציון ההמלצה הקודמת", "0.00"),
            ("אות מדיניות (נ״ב)", "=policy_gap_bp", "פורוורד שנה-בעוד-שנה מול תחזית ריבית ב״י; שלילי = השוק מגלם הפחתות מעבר לב״י", "0"),
            ("טווח מומלץ — המודל", f'=IF(B{r + 3}="","",IF(AND(p_policy_adj_on=1,B{r + 4}<>"",N(B{r + 4})<=-p_policy_gap_bp),IF(COUNTIFS({T["tenor"]},"<"&B{r + 3},{T["ok"]},"כן")>0,MAX(IF(({T["tenor"]}<B{r + 3})*({T["ok"]}="כן"),{T["tenor"]})),B{r + 3}),B{r + 3}))', "צעד טווח אחד למטה כשהאות המדיניות פעיל", "0.00"),
            (f"מח״מ מומלץ {heb} — המודל", f'=IF(B{r + 5}="","",INDEX({T["dur"]},MATCH(B{r + 5},{T["tenor"]},0)))', "מח״מ (פארי) של הטווח המומלץ", "0.00"),
            (f"מח״מ {heb} — סימולציה", f'=IF({sim}="","",{sim})', "ריק = לפי המודל", "0.00"),
            (f"מח״מ {heb} בשימוש", f'=IF({sim}="",B{r + 6},{sim})', "מזין את סינתטי/בולט ואת תיק המודל", "0.00")]
    for i, (a, b, c, f) in enumerate(rows):
        rr = r + 1 + i
        if i == 4: arr(ws, f"B{rr}", b)
        else: put(ws, f"B{rr}", b, BOLD if i in (5, 7) else BASE, GRN if i == 5 else None, f, BOX)
        ws[f"B{rr}"].number_format = f; ws[f"B{rr}"].border = BOX
        put(ws, f"A{rr}", a, BOLD if i in (5, 7) else BASE, None, None, BOX); put(ws, f"C{rr}", c, NOTE)
    ctx.name(f"m_dur_{kind}_tenor", f"'05_מחמ'!$B${r + 5}"); ctx.name(f"m_dur_{kind}", f"'05_מחמ'!$B${r + 6}"); ctx.name(f"u_dur_{kind}", f"'05_מחמ'!$B${r + 8}")
    put(ws, f"A{r + 9}", "למה", BOLD)
    put(ws, f"B{r + 9}", f'=IF(B{r + 1}="","אין טווח שעומד במגבלות","ציון מקסימלי "&TEXT(INDEX({T["score"]},MATCH(B{r + 1},{T["tenor"]},0)),"0.00")&" בטווח "&TEXT(B{r + 1},"0.##")&IF(B{r + 3}<>B{r + 1},"; נשאר ב-"&TEXT(B{r + 3},"0.##")&" (סף יציבות)","")&IF(B{r + 5}<>B{r + 3},"; אות מדיניות → צעד למטה ל-"&TEXT(B{r + 5},"0.##"),"")&IF(B{r + 4}="","; אין תחזית ריבית ב״י",""))', BASE)
    ctx.name(f"m_dur_{kind}_why", f"'05_מחמ'!$B${r + 9}")
    # סימולציה: מדדים במח״מ הידני (אינטרפולציה על עמודת המח״מ)
    put(ws, f"E{r}", "סימולציה — מדדים במח״מ שבשימוש", BOLD)
    hdr(ws, r + 1, ["", "מח״מ", "תשואה 12 ח׳ %", "Breakeven (נ״ב)", "P&L +תרחיש %", "P&L −תרחיש %"], 5)
    for j, (lab, ref) in enumerate((("המודל", f"B{r + 6}"), ("בשימוש", f"B{r + 8}"))):
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


# ---------------------------------------------------------------- מקומות שמורים לגיליונות הבאים (Tasks 8–10)
def sheet_placeholder(ctx: Ctx, title: str, text: str):
    ws = ctx.sheet(title); put(ws, "A1", text, TITLE); return ws


def manual_tables_stub(ctx: Ctx):
    """טבלאות ידניות (החזקות, כלול בעקום?) — ממוקמות ב-09; עד שהגיליון נבנה (Task 10) נוצרות כאן בבסיסן כדי שהשמות יהיו תקפים."""
    ws = ctx.sheet("09_תיק_ותרחישים")
    put(ws, "A1", "תיק ותרחישים", TITLE)
    hdr(ws, 4, ["מפתח (מס׳ ני״ע)", "כמות נקובה מוחזקת", "שם", "סטטוס", "", "מפתח", "כלול בעקום? (כן/לא)"])
    hold = ctx.prev.get("holdings", {}); excl = ctx.prev.get("exclude", {})
    for i in range(60):
        r = 5 + i; k = list(hold)[i] if i < len(hold) else None; e = list(excl)[i] if i < len(excl) else None
        put(ws, f"A{r}", k, BLUE, INP, "@", BOX); put(ws, f"B{r}", hold.get(k) if k else None, BLUE, INP, "#,##0", BOX)
        put(ws, f"C{r}", f'=IF(A{r}="","",IFERROR(INDEX(b_name,MATCH(TEXT(A{r},"0"),b_key,0)),"לא נמצא"))', BASE, None, None, BOX)
        put(ws, f"D{r}", f'=IF(A{r}="","",IF(C{r}="לא נמצא","לא נמצא",""))', BASE, None, None, BOX)
        put(ws, f"F{r}", e, BLUE, INP, "@", BOX); put(ws, f"G{r}", excl.get(e) if e else None, BLUE, INP, None, BOX)
    ctx.name("h_key", "'09_תיק_ותרחישים'!$A$5:$A$64"); ctx.name("h_amt", "'09_תיק_ותרחישים'!$B$5:$B$64")
    ctx.name("ex_key", "'09_תיק_ותרחישים'!$F$5:$F$64"); ctx.name("ex_val", "'09_תיק_ותרחישים'!$G$5:$G$64")
    ctx.name("n_not_found", 'COUNTIF(\'09_תיק_ותרחישים\'!$D$5:$D$64,"לא נמצא")')
    return ws


def names_stub(ctx: Ctx):
    """שמות שגיליונות מאוחרים (04) מגדירים — ערך זמני עד שייבנו."""
    if "infl_12m" not in ctx.names: ctx.name("infl_12m", "p_infl_user_1y")


# ---------------------------------------------------------------- carry-over
def carry_over(prev_path: Path | str | None) -> dict:
    """קורא קלטים ידניים מקובץ קודם (v1 או v2). מחזיר {params, infl_user, dw, holdings, exclude, notes}."""
    out = {"params": {}, "infl_user": {}, "dw": {}, "holdings": {}, "exclude": {}, "notes": {}}
    if not prev_path or not Path(prev_path).exists(): return out
    wb = openpyxl.load_workbook(prev_path, data_only=True)
    if "01_הגדרות" in wb.sheetnames:
        ws = wb["01_הגדרות"]; is_v2 = ws["D4"].value == "מפתח"
        if is_v2:
            for r in range(5, ws.max_row + 1):
                k = ws.cell(r, 4).value; v = ws.cell(r, 2).value
                if isinstance(k, str) and k and v not in (None, ""): out["params"][k] = v
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
            mp = {"משקל צמוד מינימום": "linked_min", "משקל צמוד מקסימום": "linked_max", "אופק Carry/Roll-down (שנים)": "horizon_years"}
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
        p9 = wb["09_תיק_ותרחישים"]
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
    sheet_settings(ctx); sheet_input(ctx); sheet_curves(ctx); manual_tables_stub(ctx); sheet_split(ctx); sheet_duration(ctx); names_stub(ctx)
    for t in ORDER:
        if t not in ctx.wb.sheetnames: sheet_placeholder(ctx, t, t)
    ctx.wb._sheets = [ctx.wb[t] for t in ORDER]
    ctx.wb.active = 0
    out = Path(out); tmp = out.with_suffix(".tmp.xlsx") if pq else out
    ctx.wb.save(tmp)
    n_formulas = sum(1 for ws in ctx.wb.worksheets for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("="))
    if pq:
        import build_live
        targets = {q: sheet for sheet, q, f, _ in DATA_SHEETS}
        schemas = {q: [(c, "date" if c in DATE_COLS else "number") for c in ctx.frames[sheet].columns if c in DATE_COLS or pd.api.types.is_numeric_dtype(ctx.frames[sheet][c])]
                   for sheet, q, f, _ in DATA_SHEETS if len(ctx.frames[sheet].columns)}
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
