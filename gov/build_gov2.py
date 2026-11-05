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
                    ("cur_vol_n", "M"), ("cur_vol_r", "N"), ("cur_vol_bei", "O"), ("cur_pct_bei", "T")]:
        ctx.name(nm, Q03(col))
    lc = LineChart(); lc.title = "עקום שקלי וריאלי"; lc.y_axis.title = "%"; lc.x_axis.title = "שנים"; lc.height = 8; lc.width = 18
    lc.add_data(Reference(ws, min_col=2, min_row=5, max_row=17), titles_from_data=True); lc.add_data(Reference(ws, min_col=3, min_row=5, max_row=17), titles_from_data=True)
    lc.set_categories(Reference(ws, min_col=1, min_row=6, max_row=17)); ws.add_chart(lc, "A21")
    lb = LineChart(); lb.title = "BEI לפי טווח"; lb.height = 8; lb.width = 18
    lb.add_data(Reference(ws, min_col=4, min_row=5, max_row=17), titles_from_data=True); lb.set_categories(Reference(ws, min_col=1, min_row=6, max_row=17)); ws.add_chart(lb, "H21")


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
    sheet_settings(ctx); sheet_input(ctx); sheet_curves(ctx); manual_tables_stub(ctx); names_stub(ctx)
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
