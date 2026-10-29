#!/usr/bin/env python3
"""build_gov.py — בונה את קובץ הממשלתי (ממשלתי.xlsx) מתוך data/gov/*.csv (gov/gov_data.py).
שימוש: python gov/build_gov.py data <פלט.xlsx> [--settings קובץ_קיים.xlsx]
   --settings  מעתיק מגרסה קודמת: פרמטרים (01), החזקות / ״כלול בעקום?״ / הערות (02, לפי שם הנייר), בחירות ב-07.
בענן (המשימה המתוזמנת 23:10) הקובץ נבנה מחדש בכל ערב ל-corp-data/00_Gov/ממשלתי.xlsx עם --settings על עצמו — השינויים של יובב נשמרים.

מבנה (לפי המודל של יובב, בלי מערכת המסחר): 01_הגדרות, 02_קלט_אגח (נתונים מהמאגר שלנו + נוסחאות), 03_עקומים, 04_ציפיות_אינפלציה,
05_אנליטיקה_נייר, 06_תיק_והחלטה, 07_סינטטי_מול_בולט, 08_מדריך + חדשים: 09_היסטוריה, 10_ארהב, 11_בנק_ישראל.
כל הנוסחאות קלאסיות (INDEX/MATCH + נוסחאות מערך) — עובדות בכל גרסת אקסל וב-LibreOffice, ולא תלויות ב-LET/FILTER."""
import datetime as dt, json, os, sys
import pandas as pd, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.chart import LineChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

NAVY = "1F3864"; HDR = PatternFill("solid", fgColor=NAVY); SUB = PatternFill("solid", fgColor="D9E2F3"); INP = PatternFill("solid", fgColor="FFF2CC")
YEL = PatternFill("solid", fgColor="FFFF99"); GRN = PatternFill("solid", fgColor="E2EFDA"); RED = PatternFill("solid", fgColor="F8CBAD")
WH = Font(bold=True, color="FFFFFF", name="Arial", size=10); BASE = Font(name="Arial", size=10); BOLD = Font(name="Arial", size=10, bold=True)
BLUE = Font(name="Arial", size=10, color="0000FF"); NOTE = Font(name="Arial", size=9, italic=True, color="595959"); TITLE = Font(name="Arial", size=14, bold=True)
thin = Side(style="thin", color="BFBFBF"); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
TENORS = [0.25, 0.5, 1, 2, 3, 4, 5, 7, 10, 15, 20, 30]
NB = 120            # שורות קלט (02) — טווח הנוסחאות
R0, R1 = 6, 6 + NB - 1   # שורות הנתונים ב-02


def put(ws, cell, v, font=BASE, fill=None, fmt=None, border=None, align=None):
    c = ws[cell]; c.value = v; c.font = font
    if fill: c.fill = fill
    if fmt: c.number_format = fmt
    if border: c.border = border
    if align: c.alignment = align
    return c


def hdr(ws, row, labels, c0=1, h=30):
    for i, t in enumerate(labels):
        c = ws.cell(row, c0 + i, t); c.fill = HDR; c.font = WH; c.border = BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = h


def rtl(ws): ws.sheet_view.rightToLeft = True


def arr(ws, cell, formula):
    ws[cell] = ArrayFormula(cell, formula)


def interp(x, xs, ys, lo, hi):
    """אינטרפולציה ליניארית: lo/hi תאים עם הנקודה התחתונה/עליונה (−1 = אין תחתונה, 9999 = אין עליונה)."""
    ylo = f"INDEX({ys},MATCH({lo},{xs},0))"; yhi = f"INDEX({ys},MATCH({hi},{xs},0))"
    return f'IFERROR(IF({lo}=-1,{yhi},IF({hi}=9999,{ylo},IF({lo}={hi},{ylo},{ylo}+({x}-{lo})*({yhi}-{ylo})/({hi}-{lo})))),"")'


def lo_formula(x, xs): return f'IF(COUNTIF({xs},"<="&{x})=0,-1,MAX(IF(ISNUMBER({xs})*({xs}<={x}),{xs})))'
def hi_formula(x, xs): return f'IF(COUNTIF({xs},">="&{x})=0,9999,MIN(IF(ISNUMBER({xs})*({xs}>={x}),{xs})))'


def main():
    data = sys.argv[1]; out = sys.argv[2]
    prev = None
    if "--settings" in sys.argv:
        prev = openpyxl.load_workbook(sys.argv[sys.argv.index("--settings") + 1], data_only=True)
    gd = os.path.join(data, "gov")
    bonds = pd.read_csv(os.path.join(gd, "bonds_latest.csv"), encoding="utf-8-sig")
    curves = pd.read_csv(os.path.join(gd, "curves_latest.csv"), encoding="utf-8-sig")
    hist = pd.read_csv(os.path.join(gd, "curves_history.csv"), encoding="utf-8-sig")
    def rd(n):
        try: return pd.read_csv(os.path.join(gd, n), encoding="utf-8-sig")
        except Exception: return pd.DataFrame()
    us, boi_c, boi_e = rd("us_latest.csv"), rd("boi_curves.csv"), rd("boi_expectations.csv")
    try: short_src = json.load(open(os.path.join(data, "short", "cpi_sources.json"), encoding="utf-8"))
    except Exception: short_src = {}
    date = pd.to_datetime(bonds["date"].iloc[0]).to_pydatetime()
    wb = openpyxl.Workbook()

    # ================= 01_הגדרות =================
    ws = wb.active; ws.title = "01_הגדרות"; rtl(ws)
    put(ws, "A1", "מודל ניתוח אג״ח ממשלת ישראל: שקלי, צמוד, אינפלציה גלומה, מח״מ וסינתטי מול בולט", TITLE)
    put(ws, "A2", f"מקור הנתונים: מאגר CORP (גלובס, יום מסחר {date:%d/%m/%Y}; בנק ישראל; FRED). בלי מערכת המסחר — העמודות שהגיעו ממנה מחושבות כאן (ריבית צבורה, מחיר נקי, מח״מ מתואם, קמירות).", NOTE)
    hdr(ws, 3, ["פרמטר", "ערך / הנחה", "הסבר", "טווח בשנים", "צפי אינפלציה שנתי (מנהל השקעות)", "משקל החלטה", "BEI מהעקום (לעיון)"])
    params = [("תאריך ניתוח", date, "תאריך הנתונים. אפשר לדרוס ידנית (משנה שנים לפדיון ו-Bucket; התשואות נשארות של יום הנתונים).", "dd/mm/yyyy"),
              ("אופן קלט תשואות/קופון", "נקודות אחוז 4.5", "הנתונים מהמאגר באחוזים (4.5). לא לשנות.", None),
              ("מח״מ יעד לתיק", 4.5, "יעד אסטרטגי/טקטי של הוועדה.", "0.00"),
              ("אופק Carry/Roll-down (שנים)", 1, "ברירת מחדל: שנה קדימה.", "0.00"),
              ("צפי אינפלציה 12 ח׳ מנהל השקעות", 0.016, "השוואה מול אינפלציה גלומה. ניתן לעדכן ידנית.", "0.00%"),
              ("מרווח החלטה מול BEI", 0.003, "לדוגמה 0.30% כדי להימנע מרעש.", "0.00%"),
              ("עלות עסקה שקלי (bp)", 3, "משמשת בקירוב להשוואת עסקאות.", "0"),
              ("עלות עסקה צמוד (bp)", 5, "משמשת בקירוב להשוואת עסקאות.", "0"),
              ("משקל שקלי מינימום", 0.2, "גבול השקעה פנימי.", "0%"), ("משקל שקלי מקסימום", 0.8, "גבול השקעה פנימי.", "0%"),
              ("משקל צמוד מינימום", 0.2, "גבול השקעה פנימי.", "0%"), ("משקל צמוד מקסימום", 0.8, "גבול השקעה פנימי.", "0%"),
              ("סף תלילות 10Y-2Y", 0.005, "לזיהוי עקום תלול/שטוח.", "0.00%"),
              ("שוק מקבילי +bp", 100, "תרחיש עליית תשואות.", "0"), ("שוק מקבילי -bp", -100, "תרחיש ירידת תשואות.", "0")]
    exp_def = {0.25: 0.019, 0.5: 0.021, 1: 0.02, 2: 0.021, 3: 0.021, 4: 0.023, 5: 0.025, 7: 0.025, 10: 0.025, 15: 0.025, 20: 0.025, 30: 0.025}
    for i, (lab, val, note, fmt) in enumerate(params):
        r = 4 + i
        if prev and lab != "אופן קלט תשואות/קופון":
            pv = prev["01_הגדרות"][f"B{r}"].value
            if pv not in (None, "") and not (lab == "תאריך ניתוח"): val = pv
        put(ws, f"A{r}", lab, BOLD); put(ws, f"B{r}", val, BLUE, INP if lab != "אופן קלט תשואות/קופון" else None, fmt, BOX); put(ws, f"C{r}", note, NOTE)
    for i, t in enumerate(TENORS):
        r = 5 + i
        e = exp_def[t]; w = 1
        if prev:
            pe = prev["01_הגדרות"][f"E{r}"].value; pw = prev["01_הגדרות"][f"F{r}"].value
            if isinstance(pe, (int, float)): e = pe
            if isinstance(pw, (int, float)): w = pw
        put(ws, f"D{r}", t, BASE, None, "0.00", BOX); put(ws, f"E{r}", e, BLUE, INP, "0.00%", BOX); put(ws, f"F{r}", w, BLUE, INP, "0", BOX)
        put(ws, f"G{r}", f"='03_עקומים'!D{6+i}", Font(name="Arial", size=10, color="008000"), None, "0.00%", BOX)
    put(ws, "A21", "כללי צבע במודל", BOLD)
    for i, (a, b) in enumerate([("כחול על רקע צהוב", "קלט ידני / הנחה לשינוי"), ("שחור", "נוסחה"), ("ירוק", "קישור פנימי מגיליון אחר"), ("אפור", "נתון מהמאגר (ערך)")]):
        put(ws, f"A{22+i}", a); put(ws, f"B{22+i}", b)
    put(ws, "A27", "מקורות", BOLD)
    put(ws, "A28", "גלובס (טבלאות האג״ח הממשלתיות, כולל מק״מ) דרך corp-market-data — data/gov/bonds_latest.csv; היסטוריה: גלובס יומי מ-13.10.2026 + קובץ מגמה רבעוני 1/2023–10/2026.", NOTE)
    put(ws, "A29", "בנק ישראל — מאגר הסדרות (SDMX, dataflow ZCM): עקומי אפס וציפיות; FRED — תשואות ארה״ב. קובץ זה הוא מקור הממשלתי לכל המערכת (קצרות, אקסל העל, מודלי המנפיקים).", NOTE)
    for col, w in zip("ABCDEFG", (34, 14, 70, 11, 18, 10, 12)): ws.column_dimensions[col].width = w
    AD = "'01_הגדרות'!$B$4"

    # ================= 02_קלט_אגח =================
    wi = wb.create_sheet("02_קלט_אגח"); rtl(wi)
    put(wi, "A1", "קלט אג״ח ממשלתיות — מהמאגר שלנו (גלובס), ליום המסחר האחרון", TITLE)
    put(wi, "A2", "עמודות A–V: נתונים (ערכים; חיבור חי — ראו 08_מדריך, שאילתת Power Query על gov_bonds_latest.csv). מימין לקו: נוסחאות. ״כלול בעקום?״ — ״לא״ מוציא נייר מהעקום. כמות מוחזקת — קלט לתיק.", NOTE)
    dcols = ["date", "security_id", "name", "kind", "source", "maturity", "coupon", "last_price", "adjusted_value", "ytm_gross", "ytm_net", "duration_globes",
             "volume_k_nis", "ttm", "accrued", "clean_price", "macaulay_calc", "modified", "convexity", "L", "pv_check", "in_curve"]
    dheb = ["תאריך נתונים", "מס׳ ני״ע", "שם אג״ח", "אפיק", "טבלת גלובס", "תאריך פדיון", "קופון שנתי %", "מחיר (מלוכלך)", "ערך מתואם", "תשואה ברוטו %",
            "תשואה נטו %", "מח״מ גלובס", "מחזור (אלפי ₪)", "שנים לפדיון (ליום הנתונים)", "ריבית צבורה (חישוב)", "מחיר נקי (חישוב)", "מח״מ Macaulay (חישוב)",
            "מח״מ Modified (חישוב)", "קמירות (חישוב)", "מקדם הצמדה L", "מחיר תיאורטי (בדיקה)", "כלול בעקום?"]
    fcols = ["תשואה (עשרוני)", "שנים לפדיון (לתאריך הניתוח, סליקה T+1)", "Bucket מח״מ", "כלול (אפקטיבי)", "כמות נקובה מוחזקת", "שווי שוק", "משקל בתיק", "תרומת מח״מ",
             "תרומת תשואה", "תרומת קמירות", "x עקום שקלי", "x עקום צמוד", "הערות"]
    hdr(wi, 5, dheb + fcols, 1, 44); wi.row_dimensions[4].height = 4
    put(wi, "A4", "שדות: " + ", ".join(dcols), Font(name="Arial", size=7, color="808080"))
    C = {h: L(i + 1) for i, h in enumerate(dcols)}            # אותיות עמודות הנתונים
    F = {h: L(len(dcols) + i + 1) for i, h in enumerate(fcols)}
    cY, cT, cB, cI, cQ, cMV, cW, cXd, cXy, cXc, cXn, cXr, cN = (F[h] for h in fcols)
    prev_q, prev_in, prev_note = {}, {}, {}
    if prev:
        p2 = prev["02_קלט_אגח"]; hmap = {p2.cell(5, c).value: c for c in range(1, p2.max_column + 1)}
        for r in range(6, p2.max_row + 1):
            nm = p2.cell(r, 3).value
            if not nm: continue
            nm = str(nm).strip()
            c = hmap.get("כמות נקובה מוחזקת")
            if c and isinstance(p2.cell(r, c).value, (int, float)): prev_q[nm] = p2.cell(r, c).value
            c = hmap.get("כלול בעקום?")
            if c and p2.cell(r, c).value == "לא": prev_in[nm] = "לא"
            c = hmap.get("הערות")
            if c and p2.cell(r, c).value not in (None, ""): prev_note[nm] = p2.cell(r, c).value
    gfill = PatternFill("solid", fgColor="F2F2F2")
    for k in range(NB):
        r = R0 + k
        b = bonds.iloc[k] if k < len(bonds) else None
        for h in dcols:
            v = None
            if b is not None:
                v = b[h]
                if isinstance(v, float) and pd.isna(v): v = None
                if h in ("date", "maturity") and v: v = pd.to_datetime(v).to_pydatetime()
                if h == "security_id" and v is not None: v = int(v)
            c = wi[f"{C[h]}{r}"]; c.value = v; c.font = BASE; c.border = BOX
            if h in ("date", "maturity"): c.number_format = "dd/mm/yyyy"
            elif h in ("coupon", "last_price", "adjusted_value", "ytm_gross", "ytm_net", "duration_globes", "ttm", "accrued", "clean_price", "macaulay_calc", "modified", "convexity", "pv_check"): c.number_format = "0.00"
            elif h == "L": c.number_format = "0.0000"
            elif h == "volume_k_nis": c.number_format = "#,##0"
            if h == "in_curve":
                c.fill = INP; c.font = BLUE
                if b is not None and prev_in.get(str(b["name"]).strip()): c.value = "לא"
            elif b is not None: c.fill = gfill
        nm = C["name"]; ytm = C["ytm_gross"]; mat = C["maturity"]; kind = C["kind"]
        put(wi, f"{cY}{r}", f'=IF({ytm}{r}="","",{ytm}{r}/100)', BASE, None, "0.00%", BOX)
        put(wi, f"{cT}{r}", f'=IF({mat}{r}="","",MAX(0,({mat}{r}-{AD}-1)/365))', BASE, None, "0.00", BOX)
        put(wi, f"{cB}{r}", f'=IF({cT}{r}="","",IF({cT}{r}<=1,"0-1",IF({cT}{r}<=2,"1-2",IF({cT}{r}<=5,"2-5",IF({cT}{r}<=10,"5-10","10+")))))', BASE, None, None, BOX)
        put(wi, f"{cI}{r}", f'=IF({nm}{r}="","",IF({C["in_curve"]}{r}="לא","לא","כן"))', BASE, None, None, BOX)
        q = prev_q.get(str(b["name"]).strip()) if b is not None else None
        put(wi, f"{cQ}{r}", q, BLUE, INP, "#,##0", BOX)
        put(wi, f"{cMV}{r}", f'=IFERROR({cQ}{r}*{C["last_price"]}{r}/100,0)', BASE, None, "#,##0", BOX)
        put(wi, f"{cW}{r}", f'=IFERROR({cMV}{r}/SUM(${cMV}${R0}:${cMV}${R1}),0)', BASE, None, "0.0%", BOX)
        put(wi, f"{cXd}{r}", f'=IFERROR({cW}{r}*{C["modified"]}{r},0)', BASE, None, "0.000", BOX)
        put(wi, f"{cXy}{r}", f'=IFERROR({cW}{r}*{cY}{r},0)', BASE, None, "0.00%", BOX)
        put(wi, f"{cXc}{r}", f'=IFERROR({cW}{r}*{C["convexity"]}{r},0)', BASE, None, "0.00", BOX)
        put(wi, f"{cXn}{r}", f'=IF(AND({cI}{r}="כן",OR({kind}{r}="שקלי",{kind}{r}="מק""מ"),ISNUMBER({cT}{r}),ISNUMBER({cY}{r})),{cT}{r},"")', Font(name="Arial", size=9, color="808080"), None, "0.00")
        put(wi, f"{cXr}{r}", f'=IF(AND({cI}{r}="כן",{kind}{r}="מדד",ISNUMBER({cT}{r}),ISNUMBER({cY}{r})),{cT}{r},"")', Font(name="Arial", size=9, color="808080"), None, "0.00")
        put(wi, f"{cN}{r}", prev_note.get(str(b["name"]).strip()) if b is not None else None, BLUE, INP, None, BOX)
    dv = DataValidation(type="list", formula1='"כן,לא"', allow_blank=True); wi.add_data_validation(dv); dv.add(f"{C['in_curve']}{R0}:{C['in_curve']}{R1}")
    widths = [11, 10, 20, 7, 14, 11, 8, 9, 9, 9, 9, 8, 10, 10, 9, 9, 10, 10, 9, 8, 9, 8] + [9, 10, 8, 8, 12, 11, 8, 8, 9, 8, 8, 8, 20]
    for i, w in enumerate(widths): wi.column_dimensions[L(i + 1)].width = w
    wi.freeze_panes = f"D{R0}"; wi.auto_filter.ref = f"A5:{cN}{R1}"
    XN = f"'02_קלט_אגח'!${cXn}${R0}:${cXn}${R1}"; XR = f"'02_קלט_אגח'!${cXr}${R0}:${cXr}${R1}"; YY = f"'02_קלט_אגח'!${cY}${R0}:${cY}${R1}"
    B2 = lambda h: f"'02_קלט_אגח'!${C[h]}${R0}:${C[h]}${R1}"
    F2 = lambda h: f"'02_קלט_אגח'!${F[h]}${R0}:${F[h]}${R1}"

    # ================= 03_עקומים =================
    wc = wb.create_sheet("03_עקומים"); rtl(wc)
    put(wc, "A1", "עקומי אג״ח: שקלי נומינלי (מק״מ + שקלית קבועה), צמוד ריאלי, אינפלציה גלומה ופורוורדים", TITLE)
    put(wc, "A3", "אינטרפולציה ליניארית על שנים לפדיון (לתאריך הניתוח), מניירות שסומנו ״כן״ בעמודת כלול בעקום. מחוץ לטווח הניירות — ערך הקצה. חישוב מקביל בפייתון: data/gov/curves_latest.csv.", NOTE)
    ch = ["טווח בשנים", "שקלי YTM", "צמוד YTM ריאלי", "אינפלציה גלומה BEI", "BEI מצטברת", "פורוורד שקלי", "פורוורד ריאלי", "פורוורד אינפלציה",
          "שיפוע שקלי מול 2Y", "שיפוע צמוד מול 2Y", "Roll-down שקלי bp", "Roll-down צמוד bp", "אות החלטה", "lo שקלי", "hi שקלי", "lo צמוד", "hi צמוד", "x roll", "lo roll", "hi roll"]
    hdr(wc, 5, ch, 1, 32)
    for i, t in enumerate(TENORS):
        r = 6 + i
        put(wc, f"A{r}", f"='01_הגדרות'!D{5+i}", Font(name="Arial", size=10, color="008000"), None, "0.00", BOX)
        arr(wc, f"N{r}", "=" + lo_formula(f"$A{r}", XN)); arr(wc, f"O{r}", "=" + hi_formula(f"$A{r}", XN))
        arr(wc, f"P{r}", "=" + lo_formula(f"$A{r}", XR)); arr(wc, f"Q{r}", "=" + hi_formula(f"$A{r}", XR))
        put(wc, f"B{r}", "=" + interp(f"$A{r}", XN, YY, f"N{r}", f"O{r}"), BASE, None, "0.00%", BOX)
        put(wc, f"C{r}", "=" + interp(f"$A{r}", XR, YY, f"P{r}", f"Q{r}"), BASE, None, "0.00%", BOX)
        put(wc, f"D{r}", f'=IFERROR((1+B{r})/(1+C{r})-1,"")', BASE, None, "0.00%", BOX)
        put(wc, f"E{r}", f'=IFERROR((1+D{r})^A{r}-1,"")', BASE, None, "0.00%", BOX)
        if i == 0:
            put(wc, f"F{r}", f'=IFERROR(B{r},"")', BASE, None, "0.00%", BOX); put(wc, f"G{r}", f'=IFERROR(C{r},"")', BASE, None, "0.00%", BOX)
        else:
            put(wc, f"F{r}", f'=IFERROR(((1+B{r})^A{r}/(1+B{r-1})^A{r-1})^(1/(A{r}-A{r-1}))-1,"")', BASE, None, "0.00%", BOX)
            put(wc, f"G{r}", f'=IFERROR(((1+C{r})^A{r}/(1+C{r-1})^A{r-1})^(1/(A{r}-A{r-1}))-1,"")', BASE, None, "0.00%", BOX)
        put(wc, f"H{r}", f'=IFERROR((1+F{r})/(1+G{r})-1,"")', BASE, None, "0.00%", BOX)
        put(wc, f"I{r}", f'=IFERROR(B{r}-INDEX($B$6:$B$17,MATCH(2,$A$6:$A$17,0)),"")', BASE, None, "0.00%", BOX)
        put(wc, f"J{r}", f'=IFERROR(C{r}-INDEX($C$6:$C$17,MATCH(2,$A$6:$A$17,0)),"")', BASE, None, "0.00%", BOX)
        put(wc, f"R{r}", f"=MAX($A{r}-'01_הגדרות'!$B$7,MIN($A$6:$A$17))", Font(name="Arial", size=9, color="808080"), None, "0.00")
        arr(wc, f"S{r}", "=" + lo_formula(f"$R{r}", "$A$6:$A$17")); arr(wc, f"T{r}", "=" + hi_formula(f"$R{r}", "$A$6:$A$17"))
        put(wc, f"K{r}", f'=IFERROR((B{r}-' + interp(f"$R{r}", "$A$6:$A$17", "$B$6:$B$17", f"S{r}", f"T{r}") + ')*10000,"")', BASE, None, "0", BOX)
        put(wc, f"L{r}", f'=IFERROR((C{r}-' + interp(f"$R{r}", "$A$6:$A$17", "$C$6:$C$17", f"S{r}", f"T{r}") + ')*10000,"")', BASE, None, "0", BOX)
        put(wc, f"M{r}", f'=IF(D{r}="","",IF(\'01_הגדרות\'!$B$8>D{r}+\'01_הגדרות\'!$B$9,"עודף צמוד",IF(\'01_הגדרות\'!$B$8<D{r}-\'01_הגדרות\'!$B$9,"עודף שקלי","נייטרלי")))', BASE, None, None, BOX)
        for col in "NOPQST": wc[f"{col}{r}"].font = Font(name="Arial", size=9, color="808080")
    for i, w in enumerate([10, 10, 12, 12, 11, 11, 11, 12, 12, 12, 12, 12, 12, 8, 8, 8, 8, 7, 7, 7]): wc.column_dimensions[L(i + 1)].width = w
    wc.column_dimensions.group("N", "T", outline_level=1, hidden=True)
    put(wc, "A19", "עזר (עמודות N–T, מכווצות): נקודות העקום הסמוכות לכל טווח. R = הטווח באופק ה-roll-down.", NOTE)
    # גרף העקומים
    lc = LineChart(); lc.title = "עקום שקלי וצמוד"; lc.y_axis.title = "תשואה"; lc.x_axis.title = "שנים לפדיון"; lc.height = 8; lc.width = 18
    lc.add_data(Reference(wc, min_col=2, min_row=5, max_row=17), titles_from_data=True); lc.add_data(Reference(wc, min_col=3, min_row=5, max_row=17), titles_from_data=True)
    lc.set_categories(Reference(wc, min_col=1, min_row=6, max_row=17)); wc.add_chart(lc, "A21")
    lb = LineChart(); lb.title = "אינפלציה גלומה (BEI) לפי טווח"; lb.height = 8; lb.width = 18
    lb.add_data(Reference(wc, min_col=4, min_row=5, max_row=17), titles_from_data=True); lb.set_categories(Reference(wc, min_col=1, min_row=6, max_row=17)); wc.add_chart(lb, "H21")

    # ================= 04_ציפיות_אינפלציה =================
    we = wb.create_sheet("04_ציפיות_אינפלציה"); rtl(we)
    put(we, "A1", "ציפיות אינפלציה: BEI מול תחזית מנהל השקעות, בנק ישראל והטיית שקלי/צמוד", TITLE)
    put(we, "A3", "פער חיובי של התחזית מול BEI = עדיפות יחסית לצמוד; שלילי = לשקלי. התחזיות לפי טווח — בגיליון 01 (עמודה E).", NOTE)
    hdr(we, 5, ["טווח בשנים", "BEI מהעקום", "תחזית אינפלציה שנתית", "פער תחזית-BEI", "פורוורד אינפלציה", "פער פורוורד מול תחזית", "אות החלטה", "משקל החלטה", "BEI מצטברת", "תחזית מצטברת"])
    for i in range(12):
        r = 6 + i
        put(we, f"A{r}", f"='03_עקומים'!A{r}", BASE, None, "0.00", BOX); put(we, f"B{r}", f"='03_עקומים'!D{r}", BASE, None, "0.00%", BOX)
        put(we, f"C{r}", f"='01_הגדרות'!E{5+i}", Font(name="Arial", size=10, color="008000"), None, "0.00%", BOX)
        put(we, f"D{r}", f'=IFERROR(C{r}-B{r},"")', BASE, None, "0.00%", BOX); put(we, f"E{r}", f"='03_עקומים'!H{r}", BASE, None, "0.00%", BOX)
        put(we, f"F{r}", f'=IFERROR(C{r}-E{r},"")', BASE, None, "0.00%", BOX)
        put(we, f"G{r}", f'=IF(B{r}="","",IF(D{r}>\'01_הגדרות\'!$B$9,"צמוד",IF(D{r}<-\'01_הגדרות\'!$B$9,"שקלי","נייטרלי")))', BASE, None, None, BOX)
        put(we, f"H{r}", f"='01_הגדרות'!F{5+i}", Font(name="Arial", size=10, color="008000"), None, "0", BOX)
        put(we, f"I{r}", f"='03_עקומים'!E{r}", BASE, None, "0.00%", BOX); put(we, f"J{r}", f'=IFERROR((1+C{r})^A{r}-1,"")', BASE, None, "0.00%", BOX)
    put(we, "A20", "ציפיות לאינפלציה — בנק ישראל (מאגר הסדרות, חודשי; ערך אחרון)", BOLD)
    hdr(we, 21, ["מקור", "תקופה", "ערך %"])
    r = 22
    for _, x in boi_e.iterrows():
        put(we, f"A{r}", x["name"], BASE, None, None, BOX); put(we, f"B{r}", x["period"], BASE, None, None, BOX); put(we, f"C{r}", float(x["value"]), BASE, None, "0.00", BOX); r += 1
    if not len(boi_e): put(we, "A22", "לא נמשכו ציפיות מבנק ישראל בריצה זו", NOTE); r = 23
    put(we, f"A{r+1}", "השוואה: BEI לשנה מהעקום שלנו", BOLD); put(we, f"C{r+1}", "=IFERROR('03_עקומים'!D8*100,\"\")", BASE, None, "0.00", BOX)
    put(we, f"A{r+2}", "ציפיות המשמשות לגיליון ׳מחמ מתחת ל 1׳ באקסל העל", BOLD)
    put(we, f"B{r+2}", short_src.get("used", ""), BASE); put(we, f"C{r+2}", (short_src.get("expectations", {}).get(short_src.get("used", ""), {}) or {}).get("value"), BASE, None, "0.00")
    put(we, f"A{r+3}", "תוואי המדדים החודשי (שנתי × עונתיות) נבנה ב-corp-market-data (short/cpi_expected.py) ומוזן לגיליון הקצרות; מקור ברירת המחדל: BEI לשנה מקובץ זה (gov).", NOTE)
    for col, w in zip("ABCDEFGHIJ", (34, 14, 14, 13, 13, 15, 11, 10, 12, 12)): we.column_dimensions[col].width = w

    # ================= 05_אנליטיקה_נייר =================
    wa = wb.create_sheet("05_אנליטיקה_נייר"); rtl(wa)
    put(wa, "A1", "אנליטיקה ברמת נייר: עקום, BEI, Carry/Roll-down, תרחישי ריבית וניקוד", TITLE)
    put(wa, "A3", "הניקוד הוא מנגנון דירוג לדיון: ערך יחסי שקלי/צמוד, Carry/Roll-down, נזילות (מחזור) וסיכון מח״מ. P&L בתרחישים לפי מח״מ מתואם וקמירות.", NOTE)
    ah = ["מס׳ ני״ע", "שם", "אפיק", "שנים לפדיון", "מח״מ מתואם", "קמירות", "YTM", "YTM עקום", "פער מול העקום bp", "BEI בטווח", "צפי אינפלציה", "פער צפי-BEI",
          "YTM שקול נומינלי", "Roll-down bp", "תשואה צפויה 12M", "P&L +bp", "P&L -bp", "נזילות (מחזור אלפי ₪)", "ניקוד כולל", "פעולה מוצעת", "lo", "hi"]
    hdr(wa, 5, ah, 1, 32)
    for k in range(NB):
        r = R0 + k; s = f"'02_קלט_אגח'!"
        put(wa, f"A{r}", f'=IF({s}{C["name"]}{r}="","",{s}{C["security_id"]}{r})', BASE, None, "0", BOX)
        put(wa, f"B{r}", f'=IF({s}{C["name"]}{r}="","",{s}{C["name"]}{r})', BASE, None, None, BOX)
        put(wa, f"C{r}", f'=IF({s}{C["name"]}{r}="","",{s}{C["kind"]}{r})', BASE, None, None, BOX)
        put(wa, f"D{r}", f'=IF(B{r}="","",{s}{cT}{r})', BASE, None, "0.00", BOX)
        put(wa, f"E{r}", f'=IF(B{r}="","",{s}{C["modified"]}{r})', BASE, None, "0.00", BOX)
        put(wa, f"F{r}", f'=IF(B{r}="","",{s}{C["convexity"]}{r})', BASE, None, "0.0", BOX)
        put(wa, f"G{r}", f'=IF(B{r}="","",{s}{cY}{r})', BASE, None, "0.00%", BOX)
        arr(wa, f"U{r}", f'=IF(D{r}="",-1,' + lo_formula(f"$D{r}", "'03_עקומים'!$A$6:$A$17") + ")")
        arr(wa, f"V{r}", f'=IF(D{r}="",9999,' + hi_formula(f"$D{r}", "'03_עקומים'!$A$6:$A$17") + ")")
        yc = f'IF(C{r}="מדד",' + interp(f"$D{r}", "'03_עקומים'!$A$6:$A$17", "'03_עקומים'!$C$6:$C$17", f"U{r}", f"V{r}") + "," + interp(f"$D{r}", "'03_עקומים'!$A$6:$A$17", "'03_עקומים'!$B$6:$B$17", f"U{r}", f"V{r}") + ")"
        put(wa, f"H{r}", f'=IF(OR(B{r}="",D{r}=""),"",{yc})', BASE, None, "0.00%", BOX)
        put(wa, f"I{r}", f'=IF(OR(G{r}="",H{r}=""),"",(G{r}-H{r})*10000)', BASE, None, "0", BOX)
        put(wa, f"J{r}", f'=IF(OR(B{r}="",D{r}=""),"",' + interp(f"$D{r}", "'03_עקומים'!$A$6:$A$17", "'03_עקומים'!$D$6:$D$17", f"U{r}", f"V{r}") + ")", BASE, None, "0.00%", BOX)
        put(wa, f"K{r}", f'=IF(OR(B{r}="",D{r}=""),"",' + interp(f"$D{r}", "'01_הגדרות'!$D$5:$D$16", "'01_הגדרות'!$E$5:$E$16", f"U{r}", f"V{r}") + ")", BASE, None, "0.00%", BOX)
        put(wa, f"L{r}", f'=IFERROR(K{r}-J{r},"")', BASE, None, "0.00%", BOX)
        put(wa, f"M{r}", f'=IF(B{r}="","",IF(C{r}="מדד",(1+G{r})*(1+K{r})-1,G{r}))', BASE, None, "0.00%", BOX)
        rd_ = f'IF(C{r}="מדד",' + interp(f"$D{r}", "'03_עקומים'!$A$6:$A$17", "'03_עקומים'!$L$6:$L$17", f"U{r}", f"V{r}") + "," + interp(f"$D{r}", "'03_עקומים'!$A$6:$A$17", "'03_עקומים'!$K$6:$K$17", f"U{r}", f"V{r}") + ")"
        put(wa, f"N{r}", f'=IF(OR(B{r}="",D{r}=""),"",{rd_})', BASE, None, "0", BOX)
        put(wa, f"O{r}", f"=IFERROR(M{r}*'01_הגדרות'!$B$7+E{r}*N{r}/10000,\"\")", BASE, None, "0.00%", BOX)
        put(wa, f"P{r}", f"=IFERROR(-E{r}*('01_הגדרות'!$B$17/10000)+0.5*F{r}*('01_הגדרות'!$B$17/10000)^2,\"\")", BASE, None, "0.00%", BOX)
        put(wa, f"Q{r}", f"=IFERROR(-E{r}*('01_הגדרות'!$B$18/10000)+0.5*F{r}*('01_הגדרות'!$B$18/10000)^2,\"\")", BASE, None, "0.00%", BOX)
        put(wa, f"R{r}", f'=IF(B{r}="","",{s}{C["volume_k_nis"]}{r})', BASE, None, "#,##0", BOX)
        put(wa, f"S{r}", f'=IF(OR(B{r}="",O{r}=""),"",MAX(0,MIN(100,50+IF(C{r}="מדד",L{r}*10000,-L{r}*10000)*0.35+O{r}*100*1.5+MIN(5,R{r}/5000)*2-E{r}*0.35)))', BASE, None, "0", BOX)
        put(wa, f"T{r}", f'=IF(S{r}="","",IF(S{r}>=65,"מועמד להגדלה",IF(S{r}<=40,"מועמד להפחתה","להחזיק/מעקב")))', BASE, None, None, BOX)
        for col in "UV": wa[f"{col}{r}"].font = Font(name="Arial", size=9, color="808080")
    for i, w in enumerate([10, 20, 7, 9, 9, 8, 8, 9, 10, 9, 10, 10, 11, 10, 11, 9, 9, 12, 8, 14, 6, 6]): wa.column_dimensions[L(i + 1)].width = w
    wa.column_dimensions.group("U", "V", outline_level=1, hidden=True); wa.freeze_panes = f"C{R0}"; wa.auto_filter.ref = f"A5:T{R1}"
    wa.conditional_formatting.add(f"I{R0}:I{R1}", FormulaRule(formula=[f'AND(I{R0}<>"",I{R0}>=10)'], fill=GRN))
    wa.conditional_formatting.add(f"I{R0}:I{R1}", FormulaRule(formula=[f'AND(I{R0}<>"",I{R0}<=-10)'], fill=RED))

    # ================= 06_תיק_והחלטה =================
    wp = wb.create_sheet("06_תיק_והחלטה"); rtl(wp)
    put(wp, "A1", "דשבורד החלטה לוועדת השקעות — מח״מ, שקלי/צמוד, עקום וסיכון", TITLE)
    put(wp, "A2", "מדדי התיק נבנים מעמודת ״כמות נקובה מוחזקת״ ב-02 (ריק = אין תיק). הסימונים ב-07/08/09 אינם תלויים בתיק.", NOTE)
    put(wp, "A3", "מדדי תיק נוכחי", BOLD)
    MV, Wt, Kd = F2("שווי שוק"), F2("משקל בתיק"), B2("kind")
    cells = [("A4", "שווי שוק"), ("B4", f"=SUM({MV})"), ("C4", "מח״מ תיק (מתואם)"), ("D4", f"=SUM({F2('תרומת מח״מ')})"), ("E4", "תשואה משוקללת"), ("F4", f"=SUM({F2('תרומת תשואה')})"),
             ("A5", "משקל שקלי (כולל מק״מ)"), ("B5", f'=SUMIFS({Wt},{Kd},"שקלי")+SUMIFS({Wt},{Kd},"מק""מ")'), ("C5", "משקל צמוד"), ("D5", f'=SUMIFS({Wt},{Kd},"מדד")'), ("E5", "קמירות משוקללת"), ("F5", f"=SUM({F2('תרומת קמירות')})"),
             ("A6", "יעד מח״מ"), ("B6", "='01_הגדרות'!$B$6"), ("C6", "פער מח״מ"), ("D6", "=D4-B6"), ("E6", "צפי אינפלציה 12M"), ("F6", "='01_הגדרות'!$B$8"),
             ("A7", "BEI 2Y"), ("B7", "=INDEX('03_עקומים'!$D$6:$D$17,MATCH(2,'03_עקומים'!$A$6:$A$17,0))"), ("C7", "BEI 5Y"), ("D7", "=INDEX('03_עקומים'!$D$6:$D$17,MATCH(5,'03_עקומים'!$A$6:$A$17,0))"),
             ("E7", "BEI 10Y"), ("F7", "=INDEX('03_עקומים'!$D$6:$D$17,MATCH(10,'03_עקומים'!$A$6:$A$17,0))"),
             ("A8", "אות שקלי/צמוד"), ("B8", '=IF(F6>D7+\'01_הגדרות\'!$B$9,"הטיה לצמוד",IF(F6<D7-\'01_הגדרות\'!$B$9,"הטיה לשקלי","נייטרלי"))'),
             ("C8", "אות מח״מ"), ("D8", '=IF(B4=0,"אין תיק",IF(D4>\'01_הגדרות\'!$B$6+0.25,"ארוך מהיעד",IF(D4<\'01_הגדרות\'!$B$6-0.25,"קצר מהיעד","קרוב ליעד")))'),
             ("E8", "אות עקום (10Y-2Y שקלי)"), ("F8", "=IF((INDEX('03_עקומים'!$B$6:$B$17,MATCH(10,'03_עקומים'!$A$6:$A$17,0))-INDEX('03_עקומים'!$B$6:$B$17,MATCH(2,'03_עקומים'!$A$6:$A$17,0)))>'01_הגדרות'!$B$16,\"תלול\",\"שטוח\")"),
             ("A9", "10Y-2Y שקלי (bp)"), ("B9", "=(INDEX('03_עקומים'!$B$6:$B$17,MATCH(10,'03_עקומים'!$A$6:$A$17,0))-INDEX('03_עקומים'!$B$6:$B$17,MATCH(2,'03_עקומים'!$A$6:$A$17,0)))*10000"),
             ("C9", "10Y-2Y צמוד (bp)"), ("D9", "=(INDEX('03_עקומים'!$C$6:$C$17,MATCH(10,'03_עקומים'!$A$6:$A$17,0))-INDEX('03_עקומים'!$C$6:$C$17,MATCH(2,'03_עקומים'!$A$6:$A$17,0)))*10000"),
             ("E9", "ישראל-ארה״ב 10Y (bp)"), ("F9", "=IFERROR((INDEX('03_עקומים'!$B$6:$B$17,MATCH(10,'03_עקומים'!$A$6:$A$17,0))*100-'10_ארהב'!$C$10)*100,\"\")")]
    fm = {"B4": "#,##0", "D4": "0.00", "F4": "0.00%", "B5": "0.0%", "D5": "0.0%", "F5": "0.0", "B6": "0.00", "D6": "0.00", "F6": "0.00%", "B7": "0.00%", "D7": "0.00%", "F7": "0.00%", "B9": "0", "D9": "0", "F9": "0"}
    for cell, v in cells:
        put(wp, cell, v, BOLD if cell[0] in "ACE" else BASE, YEL if cell in ("B8", "D8", "F8") else None, fm.get(cell), BOX)
    put(wp, "A11", "החלטה לפי Bucket (שנים לפדיון)", BOLD)
    hdr(wp, 12, ["Bucket", "משקל שקלי", "משקל צמוד", "משקל כולל", "YTM שקלי ממוצע", "YTM צמוד ממוצע", "BEI ממוצע", "ניקוד שקלי", "ניקוד צמוד", "עדיפות אפיק", "מח״מ ממוצע", "Carry/Roll ממוצע", "P&L +bp", "P&L -bp"])
    A5 = lambda col: f"'05_אנליטיקה_נייר'!${col}${R0}:${col}${R1}"; BK = F2("Bucket מח״מ")
    for i, bk in enumerate(["0-1", "1-2", "2-5", "5-10", "10+"]):
        r = 13 + i
        put(wp, f"A{r}", bk, BOLD, None, None, BOX)
        put(wp, f"B{r}", f'=SUMIFS({Wt},{Kd},"שקלי",{BK},$A{r})+SUMIFS({Wt},{Kd},"מק""מ",{BK},$A{r})', BASE, None, "0.0%", BOX)
        put(wp, f"C{r}", f'=SUMIFS({Wt},{Kd},"מדד",{BK},$A{r})', BASE, None, "0.0%", BOX); put(wp, f"D{r}", f"=B{r}+C{r}", BASE, None, "0.0%", BOX)
        put(wp, f"E{r}", f'=IFERROR(AVERAGEIFS({A5("G")},{A5("C")},"שקלי",{BK},$A{r}),"")', BASE, None, "0.00%", BOX)
        put(wp, f"F{r}", f'=IFERROR(AVERAGEIFS({A5("G")},{A5("C")},"מדד",{BK},$A{r}),"")', BASE, None, "0.00%", BOX)
        put(wp, f"G{r}", f'=IFERROR(AVERAGEIFS({A5("J")},{BK},$A{r}),"")', BASE, None, "0.00%", BOX)
        put(wp, f"H{r}", f'=IFERROR(AVERAGEIFS({A5("S")},{A5("C")},"שקלי",{BK},$A{r}),"")', BASE, None, "0", BOX)
        put(wp, f"I{r}", f'=IFERROR(AVERAGEIFS({A5("S")},{A5("C")},"מדד",{BK},$A{r}),"")', BASE, None, "0", BOX)
        put(wp, f"J{r}", f'=IF(OR(H{r}="",I{r}=""),"",IF(H{r}>I{r}+5,"שקלי",IF(I{r}>H{r}+5,"צמוד","נייטרלי")))', BASE, YEL, None, BOX)
        put(wp, f"K{r}", f'=IFERROR(AVERAGEIFS({A5("E")},{BK},$A{r}),"")', BASE, None, "0.00", BOX)
        put(wp, f"L{r}", f'=IFERROR(AVERAGEIFS({A5("O")},{BK},$A{r}),"")', BASE, None, "0.00%", BOX)
        put(wp, f"M{r}", f'=IFERROR(AVERAGEIFS({A5("P")},{BK},$A{r}),"")', BASE, None, "0.00%", BOX)
        put(wp, f"N{r}", f'=IFERROR(AVERAGEIFS({A5("Q")},{BK},$A{r}),"")', BASE, None, "0.00%", BOX)
    put(wp, "A20", "צ׳ק ליסט החלטה לוועדה", BOLD); hdr(wp, 21, ["שאלה", "בדיקה במודל", "כלל החלטה מקצועי"])
    for i, (a, b, c) in enumerate([("שקלי מול צמוד", "פער צפי אינפלציה מול BEI לפי טווח (04)", "צמוד אם צפי > BEI + מרווח; שקלי אם צפי < BEI − מרווח."),
                                    ("מח״מ", "פער מח״מ תיק מול יעד + צורת העקום (03)", "להאריך כשה-Carry/Roll עדיף והסיכון לתזוזת תשואות סביר."),
                                    ("סינתטי מול בולט", "גיליון 07: תשואה, קמירות, roll-down ונזילות", "סינתטי עדיף אם נותן קמירות/roll-down בלי ויתור תשואה מהותי ונזילות מספקת."),
                                    ("סיכון", "תרחישי +bp/-bp ברמת נייר ובאגרגציה (05, 06)", "לא לקבל החלטת תשואה בלי לראות הפסד בתרחיש עליית תשואות."),
                                    ("הקשר", "09 היסטוריה (איפה BEI/תשואות מול 3 שנים), 10 ארה״ב, 11 בנק ישראל", "פער ישראל–ארה״ב ועקומי האפס של בנק ישראל כבדיקת סבירות לעקום שלנו.")]):
        put(wp, f"A{22+i}", a, BOLD, None, None, BOX); put(wp, f"B{22+i}", b, BASE, None, None, BOX); put(wp, f"C{22+i}", c, BASE, None, None, BOX)
    for col, w in zip("ABCDEFGHIJKLMN", (22, 14, 20, 14, 22, 14, 11, 10, 10, 11, 11, 13, 10, 10)): wp.column_dimensions[col].width = w

    # ================= 07_סינטטי_מול_בולט =================
    wsb = wb.create_sheet("07_סינטטי_מול_בולט"); rtl(wsb)
    put(wsb, "A1", "השוואת מח״מ סינתטי מול בולט", TITLE)
    put(wsb, "A3", "בוחרים (תאים צהובים) נייר קצר, נייר ארוך ונייר בולט לפי מס׳ ני״ע מגיליון 02. המודל מחשב את המשקל הדרוש להשגת מח״מ היעד ומשווה תשואה, קמירות, roll-down ונזילות.", NOTE)
    hdr(wsb, 5, ["פרמטר", "ערך", "הערה"])
    # ברירות מחדל: שקלי קצר, ארוך (~10 שנים) ובולט הקרוב ליעד
    shk = bonds[bonds["kind"].isin(["שקלי", "מק\"מ"])].dropna(subset=["modified"]).sort_values("modified")
    def pick(target):
        if shk.empty: return None
        row = shk.iloc[(shk["modified"] - target).abs().argsort().iloc[0]]; return int(row["security_id"]) if pd.notna(row["security_id"]) else None
    tgt = params[2][1] if not prev else (prev["01_הגדרות"]["B6"].value or 4.5)
    defs = [("מח״מ יעד", "='01_הגדרות'!$B$6", "נמשך מהגדרות", "0.00"), ("אפיק להשוואה", "שקלי", "שקלי או מדד", None),
            ("מס׳ ני״ע קצר", pick(1.0), "נייר קצר / מק״מ", "0"), ("מס׳ ני״ע ארוך", pick(9.0), "נייר ארוך יותר מהיעד", "0"),
            ("מס׳ ני״ע בולט", pick(float(tgt)), "נייר יחיד במח״מ קרוב ליעד", "0"), ("סטיית מח״מ מקסימלית", 0.3, "mismatch מותר סביב היעד", "0.00"),
            ("עלות/חיכוך החלטה bp", 5, "פער תשואה מינימלי", "0")]
    for i, (a, v, n, f) in enumerate(defs):
        r = 6 + i
        if prev and i in (1, 2, 3, 4, 5, 6):
            pv = prev["07_סינטטי_מול_בולט"][f"B{r}"].value
            if pv not in (None, ""): v = pv
        put(wsb, f"A{r}", a, BOLD, None, None, BOX); put(wsb, f"B{r}", v, BLUE, INP if i else None, f, BOX); put(wsb, f"C{r}", n, NOTE)
    put(wsb, "A15", "רכיבי אסטרטגיה", BOLD)
    hdr(wsb, 16, ["רכיב", "מס׳ ני״ע", "שם", "מח״מ מתואם", "YTM", "קמירות", "Roll-down bp", "נזילות (מחזור)", "משקל", "תרומת תשואה", "תרומת קמירות"])
    ID = B2("security_id")
    def lk(col, key): return f'IFERROR(INDEX({col},MATCH({key},{ID},0)),"")'
    for i, (lab, src) in enumerate([("קצר", "$B$8"), ("ארוך", "$B$9"), ("בולט", "$B$10")]):
        r = 17 + i
        put(wsb, f"A{r}", lab, BOLD, None, None, BOX); put(wsb, f"B{r}", f"={src}", BASE, None, "0", BOX)
        put(wsb, f"C{r}", "=" + lk(B2("name"), f"B{r}"), BASE, None, None, BOX); put(wsb, f"D{r}", "=" + lk(B2("modified"), f"B{r}"), BASE, None, "0.00", BOX)
        put(wsb, f"E{r}", "=" + lk(F2("תשואה (עשרוני)"), f"B{r}"), BASE, None, "0.00%", BOX); put(wsb, f"F{r}", "=" + lk(B2("convexity"), f"B{r}"), BASE, None, "0.0", BOX)
        put(wsb, f"G{r}", f'=IFERROR(INDEX({A5("N")},MATCH(B{r},{A5("A")},0)),"")', BASE, None, "0", BOX); put(wsb, f"H{r}", "=" + lk(B2("volume_k_nis"), f"B{r}"), BASE, None, "#,##0", BOX)
    put(wsb, "I17", '=IFERROR(1-I18,"")', BASE, None, "0.0%", BOX); put(wsb, "I18", '=IFERROR(($B$6-D17)/(D18-D17),"")', BASE, None, "0.0%", BOX); put(wsb, "I19", 1, BASE, None, "0.0%", BOX)
    for r in (17, 18, 19):
        put(wsb, f"J{r}", f'=IFERROR(I{r}*E{r},"")', BASE, None, "0.00%", BOX); put(wsb, f"K{r}", f'=IFERROR(I{r}*F{r},"")', BASE, None, "0.0", BOX)
    put(wsb, "A22", "סיכום החלטה", BOLD); hdr(wsb, 23, ["מדד", "סינתטי", "בולט", "פער סינתטי-בולט", "פרשנות"])
    summ = [("מח״מ", "=I17*D17+I18*D18", "=D19", '=IF(ABS(D24)<=$B$11,"תקין","סטיית מח״מ דורשת תיקון")', "0.00"),
            ("YTM", "=SUM(J17:J18)", "=E19", '=IF(D25>$B$12/10000,"יתרון סינתטי",IF(D25<-$B$12/10000,"יתרון בולט","דומה"))', "0.00%"),
            ("קמירות", "=SUM(K17:K18)", "=F19", '=IF(D26>0,"סינתטי קמור יותר",IF(D26<0,"בולט קמור יותר","דומה"))', "0.0"),
            ("Roll-down bp", '=IFERROR(I17*G17+I18*G18,"")', '=IFERROR(G19,"")', '=IF(D27>0,"יתרון roll-down לסינתטי",IF(D27<0,"יתרון roll-down לבולט","דומה"))', "0"),
            ("נזילות (מחזור)", "=I17*H17+I18*H18", "=H19", '=IF(D28>=0,"סינתטי לא נחות בנזילות","בולט נזיל יותר")', "#,##0"),
            ("תשואה צפויה 12M", "=IFERROR(B25*'01_הגדרות'!$B$7+B24*B27/10000,\"\")", "=IFERROR(C25*'01_הגדרות'!$B$7+C24*C27/10000,\"\")", '=IF(D29>0,"יתרון צפוי לסינתטי",IF(D29<0,"יתרון צפוי לבולט","דומה"))', "0.00%")]
    for i, (a, b, c, e, f) in enumerate(summ):
        r = 24 + i
        put(wsb, f"A{r}", a, BOLD, None, None, BOX); put(wsb, f"B{r}", b, BASE, None, f, BOX); put(wsb, f"C{r}", c, BASE, None, f, BOX)
        put(wsb, f"D{r}", f"=IFERROR(B{r}-C{r},\"\")", BASE, None, f, BOX); put(wsb, f"E{r}", e, BASE, None, None, BOX)
    put(wsb, "A30", "החלטה", BOLD, None, None, BOX)
    put(wsb, "B30", '=IFERROR(IF(AND(ABS(D24)<=$B$11,D29>0),"סינתטי עדיף",IF(D29<0,"בולט עדיף","אין יתרון חד")),"")', BOLD, YEL, None, BOX)
    for col, w in zip("ABCDEFGHIJK", (22, 12, 22, 12, 10, 10, 12, 14, 9, 12, 12)): wsb.column_dimensions[col].width = w

    # ================= 09_היסטוריה =================
    wh = wb.create_sheet("09_היסטוריה"); rtl(wh)
    put(wh, "A1", "היסטוריית העקום: שקלי, ריאלי ו-BEI ב-1/2/5/10 שנים", TITLE)
    put(wh, "A2", "גלובס יומי (מ-13.10.2026, אותו חישוב כמו 03) + נקודות רבעוניות מקובץ מגמה (1/2023–10/2026; שנים לפדיון נגזרו משם הסדרה). ערכים באחוזים.", NOTE)
    hh = ["תאריך", "מקור", "שקלי 1Y", "שקלי 2Y", "שקלי 5Y", "שקלי 10Y", "ריאלי 1Y", "ריאלי 2Y", "ריאלי 5Y", "ריאלי 10Y", "BEI 1Y", "BEI 2Y", "BEI 5Y", "BEI 10Y"]
    hdr(wh, 4, hh)
    hcols = ["date", "src", "nom_1", "nom_2", "nom_5", "nom_10", "real_1", "real_2", "real_5", "real_10", "bei_1", "bei_2", "bei_5", "bei_10"]
    for i, (_, x) in enumerate(hist.iterrows()):
        r = 5 + i
        for j, h in enumerate(hcols):
            v = x[h]
            if h == "date": v = pd.to_datetime(v).to_pydatetime()
            elif isinstance(v, float) and pd.isna(v): v = None
            c = wh.cell(r, 1 + j, v); c.font = BASE; c.border = BOX
            if h == "date": c.number_format = "dd/mm/yyyy"
            elif j >= 2: c.number_format = "0.00"
    nh = len(hist); last = 4 + nh
    put(wh, f"A{last+2}", "מצב נוכחי מול ההיסטוריה", BOLD); hdr(wh, last + 3, ["סדרה", "נוכחי", "מינימום", "מקסימום", "ממוצע", "אחוזון נוכחי"])
    for i, (lab, col) in enumerate([("שקלי 2Y", "D"), ("שקלי 10Y", "F"), ("ריאלי 2Y", "H"), ("ריאלי 10Y", "J"), ("BEI 2Y", "L"), ("BEI 5Y", "M"), ("BEI 10Y", "N")]):
        r = last + 4 + i; rng = f"${col}$5:${col}${last}"
        put(wh, f"A{r}", lab, BOLD, None, None, BOX); put(wh, f"B{r}", f"={col}{last}", BASE, None, "0.00", BOX)
        put(wh, f"C{r}", f"=MIN({rng})", BASE, None, "0.00", BOX); put(wh, f"D{r}", f"=MAX({rng})", BASE, None, "0.00", BOX); put(wh, f"E{r}", f"=AVERAGE({rng})", BASE, None, "0.00", BOX)
        put(wh, f"F{r}", f"=IFERROR(PERCENTRANK({rng},B{r}),\"\")", BASE, None, "0%", BOX)
    if nh > 1:
        c1 = LineChart(); c1.title = "אינפלציה גלומה (BEI) 2/5/10 שנים"; c1.height = 8; c1.width = 20
        for col in (12, 13, 14): c1.add_data(Reference(wh, min_col=col, min_row=4, max_row=last), titles_from_data=True)
        c1.set_categories(Reference(wh, min_col=1, min_row=5, max_row=last)); wh.add_chart(c1, "P4")
        c2 = LineChart(); c2.title = "תשואה שקלית 2/10 שנים וריאלית 10"; c2.height = 8; c2.width = 20
        for col in (4, 6, 10): c2.add_data(Reference(wh, min_col=col, min_row=4, max_row=last), titles_from_data=True)
        c2.set_categories(Reference(wh, min_col=1, min_row=5, max_row=last)); wh.add_chart(c2, "P22")
    for i, w in enumerate([11, 14] + [9] * 12): wh.column_dimensions[L(i + 1)].width = w
    wh.freeze_panes = "C5"

    # ================= 10_ארהב =================
    wu = wb.create_sheet("10_ארהב"); rtl(wu)
    put(wu, "A1", "ארה״ב — ריבית הפד, עקום UST, TIPS ו-Breakeven (FRED)", TITLE)
    put(wu, "A2", "ערך אחרון זמין ב-FRED (יום עסקים אחרון בארה״ב), באחוזים, ושינוי בנ״ב מול חודש ו-3 חודשים. לא מהדוחות — מקור חיצוני.", NOTE)
    hdr(wu, 4, ["סדרה", "שם", "תאריך", "ערך %", "שינוי 1M (bp)", "שינוי 3M (bp)", "מקור"])
    for i, (_, x) in enumerate(us.iterrows()):
        r = 5 + i
        put(wu, f"A{r}", x["series"], BASE, None, None, BOX); put(wu, f"B{r}", x["name"], BASE, None, None, BOX)
        put(wu, f"C{r}", x.get("date"), BASE, None, None, BOX); put(wu, f"D{r}", None if pd.isna(x.get("value")) else float(x["value"]), BASE, None, "0.00", BOX)
        put(wu, f"E{r}", None if pd.isna(x.get("chg_1m_bp")) else float(x["chg_1m_bp"]), BASE, None, "+0;-0;0", BOX); put(wu, f"F{r}", None if pd.isna(x.get("chg_3m_bp")) else float(x["chg_3m_bp"]), BASE, None, "+0;-0;0", BOX)
        put(wu, f"G{r}", x.get("source"), NOTE, None, None, BOX)
    if not len(us): put(wu, "A5", "לא נמשכו נתוני FRED בריצה זו", NOTE)
    # טבלת השוואה ישראל–ארה"ב (D10 = UST 10 שנים אם הסדר נשמר; נמצא לפי MATCH)
    put(wu, "A18", "השוואה ישראל–ארה״ב", BOLD); hdr(wu, 19, ["טווח", "ישראל שקלי %", "UST %", "פער (bp)", "ישראל ריאלי %", "TIPS %", "פער ריאלי (bp)", "BEI ישראל %", "BEI ארה״ב %", "פער BEI (bp)"])
    def usv(sid): return f'IFERROR(INDEX($D$5:$D$16,MATCH("{sid}",$A$5:$A$16,0)),"")'
    for i, (t, un, ur, ub) in enumerate([(2, "DGS2", None, None), (5, "DGS5", "DFII5", "T5YIE"), (10, "DGS10", "DFII10", "T10YIE"), (30, "DGS30", None, None)]):
        r = 20 + i; il = f"INDEX('03_עקומים'!$B$6:$B$17,MATCH({t},'03_עקומים'!$A$6:$A$17,0))*100"; ilr = f"INDEX('03_עקומים'!$C$6:$C$17,MATCH({t},'03_עקומים'!$A$6:$A$17,0))*100"
        ilb = f"INDEX('03_עקומים'!$D$6:$D$17,MATCH({t},'03_עקומים'!$A$6:$A$17,0))*100"
        put(wu, f"A{r}", t, BOLD, None, "0", BOX); put(wu, f"B{r}", f'=IFERROR({il},"")', BASE, None, "0.00", BOX); put(wu, f"C{r}", "=" + usv(un), BASE, None, "0.00", BOX)
        put(wu, f"D{r}", f'=IFERROR((B{r}-C{r})*100,"")', BASE, None, "0", BOX); put(wu, f"E{r}", f'=IFERROR({ilr},"")', BASE, None, "0.00", BOX)
        put(wu, f"F{r}", ("=" + usv(ur)) if ur else None, BASE, None, "0.00", BOX); put(wu, f"G{r}", f'=IFERROR((E{r}-F{r})*100,"")', BASE, None, "0", BOX)
        put(wu, f"H{r}", f'=IFERROR({ilb},"")', BASE, None, "0.00", BOX); put(wu, f"I{r}", ("=" + usv(ub)) if ub else None, BASE, None, "0.00", BOX); put(wu, f"J{r}", f'=IFERROR((H{r}-I{r})*100,"")', BASE, None, "0", BOX)
    # תא ייחוס ל-06: UST 10Y
    put(wu, "C10", "=" + usv("DGS10"), BASE, None, "0.00") if wu["A10"].value != "DGS10" else None
    for col, w in zip("ABCDEFGHIJ", (10, 26, 12, 9, 12, 12, 22, 12, 11, 12)): wu.column_dimensions[col].width = w

    # ================= 11_בנק_ישראל =================
    wbi = wb.create_sheet("11_בנק_ישראל"); rtl(wbi)
    put(wbi, "A1", "בנק ישראל — עקומי אפס (שקלי/ריאלי) מול העקום שלנו", TITLE)
    put(wbi, "A2", "עקומי האפס של בנק ישראל (מאגר הסדרות, ZCM, חודשי — ממוצע החודש האחרון שפורסם) לעומת עקום התשואות לפדיון שלנו (03). הפרש סביר של 10–30 נ״ב בגלל שיטה (אפס מול YTM) ותאריך.", NOTE)
    hdr(wbi, 4, ["טווח (שנים)", "תקופה", "אפס שקלי %", "אפס ריאלי %", "BEI אפס %", "שלנו שקלי %", "שלנו ריאלי %", "שלנו BEI %", "הפרש שקלי (bp)", "הפרש ריאלי (bp)", "lo", "hi"])
    for i, (_, x) in enumerate(boi_c.iterrows()):
        r = 5 + i; t = int(x["tenor"])
        put(wbi, f"A{r}", t, BOLD, None, "0", BOX); put(wbi, f"B{r}", x["period"], BASE, None, None, BOX)
        put(wbi, f"C{r}", None if pd.isna(x.get("nominal")) else float(x["nominal"]), BASE, None, "0.00", BOX); put(wbi, f"D{r}", None if pd.isna(x.get("real")) else float(x["real"]), BASE, None, "0.00", BOX)
        put(wbi, f"E{r}", None if pd.isna(x.get("bei")) else float(x["bei"]), BASE, None, "0.00", BOX)
        arr(wbi, f"K{r}", "=" + lo_formula(f"$A{r}", "'03_עקומים'!$A$6:$A$17")); arr(wbi, f"L{r}", "=" + hi_formula(f"$A{r}", "'03_עקומים'!$A$6:$A$17"))
        put(wbi, f"F{r}", "=IFERROR(" + interp(f"$A{r}", "'03_עקומים'!$A$6:$A$17", "'03_עקומים'!$B$6:$B$17", f"K{r}", f"L{r}") + '*100,"")', BASE, None, "0.00", BOX)
        put(wbi, f"G{r}", "=IFERROR(" + interp(f"$A{r}", "'03_עקומים'!$A$6:$A$17", "'03_עקומים'!$C$6:$C$17", f"K{r}", f"L{r}") + '*100,"")', BASE, None, "0.00", BOX)
        put(wbi, f"H{r}", f'=IFERROR(((1+F{r}/100)/(1+G{r}/100)-1)*100,"")', BASE, None, "0.00", BOX)
        put(wbi, f"I{r}", f'=IFERROR((F{r}-C{r})*100,"")', BASE, None, "0", BOX); put(wbi, f"J{r}", f'=IFERROR((G{r}-D{r})*100,"")', BASE, None, "0", BOX)
        for col in "KL": wbi[f"{col}{r}"].font = Font(name="Arial", size=9, color="808080")
    if not len(boi_c): put(wbi, "A5", "לא נמשכו עקומי בנק ישראל בריצה זו", NOTE)
    wbi.column_dimensions.group("K", "L", outline_level=1, hidden=True)
    for col, w in zip("ABCDEFGHIJ", (11, 10, 11, 11, 10, 11, 11, 10, 13, 13)): wbi.column_dimensions[col].width = w

    # ================= 08_מדריך =================
    wg = wb.create_sheet("08_מדריך"); rtl(wg); wg.column_dimensions["A"].width = 130
    lines = [("מדריך שימוש — קובץ הממשלתי (מקור הממשלתי של מערכת CORP)", "t"), ("", ""),
             ("מה הקובץ עושה", "h"),
             ("עקום שקלי (מק״מ + שקלית קבועה) ועקום צמוד מאג״ח ממשלת ישראל, אינפלציה גלומה (BEI) ופורוורדים לפי טווח, אנליטיקה לכל נייר (Carry/Roll-down, תרחישים, ניקוד), דשבורד תיק, השוואת סינתטי מול בולט, היסטוריה מ-2023, ארה״ב ובנק ישראל.", ""),
             ("הנתונים: מאגר CORP (גלובס, יומי; מק״מ נוספו לטבלאות שנמשכות). אין תלות במערכת המסחר: ריבית צבורה, מחיר נקי, מח״מ מתואם וקמירות מחושבים בפייתון (אג״ח בולט, קופון שנתי, ריבית שנתית אפקטיבית, ACT/365, סליקה T+1, אקס-קופון 7 ימים) — אומתו מול מח״מ ומחיר גלובס (הפרש מח״מ חציוני 0.003).", ""),
             ("", ""), ("שלבי עבודה", "h"),
             ("1. 01_הגדרות — תאריך הניתוח (ברירת מחדל: יום הנתונים), צפי אינפלציה לפי טווח (E), מח״מ יעד, מרווח החלטה, תרחישי ריבית. עמודה G מציגה את ה-BEI מהעקום לצד התחזית שלך.", ""),
             ("2. 02_קלט_אגח — הנתונים מהמאגר (אפור). ״לא״ בעמודת ׳כלול בעקום?׳ מוציא נייר מהעקום. ׳כמות נקובה מוחזקת׳ = התיק (אופציונלי).", ""),
             ("3. 03_עקומים — שקלי/ריאלי/BEI/פורוורד/roll-down לפי 12 טווחים, עם גרפים. 04 — תחזית מול BEI, ציפיות בנק ישראל וההזנה לגיליון הקצרות.", ""),
             ("4. 05_אנליטיקה_נייר — כל נייר מול העקום (פער בנ״ב), תשואה צפויה ל-12 חודשים, P&L בתרחישים, ניקוד. 06 — תמונת ועדה. 07 — סינתטי מול בולט.", ""),
             ("5. 09_היסטוריה — איפה העקום וה-BEI מול 3 שנים (אחוזון). 10 — ארה״ב ופער ישראל–ארה״ב. 11 — עקומי האפס של בנק ישראל כבדיקת סבירות.", ""),
             ("", ""), ("איך הקובץ מזין את שאר המערכת", "h"),
             ("הכל נבנה מאותו מאגר (corp-market-data/data/gov): גיליון ׳מחמ מתחת ל 1׳ באקסל העל מקבל את ציפיות האינפלציה לשנה מכאן (BEI לשנה, עם עונתיות חודשית), ההפרש הממשלתי לסדרות הקונצרניות מחושב מאותם עקומים, ומודלי המנפיקים יקבלו את הנחת האינפלציה מאותו מקור (שלב הבא).", ""),
             ("עדכון: הנתונים מתרעננים בענן בכל ערב (data/gov/*.csv). הקובץ עצמו נבנה מחדש בפקודה python gov/build_gov.py data <פלט> --settings <הקובץ הקודם> (שומר פרמטרים, החזקות ובחירות ב-07).", ""),
             ("", ""), ("חיבור חי (אופציונלי, Power Query)", "h"),
             ("בדומה לאקסל העל: לטעון את C:\\CORP\\03_Market\\data\\gov_bonds_latest.csv לגיליון 02 בתא A6 (אותן 22 עמודות, בלי כותרות), והנוסחאות מימין מתעדכנות. הקובץ מסתנכרן ל-C:\\CORP עם שאר נתוני השוק.", ""),
             ("", ""), ("מוסכמות", "h"),
             ("תשואות בגיליונות 03–07 בשבר (0.0308 = 3.08%); ב-02, 09–11 באחוזים. BEI = (1+שקלי)/(1+ריאלי)−1. עקום על שנים לפדיון (לא מח״מ) — כמו במודל המקורי; ההפרש הממשלתי במערכת הקונצרנית מחושב על מח״מ (spreads.py).", ""),
             ("נייר בלי תשואה בגלובס (למשל ממשל שקלית 0530 חדשה) נשאר בטבלה בלי אנליטיקה ולא נכנס לעקום.", "")]
    for txt, kind in lines:
        wg.append([txt]); c = wg.cell(wg.max_row, 1); c.font = Font(name="Arial", size=14 if kind == "t" else 11 if kind == "h" else 10, bold=kind in ("t", "h")); c.alignment = Alignment(wrap_text=True, vertical="top")
    wb.move_sheet("08_מדריך", offset=wb.index(wb["09_היסטוריה"]) - wb.index(wb["08_מדריך"]))
    wb.active = 0
    wb.save(out); print("saved", out, len(bonds), "ניירות", f"{date:%d/%m/%Y}")


if __name__ == "__main__":
    main()
