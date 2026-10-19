#!/usr/bin/env python3
"""build_master.py - בונה את אקסל העל v1 (תמונת מצב): קורא Export מקבצי המנפיקים + data/spreads.
שימוש: python build_master.py <תיקיית מנפיקים> <spreads.csv> <פלט.xlsx>
הנתונים נטענים כערכים (תמונת מצב). לחיבור חי — קוד Power Query בגיליון 'מדריך'."""
import sys, glob, datetime as dt
import pandas as pd, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.formatting.rule import CellIsRule, FormulaRule

iss_dir, sp_csv, out = sys.argv[1:4]
sec_csv = sys.argv[4] if len(sys.argv) > 4 else None
NROW = 600            # מקסימום סדרות בהשוואה
DMAX = 3000           # מקסימום שורות בגיליונות נתונים (לטווחי הנוסחאות)

# ---------- קריאה ----------
issuers, series = [], []
for f in sorted(glob.glob(f"{iss_dir}/**/*.xlsx", recursive=True)):
    if "_archive" in f or "/_build" in f.replace("\\", "/") or "~$" in f or f.endswith("_raw.xlsx"): continue
    wb = openpyxl.load_workbook(f, data_only=True)
    if "Export" not in wb.sheetnames: continue
    ws = wb["Export"]
    def tbl(n):
        rows = list(ws[ws.tables[n].ref]); h = [c.value for c in rows[0]]
        return h, [[c.value for c in r] for r in rows[1:]]
    h, r = tbl("tblIssuer"); issuers_h = h; issuers += r
    h, r = tbl("tblSeries"); series_h = h; series += [x for x in r if x[1] not in (None, "")]
    print("נקלט:", f, len(r), "סדרות")
mk = pd.read_csv(sp_csv, dtype={"security_id": str}, encoding="utf-8-sig")
last = mk["date"].max(); prev_dates = sorted(mk["date"].unique())
prevd = prev_dates[-2] if len(prev_dates) > 1 else None
cur = mk[mk.date == last].copy()
if prevd:
    pv = mk[mk.date == prevd][["security_id", "spread"]].rename(columns={"spread": "spread_prev"})
    cur = cur.merge(pv, on="security_id", how="left")
else:
    cur["spread_prev"] = None
cur["security_id"] = pd.to_numeric(cur["security_id"], errors="coerce")
# מחיר/שינוי יומי מקובץ השוק היומי
daily_path = sp_csv.replace("spreads", "daily").rsplit("/", 1)[0] + f"/{last[:4]}/{last}.csv"
try:
    dd = pd.read_csv(daily_path, dtype={"security_id": str}, encoding="utf-8-sig")
    dd["security_id"] = pd.to_numeric(dd["security_id"], errors="coerce")
    cur = cur.merge(dd[["security_id", "last_price", "chg_pct", "volume_k_nis", "maturity"]], on="security_id", how="left")
except Exception as e:
    print("אזהרה: אין קובץ יומי", e)
    for c in ("last_price", "chg_pct", "volume_k_nis", "maturity"): cur[c] = None

# ---------- סגנונות ----------
NAVY = "1F3864"; HDR = PatternFill("solid", fgColor=NAVY); SUB = PatternFill("solid", fgColor="D9E2F3")
INP = PatternFill("solid", fgColor="FFF2CC"); WH = Font(bold=True, color="FFFFFF", name="Arial", size=10)
BASE = Font(name="Arial", size=10); BOLD = Font(name="Arial", size=10, bold=True)
thin = Side(style="thin", color="BFBFBF"); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
def hdr(ws, row, labels, c0=1):
    for i, t in enumerate(labels):
        c = ws.cell(row, c0 + i, t); c.fill = HDR; c.font = WH; c.border = BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
def rtl(ws): ws.sheet_view.rightToLeft = True

wb = openpyxl.Workbook()

# ---------- פרמטרים ----------
wp = wb.active; wp.title = "פרמטרים"; rtl(wp)
wp["A1"] = "פרמטרים והתראות"; wp["A1"].font = Font(name="Arial", size=14, bold=True)
wp["A2"] = "תאים צהובים ניתנים לשינוי. כל ההתראות בגיליון השוואה נגזרות מהסף כאן."; wp["A2"].font = BASE
hdr(wp, 4, ["פרמטר", "ערך", "הסבר"])
params = [
 ("p_headroom", "Headroom מינימלי בקובננטים", 0.15, "מתחת לערך — התראת ׳קרוב להפרה׳", "0%"),
 ("p_liq12", "כיסוי נזילות 12 חודשים מינימלי", 1.0, "מקורות / שימושים; מתחת ל-1 — התראה", "0.00"),
 ("p_gap", "פער מקסימלי: דירוג פנימי מול סוכנות (נוטצ'ים)", -2, "פנימי נמוך מהסוכנות בלפחות X נוטצ'ים", "0"),
 ("p_age", "ימים מפרסום הדוח — מקסימום", 120, "דוח ישן יותר — התראת עדכניות", "0"),
 ("p_dsp", "הרחבת מרווח יומית (נ״ב)", 25, "שינוי מרווח מול יום המסחר הקודם, בנקודות בסיס", "0"),
 ("p_excess", "מרווח עודף מול דירוג (נ״ב)", 100, "מעל חציון הדירוג הפנימי באותה קבוצה — ׳זול׳ מעניין; מתחת ל- (−X) — ׳יקר׳", "0"),
 ("p_ndcap", "חוב נטו / CAP — סף אזהרה", 0.70, "מעל — התראת מינוף", "0%"),
 ("p_minvol", "מחזור יומי מינימלי (אלפי ש״ח)", 50, "מתחת — מסומן ׳נזילות נמוכה׳ בהשוואה", "0"),
]
PR = {}
for i, (key, lab, val, note, fmt) in enumerate(params):
    r = 5 + i; wp[f"A{r}"] = lab; c = wp[f"B{r}"]; c.value = val; c.fill = INP; c.number_format = fmt; c.border = BOX
    wp[f"C{r}"] = note; PR[key] = f"פרמטרים!$B${r}"
    for col in "AC": wp[f"{col}{r}"].font = BASE
wp.column_dimensions["A"].width = 46; wp.column_dimensions["B"].width = 12; wp.column_dimensions["C"].width = 70

# ---------- נתונים ----------
def dump(name, headers, rows, fmt=None):
    ws = wb.create_sheet(name); rtl(ws)
    hdr(ws, 1, headers)
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            if isinstance(v, float) and pd.isna(v): v = None
            if v is not None and not isinstance(v, (int, float, str, dt.datetime, dt.date)): v = str(v)
            c = ws.cell(2 + i, 1 + j, v); c.font = BASE
            if isinstance(v, dt.datetime): c.number_format = "dd/mm/yyyy"
    ws.freeze_panes = "A2"
    for j in range(len(headers)): ws.column_dimensions[L(j + 1)].width = 16
    t = Table(displayName={"נתוני_מנפיקים": "tblIssuer", "נתוני_סדרות": "tblSeries", "נתוני_שוק": "tblMarket"}[name],
              ref=f"A1:{L(len(headers))}{max(2, len(rows) + 1)}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False); ws.add_table(t)
    return ws
dump("נתוני_מנפיקים", issuers_h, issuers); dump("נתוני_סדרות", series_h, series)
mk_cols = ["date", "source", "security_id", "name", "last_price", "chg_pct", "volume_k_nis", "duration", "ytm_gross",
           "gov_ytm", "spread", "spread_prev", "extrap", "curve", "maturity", "rating_maalot", "rating_midroog"]
for c in mk_cols:
    if c not in cur.columns: cur[c] = None
dump("נתוני_שוק", mk_cols, cur[mk_cols].where(pd.notna(cur[mk_cols]), None).values.tolist())
CI = {h: L(i + 1) for i, h in enumerate(issuers_h)}; CS = {h: L(i + 1) for i, h in enumerate(series_h)}
CM = {h: L(i + 1) for i, h in enumerate(mk_cols)}
def IS(h): return f"נתוני_מנפיקים!${CI[h]}$2:${CI[h]}${DMAX}"
def SE(h): return f"נתוני_סדרות!${CS[h]}$2:${CS[h]}${DMAX}"
def MK(h): return f"נתוני_שוק!${CM[h]}$2:${CM[h]}${DMAX}"


# ---------- מפת ענפים ----------
DEF_SEC = {"RE_IL": ("נדל״ן", "נדל״ן מניב"), "RE_US": ("נדל״ן", "נדל״ן מניב"), "DEV": ("נדל״ן", "נדל״ן מגורים (ייזום)"),
           "IND": ("תעשייה ומסחר", "תעשייה ומסחר"), "ENR": ("אנרגיה ותשתיות", "אנרגיה ותשתיות"),
           "HLD": ("אחזקות", "אחזקות"), "NBC": ("אשראי חוץ בנקאי", "אשראי חוץ בנקאי")}
usr = {}
if sec_csv:
    try:
        sd = pd.read_csv(sec_csv, encoding="utf-8-sig")
        for _, rw in sd.iterrows(): usr[int(rw.iloc[0])] = (str(rw.iloc[1]), str(rw.iloc[2]))
    except Exception as e: print("אזהרה: sectors.csv", e)
ci0 = {h: i for i, h in enumerate(issuers_h)}
secmap = []
for row in issuers:
    iid = int(row[ci0["issuer_id"]]); g = row[ci0["group_code"]]
    sec, sub = usr.get(iid, DEF_SEC.get(g, ("לא מסווג", "לא מסווג")))
    secmap.append([iid, row[ci0["issuer_name"]], g, sec, sub])
SMAX = 305
wsm = wb.create_sheet("מפת_ענפים"); rtl(wsm)
wsm["A1"] = "מפת ענפים ותתי-ענפים"; wsm["A1"].font = Font(name="Arial", size=14, bold=True)
wsm["A2"] = "תאים צהובים: ענף ותת-ענף לכל מנפיק. תתי-הענפים קובעים את קבוצת ההשוואה (חציונים). לשינוי קבוע ערכו את sectors.csv בתיקיית 00_Master ובנו מחדש."; wsm["A2"].font = BASE
hdr(wsm, 4, ["מס׳ מנפיק", "מנפיק", "קוד קבוצה", "ענף", "תת-ענף"])
for k in range(SMAX - 4):
    r = 5 + k
    for j in range(5):
        v = secmap[k][j] if k < len(secmap) else None
        c = wsm.cell(r, 1 + j, v); c.font = BASE
        if j >= 3: c.fill = INP; c.border = BOX
secs = ["הכל"] + sorted({m[3] for m in secmap}); subs = ["הכל"] + sorted({m[4] for m in secmap})
wsm["G4"] = "רשימת ענפים"; wsm["H4"] = "רשימת תתי-ענפים"
for c in ("G4", "H4"): wsm[c].fill = HDR; wsm[c].font = WH
for i, v in enumerate(secs): wsm.cell(5 + i, 7, v).font = BASE
for i, v in enumerate(subs): wsm.cell(5 + i, 8, v).font = BASE
for col, w in zip("ABCDEFGH", (12, 30, 10, 20, 24, 3, 18, 24)): wsm.column_dimensions[col].width = w
def SECLK(col, idexpr): return f'IFERROR(INDEX(מפת_ענפים!${col}$5:${col}${SMAX},MATCH({idexpr},מפת_ענפים!$A$5:$A${SMAX},0)),"לא מסווג")'

# ---------- השוואה ----------
wc = wb.create_sheet("השוואה_אגחים", 0); rtl(wc)
wc["A1"] = "אקסל על — השוואת סדרות אג״ח קונצרני"; wc["A1"].font = Font(name="Arial", size=14, bold=True)
wc["A2"] = f"נתוני שוק: {last}  |  גרסת מודל M1.0  |  שורה לכל סדרה שיש לה מודל מנפיק. מסננים בכותרת."; wc["A2"].font = BASE
cols = [  # כותרת, מקור, פורמט, רוחב
 ("מס׳ ני״ע", "sid", "0", 11), ("מנפיק", "iss:issuer_name", None, 26), ("ענף", "sec:D", None, 14), ("תת-ענף", "sec:E", None, 18), ("סדרה", "ser:series_name", None, 18),
 ("מח״מ", "mk:duration", "0.00", 7), ("תשואה ברוטו %", "mk:ytm_gross", "0.00", 9), ("הפרש ממשלתי (נ״ב)", "sp_bp", "0", 10),
 ("הצמדה", "ser:linkage", None, 11), ("פירעון סופי", "ser:maturity", "mm/yyyy", 10),
 ("דירוג פנימי S&P", "ser:internal_sp", None, 9), ("דירוג פנימי Moody's", "ser:internal_moodys", None, 9),
 ("דירוג סוכנות", "agency", None, 10), ("פער פנימי–סוכנות", "ser:internal_vs_agency", "0", 9),
 ("חוב נטו / EBITDA", "iss:nd_ebitda", "0.0", 9), ("חוב נטו / CAP", "iss:nd_cap", "0.0%", 9),
 ("EBITDA / ריבית", "iss:ebitda_int", "0.00", 9), ("כיסוי נזילות 12 ח׳", "iss:liq12", "0.00", 9),
 ("Headroom מינימלי", "ser:min_cov_headroom", "0%", 9), ("בטוחה", "ser:collateral_type", None, 14),
 ("מחיר", "mk:last_price", "0.00", 8), ("שינוי יומי %", "mk:chg_pct", "0.00", 8), ("מחזור (אלפי ₪)", "mk:volume_k_nis", "#,##0", 9),

 ("תשואה ממשלתי באותו מח״מ", "mk:gov_ytm", "0.00", 10),
 ("שינוי מרווח יומי (נ״ב)", "dsp_bp", "0", 10), ("חציון קבוצת הדירוג (נ״ב)", "peer", "0", 11), ("מרווח עודף (נ״ב)", "excess", "0", 10),
 ("חציון תת-ענף (נ״ב)", "subpeer", "0", 11), ("מרווח מול תת-ענף (נ״ב)", "subex", "0", 11),
 ("התראות", "alerts", None, 60),
]
hdr(wc, 4, [c[0] for c in cols]); wc.row_dimensions[4].height = 42
CL = {c[1]: L(i + 1) for i, c in enumerate(cols)}
for i, c in enumerate(cols): wc.column_dimensions[L(i + 1)].width = c[3]
def lk(rng, key, r): return f"MATCH({key},{rng},0)"
for k in range(NROW):
    r = 5 + k; sr = 2 + k
    sid = f"$A{r}"
    f = {}
    f["sid"] = f'=IF(נתוני_סדרות!{CS["security_id"]}{sr}="","",נתוני_סדרות!{CS["security_id"]}{sr})'
    ser_match = f'MATCH({sid},{SE("security_id")},0)'
    ISS_ID = f'INDEX({SE("issuer_id")},{ser_match})'
    iss_match = f'MATCH(INDEX({SE("issuer_id")},{ser_match}),{IS("issuer_id")},0)'
    mk_match = f'MATCH({sid},{MK("security_id")},0)'
    for hd, src, fmt, w in cols:
        if src == "sid": continue
        col = CL[src]
        if src.startswith("ser:"): f[src] = f'=IF({sid}="","",IFERROR(INDEX({SE(src[4:])},{ser_match})&"","")*1)' if False else f'=IF({sid}="","",IFERROR(IF(INDEX({SE(src[4:])},{ser_match})="","",INDEX({SE(src[4:])},{ser_match})),""))'
        elif src.startswith("iss:"): f[src] = f'=IF({sid}="","",IFERROR(IF(INDEX({IS(src[4:])},{iss_match})="","",INDEX({IS(src[4:])},{iss_match})),""))'
        elif src.startswith("sec:"): f[src] = f'=IF({sid}="","",{SECLK(src[4:], ISS_ID)})'
        elif src.startswith("mk:"): f[src] = f'=IF({sid}="","",IFERROR(IF(INDEX({MK(src[3:])},{mk_match})="","",INDEX({MK(src[3:])},{mk_match})),""))'
    f["agency"] = (f'=IF({sid}="","",IFERROR(IF(INDEX({SE("maalot")},{ser_match})<>"",INDEX({SE("maalot")},{ser_match}),'
                   f'IF(INDEX({SE("midroog")},{ser_match})<>"",INDEX({SE("midroog")},{ser_match}),"")),""))')
    sp = CL["mk:ytm_gross"]
    f["sp_bp"] = f'=IF({sid}="","",IFERROR(INDEX({MK("spread")},{mk_match})*100,""))'
    f["dsp_bp"] = f'=IF({sid}="","",IFERROR(IF(INDEX({MK("spread_prev")},{mk_match})="","",(INDEX({MK("spread")},{mk_match})-INDEX({MK("spread_prev")},{mk_match}))*100),""))'
    f["peer"] = f'=IF(OR({sid}="",{CL["ser:internal_sp"]}{r}=""),"",IFERROR(INDEX(\'דירוג_פרמיה\'!$C$5:$C$40,MATCH({CL["ser:internal_sp"]}{r},\'דירוג_פרמיה\'!$A$5:$A$40,0)),""))'
    SPR = f'${CL["sp_bp"]}$5:${CL["sp_bp"]}${4+NROW}'; SUBR = f'${CL["sec:E"]}$5:${CL["sec:E"]}${4+NROW}'
    f["subpeer"] = None; f["subex"] = f'=IF(OR({CL["sp_bp"]}{r}="",{CL["subpeer"]}{r}=""),"",{CL["sp_bp"]}{r}-{CL["subpeer"]}{r})'
    f["excess"] = f'=IF(OR({CL["sp_bp"]}{r}="",{CL["peer"]}{r}=""),"",{CL["sp_bp"]}{r}-{CL["peer"]}{r})'
    # התראות — שרשור טקסט
    age = f'(TODAY()-INDEX({IS("report_date")},{iss_match}))'
    parts = [
      f'IF(INDEX({IS("data_status")},{iss_match})="שגיאה","נתונים בשגיאה; ","")',
      f'IF(AND({CL["ser:min_cov_headroom"]}{r}<>"",N({CL["ser:min_cov_headroom"]}{r})<{PR["p_headroom"]}),"Headroom נמוך; ","")',
      f'IF(N(INDEX({IS("cov_breaches")},{iss_match}))>0,"הפרת קובננט; ","")',
      f'IF(AND({CL["iss:liq12"]}{r}<>"",N({CL["iss:liq12"]}{r})<{PR["p_liq12"]}),"נזילות 12 ח׳ <1; ","")',
      f'IF(AND({CL["ser:internal_vs_agency"]}{r}<>"",N({CL["ser:internal_vs_agency"]}{r})<={PR["p_gap"]}),"פער דירוג; ","")',
      f'IF({age}>{PR["p_age"]},"דוח ישן; ","")',
      f'IF(AND({CL["iss:nd_cap"]}{r}<>"",N({CL["iss:nd_cap"]}{r})>{PR["p_ndcap"]}),"מינוף גבוה; ","")',
      f'IF(AND({CL["dsp_bp"]}{r}<>"",N({CL["dsp_bp"]}{r})>={PR["p_dsp"]}),"הרחבת מרווח; ","")',
      f'IF(AND({CL["excess"]}{r}<>"",N({CL["excess"]}{r})>={PR["p_excess"]}),"מרווח עודף (זול); ","")',
      f'IF(AND({CL["excess"]}{r}<>"",N({CL["excess"]}{r})<=-{PR["p_excess"]}),"מרווח נמוך (יקר); ","")',
      f'IF(AND({CL["mk:volume_k_nis"]}{r}<>"",N({CL["mk:volume_k_nis"]}{r})<{PR["p_minvol"]}),"סחירות נמוכה; ","")',
    ]
    f["alerts"] = f'=IF({sid}="","",IFERROR({"&".join(parts)},""))'
    for i, (hd, src, fmt, w) in enumerate(cols):
        c = wc.cell(r, i + 1, f[src]); c.font = BASE
        if fmt: c.number_format = fmt
    pc = CL["subpeer"]
    wc[f"{pc}{r}"] = ArrayFormula(f"{pc}{r}", f'=IF(OR({sid}="",{CL["sec:E"]}{r}=""),"",IFERROR(MEDIAN(IF(({SUBR}={CL["sec:E"]}{r})*ISNUMBER({SPR}),{SPR})),""))')
    wc[f"{pc}{r}"].number_format = "0"; wc[f"{pc}{r}"].font = BASE
wc.freeze_panes = "D5"; wc.auto_filter.ref = f"A4:{L(len(cols))}{4 + NROW}"
red = PatternFill("solid", fgColor="F8CBAD")
wc.conditional_formatting.add(f"{CL['alerts']}5:{CL['alerts']}{4+NROW}", FormulaRule(formula=[f'LEN({CL["alerts"]}5)>0'], fill=red))
for key in ("ser:min_cov_headroom",):
    wc.conditional_formatting.add(f"{CL[key]}5:{CL[key]}{4+NROW}", FormulaRule(formula=[f'AND({CL[key]}5<>"",{CL[key]}5<{PR["p_headroom"]})'], fill=red))

# ---------- דירוג_פרמיה ----------
wr = wb.create_sheet("דירוג_פרמיה", 1); rtl(wr)
wr["A1"] = "מרווח ממשלתי חציוני לפי דירוג פנימי (S&P)"; wr["A1"].font = Font(name="Arial", size=14, bold=True)
wr["A2"] = "חציון על פני סדרות עם מודל בלבד. מתחת ל-3 סדרות בקבוצה — החציון חלש; להסתמך רק כשיש מספיק מנפיקים."; wr["A2"].font = BASE
hdr(wr, 4, ["דירוג פנימי", "מס׳ סדרות", "חציון מרווח (נ״ב)", "מינימום", "מקסימום"])
scale = ["AAA","AA+","AA","AA-","A+","A","A-","BBB+","BBB","BBB-","BB+","BB","BB-","B+","B","B-","CCC+","CCC","CCC-","CC","C","D"]
rg = lambda k: f"השוואה_אגחים!${CL[k]}$5:${CL[k]}${4+NROW}"
for i, g in enumerate(scale):
    r = 5 + i; wr[f"A{r}"] = g; wr[f"A{r}"].font = BOLD
    wr[f"B{r}"] = f'=COUNTIFS({rg("ser:internal_sp")},A{r},{rg("sp_bp")},">-9999")'
    wr[f"C{r}"] = ArrayFormula(f"C{r}", f'=IF(B{r}=0,"",MEDIAN(IF(({rg("ser:internal_sp")}=A{r})*ISNUMBER({rg("sp_bp")}),{rg("sp_bp")})))')
    wr[f"D{r}"] = ArrayFormula(f"D{r}", f'=IF(B{r}=0,"",MIN(IF(({rg("ser:internal_sp")}=A{r})*ISNUMBER({rg("sp_bp")}),{rg("sp_bp")})))')
    wr[f"E{r}"] = ArrayFormula(f"E{r}", f'=IF(B{r}=0,"",MAX(IF(({rg("ser:internal_sp")}=A{r})*ISNUMBER({rg("sp_bp")}),{rg("sp_bp")})))')
    for col in "BCDE":
        wr[f"{col}{r}"].font = BASE; wr[f"{col}{r}"].number_format = "0" if col != "B" else "0"
for col, w in zip("ABCDE", (14, 12, 18, 12, 12)): wr.column_dimensions[col].width = w

# ---------- סינון_שוק ----------
wm = wb.create_sheet("סינון_שוק", 2); rtl(wm)
wm["A1"] = "כל האג״ח הקונצרניות בשוק — מרווח ממשלתי ודירוג סוכנות"; wm["A1"].font = Font(name="Arial", size=14, bold=True)
wm["A2"] = "כולל סדרות שאין להן עדיין מודל מנפיק. ׳יש מודל׳ מסמן סדרות שנמצאות בהשוואה."; wm["A2"].font = BASE
mh = ["מס׳ ני״ע", "סדרה", "סוג", "מח״מ", "תשואה ברוטו %", "ממשלתי באותו מח״מ", "הפרש (נ״ב)", "שינוי יומי (נ״ב)", "מחיר", "מחזור (אלפי ₪)",
      "פירעון", "מעלות", "מידרוג", "יש מודל", "מחוץ לטווח העקום", "מצוקה"]
hdr(wm, 4, mh); wm.row_dimensions[4].height = 32
src_type = {"corp_cpi": "צמוד", "corp_unlinked": "שקלי"}
for k in range(len(cur)):
    r = 5 + k; sr = 2 + k
    def m(h): return f"נתוני_שוק!{CM[h]}{sr}"
    vals = [f"={m('security_id')}", f"={m('name')}", f'=IF({m("source")}="corp_cpi","צמוד","שקלי")', f"={m('duration')}", f"={m('ytm_gross')}",
            f"={m('gov_ytm')}", f"={m('spread')}*100", f'=IF({m("spread_prev")}="","",({m("spread")}-{m("spread_prev")})*100)',
            f'=IF({m("last_price")}="","",{m("last_price")})', f'=IF({m("volume_k_nis")}="","",{m("volume_k_nis")})', f'=IF({m("maturity")}="","",{m("maturity")})',
            f'=IF({m("rating_maalot")}="","",{m("rating_maalot")})', f'=IF({m("rating_midroog")}="","",{m("rating_midroog")})',
            f'=IF(COUNTIF({SE("security_id")},A{r})>0,"כן","")', f'=IF({m("extrap")}=1,"כן","")', f'=IF(OR({m("ytm_gross")}>12,G{r}>600),"כן","")']
    fm = ["0", None, None, "0.00", "0.00", "0.00", "0", "0", "0.00", "#,##0", "mm/yyyy", None, None, None, None, None]
    for j, v in enumerate(vals):
        c = wm.cell(r, 1 + j, v); c.font = BASE
        if fm[j]: c.number_format = fm[j]
for j, w in enumerate([11, 22, 7, 7, 10, 11, 10, 10, 8, 10, 9, 9, 9, 8, 10, 7]): wm.column_dimensions[L(j + 1)].width = w
wm.freeze_panes = "C5"; wm.auto_filter.ref = f"A4:{L(len(mh))}{4+len(cur)}"
wm.conditional_formatting.add(f"A5:P{4+len(cur)}", FormulaRule(formula=['$N5="כן"'], fill=SUB))

# ---------- Watchlist ----------
ww = wb.create_sheet("Watchlist", 3); rtl(ww)
ww["A1"] = "Watchlist"; ww["A1"].font = Font(name="Arial", size=14, bold=True)
ww["A2"] = "מזינים מס׳ ני״ע (תא צהוב) והערה. השורה נמשכת אוטומטית מהשוואה, ואם אין מודל — מנתוני השוק."; ww["A2"].font = BASE
wh = ["מס׳ ני״ע", "הערה / סיבה", "סדרה", "מנפיק", "דירוג פנימי", "דירוג סוכנות", "הפרש (נ״ב)", "מרווח עודף (נ״ב)", "שינוי יומי (נ״ב)", "תשואה %", "מח״מ", "התראות"]
hdr(ww, 4, wh)
seed = [(int(s[1]), "דוגמה — אדגר") for s in series[:4]]
for k in range(60):
    r = 5 + k
    a = ww.cell(r, 1, seed[k][0] if k < len(seed) else None); a.fill = INP; a.border = BOX; a.font = BASE
    b = ww.cell(r, 2, seed[k][1] if k < len(seed) else None); b.fill = INP; b.border = BOX; b.font = BASE
    cm = f"MATCH($A{r},השוואה_אגחים!$A$5:$A${4+NROW},0)"; mm = f"MATCH($A{r},{MK('security_id')},0)"
    def cmp(key): return f'IFERROR(INDEX(השוואה_אגחים!${CL[key]}$5:${CL[key]}${4+NROW},{cm}),"")'
    ww.cell(r, 3, f'=IF($A{r}="","",IFERROR(INDEX({MK("name")},{mm}),{cmp("ser:series_name")}))')
    ww.cell(r, 4, f'=IF($A{r}="","",{cmp("iss:issuer_name")})')
    ww.cell(r, 5, f'=IF($A{r}="","",{cmp("ser:internal_sp")})')
    ww.cell(r, 6, f'=IF($A{r}="","",IF({cmp("agency")}<>"",{cmp("agency")},IFERROR(IF(INDEX({MK("rating_maalot")},{mm})<>"",INDEX({MK("rating_maalot")},{mm}),INDEX({MK("rating_midroog")},{mm})),"")))')
    ww.cell(r, 7, f'=IF($A{r}="","",IFERROR(INDEX({MK("spread")},{mm})*100,""))').number_format = "0"
    ww.cell(r, 8, f'=IF($A{r}="","",{cmp("excess")})').number_format = "0"
    ww.cell(r, 9, f'=IF($A{r}="","",{cmp("dsp_bp")})').number_format = "0"
    ww.cell(r, 10, f'=IF($A{r}="","",IFERROR(INDEX({MK("ytm_gross")},{mm}),""))').number_format = "0.00"
    ww.cell(r, 11, f'=IF($A{r}="","",IFERROR(INDEX({MK("duration")},{mm}),""))').number_format = "0.00"
    ww.cell(r, 12, f'=IF($A{r}="","",{cmp("alerts")})')
    for c in range(3, 13): ww.cell(r, c).font = BASE
for j, w in enumerate([11, 26, 20, 26, 9, 10, 10, 11, 11, 9, 8, 60]): ww.column_dimensions[L(j + 1)].width = w
ww.freeze_panes = "C5"

# ---------- השוואה_מנפיקים ----------
NI = 300
wi = wb.create_sheet("השוואה_מנפיקים", 0); rtl(wi)
wi["A1"] = "אקסל על — השוואת מנפיקים לפי ענף ותת-ענף"; wi["A1"].font = Font(name="Arial", size=14, bold=True)
wi["A2"] = "שורה לכל מנפיק שנבדק. חוב ונתונים פיננסיים ב-₪ מיליון. חציונים לפי תת-ענף (מתחת ל-3 מנפיקים — חלש). מסננים בכותרת."; wi["A2"].font = BASE
icols = [  # כותרת, מפתח, פורמט, רוחב
 ("ענף", "sec", None, 14), ("תת-ענף", "sub", None, 18), ("מנפיק", "iss:issuer_name", None, 28), ("מס׳ סדרות", "nser", "0", 7),
 ("יתרת חוב סדרות (פארי)", "par", "#,##0", 10),
 ("דירוג פנימי S&P", "iss:rating_sp", None, 8), ("דירוג פנימי Moody's", "iss:rating_moodys", None, 8), ("דירוג סוכנות", "agency", None, 9),
 ("פער פנימי–סוכנות", "iss:internal_vs_agency", "0", 8), ("תחזית סוכנות", "iss:maalot_outlook", None, 8),
 ("חוב נטו", "iss:net_debt_ils", "#,##0", 9), ("הון", "iss:equity_ils", "#,##0", 9), ("סך נכסים", "iss:total_assets_ils", "#,##0", 9),
 ("EBITDA LTM", "iss:ebitda_ils", "#,##0", 9),
 ("חוב נטו / EBITDA", "iss:nd_ebitda", "0.0", 8), ("חוב נטו / CAP", "iss:nd_cap", "0.0%", 8), ("EBITDA / ריבית", "iss:ebitda_int", "0.00", 8),
 ("FFO / חוב נטו", "iss:ffo_nd", "0.0%", 8), ("כיסוי נזילות 12 ח׳", "iss:liq12", "0.00", 8), ("כיסוי נזילות 24 ח׳", "iss:liq24", "0.00", 8),
 ("מח״מ חוב (שנים)", "iss:wam_years", "0.0", 8),
 ("גיוס נדרש 3 שנים (בסיס)", "iss:fc_need_3y", "#,##0", 9), ("שנת גיוס ראשונה (בסיס)", "iss:fc_first_need_year", '0;-0;"אין"', 8),
 ("נזילות מתחת למינימום בלי גיוס — שנה (בסיס)", "iss:fc_runout_year", '0;-0;"אין"', 8),
 ("כיסוי פנימי 3 שנים (הגנה)", "iss:fc_selfcov_3y", "0%", 8), ("צורך שנתי / קצב גיוס היסטורי (לחץ)", "iss:fc_need_hist_3y", "0%", 9),
 ("צורך 3 שנים / יכולת גיוס מובטח נוסף (לחץ)", "iss:fc_need_cap_3y", "0%", 9),
 ("Headroom מינימלי", "iss:min_cov_headroom", "0%", 8), ("הפרות קובננט", "iss:cov_breaches", "0", 7),
 ("תאריך דוח", "iss:report_date", "dd/mm/yy", 9),
 ("מס׳ מנפיקים בתת-ענף", "grpn", "0", 8),
 ("חציון תת-ענף: חוב נטו / EBITDA", "med_e", "0.0", 10), ("חציון תת-ענף: חוב נטו / CAP", "med_c", "0.0%", 10), ("חציון תת-ענף: EBITDA / ריבית", "med_i", "0.00", 10),
 ("חוב נטו/CAP מול חציון (נק׳ אחוז)", "d_cap", "0.0", 10), ("דירוג בתת-ענף (1 = מינוף נמוך)", "rank", "0", 9),
 ("התראות", "alerts", None, 55), ("_עזר", "helper", "0", 6),
]
hdr(wi, 4, [c[0] for c in icols]); wi.row_dimensions[4].height = 54
IC = {c[1]: L(i + 1) for i, c in enumerate(icols)}
for i, c in enumerate(icols): wi.column_dimensions[L(i + 1)].width = c[3]
irng = lambda k: f"${IC[k]}$5:${IC[k]}${4+NI}"
SELS, SELB = "השוואה_ענפית!$C$3", "השוואה_ענפית!$C$4"
for k in range(NI):
    r = 5 + k; sr = 2 + k
    idc = f"נתוני_מנפיקים!{CI['issuer_id']}{sr}"
    f = {}
    def g(h): return f'IF(nתוני_מנפיקים!{CI[h]}{sr}="","",נתוני_מנפיקים!{CI[h]}{sr})'.replace("nתוני", "נתוני")
    for hd, key, fmt, w in icols:
        if key.startswith("iss:"): f[key] = f'=IF({idc}="","",{g(key[4:])})'
    f["sec"] = f'=IF({idc}="","",{SECLK("D", idc)})'; f["sub"] = f'=IF({idc}="","",{SECLK("E", idc)})'
    f["nser"] = f'=IF({idc}="","",COUNTIF({SE("issuer_id")},{idc}))'
    f["par"] = f'=IF({idc}="","",SUMIF({SE("issuer_id")},{idc},{SE("par_m")}))'
    f["agency"] = f'=IF({idc}="","",IF(נתוני_מנפיקים!{CI["maalot"]}{sr}<>"",נתוני_מנפיקים!{CI["maalot"]}{sr},IF(נתוני_מנפיקים!{CI["midroog"]}{sr}<>"",נתוני_מנפיקים!{CI["midroog"]}{sr},"")))'
    f["grpn"] = f'=IF({idc}="","",COUNTIF({irng("sub")},{IC["sub"]}{r}))'
    for key, base in (("med_e", "iss:nd_ebitda"), ("med_c", "iss:nd_cap"), ("med_i", "iss:ebitda_int")):
        f[key] = None
    f["d_cap"] = f'=IF(OR({IC["iss:nd_cap"]}{r}="",{IC["med_c"]}{r}=""),"",({IC["iss:nd_cap"]}{r}-{IC["med_c"]}{r})*100)'
    f["rank"] = f'=IF(OR({idc}="",{IC["iss:nd_cap"]}{r}=""),"",1+COUNTIFS({irng("sub")},{IC["sub"]}{r},{irng("iss:nd_cap")},"<"&{IC["iss:nd_cap"]}{r}))'
    age = f'(TODAY()-{IC["iss:report_date"]}{r})'
    parts = [
      f'IF(N({IC["iss:cov_breaches"]}{r})>0,"הפרת קובננט; ","")',
      f'IF(AND({IC["iss:min_cov_headroom"]}{r}<>"",N({IC["iss:min_cov_headroom"]}{r})<{PR["p_headroom"]}),"Headroom נמוך; ","")',
      f'IF(AND({IC["iss:liq12"]}{r}<>"",N({IC["iss:liq12"]}{r})<{PR["p_liq12"]}),"נזילות 12 ח׳ <1; ","")',
      f'IF(AND({IC["iss:internal_vs_agency"]}{r}<>"",N({IC["iss:internal_vs_agency"]}{r})<={PR["p_gap"]}),"פער דירוג; ","")',
      f'IF(AND({IC["iss:nd_cap"]}{r}<>"",N({IC["iss:nd_cap"]}{r})>{PR["p_ndcap"]}),"מינוף גבוה; ","")',
      f'IF(AND({IC["iss:report_date"]}{r}<>"",{age}>{PR["p_age"]}),"דוח ישן; ","")',
      f'IF(AND({IC["d_cap"]}{r}<>"",N({IC["grpn"]}{r})>=3,N({IC["d_cap"]}{r})>=10),"מינוף גבוה מהעמיתים; ","")',
      f'IF(AND({IC["grpn"]}{r}<>"",N({IC["grpn"]}{r})<3),"קבוצת השוואה קטנה; ","")',
      f'IF(N({IC["iss:fc_runout_year"]}{r})=1,"נזילות מתחת למינימום בשנה 1; ","")',
      f'IF(AND({IC["iss:fc_need_hist_3y"]}{r}<>"",N({IC["iss:fc_need_hist_3y"]}{r})>1),"צורך גיוס מעל הקצב ההיסטורי; ","")',
    ]
    f["alerts"] = f'=IF({idc}="","",{"&".join(parts)})'
    f["helper"] = (f'=IF(AND({idc}<>"",OR({SELS}="הכל",{IC["sec"]}{r}={SELS}),OR({SELB}="הכל",{IC["sub"]}{r}={SELB})),'
                   f'N({IC["helper"]}{r-1})+1,"")') if k else f'=IF(AND({idc}<>"",OR({SELS}="הכל",{IC["sec"]}{r}={SELS}),OR({SELB}="הכל",{IC["sub"]}{r}={SELB})),1,"")'
    f["helper"] = f['helper'].replace(f'N({IC["helper"]}{r-1})+1', f'MAX({IC["helper"]}$4:{IC["helper"]}{r-1})+1')
    for i, (hd, key, fmt, w) in enumerate(icols):
        c = wi.cell(r, i + 1, f[key]); c.font = BASE
        if fmt: c.number_format = fmt
    for key, base in (("med_e", "iss:nd_ebitda"), ("med_c", "iss:nd_cap"), ("med_i", "iss:ebitda_int")):
        col = IC[key]
        wi[f"{col}{r}"] = ArrayFormula(f"{col}{r}", f'=IF(OR({idc}="",{IC["sub"]}{r}=""),"",IFERROR(MEDIAN(IF(({irng("sub")}={IC["sub"]}{r})*ISNUMBER({irng(base)}),{irng(base)})),""))')
        wi[f"{col}{r}"].font = BASE; wi[f"{col}{r}"].number_format = [c for c in icols if c[1] == key][0][2]
wi.freeze_panes = "D5"; wi.auto_filter.ref = f"A4:{L(len(icols))}{4+NI}"
wi.conditional_formatting.add(f"{IC['alerts']}5:{IC['alerts']}{4+NI}", FormulaRule(formula=[f'LEN({IC["alerts"]}5)>0'], fill=PatternFill("solid", fgColor="F8CBAD")))
wi.column_dimensions[IC["helper"]].hidden = True

# ---------- השוואה_ענפית ----------
wv = wb.create_sheet("השוואה_ענפית", 1); rtl(wv)
wv["A1"] = "השוואה ענפית — מנפיקים באותו ענף / תת-ענף"; wv["A1"].font = Font(name="Arial", size=14, bold=True)
wv["A2"] = "בוחרים ענף ותת-ענף בתאים הצהובים (׳הכל׳ = ללא סינון). הרשימה נמשכת מהשוואה_מנפיקים, ובתחתיתה חציון הקבוצה."; wv["A2"].font = BASE
for r, lab in ((3, "ענף"), (4, "תת-ענף")):
    wv.cell(r, 2, lab).font = BOLD; c = wv.cell(r, 3, "הכל"); c.fill = INP; c.border = BOX; c.font = BOLD
from openpyxl.worksheet.datavalidation import DataValidation
dv1 = DataValidation(type="list", formula1=f"=מפת_ענפים!$G$5:$G${4+len(secs)}", allow_blank=False)
dv2 = DataValidation(type="list", formula1=f"=מפת_ענפים!$H$5:$H${4+len(subs)}", allow_blank=False)
wv.add_data_validation(dv1); wv.add_data_validation(dv2); dv1.add("C3"); dv2.add("C4")
vkeys = ["sec", "sub", "iss:issuer_name", "nser", "par", "iss:rating_sp", "iss:rating_moodys", "agency", "iss:internal_vs_agency",
         "iss:net_debt_ils", "iss:ebitda_ils", "iss:nd_ebitda", "iss:nd_cap", "iss:ebitda_int", "iss:ffo_nd", "iss:liq12", "iss:liq24",
         "iss:wam_years", "iss:fc_need_3y", "iss:fc_first_need_year", "iss:fc_selfcov_3y", "iss:fc_need_hist_3y", "iss:min_cov_headroom", "rank", "alerts"]
meta = {c[1]: c for c in icols}
hdr(wv, 6, [meta[k][0] for k in vkeys]); wv.row_dimensions[6].height = 42
NV = 40
for i, k in enumerate(vkeys): wv.column_dimensions[L(i + 1)].width = meta[k][3]
for n in range(1, NV + 1):
    r = 6 + n
    mt = f"MATCH({n},'השוואה_מנפיקים'!${IC['helper']}$5:${IC['helper']}${4+NI},0)"
    for i, k in enumerate(vkeys):
        c = wv.cell(r, i + 1, f"=IFERROR(INDEX('השוואה_מנפיקים'!${IC[k]}$5:${IC[k]}${4+NI},{mt}),\"\")"); c.font = BASE
        if meta[k][2]: c.number_format = meta[k][2]
rm = 6 + NV + 1
wv.cell(rm, 3, "חציון הקבוצה").font = BOLD
for i, k in enumerate(vkeys):
    if k in ("iss:net_debt_ils", "iss:ebitda_ils", "iss:nd_ebitda", "iss:nd_cap", "iss:ebitda_int", "iss:ffo_nd", "iss:liq12", "iss:liq24", "iss:wam_years", "par", "iss:fc_selfcov_3y", "iss:fc_need_hist_3y"):
        col = L(i + 1); c = wv.cell(rm, i + 1, f'=IFERROR(MEDIAN({col}7:{col}{6+NV}),"")'); c.font = BOLD; c.number_format = meta[k][2]; c.fill = SUB
wv.freeze_panes = "D7"
wv.conditional_formatting.add(f"{L(len(vkeys))}7:{L(len(vkeys))}{6+NV}", FormulaRule(formula=[f'LEN({L(len(vkeys))}7)>0'], fill=PatternFill("solid", fgColor="F8CBAD")))


# ---------- מה_חדש ----------
upd = []
for f in sorted(glob.glob(f"{iss_dir}/**/_log/updates.csv", recursive=True)):
    try: upd.append(pd.read_csv(f, encoding="utf-8-sig"))
    except Exception as e: print("אזהרה: יומן", f, e)
extra = sys.argv[5] if len(sys.argv) > 5 else None
if extra:
    try: upd.append(pd.read_csv(extra, encoding="utf-8-sig"))
    except Exception as e: print("אזהרה: יומן", e)
wn = wb.create_sheet("מה_חדש", 2); rtl(wn)
wn["A1"] = "מה התחדש — יומן עדכוני מודלים"; wn["A1"].font = Font(name="Arial", size=14, bold=True)
wn["A2"] = "נקרא מ-_log\\updates.csv בתיקיית כל מנפיק. החדש ראשון. ׳טיוטה׳ = ממתין לאישורך."; wn["A2"].font = BASE
hdr(wn, 4, ["תאריך", "מנפיק", "קובץ שנקלט", "סוג", "תקופה", "מה השתנה", "סטטוס"])
if upd:
    U = pd.concat(upd, ignore_index=True).sort_values("date", ascending=False)
    names = {int(r[ci0["issuer_id"]]): r[ci0["issuer_name"]] for r in issuers}
    for i, (_, u) in enumerate(U.iterrows()):
        vals = [u["date"], names.get(int(u["issuer_id"]), u["issuer_id"]), u["file"], u["doc_type"], u["period"], u["what_changed"], u["status"]]
        for j, v in enumerate(vals):
            c = wn.cell(5 + i, 1 + j, None if pd.isna(v) else v); c.font = BASE; c.alignment = Alignment(wrap_text=True, vertical="top")
for col, w in zip("ABCDEFG", (11, 26, 36, 9, 9, 70, 9)): wn.column_dimensions[col].width = w
wn.freeze_panes = "A5"
# ---------- מדריך ----------
wg = wb.create_sheet("מדריך", 4); rtl(wg)
wg.column_dimensions["A"].width = 120
lines = [
 ("אקסל על v1 — מדריך", "t"),
 ("גרסה: v1 (תמונת מצב). הנתונים בגיליונות נתוני_* נטענו כערכים בעת הבנייה. לחיבור חי — שלוש שאילתות Power Query למטה.", ""),
 ("", ""),
 ("מבנה", "h"),
 ("השוואה_מנפיקים — שורה לכל מנפיק: ענף, תת-ענף, דירוגים, חוב ונתונים פיננסיים, יחסי אשראי, חציון תת-ענף, דירוג בתת-ענף והתראות.", ""),
 ("השוואה_ענפית — בוחרים ענף/תת-ענף ורואים רק את המנפיקים באותה קבוצה, עם חציון.", ""),
 ("מה_חדש — יומן עדכונים: איזה דוח נקלט, מה השתנה במודל, וסטטוס אישור.", ""),
 ("מפת_ענפים — שיוך כל מנפיק לענף ולתת-ענף (נקבע מ-sectors.csv).", ""),
 ("השוואה_אגחים — שורה לכל סדרה עם מודל מנפיק: דירוג פנימי/סוכנות, מדדי אשראי, שוק, הפרש ממשלתי, מרווח עודף מול חציון הדירוג, התראות.", ""),
 ("דירוג_פרמיה — חציון המרווח לכל דירוג פנימי. נותן משמעות רק עם מספר מנפיקים בקבוצה.", ""),
 ("סינון_שוק — כל האג״ח הקונצרניות עם הפרש ממשלתי ודירוגי סוכנות, כולל סדרות בלי מודל. מסננים לפי דירוג, מח״מ, הפרש.", ""),
 ("Watchlist — מזינים מספר ני״ע; מושך את ההשוואה והשוק.", ""),
 ("פרמטרים — סף כל ההתראות.", ""),
 ("", ""),
 ("שיטת המרווח", "h"),
 ("הפרש = תשואה ברוטו של האג״ח פחות תשואה ברוטו של ממשלתי באותו מח״מ (אינטרפולציה ליניארית). צמוד מול עקום צמוד, שקלי מול עקום שקלי קבוע. מחושב ב-GitHub (spreads.py).", ""),
 ("מרווח עודף = ההפרש פחות חציון ההפרש של אותו דירוג פנימי. חיובי — האג״ח נסחרת מעל מה שהדירוג הפנימי מצדיק (מעניינת); שלילי — יקרה יחסית.", ""),
 ("הערה: ההפרש מוצג בנקודות בסיס. סדרות במצוקה (תשואה >12% או הפרש >600 נ״ב) מסומנות ׳מצוקה׳ בסינון_שוק ולא לבסס עליהן השוואה.", ""),
 ("", ""),
 ("חיבור חי (Power Query)", "h"),
 ("שלוש שאילתות. לכל אחת: Data → Get Data → From Other Sources → Blank Query → Advanced Editor → הדבקה → Close & Load To → Existing worksheet → תא A1 של הגיליון הנקוב.", ""),
 ("שימו לב: שאילתות 2–3 מקבלות טוקן (כמו בשאילתת השוק הקיימת). שאילתה 1 (תיקיית המנפיקים) לא צריכה טוקן.", ""),
 ("", ""),
]
M1 = '''let
    Folder = "C:\\CORP\\02_Issuers",
    Files = Folder.Files(Folder),
    Xl = Table.SelectRows(Files, each [Extension] = ".xlsx" and not Text.StartsWith([Name], "~$")
                                       and not Text.Contains([Folder Path], "_archive") and not Text.Contains([Folder Path], "_build")),
    Wb = Table.AddColumn(Xl, "T", each try Excel.Workbook([Content], true){[Item="tblIssuer", Kind="Table"]}[Data] otherwise null),
    Ok = Table.SelectRows(Wb, each [T] <> null),
    Out = Table.Combine(Ok[T])
in
    Out
// טוענים לגיליון נתוני_מנפיקים, תא A1. לסדרות: אותו קוד עם Item="tblSeries" → גיליון נתוני_סדרות.'''
M3 = '''let
    Token = "הדבק_כאן_את_הטוקן",
    Base = "https://raw.githubusercontent.com/yovav81/corp-market-data/main/data/",
    Get = (p) => Table.PromoteHeaders(Csv.Document(Web.Contents(Base & p, [Headers = [Authorization = "token " & Token]]),
                 [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]), [PromoteAllScalars = true]),
    Sp = Get("spreads/2026.csv"),
    Typed = Table.TransformColumnTypes(Sp, {{"duration", type number}, {"ytm_gross", type number}, {"gov_ytm", type number},
                 {"spread", type number}, {"extrap", Int64.Type}}, "en-US"),
    Dates = List.Sort(List.Distinct(Typed[date]), Order.Descending),
    Cur = Table.SelectRows(Typed, each [date] = Dates{0}),
    Prev = if List.Count(Dates) > 1 then Table.SelectRows(Typed, each [date] = Dates{1}) else Table.FirstN(Typed, 0),
    Pv = Table.RenameColumns(Table.SelectColumns(Prev, {"security_id", "spread"}), {{"spread", "spread_prev"}}),
    J = Table.NestedJoin(Cur, {"security_id"}, Pv, {"security_id"}, "pv", JoinKind.LeftOuter),
    Ex = Table.ExpandTableColumn(J, "pv", {"spread_prev"}),
    Daily = Get("daily/" & Text.Start(Dates{0}, 4) & "/" & Dates{0} & ".csv"),
    D2 = Table.TransformColumnTypes(Table.SelectColumns(Daily, {"security_id", "last_price", "chg_pct", "volume_k_nis", "maturity"}),
                 {{"last_price", type number}, {"chg_pct", type number}, {"volume_k_nis", type number}}, "en-US"),
    Out = Table.NestedJoin(Ex, {"security_id"}, D2, {"security_id"}, "d", JoinKind.LeftOuter),
    Final = Table.ExpandTableColumn(Out, "d", {"last_price", "chg_pct", "volume_k_nis", "maturity"})
in
    Final
// סדר העמודות בגיליון נתוני_שוק חייב להתאים לכותרות הקיימות: אחרי הטעינה לבדוק שהכותרות זהות (ראו שורה 1 בגיליון).'''
for txt, kind in lines:
    wg.append([txt])
    c = wg.cell(wg.max_row, 1); c.font = Font(name="Arial", size=14 if kind == "t" else 11 if kind == "h" else 10, bold=kind in ("t", "h")); c.alignment = Alignment(wrap_text=True, vertical="top")
for title, code in (("שאילתה 1 — מנפיקים (Export מכל הקבצים)", M1), ("שאילתה 3 — שוק והפרשים (אחרון + קודם)", M3)):
    wg.append([title]); wg.cell(wg.max_row, 1).font = BOLD
    for ln in code.split("\n"):
        wg.append([ln]); c = wg.cell(wg.max_row, 1); c.font = Font(name="Consolas", size=9); c.alignment = Alignment(horizontal="left")
    wg.append([""])
wb.move_sheet("מדריך", offset=wb.index(wb["פרמטרים"]) - wb.index(wb["מדריך"]))
wb.active = 0
wb.save(out); print("saved", out, "סדרות", len(series), "שוק", len(cur), "תאריך", last)
