"""short_sheet.py — הגיליון ׳מחמ מתחת ל 1׳ באקסל העל: תשואה מחושבת מחדש לאג"ח קצרות מתוך לוח הסילוקין בפועל,
עם תוואי מדדים צפויים חי (ניתן לדריסה ידנית בגיליון) ותזרים לכל סדרה. נקרא מ-build_master.py.
מקורות: data/short/short_bonds.csv, cashflows.csv, cpi_expected.csv, cpi_sources.json (נבנים ב-GitHub מדי ערב, short/*.py)."""
import datetime as dt, json, os
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import FormulaRule

SLOTS = 8


def _d(s):
    try: return pd.to_datetime(str(s)).to_pydatetime() if s not in (None, "") and not (isinstance(s, float) and pd.isna(s)) else None
    except Exception: return None


def build(wb, sp_csv, *, index, styles, SE, PR):
    """index: מיקום הגיליון; styles: dict(HDR, SUB, INP, WH, BASE, BOLD, BOX, red); SE(h): טווח בנתוני_סדרות; PR: פרמטרים."""
    HDR, SUB, INP, WH, BASE, BOLD, BOX, RED = (styles[k] for k in ("HDR", "SUB", "INP", "WH", "BASE", "BOLD", "BOX", "red"))
    sd = os.path.join(os.path.dirname(os.path.dirname(sp_csv)), "short")
    try:
        sb = pd.read_csv(os.path.join(sd, "short_bonds.csv"), encoding="utf-8-sig")
        cf = pd.read_csv(os.path.join(sd, "cashflows.csv"), encoding="utf-8-sig")
        cpi = pd.read_csv(os.path.join(sd, "cpi_expected.csv"), encoding="utf-8-sig")
        src = json.load(open(os.path.join(sd, "cpi_sources.json"), encoding="utf-8"))
    except Exception as e:
        print("אזהרה: אין נתוני אג״ח קצרות (data/short) —", e); return None
    gd = os.path.join(os.path.dirname(sd), "gov")
    try: gcv = pd.read_csv(os.path.join(gd, "curves_latest.csv"), encoding="utf-8-sig")
    except Exception: gcv = pd.DataFrame()
    ws = wb.create_sheet("מחמ מתחת ל 1", index); ws.sheet_view.rightToLeft = True
    ws["A1"] = "אג״ח קונצרניות עם מח״מ מתחת ל-1 — תשואה מחושבת מחדש מלוח הסילוקין, עם מדד צפוי"; ws["A1"].font = Font(name="Arial", size=14, bold=True)
    ws["A2"] = ("התשואה הריאלית של גלובס/הבורסה נכונה חשבונאית אבל מתעלמת מהמדדים שעוד לא פורסמו: באג״ח צמודה קצרה שני מדדים צפויים יכולים לשנות את התשואה "
                "השקלית הצפויה בנקודות אחוז. כאן כל סדרה נבנית מחדש מתזרים הפירעונות (ביזפורטל), כל תשלום מוצמד למדד הידוע במועדו לפי טבלת המדדים הצפויים למטה, "
                "ומחושבות תשואה ריאלית, תשואה שקלית צפויה, מח״מ ומרווח מול ממשלתי שקלי באותו מח״מ. תאים צהובים — דריסה ידנית של מדד צפוי (למשל תחזית בנק); "
                "השינוי מתגלגל לכל הסדרות. מוסכמות: סליקה T+1, ACT/365, ריבית שנתית אפקטיבית; תשלום שמועד האקס שלו עבר הושמט.")
    ws["A2"].font = BASE; ws["A2"].alignment = Alignment(wrap_text=True, vertical="top"); ws.merge_cells("A2:P2"); ws.row_dimensions[2].height = 64
    def put(cell, v, font=BASE, fill=None, fmt=None, border=None, align=None):
        c = ws[cell]; c.value = v; c.font = font
        if fill: c.fill = fill
        if fmt: c.number_format = fmt
        if border: c.border = border
        if align: c.alignment = align
        return c
    def hdr(row, labels, c0=1):
        for i, t in enumerate(labels):
            c = ws.cell(row, c0 + i, t); c.fill = HDR; c.font = WH; c.border = BOX
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # ---- פרמטרים (שורה אחת) ----
    date = _d(sb["date"].iloc[0]) if len(sb) else dt.datetime.today()
    ex = src.get("expectations", {}); used_key = src.get("used", "gov"); used = ex.get(used_key) or {}
    put("A4", "תאריך נתונים", BOLD); put("B4", date, fmt="dd/mm/yyyy", border=BOX)
    put("C4", "מועד ערך (T+1)", BOLD); put("D4", "=B4+1", fmt="dd/mm/yyyy", border=BOX)
    put("E4", "המדד הידוע", BOLD); put("F4", src.get("known_month"), border=BOX); put("G4", src.get("known_index"), fmt="0.0", border=BOX)
    put("H4", "ציפיות לשנה", BOLD); put("I4", used.get("value"), fmt="0.00", border=BOX)
    put("J4", ("מקובץ הממשלתי (BEI לשנה)" if used_key == "gov" else used.get("name", "")) + f" · {used.get('period', '')}", Font(name="Arial", size=9, italic=True))
    put("N4", "טבלת המדדים (דריסה ידנית) והעקום הממשלתי — מתחת לרשימת הסדרות ↓", Font(name="Arial", size=9, italic=True, color="1F3864"))
    gcv_n = len(gcv)
    # ---- טבלת הסדרות ----
    HR = 6
    put(f"A{HR-1}", f"סדרות עם מח״מ גלובס מתחת ל-1 ({len(sb)}; חושבו {(sb['status']=='ok').sum()})", Font(name="Arial", size=12, bold=True))
    H = ["מס׳ ני״ע", "סדרה", "הצמדה", "סוג פדיון", "ריבית %", "תדירות", "מעלות", "מידרוג", "יש מודל", "פדיון סופי", "מחיר", "ערך מתואם",
         "מח״מ גלובס", "תשואה ברוטו גלובס %", "הפרש ממשלתי גלובס (נ״ב)", "תשלומים נותרים", "תשלום קרוב", "% קרן בתשלום הקרוב", "אקס קרוב",
         "מקדם הצמדה L", "מח״מ מחושב", "תשואה ריאלית מחושבת %", "פער מול גלובס (נ״ב)", "אינפלציה צפויה עד הפדיון (שנתי) %",
         "תשואה שקלית צפויה %", "ממשלתי שקלי באותו מח״מ %", "מרווח שקלי צפוי (נ״ב)", "ממשלתי צמוד באותו מח״מ %", "מרווח ריאלי (נ״ב)",
         "ימים לפדיון", "רווח כולל לפדיון % (לא שנתי)", "רגישות: +0.1% במדד = נ״ב בתשואה השנתית",
         "דגלים / סטטוס", "מחזור (אלפי ₪)"]
    NC = len(H)
    hdr(HR, H); ws.row_dimensions[HR].height = 48
    # בלוק תזרים
    C0 = NC + 2                      # עמודת מועד הערך
    cD, cR, cF, cN = C0, C0 + SLOTS + 1, C0 + 2 * (SLOTS + 1), C0 + 2 * (SLOTS + 1) + SLOTS
    hdr(HR, ["מועד ערך"] + [f"תשלום {i+1}" for i in range(SLOTS)] + ["−מחיר"] + [f"ריאלי {i+1} (×L)" for i in range(SLOTS)]
        + [f"מקדם מדד {i+1}" for i in range(SLOTS)] + ["−מחיר"] + [f"שקלי צפוי {i+1}" for i in range(SLOTS)], cD)
    put(f"{L(cD)}{HR-1}", "תזרים לכל סדרה ל-100 ע.נ. נותרת (עמודות מקובצות — לחצו + לפתיחה)", BOLD)
    cfg = {int(k): g.sort_values("slot") for k, g in cf.groupby("security_id")} if len(cf) else {}
    # מיקומי הטבלאות שמתחת לרשימה (ידועים מראש כדי שנוסחאות התזרים יפנו אליהן)
    RL_ = HR + len(sb)
    GR0 = RL_ + 8                                   # כותרת העקום הממשלתי (אחרי 5 שורות ההסבר)
    GT0 = GR0 + 2; GT1 = GT0 + max(1, gcv_n) - 1     # שורות העקום
    R0 = GT1 + 3                                     # כותרת טבלת המדדים
    known = cpi[cpi["kind"] == "ידוע"].tail(6); expc = cpi[cpi["kind"] != "ידוע"]
    cpi_rows = pd.concat([known, expc]); first = R0 + 2; lastr = first + len(cpi_rows) - 1
    KEY, CUM = f"$A${first}:$A${lastr}", f"$H${first}:$H${lastr}"; LASTCUM = f"$H${lastr}"
    GTEN, GNOM, GREAL = f"$B${GT0}:$B${GT1}", f"$C${GT0}:$C${GT1}", f"$D${GT0}:$D${GT1}"
    def ginterp(x, ys):
        m = f"MATCH({x},{GTEN},1)"
        return (f'IF({x}="","",IF(ISNA({m}),INDEX({ys},1),IF({m}=ROWS({GTEN}),INDEX({ys},ROWS({GTEN})),'
                f'INDEX({ys},{m})+({x}-INDEX({GTEN},{m}))*(INDEX({ys},{m}+1)-INDEX({ys},{m}))/(INDEX({GTEN},{m}+1)-INDEX({GTEN},{m})))))')
    lk = {"corp_cpi": "מדד", "corp_unlinked": "שקלי", "corp_fx": "מט״ח"}
    r = HR + 1
    for _, b in sb.iterrows():
        sid = int(b["security_id"]); ok = b["status"] == "ok"
        g = cfg.get(sid); n = len(g) if g is not None else 0
        linked = ok and "מדד" in str(b.get("linkage"))
        v = [sid, b["name"], (str(b["linkage"]) if ok else lk.get(b["source"], "")), b.get("redemption_type"), b.get("coupon"), b.get("coupon_freq"),
             b.get("rating_maalot"), b.get("rating_midroog"), f'=IF(COUNTIF({SE("security_id")},A{r})>0,"כן","")', _d(b.get("maturity")),
             b.get("last_price"), b.get("adjusted_value"), b.get("dur_globes"), b.get("ytm_globes"),
             (None if pd.isna(b.get("spread_globes")) else float(b["spread_globes"]) * 100), n if ok else None,
             _d(b.get("next_pay_date")), b.get("next_principal_pct"), _d(b.get("next_ex")), b.get("L") if ok else None, b.get("dur_calc")]
        if ok and n:
            dates = f"${L(cD)}{r}:${L(cD+SLOTS)}{r}"
            v += [f"=IFERROR(XIRR({L(cR)}{r}:{L(cR+SLOTS)}{r},{dates})*100,\"\")",
                  f'=IF(OR(V{r}="",N{r}=""),"",(V{r}-N{r})*100)',
                  (f"=IFERROR((INDEX({L(cF)}{r}:{L(cF+SLOTS-1)}{r},{n})^(365/({L(cD+n)}{r}-{L(cD)}{r}))-1)*100,\"\")" if linked else None),
                  f"=IFERROR(XIRR({L(cN)}{r}:{L(cN+SLOTS)}{r},{dates})*100,\"\")",
                  f"__GOVN_{r}__", f'=IF(OR(Y{r}="",Z{r}=""),"",(Y{r}-Z{r})*100)',
                  f"__GOVR_{r}__", (f'=IF(OR(V{r}="",AB{r}=""),"",(V{r}-AB{r})*100)' if linked else f'=IF(OR(V{r}="",Z{r}=""),"",(V{r}-Z{r})*100)')]
        else:
            v += [None] * 8
        if ok and n:
            v += [f'=IF({L(cD+n)}{r}="","",{L(cD+n)}{r}-$D$4)', f"=IFERROR(SUM({L(cN+1)}{r}:{L(cN+SLOTS)}{r})/K{r}*100-100,\"\")", b.get("cpi_sens_bp")]
        else:
            v += [None, None, None]
        v += [b.get("flags") if ok else b.get("status"), b.get("volume_k_nis")]
        fm = ["0", None, None, None, "0.00", None, None, None, None, "dd/mm/yyyy", "0.00", "0.00", "0.00", "0.00", "0", "0", "dd/mm/yyyy", "0.0", "dd/mm/yyyy",
              "0.0000", "0.00", "0.00", "0", "0.00", "0.00", "0.00", "0", "0.00", "0", "0", "0.00", "0", None, "#,##0"]
        for j, x in enumerate(v):
            if isinstance(x, float) and pd.isna(x): x = None
            if x == f"__GOVN_{r}__": x = ("=" + ginterp(f"U{r}", GNOM)) if gcv_n else b.get("gov_ils_at_dur")
            if x == f"__GOVR_{r}__": x = ("=" + ginterp(f"U{r}", GREAL)) if gcv_n else b.get("gov_cpi_at_dur")
            c = ws.cell(r, 1 + j, x); c.font = BASE
            if fm[j]: c.number_format = fm[j]
        # תזרים
        if ok and n:
            ws.cell(r, cD, f"=$D$4").number_format = "dd/mm/yyyy"
            lastd = _d(g["pay_date"].iloc[-1])
            ws.cell(r, cR, f"=-K{r}").number_format = "0.00"; ws.cell(r, cN, f"=-K{r}").number_format = "0.00"
            for i in range(SLOTS):
                if i < n:
                    x = g.iloc[i]
                    ws.cell(r, cD + 1 + i, _d(x["pay_date"])).number_format = "dd/mm/yyyy"
                    ws.cell(r, cR + 1 + i, float(x["real"])).number_format = "0.0000"
                    dc = f"{L(cD+1+i)}{r}"
                    if linked:
                        f = (f'=IFERROR(INDEX({CUM},MATCH(DATE(YEAR({dc}),MONTH({dc})-IF(DAY({dc})>=16,1,2),1),{KEY},0)),'
                             f'IF(DATE(YEAR({dc}),MONTH({dc})-IF(DAY({dc})>=16,1,2),1)<$A${first},1,{LASTCUM}))')
                    else:
                        f = 1
                    ws.cell(r, cF + i, f).number_format = "0.0000"
                    ws.cell(r, cN + 1 + i, f"={L(cR+1+i)}{r}*{L(cF+i)}{r}").number_format = "0.0000"
                else:
                    ws.cell(r, cD + 1 + i, lastd).number_format = "dd/mm/yyyy"
                    ws.cell(r, cR + 1 + i, 0); ws.cell(r, cF + i, 1); ws.cell(r, cN + 1 + i, 0)
            for cc in range(cD, cN + SLOTS + 1): ws.cell(r, cc).font = Font(name="Arial", size=9, color="808080")
        r += 1
    RL = r - 1
    widths = [10, 20, 7, 11, 7, 9, 8, 8, 7, 10, 8, 9, 7, 9, 9, 8, 10, 8, 10, 8, 8, 10, 9, 11, 10, 10, 10, 10, 9, 8, 11, 12, 40, 9]
    for j, w in enumerate(widths): ws.column_dimensions[L(j + 1)].width = max(ws.column_dimensions[L(j + 1)].width or 0, w) if j < 9 else w
    for cc in range(cD, cN + SLOTS + 1): ws.column_dimensions[L(cc)].width = 10
    ws.column_dimensions.group(L(cD), L(cN + SLOTS), outline_level=1, hidden=False)
    ws.freeze_panes = f"C{HR+1}"; ws.auto_filter.ref = f"A{HR}:{L(NC)}{RL}"
    assert RL == RL_, (RL, RL_)
    rng = f"A{HR+1}:{L(NC)}{RL}"
    ws.conditional_formatting.add(f"W{HR+1}:W{RL}", FormulaRule(formula=[f'AND(W{HR+1}<>"",ABS(W{HR+1})>={PR["p_sgap"]})'], fill=RED))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'ISNUMBER(SEARCH("מצוקה",$AG{HR+1}))'], font=Font(name="Arial", size=10, color="808080")))
    ws.conditional_formatting.add(f"Y{HR+1}:Y{RL}", FormulaRule(formula=[f'AND(AD{HR+1}<>"",AD{HR+1}<60)'], font=Font(name="Arial", size=10, italic=True, color="7F6000")))
    ws.conditional_formatting.add(f"Y{HR+1}:Y{RL}", FormulaRule(formula=[f'AND(Y{HR+1}<>"",V{HR+1}<>"",ABS(Y{HR+1}-V{HR+1})>=1)'], fill=SUB))
    # הסבר קצר מתחת לטבלה
    r = RL + 2
    for t in ["קריאה: ׳תשואה ריאלית מחושבת׳ צריכה להתלכד עם גלובס (פער של נ״ב בודדות); פער גדול = עיוות בנתון המצוטט (למשל אחרי פדיון חלקי) — לבדוק.",
              "׳תשואה שקלית צפויה׳ = אותו תזרים כשכל תשלום מוצמד למדד הידוע במועדו לפי טבלת המדדים למטה; ׳מרווח שקלי צפוי׳ = מול העקום הממשלתי השקלי (קובץ הממשלתי) באותו מח״מ.",
              "מקדם ההצמדה L נגזר מהערך המתואם ברוטו של גלובס (ערך מתואם ÷ (1 + ריבית צבורה)), לא ממדד הבסיס — דיוק של כ-0.01%. רצפת מדד (אם קיימת בשטר) לא מיושמת.",
              "טווח קצר מ-60 יום (ימים לפדיון בכתום): התשואה השנתית נשענת על מדד צפוי אחד — להסתכל על ׳רווח כולל לפדיון׳ (לא שנתי) ועל הרגישות: כמה נ״ב בתשואה השנתית משנה סטייה של 0.1% במדד.",
              "ריבית משתנה חושבה לפי הריבית הנוכחית; צמודות מט״ח — ללא תחזית שער. מוסכמות: סליקה T+1, ACT/365, ריבית שנתית אפקטיבית; תשלום שמועד האקס שלו עבר הושמט. מקור הלוחות: ביזפורטל (מתעדכן בענן בכל ערב)."]:
        put(f"A{r}", t, Font(name="Arial", size=9)); r += 1
    # ---- העקום הממשלתי (מקובץ הממשלתי) ----
    put(f"A{GR0}", "העקום הממשלתי — מקובץ הממשלתי (00_Gov/ממשלתי.xlsx, גיליון 03_עקומים; אותו מקור: data/gov/curves_latest.csv)", Font(name="Arial", size=12, bold=True))
    hdr(GR0 + 1, ["", "טווח (שנים)", "שקלי %", "ריאלי %", "BEI %", "תאריך"])
    for i in range(max(1, gcv_n)):
        rr = GT0 + i
        if gcv_n:
            x = gcv.iloc[i]
            put(f"B{rr}", float(x["tenor"]), fmt="0.00", border=BOX); put(f"C{rr}", float(x["nominal"]), fmt="0.00", border=BOX)
            put(f"D{rr}", float(x["real"]), fmt="0.00", border=BOX); put(f"E{rr}", None if pd.isna(x["bei"]) else float(x["bei"]), fmt="0.00", border=BOX)
            put(f"F{rr}", str(x["date"]), border=BOX)
        else:
            put(f"B{rr}", "אין נתוני עקום — הממשלתי בשורות הסדרות מגלובס", Font(name="Arial", size=9, italic=True))
    put(f"H{GT0}", "׳ממשלתי שקלי/צמוד באותו מח״מ׳ בטבלה למעלה מחושב באינטרפולציה על העקום הזה (שקלי = מק״מ + שקלית קבועה). כל ערב העקום מתחלף יחד עם קובץ הממשלתי.", Font(name="Arial", size=9, italic=True))
    # ---- טבלת מדדים ----
    put(f"A{R0}", "מדדים: ידועים וצפויים (תאים צהובים — דריסה ידנית; נשמרת בין בנייה לבנייה)", Font(name="Arial", size=12, bold=True))
    hdr(R0 + 1, ["חודש המדד", "מועד פרסום", "ידוע מ-", "מדד", "שינוי % אוטומטי", "דריסה ידנית %", "שינוי % בפועל", "מקדם מצטבר מהמדד הידוע", "מקור"])
    ws.row_dimensions[R0 + 1].height = 32
    r = first; last_known_row = None
    for _, x in cpi_rows.iterrows():
        y, m = str(x["ym"]).split("-"); pub = _d(x["pub_date"])
        put(f"A{r}", dt.datetime(int(y), int(m), 1), fmt="mm/yyyy", border=BOX)
        put(f"B{r}", pub, fmt="dd/mm/yyyy", border=BOX); put(f"C{r}", f"=B{r}+1", fmt="dd/mm/yyyy", border=BOX)
        isk = str(x["kind"]) == "ידוע"
        put(f"D{r}", None if pd.isna(x["index"]) else float(x["index"]), fmt="0.0", border=BOX)
        put(f"E{r}", None if pd.isna(x["monthly_pct"]) else float(x["monthly_pct"]), fmt="0.00", border=BOX, fill=None if not isk else SUB)
        put(f"F{r}", None, fill=None if isk else INP, fmt="0.00", border=BOX)
        if isk:
            put(f"G{r}", f"=E{r}", fmt="0.00", border=BOX, fill=SUB); last_known_row = r
            put(f"H{r}", None, border=BOX)
        else:
            put(f"G{r}", f'=IF(F{r}<>"",F{r},E{r})', fmt="0.00", border=BOX)
            prev = f"H{r-1}" if r - 1 != last_known_row else "1"
            put(f"H{r}", f"={prev}*(1+G{r}/100)", fmt="0.0000", border=BOX)
        put(f"I{r}", str(x["source"]) if not isk else "הלמ״ס", Font(name="Arial", size=9), border=BOX)
        r += 1
    if last_known_row: put(f"H{last_known_row}", 1.0, fmt="0.0000", border=BOX, fill=SUB)
    put(f"A{r+1}", "המדד לחודש M מתפרסם ב-15 ל-M+1 ונחשב ידוע מה-16; תשלום מוצמד למדד הידוע במועדו. תוואי: ציפיות לשנה (קובץ הממשלתי) + עונתיות הלמ״ס; דריסה בעמודה F גוברת.", Font(name="Arial", size=9))
    return ws
