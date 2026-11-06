# -*- coding: utf-8 -*-
"""rashut.py — נרמול "דוח רשות" (דיווח חודשי של מנהל קרנות לרשות ני״ע על נכסי כל קרן) לטבלות קובץ הקרן.

קלט: (א) הייצוא הפנימי של המנהל — אקסל, 26 עמודות (מספר קרן, שם קרן, שם נכס, סוג נכס, מספר מזהה, ... תאריך דוח);
     (ב) הקובץ הציבורי ממאיה — טופס ק203 "דוח חודשי" של מנהל הקרנות, TXT מופרד-טאבים (אותן עמודות + "סהכ רשומות",
         "מס. מנהל ברשם"; "%חשיפה" במקום "מקרן%"), מתפרסם כחודשיים אחרי סוף החודש, אפשר גם דחוס .txt.gz.
פלט: holdings (שורה לכל אחזקה, מסווגת), funds (שורה לכל קרן + הנתונים המדווחים ברמת הקרן), index_members.

מבנה הדוח (פוענח מהקובץ של 04/2026):
  סוג נכס < 200  = אחזקות; 201–217 = התחייבויות ויתרות; 300+ = נתוני קרן (300 שווי נקי, 301 יחידות, 319 חשיפה למניות,
                   320 חשיפה למט״ח, 325 מח״מ תיק האג״ח [בעמודת ׳כמות׳], 326/327 חשיפה חיובית/שלילית לאג״ח).
  ׳מקרן%׳        = משקל מהשווי הנקי. בנגזרים (חוזים, פורוורד, Swap, ספוט) זה החשיפה הנומינלית, לא השווי.
  מקום מסחר      = 0 ת״א; 10–13 ארה״ב; 14 קנדה; 21–24 אירופה; 30 יפן; 31 הונג-קונג; 40 אוסטרליה; 98 לא סחיר/OTC; 99 אחר.
  מטבע           = 1 ₪; 2 $; 3 £; 4 SEK; 5 CHF; 6 CAD; 10 AUD; 11 €; 13 ¥; 15 HKD.
  קוד מדרגת      = 1 מעלות (סולם S&P), 2 מידרוג (סולם Moody's), 3 Moody's, 4/5 S&P / Fitch בינ״ל, 0/6 לא מדורג.
מה שאין בדוח ומושלם מנתוני השוק שלנו (corp-market-data): הצמדה ומח״מ לכל אג״ח, מנפיק וענף (ענף רק למנפיקים המסוקרים).
חברות במדדי ת״א 35/90/125 נגזרת מאחזקות קרנות הסל העוקבות באותו דוח (תכלית סל תא35/תא90/תא125) ונשמרת ל-index_members.csv.
"""
import os, re, glob, datetime as dt
import numpy as np
import pandas as pd

G = "״"  # גרשיים עבריים — חייבים להתאים לנוסחאות בקובץ הקרן
EQ, CORP, GOV, CASH, OTHER, DERIV = "מניות", "קונצרני", "ממשלתי", "מזומן ופיקדונות", "אחר", "נגזרים"
IL, ABROAD = "ישראל", f"חו{G}ל"
CPI, ILS, FX = "צמוד מדד", "שקלי", f"מט{G}ח"
US, EU, OTH = f"ארה{G}ב", "אירופה", "אחר"
TA35, TA90, YETER = f"ת{G}א 35", f"ת{G}א 90", "יתר"

PLACE = {0: IL, 1: IL, 10: US, 11: US, 12: US, 13: US, 14: "קנדה", 21: EU, 22: EU, 23: EU, 24: EU, 30: "יפן", 31: "הונג-קונג", 40: "אוסטרליה", 98: OTH, 99: OTH}
CUR = {1: "ILS", 2: "USD", 3: "GBP", 4: "SEK", 5: "CHF", 6: "CAD", 10: "AUD", 11: "EUR", 13: "JPY", 15: "HKD"}
# סוג נכס → (אפיק, סוג מכשיר)
ASSET = {
    1: (EQ, "מניה"), 3: (EQ, f"יה{G}ש"), 5: (EQ, "זכויות"), 6: (EQ, "כתב אופציה"), 57: (EQ, "מניה לא סחירה"),
    10: (CORP, f"אג{G}ח"), 9: (CORP, f"אג{G}ח להמרה"), 22: (CORP, f"אג{G}ח לא סחיר / בהסדר"), 11: (CORP, f"ני{G}ע מסחרי"),
    102: (CORP, f"נע{G}מ"), 104: (CORP, f"נע{G}מ"), 109: (CORP, f"נע{G}מ"), 110: (CORP, f"נע{G}מ"),
    12: (GOV, f"אג{G}ח ממשלתי"), 13: (GOV, f"מק{G}מ"), 19: (GOV, f"אג{G}ח ממשלתי חו{G}ל"),
    100: (CASH, "פיקדון"), 105: (CASH, "פיקדון לזמן קצוב"), 56: (CASH, f"מט{G}ח ספוט"), 111: (CASH, "בטוחות"),
    17: ("ETF", f"קרן חו{G}ל"), 18: ("ETF", "קרן נאמנות"), 20: ("ETF", f"קרן סל חו{G}ל"), 43: ("ETF", f"קרן סל חו{G}ל"), 44: ("ETF", f"קרן סל חו{G}ל"),
    50: ("ETF", "קרן סל"), 51: ("ETF", "קרן סל"), 54: ("ETF", "קרן סל"), 55: ("ETF", "קרן סל"),
    14: (DERIV, f"אופציות מעו{G}ף"), 15: (DERIV, f"אופציות מעו{G}ף"), 33: (EQ, "חוזה עתידי על מדד"), 34: (GOV, f"חוזה עתידי אג{G}ח"),
    35: (DERIV, "חוזה עתידי סחורות"), 38: (DERIV, f"פורוורד מט{G}ח"), 45: (DERIV, "Swap"), 49: (DERIV, "Swap"),
}
NOTIONAL_CODES = {14, 15, 33, 34, 35, 38, 45, 49, 56}      # מקרן% = נומינלי
FX_ONLY_CODES = {38, 45, 49, 56}                          # לא נכסים — רק חשיפת מט״ח (משקל 0)
NO_FX_CODES = {33, 34, 35, 45, 49}                        # חוזים: אין חשיפת מט״ח על הנומינלי (רק על הביטחונות)

RX_US = re.compile(r"S&P|SP ?500|SP500|SP ?\d|NDX|NASDAQ|NASD|RUSSEL|DOW|DJ |US |USA|ארה.ב|אמריק|PHLX|ISE |INDXX|TREASUR|IPOX|BIG ?TECH|ביג ?טק|MIDCAP|SEMICON|CLOUD|CYBER|ROBOTICS|CAPMARKETS|REITS|UTILITIES|HEALTHCARE|TECH|AD ", re.I)
RX_EU = re.compile(r"DAX|STOXX|CAC|FTSE|EURO|אירופ|URO|MDAX|גרמניה|UK ", re.I)
RX_IL = re.compile(r"ת.?א ?\d|תל ?בונד|תלבונד|תלגוב|תל ?גוב|ישראל|SME|בנקים|מעו.ף|TA ?\d", re.I)
RX_FOREIGN_OTHER = re.compile(r"NIKKEI|HANG|CHINA|סין|INDIA|הודו|JAPAN|יפן|MSCI|WORLD|ACWI|GLOBAL|גלובל|עולמי|EMERG|EM |KOSPI|TOPIX|BITCOIN|ETHER|BLOCKCHAIN|GOLD|זהב|LBMA|OIL|נפט|SILVER|QUANTUM|MOMENTUM|SUPERINVESTOR|GREEN|EGAMING|TOUR", re.I)
RX_BOND_ETF = re.compile(r"בונד|BOND|תלגוב|תל ?גוב|אג.?ח|מק.?מ|TREASUR|TRBD|SHORT DUR|CRDT|CREDIT|ממשל|שקלי|צמוד", re.I)
RX_GOV_ETF = re.compile(r"תלגוב|תל ?גוב|ממשל|TREASUR|TRBD|מק.?מ", re.I)
RX_HEDGED = re.compile(r"\bממ\b|ממ$| ממ |מנוטרל|נטרול|\(ממ\)|מט.?ח מנוטר", re.I)
RX_CPI = re.compile(r"צמוד|מדד|גליל|CPI", re.I)
RX_GOV_CPI = re.compile(r"צמוד|גליל", re.I)
RX_GOV_FLOAT = re.compile(r"משתנה|גילון", re.I)


def _s(x):
    return "" if x is None or (isinstance(x, float) and np.isnan(x)) else str(x).strip()


def _manager(fund_name):
    m = re.match(r"^([א-ת]+)", fund_name.replace("(", " ").replace(")", " ").strip())
    return m.group(1) if m else fund_name.split()[0]


def _is_month_end(d):
    return ((d.replace(day=28) + dt.timedelta(days=4)).replace(day=1) - d).days == 1


def _holdings_date(filing, maya_mode):
    """׳תאריך דוח׳ (ddmmyyyy): בקובץ ממאיה (ק203) זהו תאריך האחזקות עצמו (סוף החודש, או יום הפירוק בקרן שנסגרה);
    בייצוא הפנימי של המנהל זהו תאריך ההגשה — ואז האחזקות הן לסוף החודש הקודם."""
    d = dt.datetime.strptime(str(filing).strip().zfill(8), "%d%m%Y")
    if maya_mode or _is_month_end(d):
        return d
    return d.replace(day=1) - dt.timedelta(days=1)


def _rating(grade, agency):
    g = _s(grade)
    if not g or agency in (0, 6) or not re.match(r"^[A-DP][A-Za-z0-9+\-]*$", g):
        return ""
    return g


def _region_from_name(name):
    if RX_IL.search(name):
        return IL
    if RX_US.search(name):
        return US
    if RX_EU.search(name):
        return EU
    if RX_FOREIGN_OTHER.search(name):
        return OTH
    return ""


def _makam_maturity(name, ref):
    """׳מלווה קצר מועד 616׳ / ׳מ.ק.מ. 1017׳ → חודש+סדרה+ספרת שנה (M N Y)."""
    m = re.search(r"(\d{3,4})\s*$", name)
    if not m:
        return None
    s = m.group(1)
    mon, year_d = int(s[:-2]), int(s[-1])
    year = ref.year - ref.year % 10 + year_d
    if year < ref.year - 1:
        year += 10
    try:
        return dt.datetime(year, mon, 1)
    except ValueError:
        return None


# ------------------------------------------------------------------ נתוני שוק להשלמה
class Market:
    def __init__(self, cmd, master_xlsx=None):
        self.cmd = cmd
        p = os.path.join(cmd, "data", "hist", "bond_static.csv")
        self.static = pd.read_csv(p, dtype={"security_id": str}).set_index("security_id") if os.path.exists(p) else pd.DataFrame()
        # מפת הצמדה לכל השוק: bond_static + טבלאות גלובס (latest/history/spreads) — source/curve → cpi/ils/fx
        self.linkage = {}
        src_map = {"corp_cpi": "cpi", "corp_unlinked": "ils", "corp_fx": "fx", "gov_cpi": "cpi", "gov_unlinked_fixed": "ils", "gov_makam": "ils", "gov_float": "ils"}
        for f in [os.path.join(cmd, "data", "latest.csv")] + sorted(glob.glob(os.path.join(cmd, "data", "history", "*.csv"))):
            if os.path.exists(f):
                t = pd.read_csv(f, dtype={"security_id": str}, usecols=["security_id", "source"]).drop_duplicates("security_id")
                self.linkage.update({k: src_map[v] for k, v in zip(t.security_id, t.source) if v in src_map})
        for f in sorted(glob.glob(os.path.join(cmd, "data", "spreads", "*.csv"))):
            t = pd.read_csv(f, dtype={"security_id": str}, usecols=["security_id", "curve"]).drop_duplicates("security_id")
            self.linkage.update({k: ("cpi" if v == "gov_cpi" else "ils") for k, v in zip(t.security_id, t.curve) if isinstance(v, str)})
        if len(self.static):
            self.linkage.update({k: v for k, v in self.static["linkage"].items() if isinstance(v, str)})
        p = os.path.join(cmd, "data", "hist", "spread_stats.csv")
        self.stats = pd.read_csv(p, dtype={"security_id": str}).set_index("security_id") if os.path.exists(p) else pd.DataFrame()
        p = os.path.join(cmd, "data", "hist", "quarterly_spreads.csv")
        self.quarterly = pd.read_csv(p, dtype={"security_id": str}, parse_dates=["date"]) if os.path.exists(p) else pd.DataFrame()
        self.daily_files = sorted(glob.glob(os.path.join(cmd, "data", "daily", "*", "*.csv")))
        p = os.path.join(cmd, "data", "gov", "bonds_latest.csv")
        self.gov = pd.read_csv(p) if os.path.exists(p) else pd.DataFrame()
        if len(self.gov):
            self.gov = self.gov[self.gov.security_id.notna()].copy()
            self.gov["security_id"] = self.gov.security_id.astype(float).astype(int).astype(str)
            self.gov["date"] = pd.to_datetime(self.gov["date"])
            self.gov = self.gov.set_index("security_id")
        self.sector = {}
        if master_xlsx and os.path.exists(master_xlsx):
            from openpyxl import load_workbook
            wb = load_workbook(master_xlsx, read_only=True, data_only=True)
            if "נתוני_אגח_עכשיו" in wb.sheetnames:
                rows = wb["נתוני_אגח_עכשיו"].iter_rows(values_only=True)
                h = next(rows)
                i, j, k = h.index("security_id"), h.index("sub_sector"), h.index("issuer_name")
                for r in rows:
                    if r[i]:
                        self.sector[str(r[i]).strip()] = (r[j], r[k])
            wb.close()

    def daily_snapshot(self, ref):
        """טבלת היום האחרון ≤ ref מההיסטוריה היומית (קונצרני + ממשלתי), אם קיימת."""
        best = None
        for f in self.daily_files:
            d = dt.datetime.strptime(os.path.basename(f)[:10], "%Y-%m-%d")
            if d <= ref and (ref - d).days <= 21:
                best = f
        if not best:
            return None
        t = pd.read_csv(best, dtype={"security_id": str})
        return t.drop_duplicates("security_id").set_index("security_id")

    def quarterly_duration(self, sid, ref):
        """מח״מ בתאריך ref מהנקודות הרבעוניות: אינטרפולציה בין שתי הנקודות הסמוכות (או הסטה בזמן מהקרובה)."""
        if not len(self.quarterly):
            return None, ""
        q = self.quarterly[(self.quarterly.security_id == sid) & self.quarterly.duration.notna()]
        if q.empty:
            return None, ""
        before, after = q[q.date <= ref], q[q.date > ref]
        if len(before) and len(after):
            a, b = before.iloc[-1], after.iloc[0]
            w = (ref - a.date).days / max((b.date - a.date).days, 1)
            return float(a.duration + (b.duration - a.duration) * w), "משוער (אינטרפולציה רבעונית)"
        row = (after.iloc[0] if len(after) else before.iloc[-1])
        shift = (row.date - ref).days / 365.25
        return max(float(row.duration) + shift, 0.0), "משוער (הסטה בזמן מנקודה רבעונית)"


# ------------------------------------------------------------------ נרמול
def read_report(path):
    """קורא דוח רשות: xlsx (הייצוא הפנימי של המנהל) או txt/csv מופרד-טאבים כפי שמתפרסם במאיה (טופס ק203)."""
    dt_ = {"מספר קרן": str, "מספר מזהה": str, "תאריך דוח": str}
    if str(path).lower().endswith((".txt", ".csv", ".tsv", ".txt.gz", ".csv.gz")):
        raw = pd.read_csv(path, sep="\t", dtype=dt_, encoding="utf-8-sig", keep_default_na=False, na_values=[""],
                          compression="gzip" if str(path).lower().endswith(".gz") else None)
        raw.columns = [c.strip() for c in raw.columns]
        raw = raw.rename(columns={"%חשיפה": "מקרן%"})
        for c in raw.columns:
            if pd.api.types.is_string_dtype(raw[c]) or raw[c].dtype == object:
                raw[c] = raw[c].astype(str).str.strip().replace({"nan": "", "<NA>": ""})
        for c in ("סוג נכס", "סוג מזהה", "סוג מחיר", "מקום מסחר", "שוק מסחר", "מטבע", "לפי מדיניות", "קוד מדרגת", "קוד בנק", "מס. רשומה"):
            raw[c] = pd.to_numeric(raw[c], errors="coerce").fillna(0).astype(int)
        for c in ("כמות", "שווי בשקלים", "מחיר", "שווי במטבע מסחר", "מקרן%", "כמות שהושאלה", "שער המטבע", "כמות אגח לפדיון", "ריבית שנצברה"):
            raw[c] = pd.to_numeric(raw[c], errors="coerce")
        raw["מספר קרן"] = raw["מספר קרן"].str.zfill(8)
    else:
        raw = pd.read_excel(path, dtype=dt_)
    return raw


def normalize(path, cmd, master_xlsx=None, index_members_csv=None, mkt=None):
    raw = read_report(path)
    for c in ("שם קרן", "שם נכס", "TICKER"):
        raw[c] = raw[c].astype(str).str.strip()
    raw["מספר מזהה"] = raw["מספר מזהה"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    mkt = mkt or Market(cmd, master_xlsx)

    # --- נתוני קרן (300+) ---
    info = raw[raw["סוג נכס"] >= 300]
    piv = info.pivot_table(index="מספר קרן", columns="סוג נכס", values="שווי בשקלים", aggfunc="first")
    units = info[info["סוג נכס"] == 301].set_index("מספר קרן")["כמות"]
    dur_rep = info[info["סוג נכס"] == 325].set_index("מספר קרן")["כמות"]
    first = raw.drop_duplicates("מספר קרן").set_index("מספר קרן")
    dates = pd.to_datetime(first["תאריך דוח"].astype(str).str.zfill(8), format="%d%m%Y", errors="coerce").dropna()
    maya_mode = dates.map(_is_month_end).mean() > 0.5 if len(dates) else False
    funds = []
    for fid in piv.index:
        nav = float(piv.loc[fid].get(300, np.nan) or 0)
        name = first.loc[fid, "שם קרן"]
        filing = first.loc[fid, "תאריך דוח"]
        hd = _holdings_date(filing, maya_mode)
        u = float(units.get(fid, 0) or 0)
        funds.append([fid, name, _manager(name), "", nav / 1000.0, hd, dt.datetime.strptime(filing, "%d%m%Y"),
                      u, (nav / u if u else None),
                      (float(piv.loc[fid].get(319, 0) or 0) / nav if nav else None),
                      (float(piv.loc[fid].get(320, 0) or 0) / nav if nav else None),
                      float(dur_rep.get(fid, 0) or 0) or None,
                      (float(piv.loc[fid].get(326, 0) or 0) / nav if nav else None),
                      (float(piv.loc[fid].get(327, 0) or 0) / nav if nav else None)])
    funds_df = pd.DataFrame(funds, columns=["מס_קרן", "שם", "מנהל", "סיווג", "שווי_נכסים_אלפי_ש״ח", "תאריך_דוח", "תאריך_הגשה", "יחידות", "מחיר_יחידה",
                                            "חשיפה_מניות_מדווחת", "חשיפה_מטח_מדווחת", "מחמ_אגח_מדווח", "חשיפה_אגח_חיובית", "חשיפה_אגח_שלילית"])
    nav_map = funds_df.set_index("מס_קרן")["שווי_נכסים_אלפי_ש״ח"] * 1000
    hd_map = funds_df.set_index("מס_קרן")["תאריך_דוח"]
    ref = hd_map.max()

    # --- חברות במדדים מקרנות הסל העוקבות ---
    members = {}
    eq = raw[(raw["סוג נכס"] == 1) & (raw["מקום מסחר"] == 0)]
    for fid, idx in (("01143700", TA35), ("01143783", TA90)):
        ids = set(eq[eq["מספר קרן"] == fid]["מספר מזהה"])
        for s in ids:
            members[s] = idx
    if not members and index_members_csv and os.path.exists(index_members_csv):
        m = pd.read_csv(index_members_csv, dtype=str)
        members = dict(zip(m.security_id, m["index"]))
    elif members and index_members_csv:
        pd.DataFrame({"security_id": list(members), "index": list(members.values()), "as_of": ref.date()}).to_csv(index_members_csv, index=False, encoding="utf-8-sig")

    # --- אחזקות ---
    snap = mkt.daily_snapshot(ref)
    hold = raw[raw["סוג נכס"] < 300].copy()
    rows = []
    for r in hold.itertuples(index=False):
        code, fid, name = int(r[3]), r[0], r[2]
        sid, place, cur = r[4], int(r[13]), CUR.get(int(r[15]), "אחר")
        w = float(r[10] or 0) / 100.0
        val_k = float(r[7] or 0) / 1000.0
        asset, instr = ASSET.get(code, (OTHER, f"קוד {code}"))
        if 201 <= code <= 217:
            asset, instr = OTHER, "התחייבויות ויתרות"
        region, sub, link, dur, dur_src, issuer, sector, rating = "", "", "", None, "", "", "", ""
        is_il = place in (0, 1)

        if asset == EQ:
            region = IL if is_il and code != 33 else PLACE.get(place, OTH)
            if code == 33:
                region = _region_from_name(name) or PLACE.get(place, OTH)
            if region == IL:
                sub = members.get(sid, YETER if members else "")
            else:
                sub = region if region in (US, EU) else OTH
                region = ABROAD
            issuer = name
        elif asset == "ETF":
            reg = _region_from_name(name)
            if RX_BOND_ETF.search(name):
                asset = GOV if RX_GOV_ETF.search(name) else CORP
                link = CPI if RX_CPI.search(name) else (ILS if (reg in (IL, "") and not re.search(r"דולר|USD|\$", name)) else FX)
                region = IL if reg in (IL, "") else ABROAD
            else:
                asset = EQ
                if reg in (IL, ""):
                    region, sub = IL, ""
                else:
                    region, sub = ABROAD, (reg if reg in (US, EU) else OTH)
            issuer = name
        elif asset == CORP:
            region = IL if is_il else ABROAD
            rating = _rating(r[21], int(r[22]))
            if is_il:
                lk = mkt.linkage.get(sid, "")
                link = {"cpi": CPI, "ils": ILS, "fx": FX}.get(lk, FX if cur != "ILS" else "לא ידוע")
                if code in (102, 104, 109, 110, 11):
                    link = ILS
                if link == "לא ידוע" and re.search(r"פיקדונ|פקדונ|פיקדון|פקדון|פקד |פקד'|פק אג|ריביות|פאי ", name):
                    link, instr = ILS, f"אג{G}ח פיקדון (ריבית משתנה)"
                if sid in mkt.stats.index:
                    issuer = _s(mkt.stats.loc[sid, "issuer"])
                sector, iss2 = mkt.sector.get(sid, ("לא מסווג", None))
                issuer = issuer or _s(iss2) or re.sub(r"\s*אג.?ח.*$", "", name)
                sector = sector or "לא מסווג"
                if snap is not None and sid in snap.index and pd.notna(snap.loc[sid, "duration"]):
                    dur, dur_src = float(snap.loc[sid, "duration"]), "נתוני שוק יומיים"
                else:
                    dur, dur_src = mkt.quarterly_duration(sid, ref)
            else:
                link, issuer, sector = FX, re.sub(r"\s+\d.*$", "", name), f"חו{G}ל"
        elif asset == GOV:
            region = IL if is_il and code != 34 else ABROAD
            issuer = "מדינת ישראל" if region == IL else name
            if region == IL:
                link = CPI if RX_GOV_CPI.search(name) else ILS
                if RX_GOV_FLOAT.search(name):
                    instr = "ממשלתי בריבית משתנה"
                if snap is not None and sid in snap.index and pd.notna(snap.loc[sid, "duration"]):
                    dur, dur_src = float(snap.loc[sid, "duration"]), "נתוני שוק יומיים"
                elif code == 13:
                    mat = _makam_maturity(name, ref)
                    if mat:
                        dur, dur_src = max((mat - ref).days / 365.25, 0.0), "משוער (זמן לפדיון)"
                elif sid in mkt.gov.index:
                    g = mkt.gov.loc[sid]
                    d0 = g["duration_globes"] if pd.notna(g["duration_globes"]) else g["modified"]
                    dur, dur_src = max(float(d0) + (g["date"] - ref).days / 365.25, 0.0), "משוער (הסטה בזמן מנתון עדכני)"
            else:
                link = FX
                mt = re.search(r"(\d{2})/(\d{2})/(\d{2})\s*$", name)
                if code == 19 and mt:
                    m_, d_, y_ = (int(x) for x in mt.groups())
                    dur, dur_src = max((dt.datetime(2000 + y_, m_, d_) - ref).days / 365.25, 0.0), "משוער (זמן לפדיון)"
        elif asset == CASH:
            region = IL
            link = ILS if cur == "ILS" else FX
        elif asset == DERIV:
            region = ""

        # משקל: נגזרי מט״ח/Swap/ספוט אינם נכסים → 0; חוזים על מדדים/אג״ח → נומינלי (כך גם הרשות מודדת חשיפה)
        weight = 0.0 if code in FX_ONLY_CODES else w
        fx_w = 0.0
        if code in FX_ONLY_CODES:
            fx_w = w if code not in (45, 49) else 0.0
        elif code not in NO_FX_CODES:
            if cur != "ILS" or link == FX:
                fx_w = w
            elif asset in (EQ, CORP, GOV) and r[3] in (17, 18, 20, 43, 44, 50, 51, 54, 55) and region == ABROAD and not RX_HEDGED.search(name):
                fx_w = w
        if dur is not None and not (0 <= dur <= 40):      # ערך לא סביר בנתוני המקור → מתעלמים
            dur, dur_src = None, "נפסל (ערך לא סביר במקור)"
        group = asset if asset in (EQ, CORP, GOV) else OTHER
        rows.append([fid, r[1], hd_map.get(fid), asset, region, sub, link, sid, name, issuer, sector, rating, cur,
                     dur, val_k, weight, (dur * weight if dur is not None else 0.0), fx_w, group, code, instr, dur_src, float(r[6] or 0)])

    cols = ["מס_קרן", "שם_קרן", "תאריך_דוח", "אפיק", "אזור", "תת_סיווג", "הצמדה", "מס_נייר", "שם_נייר", "מנפיק", "ענף", "דירוג", "מטבע",
            "מח״מ", "שווי_אלפי_ש״ח", "משקל", "מחמ_x_משקל", "משקל_מט״ח", "קבוצה", "סוג_נכס_קוד", "סוג_מכשיר", "מקור_מחמ", "כמות"]
    h = pd.DataFrame(rows, columns=cols)
    h["מנהל"] = h["מס_קרן"].map(funds_df.set_index("מס_קרן")["מנהל"])
    # מפתח: קרן|קבוצה|מספר רץ לפי שווי יורד (לפירוט בכרטסת)
    h = h.sort_values(["מס_קרן", "קבוצה", "שווי_אלפי_ש״ח"], ascending=[True, True, False])
    h["רץ"] = h.groupby(["מס_קרן", "קבוצה"]).cumcount() + 1
    h["מפתח"] = h["מס_קרן"] + "|" + h["קבוצה"] + "|" + h["רץ"].astype(str)
    h["משקל_עם_מחמ"] = np.where(h["מח״מ"].notna(), h["משקל"], 0.0)
    out_cols = ["מפתח", "מס_קרן", "שם_קרן", "מנהל", "תאריך_דוח", "אפיק", "אזור", "תת_סיווג", "הצמדה", "מס_נייר", "שם_נייר", "מנפיק", "ענף", "דירוג",
                "מטבע", "מח״מ", "שווי_אלפי_ש״ח", "משקל", "מחמ_x_משקל", "משקל_מט״ח", "קבוצה", "סוג_נכס_קוד", "סוג_מכשיר", "מקור_מחמ", "כמות", "משקל_עם_מחמ"]
    h = h[out_cols]
    # סיווג קרן גס לפי האחזקות (לרשימת הקרנות)
    agg = h.groupby(["מס_קרן", "אפיק"])["משקל"].sum().unstack(fill_value=0)
    rep_eq = funds_df.set_index("מס_קרן")["חשיפה_מניות_מדווחת"].fillna(0)
    names_ = funds_df.set_index("מס_קרן")["שם"]
    def classify(fid):
        a = agg.loc[fid] if fid in agg.index else {}
        e, c, g = abs(rep_eq.get(fid, 0)), a.get(CORP, 0), a.get(GOV, 0)
        if "כספית" in names_.get(fid, ""): return "כספית"
        if e > 0.5: return "מניות"
        if e > 0.35: return "מעורב"
        if c + g < 0.3: return "מזומן / אחר"
        if e > 0.02: return f"אג{G}ח כללי עם מניות"
        if c > 0.6: return f"אג{G}ח קונצרני"
        if g > 0.6: return f"אג{G}ח ממשלתי"
        return f"אג{G}ח כללי"
    funds_df["סיווג"] = funds_df["מס_קרן"].map(classify)
    return h, funds_df, members


def latest_report(folder):
    """הקובץ העדכני ביותר לפי YYYY-MM בשם הקובץ (ואם אין — לפי זמן שינוי)."""
    files = [f for ext in ("*.xlsx", "*.txt", "*.txt.gz") for f in glob.glob(os.path.join(folder, ext))]
    def key(f):
        m = re.search(r"(20\d{2})-(\d{2})", os.path.basename(f))
        return (m.group(1) + m.group(2) if m else "000000", os.path.getmtime(f))
    return max(files, key=key) if files else None


if __name__ == "__main__":
    import sys
    h, f, m = normalize(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else None)
    print(len(h), "אחזקות,", len(f), "קרנות,", len(m), "מניות במדדים")
    print(f.head(10).to_string())
