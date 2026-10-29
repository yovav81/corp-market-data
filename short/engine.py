"""
engine.py — מנוע חישוב לאג"ח קצרות (מח"מ < 1): תזרים מחדש, תשואה לפדיון ומח"מ מתוך לוח הסילוקין בפועל.

מוסכמות (תואמות לבורסה/גלובס, אומתו מול תשואות ביזפורטל/גלובס ל-28.10.2026):
  * מועד ערך = יום המסחר + 1 (סליקה T+1). ימים לפי ACT/365, ריבית שנתית אפקטיבית.
  * המחיר הוא ל-100 ע.נ. של הקרן *הנותרת* (אחרי פדיונות חלקיים), וכולל ריבית צבורה והצמדה למדד הידוע.
  * אג"ח צמודת מט"ח: אותו L (השער הגלום בערך המתואם), בלי תחזית שער — התשואה במונחי המטבע.
  * אג"ח צמודה: מקדם ההצמדה L נגזר מהערך המתואם ברוטו של גלובס: L = ערך_מתואם / (100 × (1 + ריבית צבורה)).
    תזרים ריאלי (במונחי המדד הידוע היום) = (קרן + ריבית) × L ; תשואה ריאלית = IRR של התזרים הריאלי.
    תזרים נומינלי צפוי = תזרים ריאלי × (מדד ידוע במועד התשלום / מדד ידוע היום), כאשר המדד לחודש M מתפרסם ב-15 ל-M+1
    ונחשב "ידוע" מה-16 ואילך. המדדים העתידיים נלקחים מטבלת המדדים הצפויים (cpi_expected.csv).
  * תשלום שמועד האקס שלו כבר עבר (או ≤ מועד הערך) אינו שייך לקונה — מושמט.
"""
from __future__ import annotations
import datetime as dt
from dataclasses import dataclass, field
from typing import Iterable

FREQ_MONTHS = {"שנתי": 12, "חצי שנתי": 6, "רבעוני": 3, "חודשי": 1, "דו שנתי": 24, "דו-שנתי": 24}
DEFAULT_EX_DAYS = 6   # אם מועד האקס של תשלום עתידי לא ידוע — כ-6 ימים לפני התשלום (מקובל בבורסה)


# ---------- מדדים ----------
def cpi_month_known_at(d: dt.date) -> tuple[int, int]:
    """(שנה, חודש) של המדד הידוע בתאריך d: המדד לחודש M מתפרסם ב-15 ל-M+1 ונחשב ידוע מה-16."""
    y, m = d.year, d.month
    # לפני ה-16: ידוע מדד M-2 ; מה-16: ידוע מדד M-1
    back = 1 if d.day >= 16 else 2
    m -= back
    while m <= 0:
        m += 12; y -= 1
    return y, m


def add_months(y: int, m: int, k: int) -> tuple[int, int]:
    m += k
    while m > 12: m -= 12; y += 1
    while m <= 0: m += 12; y -= 1
    return y, m


def index_factor(pay_date: dt.date, today: dt.date, exp_monthly: dict[tuple[int, int], float]) -> tuple[float, list]:
    """יחס המדד הידוע במועד התשלום למדד הידוע היום. exp_monthly: {(שנה,חודש): שינוי חודשי צפוי (שבר)}.
    מחזיר (מקדם, רשימת החודשים שנכנסו לחישוב). חודש חסר בטבלה = 0 (ומצוין ברשימה כ-None)."""
    y0, m0 = cpi_month_known_at(today); y1, m1 = cpi_month_known_at(pay_date)
    f, used = 1.0, []
    y, m = y0, m0
    while (y, m) < (y1, m1):
        y, m = add_months(y, m, 1)
        e = exp_monthly.get((y, m))
        used.append(((y, m), e))
        f *= 1.0 + (e or 0.0)
    return f, used


# ---------- תזרים ----------
@dataclass
class Flow:
    date: dt.date
    principal: float      # ל-100 ע.נ. נותרת, נומינלי
    interest: float
    ex_date: dt.date | None = None
    included: bool = True
    note: str = ""

    @property
    def nominal(self): return self.principal + self.interest


@dataclass
class Bond:
    security_id: int
    name: str
    linkage: str                   # "מדד" / "שקלי" / "דולר" / ...
    coupon: float                  # ריבית שנתית נקובה (%)
    coupon_type: str               # קבועה / משתנה
    freq_months: int               # חודשים בין תשלומי ריבית
    flows: list[Flow] = field(default_factory=list)


def flows_from_schedule(rows: Iterable[dict], next_ex: dt.date | None) -> list[Flow]:
    """rows: שורות לוח הסילוקין (pay_date, principal_amt, interest_amt) בסכומים כספיים. מנרמל ל-100 ע.נ. נותרת."""
    def num(v):
        try:
            x = float(v)
            return 0.0 if x != x else x      # NaN (תא ריק / "--") = 0
        except (TypeError, ValueError):
            return 0.0
    rows = sorted(rows, key=lambda r: r["pay_date"])
    tot = sum(num(r.get("principal_amt")) for r in rows)
    out = []
    for i, r in enumerate(rows):
        p = num(r.get("principal_amt")); it = num(r.get("interest_amt"))
        if tot <= 0:
            raise ValueError("סך הקרן בלוח הוא 0")
        ex = next_ex if (i == 0 and next_ex) else r["pay_date"] - dt.timedelta(days=DEFAULT_EX_DAYS)
        out.append(Flow(r["pay_date"], 100.0 * p / tot, 100.0 * it / tot, ex))
    return out


def accrued_fraction(val_date: dt.date, flows: list[Flow], freq_months: int) -> float:
    """חלק התקופה שחלף עד התשלום הקרוב (לריבית צבורה), לפי התדירות. 0 אם אין תשלומי ריבית."""
    nxt = next((f for f in flows if f.date > val_date), None)
    if not nxt or freq_months <= 0: return 0.0
    y, m = add_months(nxt.date.year, nxt.date.month, -freq_months)
    day = min(nxt.date.day, [31, 29 if y % 4 == 0 and (y % 100 != 0 or y % 400 == 0) else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1])
    prev = dt.date(y, m, day)
    if val_date <= prev: return 0.0
    return min(1.0, (val_date - prev).days / max(1, (nxt.date - prev).days))


def linkage_factor(adjusted_value: float, val_date: dt.date, flows: list[Flow], freq_months: int, coupon: float) -> float:
    """מקדם ההצמדה הגלום בערך המתואם ברוטו (ל-100 ע.נ. נותרת): ערך_מתואם / (1 + ריבית צבורה)."""
    per = coupon / 100.0 * freq_months / 12.0 if freq_months else 0.0
    acc = per * accrued_fraction(val_date, flows, freq_months)
    return adjusted_value / (100.0 * (1.0 + acc))


# ---------- תשואה ומח"מ ----------
def xirr(cashflows: list[tuple[dt.date, float]], guess: float = 0.05) -> float | None:
    """IRR שנתי (ACT/365, ריבית דריבית שנתית). cashflows: [(תאריך, סכום)], הראשון שלילי (המחיר)."""
    cfs = [(d, v) for d, v in cashflows if v != 0]
    if len(cfs) < 2 or not any(v < 0 for _, v in cfs) or not any(v > 0 for _, v in cfs): return None
    t0 = min(d for d, _ in cfs)
    ts = [((d - t0).days / 365.0, v) for d, v in cfs]
    def f(r): return sum(v / (1 + r) ** t for t, v in ts)
    lo, hi = -0.99, 10.0
    flo, fhi = f(lo), f(hi)
    if flo * fhi > 0: return None
    for _ in range(200):
        mid = (lo + hi) / 2; fm = f(mid)
        if abs(fm) < 1e-10: return mid
        if flo * fm < 0: hi, fhi = mid, fm
        else: lo, flo = mid, fm
    return (lo + hi) / 2


def macaulay(val_date: dt.date, flows: list[tuple[dt.date, float]], y: float) -> float | None:
    pv = [(((d - val_date).days / 365.0), v / (1 + y) ** ((d - val_date).days / 365.0)) for d, v in flows if v > 0]
    s = sum(p for _, p in pv)
    return sum(t * p for t, p in pv) / s if s else None


@dataclass
class Result:
    val_date: dt.date
    L: float
    flows_real: list[tuple[dt.date, float]]      # ריאלי = במונחי המדד הידוע היום (ל-100 ע.נ. נותרת)
    flows_nominal: list[tuple[dt.date, float]]   # נומינלי צפוי
    factors: list[float]
    y_real: float | None
    y_nominal: float | None
    dur: float | None
    excluded: list[Flow]
    months_used: list
    infl_period: float | None      # אינפלציה (צפויה) מצטברת עד התשלום האחרון, בשבר
    infl_annual: float | None      # אותה אינפלציה במונחים שנתיים


def compute(bond: Bond, price: float, adjusted_value: float | None, trade_date: dt.date,
            exp_monthly: dict[tuple[int, int], float], settle_days: int = 1) -> Result:
    val = trade_date + dt.timedelta(days=settle_days)
    live, dead = [], []
    for fl in bond.flows:
        if fl.date <= val or (fl.ex_date and fl.ex_date <= trade_date):
            fl.included = False; fl.note = "מועד האקס עבר"; dead.append(fl)
        else:
            live.append(fl)
    if not live:
        return Result(val, 1.0, [], [], [], None, None, None, dead, [], None, None)
    linked = bond.linkage in ("מדד", "צמוד", "צמוד מדד") or bond.linkage.startswith("מדד")
    fx = any(k in bond.linkage for k in ("דולר", "מט", "אירו", "יורו"))   # צמוד מט"ח: L = שער גלום בערך המתואם, ללא תחזית שער
    if (linked or fx) and adjusted_value:
        L = linkage_factor(adjusted_value, val, live, bond.freq_months, bond.coupon)
    else:
        L = 1.0
    real = [(f.date, f.nominal * L) for f in live]
    facs, used, nom = [], [], []
    for d, v in real:
        if linked:
            k, u = index_factor(d, trade_date, exp_monthly)
        else:
            k, u = 1.0, []
        facs.append(k); used += u; nom.append((d, v * k))
    y_r = xirr([(val, -price)] + real)
    y_n = xirr([(val, -price)] + nom) if linked else y_r
    dur = macaulay(val, real, y_r) if y_r is not None else None
    last = max(d for d, _ in real)
    kf = facs[[d for d, _ in real].index(last)]
    T = (last - val).days / 365.0
    infl_p = kf - 1.0 if linked else None
    infl_a = (kf ** (1 / T) - 1.0) if (linked and T > 0) else None
    return Result(val, L, real, nom, facs, y_r, y_n, dur, dead, used, infl_p, infl_a)
