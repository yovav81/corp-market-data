"""נרמול דירוגים לסולם אחד (סולם S&P) וקיבוץ לקבוצות השוואה.

מקורות הכתיב:
  קובץ ההיסטוריה (מגמה):  ilAA- / Aa3.il / ilA-1+ / P-1.il
  גלובס (נתונים יומיים):  AA-il   / Aa3il  / A-1+il
"""
import re

MIDROOG = {"aaa": "AAA", "aa1": "AA+", "aa2": "AA", "aa3": "AA-", "a1": "A+", "a2": "A", "a3": "A-",
           "baa1": "BBB+", "baa2": "BBB", "baa3": "BBB-", "ba1": "BB+", "ba2": "BB", "ba3": "BB-",
           "b1": "B+", "b2": "B", "b3": "B-", "caa1": "CCC+", "caa2": "CCC", "caa3": "CCC-", "ca": "CC", "c": "C"}
ORDER = ["AAA", "AA+", "AA", "AA-", "A+", "A", "A-", "BBB+", "BBB", "BBB-", "BB+", "BB", "BB-",
         "B+", "B", "B-", "CCC+", "CCC", "CCC-", "CC", "C", "D"]
RANK = {r: i for i, r in enumerate(ORDER)}
SHORT_TERM = {"A-1+", "A-1", "A-2", "A-3", "P-1", "P-2", "P-3"}


def normalize(r):
    """מחזיר דירוג בסולם S&P ('AA-') או None (ללא דירוג / קצר-טווח / WR)."""
    if r is None or (isinstance(r, float) and r != r):
        return None
    s = str(r).strip()
    if not s or s.lower() in ("nan", "none", "-", "nr", "wr", "wr.il"):
        return None
    s = re.sub(r"[\s\(\)]", "", s)
    s = re.sub(r"^il", "", s, flags=re.I)          # ilAA-
    s = re.sub(r"\.?il$", "", s, flags=re.I)       # AA-il / Aa3.il
    if s.upper() in SHORT_TERM or s.upper().startswith(("A-1", "A-2", "A-3", "P-1", "P-2", "P-3")):
        return None
    if re.fullmatch(r"[A-Z]{1,3}[+-]?", s):          # כתיב מעלות/S&P: AA-, BBB+
        return s if s in RANK else None
    key = s.lower()                                  # כתיב מידרוג: Aa3, A1, Baa2
    return MIDROOG.get(key)


def lower_of(a, b):
    """הדירוג הנמוך מבין שניים ('עדיפות לנמוך')."""
    na, nb = normalize(a), normalize(b)
    if na is None:
        return nb
    if nb is None:
        return na
    return na if RANK[na] >= RANK[nb] else nb


def group(r):
    """קבוצת השוואה גסה: AAA / AA / A / BBB- (BBB ומטה) / NR."""
    n = normalize(r)
    if n is None:
        return "NR"
    if n == "AAA":
        return "AAA"
    if n.startswith("AA"):
        return "AA"
    if n.startswith("A"):
        return "A"
    return "BBB-"


def linkage(x):
    """מיפוי סוג הצמדה / מקור גלובס לקוד אחיד: cpi / ils / fx."""
    s = str(x or "").strip()
    if s in ("מדד", "corp_cpi", "cpi"):
        return "cpi"
    if s in ("שקלי", "corp_unlinked", "ils"):
        return "ils"
    if s in ("דולר", "יורו", "corp_fx", "fx"):
        return "fx"
    return None


def dur_bucket(d):
    try:
        d = float(d)
    except (TypeError, ValueError):
        return None
    if d != d:
        return None
    if d < 1.5:
        return "0-1.5"
    if d < 3:
        return "1.5-3"
    if d < 5:
        return "3-5"
    return "5+"
