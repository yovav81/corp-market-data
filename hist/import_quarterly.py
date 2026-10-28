#!/usr/bin/env python3
"""ייבוא קובץ ההיסטוריה הרבעוני (מגמה) לפורמט ארוך.

    python hist/import_quarterly.py <קובץ.xlsx> [data]

מבנה הקובץ המקורי: שורה 1 תאריך, שורה 2 מפתח (close/yield/duration/histspread), שורה 3 כותרת עברית,
עמודות A–G מזהים ונתונים סטטיים, ומעמודה H בלוקים של 4 עמודות לכל תאריך.
כלל מחייב: התאריך בשורה 1 הוא התאריך של העמודה — העמודות אינן בהכרח בסדר כרונולוגי.

פלט:
  data/hist/quarterly_spreads.csv  — שורה לכל אג"ח×תאריך: date, security_id, name, close, ytm_gross, duration, spread (באחוזים, כמו גלובס), ok, reason
  data/hist/bond_static.csv        — נתונים סטטיים לכל אג"ח: הצמדה, דירוגים (מנורמלים), הון רשום
סינונים (ok=0): ערכי זבל (מרווח > 50%), מרווח < -1%, תשואה שלילית/קיצונית, אג"ח ממשלתיות ומק"מ.
"""
import re
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))
from ratings import normalize, lower_of, linkage  # noqa: E402

GOV_RE = re.compile(r"ממשל|מ\.?ק\.?מ")
MAX_SPREAD, MIN_SPREAD, MAX_YTM = 50.0, -1.0, 50.0   # באחוזים


def parse_date(s):
    """'21.10.26' / '05.06.2026' → '2026-06-05' (יום.חודש.שנה)."""
    d, m, y = str(s).strip().split(".")
    y = int(y)
    y = y + 2000 if y < 100 else y
    return f"{y:04d}-{int(m):02d}-{int(d):02d}"


def main(xlsx, data="data"):
    raw = pd.read_excel(xlsx, header=None)
    dates, keys = raw.iloc[0].tolist(), raw.iloc[1].tolist()
    body = raw.iloc[3:].reset_index(drop=True)
    static_cols = ["symbol", "name", "linked_asset", "lowprior_il", "maalot_rating_il", "midroog_rating_il", "listed_capital"]
    assert keys[:7] == static_cols, keys[:7]

    st = body.iloc[:, :7].copy(); st.columns = static_cols
    st["security_id"] = pd.to_numeric(st["symbol"], errors="coerce").astype("Int64").astype(str)
    st["is_gov"] = st["name"].astype(str).str.contains(GOV_RE).astype(int)
    st["linkage"] = st["linked_asset"].map(linkage)
    st["rating_maalot"] = st["maalot_rating_il"].map(normalize)
    st["rating_midroog"] = st["midroog_rating_il"].map(normalize)
    st["rating_low"] = [lower_of(a, b) for a, b in zip(st["maalot_rating_il"], st["midroog_rating_il"])]
    static = st[["security_id", "name", "linkage", "linked_asset", "rating_low", "rating_maalot", "rating_midroog",
                 "listed_capital", "is_gov"]]

    # בלוקים של 4 עמודות — לפי התאריך בשורה 1, לא לפי מיקום
    blocks = {}
    for i in range(7, raw.shape[1]):
        d = parse_date(dates[i]); blocks.setdefault(d, {})[keys[i]] = i
    rows = []
    for d, cols in sorted(blocks.items()):
        assert set(cols) == {"close", "yield", "duration", "histspread"}, (d, cols)
        part = pd.DataFrame({
            "date": d, "security_id": st["security_id"], "name": st["name"],
            "close": pd.to_numeric(body.iloc[:, cols["close"]], errors="coerce"),
            "ytm_gross": pd.to_numeric(body.iloc[:, cols["yield"]], errors="coerce") * 100,
            "duration": pd.to_numeric(body.iloc[:, cols["duration"]], errors="coerce"),
            "spread": pd.to_numeric(body.iloc[:, cols["histspread"]], errors="coerce") * 100,
            "is_gov": st["is_gov"],
        })
        rows.append(part)
    q = pd.concat(rows, ignore_index=True)
    q = q[q["spread"].notna() | q["ytm_gross"].notna() | q["close"].notna()]
    reason = pd.Series("", index=q.index)
    reason[q["is_gov"] == 1] = "gov"
    reason[(reason == "") & q["spread"].isna()] = "no_spread"
    reason[(reason == "") & (q["spread"] > MAX_SPREAD)] = "junk"
    reason[(reason == "") & (q["spread"] < MIN_SPREAD)] = "neg_spread"
    reason[(reason == "") & ((q["ytm_gross"] <= 0) | (q["ytm_gross"] > MAX_YTM))] = "bad_ytm"
    q["ok"] = (reason == "").astype(int); q["reason"] = reason
    for c in ("ytm_gross", "spread"):
        q[c] = q[c].round(3)
    q["duration"] = q["duration"].round(3)
    q = q.drop(columns=["is_gov"]).sort_values(["date", "security_id"])

    out = Path(data) / "hist"; out.mkdir(parents=True, exist_ok=True)
    q.to_csv(out / "quarterly_spreads.csv", index=False, encoding="utf-8-sig")
    static.to_csv(out / "bond_static.csv", index=False, encoding="utf-8-sig")
    print(f"תאריכים: {sorted(blocks)}")
    print(f"שורות: {len(q)}, תקינות: {int(q['ok'].sum())}, לפי סיבת פסילה: {q[q.ok == 0]['reason'].value_counts().to_dict()}")
    print(f"אג\"ח: {static.shape[0]}, מהן ממשלתיות/מק\"מ: {int(static.is_gov.sum())}")
    return 0


if __name__ == "__main__":
    sys.exit(main(*sys.argv[1:]))
