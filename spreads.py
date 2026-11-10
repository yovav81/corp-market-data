#!/usr/bin/env python3
"""חישוב הפרש ממשלתי יומי לכל אג"ח קונצרנית, מנתוני הגלם ב-data/history.

צמוד  -> מול עקום ממשלתי צמוד (gov_cpi);  שקלי -> מול עקום ממשלתי שקלי קבוע (gov_unlinked_fixed).
אינטרפולציה ליניארית לפי מח"מ; מחוץ לטווח העקום — ערך קצה, ומסומן extrap=1.
מחושב מחדש לכל ההיסטוריה בכל הרצה, לכן שינוי שיטה משפיע גם על העבר.
סדרה שגלובס מפרסם לה מחיר בלי מח"מ/תשואה ויש לה לוח תשלומים ב-calc/schedules.csv — מחושבת מהלוח (calc/bond_calc.py) ונכתבת באותו קובץ עם calc=1.
ערך אמיתי של גלובס לעולם לא נדרס: חישוב רק כשגלובס לא נתן מח"מ או תשואה. calc=0 בכל שורה שמקורה בגלובס.
פלט: data/spreads/<שנה>.csv; data/calc/coverage.csv — לכל סדרה בלי מח"מ/תשואה בגלובס ביום האחרון: חושבה או לא, ולמה לא.
"""
import sys
from pathlib import Path
import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from calc.bond_calc import load_schedules, calc_rows

PAIRS = {"corp_cpi": "gov_cpi", "corp_unlinked": "gov_unlinked_fixed"}
MAX_YTM, MIN_DUR = 50.0, 0.3   # סינון אג"ח במצוקה / מח"מ קצר מאוד
COLS = ["date", "source", "security_id", "name", "duration", "ytm_gross", "gov_ytm", "spread", "extrap", "curve",
        "rating_maalot", "rating_midroog", "calc"]


def write_coverage(data, raw, res, sched):
    """data/calc/coverage.csv — סדרות קונצרניות שגלובס פרסם להן מחיר בלי מח"מ/תשואה ביום האחרון, והאם חושבו."""
    last = raw["date"].max()
    r = raw[(raw["date"] == last) & raw["source"].isin(PAIRS) & (raw["duration"].isna() | raw["ytm_gross"].isna())]
    done = set(res[(res["date"] == last) & (res["calc"] == 1)]["security_id"].astype(str))
    reasons = {}
    sk = Path(__file__).resolve().parent / "calc" / "skipped.csv"
    if sk.exists():
        reasons = {str(a): b for a, b in pd.read_csv(sk, dtype=str, encoding="utf-8-sig")[["security_id", "reason"]].itertuples(index=False)}
    rows = []
    for x in r.itertuples():
        sid = str(x.security_id)
        if sid in done: st, why = "חושב", "לוח תשלומים מהמודל (calc/schedules.csv)"
        elif sid in sched: st, why = "לא חושב", "יש לוח אך החישוב נכשל או שהתוצאה מחוץ לטווח (מח\"מ < 0.3 / תשואה > 50%)"
        elif sid in reasons: st, why = "לא חושב", reasons[sid]
        else: st, why = "לא חושב", "אין מודל מנפיק — אין לוח תשלומים"
        rows.append({"date": last, "security_id": sid, "name": x.name, "source": x.source, "last_price": x.last_price, "status": st, "reason": why})
    (Path(data) / "calc").mkdir(exist_ok=True)
    pd.DataFrame(rows, columns=["date", "security_id", "name", "source", "last_price", "status", "reason"]).sort_values(["status", "source", "name"]).to_csv(
        Path(data) / "calc" / "coverage.csv", index=False, encoding="utf-8-sig")
    print(f"calc: {len(done)} חושבו, {len(rows) - len(done)} לא חושבו (data/calc/coverage.csv)")


def main(data="data"):
    data = Path(data)
    files = sorted((data / "history").glob("*.csv"))
    if not files:
        print("אין היסטוריה"); return 1
    h = pd.concat([pd.read_csv(f, dtype=str, encoding="utf-8-sig") for f in files], ignore_index=True)
    for c in ("duration", "ytm_gross", "ytm_net"):
        h[c] = pd.to_numeric(h[c], errors="coerce")
    raw = h.copy()
    for c in ("last_price",):
        raw[c] = pd.to_numeric(raw[c], errors="coerce")
    sched = load_schedules()
    ok = h["ytm_gross"].notna() & h["duration"].notna() & (h["ytm_gross"] > 0) & (h["ytm_gross"] <= MAX_YTM) & (h["duration"] >= MIN_DUR)
    h = h[ok]
    out = []
    rawd = dict(tuple(raw.groupby("date")))
    for day, g in h.groupby("date"):
        rd = rawd.get(day, raw.iloc[0:0])
        for corp, gov in PAIRS.items():
            curve = g[g.source == gov].groupby("duration", as_index=False)["ytm_gross"].mean().sort_values("duration")
            c = g[g.source == corp].copy()
            if len(curve) < 2:
                continue
            x, y = curve["duration"].to_numpy(), curve["ytm_gross"].to_numpy()
            if not c.empty:
                c["gov_ytm"] = np.interp(c["duration"], x, y)
                c["extrap"] = ((c["duration"] < x[0]) | (c["duration"] > x[-1])).astype(int)
                c["spread"] = c["ytm_gross"] - c["gov_ytm"]
                c["curve"] = gov
                c["calc"] = 0
                out.append(c[COLS])
            if corp == "corp_unlinked" and sched:
                k = calc_rows(day, rd[rd.source == corp], sched, x, y, gov, MAX_YTM, MIN_DUR)
                if not k.empty:
                    out.append(k[COLS])
    if not out:
        print("לא חושב דבר"); return 1
    res = pd.concat(out, ignore_index=True)
    for c in ("ytm_gross", "gov_ytm", "spread"):
        res[c] = res[c].round(3)
    res["calc"] = res["calc"].astype(int)
    write_coverage(data, raw, res, sched)
    (data / "spreads").mkdir(exist_ok=True)
    for yr, g in res.groupby(res["date"].str[:4]):
        g.to_csv(data / "spreads" / f"{yr}.csv", index=False, encoding="utf-8-sig")
        print(f"{yr}: {len(g)} שורות, {g['date'].nunique()} ימים")
    return 0


if __name__ == "__main__":
    sys.exit(main(*sys.argv[1:]))
