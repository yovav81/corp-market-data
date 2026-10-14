#!/usr/bin/env python3
"""חישוב הפרש ממשלתי יומי לכל אג"ח קונצרנית, מנתוני הגלם ב-data/history.

צמוד  -> מול עקום ממשלתי צמוד (gov_cpi);  שקלי -> מול עקום ממשלתי שקלי קבוע (gov_unlinked_fixed).
אינטרפולציה ליניארית לפי מח"מ; מחוץ לטווח העקום — ערך קצה, ומסומן extrap=1.
מחושב מחדש לכל ההיסטוריה בכל הרצה, לכן שינוי שיטה משפיע גם על העבר.
פלט: data/spreads/<שנה>.csv
"""
import sys
from pathlib import Path
import numpy as np
import pandas as pd

PAIRS = {"corp_cpi": "gov_cpi", "corp_unlinked": "gov_unlinked_fixed"}
MAX_YTM, MIN_DUR = 50.0, 0.3   # סינון אג"ח במצוקה / מח"מ קצר מאוד


def main(data="data"):
    data = Path(data)
    files = sorted((data / "history").glob("*.csv"))
    if not files:
        print("אין היסטוריה"); return 1
    h = pd.concat([pd.read_csv(f, dtype=str, encoding="utf-8-sig") for f in files], ignore_index=True)
    for c in ("duration", "ytm_gross", "ytm_net"):
        h[c] = pd.to_numeric(h[c], errors="coerce")
    ok = h["ytm_gross"].notna() & h["duration"].notna() & (h["ytm_gross"] > 0) & (h["ytm_gross"] <= MAX_YTM) & (h["duration"] >= MIN_DUR)
    h = h[ok]
    out = []
    for day, g in h.groupby("date"):
        for corp, gov in PAIRS.items():
            curve = g[g.source == gov].groupby("duration", as_index=False)["ytm_gross"].mean().sort_values("duration")
            c = g[g.source == corp].copy()
            if len(curve) < 2 or c.empty:
                continue
            x, y = curve["duration"].to_numpy(), curve["ytm_gross"].to_numpy()
            c["gov_ytm"] = np.interp(c["duration"], x, y)
            c["extrap"] = ((c["duration"] < x[0]) | (c["duration"] > x[-1])).astype(int)
            c["spread"] = c["ytm_gross"] - c["gov_ytm"]
            c["curve"] = gov
            out.append(c[["date", "source", "security_id", "name", "duration", "ytm_gross",
                          "gov_ytm", "spread", "extrap", "curve", "rating_maalot", "rating_midroog"]])
    if not out:
        print("לא חושב דבר"); return 1
    res = pd.concat(out, ignore_index=True)
    for c in ("ytm_gross", "gov_ytm", "spread"):
        res[c] = res[c].round(3)
    (data / "spreads").mkdir(exist_ok=True)
    for yr, g in res.groupby(res["date"].str[:4]):
        g.to_csv(data / "spreads" / f"{yr}.csv", index=False, encoding="utf-8-sig")
        print(f"{yr}: {len(g)} שורות, {g['date'].nunique()} ימים")
    return 0


if __name__ == "__main__":
    sys.exit(main(*sys.argv[1:]))
