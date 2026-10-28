#!/usr/bin/env python3
"""פערי מרווח בין סדרות של אותו מנפיק — מה "נורמלי" לכל זוג, ומי חורג ממנו היום.

    python hist/intra_issuer.py [data]

לכל מנפיק עם 2+ סדרות שיש להן מרווח נוכחי: לכל זוג סדרות (א,ב) מחושבת סדרת ההפרש (מרווח א − מרווח ב) בתאריכי
המדגם המשותפים (data/hist/sample_points.csv), ממוצע וסטיית תקן של ההפרש, ההפרש הנוכחי, ו-Z של הזוג.
Z גבוה = סדרה א' התרחבה מול אחותה הרבה מעבר למה שהיה ביניהן ב-3 השנים (או ב' התכווצה) — או הזדמנות או סיבה (בטוחה, מח"מ, הצמדה).
זוגות בהצמדה שונה מסומנים (mixed=1) כי פער שקלי/צמוד זז עם ציפיות האינפלציה — פחות אמין.

פלט: data/hist/intra_issuer.csv — שורה לכל זוג (a = הסדרה הרחבה יותר היום), ממוין לפי |Z|.
"""
import itertools
import sys
from pathlib import Path

import numpy as np
import pandas as pd

MIN_N, MIN_DUR = 6, 0.75


def main(data="data"):
    data = Path(data) / "hist"
    st = pd.read_csv(data / "spread_stats.csv", dtype={"security_id": str}, encoding="utf-8-sig")
    pts = pd.read_csv(data / "sample_points.csv", dtype={"security_id": str}, encoding="utf-8-sig")
    hist = pts[pts["src"] != "cur"].pivot_table(index="date", columns="security_id", values="spread")
    cur = st[st["cur_spread"].notna() & (st["duration"] >= MIN_DUR)].set_index("security_id")
    rows = []
    for issuer, g in cur.groupby("issuer"):
        if len(g) < 2:
            continue
        for a, b in itertools.combinations(g.index, 2):
            if a not in hist.columns or b not in hist.columns:
                continue
            d = (hist[a] - hist[b]).dropna()
            if len(d) < MIN_N:
                continue
            cur_d = g.at[a, "cur_spread"] - g.at[b, "cur_spread"]
            if cur_d < 0:                       # a = הרחבה יותר היום
                a, b, d, cur_d = b, a, -d, -cur_d
            m, sd = d.mean(), d.std(ddof=1)
            z = (cur_d - m) / sd if sd > 0 else np.nan
            rows.append(dict(issuer=issuer, covered=g.at[a, "covered"], sid_a=a, name_a=g.at[a, "name"], sid_b=b, name_b=g.at[b, "name"],
                             link_a=g.at[a, "linkage"], link_b=g.at[b, "linkage"], mixed=int(g.at[a, "linkage"] != g.at[b, "linkage"]),
                             dur_a=g.at[a, "duration"], dur_b=g.at[b, "duration"], spread_a=g.at[a, "cur_spread"], spread_b=g.at[b, "cur_spread"],
                             n=len(d), mean_diff_bp=round(m * 100), std_bp=round(sd * 100), min_diff_bp=round(d.min() * 100), max_diff_bp=round(d.max() * 100),
                             cur_diff_bp=round(cur_d * 100), dev_bp=round((cur_d - m) * 100), z_pair=round(z, 2) if pd.notna(z) else np.nan,
                             cur_date=g.at[a, "cur_date"], z_a=g.at[a, "z"], z_b=g.at[b, "z"]))
    out = pd.DataFrame(rows)
    if len(out):
        out = out.reindex(out["z_pair"].abs().sort_values(ascending=False).index)
    out.to_csv(data / "intra_issuer.csv", index=False, encoding="utf-8-sig")
    print(f"זוגות: {len(out)}, מנפיקים: {out['issuer'].nunique() if len(out) else 0}, |Z|≥2: {int((out['z_pair'].abs() >= 2).sum()) if len(out) else 0}")
    return 0


if __name__ == "__main__":
    sys.exit(main(*sys.argv[1:]))
