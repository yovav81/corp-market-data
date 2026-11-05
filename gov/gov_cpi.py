"""gov_cpi.py — תוואי המדדים לקובץ הממשלתי (data/gov/cpi_path.csv).

מקור: data/short/cpi_expected.csv (short/cpi_expected.py: מדדים ידועים מהלמ"ס + חודשים צפויים = אינפלציה שנתית × עונתיות).
פלט: ym, pub_date, kind (ידוע/צפוי), index, monthly_pct, season_pct, cum_factor, source, as_of
  season_pct = monthly_pct − ממוצע 12 החודשים הצפויים הראשונים (לחודשים צפויים; לידועים 0).
  באקסל, לכל תחזית שנתית נבחרת: חודש_צפוי = תחזית/12 + season_pct — כך התוואי מתאים לכל מקור תחזית, לא רק ל-BEI.
"""
from __future__ import annotations
import sys
from pathlib import Path
import pandas as pd

COLS = ["ym", "pub_date", "kind", "index", "monthly_pct", "season_pct", "cum_factor", "source", "as_of"]


def build_cpi_path(data: Path | str, horizon_months: int = 24) -> pd.DataFrame:
    src = Path(data) / "short" / "cpi_expected.csv"
    df = pd.read_csv(src, encoding="utf-8-sig", dtype={"ym": str, "pub_date": str, "kind": str, "source": str, "as_of": str})
    df = df.sort_values("ym").reset_index(drop=True)
    df["monthly_pct"] = pd.to_numeric(df["monthly_pct"], errors="coerce")
    df["index"] = pd.to_numeric(df["index"], errors="coerce")
    df["cum_factor"] = pd.to_numeric(df["cum_factor"], errors="coerce")
    known = df[df["kind"] == "ידוע"]; exp = df[df["kind"] == "צפוי"].head(horizon_months).copy()
    mean12 = exp["monthly_pct"].head(12).mean() if len(exp) else 0.0
    exp["season_pct"] = (exp["monthly_pct"] - mean12).round(6)
    known = known.copy(); known["season_pct"] = 0.0
    out = pd.concat([known, exp], ignore_index=True)
    for c in COLS:
        if c not in out.columns: out[c] = ""
    return out[COLS]


if __name__ == "__main__":
    data = Path(sys.argv[1] if len(sys.argv) > 1 else "data")
    out = build_cpi_path(data); out.to_csv(data / "gov" / "cpi_path.csv", index=False, encoding="utf-8-sig")
    print(f"cpi_path: {len(out)} חודשים ({(out.kind == 'ידוע').sum()} ידועים, {(out.kind == 'צפוי').sum()} צפויים)")
