#!/usr/bin/env python3
"""סטטיסטיקת מרווחים היסטורית לכל אג"ח — מוחלטת ויחסית לקבוצת ההשוואה.

    python hist/hist_spreads.py [data]

מקורות (שניהם מרווח באחוזים מול אג"ח ממשלתית באותו מח"מ; נבדקו זה מול זה ל-21.10.26 — חציון הפרש 0.00, רבעונים ±2 נ"ב):
  1. data/hist/quarterly_spreads.csv — נקודות רבעוניות 2023–2026 (מגמה), רק ok=1
  2. data/spreads/<שנה>.csv            — מרווח יומי מגלובס (מ-13.10.2026 ואילך)

מדגם ההיסטוריה לכל אג"ח: הנקודות הרבעוניות + נקודת סוף-חודש אחת לכל חודש מנתוני גלובס (חודשים שהסתיימו).
כך המדגם הופך בהדרגה מרבעוני לחודשי בלי שהימים האחרונים ישלטו בממוצע. חודש שיש לו גם נקודה רבעונית וגם
סוף-חודש מגלובס — גלובס גובר. "נוכחי" = יום המסחר האחרון בגלובס (לא חלק מהמדגם).

קבוצת השוואה (peer): הצמדה × קבוצת דירוג (AAA/AA/A/BBB-/NR) × טווח מח"מ (0-1.5/1.5-3/3-5/5+).
הדירוג נלקח כפי שהוא היום (קירוב: מניח שהדירוג לא השתנה לאורך 3 השנים).
מרווח יחסי = מרווח − חציון הקבוצה באותו תאריך; Z יחסי מבודד את סיפור המנפיק מתנועת השוק.

פלט:
  data/hist/spread_stats.csv  — שורה לכל אג"ח (ראו COLS)
  data/hist/peer_medians.csv  — חציון מרווח לכל קבוצה בכל תאריך (מדד תנאי אשראי)
  data/hist/sample_points.csv — כל נקודות המדגם (לגרפים/בדיקה)
  data/hist/credit_conditions.csv — מדד תנאי אשראי: חציון מרווח לפי הצמדה × קבוצת דירוג (כל המח"מים) בכל תאריך, + מס' סדרות
"""
import json
import sys
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))
from ratings import normalize, lower_of, group, linkage, dur_bucket  # noqa: E402

MIN_N = 4               # מינימום נקודות היסטוריה לחישוב Z
MAX_SPREAD, MIN_SPREAD = 50.0, -1.0
COLS = ["security_id", "name", "issuer", "covered", "linkage", "rating_low", "rating_group", "dur_bucket", "duration",
        "cur_date", "cur_spread", "cur_ytm", "n_hist", "hist_from", "hist_to", "mean", "median", "std", "min", "min_date",
        "max", "max_date", "pct_rank", "z", "dev_bp", "chg_1d_bp", "chg_5d_bp", "chg_vs_lastq_bp", "lastq_date",
        "peer_group", "peer_n", "peer_cur_median", "rel_cur", "rel_mean", "rel_std", "rel_z", "rel_dev_bp",
        "rating_latest", "rating_hist", "rating_change", "src_q", "src_g"]


def issuer_of(name):
    """'אלוני חץ אגח יא' → 'אלוני חץ' (הקידומת לפני 'אגח'/'אג')."""
    s = str(name)
    for tok in (" אגח", " אג ", " אג'", " אג\""):
        i = s.find(tok)
        if i > 0:
            return s[:i].strip()
    parts = s.split()
    return " ".join(parts[:-1]) if len(parts) > 1 else s


def load_issuers(here):
    p = here / "maya" / "issuers.json"
    if not p.exists():
        return []
    out = []
    for it in json.loads(p.read_text(encoding="utf-8")):
        names = {it.get("name", "")} | set(it.get("bond_prefixes", []))
        out.append((it.get("key"), {n.replace("״", '"').replace("׳", "'") for n in names if n}))
    return out


def covered_key(issuer, issuers):
    s = issuer.replace("״", '"').replace("׳", "'")
    for key, names in issuers:
        for n in names:
            if s == n or s.startswith(n) or n.startswith(s) and len(s) >= 4:
                return key
    return ""


def main(data="data"):
    data = Path(data); here = Path(__file__).resolve().parent.parent
    hist_dir = data / "hist"
    q = pd.read_csv(hist_dir / "quarterly_spreads.csv", dtype={"security_id": str}, encoding="utf-8-sig")
    static = pd.read_csv(hist_dir / "bond_static.csv", dtype={"security_id": str}, encoding="utf-8-sig")
    q = q[q["ok"] == 1][["date", "security_id", "name", "spread", "duration", "ytm_gross"]].copy()
    q["src"] = "q"

    files = sorted((data / "spreads").glob("*.csv"))
    g = pd.concat([pd.read_csv(f, dtype={"security_id": str}, encoding="utf-8-sig") for f in files], ignore_index=True) if files else pd.DataFrame()
    if not g.empty:
        for c in ("spread", "duration", "ytm_gross"):
            g[c] = pd.to_numeric(g[c], errors="coerce")
        g = g[g["spread"].between(MIN_SPREAD, MAX_SPREAD)].copy()
        g["linkage_g"] = g["source"].map(linkage)
        g["rating_g"] = [lower_of(a, b) for a, b in zip(g["rating_maalot"], g["rating_midroog"])]
        g = g.sort_values(["security_id", "date"])
        cur_date = g["date"].max()
        cur = g[g["date"] == cur_date].drop_duplicates("security_id").set_index("security_id")
        dates = sorted(g["date"].unique())
        prev1 = g[g["date"] == dates[-2]].drop_duplicates("security_id").set_index("security_id")["spread"] if len(dates) > 1 else pd.Series(dtype=float)
        prev5 = g[g["date"] == dates[-6]].drop_duplicates("security_id").set_index("security_id")["spread"] if len(dates) > 5 else pd.Series(dtype=float)
        first_rating = g.drop_duplicates("security_id").set_index("security_id")["rating_g"]
        # נקודת סוף-חודש לכל חודש שהסתיים
        g["ym"] = g["date"].str[:7]
        closed = g[g["ym"] < cur_date[:7]]
        me = closed.sort_values("date").groupby(["security_id", "ym"]).tail(1)
        me = me[["date", "security_id", "name", "spread", "duration", "ytm_gross"]].copy(); me["src"] = "g"
    else:
        cur_date, cur, prev1, prev5, first_rating = None, pd.DataFrame(), pd.Series(dtype=float), pd.Series(dtype=float), pd.Series(dtype=str)
        me = pd.DataFrame(columns=list(q.columns))

    # מדגם: רבעוני + סוף-חודש גלובס; באותו חודש — גלובס גובר
    sample = pd.concat([q, me], ignore_index=True)
    sample["ym"] = sample["date"].str[:7]
    sample = sample.sort_values(["security_id", "ym", "src"])           # 'g' < 'q' → גלובס ראשון
    sample = sample.drop_duplicates(["security_id", "ym"], keep="first").drop(columns="ym")

    # מאפיינים סטטיים: מההיסטוריה, ואם אין — מגלובס
    st = static.set_index("security_id")
    ids = sorted(set(sample["security_id"]) | set(cur.index))
    meta = pd.DataFrame(index=ids)
    meta["name"] = [cur["name"].get(i) if i in cur.index else st["name"].get(i) for i in ids]
    meta["linkage"] = [st["linkage"].get(i) if i in st.index and pd.notna(st["linkage"].get(i)) else (cur["linkage_g"].get(i) if i in cur.index else None) for i in ids]
    meta["rating_latest"] = [cur["rating_g"].get(i) if i in cur.index else None for i in ids]
    meta["rating_low"] = [meta["rating_latest"][i] if pd.notna(meta["rating_latest"][i]) else (st["rating_low"].get(i) if i in st.index else None) for i in ids]
    meta["rating_group"] = meta["rating_low"].map(group)
    meta["is_gov"] = [int(st["is_gov"].get(i, 0)) if i in st.index else 0 for i in ids]
    meta = meta[meta["is_gov"] == 0]
    issuers = load_issuers(here)
    meta["issuer"] = meta["name"].map(issuer_of)
    meta["covered"] = meta["issuer"].map(lambda s: covered_key(s, issuers))

    # חציון קבוצת השוואה לכל תאריך (על המדגם + היום הנוכחי)
    pts = sample.merge(meta[["linkage", "rating_group"]], left_on="security_id", right_index=True, how="inner")
    if not cur.empty:
        c = cur.reset_index()[["date", "security_id", "name", "spread", "duration", "ytm_gross"]].copy(); c["src"] = "cur"
        pts = pd.concat([pts, c.merge(meta[["linkage", "rating_group"]], left_on="security_id", right_index=True, how="inner")], ignore_index=True)
    pts["dur_bucket"] = pts["duration"].map(dur_bucket)
    pts = pts[pts["linkage"].notna()]
    # נקודה בלי מח"מ (למשל 01/2024 בקובץ ההיסטוריה) נשארת בהיסטוריה של הסדרה, אך בלי קבוצת השוואה
    pts["peer_group"] = np.where(pts["dur_bucket"].notna(), pts["linkage"] + "|" + pts["rating_group"] + "|" + pts["dur_bucket"].astype(str), None)
    pm = pts.groupby(["date", "peer_group"])["spread"].agg(peer_median="median", peer_n="size").reset_index()
    pts = pts.merge(pm, on=["date", "peer_group"], how="left")
    pts["rel"] = np.where(pts["peer_n"] >= 3, pts["spread"] - pts["peer_median"], np.nan)

    rows = []
    for sid, grp in pts.groupby("security_id"):
        h = grp[grp["src"] != "cur"].sort_values("date")
        c = grp[grp["src"] == "cur"]
        m = meta.loc[sid]
        r = {"security_id": sid, "name": m["name"], "issuer": m["issuer"], "covered": m["covered"], "linkage": m["linkage"],
             "rating_low": m["rating_low"], "rating_group": m["rating_group"], "rating_latest": m["rating_latest"],
             "n_hist": len(h), "src_q": int((h["src"] == "q").sum()), "src_g": int((h["src"] == "g").sum())}
        if len(h):
            r.update(hist_from=h["date"].iloc[0], hist_to=h["date"].iloc[-1], mean=h["spread"].mean(), median=h["spread"].median(),
                     std=h["spread"].std(ddof=1) if len(h) > 1 else np.nan, min=h["spread"].min(), min_date=h.loc[h["spread"].idxmin(), "date"],
                     max=h["spread"].max(), max_date=h.loc[h["spread"].idxmax(), "date"],
                     rel_mean=h["rel"].mean(), rel_std=h["rel"].std(ddof=1) if h["rel"].notna().sum() > 1 else np.nan)
            lastq = h[h["src"] == "q"]
            if len(lastq):
                r["lastq_date"] = lastq["date"].iloc[-1]
        if len(c):
            c = c.iloc[0]
            r.update(cur_date=c["date"], cur_spread=c["spread"], cur_ytm=c["ytm_gross"], duration=c["duration"], dur_bucket=c["dur_bucket"],
                     peer_group=c["peer_group"], peer_n=c["peer_n"], peer_cur_median=c["peer_median"], rel_cur=c["rel"])
            if sid in prev1.index:
                r["chg_1d_bp"] = (c["spread"] - prev1[sid]) * 100
            if sid in prev5.index:
                r["chg_5d_bp"] = (c["spread"] - prev5[sid]) * 100
            if len(h):
                r["dev_bp"] = (c["spread"] - r["mean"]) * 100
                r["pct_rank"] = (h["spread"] < c["spread"]).mean()
                if len(h) >= MIN_N and r["std"] > 0:
                    r["z"] = (c["spread"] - r["mean"]) / r["std"]
                if pd.notna(c["rel"]) and pd.notna(r.get("rel_std")) and r["rel_std"] > 0 and h["rel"].notna().sum() >= MIN_N:
                    r["rel_z"] = (c["rel"] - r["rel_mean"]) / r["rel_std"]
                    r["rel_dev_bp"] = (c["rel"] - r["rel_mean"]) * 100
                lq = h[h["src"] == "q"]
                if len(lq):
                    r["chg_vs_lastq_bp"] = (c["spread"] - lq["spread"].iloc[-1]) * 100
            # שינוי דירוג — רק כפי שנצפה בנתוני גלובס לאורך זמן (דירוגי מגמה בקובץ ההיסטוריה שונים לעיתים בכתיב/ברמת המכשיר)
            fr = first_rating.get(sid); lr = m["rating_latest"]
            if isinstance(fr, str) and isinstance(lr, str) and fr != lr:
                r["rating_change"] = f"{fr}→{lr}"
            r["rating_hist"] = st["rating_low"].get(sid) if sid in st.index else None
        else:
            r["dur_bucket"] = dur_bucket(h["duration"].iloc[-1]) if len(h) else None
        rows.append(r)
    out = pd.DataFrame(rows).reindex(columns=COLS)
    for c in ("mean", "median", "std", "min", "max", "rel_mean", "rel_std", "rel_cur", "peer_cur_median", "cur_spread", "cur_ytm", "duration"):
        out[c] = pd.to_numeric(out[c]).round(3)
    for c in ("z", "rel_z", "pct_rank"):
        out[c] = pd.to_numeric(out[c]).round(2)
    for c in ("dev_bp", "chg_1d_bp", "chg_5d_bp", "chg_vs_lastq_bp", "rel_dev_bp"):
        out[c] = pd.to_numeric(out[c]).round(0)
    out = out.sort_values(["rel_z", "z"], ascending=False, na_position="last")
    out.to_csv(hist_dir / "spread_stats.csv", index=False, encoding="utf-8-sig")
    pm.sort_values(["peer_group", "date"]).to_csv(hist_dir / "peer_medians.csv", index=False, encoding="utf-8-sig")
    pts.sort_values(["security_id", "date"]).to_csv(hist_dir / "sample_points.csv", index=False, encoding="utf-8-sig")
    # מדד תנאי אשראי — חציון לפי הצמדה × קבוצת דירוג, בלי פיצול מח"מ (מספיק סדרות בכל נקודה)
    wide = pts[pts["rating_group"] != "NR"].copy()
    wide["series"] = wide["linkage"] + "|" + wide["rating_group"]
    cc = wide.groupby(["date", "series"])["spread"].agg(median="median", n="size").reset_index()
    cc = cc[cc["n"] >= 5]
    cc["median"] = cc["median"].round(3)
    cc.sort_values(["series", "date"]).to_csv(hist_dir / "credit_conditions.csv", index=False, encoding="utf-8-sig")
    print(f"אג\"ח: {len(out)}, עם נוכחי: {int(out['cur_spread'].notna().sum())}, עם Z: {int(out['z'].notna().sum())}, "
          f"עם Z יחסי: {int(out['rel_z'].notna().sum())}, מסוקרות: {int((out['covered'] != '').sum())}, נוכחי ל-{cur_date}")
    return 0


if __name__ == "__main__":
    sys.exit(main(*sys.argv[1:]))
