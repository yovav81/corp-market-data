#!/usr/bin/env python3
"""מרווחים חריגים שנפתחו מול ההיסטוריה — טבלה יומית + הודעת טלגרם. כל המנפיקים בשוק, לא רק המסוקרים.

    python hist/spread_alerts.py [data] [--no-state]

תנאי התראה (ספים ב-PARAMS; אפשר לשנות דרך משתני סביבה ALERT_<שם>):
  * יש מרווח נוכחי (גלובס) ולפחות MIN_N נקודות היסטוריה
  * Z מוחלט ≥ Z_ABS  וגם  סטייה מהממוצע ≥ DEV_BP נ"ב        — נפתח מול ההיסטוריה של עצמו
  * Z יחסי ≥ Z_REL                                            — לא מוסבר ע"י תנועה של כל קבוצת ההשוואה
  * מח"מ ≥ MIN_DUR, מרווח נוכחי ≤ MAX_SPREAD (לא אג"ח במצוקה); ללא דירוג — רק אם INCLUDE_NR=1 (נשארות בטבלה, לא בהודעה)
סינון "הסבר ידוע": הורדת דירוג שנצפתה בגלובס → נפסל. דיווח חדש במאיה ב-MAYA_DAYS הימים האחרונים (מנפיקים מסוקרים בלבד,
מ-maya/state.json) → נשאר אך מסומן "לבדוק דיווח". לשאר המנפיקים אין לנו מעקב דיווחים — ההודעה אומרת זאת.
מניעת חזרות: data/hist/alert_state.json — התראה חוזרת רק אחרי REPEAT_DAYS ימים או אם ה-Z עלה ב-REPEAT_DZ.

פלט: data/hist/alerts_latest.csv (הטבלה המלאה, כולל מה שנשלח בעבר), data/hist/alert_message.txt (הודעת טלגרם; ריק = אין מה לשלוח)
"""
import datetime as dt
import json
import os
import sys
from pathlib import Path

import pandas as pd

PARAMS = dict(MIN_N=6, Z_ABS=1.5, DEV_BP=40, Z_REL=1.0, MIN_DUR=1.0, MAX_SPREAD=12.0, INCLUDE_NR=0, MAYA_DAYS=14,
              REPEAT_DAYS=7, REPEAT_DZ=0.5, MAX_LINES=30)
for k in list(PARAMS):
    v = os.environ.get(f"ALERT_{k}")
    if v:
        PARAMS[k] = type(PARAMS[k])(v)

LINK = {"cpi": "צמוד", "ils": "שקלי", "fx": "מט\"ח"}


def recent_maya(here, days):
    """{מפתח מנפיק: תאריך הדיווח האחרון} עבור דיווחים ב-days הימים האחרונים."""
    p = here / "maya" / "state.json"; q = here / "maya" / "issuers.json"
    if not (p.exists() and q.exists()):
        return {}
    state = json.loads(p.read_text(encoding="utf-8"))
    key_of = {str(i.get("maya_id")): i.get("key") for i in json.loads(q.read_text(encoding="utf-8"))}
    cutoff = (dt.date.today() - dt.timedelta(days=days)).isoformat()
    out = {}
    for cid, v in state.items():
        last = max((d for d in (v.get("seen") or {}).values() if d), default=None)
        if last and last >= cutoff and key_of.get(str(cid)):
            out[key_of[str(cid)]] = last
    return out


def fmt_row(r):
    cov = "★" if isinstance(r["covered"], str) and r["covered"] else ""
    rating = r["rating_low"] if isinstance(r["rating_low"], str) else "NR"
    flag = ""
    if r.get("maya_recent"):
        flag = f" ⚠דיווח {r['maya_recent'][5:]}"
    return (f"{cov}<b>{r['name']}</b> ({r['rating_low'] or 'NR'}, {LINK.get(r['linkage'], '')}, מח\"מ {r['duration']:.1f}): "
            f"{r['cur_spread']:.2f}% מול ממוצע {r['mean']:.2f}% (+{r['dev_bp']:.0f} נ\"ב, Z {r['z']:.1f}, יחסי {r['rel_z']:.1f}, "
            f"אחוזון {r['pct_rank']*100:.0f}){flag}")


def main(data="data", *flags):
    P = PARAMS
    data = Path(data); here = Path(__file__).resolve().parent.parent
    s = pd.read_csv(data / "hist" / "spread_stats.csv", dtype={"security_id": str}, encoding="utf-8-sig")
    s["covered"] = s["covered"].fillna("")
    ok = (s["cur_spread"].notna() & (s["n_hist"] >= P["MIN_N"]) & (s["z"] >= P["Z_ABS"]) & (s["dev_bp"] >= P["DEV_BP"])
          & (s["rel_z"] >= P["Z_REL"]) & (s["duration"] >= P["MIN_DUR"]) & (s["cur_spread"] <= P["MAX_SPREAD"]))
    a = s[ok].copy()
    # הורדת דירוג שנצפתה — הסבר ידוע → החוצה (נשמר בטבלה עם סיבה)
    from ratings import normalize, RANK  # noqa: E402
    def downgraded(x):
        if not isinstance(x, str) or "→" not in x:
            return False
        f, t = (normalize(p) for p in x.split("→"))
        return bool(f and t and RANK[t] > RANK[f])
    a["downgrade"] = a["rating_change"].map(downgraded)
    maya = recent_maya(here, P["MAYA_DAYS"])
    a["maya_recent"] = a["covered"].map(lambda k: maya.get(k, "") if k else "")
    a["note"] = ""
    a.loc[a["downgrade"], "note"] = "הורדת דירוג — לא נשלח"
    a.loc[~a["downgrade"] & (a["maya_recent"] != ""), "note"] = "דיווח חדש במאיה — לבדוק"
    if not P["INCLUDE_NR"]:
        a.loc[a["rating_low"].isna() & (a["note"] == ""), "note"] = "ללא דירוג — לא נשלח"
    a = a.sort_values(["rel_z", "z"], ascending=False)

    # מצב התראות קודמות
    sp = data / "hist" / "alert_state.json"
    state = json.loads(sp.read_text(encoding="utf-8")) if sp.exists() and "--no-state" not in flags else {}
    today = str(a["cur_date"].iloc[0]) if len(a) else dt.date.today().isoformat()
    send = []
    for _, r in a.iterrows():
        if r["downgrade"] or (not P["INCLUDE_NR"] and not isinstance(r["rating_low"], str)):
            continue
        prev = state.get(r["security_id"])
        if prev:
            days = (dt.date.fromisoformat(today) - dt.date.fromisoformat(prev["date"])).days
            if days < P["REPEAT_DAYS"] and r["z"] < prev["z"] + P["REPEAT_DZ"]:
                continue
        send.append(r)
    a["sent_today"] = a["security_id"].isin([r["security_id"] for r in send]).astype(int)
    cols = ["security_id", "name", "issuer", "covered", "rating_low", "linkage", "duration", "cur_date", "cur_spread", "mean", "std",
            "min", "max", "dev_bp", "z", "rel_z", "rel_dev_bp", "pct_rank", "chg_5d_bp", "chg_vs_lastq_bp", "peer_cur_median",
            "rating_change", "maya_recent", "note", "sent_today"]
    a[cols].to_csv(data / "hist" / "alerts_latest.csv", index=False, encoding="utf-8-sig")

    msg = ""
    if send:
        d = dt.date.fromisoformat(today).strftime("%d.%m.%Y")
        lines = [f"📈 <b>מרווחים חריגים שנפתחו — {d}</b>",
                 f"<i>מול ממוצע 3 שנים (Z≥{P['Z_ABS']}, +{P['DEV_BP']} נ\"ב) וגם מול קבוצת הדירוג/הצמדה/מח\"מ (Z יחסי ≥{P['Z_REL']}). "
                 f"הורדות דירוג סוננו. ★ = מנפיק מסוקר. לשאר המנפיקים אין מעקב דיווחים — לבדוק במאיה לפני פעולה.</i>", ""]
        for r in send[:P["MAX_LINES"]]:
            lines.append("• " + fmt_row(r))
        if len(send) > P["MAX_LINES"]:
            lines.append(f"… ועוד {len(send) - P['MAX_LINES']} בטבלה המלאה (data/hist/alerts_latest.csv)")
        msg = "\n".join(lines)
        for r in send:
            state[r["security_id"]] = {"date": today, "z": float(r["z"]), "name": r["name"]}
        if "--no-state" not in flags:
            sp.write_text(json.dumps(state, ensure_ascii=False, indent=1), encoding="utf-8")
    (data / "hist" / "alert_message.txt").write_text(msg, encoding="utf-8")
    print(f"עומדות בתנאים: {len(a)}, מהן הורדת דירוג: {int(a['downgrade'].sum())}, נשלחות היום: {len(send)}")
    return 0


if __name__ == "__main__":
    sys.path.insert(0, str(Path(__file__).parent))
    sys.exit(main(*sys.argv[1:]))
