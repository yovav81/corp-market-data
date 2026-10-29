"""
short_yields.py — חישוב מחדש של תשואה ומח"מ לאג"ח קונצרניות עם מח"מ מתחת ל-1, מתוך לוח הסילוקין בפועל
(data/short/schedules.csv), נתוני גלובס (data/latest.csv) ותוואי המדדים הצפויים (data/short/cpi_expected.csv).

פלט:
  data/short/short_bonds.csv  שורה לסדרה: נתוני גלובס, תזרים מחדש, תשואה ריאלית מחושבת, תשואה נומינלית צפויה, מח"מ מחושב,
                              פער מול גלובס, השוואה לעקומים הממשלתיים (שקלי/צמוד) באותו מח"מ, דגלים
  data/short/cashflows.csv    שורה לכל תשלום עתידי לסדרה (ל-100 ע.נ. נותרת): נומינלי, ריאלי (×L), מקדם מדד, נומינלי צפוי
הרצה: python short/short_yields.py data
"""
from __future__ import annotations
import datetime as dt, json, os, sys
from pathlib import Path
import numpy as np, pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from short.engine import Bond, Flow, FREQ_MONTHS, compute, flows_from_schedule

DUR_MAX = float(os.environ.get("SHORT_DUR", "1.0"))
GAP_BP = float(os.environ.get("SHORT_GAP_BP", "15"))
OUT_COLS = ["date", "security_id", "name", "source", "linkage", "redemption_type", "coupon", "coupon_type", "coupon_freq",
            "maturity", "last_price", "adjusted_value", "volume_k_nis", "rating_maalot", "rating_midroog",
            "dur_globes", "ytm_globes", "gov_ytm_globes", "spread_globes",
            "L", "n_flows", "next_pay_date", "next_pay_amt", "next_principal_pct", "next_ex", "final_pay_date",
            "dur_calc", "ytm_real_calc", "ytm_nominal_exp", "infl_period_pct", "infl_annual_pct",
            "gap_bp", "gov_cpi_at_dur", "gov_ils_at_dur", "spread_real_bp", "spread_nominal_bp", "flags", "months", "status"]


def _d(s):
    if s is None or (isinstance(s, float) and np.isnan(s)): return None
    try: return pd.to_datetime(str(s)).date()
    except Exception: return None


def curve(mk: pd.DataFrame, src: str):
    g = mk[(mk["source"] == src)].copy()
    g["duration"] = pd.to_numeric(g["duration"], errors="coerce"); g["ytm_gross"] = pd.to_numeric(g["ytm_gross"], errors="coerce")
    g = g.dropna(subset=["duration", "ytm_gross"]).groupby("duration", as_index=False)["ytm_gross"].mean().sort_values("duration")
    return g["duration"].to_numpy(), g["ytm_gross"].to_numpy()


def main(data="data"):
    d = Path(data); sd = d / "short"
    mk = pd.read_csv(d / "latest.csv", encoding="utf-8-sig")
    mk["security_id"] = pd.to_numeric(mk["security_id"], errors="coerce")
    date = _d(mk["date"].max())
    terms = pd.read_csv(sd / "bond_terms.csv", encoding="utf-8-sig"); terms["security_id"] = pd.to_numeric(terms["security_id"], errors="coerce")
    sched = pd.read_csv(sd / "schedules.csv", encoding="utf-8-sig"); sched["security_id"] = pd.to_numeric(sched["security_id"], errors="coerce")
    cpi = pd.read_csv(sd / "cpi_expected.csv", encoding="utf-8-sig")
    exp = {}
    for _, r in cpi.iterrows():
        if str(r["kind"]).startswith("צפוי"):
            y, m = str(r["ym"]).split("-"); exp[(int(y), int(m))] = float(r["monthly_pct"]) / 100
    # עקומים ממשלתיים
    xi, yi = curve(mk, "gov_cpi"); xs, ys = curve(mk, "gov_unlinked_fixed")
    # מרווח גלובס (מהחישוב היומי)
    try:
        sp = pd.read_csv(d / "spreads" / f"{date.year}.csv", encoding="utf-8-sig", dtype={"security_id": str})
        sp = sp[sp["date"] == date.isoformat()]; sp["security_id"] = pd.to_numeric(sp["security_id"], errors="coerce")
        sp = sp.set_index("security_id")[["gov_ytm", "spread"]]
    except Exception:
        sp = pd.DataFrame(columns=["gov_ytm", "spread"])
    corp = mk[mk["source"].astype(str).str.startswith("corp") & (pd.to_numeric(mk["duration"], errors="coerce") < DUR_MAX)]
    out, cfs = [], []
    for _, r in corp.sort_values(["source", "name"]).iterrows():
        sid = int(r["security_id"]); t = terms[terms["security_id"] == sid]
        row = {"date": date.isoformat(), "security_id": sid, "name": r["name"], "source": r["source"], "maturity": r.get("maturity"),
               "last_price": r.get("last_price"), "adjusted_value": r.get("adjusted_value"), "volume_k_nis": r.get("volume_k_nis"),
               "rating_maalot": r.get("rating_maalot"), "rating_midroog": r.get("rating_midroog"),
               "dur_globes": r.get("duration"), "ytm_globes": r.get("ytm_gross"),
               "gov_ytm_globes": sp["gov_ytm"].get(sid) if len(sp) else None, "spread_globes": sp["spread"].get(sid) if len(sp) else None}
        flags = []
        price = pd.to_numeric(r.get("last_price"), errors="coerce"); adj = pd.to_numeric(r.get("adjusted_value"), errors="coerce")
        if pd.isna(price) or price <= 0:
            row.update(status="אין מחיר"); out.append(row); continue
        if t.empty or t.iloc[0]["status"] != "ok":
            row.update(status="אין לוח סילוקין" if t.empty else f"לוח: {t.iloc[0]['status']}"); out.append(row); continue
        t = t.iloc[0]
        s = sched[sched["security_id"] == sid]
        rows = [{"pay_date": _d(x["pay_date"]), "principal_amt": x["principal_amt"], "interest_amt": x["interest_amt"]} for _, x in s.iterrows()]
        rows = [x for x in rows if x["pay_date"]]
        if not rows or sum(float(x["principal_amt"] or 0) for x in rows) <= 0:
            row.update(status="לוח ריק"); out.append(row); continue
        linkage = str(t.get("linkage") or ("מדד" if r["source"] == "corp_cpi" else "שקלי"))
        if r["source"] == "corp_cpi" and "מדד" not in linkage: linkage = "מדד"
        freq = FREQ_MONTHS.get(str(t.get("coupon_freq") or "").strip(), 6 if r["source"] != "corp_fx" else 6)
        coupon = float(pd.to_numeric(t.get("coupon"), errors="coerce") or 0)
        ctype = str(t.get("coupon_type") or "")
        b = Bond(sid, r["name"], linkage, coupon, ctype, freq, flows_from_schedule(rows, _d(t.get("next_ex"))))
        linked = "מדד" in linkage
        res = compute(b, float(price), float(adj) if ((linked or r["source"] == "corp_fx") and not pd.isna(adj)) else None, date, exp)
        row.update(linkage=linkage, redemption_type=t.get("redemption_type"), coupon=coupon, coupon_type=ctype, coupon_freq=t.get("coupon_freq"),
                   L=round(res.L, 5), n_flows=len(res.flows_real))
        if not res.flows_real:
            row.update(status="כל התשלומים אחרי אקס"); out.append(row); continue
        live = [f for f in b.flows if f.included]
        nx = live[0]
        row.update(next_pay_date=nx.date.isoformat(), next_pay_amt=round(nx.nominal * res.L, 4), next_principal_pct=round(nx.principal, 2),
                   next_ex=(nx.ex_date.isoformat() if nx.ex_date else None), final_pay_date=live[-1].date.isoformat(),
                   dur_calc=round(res.dur, 3) if res.dur is not None else None,
                   ytm_real_calc=round(res.y_real * 100, 3) if res.y_real is not None else None,
                   ytm_nominal_exp=round(res.y_nominal * 100, 3) if res.y_nominal is not None else None,
                   infl_period_pct=round(res.infl_period * 100, 3) if res.infl_period is not None else None,
                   infl_annual_pct=round(res.infl_annual * 100, 3) if res.infl_annual is not None else None,
                   months=" ".join(f"{y}-{m:02d}" + ("" if e is not None else "?") for (y, m), e in dict((k, v) for k, v in res.months_used).items()))
        yg = pd.to_numeric(r.get("ytm_gross"), errors="coerce")
        if res.y_real is not None and not pd.isna(yg):
            gap = (res.y_real * 100 - yg) * 100; row["gap_bp"] = round(gap, 1)
            if abs(gap) > GAP_BP: flags.append(f"פער מול גלובס {gap:+.0f} נ״ב")
        if res.dur is not None:
            gi = float(np.interp(res.dur, xi, yi)) if len(xi) else None; gs = float(np.interp(res.dur, xs, ys)) if len(xs) else None
            row.update(gov_cpi_at_dur=round(gi, 3) if gi is not None else None, gov_ils_at_dur=round(gs, 3) if gs is not None else None)
            if res.y_real is not None:
                if linked and gi is not None: row["spread_real_bp"] = round((res.y_real * 100 - gi) * 100, 0)
                if not linked and gs is not None: row["spread_real_bp"] = round((res.y_real * 100 - gs) * 100, 0)
            if res.y_nominal is not None and gs is not None and r["source"] != "corp_fx":
                row["spread_nominal_bp"] = round((res.y_nominal * 100 - gs) * 100, 0)
        if res.excluded: flags.append("תשלום אחרי אקס הושמט")
        if "משתנה" in ctype: flags.append("ריבית משתנה — חושב לפי הריבית הנוכחית")
        if r["source"] == "corp_fx": flags.append("צמוד מט״ח — ללא תחזית שער")
        if (not pd.isna(yg) and yg > 50) or price < 70: flags.append("מצוקה")
        if any(e is None for _, e in res.months_used): flags.append("חסר מדד צפוי לחלק מהחודשים")
        if res.dur is not None and nx.date and (nx.date - date).days <= 10: flags.append("תשלום בתוך 10 ימים")
        row.update(flags="; ".join(flags), status="ok")
        out.append(row)
        for i, (f, (dd, real), (_, nom), k) in enumerate(zip(live, res.flows_real, res.flows_nominal, res.factors)):
            cfs.append({"security_id": sid, "name": r["name"], "slot": i + 1, "pay_date": dd.isoformat(), "principal": round(f.principal, 4),
                        "interest": round(f.interest, 4), "nominal": round(f.nominal, 4), "L": round(res.L, 5), "real": round(real, 4),
                        "index_factor": round(k, 6), "nominal_exp": round(nom, 4), "ex_date": f.ex_date.isoformat() if f.ex_date else None})
    df = pd.DataFrame(out, columns=OUT_COLS)
    df.to_csv(sd / "short_bonds.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(cfs).to_csv(sd / "cashflows.csv", index=False, encoding="utf-8-sig")
    ok = df[df["status"] == "ok"]
    print(f"{date}: {len(df)} סדרות מתחת למח״מ {DUR_MAX}, חושבו {len(ok)}; פער חציוני מול גלובס {ok['gap_bp'].abs().median():.1f} נ״ב; "
          f"{(ok['gap_bp'].abs() > GAP_BP).sum()} עם פער מעל {GAP_BP:.0f} נ״ב")
    print(df["status"].value_counts().to_string())


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "data")
