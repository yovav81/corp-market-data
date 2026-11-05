"""gov_decide.py — כללי ההחלטה של קובץ הממשלתי v2 בפייתון (מראה של הנוסחאות באקסל; אפיון CORP-gov-design-spec סעיף 5).

הפונקציות טהורות (ללא I/O) מלבד decide(data) שקוראת את data/gov/*.csv וכותבת:
  decisions_latest.csv   key, value, why, as_of            — פסקי הדין של היום עם "למה"
  decisions_history.csv  date, as_of, <key...>  (שורה ליום) — מקור "ערך אתמול" בדשבורד (Power Query: GovDecisions)
יחידות: תשואות באחוזים (4.21), פערים בנ"ב, משקלים כשברים (0.55), מח"מ בשנים.
"""
from __future__ import annotations
import datetime as dt, json, os, sys
from pathlib import Path
import numpy as np, pandas as pd

TENORS = [0.25, 0.5, 1, 2, 3, 4, 5, 7, 10, 15, 20, 30]
NOMINAL_KINDS = ("שקלי", "מק\"מ"); REAL_KINDS = ("מדד",)
SOURCE_SERIES = {"ממוצע החזאים (ב״י)": "EXP_HAZ_MAD_AVG_01Y_MA", "ציפיות שוק ההון (ב״י)": "EXP_EXP_01Y_BI_MA", "חוזי אינפלציה (ב״י)": "EXP_EXP_ISSWIT1_01Y_MA"}
NLIST = 30                                   # ניירות לאפיק בסריקת המבנים (כמו רשימת 06 באקסל) — הקצרים במח"מ
STALE_DAYS = 7                               # נתון חו"ל ישן מ-7 ימים מיום הנתונים נחשב "ישן"
DECISION_KEYS = ["linked_weight", "dur_nominal", "dur_real", "dur_weighted", "struct_nominal", "struct_real", "top5_nominal", "top5_real", "foreign_5y", "foreign_10y"]


def load_settings(path: Path | str | None = None, overrides: dict | None = None) -> dict:
    p = Path(path) if path else Path(__file__).with_name("settings_default.json")
    s = {x["key"]: x["value"] for x in json.load(open(p, encoding="utf-8"))}
    for t in TENORS: s.setdefault(f"dw_{t:g}", 1 if t in (1, 2, 3, 5) else 0)
    s.setdefault("exclude", set()); s.setdefault("infl_user", {})
    if overrides: s.update(overrides)
    return s


def overrides_from_workbook(xlsx: Path | str) -> dict:
    """ההגדרות של יובב מהחוברת (p_*, משקלי החלטה, תחזית לפי טווח, 'כלול בעקום?'=לא) כ-overrides ל-load_settings."""
    import build_gov2   # ייבוא עצל — build_gov2 מייבא מודול זה
    co = build_gov2.carry_over(xlsx); o = dict(co["params"])
    for t, w in co["dw"].items(): o[f"dw_{float(t):g}"] = w
    o["infl_user"] = {float(t): v for t, v in co["infl_user"].items()}
    o["exclude"] = set(co["exclude"])
    return o


def _num(v, default=np.nan):
    try:
        x = float(v); return default if np.isnan(x) else x
    except (TypeError, ValueError): return default


def interp(x, xs, ys):
    xs = np.asarray(xs, float); ys = np.asarray(ys, float); o = np.argsort(xs)
    return float(np.interp(x, xs[o], ys[o]))


def par_duration(t: float, y_pct: float) -> float:
    """מח"מ מתואם של אג"ח בפארי בטווח t ותשואה y (קופון שנתי)."""
    y = y_pct / 100
    return t if abs(y) < 1e-9 else (1 - (1 + y) ** -t) / y


# ---------- 5.1 שקלי מול צמוד ----------
def linked_weight(gap_bp: float | None, bei_pct: float | None, s: dict) -> tuple[float, str]:
    neutral = float(s["linked_neutral"]); lo, hi = float(s["linked_min"]), float(s["linked_max"])
    if gap_bp is None or (isinstance(gap_bp, float) and np.isnan(gap_bp)):
        return neutral, "אין תחזית אינפלציה — נייטרלי"
    margin = float(s["decision_margin_bp"])
    if abs(gap_bp) <= margin:
        tilt, why = 0.0, f"פער {gap_bp:+.0f} נ״ב בתוך מרווח ההחלטה (±{margin:.0f}) → נייטרלי"
    else:
        tilt = max(-1.0, min(1.0, gap_bp / float(s["tilt_scale_bp"]))) * float(s["tilt_max"])
        why = f"פער תחזית−BEI {gap_bp:+.0f} נ״ב → הטיה {tilt:+.0%}"
    if int(_num(s.get("pct_adj_on"), 0)) and bei_pct is not None and not (isinstance(bei_pct, float) and np.isnan(bei_pct)):
        adj = float(s["pct_adj"])
        if bei_pct > float(s["pct_hi"]): tilt -= adj; why += f"; BEI באחוזון {bei_pct:.0f} (גבוה) → −{adj:.0%}"
        elif bei_pct < float(s["pct_lo"]): tilt += adj; why += f"; BEI באחוזון {bei_pct:.0f} (נמוך) → +{adj:.0%}"
    w = max(lo, min(hi, neutral + tilt))
    return round(w, 4), why + f" → משקל צמוד {w:.0%}"


def expected_12m(ytm_pct: float, roll_bp: float, dur: float) -> float:
    """תשואה צפויה ל-12 חודשים (%) = carry + אפקט roll-down (מח"מ × ירידת התשואה על העקום). האופק קבוע: שנה."""
    return ytm_pct + dur * roll_bp / 100


def linked_nominal_equiv(real_pct: float, infl_12m_pct: float, roll_bp: float, dur: float) -> float:
    return ((1 + real_pct / 100) * (1 + infl_12m_pct / 100) - 1) * 100 + dur * roll_bp / 100


# ---------- 5.2 מח"מ ----------
def duration_table(curve: pd.DataFrame, vol: pd.DataFrame, kind: str, s: dict) -> pd.DataFrame:
    ycol, rcol, vcol = ("nominal", "roll_nominal_bp", "nominal_vol_bp") if kind == "nominal" else ("real", "roll_real_bp", "real_vol_bp")
    c = curve.set_index("tenor"); v = vol.set_index("tenor")
    y1 = float(c.loc[1, ycol]); rows = []
    for t in TENORS:
        y = float(c.loc[t, ycol]); roll = float(c.loc[t, rcol]); d = par_duration(t, y)
        e = expected_12m(y, roll, d); excess = (e - y1) * 100
        volbp = float(v.loc[t, vcol]) if t in v.index else interp(t, v.index, v[vcol])
        risk = d * volbp
        score = (excess - float(s["dur_friction_bp"])) / risk if risk > 0 else np.nan
        allowed = float(s["dur_min"]) <= d <= float(s["dur_max"]) and excess >= float(s["dur_min_excess_bp"])
        ex_r, d_r = round(excess, 1), round(d, 3)
        rows.append({"tenor": t, "ytm": round(y, 4), "roll_bp": round(roll, 1), "dur": d_r, "exp_12m": round(e, 4), "excess_bp": ex_r,
                     "breakeven_bp": round(ex_r / d_r, 2) if d_r else np.nan, "risk_bp": round(risk, 1), "score": round(score, 4) if not np.isnan(score) else np.nan, "allowed": bool(allowed)})
    return pd.DataFrame(rows)


def duration_pick(tbl: pd.DataFrame, s: dict, prev_tenor: float | None, policy_gap_bp: float | None) -> tuple[float | None, str]:
    cand = tbl[tbl["allowed"] & tbl["score"].notna()].sort_values("tenor")
    if cand.empty: return None, "אין טווח שעומד במגבלות המח״מ ובעודף המינימלי"
    best = cand.loc[cand["score"].idxmax()]; tenor = float(best["tenor"]); why = f"ציון מקסימלי {best['score']:.2f} בטווח {tenor:g}"
    if prev_tenor is not None and (cand["tenor"] == prev_tenor).any():
        ps = float(cand.loc[cand["tenor"] == prev_tenor, "score"].iloc[0])
        if float(best["score"]) < ps * (1 + float(s["dur_stability"])) and tenor != prev_tenor:
            tenor = float(prev_tenor); why = f"נשאר {tenor:g} (ציון {ps:.2f}); המועמד {best['tenor']:g} לא עובר סף יציבות {float(s['dur_stability']):.0%}"
    if int(_num(s.get("policy_adj_on"), 0)) and policy_gap_bp is not None and not np.isnan(policy_gap_bp) and policy_gap_bp <= -float(s["policy_gap_bp"]):
        lower = cand[cand["tenor"] < tenor]
        if not lower.empty:
            tenor = float(lower["tenor"].max()); why += f"; השוק מגלם הפחתות {-policy_gap_bp:.0f} נ״ב מעבר לתחזית ב״י → צעד למטה ל-{tenor:g}"
    return tenor, why


# ---------- 5.3 סינתטי מול בולט ----------
def ensure_key(df: pd.DataFrame) -> pd.DataFrame:
    """key = מס' ני"ע כטקסט, או השם כשאין מס' ני"ע (כמו gov_history.bond_key)."""
    df = df.copy()
    sid = pd.to_numeric(df.get("security_id"), errors="coerce") if "security_id" in df else pd.Series(np.nan, index=df.index)
    name = df["name"].astype(str).str.strip() if "name" in df else pd.Series("", index=df.index)
    k = df["key"].astype(str) if "key" in df else pd.Series("", index=df.index)
    bad = k.isin(["", "nan", "None", "<NA>"])
    df["key"] = np.where(bad, np.where(sid.notna(), sid.fillna(0).astype(int).astype(str), name), k)
    return df


def _segment_all(bonds: pd.DataFrame, kind: str, exclude: set | None = None) -> pd.DataFrame:
    kinds = REAL_KINDS if kind in REAL_KINDS or kind == "real" else NOMINAL_KINDS
    b = bonds[bonds["kind"].isin(kinds)].copy()
    b = b[(b.get("in_curve", pd.Series("", index=b.index)).astype(str) != "לא")]
    for c in ("modified", "ytm_gross", "ttm", "convexity", "volume_k_nis"):
        b[c] = pd.to_numeric(b[c], errors="coerce")
    b = ensure_key(b)
    if exclude: b = b[~b["key"].isin({str(k) for k in exclude})]
    return b.dropna(subset=["modified", "ytm_gross", "ttm"]).reset_index(drop=True)


def _segment(bonds: pd.DataFrame, kind: str, exclude: set | None = None) -> pd.DataFrame:
    """האפיק לסריקת מבנים — כמו רשימת 06 באקסל: עד NLIST ניירות, הקצרים במח"מ."""
    return _segment_all(bonds, kind, exclude).sort_values("modified").head(NLIST).reset_index(drop=True)


def segment_overflow(bonds: pd.DataFrame, kind: str, exclude: set | None = None) -> int:
    return max(0, len(_segment_all(bonds, kind, exclude)) - NLIST)


def pairs(bonds: pd.DataFrame, target: float) -> pd.DataFrame:
    b = bonds.dropna(subset=["modified"]); rows = []
    for _, si in b.iterrows():
        for _, lj in b.iterrows():
            if si["modified"] < target < lj["modified"]:
                w = (lj["modified"] - target) / (lj["modified"] - si["modified"])
                rows.append({"short_id": si["key"], "long_id": lj["key"], "w_short": w, "w_long": 1 - w})
    return pd.DataFrame(rows, columns=["short_id", "long_id", "w_short", "w_long"])


def _roll_bp(curve: pd.DataFrame, kind: str, ttm: float, h: float = 1.0) -> float:
    """roll-down לנייר: אינטרפולציה של עמודת ה-roll-down של העקום (לפי טווח) בשנים לפדיון של הנייר — כמו עמודת Roll-down ב-02 באקסל."""
    col = "roll_real_bp" if kind in REAL_KINDS or kind == "real" else "roll_nominal_bp"
    if col in curve.columns: return interp(ttm, curve["tenor"], curve[col])
    ycol = "real" if col == "roll_real_bp" else "nominal"
    return (interp(ttm, curve["tenor"], curve[ycol]) - interp(max(ttm - h, TENORS[0]), curve["tenor"], curve[ycol])) * 100


def structure_metrics(bonds: pd.DataFrame, items: pd.DataFrame, curve: pd.DataFrame, s: dict, target: float) -> pd.DataFrame:
    """items: עמודות ids (רשימה), weights (רשימה). מחזיר מדדים לכל מבנה; P&L ב-% לתרחישים."""
    b = bonds.assign(key=bonds["key"].astype(str)).set_index("key"); h = 1.0
    shock = float(s["bb_shock_bp"]) / 100; twist = float(s["bb_twist_bp"]) / 100; span = float(s["bb_twist_span"])
    out = []
    for _, it in items.iterrows():
        ids, ws = list(it["ids"]), list(it["weights"]); rec = {"ids": ids, "weights": ws}
        ytm = conv = exp = up = dn = st = fl = 0.0; liq = np.inf
        for sid, w in zip(ids, ws):
            r = b.loc[sid]; kind = r["kind"]; D = float(r["modified"]); C = _num(r.get("convexity"), 0.0); y = float(r["ytm_gross"]); ttm = float(r["ttm"])
            roll = _roll_bp(curve, kind, ttm, h)
            ytm += w * y; conv += w * C; exp += w * expected_12m(y, roll, D); liq = min(liq, _num(r.get("volume_k_nis"), 0.0))
            pnl = lambda dy: (-D * dy + 0.5 * C * dy * dy) * 100          # dy בשבר עשרוני → %
            up += w * pnl(shock / 100); dn += w * pnl(-shock / 100)
            tw = twist * (ttm - target) / span
            st += w * pnl(tw / 100); fl += w * pnl(-tw / 100)
        rec.update(ytm=round(ytm, 4), convexity=round(conv, 3), roll_bp=round((exp - ytm) * 100 / max(1e-9, sum(ws[i] * float(b.loc[ids[i], "modified"]) for i in range(len(ids)))), 1),
                   liquidity=liq, exp_12m=round(exp, 4), pnl_up=round(up, 4), pnl_dn=round(dn, 4), pnl_steep=round(st, 4), pnl_flat=round(fl, 4))
        rec["score"] = round(np.mean([exp, exp + up, exp + dn, exp + st, exp + fl]), 4)
        rec["ok"] = bool(liq >= float(s["bb_min_volume_k"]))
        out.append(rec)
    return pd.DataFrame(out, columns=["ids", "weights", "ytm", "convexity", "roll_bp", "liquidity", "exp_12m", "pnl_up", "pnl_dn", "pnl_steep", "pnl_flat", "score", "ok"])


def best_structure(bonds: pd.DataFrame, kind: str, target: float, curve: pd.DataFrame, s: dict) -> dict:
    seg = _segment(bonds, kind, s.get("exclude"))
    res = {"kind": "אין מבנה", "ids": [], "weights": [], "score": np.nan, "margin_bp": np.nan, "why": "אין ניירות באפיק", "pair": None, "bullet": None}
    if seg.empty or target is None or np.isnan(target): return res
    pr = pairs(seg, target)
    pm = structure_metrics(seg, pr.assign(ids=[[a, b] for a, b in zip(pr["short_id"], pr["long_id"])], weights=[[w, 1 - w] for w in pr["w_short"]])[["ids", "weights"]], curve, s, target) if len(pr) else pd.DataFrame()
    bl = seg[(seg["modified"] - target).abs() <= float(s["bb_tol"])]
    bm = structure_metrics(seg, pd.DataFrame({"ids": [[i] for i in bl["key"]], "weights": [[1.0]] * len(bl)}), curve, s, target) if len(bl) else pd.DataFrame()
    pm_ok = pm[pm["ok"]] if len(pm) else pm; bm_ok = bm[bm["ok"]] if len(bm) else bm
    bp = pm_ok.loc[pm_ok["score"].idxmax()] if len(pm_ok) else None; bb = bm_ok.loc[bm_ok["score"].idxmax()] if len(bm_ok) else None
    res["pair"] = None if bp is None else bp.to_dict(); res["bullet"] = None if bb is None else bb.to_dict()
    fr = float(s["bb_friction_bp"])
    if bp is None and bb is None: res.update(kind="אין זוג מתאים", why=f"אין נייר ארוך וקצר סביב היעד {target:.2f}, ואין בולט נזיל בטווח הסטייה"); return res
    if bp is None:
        res.update(kind="אין זוג מתאים", ids=list(bb["ids"]), weights=[1.0], score=float(bb["score"]), margin_bp=np.nan,
                   why=f"אין נייר ארוך/קצר מהיעד {target:.2f}; הבולט {list(bb['ids'])} משמש כמבנה"); return res
    if bb is None:
        res.update(kind="סינתטי", ids=list(bp["ids"]), weights=[round(w, 4) for w in bp["weights"]], score=float(bp["score"]), margin_bp=np.nan, why="אין בולט בטווח הסטייה; הזוג הטוב ביותר נבחר"); return res
    margin = (float(bp["score"]) - float(bb["score"])) * 100
    if margin > fr: res.update(kind="סינתטי", ids=list(bp["ids"]), weights=[round(w, 4) for w in bp["weights"]], score=float(bp["score"]))
    elif margin < -fr: res.update(kind="בולט", ids=list(bb["ids"]), weights=[1.0], score=float(bb["score"]))
    else: res.update(kind="דומה", ids=list(bb["ids"]), weights=[1.0], score=float(bb["score"]))
    res["margin_bp"] = round(margin, 1)
    res["why"] = f"ממוצע 5 תרחישים: סינתטי {float(bp['score']):.2f}% מול בולט {float(bb['score']):.2f}% (פער {margin:+.0f} נ״ב, חיכוך {fr:.0f})"
    return res


# ---------- 5.4 סדרות ----------
def series_rank(bonds: pd.DataFrame, bond_history: pd.DataFrame, kind: str, curve: pd.DataFrame, s: dict) -> pd.DataFrame:
    seg = _segment_all(bonds, kind, s.get("exclude")); h = 1.0
    col = "real" if kind in REAL_KINDS or kind == "real" else "nominal"
    rows = []
    hist = bond_history[bond_history["kind"].isin(REAL_KINDS if col == "real" else NOMINAL_KINDS)] if len(bond_history) else bond_history
    for _, r in seg.iterrows():
        cy = interp(r["ttm"], curve["tenor"], curve[col]); spread = (r["ytm_gross"] - cy) * 100
        hk = hist[hist["key"].astype(str) == str(r["key"])]["spread_bp"].dropna() if len(hist) else pd.Series(dtype=float)
        n = int(len(hk)); z = (spread - hk.mean()) / hk.std(ddof=1) if (n >= int(s["ser_min_obs"]) and hk.std(ddof=1) > 0) else np.nan
        rows.append({"security_id": r["security_id"], "key": r["key"], "name": r["name"], "ttm": round(float(r["ttm"]), 3), "dur": float(r["modified"]),
                     "spread_bp": round(spread, 1), "n_obs": n, "z": round(z, 3) if not np.isnan(z) else np.nan,
                     "roll_bp": _roll_bp(curve, kind, float(r["ttm"]), h), "vol": _num(r.get("volume_k_nis"), 0.0)})
    df = pd.DataFrame(rows)
    if df.empty: return pd.DataFrame(columns=["security_id", "key", "name", "ttm", "dur", "spread_bp", "n_obs", "z", "roll_pct", "liq_pct", "score", "rank"])
    df["roll_pct"] = df["roll_bp"].rank(pct=True).round(3); df["liq_pct"] = df["vol"].rank(pct=True).round(3)
    wz, wr, wl = float(s["ser_w_z"]), float(s["ser_w_roll"]), float(s["ser_w_liq"])
    def score(r):
        terms = [(wr, 2 * r["roll_pct"] - 1), (wl, 2 * r["liq_pct"] - 1)]
        if not np.isnan(r["z"]): terms.append((wz, max(-3.0, min(3.0, r["z"])) / 3))
        wsum = sum(w for w, _ in terms); return round(sum(w * v for w, v in terms) / wsum, 4)
    df["score"] = df.apply(score, axis=1); df["rank"] = df["score"].rank(ascending=False, method="first").astype(int)
    return df[["security_id", "key", "name", "ttm", "dur", "spread_bp", "n_obs", "z", "roll_pct", "liq_pct", "score", "rank"]].sort_values("rank").reset_index(drop=True)


def switch_pairs(ranked: pd.DataFrame, s: dict) -> pd.DataFrame:
    rows = []; tol = float(s["switch_dur_tol"]); pick = float(s["switch_pickup_bp"])
    for _, buy in ranked[ranked["rank"] <= 8].iterrows():
        for _, sell in ranked.iterrows():
            if str(sell["key"]) == str(buy["key"]): continue
            dd = abs(float(sell["dur"]) - float(buy["dur"])); pk = float(buy["spread_bp"]) - float(sell["spread_bp"])
            if dd <= tol and pk >= pick:
                rows.append({"sell_id": sell["key"], "buy_id": buy["key"], "pickup_bp": round(pk, 1), "d_diff": round(dd, 3)})
    return pd.DataFrame(rows, columns=["sell_id", "buy_id", "pickup_bp", "d_diff"]).sort_values("pickup_bp", ascending=False).reset_index(drop=True)


# ---------- 5.5 חו"ל ----------
def foreign_verdict(foreign: pd.DataFrame, curve: pd.DataFrame, s: dict, asof: str | None = None) -> pd.DataFrame:
    """stale = הדגל מהמשיכה או תאריך ישן מ-STALE_DAYS לפני יום הנתונים. פסק דין לטווח: פחות מ-fx_min_markets שווקים טריים → 'נתון ישן';
    אחרת 'חו״ל מגודר' אם לפחות fx_min_markets שווקים טריים עוברים את סף הפיקאפ, אחרת 'מקומי'."""
    il1 = interp(1, curve["tenor"], curve["nominal"]); rows = []
    ref = pd.to_datetime(asof) if asof else (pd.to_datetime(curve["date"].iloc[0]) if "date" in curve else None)
    for _, r in foreign.iterrows():
        t = float(r["tenor"]); hedged = float(r["yield"]) - (float(r["rate_1y"]) - il1)
        il_t = interp(t, curve["tenor"], curve["nominal"])
        stale = int(_num(r.get("stale"), 0))
        try:
            if ref is not None and pd.notna(r.get("date")) and (ref - pd.to_datetime(r["date"])).days > STALE_DAYS: stale = 1
        except Exception: pass
        rows.append({"tenor": t, "market": r["market"], "yield": float(r["yield"]), "hedged": round(hedged, 4), "pickup_bp": round((hedged - il_t) * 100, 1), "stale": stale})
    df = pd.DataFrame(rows)
    if df.empty: return df.assign(verdict=[])
    thr, mk = float(s["fx_pickup_bp"]), int(s["fx_min_markets"])
    fresh = df[(df["stale"] == 0) & df["pickup_bp"].notna()].groupby("tenor")["market"].nunique()
    passing = df[(df["stale"] == 0) & (df["pickup_bp"] >= thr)].groupby("tenor")["market"].nunique()
    df["verdict"] = df["tenor"].map(lambda t: "נתון ישן" if fresh.get(t, 0) < mk else ("חו״ל מגודר" if passing.get(t, 0) >= mk else "מקומי"))
    return df


# ---------- הרצה על data/gov ----------
def _read(g: Path, name: str) -> pd.DataFrame:
    f = g / name
    return pd.read_csv(f, encoding="utf-8-sig") if f.exists() else pd.DataFrame()


def inflation_forecast(source: str, boi_exp: pd.DataFrame, cpi_path: pd.DataFrame, user_1y: float | None) -> tuple[float | None, str]:
    if source == "תחזית מנהל השקעות" and user_1y is not None and not np.isnan(user_1y): return float(user_1y), "תחזית מנהל השקעות"
    if source == "שקול 12 ח׳ (תוואי מדדים)" and len(cpi_path):
        e = cpi_path[cpi_path["kind"] == "צפוי"].head(12)
        if len(e) == 12: return float((np.prod(1 + pd.to_numeric(e["monthly_pct"]) / 100) - 1) * 100), "תוואי המדדים 12 ח׳"
    ser = SOURCE_SERIES.get(source, SOURCE_SERIES["ממוצע החזאים (ב״י)"])
    if len(boi_exp) and (boi_exp["series"] == ser).any(): return float(boi_exp.loc[boi_exp["series"] == ser, "value"].iloc[0]), source if source in SOURCE_SERIES else "ממוצע החזאים (ב״י) (ברירת מחדל)"
    return None, "אין מקור תחזית זמין"


def decide(data: Path | str, s: dict | None = None) -> dict:
    s = s or load_settings(); d = Path(data); g = d / "gov"
    bonds, curve, vol = _read(g, "bonds_latest.csv"), _read(g, "curves_latest.csv"), _read(g, "vol.csv")
    hist, cpi, foreign, boi_exp, policy, bh = _read(g, "curves_history.csv"), _read(g, "cpi_path.csv"), _read(g, "foreign_latest.csv"), _read(g, "boi_expectations.csv"), _read(g, "boi_policy.csv"), _read(g, "bond_history.csv")
    prev_hist = _read(g, "decisions_history.csv")
    date = str(curve["date"].iloc[0]) if len(curve) and "date" in curve else dt.date.today().isoformat()
    prev = prev_hist[prev_hist["date"] < date].sort_values("date").tail(1) if len(prev_hist) else pd.DataFrame()
    if vol.empty: vol = pd.DataFrame([{"tenor": t, "nominal_vol_bp": s["vol_floor_nominal_bp"], "real_vol_bp": s["vol_floor_real_bp"], "bei_vol_bp": 30, "n_daily": 0, "method": "floor"} for t in TENORS])
    out = {}; why = {}
    # 5.1
    iu = {float(k): float(v) for k, v in (s.get("infl_user") or {}).items()}
    fc, fsrc = inflation_forecast(str(s.get("infl_source", "")), boi_exp, cpi, iu.get(1.0, np.nan))
    c = curve.set_index("tenor")
    weights = {t: float(s.get(f"dw_{t:g}", 0) or 0) for t in TENORS}
    gap = None
    if fc is not None and sum(weights.values()) > 0:
        manual = str(s.get("infl_source", "")) == "תחזית מנהל השקעות" and iu
        gap = sum(weights[t] * ((iu.get(float(t), fc) if manual else fc) - float(c.loc[t, "bei"])) for t in TENORS) / sum(weights.values()) * 100
    bei_pct = None
    if len(hist) and "bei_2" in hist:
        today_bei2 = float(c.loc[2, "bei"]); hb = pd.to_numeric(hist["bei_2"], errors="coerce").dropna()
        if len(hb) >= 5: bei_pct = float((hb < today_bei2).mean() * 100)
    w, w_why = linked_weight(gap, bei_pct, s); out["linked_weight"] = w; why["linked_weight"] = f"מקור: {fsrc} {fc if fc is None else round(fc, 2)}%; " + w_why
    # 5.2
    fwd1y1y = float(c.loc[2, "fwd_nominal"]) if "fwd_nominal" in c else np.nan
    boi_fc = _num(policy.loc[policy["item"] == "forecast_end_next_year", "value"].iloc[0], np.nan) if len(policy) and (policy["item"] == "forecast_end_next_year").any() else np.nan
    policy_gap = (fwd1y1y - boi_fc) * 100 if not (np.isnan(fwd1y1y) or np.isnan(boi_fc)) else None
    for kind, key in (("nominal", "dur_nominal"), ("real", "dur_real")):
        tbl = duration_table(curve, vol, kind, s)
        pt = _num(prev[key].iloc[0], np.nan) if len(prev) and key in prev else np.nan
        prev_tenor = None if np.isnan(pt) else float(tbl.loc[(tbl["dur"] - pt).abs().idxmin(), "tenor"])
        tenor, dwhy = duration_pick(tbl, s, prev_tenor, policy_gap)
        dur = float(tbl.loc[tbl["tenor"] == tenor, "dur"].iloc[0]) if tenor is not None else np.nan
        out[key] = round(dur, 2) if not np.isnan(dur) else None; out[key + "_tenor"] = tenor; why[key] = dwhy + (f"; אות מדיניות {policy_gap:+.0f} נ״ב" if policy_gap is not None else "; אין תחזית ריבית ב״י")
    out["dur_weighted"] = round(w * (out["dur_real"] or 0) + (1 - w) * (out["dur_nominal"] or 0), 2) if out["dur_nominal"] is not None and out["dur_real"] is not None else None
    why["dur_weighted"] = f"{w:.0%} × צמוד {out['dur_real']} + {1 - w:.0%} × שקלי {out['dur_nominal']}"
    # 5.3
    for kind, key in (("nominal", "struct_nominal"), ("real", "struct_real")):
        tgt = out["dur_nominal" if kind == "nominal" else "dur_real"]
        res = best_structure(bonds, kind, float(tgt) if tgt is not None else np.nan, curve, s)
        out[key] = res["kind"] + ("" if not res["ids"] else " " + "+".join(f"{i}×{wt:.0%}" for i, wt in zip(res["ids"], res["weights"])))
        why[key] = res["why"]
    # 5.4
    for kind, key in (("nominal", "top5_nominal"), ("real", "top5_real")):
        r = series_rank(bonds, bh if len(bh) else pd.DataFrame(columns=["key", "kind", "spread_bp"]), kind, curve, s)
        out[key] = "; ".join(str(x.key) for x in r.head(5).itertuples()); why[key] = "ציון = Z/3×0.4 + roll×0.3 + נזילות×0.3" if len(r) else "אין סדרות"
    # 5.5
    if len(foreign):
        fv = foreign_verdict(foreign, curve, s, asof=date)
        for t, key in ((5, "foreign_5y"), (10, "foreign_10y")):
            vt = fv[fv["tenor"] == t]
            out[key] = vt["verdict"].iloc[0] if len(vt) else "אין נתון"
            why[key] = "; ".join(f"{x.market} {x.pickup_bp:+.0f} נ״ב{' (ישן)' if x.stale else ''}" for x in vt.itertuples())
    else:
        out["foreign_5y"] = out["foreign_10y"] = "אין נתון"; why["foreign_5y"] = why["foreign_10y"] = "אין קובץ חו״ל"
    # כתיבה
    asof = dt.date.today().isoformat()
    pd.DataFrame([{"key": k, "value": out.get(k), "why": why.get(k, ""), "as_of": asof} for k in DECISION_KEYS]).to_csv(g / "decisions_latest.csv", index=False, encoding="utf-8-sig")
    row = {"date": date, "as_of": asof, **{k: out.get(k) for k in DECISION_KEYS}, "dur_nominal_tenor": out.get("dur_nominal_tenor"), "dur_real_tenor": out.get("dur_real_tenor")}
    hist_df = prev_hist[prev_hist["date"] != date] if len(prev_hist) else pd.DataFrame()
    hist_df = pd.concat([hist_df, pd.DataFrame([row])], ignore_index=True).sort_values("date")
    hist_df.to_csv(g / "decisions_history.csv", index=False, encoding="utf-8-sig")
    return out


if __name__ == "__main__":
    a = sys.argv[1:]; data = a[0] if a and not a[0].startswith("--") else "data"
    s = load_settings()
    if "--settings" in a:   # ההגדרות של יובב מהחוברת (corp-data/00_Gov/ממשלתי.xlsx)
        xl = a[a.index("--settings") + 1]
        if Path(xl).exists(): s = load_settings(overrides=overrides_from_workbook(xl)); print("הגדרות מהחוברת:", xl)
    res = decide(data, s)
    for k in DECISION_KEYS: print(f"{k}: {res.get(k)}")
