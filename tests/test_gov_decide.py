import json
import numpy as np, pandas as pd, pytest
from conftest import ROOT, write_cpi_fixture
import gov_decide

S = gov_decide.load_settings()
TEN = gov_decide.TENORS


def _curve(nom_fn, real_fn):
    rows = []
    for t in TEN:
        n, r = nom_fn(t), real_fn(t)
        rows.append({"tenor": t, "nominal": n, "real": r, "bei": ((1 + n / 100) / (1 + r / 100) - 1) * 100})
    df = pd.DataFrame(rows)
    df["roll_nominal_bp"] = [(nom_fn(t) - nom_fn(max(t - 1, TEN[0]))) * 100 for t in TEN]
    df["roll_real_bp"] = [(real_fn(t) - real_fn(max(t - 1, TEN[0]))) * 100 for t in TEN]
    df["fwd_nominal"] = df["nominal"]; df["fwd_real"] = df["real"]
    return df


def _vol(level=60.0):
    return pd.DataFrame([{"tenor": t, "nominal_vol_bp": level, "real_vol_bp": level * 0.8, "bei_vol_bp": level * 0.5, "n_daily": 70, "method": "daily"} for t in TEN])


def _bonds(kind="שקלי"):
    rows = []
    for i, (sid, ttm, y, vol) in enumerate([(101, 0.9, 3.1, 900), (102, 2.0, 3.4, 2000), (103, 4.0, 3.8, 3000), (104, 7.0, 4.1, 2500), (105, 12.0, 4.4, 1500)]):
        d = (1 - (1 + y / 100) ** -ttm) / (y / 100)
        rows.append({"security_id": sid, "key": str(sid), "name": f"נייר {sid}", "kind": kind, "ttm": ttm, "ytm_gross": y, "modified": round(d, 4),
                     "convexity": round(d * d, 3), "volume_k_nis": vol, "in_curve": ""})
    return pd.DataFrame(rows)


# ---------- 5.1 ----------
def test_linked_weight_neutral_inside_margin():
    w, why = gov_decide.linked_weight(20, 50, S); assert w == 0.5 and "נייטרלי" in why


def test_linked_weight_tilt():
    assert gov_decide.linked_weight(50, 50, S)[0] == pytest.approx(0.80)
    assert gov_decide.linked_weight(-100, 50, S)[0] == pytest.approx(0.20)   # 0.5−0.30 חסום ב-0.20


def test_pct_adjustment():
    assert gov_decide.linked_weight(50, 90, S)[0] == pytest.approx(0.75)
    assert gov_decide.linked_weight(50, 90, {**S, "pct_adj_on": 0})[0] == pytest.approx(0.80)


def test_linked_nominal_equiv():
    assert gov_decide.linked_nominal_equiv(1.5, 2.0, 0.0, 4.0) == pytest.approx(((1.015 * 1.02) - 1) * 100)
    assert gov_decide.expected_12m(4.0, 10.0, 5.0) == pytest.approx(4.0 + 5.0 * 0.10)


# ---------- 5.2 ----------
def test_duration_table_columns_and_breakeven():
    cur = _curve(lambda t: 3.0 + 0.15 * t, lambda t: 1.0 + 0.08 * t)
    tbl = gov_decide.duration_table(cur, _vol(), "nominal", S)
    assert list(tbl.columns) == ["tenor", "ytm", "roll_bp", "dur", "exp_12m", "excess_bp", "breakeven_bp", "risk_bp", "score", "allowed"]
    r5 = tbl[tbl["tenor"] == 5].iloc[0]
    assert r5["breakeven_bp"] == pytest.approx(r5["excess_bp"] / r5["dur"], abs=0.01)
    assert bool(r5["allowed"]) and not bool(tbl[tbl["tenor"] == 30].iloc[0]["allowed"])   # מח״מ מעל המקסימום


def test_duration_pick_max_score():
    tbl = pd.DataFrame({"tenor": [1, 2, 5, 10], "score": [0.1, 0.5, 0.9, 0.7], "allowed": [True, True, True, True], "dur": [1, 1.9, 4.5, 8.0], "excess_bp": [0, 20, 60, 80]})
    t, why = gov_decide.duration_pick(tbl, S, None, None); assert t == 5 and "ציון" in why


def test_duration_stability():
    tbl = pd.DataFrame({"tenor": [2, 5, 7], "score": [0.3, 0.80, 0.84], "allowed": [True] * 3, "dur": [1.9, 4.5, 6.0], "excess_bp": [20, 60, 70]})
    assert gov_decide.duration_pick(tbl, S, 5, None)[0] == 5          # שיפור 5% < סף 10%
    assert gov_decide.duration_pick(tbl, S, 5, None, )[0] == 5
    tbl.loc[tbl["tenor"] == 7, "score"] = 0.95
    assert gov_decide.duration_pick(tbl, S, 5, None)[0] == 7           # שיפור 19% > סף


def test_duration_policy_step_down():
    tbl = pd.DataFrame({"tenor": [2, 3, 5, 7], "score": [0.3, 0.5, 0.9, 0.7], "allowed": [True] * 4, "dur": [1.9, 2.8, 4.5, 6.0], "excess_bp": [20, 30, 60, 70]})
    assert gov_decide.duration_pick(tbl, S, None, -40)[0] == 3         # השוק מגלם 40 נ״ב הפחתות מעבר לב״י → צעד למטה
    assert gov_decide.duration_pick(tbl, S, None, -10)[0] == 5
    assert gov_decide.duration_pick(tbl, {**S, "policy_adj_on": 0}, None, -40)[0] == 5


# ---------- 5.3 ----------
def test_pairs_weights_sum_and_duration():
    b = _bonds(); p = gov_decide.pairs(b, 4.0)
    assert len(p) > 0
    assert np.allclose(p["w_short"] + p["w_long"], 1)
    d = dict(zip(b["key"], b["modified"]))
    assert np.allclose(p["w_short"] * p["short_id"].map(d) + p["w_long"] * p["long_id"].map(d), 4.0)


def test_scan_no_long_bond():
    cur = _curve(lambda t: 3.0 + 0.15 * t, lambda t: 1.0 + 0.08 * t)
    res = gov_decide.best_structure(_bonds(), "שקלי", 20.0, cur, S)
    assert res["kind"] == "אין זוג מתאים" and res["ids"] == []


def test_twist_zero_at_target():
    cur = _curve(lambda t: 3.0 + 0.15 * t, lambda t: 1.0 + 0.08 * t)
    b = _bonds(); target = float(b.loc[b["security_id"] == 103, "modified"].iloc[0])
    b.loc[b["security_id"] == 103, "ttm"] = target             # נייר שה-ttm שלו = היעד בדיוק
    bullets = pd.DataFrame([{"ids": ["103"], "weights": [1.0]}])
    m = gov_decide.structure_metrics(b, bullets, cur, S, target)
    assert m.iloc[0]["pnl_steep"] == pytest.approx(0.0) and m.iloc[0]["pnl_flat"] == pytest.approx(0.0)
    assert m.iloc[0]["pnl_up"] < 0 < m.iloc[0]["pnl_dn"]


def test_best_structure_verdict_fields():
    cur = _curve(lambda t: 3.0 + 0.15 * t, lambda t: 1.0 + 0.08 * t)
    res = gov_decide.best_structure(_bonds(), "שקלי", 4.0, cur, S)
    assert res["kind"] in ("סינתטי", "בולט", "דומה") and len(res["ids"]) in (1, 2) and abs(sum(res["weights"]) - 1) < 1e-9
    assert "margin_bp" in res and "why" in res


# ---------- 5.4 ----------
def _hist(bonds, n, spread_mean=0.0, sd=5.0):
    rows = []
    for _, b in bonds.iterrows():
        for i in range(n):
            rows.append({"date": f"2026-10-{13 + i:02d}", "key": b["key"], "kind": b["kind"], "spread_bp": spread_mean + sd * (1 if i % 2 else -1)})
    return pd.DataFrame(rows)


def test_series_rank_no_history():
    cur = _curve(lambda t: 3.0 + 0.15 * t, lambda t: 1.0 + 0.08 * t)
    r = gov_decide.series_rank(_bonds(), _hist(_bonds(), 3), "שקלי", cur, S)
    assert list(r.columns) == ["security_id", "key", "name", "ttm", "dur", "spread_bp", "n_obs", "z", "roll_pct", "liq_pct", "score", "rank"]
    assert r["z"].isna().all() and r["score"].notna().all() and set(r["rank"]) == {1, 2, 3, 4, 5}


def test_series_rank_with_history_prefers_cheap():
    cur = _curve(lambda t: 3.0 + 0.15 * t, lambda t: 1.0 + 0.08 * t)
    b = _bonds(); b.loc[b["security_id"] == 103, "ytm_gross"] += 0.30      # נייר 103 זול ב-30 נ״ב מול העקום
    r = gov_decide.series_rank(b, _hist(b, 12), "שקלי", cur, S)
    assert r["z"].notna().all() and r.loc[r["security_id"] == 103, "rank"].iloc[0] == 1


def test_switch_pairs_pickup():
    cur = _curve(lambda t: 3.0 + 0.15 * t, lambda t: 1.0 + 0.08 * t)
    b = _bonds(); b.loc[b["security_id"] == 103, "ytm_gross"] += 0.30
    b = pd.concat([b, pd.DataFrame([{"security_id": 106, "key": "106", "name": "נייר 106", "kind": "שקלי", "ttm": 4.2, "ytm_gross": 3.80,
                                     "modified": float(b.loc[b.security_id == 103, "modified"].iloc[0]) + 0.1, "convexity": 15.0, "volume_k_nis": 1000, "in_curve": ""}])], ignore_index=True)
    r = gov_decide.series_rank(b, _hist(b, 12), "שקלי", cur, S)
    sw = gov_decide.switch_pairs(r, S)
    assert list(sw.columns) == ["sell_id", "buy_id", "pickup_bp", "d_diff"]
    assert ((sw["buy_id"] == "103") & (sw["sell_id"] == "106")).any() and (sw["pickup_bp"] >= S["switch_pickup_bp"]).all()


# ---------- 5.5 ----------
def test_foreign_hedged_formula():
    cur = _curve(lambda t: 3.1 + 0.1 * (t - 1), lambda t: 1.0 + 0.08 * t)   # il_1y = 3.1
    fo = pd.DataFrame([{"market": m, "tenor": t, "yield": 4.6 if m == "US" else 2.5, "rate_1y": 4.4 if m == "US" else 2.2, "date": "2026-11-04", "source": "x", "as_of": "2026-11-05", "stale": 0}
                       for m in ("US", "EUR", "GBP") for t in (5, 10)])
    v = gov_decide.foreign_verdict(fo, cur, S)
    us5 = v[(v["market"] == "US") & (v["tenor"] == 5)].iloc[0]
    assert us5["hedged"] == pytest.approx(3.3) and us5["pickup_bp"] == pytest.approx((3.3 - 3.5) * 100)
    assert set(v.columns) == {"tenor", "market", "yield", "hedged", "pickup_bp", "stale", "verdict"}
    assert (v["verdict"] == "מקומי").all()


def test_foreign_verdict_requires_two_markets():
    cur = _curve(lambda t: 3.1 + 0.1 * (t - 1), lambda t: 1.0 + 0.08 * t)
    fo = pd.DataFrame([{"market": m, "tenor": 5, "yield": y, "rate_1y": r1, "date": "2026-11-04", "source": "x", "as_of": "2026-11-05", "stale": st}
                       for m, y, r1, st in (("US", 5.2, 4.4, 0), ("EUR", 2.6, 2.2, 0), ("GBP", 4.6, 4.0, 1))])
    v = gov_decide.foreign_verdict(fo, cur, S)       # US pickup +40, EUR 0, GBP ישן → שוק אחד בלבד
    assert (v["verdict"] == "מקומי").all()
    fo.loc[fo["market"] == "EUR", "yield"] = 3.2       # EUR hedged 4.1 → pickup +60
    v = gov_decide.foreign_verdict(fo, cur, S)
    assert (v["verdict"] == "חו״ל מגודר").all()


# ---------- decide ----------
def test_decide_writes_history(monkeypatch, tmp_data):
    import gov_data, gov_foreign
    from test_gov_data_outputs import _patch_network
    _patch_network(monkeypatch); write_cpi_fixture(tmp_data)
    gov_data.main(str(tmp_data))
    res = gov_decide.decide(tmp_data)
    assert {"linked_weight", "dur_nominal", "dur_real", "dur_weighted", "struct_nominal", "struct_real", "top5_nominal", "top5_real", "foreign_5y", "foreign_10y"} <= set(res)
    res2 = gov_decide.decide(tmp_data)
    hist = pd.read_csv(tmp_data / "gov" / "decisions_history.csv", encoding="utf-8-sig")
    assert len(hist) == 1 and "linked_weight" in hist.columns and "date" in hist.columns
    latest = pd.read_csv(tmp_data / "gov" / "decisions_latest.csv", encoding="utf-8-sig")
    assert list(latest.columns) == ["key", "value", "why", "as_of"]
