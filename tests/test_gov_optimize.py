"""מנוע התרחישים והאופטימיזציה (gov_decide): הסתברויות, תשואת רגליים, בחירת התיק."""
import numpy as np, pandas as pd, pytest
import gov_decide as gd
from test_gov_decide import _curve

S = gd.load_settings()
CUR = _curve(lambda t: 3.0 + 0.12 * min(t, 10), lambda t: 1.0 + 0.10 * min(t, 10))
FLAT = _curve(lambda t: 4.0, lambda t: 2.0)


def test_shock_at_nodes_and_interpolation():
    assert gd.shock_at(1, -50, -35, -15) == -50 and gd.shock_at(2, -50, -35, -15) == -50
    assert gd.shock_at(3.5, -50, -35, -15) == pytest.approx(-42.5)
    assert gd.shock_at(7, 0, 20, 45) == pytest.approx(30.0)
    assert gd.shock_at(10, 0, 20, 45) == 45 and gd.shock_at(30, 0, 20, 45) == 45


def test_probs_base_without_signals():
    p, _ = gd.scenario_probs(S, {}, y1=3.1)
    assert len(p) == 9 and sum(p) == pytest.approx(1.0)
    assert p[0] == pytest.approx(0.30) and p[3] == pytest.approx(0.05)


def test_probs_valuation_signal_moves_weight_to_long_rally():
    base, _ = gd.scenario_probs(S, {}, y1=3.1)
    p, notes = gd.scenario_probs(S, {"real10_pct": 100.0}, y1=3.1)
    k = float(S["opt_k_val"])
    assert sum(p) == pytest.approx(1.0)
    assert p[6] == pytest.approx(base[6] + k / 2) and p[1] == pytest.approx(base[1] + k / 2)
    assert p[5] == pytest.approx(base[5] - k / 2) and p[2] == pytest.approx(base[2] - k / 2)
    assert any("תמחור" in n for n in notes)


def test_probs_policy_fiscal_dispersion_and_floor():
    sig = {"rate_now": 3.25, "rate_fc_mean": 3.0, "rate_fc_min": 2.75, "rate_fc_max": 3.5}
    base, _ = gd.scenario_probs(S, {}, y1=3.375)
    p, _ = gd.scenario_probs(S, sig, y1=3.375)          # ריבית מגולמת בסוף השנה 3.5 מול תחזית 3.0 → פער +50 → אות מלא
    kp, kd = float(S["opt_k_policy"]), float(S["opt_k_disp"])
    assert p[4] == pytest.approx(base[4] + kp + kd / 2) and p[0] == pytest.approx(base[0] - kd)
    hi = gd.scenario_probs({**S, "fiscal_risk": 3}, {}, y1=3.1)[0]; assert hi[5] == pytest.approx(base[5] + float(S["opt_k_fiscal"]))
    floor = gd.scenario_probs({**S, "opt_k_val": 0.9}, {"real10_pct": 0.0}, y1=3.1)[0]
    assert min(floor) > 0 and sum(floor) == pytest.approx(1.0) and floor[1] == pytest.approx(floor[6])
    off = gd.scenario_probs({**S, "opt_signals_on": 0}, {"real10_pct": 100.0}, y1=3.1)[0]; assert off == pytest.approx(base)
    user = gd.scenario_probs({**S, "user_rate_fc": 3.5}, sig, y1=3.375)[0]; assert user[4] == pytest.approx(base[4] + kd / 2)   # תחזית ידנית 3.5 = המגולם → אין אות ריבית


def test_leg_returns_flat_curve():
    ln = gd.leg_returns(FLAT, "nominal", S, 2.0); lr = gd.leg_returns(FLAT, "real", S, 2.0)
    assert list(ln.index) == gd.GRID and ln["r1"].tolist() == pytest.approx([4.0] * 7) and lr["r1"].tolist() == pytest.approx([4.0] * 7)
    d1 = gd.par_duration(4, 4.0)
    assert ln.loc[5, "r3"] == pytest.approx(4.0 - d1 * 0.5 + 0.5 * d1 * (d1 + 1) * 50 ** 2 / 1e6)
    assert ln.loc[1, "r4"] == pytest.approx(4.0)                    # שנה אחת — נפדה, אין זעזוע
    assert lr.loc[5, "r8"] == pytest.approx(2.0 + 2.0 + 1.0)         # ריאלי ללא שינוי + אינפלציה 3%
    assert ln.loc[5, "r8"] == pytest.approx(4.0 - d1 * 0.4 + 0.5 * d1 * (d1 + 1) * 40 ** 2 / 1e6)   # שקלי נפגע מעליית הציפיות


def test_optimize_respects_floor_band_and_duration_limits():
    r = gd.optimize(CUR, S, {}, anchor_w=0.5, infl_fc=2.0)
    assert r["status"] == "תקין" and r["worst"] >= float(S["risk_floor_pct"]) - 1e-9
    assert 0.4 - 1e-9 <= r["w"] <= 0.6 + 1e-9 and 1 <= r["leg_n"] <= gd.NLEG and 1 <= r["leg_r"] <= gd.NLEG
    assert float(S["dur_min"]) <= r["dur_n"] <= float(S["dur_max"]) and float(S["dur_min"]) <= r["dur_r"] <= float(S["dur_max"])
    c = r["cands"]; assert len(c) == 5 * gd.NLEG * gd.NLEG
    best = c[c["feasible"]]["exp"].max(); assert r["exp"] >= best - float(S["opt_tol_bp"]) / 100 - 1e-9
    near = c[c["feasible"] & (c["exp"] >= best - float(S["opt_tol_bp"]) / 100)]; assert r["worst"] == pytest.approx(near["worst"].max())


def test_optimize_higher_floor_is_safer_and_fallback():
    lo = gd.optimize(CUR, {**S, "risk_floor_pct": -5.0}, {}, 0.5, 2.0); hi = gd.optimize(CUR, {**S, "risk_floor_pct": 2.0}, {}, 0.5, 2.0)
    assert hi["worst"] >= lo["worst"] - 1e-9 and hi["exp"] <= lo["exp"] + float(S["opt_tol_bp"]) / 100 + 1e-9
    none = gd.optimize(CUR, {**S, "risk_floor_pct": 50.0}, {}, 0.5, 2.0)
    assert none["status"] != "תקין" and none["worst"] == pytest.approx(none["cands"][none["cands"]["dur_ok"]]["worst"].max())


def test_optimize_keeps_yesterday_inside_tolerance():
    r = gd.optimize(CUR, S, {}, 0.5, 2.0); c = r["cands"]; best = c[c["feasible"]]["exp"].max()
    near = c[c["feasible"] & (c["exp"] >= best - float(S["opt_tol_bp"]) / 100)]
    other = near[(near["w"] != r["w"]) | (near["leg_n"] != r["leg_n"]) | (near["leg_r"] != r["leg_r"])]
    assert len(other), "הפיקסצ'ר צריך לפחות שני תיקים בתוך סף האדישות"
    o = other.iloc[-1]; kept = gd.optimize(CUR, S, {}, 0.5, 2.0, prev=(float(o.w), int(o.leg_n), int(o.leg_r)))
    assert (kept["w"], kept["leg_n"], kept["leg_r"]) == (pytest.approx(o.w), o.leg_n, o.leg_r) and kept["kept"]
    far = gd.optimize(CUR, S, {}, 0.5, 2.0, prev=(0.5, 1, 1))          # התיק של אתמול רחוק מהטוב ביותר → מתחלף
    assert not far["kept"] and (far["w"], far["leg_n"], far["leg_r"]) == (r["w"], r["leg_n"], r["leg_r"])


def test_eval_portfolio_matches_candidates_and_interpolates():
    r = gd.optimize(CUR, S, {}, 0.5, 2.0); c = r["cands"]; row = c[(c.leg_n <= 7) & (c.leg_r <= 7)].iloc[30]      # רגליים בודדות — על הרשת
    e = gd.eval_portfolio(CUR, S, {}, float(row.w), float(row.dn), float(row.dr), 2.0)
    assert e["exp"] == pytest.approx(row.exp) and e["worst"] == pytest.approx(row.worst)
    ln = gd.leg_returns(CUR, "nominal", S, 2.0); mid = (ln.loc[3, "dur"] + ln.loc[4, "dur"]) / 2
    e2 = gd.eval_portfolio(CUR, S, {}, 0.0, mid, ln.loc[3, "dur"], 2.0)
    assert e2["r"][0] == pytest.approx((ln.loc[3, "r1"] + ln.loc[4, "r1"]) / 2)


def test_leg_options_singles_then_mixes():
    lo = gd.leg_options(CUR, "real", S, 2.0); L = gd.leg_returns(CUR, "real", S, 2.0)
    assert list(lo.index) == list(range(1, 13)) and lo.loc[3, "label"] == "3" and lo.loc[11, "label"] == "4+7" and pd.isna(lo.loc[11, "tenor"])
    assert lo.loc[11, "dur"] == pytest.approx((L.loc[4, "dur"] + L.loc[7, "dur"]) / 2) and lo.loc[11, "r5"] == pytest.approx((L.loc[4, "r5"] + L.loc[7, "r5"]) / 2)


def test_decide_writes_optimizer_fields_and_migrates_history(tmp_data, monkeypatch):
    from conftest import write_cpi_fixture
    from test_gov_data_outputs import _patch_network
    import gov_data
    write_cpi_fixture(tmp_data); _patch_network(monkeypatch); gov_data.main(str(tmp_data))
    g = tmp_data / "gov"; h = pd.read_csv(g / "decisions_history.csv", encoding="utf-8-sig")
    old = h.iloc[[0]].copy(); old["date"] = "2026-10-01"; old["dur_nominal_tenor"] = 3.0; old["dur_real_tenor"] = 4.0
    old.drop(columns=["opt_leg_n", "opt_leg_r"]).to_csv(g / "decisions_history.csv", index=False, encoding="utf-8-sig")     # היסטוריה בפורמט v2.0
    res = gd.decide(tmp_data); h2 = pd.read_csv(g / "decisions_history.csv", encoding="utf-8-sig")
    assert h2.loc[h2.date == "2026-10-01", "opt_leg_n"].iloc[0] == 3 and h2.loc[h2.date == "2026-10-01", "opt_leg_r"].iloc[0] == 4
    assert res["opt_status"] in ("תקין", "אין תיק שעומד בתקציב הסיכון — נבחר הבטוח ביותר") and res["opt_exp"] is not None
    last = h2.iloc[-1]; assert {"rule_linked", "rule_dur_nominal_tenor", "opt_leg_n", "opt_exp", "opt_worst"} <= set(h2.columns) and last["opt_leg_n"] >= 1
    off = gd.decide(tmp_data, {**gd.load_settings(), "opt_on": 0}); assert off["opt_status"] == "כבוי" and off["linked_weight"] == off["rule_linked"]
