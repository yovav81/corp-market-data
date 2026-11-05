import json
import pandas as pd
from conftest import write_cpi_fixture, ROOT
import gov_data, gov_foreign


def _patch_network(monkeypatch):
    monkeypatch.setattr(gov_data, "fetch_fred", lambda: pd.DataFrame([{"series": s, "name": n, "date": "2026-11-04", "value": 4.0 + i * 0.1, "chg_1m_bp": 0, "chg_3m_bp": 0, "source": "FRED"}
                                                                      for i, (s, n) in enumerate(gov_data.FRED)]))
    monkeypatch.setattr(gov_data, "fetch_boi", lambda: (pd.DataFrame([{"tenor": t, "period": "2026-09", "nominal": 3.2, "real": 1.6, "bei": 1.57} for t in (1, 2, 5, 10)]),
                                                        pd.DataFrame([{"series": "EXP_HAZ_MAD_AVG_01Y_MA", "name": "ממוצע החזאים — 12 חודשים", "period": "2026-10", "value": 2.08}])))
    monkeypatch.setattr(gov_foreign, "fetch_us", lambda: ("2026-11-04", {1: 4.38, 2: 4.59, 3: 4.6, 5: 4.7, 7: 4.8, 10: 4.91, 20: 5.2, 30: 5.28}))
    monkeypatch.setattr(gov_foreign, "fetch_ecb", lambda tenors: ("2026-11-04", {t: 2.0 + 0.05 * t for t in gov_foreign.TENORS}))
    monkeypatch.setattr(gov_foreign, "fetch_boe", lambda: ("2026-11-04", {t: 4.0 + 0.05 * t for t in gov_foreign.TENORS}))


def test_settings_default_json():
    s = json.loads((ROOT / "gov" / "settings_default.json").read_text(encoding="utf-8"))
    keys = {x["key"]: x for x in s}
    for k in ["linked_neutral", "tilt_scale_bp", "tilt_max", "decision_margin_bp", "linked_min", "linked_max", "infl_source", "dur_min", "dur_max",
              "dur_min_excess_bp", "dur_friction_bp", "dur_stability", "policy_adj_on", "policy_gap_bp", "vol_floor_nominal_bp", "vol_floor_real_bp",
              "bb_tol", "bb_friction_bp", "bb_shock_bp", "bb_twist_bp", "bb_twist_span", "bb_min_volume_k", "ser_w_z", "ser_w_roll", "ser_w_liq",
              "ser_min_obs", "switch_pickup_bp", "switch_dur_tol", "fx_pickup_bp", "fx_min_markets", "makam_min", "float_weight", "bench_dur",
              "bench_linked", "scen_parallel_1", "scen_parallel_2", "scen_bei_bp", "scen_real_bp", "scen_infl_pct", "horizon_years"]:
        assert k in keys, k
        assert all(f in keys[k] for f in ("value", "label", "note", "group", "fmt")), k
    assert keys["linked_neutral"]["value"] == 0.5 and keys["tilt_scale_bp"]["value"] == 50 and keys["tilt_max"]["value"] == 0.30
    assert keys["dur_max"]["value"] == 8 and keys["bb_twist_span"]["value"] == 8 and keys["fx_pickup_bp"]["value"] == 20


def test_gov_data_outputs(monkeypatch, tmp_data):
    _patch_network(monkeypatch); write_cpi_fixture(tmp_data)
    gov_data.main(str(tmp_data))
    g = tmp_data / "gov"
    bonds = pd.read_csv(g / "bonds_latest.csv", encoding="utf-8-sig"); assert "key" in bonds.columns and bonds["key"].notna().all()
    assert set(pd.read_csv(g / "bond_history.csv", encoding="utf-8-sig").columns) >= {"date", "key", "spread_bp"}
    vol = pd.read_csv(g / "vol.csv", encoding="utf-8-sig"); assert list(vol.columns)[:6] == ["tenor", "nominal_vol_bp", "real_vol_bp", "bei_vol_bp", "n_daily", "method"]
    assert "season_pct" in pd.read_csv(g / "cpi_path.csv", encoding="utf-8-sig").columns
    fo = pd.read_csv(g / "foreign_latest.csv", encoding="utf-8-sig"); assert set(fo["market"]) == {"US", "EUR", "GBP"}
    pol = pd.read_csv(g / "boi_policy.csv", encoding="utf-8-sig")
    assert list(pol.columns) == ["item", "value", "as_of", "source"] and (pol["item"] == "il_1y").any()
    assert pd.to_numeric(pol.loc[pol["item"] == "il_1y", "value"]).iloc[0] > 0
    assert (pol["item"] == "rate_current").any()
    assert (g / "decisions_latest.csv").exists() and (g / "decisions_history.csv").exists()
