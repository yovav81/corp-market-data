"""gov_signals: היסטוריה חודשית ארוכה, ארה"ב חודשי, כלכלנים, אותות לאופטימיזציה."""
import pandas as pd, numpy as np, pytest
import gov_signals as gs

ZCM = "SERIES_CODE,TIME_PERIOD,OBS_VALUE\n" + "".join(
    f"ZC_TSB_ZND_{t}Y_MA,2024-{m:02d},{3.0 + 0.1 * t + 0.01 * m}\nZC_TSB_ZRD_{t}Y_MA,2024-{m:02d},{1.0 + 0.05 * t + 0.01 * m}\n" for t in (2, 10) for m in range(1, 13)) + "OTHER_CODE,2024-01,9\n"
TRE = 'Date,"1 Mo","5 Yr","10 Yr"\n01/02/2024,5.5,3.90,4.00\n01/31/2024,5.5,4.10,4.20\n02/15/2024,5.4,4.30,4.40\n'


def test_parse_zcm_monthly():
    df = gs.parse_zcm_monthly(ZCM)
    assert set(df.columns) == {"period", "tenor", "nominal", "real", "bei"} and len(df) == 24
    r = df[(df.period == "2024-03") & (df.tenor == 10)].iloc[0]
    assert r.nominal == pytest.approx(4.03) and r.real == pytest.approx(1.53) and r.bei == pytest.approx(((1.0403 / 1.0153) - 1) * 100, abs=1e-3)


def test_parse_treasury_monthly():
    df = gs.parse_treasury_monthly(TRE)
    assert list(df.period) == ["2024-01", "2024-02"]
    assert df.us_10y.iloc[0] == pytest.approx(4.10) and df.us_5y.iloc[1] == pytest.approx(4.30)


def test_pct_rank():
    assert gs.pct_rank([1, 2, 3, 4], 3.5) == 75.0 and gs.pct_rank([1, 2, 3, 4], 0) == 0.0
    assert gs.pct_rank([], 1) is None and gs.pct_rank([1, 2], None) is None


def test_build_signals():
    cm = gs.parse_zcm_monthly(ZCM); us = gs.parse_treasury_monthly(TRE)
    curve = pd.DataFrame({"tenor": [1, 2, 10], "nominal": [3.1, 3.3, 4.2], "real": [1.2, 1.4, 1.70], "bei": [1.9, 1.9, 1.9]})
    econ = pd.DataFrame({"house": ["א", "ב"], "rate_12m": [3.0, 3.5], "infl_12m": [2.2, np.nan]})
    policy = pd.DataFrame({"item": ["rate_current", "forecast_end_next_year", "forecast_infl_4q"], "value": ["3.25", "3.0", "2.0"]})
    sig = gs.build_signals(cm, us, econ, policy, curve, us10_now=4.5).set_index("key")["val"]
    assert sig["real10_pct"] == pytest.approx(100.0)            # 1.70 מעל כל ההיסטוריה
    assert sig["hist_months"] == 12
    assert sig["rate_now"] == 3.25 and sig["rate_fc_n"] == 3 and sig["rate_fc_mean"] == pytest.approx((3.0 + 3.5 + 3.0) / 3, abs=1e-3)
    assert sig["rate_fc_min"] == 3.0 and sig["rate_fc_max"] == 3.5 and sig["infl_fc_mean"] == pytest.approx(2.1)
    assert sig["ilus10_now"] == pytest.approx(-30.0)
    assert pd.isna(sig["ilus10_pct"])                            # רק חודשיים חופפים — פחות מ-12 → אין אחוזון
    us12 = pd.DataFrame({"period": [f"2024-{m:02d}" for m in range(1, 13)], "us_5y": 4.0, "us_10y": [4.0 + 0.1 * m for m in range(1, 13)]})
    sig2 = gs.build_signals(cm, us12, econ, policy, curve, us10_now=4.5).set_index("key")["val"]
    assert 0 <= sig2["ilus10_pct"] <= 100 and not pd.isna(sig2["ilus10_med"])


def test_build_signals_without_history():
    curve = pd.DataFrame({"tenor": [1, 2, 10], "nominal": [3.1, 3.3, 4.2], "real": [1.2, 1.4, 1.6], "bei": [1.9, 1.9, 1.9]})
    sig = gs.build_signals(pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), curve).set_index("key")["val"]
    assert pd.isna(sig["real10_pct"]) and pd.isna(sig["rate_fc_mean"]) and sig["real10_now"] == 1.6


def test_merge_us_keeps_old_and_replaces_overlap():
    old = pd.DataFrame({"period": ["2023-12", "2024-01"], "us_5y": [3.8, 9.9], "us_10y": [3.9, 9.9]})
    new = gs.parse_treasury_monthly(TRE)
    m = gs.merge_monthly(old, new)
    assert list(m.period) == ["2023-12", "2024-01", "2024-02"] and m.us_10y.iloc[1] == pytest.approx(4.10)
