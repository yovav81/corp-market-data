import io, zipfile
import openpyxl, pandas as pd, pytest
import gov_foreign

ECB_CSV = """KEY,FREQ,REF_AREA,CURRENCY,PROVIDER_FM,INSTRUMENT_FM,PROVIDER_FM_ID,DATA_TYPE_FM,TIME_PERIOD,OBS_VALUE,OBS_STATUS
YC.B.U2.EUR.4F.G_N_A.SV_C_YM.SR_5Y,B,U2,EUR,4F,G_N_A,SV_C_YM,SR_5Y,2026-11-03,2.4123,A
YC.B.U2.EUR.4F.G_N_A.SV_C_YM.SR_5Y,B,U2,EUR,4F,G_N_A,SV_C_YM,SR_5Y,2026-11-04,2.4387,A
YC.B.U2.EUR.4F.G_N_A.SV_C_YM.SR_5Y,B,U2,EUR,4F,G_N_A,SV_C_YM,SR_5Y,2026-11-05,2.4501,A
"""


def _boe_zip() -> bytes:
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "4. spot curve"
    ws["A1"] = "Spot curve"; ws["A4"] = "years:"
    tenors = [0.5, 1, 1.5, 2, 3, 5, 7, 10, 15, 20, 25, 30]
    for j, t in enumerate(tenors): ws.cell(4, 2 + j, t)
    import datetime as dt
    for i, (d, base) in enumerate([(dt.datetime(2026, 11, 3), 4.0), (dt.datetime(2026, 11, 4), 4.1)]):
        ws.cell(6 + i, 1, d)
        for j, t in enumerate(tenors): ws.cell(6 + i, 2 + j, round(base + 0.05 * t, 4))
    bio = io.BytesIO(); wb.save(bio)
    z = io.BytesIO()
    with zipfile.ZipFile(z, "w") as zf: zf.writestr("GLC Nominal daily data current month.xlsx", bio.getvalue())
    return z.getvalue()


def test_parse_ecb_csv():
    date, value = gov_foreign.parse_ecb_csv(ECB_CSV)
    assert date == "2026-11-05" and value == 2.4501


def test_parse_boe_spot():
    date, curve = gov_foreign.parse_boe_spot(_boe_zip())
    assert date == "2026-11-04"
    assert curve[1.0] == pytest.approx(4.15) and curve[10.0] == pytest.approx(4.6)


def test_foreign_fallback_stale(monkeypatch, tmp_data):
    monkeypatch.setattr(gov_foreign, "fetch_us", lambda: ("2026-11-04", {1: 4.38, 2: 4.59, 3: 4.6, 5: 4.7, 7: 4.8, 10: 4.91, 20: 5.2, 30: 5.28}))
    monkeypatch.setattr(gov_foreign, "fetch_ecb", lambda tenors: (_ for _ in ()).throw(RuntimeError("ECB down")))
    monkeypatch.setattr(gov_foreign, "fetch_boe", lambda: ("2026-11-04", {0.5: 4.0, 1: 4.1, 2: 4.2, 3: 4.3, 5: 4.4, 7: 4.5, 10: 4.6, 15: 4.7, 20: 4.8, 30: 4.9}))
    prev = pd.DataFrame([{"market": "EUR", "tenor": t, "yield": 2.0 + 0.05 * t, "rate_1y": 2.05, "date": "2026-11-03", "source": "ECB", "as_of": "2026-11-04", "stale": 0}
                         for t in gov_foreign.TENORS])
    df = gov_foreign.build_foreign(tmp_data, prev)
    assert list(df.columns) == ["market", "tenor", "yield", "rate_1y", "date", "source", "as_of", "stale"]
    eur = df[df["market"] == "EUR"]; us = df[df["market"] == "US"]; gb = df[df["market"] == "GBP"]
    assert len(eur) == len(gov_foreign.TENORS) and (eur["stale"] == 1).all() and (eur["date"] == "2026-11-03").all()
    assert (us["stale"] == 0).all() and (gb["stale"] == 0).all()
    assert us[us["tenor"] == 15]["yield"].iloc[0] == pytest.approx((4.91 + 5.2) / 2)      # 15 = אינטרפולציה 10–20
    assert (us["rate_1y"] == 4.38).all() and gb[gb["tenor"] == 10]["rate_1y"].iloc[0] == pytest.approx(4.1)


def test_foreign_market_dropped_without_prev(monkeypatch, tmp_data):
    monkeypatch.setattr(gov_foreign, "fetch_us", lambda: ("2026-11-04", {1: 4.38, 2: 4.59, 3: 4.6, 5: 4.7, 7: 4.8, 10: 4.91, 20: 5.2, 30: 5.28}))
    monkeypatch.setattr(gov_foreign, "fetch_ecb", lambda tenors: (_ for _ in ()).throw(RuntimeError("ECB down")))
    monkeypatch.setattr(gov_foreign, "fetch_boe", lambda: (_ for _ in ()).throw(RuntimeError("BoE down")))
    df = gov_foreign.build_foreign(tmp_data, None)
    assert set(df["market"]) == {"US"}


# --- מקורות חסינים יותר (10.11): אוצר ארה״ב כמקור ראשי ל-UST, ECB בבקשה אחת ---
TREASURY_CSV = """Date,"1 Mo","1.5 Month","2 Mo","3 Mo","4 Mo","6 Mo","1 Yr","2 Yr","3 Yr","5 Yr","7 Yr","10 Yr","20 Yr","30 Yr"
11/07/2026,3.95,3.94,3.93,3.90,3.88,3.80,3.62,3.55,3.58,3.70,3.85,4.08,4.62,4.68
11/09/2026,3.96,3.95,3.94,3.91,3.89,3.81,3.63,3.57,3.60,3.72,3.87,4.10,4.64,4.70
"""
ECB_MULTI = """KEY,FREQ,REF_AREA,CURRENCY,PROVIDER_FM,INSTRUMENT_FM,PROVIDER_FM_ID,DATA_TYPE_FM,TIME_PERIOD,OBS_VALUE,OBS_STATUS
YC.B.U2.EUR.4F.G_N_A.SV_C_YM.SR_1Y,B,U2,EUR,4F,G_N_A,SV_C_YM,SR_1Y,2026-11-06,1.95,A
YC.B.U2.EUR.4F.G_N_A.SV_C_YM.SR_1Y,B,U2,EUR,4F,G_N_A,SV_C_YM,SR_1Y,2026-11-09,1.97,A
YC.B.U2.EUR.4F.G_N_A.SV_C_YM.SR_10Y,B,U2,EUR,4F,G_N_A,SV_C_YM,SR_10Y,2026-11-09,2.71,A
"""


def test_parse_treasury_csv():
    date, curve = gov_foreign.parse_treasury_csv(TREASURY_CSV)
    assert date == "2026-11-09" and curve[1] == 3.63 and curve[10] == 4.10 and curve[30] == 4.70 and 3 in curve and 7 in curve and 20 in curve


def test_parse_ecb_multi_csv():
    date, curve = gov_foreign.parse_ecb_multi_csv(ECB_MULTI)
    assert date == "2026-11-09" and curve == {1: 1.97, 10: 2.71}


def test_get_retries(monkeypatch):
    import gov_data, urllib.error
    calls = {"n": 0}
    class R:
        def __init__(self): self.b = b"ok"
        def read(self): return self.b
        def __enter__(self): return self
        def __exit__(self, *a): return False
    def fake_open(req, timeout):
        calls["n"] += 1
        if calls["n"] < 3: raise urllib.error.HTTPError(req.full_url, 504, "Gateway Time-out", None, None)
        return R()
    monkeypatch.setattr(gov_data.urllib.request, "urlopen", fake_open)
    monkeypatch.setattr(gov_data.time, "sleep", lambda s: None)
    assert gov_data.get("https://x/y", timeout=5, retries=3) == "ok" and calls["n"] == 3
