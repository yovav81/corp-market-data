import pandas as pd
import gov_cpi


def _cpi_fixture(d, n_known=4, n_exp=18):
    rows = []; idx = 100.0
    for i in range(n_known):
        ym = f"2026-{i+1:02d}"; pct = [0.3, -0.2, 0.4, 0.5][i]; idx = round(idx * (1 + pct / 100), 2)
        rows.append({"ym": ym, "pub_date": f"2026-{i+2:02d}-15", "kind": "ידוע", "index": idx, "monthly_pct": pct, "cum_factor": "", "source": "הלמ״ס", "as_of": "2026-11-05"})
    cum = 1.0
    for j in range(n_exp):
        m = n_known + j; y, mo = 2026 + m // 12, m % 12 + 1
        pct = round(0.15 + 0.25 * ((j % 4) - 1.5), 3)     # עונתיות מלאכותית סביב 0.15 לחודש
        cum *= 1 + pct / 100
        rows.append({"ym": f"{y}-{mo:02d}", "pub_date": "", "kind": "צפוי", "index": "", "monthly_pct": pct, "cum_factor": round(cum, 6), "source": "BEI+עונתיות", "as_of": "2026-11-05"})
    pd.DataFrame(rows).to_csv(d / "short" / "cpi_expected.csv", index=False, encoding="utf-8-sig")


def test_season_sums_to_zero(tmp_data):
    _cpi_fixture(tmp_data)
    df = gov_cpi.build_cpi_path(tmp_data)
    exp = df[df["kind"] == "צפוי"].head(12)
    assert abs(exp["season_pct"].sum()) < 1e-9
    assert list(df.columns) == ["ym", "pub_date", "kind", "index", "monthly_pct", "season_pct", "cum_factor", "source", "as_of"]


def test_known_rows_kept(tmp_data):
    _cpi_fixture(tmp_data)
    df = gov_cpi.build_cpi_path(tmp_data)
    known = df[df["kind"] == "ידוע"]
    assert len(known) == 4 and (known["season_pct"] == 0).all() and known["index"].notna().all()


def test_horizon(tmp_data):
    _cpi_fixture(tmp_data, n_exp=30)
    df = gov_cpi.build_cpi_path(tmp_data, horizon_months=24)
    assert (df["kind"] == "צפוי").sum() == 24
