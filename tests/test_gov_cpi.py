import pandas as pd
import gov_cpi
from conftest import write_cpi_fixture as _cpi_fixture


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
