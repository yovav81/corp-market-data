from datetime import datetime
from zoneinfo import ZoneInfo

from fetch_market import market_open_now, trade_day

TZ = ZoneInfo("Asia/Jerusalem")


def t(s):
    return datetime.strptime(s, "%Y-%m-%d %H:%M").replace(tzinfo=TZ)


def test_evening_run_is_same_day():
    assert trade_day(t("2026-11-24 18:17")) == "2026-11-24"
    assert trade_day(t("2026-11-24 21:47")) == "2026-11-24"


def test_delayed_run_after_midnight_belongs_to_previous_trading_day():
    assert trade_day(t("2026-11-25 02:42")) == "2026-11-24"
    assert trade_day(t("2026-11-25 06:27")) == "2026-11-24"   # ריצה #96 — 12 שעות איחור
    assert trade_day(t("2026-11-25 08:59")) == "2026-11-24"


def test_after_open_is_today():
    assert trade_day(t("2026-11-25 09:00")) == "2026-11-25"
    assert trade_day(t("2026-11-25 12:00")) == "2026-11-25"


def test_weekend_rolls_back_to_friday():
    assert trade_day(t("2026-11-28 03:00")) == "2026-11-27"   # שבת לפנות בוקר → שישי
    assert trade_day(t("2026-11-30 03:00")) == "2026-11-27"   # שני לפנות בוקר → שישי


def test_market_open_now():
    assert market_open_now(t("2026-11-25 10:00"))
    assert market_open_now(t("2026-11-25 17:29"))
    assert not market_open_now(t("2026-11-25 17:30"))
    assert not market_open_now(t("2026-11-25 08:59"))
    assert not market_open_now(t("2026-11-28 12:00"))   # שבת
