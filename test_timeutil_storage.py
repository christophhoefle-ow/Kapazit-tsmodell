import tempfile

import pandas as pd

from kapmodell.aggregate import DAILY_COLUMNS
from kapmodell.storage import days_to_fetch, read_daily, read_log, update_log, write_day
from kapmodell.timeutil import business_day_hours, day_range


def test_hours_winter_summer_and_dst_days():
    w = business_day_hours("2026-11-16")
    assert len(w) == 24 and str(w[0]) == "2026-11-15 23:00:00+00:00"
    s = business_day_hours("2026-07-01")
    assert len(s) == 24 and str(s[0]) == "2026-06-30 22:00:00+00:00"
    assert len(business_day_hours("2026-03-29")) == 23      # Beginn Sommerzeit
    assert len(business_day_hours("2026-10-25")) == 25      # Ende Sommerzeit


def test_day_range():
    assert day_range("2026-11-30", "2026-12-02") == ["2026-11-30", "2026-12-01", "2026-12-02"]


def _row(tag, key, summe):
    return dict(Tag=tag, Schluessel_JAO=key, UENB="X", Netzelement=key, Namen=key, EIC=key, Fmax_MW=1000.0,
                Stunden=24, Auslastung_Summe=summe, Auslastung_Max=0.5, Stunden_hoch=0, Stunden_Ueberlast=0, Stunden_bindend=0)


def test_write_day_replaces_same_day_only():
    with tempfile.TemporaryDirectory() as d:
        write_day(d, "2026-11-01", pd.DataFrame([_row("2026-11-01", "A", 1.0)], columns=DAILY_COLUMNS))
        write_day(d, "2026-11-02", pd.DataFrame([_row("2026-11-02", "A", 2.0)], columns=DAILY_COLUMNS))
        write_day(d, "2026-11-01", pd.DataFrame([_row("2026-11-01", "A", 9.0)], columns=DAILY_COLUMNS))
        got = read_daily(d).set_index("Tag")
        assert len(got) == 2 and got.loc["2026-11-01", "Auslastung_Summe"] == 9.0 and got.loc["2026-11-02", "Auslastung_Summe"] == 2.0


def test_days_to_fetch_skips_complete_days():
    with tempfile.TemporaryDirectory() as d:
        p = f"{d}/log.csv"
        for tag, st in (("2026-11-16", "vollstaendig"), ("2026-11-17", "unvollstaendig")):
            update_log(p, {"Tag": tag, "Stunden_erwartet": 24, "Stunden_geliefert": 24, "Zeilen_DE": 1,
                           "Status": st, "Abruf_UTC": "x"})
        todo = days_to_fetch("2026-11-18", 3, read_log(p))
        assert todo == ["2026-11-15", "2026-11-17", "2026-11-18"]


def test_fallback_days_count_as_complete():
    with tempfile.TemporaryDirectory() as d:
        p = f"{d}/log.csv"
        for tag, st in (("2026-11-16", "fallback"), ("2026-11-17", "wartet")):
            update_log(p, {"Tag": tag, "Stunden_erwartet": "", "Stunden_geliefert": 0, "Zeilen_DE": 0, "Status": st, "Abruf_UTC": "x"})
        assert days_to_fetch("2026-11-18", 3, read_log(p)) == ["2026-11-15", "2026-11-17", "2026-11-18"]


def test_old_daily_files_without_new_column_remain_readable():
    with tempfile.TemporaryDirectory() as d:
        old = pd.DataFrame([_row("2026-11-01", "A", 1.0)], columns=DAILY_COLUMNS).drop(columns=["Stunden_bindend"])
        old.to_csv(f"{d}/2026-11.csv", index=False)
        got = read_daily(d)
        assert list(got.columns) == DAILY_COLUMNS and int(got.iloc[0]["Stunden_bindend"]) == 0
        write_day(d, "2026-11-02", pd.DataFrame([_row("2026-11-02", "A", 2.0)], columns=DAILY_COLUMNS))
        assert len(read_daily(d)) == 2
