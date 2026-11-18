import pandas as pd

from kapmodell.aggregate import daily_aggregate, rows_to_frame

TSOS = ["50HERTZ", "AMPRION"]


def _raw(hour, tso, name, eic, fmax, fref):
    return {"DateTimeUtc": hour, "Tso": tso, "CneName": name, "CneEic": eic, "Fmax": fmax, "Fref": fref}


def test_max_per_hour_then_hour_statistics():
    rows = [
        # Stunde 1: zwei Zeilen (Richtungen/Ausfallfälle) -> ungünstigster Fall 0.9
        _raw("2026-11-15T23:00:00Z", "AMPRION", "Leitung A ", "EIC-A", 1000, 900),
        _raw("2026-11-15T23:00:00Z", "AMPRION", "Leitung A ", "EIC-A", 1000, -500),
        # Stunde 2: Überlast 1.2
        _raw("2026-11-16T00:00:00Z", "AMPRION", "Leitung A ", "EIC-A", 1000, -1200),
        # Stunde 3: 0.3
        _raw("2026-11-16T01:00:00Z", "AMPRION", "Leitung A ", "EIC-A", 1000, 300),
        _raw("2026-11-16T01:00:00Z", "TENNETBV", "Fremd", "EIC-X", 1000, 999),   # NL-ÜNB wird ignoriert
    ]
    agg = daily_aggregate(rows_to_frame(rows, TSOS), "2026-11-16")
    assert len(agg) == 1
    r = agg.iloc[0]
    assert r["Stunden"] == 3
    assert abs(r["Auslastung_Summe"] - (0.9 + 1.2 + 0.3)) < 1e-9
    assert abs(r["Auslastung_Max"] - 1.2) < 1e-9
    assert r["Stunden_hoch"] == 2          # 0.9 und 1.2
    assert r["Stunden_Ueberlast"] == 1     # nur 1.2
    assert r["Netzelement"] == "Leitung A"  # Leerzeichen am Ende entfernt


def test_missing_eic_falls_back_to_name_key_and_zero_fmax_ignored():
    rows = [
        _raw("2026-11-15T23:00:00Z", "50HERTZ", "Ohne Code ", None, 500, 250),
        _raw("2026-11-15T23:00:00Z", "50HERTZ", "Null-Fmax", "EIC-0", 0, 100),
    ]
    agg = daily_aggregate(rows_to_frame(rows, TSOS), "2026-11-16")
    assert list(agg["Schluessel_JAO"]) == ["NAME|Ohne Code"]
    assert agg.iloc[0]["EIC"] == ""


def test_two_tso_views_of_same_eic_are_merged_worst_case():
    rows = [
        _raw("2026-11-15T23:00:00Z", "AMPRION", "Name Amprion", "EIC-T", 2120, 372),     # 0.1755
        _raw("2026-11-15T23:00:00Z", "50HERTZ", "Name 50Hertz", "EIC-T", 2182, 479),     # 0.2195
    ]
    agg = daily_aggregate(rows_to_frame(rows, TSOS), "2026-11-16")
    assert len(agg) == 1 and agg.iloc[0]["Stunden"] == 1
    assert abs(agg.iloc[0]["Auslastung_Max"] - 479 / 2182) < 1e-9
    assert agg.iloc[0]["Fmax_MW"] == 2120                       # ungünstigerer (kleinerer) Grenzwert
    assert "Name Amprion" in agg.iloc[0]["Namen"] and "Name 50Hertz" in agg.iloc[0]["Namen"]


def test_empty_input_gives_empty_frame():
    assert daily_aggregate(rows_to_frame([], TSOS), "2026-11-16").empty


def _raw2(hour, tso, name, eic, fmax, fref, imax=None, presolved=False):
    d = _raw(hour, tso, name, eic, fmax, fref)
    d["Imax"], d["Presolved"] = imax, presolved
    return d


def test_helper_rows_from_handbook_are_excluded():
    """Handbuch v2.0: Imax=9999-Zeilen und technische Randbedingungen sind keine physischen Netzelemente."""
    rows = [
        _raw2("2026-11-15T23:00:00Z", "AMPRION", "Echte Leitung", "EIC-R", 1000, 800, imax=1500),
        _raw2("2026-11-15T23:00:00Z", "AMPRION", "Echte Leitung", "EIC-R", 6500, 800, imax=9999),   # Duplikat Core/Nicht-Core
        _raw2("2026-11-15T23:00:00Z", "AMPRION", "External Constraint DE_AL_export", "", 1000, 900, imax=1000),
        _raw2("2026-11-15T23:00:00Z", "AMPRION", "NL_import", "", 1000, 900),
        _raw2("2026-11-15T23:00:00Z", "AMPRION", "Nur Imax 9999", "EIC-9", 6500, 100, imax=9999),
    ]
    agg = daily_aggregate(rows_to_frame(rows, TSOS), "2026-11-16")
    assert list(agg["Schluessel_JAO"]) == ["EIC-R"]            # nur das echte Element bleibt
    assert agg.iloc[0]["Fmax_MW"] == 1000                      # nicht der 9999-A-Wert


def test_binding_hours_count_presolved_once_per_hour():
    rows = [
        _raw2("2026-11-15T23:00:00Z", "AMPRION", "Linie", "EIC-B", 1000, 500, presolved=True),
        _raw2("2026-11-15T23:00:00Z", "AMPRION", "Linie", "EIC-B", 1000, 300, presolved=True),   # gleiche Stunde: zählt einmal
        _raw2("2026-11-16T00:00:00Z", "AMPRION", "Linie", "EIC-B", 1000, 500, presolved=False),
        _raw2("2026-11-16T01:00:00Z", "AMPRION", "Linie", "EIC-B", 1000, 500, presolved="true"),  # Text 'true' aus JSON/CSV
    ]
    r = daily_aggregate(rows_to_frame(rows, TSOS), "2026-11-16").iloc[0]
    assert r["Stunden"] == 3 and r["Stunden_bindend"] == 2
