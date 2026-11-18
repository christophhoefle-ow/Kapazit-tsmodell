import pandas as pd

from kapmodell.mapping import build_mapping, monthly_from_daily, sgm_elements


def _daily(tag, key, eic, namen, **kw):
    base = dict(Tag=tag, Schluessel_JAO=key, UENB="AMPRION", Netzelement=namen.split(" | ")[0], Namen=namen,
                EIC=eic, Fmax_MW=1000.0, Stunden=24, Auslastung_Summe=6.0, Auslastung_Max=0.5,
                Stunden_hoch=0, Stunden_Ueberlast=0)
    base.update(kw)
    return base


def _sgm(lines=(), trafos=()):
    l = pd.DataFrame(lines, columns=["NE_name", "EIC_Code"])
    t = pd.DataFrame(trafos, columns=["Name", "EIC_Code"])
    return sgm_elements(l, t)


def test_eic_then_name_then_unmatched():
    daily = pd.DataFrame([
        _daily("2026-11-01", "EIC-1", "EIC-1", "Linie Eins"),
        _daily("2026-11-01", "NAME|Linie Zwei", "", "Linie Zwei"),
        _daily("2026-11-01", "EIC-9", "EIC-9", "Unbekannt"),
    ])
    sgm = _sgm(lines=[("Linie Eins", "EIC-1"), ("LINIE ZWEI ", "EIC-ANDERS")])
    m = build_mapping(daily, sgm).set_index("Schluessel_JAO")
    assert m.loc["EIC-1", "Zuordnung"] == "EIC"
    assert m.loc["NAME|Linie Zwei", "Zuordnung"] == "Name" and m.loc["NAME|Linie Zwei", "Schluessel"] == "EIC-ANDERS"
    assert m.loc["EIC-9", "Zuordnung"] == "nicht zugeordnet" and m.loc["EIC-9", "Schluessel"] == "EIC-9"


def test_alias_has_priority_over_name_match():
    daily = pd.DataFrame([_daily("2026-11-01", "EIC-J", "EIC-J", "Gleicher Name")])
    sgm = _sgm(lines=[("Gleicher Name", "EIC-S1"), ("Anderer Name", "EIC-S2")])
    alias = pd.DataFrame([("EIC-J", "EIC-S2", "geprüft")], columns=["Schluessel_JAO", "Schluessel_SGM", "Bemerkung"])
    m = build_mapping(daily, sgm, alias)
    assert m.loc[m["Zuordnung"] == "Alias", "Schluessel"].tolist() == ["EIC-S2"]


def test_multiple_jao_rows_for_one_sgm_element_take_worst_case_and_months_add_up():
    daily = pd.DataFrame([
        _daily("2026-11-01", "EIC-A", "EIC-A", "Tie Line", Auslastung_Max=0.2, Auslastung_Summe=4.0),
        _daily("2026-11-01", "EIC-B", "EIC-B", "Tie Line", Auslastung_Max=0.7, Auslastung_Summe=9.0, Stunden_hoch=0),
        _daily("2026-11-02", "EIC-B", "EIC-B", "Tie Line", Auslastung_Max=0.9, Auslastung_Summe=10.0, Stunden_hoch=3,
               Stunden_Ueberlast=1),
        _daily("2026-12-01", "EIC-B", "EIC-B", "Tie Line", Auslastung_Max=0.4, Auslastung_Summe=5.0),
    ])
    sgm = _sgm(lines=[("Tie Line", "EIC-SGM")])
    mon = monthly_from_daily(daily, sgm).set_index(["Monat", "Schluessel"])
    nov = mon.loc[("2026-11-01", "EIC-SGM")]
    assert nov["Stunden"] == 48                     # 24 + 24 (je Tag der Maximalwert beider JAO-Zeilen)
    assert abs(nov["Auslastung_Summe"] - (9.0 + 10.0)) < 1e-9
    assert nov["Auslastung_Max"] == 0.9 and nov["Stunden_hoch"] == 3 and nov["Stunden_Ueberlast"] == 1
    assert mon.loc[("2026-12-01", "EIC-SGM"), "Stunden"] == 24


def test_empty_daily():
    assert monthly_from_daily(pd.DataFrame(columns=["Tag"]), _sgm()).empty
