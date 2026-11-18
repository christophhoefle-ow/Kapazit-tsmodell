"""Zuordnung JAO-Netzelemente <-> Stammdaten des Static Grid Model (SGM) und Monatsverdichtung.

Die Zuordnung passiert hier im Skript und NICHT in Excel. Excel braucht dadurch nur noch einen
eindeutigen Schlüssel je Leitung bzw. Trafo (Regel siehe normalize.element_key).

Zuordnungsregeln (in dieser Reihenfolge):
  0. Manuelle Zuordnung (data/manual/alias_jao_sgm.csv, von Menschen geprüft)  -> Zuordnung "Alias"
  1. EIC-Code des SGM-Elements = EIC-Code im JAO-Datensatz            -> Zuordnung "EIC"
  2. sonst: Name des SGM-Elements = einer der JAO-Namen (normalisiert,
     ohne Groß-/Kleinschreibung)                                        -> Zuordnung "Name"
  3. JAO-Elemente ohne Treffer bleiben mit eigenem Schlüssel erhalten  -> "nicht zugeordnet"

Treffen mehrere JAO-Zeilen auf dasselbe SGM-Element (z. B. beide Netzbetreiber melden dieselbe
Kuppelleitung unter verschiedenen Namen), gilt der UNGÜNSTIGSTE Wert je Kennzahl.
"""
from __future__ import annotations

import pandas as pd

from .normalize import element_key, name_key, norm_eic, norm_text

MONTHLY_COLUMNS = [
    "Monat", "Schluessel", "Zuordnung", "UENB", "Netzelement", "EIC", "Fmax_MW",
    "Stunden", "Auslastung_Summe", "Auslastung_Max", "Stunden_hoch", "Stunden_Ueberlast",
]


def sgm_elements(lines: pd.DataFrame, trafos: pd.DataFrame) -> pd.DataFrame:
    """Alle Leitungen und Trafos des SGM als Tabelle (sgm_key, eic, name, art)."""
    parts = []
    for df, art, eic_col, name_col in (
        (lines, "Leitung", "EIC_Code", "NE_name"),
        (trafos, "Trafo", "EIC_Code", "Name"),
    ):
        if df is None or df.empty:
            continue
        t = pd.DataFrame({
            "eic": df[eic_col].map(norm_eic),
            "name": df[name_col].map(norm_text),
        })
        t["sgm_key"] = [element_key(e, n) for e, n in zip(t["eic"], t["name"])]
        t["art"] = art
        parts.append(t)
    if not parts:
        return pd.DataFrame(columns=["sgm_key", "eic", "name", "art"])
    return pd.concat(parts, ignore_index=True)


def build_mapping(daily: pd.DataFrame, sgm: pd.DataFrame, alias: pd.DataFrame | None = None) -> pd.DataFrame:
    """Zuordnungstabelle (Schluessel, Schluessel_JAO, Zuordnung).

    `alias` (optional) enthält die Spalten Schluessel_JAO und Schluessel_SGM.
    """
    jao = daily.drop_duplicates("Schluessel_JAO")[["Schluessel_JAO", "EIC", "Namen"]]
    jao_eics = set(jao.loc[jao["EIC"] != "", "Schluessel_JAO"])
    name_index: dict[str, set[str]] = {}
    for jk, namen in zip(jao["Schluessel_JAO"], jao["Namen"]):
        for n in str(namen).split(" | "):
            if n:
                name_index.setdefault(name_key(n), set()).add(jk)

    rows, claimed, done = [], set(), set()
    sgm_keys = set(sgm["sgm_key"])
    if alias is not None and not alias.empty:
        for jk, sk in zip(alias["Schluessel_JAO"], alias["Schluessel_SGM"]):
            if jk in set(jao["Schluessel_JAO"]) and sk in sgm_keys:
                rows.append((sk, jk, "Alias"))
                claimed.add(jk)
                done.add(sk)
    for _, s in sgm.iterrows():
        if s["sgm_key"] in done:
            continue
        if s["eic"] and s["eic"] in jao_eics:
            rows.append((s["sgm_key"], s["eic"], "EIC"))
            claimed.add(s["eic"])
        else:
            for jk in sorted(name_index.get(name_key(s["name"]), set())):
                rows.append((s["sgm_key"], jk, "Name"))
                claimed.add(jk)
    for jk in jao["Schluessel_JAO"]:
        if jk not in claimed:
            rows.append((jk, jk, "nicht zugeordnet"))
    return pd.DataFrame(rows, columns=["Schluessel", "Schluessel_JAO", "Zuordnung"]).drop_duplicates()


def monthly_from_daily(daily: pd.DataFrame, sgm: pd.DataFrame, alias: pd.DataFrame | None = None) -> pd.DataFrame:
    """Tagesdateien -> Monatstabelle je Excel-Schlüssel (Lieferung an Blatt D_Auslastung)."""
    if daily.empty:
        return pd.DataFrame(columns=MONTHLY_COLUMNS)
    mapping = build_mapping(daily, sgm, alias)
    m = daily.merge(mapping, on="Schluessel_JAO", how="inner")

    # a) mehrere JAO-Zeilen -> ein SGM-Element: ungünstigster Wert je Tag
    day = m.sort_values(["Tag", "Schluessel", "Schluessel_JAO"]).groupby(["Tag", "Schluessel"], as_index=False).agg(
        Zuordnung=("Zuordnung", "first"), UENB=("UENB", "first"), Netzelement=("Netzelement", "first"),
        EIC=("EIC", "first"), Fmax_MW=("Fmax_MW", "min"), Stunden=("Stunden", "max"),
        Auslastung_Summe=("Auslastung_Summe", "max"), Auslastung_Max=("Auslastung_Max", "max"),
        Stunden_hoch=("Stunden_hoch", "max"), Stunden_Ueberlast=("Stunden_Ueberlast", "max"))

    # b) Tage -> Monate (Summen addieren, Maximum bilden)
    day["Monat"] = day["Tag"].str.slice(0, 7) + "-01"
    mon = day.groupby(["Monat", "Schluessel"], as_index=False).agg(
        Zuordnung=("Zuordnung", "first"), UENB=("UENB", "first"), Netzelement=("Netzelement", "first"),
        EIC=("EIC", "first"), Fmax_MW=("Fmax_MW", "min"), Stunden=("Stunden", "sum"),
        Auslastung_Summe=("Auslastung_Summe", "sum"), Auslastung_Max=("Auslastung_Max", "max"),
        Stunden_hoch=("Stunden_hoch", "sum"), Stunden_Ueberlast=("Stunden_Ueberlast", "sum"))
    mon["Auslastung_Summe"] = mon["Auslastung_Summe"].round(6)
    mon["Auslastung_Max"] = mon["Auslastung_Max"].round(6)
    return mon[MONTHLY_COLUMNS].sort_values(["Monat", "Schluessel"]).reset_index(drop=True)
