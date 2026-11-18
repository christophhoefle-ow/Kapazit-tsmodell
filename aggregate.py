"""Verdichtung der stündlichen JAO-Zeilen zu Auslastungskennzahlen je Netzelement.

Regel (identisch zur Aufbereitung im Testdaten-Blatt '2_je_Netzelement' des Projektteams):

  * Auslastung einer Zeile = |Fref| / Fmax                      (nur wenn Fmax > 0)
  * Auslastung je Stunde   = Maximum über alle Zeilen desselben Netzelements
                             (beide Richtungen, alle Ausfallfälle = "ungünstigster Fall")
  * je Tag und Element     = Stunden, Summe der Stundenwerte, Maximum,
                             Stunden >= Schwelle "hoch", Stunden >= Schwelle "Überlast"

Warum die SUMME der Stundenwerte gespeichert wird: Mittelwerte lassen sich nicht über
mehrere Tage oder Monate zusammenfassen, Summen schon (Mittel = Summe / Stunden).
"""
from __future__ import annotations

import re
from typing import Iterable

import numpy as np
import pandas as pd

from .normalize import element_key, norm_eic, norm_text

DAILY_COLUMNS = [
    "Tag", "Schluessel_JAO", "UENB", "Netzelement", "Namen", "EIC",
    "Fmax_MW", "Stunden", "Auslastung_Summe", "Auslastung_Max",
    "Stunden_hoch", "Stunden_Ueberlast", "Stunden_bindend",
]

# Laut JAO-Handbuch (v2.0, Kap. 5.7/5.12) enthalten die Domain-Seiten auch Zeilen, die KEIN physisches Netzelement
# beschreiben: Elemente mit Imax = 9999 (Grenzen Core/Nicht-Core, nur für die KPI der Nicht-Core-Austausche) sowie
# technische Randbedingungen (z. B. 'External Constraint BE_AL_export', 'NL_import', Gleichheitsbedingungen).
IMAX_UNBEGRENZT = 9999
_TECH_NAME = re.compile(r"^\s*(external constraint|equality constraint)|^[a-z]{2}_(import|export)\s*$", re.I)


def _to_float(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return np.nan


def rows_to_frame(raw_rows: Iterable[dict], tsos: Iterable[str], hour_utc=None) -> pd.DataFrame:
    """Rohzeilen der JAO-API in ein schlankes DataFrame umwandeln.

    Schlüsselnamen werden klein geschrieben verglichen, damit sowohl die CSV-Schreibweise
    ('CneName') als auch JSON-Varianten ('cneName') funktionieren.
    Es bleiben nur die Zeilen der gewünschten Übertragungsnetzbetreiber übrig.
    """
    wanted = {norm_text(t).upper() for t in tsos}
    recs = []
    for raw in raw_rows:
        d = {str(k).lower(): v for k, v in raw.items()}
        tso = norm_text(d.get("tso")).upper()
        if tso not in wanted:
            continue
        try:
            fmax = float(d.get("fmax"))
            fref = float(d.get("fref"))
        except (TypeError, ValueError):
            continue
        ts = hour_utc if hour_utc is not None else d.get("datetimeutc")
        presolved = str(d.get("presolved")).strip().lower() in ("true", "1", "yes")
        recs.append((ts, tso, d.get("cnename"), d.get("cneeic"), fmax, fref, _to_float(d.get("imax")), presolved))
    frame = pd.DataFrame(recs, columns=["hour", "tso", "name", "eic", "fmax", "fref", "imax", "presolved"])
    if not frame.empty:
        frame["hour"] = pd.to_datetime(frame["hour"], utc=True)
    return frame


def daily_aggregate(frame: pd.DataFrame, tag: str,
                    schwelle_hoch: float = 0.80, schwelle_ueberlast: float = 1.00) -> pd.DataFrame:
    """Eine Tagesdatei aus allen Stundenzeilen des Tages erzeugen (eine Zeile je Netzelement)."""
    if frame is None or frame.empty:
        return pd.DataFrame(columns=DAILY_COLUMNS)

    f = frame.copy()
    if "imax" not in f:
        f["imax"] = np.nan
    if "presolved" not in f:
        f["presolved"] = False
    f["name"] = f["name"].map(norm_text)
    f["eic"] = f["eic"].map(norm_eic)
    f["key"] = [element_key(e, n) for e, n in zip(f["eic"], f["name"])]
    # Hilfszeilen ausschließen (siehe Kommentar oben); fehlendes Imax (NaN) bleibt erhalten
    f = f[(f["fmax"] > 0) & ~(f["imax"] >= IMAX_UNBEGRENZT) & ~f["name"].str.match(_TECH_NAME)].copy()
    f["u"] = f["fref"].abs() / f["fmax"]
    if f.empty:
        return pd.DataFrame(columns=DAILY_COLUMNS)

    # 1) ungünstigster Fall je Stunde und Element ('bindend' = in dieser Stunde mind. eine Zeile mit Presolved = TRUE)
    f["presolved"] = f["presolved"].astype(bool)
    hourly = f.groupby(["key", "hour"], as_index=False).agg(u=("u", "max"), fmax=("fmax", "min"), bind=("presolved", "max"))

    # 2) Kennzahlen je Element über alle Stunden des Tages
    hourly["hoch"] = (hourly["u"] >= schwelle_hoch).astype(int)
    hourly["ueber"] = (hourly["u"] >= schwelle_ueberlast).astype(int)
    kenn = hourly.groupby("key", as_index=False).agg(
        Fmax_MW=("fmax", "min"),
        Stunden=("u", "size"),
        Auslastung_Summe=("u", "sum"),
        Auslastung_Max=("u", "max"),
        Stunden_hoch=("hoch", "sum"),
        Stunden_Ueberlast=("ueber", "sum"),
        Stunden_bindend=("bind", "sum"),
    )

    # 3) Beschreibende Felder: Hauptname (häufigster), alle Namensvarianten, ÜNB des Hauptnamens
    cnt = f.groupby(["key", "name", "tso"]).size().reset_index(name="n")
    cnt = cnt.sort_values(["key", "n", "name", "tso"], ascending=[True, False, True, True])
    haupt = cnt.drop_duplicates("key")[["key", "name", "tso"]].rename(
        columns={"name": "Netzelement", "tso": "UENB"})
    namen = f.groupby("key")["name"].agg(lambda s: " | ".join(sorted({x for x in s if x}))).reset_index(
        name="Namen")

    out = kenn.merge(haupt, on="key").merge(namen, on="key")
    out["EIC"] = [k if not k.startswith("NAME|") else "" for k in out["key"]]
    out["Tag"] = tag
    out = out.rename(columns={"key": "Schluessel_JAO"})
    out["Stunden"] = out["Stunden"].astype(int)
    out["Stunden_hoch"] = out["Stunden_hoch"].astype(int)
    out["Stunden_Ueberlast"] = out["Stunden_Ueberlast"].astype(int)
    out["Stunden_bindend"] = out["Stunden_bindend"].astype(int)
    return out[DAILY_COLUMNS].sort_values("Schluessel_JAO").reset_index(drop=True)
