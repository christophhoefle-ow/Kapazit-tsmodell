#!/usr/bin/env python3
"""Schlägt für JAO-Elemente ohne Stammdaten-Treffer ähnlich benannte SGM-Elemente vor.

Das Ergebnis (data/manual/alias_vorschlaege.csv) ist NUR eine Arbeitshilfe. Nichts davon wird
automatisch angewendet: Ein Fachmensch prüft jede Zeile und kopiert bestätigte Zuordnungen nach
data/manual/alias_jao_sgm.csv (Spalten: Schluessel_JAO, Schluessel_SGM, Bemerkung).
"""
from __future__ import annotations

import difflib
import sys

import pandas as pd

import _common as common
from kapmodell.mapping import sgm_elements
from kapmodell.normalize import name_key


def main() -> int:
    proc = common.ROOT / "data" / "processed"
    mon = pd.read_csv(proc / "auslastung_monatlich.csv", dtype=str, keep_default_na=False)
    offen = mon[mon["Zuordnung"] == "nicht zugeordnet"].drop_duplicates("Schluessel")
    sgm = sgm_elements(pd.read_csv(proc / "sgm_leitungen.csv", dtype=str, keep_default_na=False),
                       pd.read_csv(proc / "sgm_trafos.csv", dtype=str, keep_default_na=False))
    namen = {name_key(n): (k, n) for k, n in zip(sgm["sgm_key"], sgm["name"])}
    rows = []
    for _, r in offen.iterrows():
        best = difflib.get_close_matches(name_key(r["Netzelement"]), list(namen), n=1, cutoff=0.6)
        if best:
            sk, sname = namen[best[0]]
            score = difflib.SequenceMatcher(None, name_key(r["Netzelement"]), best[0]).ratio()
            rows.append((r["Schluessel"], r["Netzelement"], r["UENB"], r["Auslastung_Max"], sk, sname, round(score, 2)))
        else:
            rows.append((r["Schluessel"], r["Netzelement"], r["UENB"], r["Auslastung_Max"], "", "", 0))
    out = pd.DataFrame(rows, columns=["Schluessel_JAO", "JAO_Name", "UENB", "Auslastung_Max",
                                      "Vorschlag_Schluessel_SGM", "Vorschlag_SGM_Name", "Aehnlichkeit"])
    out = out.sort_values("Auslastung_Max", ascending=False, key=lambda s: pd.to_numeric(s))
    out.to_csv(common.ROOT / "data" / "manual" / "alias_vorschlaege.csv", index=False,
               lineterminator="\n", encoding="utf-8")
    print(f"{len(out)} offene Elemente, {int((out['Aehnlichkeit'] > 0).sum())} mit Vorschlag -> data/manual/alias_vorschlaege.csv")
    return 0


if __name__ == "__main__":
    sys.exit(main())
