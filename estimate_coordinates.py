#!/usr/bin/env python3
"""Lage der Umspannwerke bestimmen -> data/processed/umspannwerke_koordinaten.csv (Blatt D_Koordinaten).

Rangfolge je Umspannwerk:
  1. data/manual/koordinaten_manuell.csv   exakte Werte (z. B. aus OpenStreetMap); überschreiben alles
  2. data/manual/koordinaten_anker.csv     ungefähre Ortslage (Näherung, ungeprüft); wird gegen das Netz geprüft
  3. Schätzung aus Leitungslängen und Netztopologie (siehe src/kapmodell/geo.py)
Knoten, deren Name bei mehreren Betreibern vorkommt und verschiedene Standorte verbindet, bleiben ohne Koordinate.
Zusätzlich entsteht data/manual/koordinaten_pruefbericht.csv (entfernte Anker, mehrdeutige Namen) zur fachlichen Prüfung.
"""
from __future__ import annotations

import sys

import pandas as pd

import _common as common
from kapmodell import geo

KLASSE = [(25, "geschätzt (Fehler meist unter 25 km)"), (50, "geschätzt (Fehler meist unter 50 km)"),
          (1e9, "geschätzt (unsicher, über 50 km möglich)")]


def main() -> int:
    proc, man = common.ROOT / "data" / "processed", common.ROOT / "data" / "manual"
    lines = pd.read_csv(proc / "sgm_leitungen.csv", dtype=str, keep_default_na=False)
    nodes = list(pd.read_csv(proc / "sgm_umspannwerke.csv", dtype=str, keep_default_na=False)["Umspannwerk"])
    anker = pd.read_csv(man / "koordinaten_anker.csv", dtype={"Umspannwerk": str})
    manuell = pd.read_csv(man / "koordinaten_manuell.csv", dtype={"Umspannwerk": str})
    guessed = {r.Umspannwerk: (float(r.Breitengrad), float(r.Laengengrad)) for r in anker.itertuples() if r.Umspannwerk in nodes}
    exact = {r.Umspannwerk: (float(r.Breitengrad), float(r.Laengengrad)) for r in manuell.itertuples() if r.Umspannwerk in nodes}

    edges = geo.edge_table(lines)
    ambiguous = {n for n, _ in geo.find_ambiguous_nodes(nodes, edges, {**guessed, **exact})} - set(exact)
    edges2 = edges[~edges["u"].isin(ambiguous) & ~edges["v"].isin(ambiguous)]
    kept, removed = geo.clean_anchors(nodes, edges2, {a: c for a, c in guessed.items() if a not in ambiguous})
    anchors = {**kept, **exact}
    est, alpha, _ = geo.estimate_positions(nodes, edges2, anchors)
    e = est.set_index("Umspannwerk")

    rows = []
    for n in nodes:
        if n in exact:
            rows.append((n, *exact[n], "exakt (manuell oder OpenStreetMap)", "data/manual/koordinaten_manuell.csv"))
        elif n in e.index and bool(e.loc[n, "Anker"]):
            rows.append((n, e.loc[n, "Breitengrad"], e.loc[n, "Laengengrad"], "Anker (ca., Näherung ungeprüft)", "Ortslage der gleichnamigen Gemeinde"))
        elif n in e.index:
            klasse = next(t for lim, t in KLASSE if e.loc[n, "Netzabstand_Anker_km"] <= lim)
            rows.append((n, e.loc[n, "Breitengrad"], e.loc[n, "Laengengrad"], klasse, "Schätzung aus Leitungslängen und Netztopologie"))
        else:
            grund = "nicht eindeutig (Name verbindet vermutlich mehrere Standorte)" if n in ambiguous else "keine Schätzung (nicht im zusammenhängenden Netz)"
            rows.append((n, "", "", grund, ""))
    out = pd.DataFrame(rows, columns=["Umspannwerk", "Breitengrad", "Laengengrad", "Genauigkeit", "Quelle"])
    for c in ("Breitengrad", "Laengengrad"):
        out[c] = pd.to_numeric(out[c], errors="coerce").round(4)
    out.to_csv(proc / "umspannwerke_koordinaten.csv", index=False, lineterminator="\n", encoding="utf-8")
    report = [(n, "Mehrdeutiger Name: verbindet vermutlich verschiedene Standorte (gleicher Name bei mehreren Betreibern?)") for n in sorted(ambiguous)]
    report += [(n, f"Anker entfernt: passt nicht zu den Leitungslängen ({k} Verstöße), Koordinate oder Namensverwechslung prüfen") for n, k in removed]
    pd.DataFrame(report, columns=["Umspannwerk", "Befund"]).to_csv(man / "koordinaten_pruefbericht.csv", index=False, lineterminator="\n", encoding="utf-8")
    print(f"{len(out)} Umspannwerke: {sum(1 for r in rows if r[3].startswith('Anker'))} Anker, "
          f"{sum(1 for r in rows if r[3].startswith('geschätzt'))} geschätzt, {int(out['Breitengrad'].isna().sum())} ohne Koordinate "
          f"({len(ambiguous)} mehrdeutig, {len(removed)} Anker entfernt); alpha = {alpha:.2f}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
