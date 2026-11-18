#!/usr/bin/env python3
"""Plausibilitätsprüfung der Netzstammdaten (SGM-CSV) vor jeder Übernahme.

Gibt Warnungen aus und endet mit Fehlercode 1, wenn Pflichtspalten fehlen oder Werte unmöglich sind.
"""
from __future__ import annotations

import sys

import pandas as pd

import _common as common

LINE_COLS = ["NE_name", "EIC_Code", "UENB", "Umspannwerk_1", "Umspannwerk_2", "Spannung_kV",
             "Imax_Periode_1_A", "Imax_Periode_2_A", "Imax_Periode_3_A", "Imax_Periode_4_A",
             "Imax_Periode_5_A", "Imax_Periode_6_A", "Imax_fix_A", "DLR_min_A", "DLR_max_A",
             "R_Ohm", "X_Ohm", "B_uS", "Laenge_km", "Kommentar_JAO", "Typ", "Nachbarland"]
TRAFO_COLS = ["Name", "EIC_Code", "UENB", "Imax_min_A", "Imax_max_A", "Imax_fix_A", "U_primaer_kV",
              "U_sekundaer_kV", "R_Ohm", "X_Ohm", "B_uS", "G_uS", "Stufen_RAO", "Winkel_Grad",
              "Sym_Asym", "du_Phase_Prozent", "du_Winkel_Prozent", "Kommentar_JAO", "Umspannwerk", "Typ"]
LINE_TYPES = {"Leitung", "Kuppelleitung zwischen dt. ÜNB", "Grenzkuppelleitung"}


def check(lines: pd.DataFrame, trafos: pd.DataFrame) -> tuple[list[str], list[str]]:
    errors, warnings = [], []
    for name, df, cols in (("sgm_leitungen", lines, LINE_COLS), ("sgm_trafos", trafos, TRAFO_COLS)):
        missing = [c for c in cols if c not in df.columns]
        if missing:
            errors.append(f"{name}: Spalten fehlen: {missing}")
        elif list(df.columns[:len(cols)]) != cols:
            errors.append(f"{name}: Spaltenreihenfolge weicht ab (Excel liest nach Position!)")
    if errors:
        return errors, warnings
    spannung = pd.to_numeric(lines["Spannung_kV"], errors="coerce")
    if spannung.isna().any():
        # Lücke in den Quelldaten: Smax bleibt für diese Zeilen leer, das Modell läuft weiter
        warnings.append(f"sgm_leitungen: {int(spannung.isna().sum())} Zeilen ohne Spannung (Smax bleibt leer)")
    unbekannt = set(lines["Typ"]) - LINE_TYPES
    if unbekannt:
        errors.append(f"sgm_leitungen: unbekannte Typen {sorted(unbekannt)}")
    gk = lines[lines["Typ"] == "Grenzkuppelleitung"]
    if (gk["Nachbarland"] == "").any():
        warnings.append("Grenzkuppelleitungen ohne Nachbarland vorhanden")
    mit_eic = lines[lines["EIC_Code"] != ""]
    dubletten = mit_eic.duplicated(["EIC_Code", "NE_name"], keep=False).sum()
    if dubletten:
        warnings.append(f"sgm_leitungen: {int(dubletten)} Zeilen mit gleicher EIC und gleichem Namen "
                        "(mögliche Dubletten oder zwei Pole/Stränge – fachlich prüfen, nicht automatisch löschen)")
    if trafos[trafos["EIC_Code"] != ""]["EIC_Code"].duplicated().any():
        warnings.append("sgm_trafos: doppelte EIC-Codes")
    ohne = int((lines["EIC_Code"] == "").sum())
    if ohne:
        warnings.append(f"sgm_leitungen: {ohne} Zeilen ohne EIC (Zuordnung nur über den Namen)")
    return errors, warnings


def main() -> int:
    proc = common.ROOT / "data" / "processed"
    lines = pd.read_csv(proc / "sgm_leitungen.csv", dtype=str, keep_default_na=False)
    trafos = pd.read_csv(proc / "sgm_trafos.csv", dtype=str, keep_default_na=False)
    errors, warnings = check(lines, trafos)
    for w in warnings:
        print("WARNUNG:", w)
    for e in errors:
        print("FEHLER:", e)
    print(f"{len(lines)} Stromkreise, {len(trafos)} Trafos geprüft – {len(errors)} Fehler, {len(warnings)} Warnungen")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
