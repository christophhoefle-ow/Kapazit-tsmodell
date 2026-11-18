#!/usr/bin/env python3
"""Prüft ein BERECHNETES Modell (mit zwischengespeicherten Werten) gegen eine unabhängige Neuberechnung in pandas.

    python scripts/check_model.py Kapazitaetsmodell_DE_V3.xlsx

Voraussetzung: Die Datei wurde einmal in Excel geöffnet und gespeichert (oder per LibreOffice berechnet),
damit die Formelwerte in der Datei stehen. Geprüft werden – mit den Standard-Parametern
(Imax-Ansatz MIN, Schwellen 80 % / 5 %, Mindeststunden 1, Fenster enthält alle Daten):

  * Status je Stromkreis und Trafo (Anzahlen)
  * je Umspannwerk: Typ, Anzahl Leitungen/Grenzkuppeln/Trafos, bewertete/Engpass/hoch belastete Elemente,
    höchste Auslastung, Netz-Ampel, Σ Smax (Leitungen, Grenzkuppel), Σ Snenn (Trafos)

Endet mit Fehlercode 1, wenn etwas abweicht. Nicht geprüft: Anschluss-Ampel (D_Netzanschluss), Dashboard-Layout.
"""
from __future__ import annotations

import sys
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook

import _common as common
from kapmodell.normalize import element_key

OK_STATUS = ("Engpass", "hoch belastet", "unkritisch")


def _num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s.replace("", np.nan))


def reference(root: Path):
    p = root / "data" / "processed"
    L = pd.read_csv(p / "sgm_leitungen.csv", dtype=str, keep_default_na=False)
    T = pd.read_csv(p / "sgm_trafos.csv", dtype=str, keep_default_na=False)
    U = pd.read_csv(p / "sgm_umspannwerke.csv", dtype=str, keep_default_na=False)
    M = pd.read_csv(p / "auslastung_monatlich.csv", dtype={"Schluessel": str}, keep_default_na=False)
    for c in ("Stunden", "Auslastung_Max", "Stunden_Ueberlast"):
        M[c] = pd.to_numeric(M[c])
    agg = M.groupby("Schluessel").agg(h=("Stunden", "sum"), ue=("Stunden_Ueberlast", "sum"), mx=("Auslastung_Max", "max"))

    def status(key):
        if key not in agg.index:
            return "keine Daten"
        a = agg.loc[key]
        if a["ue"] / a["h"] >= 0.05:
            return "Engpass"
        return "hoch belastet" if a["mx"] >= 0.8 else "unkritisch"

    L["status"] = [status(element_key(e, n)) for e, n in zip(L["EIC_Code"], L["NE_name"])]
    L["util"] = [agg.loc[element_key(e, n), "mx"] if element_key(e, n) in agg.index else np.nan
                 for e, n in zip(L["EIC_Code"], L["NE_name"])]
    per = pd.concat([_num(L[f"Imax_Periode_{i}_A"]) for i in range(1, 7)], axis=1)
    has_per = per.notna().any(axis=1)
    imax = _num(L["Imax_fix_A"])
    imax = imax.where(imax.notna(), per.min(axis=1).where(has_per))
    dlr = pd.concat([_num(L["DLR_min_A"]), _num(L["DLR_max_A"])], axis=1)
    imax = imax.where(imax.notna() | has_per, dlr["DLR_min_A"].where(dlr["DLR_min_A"].notna(), dlr["DLR_max_A"]))
    L["smax"] = np.sqrt(3) * _num(L["Spannung_kV"]) * imax / 1000
    L["gk"] = L["Typ"] == "Grenzkuppelleitung"

    T["status"] = [status(element_key(e, n)) for e, n in zip(T["EIC_Code"], T["Name"])]
    T["util"] = [agg.loc[element_key(e, n), "mx"] if element_key(e, n) in agg.index else np.nan
                 for e, n in zip(T["EIC_Code"], T["Name"])]
    tfix = _num(T["Imax_fix_A"])
    tmin = pd.concat([_num(T["Imax_min_A"]), _num(T["Imax_max_A"])], axis=1).min(axis=1)
    timax = np.where(tfix.notna() & (tfix < 99999), tfix, np.where(tfix.notna(), np.nan, tmin))
    T["snenn"] = np.sqrt(3) * _num(T["U_primaer_kV"]) * timax / 1000

    rows = []
    for uw in U["Umspannwerk"]:
        a = L[(L["Umspannwerk_1"] == uw) | (L["Umspannwerk_2"] == uw)]     # ein Stromkreis zählt je Knoten einmal
        t = T[T["Umspannwerk"] == uw]
        st = list(a["status"]) + list(t["status"])
        bew = [s for s in st if s in OK_STATUS]
        utils = pd.concat([a.loc[a["status"].isin(OK_STATUS), "util"], t.loc[t["status"].isin(OK_STATUS), "util"]])
        n_lt, n_tr = int((~a["gk"]).sum()), len(t)
        rows.append(dict(
            uw=uw, typ="Abzweig (Y-Knoten)" if uw.startswith("Y-") else ("Netzknoten DE" if n_lt + n_tr > 0 else "nur Kuppelleitung (meist Ausland)"),
            n_lt=n_lt, n_gk=int(a["gk"].sum()), n_tr=n_tr, n_bew=len(bew), n_eng=st.count("Engpass"), n_hoch=st.count("hoch belastet"),
            amp="Grau" if not bew else ("Rot" if st.count("Engpass") else ("Gelb" if st.count("hoch belastet") else "Grün")),
            util=utils.max() if len(utils) else np.nan, smax_lt=a.loc[~a["gk"], "smax"].sum(),
            smax_gk=a.loc[a["gk"], "smax"].sum(), snenn=t["snenn"].sum()))
    return pd.DataFrame(rows), L, T


def main() -> int:
    if len(sys.argv) != 2:
        print(__doc__)
        return 2
    ref, L, T = reference(common.ROOT)
    wb = load_workbook(sys.argv[1], data_only=True)
    ws = wb["Umspannwerke"]
    cols = {"uw": 1, "typ": 3, "util": 6, "n_lt": 9, "n_gk": 10, "n_tr": 11, "n_bew": 12, "n_eng": 13, "n_hoch": 14, "amp": 16,
            "smax_lt": 17, "smax_gk": 18, "snenn": 19}
    x = pd.DataFrame({k: [ws.cell(r, c).value for r in range(5, 5 + len(ref))] for k, c in cols.items()})
    if x["amp"].isna().all():
        print("FEHLER: Keine berechneten Werte in der Datei. Erst in Excel öffnen und speichern (oder per LibreOffice berechnen).")
        return 2
    failed = 0

    def report(name, ok, detail=""):
        nonlocal failed
        failed += 0 if ok else 1
        print(f"  {'ok  ' if ok else 'FAIL'} {name} {detail}")

    report("Umspannwerke: Namen und Reihenfolge", bool((ref["uw"].values == x["uw"].values).all()))
    for k in ("typ", "n_lt", "n_gk", "n_tr", "n_bew", "n_eng", "n_hoch", "amp"):
        diff = int((ref[k].values != x[k].values).sum())
        report(f"Umspannwerke: {k}", diff == 0, f"({diff} Abweichungen)" if diff else "")
    for k in ("util", "smax_lt", "smax_gk", "snenn"):
        d = (pd.to_numeric(ref[k]).fillna(-1) - pd.to_numeric(x[k].replace("", np.nan)).fillna(-1)).abs().max()
        report(f"Umspannwerke: {k}", d < 1e-6, f"(max. Abweichung {d:.6f})" if d >= 1e-6 else "")
    ws2 = wb["Leitungen"]
    xs = pd.Series([ws2.cell(r, 23).value for r in range(5, 5 + len(L))])
    report("Leitungen: Status je Stromkreis", bool((xs.values == L["status"].values).all()),
           str(dict(L["status"].value_counts())))
    ws3 = wb["Trafos"]
    xt = pd.Series([ws3.cell(r, 20).value for r in range(5, 5 + len(T))])
    report("Trafos: Status je Trafo", bool((xt.values == T["status"].values).all()), str(dict(T["status"].value_counts())))
    print("\n" + ("ALLE PRÜFUNGEN BESTANDEN" if not failed else f"{failed} PRÜFUNG(EN) FEHLGESCHLAGEN"))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
