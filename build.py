"""Erzeugt Kapazitaetsmodell_DE_V3.xlsx aus den Lieferdateien in data/processed.

Aufbau in drei Schichten (Farbcode der Reiter):
  grau      D_*            Rohdaten (werden per Power Query oder Kopieren ersetzt)
  dunkelblau Leitungen, Trafos   Berechnung je Netzelement
  grün      Umspannwerke, Dashboard   Ergebnis (Ampel)
"""
from __future__ import annotations

import csv
import datetime as dt
import math
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Font
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.table import Table, TableColumn
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.chart import ScatterChart, Series
from openpyxl.chart.legend import LegendEntry
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.data_source import StrRef
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from . import connections as conn
from .outline import OUTLINE
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Border, Side
from openpyxl.utils import column_index_from_string
from .style import FONT_TITLE, INPUT_BLUE, WHITE_SEP

from . import content as C
from .style import (AMPEL, BLUE_TAB, BORDER, CENTER, GOLD, GREEN_D, GREY, L_BLUE, L_GOLD, L_GREEN, L_GREY, L_GREY2, NAVY,
                    WRAP_TOP, col, color_formula, fill, font, header, row_height, set_widths)

# --------------------------------------------------------------------------- Kapazitäten
R0 = 5                         # erste Datenzeile aller Tabellenblätter
N_LINES, N_TRAFOS, N_UW, N_NA, N_BEL = 1200, 250, 650, 500, 700
A_LAST = 20004                 # letzte Zeile in D_Auslastung, die Formeln lesen
L_LAST, T_LAST, U_LAST = R0 + N_LINES - 1, R0 + N_TRAFOS - 1, R0 + N_UW - 1
NA_LAST, BEL_LAST = R0 + N_NA - 1, R0 + N_BEL - 1

SHEETS = ["Lies_mich", "Modelllogik", "Dashboard", "Umspannwerke", "Leitungen", "Trafos", "Karte_Daten", "Karte_Linien", "Parameter",
          "Datenstand", "Glossar", "Quellen", "Automatisierung", "D_SGM_Leitungen", "D_SGM_Trafos", "D_SGM_Umspannwerke",
          "D_Auslastung", "D_Netzanschluss", "D_Belegung", "D_Koordinaten", "D_Meta"]
TAB = {"Lies_mich": "CEECFF", "Modelllogik": "CEECFF", "Glossar": "CEECFF", "Quellen": "CEECFF", "Automatisierung": "CEECFF",
       "Dashboard": "000F47", "Umspannwerke": "000F47", "Leitungen": "82BAFF", "Trafos": "82BAFF", "Karte_Daten": "82BAFF",
       "Karte_Linien": "82BAFF", "Parameter": "0B4BFF", "Datenstand": "0B4BFF"}
CSV_FILE = {"D_SGM_Leitungen": "sgm_leitungen.csv", "D_SGM_Trafos": "sgm_trafos.csv", "D_SGM_Umspannwerke": "sgm_umspannwerke.csv",
            "D_Auslastung": "auslastung_monatlich.csv", "D_Netzanschluss": "netzanschluss_uenb.csv", "D_Belegung": "belegung_mastr.csv",
            "D_Koordinaten": "umspannwerke_koordinaten.csv", "D_Meta": "meta.csv"}
_STATE = {"connections": False, "specs": []}      # wird in build_workbook gesetzt

BLUE_FONT = font(color="0B4BFF")

# --------------------------------------------------------------------------- Spaltenköpfe der Rohdaten (= CSV)
H_LINES = ["NE_name", "EIC_Code", "UENB", "Umspannwerk_1", "Umspannwerk_2", "Spannung_kV", "Imax_Periode_1_A",
           "Imax_Periode_2_A", "Imax_Periode_3_A", "Imax_Periode_4_A", "Imax_Periode_5_A", "Imax_Periode_6_A",
           "Imax_fix_A", "DLR_min_A", "DLR_max_A", "R_Ohm", "X_Ohm", "B_uS", "Laenge_km", "Kommentar_JAO", "Typ",
           "Nachbarland"]
F_LINES = ["Name des Stromkreises", "EIC-Code", "Übertragungsnetzbetreiber", "Umspannwerk 1", "Umspannwerk 2",
           "Spannung (kV)", "Imax Periode 1 (A)", "Imax Periode 2 (A)", "Imax Periode 3 (A)", "Imax Periode 4 (A)",
           "Imax Periode 5 (A)", "Imax Periode 6 (A)", "Imax fix (A)", "DLR min (A)", "DLR max (A)", "R (Ω)", "X (Ω)",
           "B (µS)", "Länge (km)", "Kommentar JAO", "Typ", "Nachbarland"]
H_TRAFOS = ["Name", "EIC_Code", "UENB", "Imax_min_A", "Imax_max_A", "Imax_fix_A", "U_primaer_kV", "U_sekundaer_kV",
            "R_Ohm", "X_Ohm", "B_uS", "G_uS", "Stufen_RAO", "Winkel_Grad", "Sym_Asym", "du_Phase_Prozent",
            "du_Winkel_Prozent", "Kommentar_JAO", "Umspannwerk", "Typ"]
F_TRAFOS = ["Name", "EIC-Code", "ÜNB", "Imax min (A)", "Imax max (A)", "Imax fix (A)", "Spannung primär (kV)",
            "Spannung sekundär (kV)", "R (Ω)", "X (Ω)", "B (µS)", "G (µS)", "Stufen für RAO", "Winkel θ (°)",
            "Sym./Asym.", "δu Phase (%)", "δu Winkel (%)", "Kommentar JAO", "Umspannwerk", "Typ"]
H_AUSL = ["Monat", "Schluessel", "Zuordnung", "UENB", "Netzelement", "EIC", "Fmax_MW", "Stunden", "Auslastung_Summe",
          "Auslastung_Max", "Stunden_hoch", "Stunden_Ueberlast"]
F_AUSL = ["Monat", "Schlüssel (EIC oder NAME|Name)", "Zuordnung", "ÜNB (laut JAO)", "Netzelement (laut JAO)",
          "EIC (laut JAO)", "kleinster Fmax im Monat (MW)", "Stunden mit Daten", "Summe der Stunden-Auslastungen",
          "höchste Auslastung", "Stunden ab 'hoch'", "Stunden ab 'Überlast'"]
H_NA = ["Umspannwerk", "UENB", "Spannungsebene", "Einspeisung_frei_von_MW", "Einspeisung_frei_bis_MW", "Last_frei_von_MW",
        "Last_frei_bis_MW", "Verfuegbar_ab_Jahr", "Stand", "Quelle_URL", "Bemerkung"]
F_NA = ["Umspannwerk (Schreibweise wie im Blatt Umspannwerke)", "ÜNB", "Spannungsebene", "Einspeisung frei von (MW)",
        "Einspeisung frei bis (MW)", "Last frei von (MW)", "Last frei bis (MW)", "verfügbar ab (Jahr)",
        "Stand der Karte", "Quelle (Link)", "Bemerkung"]
H_BEL = ["Umspannwerk", "Belegt_Einspeisung_MW", "Belegt_Last_MW", "Beantragt_Einspeisung_MW", "Beantragt_Last_MW",
         "Stand", "Quelle"]
F_BEL = ["Umspannwerk", "belegt Einspeisung (MW)", "belegt Last (MW)", "beantragt Einspeisung (MW)",
         "beantragt Last (MW)", "Stand", "Quelle"]

H_KO = ["Umspannwerk", "Breitengrad", "Laengengrad", "Genauigkeit", "Quelle"]
F_KO = ["Umspannwerk", "Breitengrad (°N)", "Längengrad (°O)", "Genauigkeit der Lage", "Quelle der Lage"]
NUM_KO = {"Breitengrad", "Laengengrad"}
NUM_LINES = {"Spannung_kV", "Imax_Periode_1_A", "Imax_Periode_2_A", "Imax_Periode_3_A", "Imax_Periode_4_A",
             "Imax_Periode_5_A", "Imax_Periode_6_A", "Imax_fix_A", "DLR_min_A", "DLR_max_A", "R_Ohm", "X_Ohm", "B_uS",
             "Laenge_km"}
NUM_TRAFOS = {"Imax_min_A", "Imax_max_A", "Imax_fix_A", "U_primaer_kV", "U_sekundaer_kV", "R_Ohm", "X_Ohm", "B_uS",
              "G_uS", "Winkel_Grad", "du_Phase_Prozent", "du_Winkel_Prozent"}
NUM_AUSL = {"Fmax_MW", "Stunden", "Auslastung_Summe", "Auslastung_Max", "Stunden_hoch", "Stunden_Ueberlast"}
NUM_NA = {"Einspeisung_frei_von_MW", "Einspeisung_frei_bis_MW", "Last_frei_von_MW", "Last_frei_bis_MW", "Verfuegbar_ab_Jahr"}
NUM_BEL = {"Belegt_Einspeisung_MW", "Belegt_Last_MW", "Beantragt_Einspeisung_MW", "Beantragt_Last_MW"}


# --------------------------------------------------------------------------- Daten laden
def _read(path: Path) -> pd.DataFrame:
    return pd.read_csv(path, dtype=str, keep_default_na=False)


def load_inputs(root: Path) -> dict:
    proc = root / "data" / "processed"
    data = {n: _read(proc / f"{n}.csv") for n in ("sgm_leitungen", "sgm_trafos", "sgm_umspannwerke",
                                                  "auslastung_monatlich", "netzanschluss_uenb", "belegung_mastr", "meta", "umspannwerke_koordinaten")}
    data["quellen"] = list(csv.reader(open(root / "config" / "quellen.csv", encoding="utf-8")))
    data["glossar"] = list(csv.reader(open(root / "config" / "glossar.csv", encoding="utf-8")))
    return data


def _cell_value(v: str, numeric: bool):
    if v == "":
        return None
    if numeric:
        f = float(v)
        return int(f) if f.is_integer() else f
    return v


# --------------------------------------------------------------------------- Rohdatenblätter
def landing_sheet(ws, title_text, desc, friendly, tech, df, numeric, widths, date_col=None, pct_cols=(), note_fill=None):
    """Rohdatenblatt: Zeile 1-2 Erklärung, Zeile 3 Klartext-Köpfe, Zeile 4 technische Köpfe (= CSV), ab Zeile 5 Daten."""
    ws.sheet_properties.tabColor = GREY
    ws["A1"] = title_text
    ws["A1"].font = font(bold=True, size=13, color=NAVY)
    ws["A2"] = desc
    ws["A2"].font = font(italic=True, size=9, color="595959")
    ws["A2"].alignment = Alignment(wrap_text=False, vertical="top")
    for j, (fh, th) in enumerate(zip(friendly, tech), start=1):
        c3 = ws.cell(3, j, fh)
        c3.font = font(italic=True, size=9, color="595959")
        c3.fill = fill(L_GREY)
        c3.alignment = CENTER
        c3.border = BORDER
        header(ws.cell(4, j, th), "data")
    ws.row_dimensions[3].height = 38
    ws.row_dimensions[4].height = 30
    for r, row in enumerate(df.itertuples(index=False), start=R0):
        for j, (v, name) in enumerate(zip(row, tech), start=1):
            if date_col and name == date_col:
                cell = ws.cell(r, j, dt.date.fromisoformat(v))
                cell.number_format = "mm/yyyy"
            else:
                cell = ws.cell(r, j, _cell_value(v, name in numeric))
                if name in pct_cols:
                    cell.number_format = "0.0%"
            cell.font = BLUE_FONT
    set_widths(ws, widths)
    ws.freeze_panes = "A5"
    last = max(R0, R0 + len(df) - 1)
    if _STATE["connections"] and ws.title in CSV_FILE:
        # Bereichsabfrage: Excel ersetzt beim Aktualisieren den Inhalt ab Zeile 4 (Kopf) durch die CSV-Datei aus dem Internet
        idx = len(_STATE["specs"]) + 1
        kinds = ["date" if t == date_col else ("number" if t in numeric else "text") for t in tech]
        ws.defined_names[f"ExternalData_{idx}"] = DefinedName(f"ExternalData_{idx}", attr_text=f"'{ws.title}'!$A$4:${col(len(tech))}${last}", hidden=True)
        _STATE["specs"].append(dict(sheet=ws.title, file=CSV_FILE[ws.title], columns=list(zip(tech, kinds))))
    else:
        ws.auto_filter.ref = f"A4:{col(len(tech))}{last}"
    _fit_width(ws, title_rows="3:4")


def build_landing_sheets(wb, data):
    # ---- Leitungen
    w = {col(i): 16 for i in range(1, 23)}
    w.update({"A": 44, "B": 20, "C": 16, "D": 18, "E": 18, "T": 28, "U": 28, "V": 12})
    landing_sheet(wb["D_SGM_Leitungen"], "D_SGM_Leitungen – Rohdaten: Stromkreise (JAO Static Grid Model)",
                  "Quelle: JAO Core Static Grid Model (9. Release, 03/2026) · Lieferung: data/processed/sgm_leitungen.csv · "
                  "Nicht von Hand ändern – wird durch Power Query oder Kopieren ersetzt (Zeile 4 = Kopf, ab Zeile 5 Daten).",
                  F_LINES, H_LINES, data["sgm_leitungen"], NUM_LINES, w)
    # ---- Trafos
    w = {col(i): 15 for i in range(1, 21)}
    w.update({"A": 36, "B": 20, "C": 16, "R": 28, "S": 18, "T": 16})
    landing_sheet(wb["D_SGM_Trafos"], "D_SGM_Trafos – Rohdaten: Transformatoren und Phasenschieber (JAO Static Grid Model)",
                  "Quelle: JAO Core Static Grid Model (9. Release, 03/2026) · Lieferung: data/processed/sgm_trafos.csv · "
                  "Nicht von Hand ändern.", F_TRAFOS, H_TRAFOS, data["sgm_trafos"], NUM_TRAFOS, w)
    # ---- Umspannwerke
    landing_sheet(wb["D_SGM_Umspannwerke"], "D_SGM_Umspannwerke – Rohdaten: Liste aller Umspannwerke / Netzknoten",
                  "Abgeleitet aus Umspannwerk 1/2 der Stromkreise und Umspannwerk der Trafos · Lieferung: "
                  "data/processed/sgm_umspannwerke.csv · Schreibweise wie im JAO-Datensatz (Varianten zwischen ÜNB möglich).",
                  ["Umspannwerk / Netzknoten"], ["Umspannwerk"], data["sgm_umspannwerke"], set(), {"A": 40})
    # ---- Auslastung
    landing_sheet(wb["D_Auslastung"], "D_Auslastung – Rohdaten: Auslastung je Netzelement und Monat (JAO Final Computation)",
                  "Quelle: JAO Publication Tool Core · Lieferung: data/processed/auslastung_monatlich.csv · Je Zeile ein Monat und ein "
                  "Netzelement. Mittelwert = Summe ÷ Stunden. Nicht von Hand ändern.",
                  F_AUSL, H_AUSL, data["auslastung_monatlich"], NUM_AUSL,
                  {"A": 11, "B": 26, "C": 16, "D": 14, "E": 44, "F": 20, "G": 14, "H": 11, "I": 16, "J": 13, "K": 12, "L": 12},
                  date_col="Monat", pct_cols=("Auslastung_Max",))
    # ---- Netzanschluss (Eingabeblatt, gold)
    ws = wb["D_Netzanschluss"]
    landing_sheet(ws, "D_Netzanschluss – Freie Anschlusskapazität laut Netzanschlusskarten der ÜNB (Bandbreiten)",
                  "Eingabe oder Lieferung aus data/processed/netzanschluss_uenb.csv. Formatbeispiel mit ERFUNDENEN Werten (nicht eintragen): "
                  "Beispielhausen | TENNETGMBH | 380/110 kV | 500 | 1000 | 0 | 500 | 2032 | 01.04.2026 | https://… · "
                  "'unter 500 MW' = von 0 bis 500. Name = Schreibweise im Blatt Umspannwerke.",
                  F_NA, H_NA, data["netzanschluss_uenb"], NUM_NA,
                  {"A": 30, "B": 14, "C": 16, "D": 14, "E": 14, "F": 14, "G": 14, "H": 12, "I": 14, "J": 50, "K": 30})
    ws.sheet_properties.tabColor = GREY
    # ---- Belegung
    landing_sheet(wb["D_Belegung"], "D_Belegung – Belegung und Pipeline je Umspannwerk (MaStR, Phase 2)",
                  "Noch nicht befüllt. Vorgesehen: Aggregat aus dem Marktstammdatenregister je Umspannwerk · "
                  "Lieferung: data/processed/belegung_mastr.csv.",
                  F_BEL, H_BEL, data["belegung_mastr"], NUM_BEL, {"A": 30, "B": 16, "C": 14, "D": 16, "E": 16, "F": 14, "G": 40})
    # ---- Koordinaten (Lage der Umspannwerke: manuell/OSM exakt, sonst geschätzt)
    landing_sheet(wb["D_Koordinaten"], "D_Koordinaten – Lage der Umspannwerke (für die Netzkarte im Dashboard)",
                  "Lieferung: data/processed/umspannwerke_koordinaten.csv · Rangfolge: exakt (manuell/OSM) vor Anker (ca.) vor Schätzung aus Leitungslängen. "
                  "Die Schätzung ist eine NÄHERUNG (Median-Abweichung rund 16 km, einzelne Knoten über 50 km). Spalte 'Genauigkeit' beachten.",
                  F_KO, H_KO, data["umspannwerke_koordinaten"], NUM_KO, {"A": 30, "B": 13, "C": 13, "D": 44, "E": 52})
    # ---- Meta (Werte immer als Text: so liefert es Power Query)
    ws = wb["D_Meta"]
    landing_sheet(ws, "D_Meta – Datenstand und Einstellungen der Pipeline",
                  "Lieferung: data/processed/meta.csv · Werte stehen als Text (Dezimalpunkt) – die Formeln lesen sie locale-unabhängig.",
                  ["Schlüssel", "Wert"], ["Schluessel", "Wert"], data["meta"], set(), {"A": 28, "B": 80})


# --------------------------------------------------------------------------- Parameter
def meta_ref(key: str, text=True) -> str:
    """Formel-Baustein: Wert aus D_Meta (Text)."""
    return f'INDEX(D_Meta!$B$5:$B$60,MATCH("{key}",D_Meta!$A$5:$A$60,0))'


PARAMS = [
    # (Name, Bezeichnung, Wert, Einheit, Erklärung, Herkunft, Wirkt auf, Format, Typ)
    ("§", "A  Zeitraum der Auslastungsdaten"),
    ("P_FensterMonate", "Auswertungsfenster (Anzahl Monate)", 3, "Monate",
     "Wie viele Monate werden für die Auslastung zusammengefasst? Eine einzelne Stunde sagt wenig (nur 0 % oder 100 % Überlast). "
     "Ab etwa einem Quartal wird das Bild stabil, 12 Monate zeigen auch die Jahreszeiten.",
     "Neu in V3 – Vorschlag, bitte im Team abstimmen", "Leitungen, Trafos, Umspannwerke", "0", "input"),
    ("P_EndMonat", "Letzter Monat im Fenster (leer = neuester Monat mit Daten)", None, "Datum",
     "Leer lassen für die aktuelle Auswertung. Ein Datum (z. B. 01.09.2026) blickt auf einen früheren Zeitraum zurück.",
     "Neu in V3", "Leitungen, Trafos, Umspannwerke", "mm/yyyy", "input"),
    ("P_Bis", "→ Fenster bis (berechnet)", f'=IF(P_EndMonat="",IF(COUNT(D_Auslastung!$A${R0}:$A${A_LAST})=0,"",MAX(D_Auslastung!$A${R0}:$A${A_LAST})),DATE(YEAR(P_EndMonat),MONTH(P_EndMonat),1))',
     "Datum", "Ergibt sich aus der Eingabe darüber und den vorhandenen Daten.", "Formel", "–", "mm/yyyy", "calc"),
    ("P_Von", "→ Fenster von (berechnet)", '=IF(P_Bis="","",EDATE(P_Bis,-(P_FensterMonate-1)))', "Datum",
     "Erster Monat des Fensters.", "Formel", "–", "mm/yyyy", "calc"),
    ("P_DatenVon", "→ Daten vorhanden ab", f'=IF(COUNT(D_Auslastung!$A${R0}:$A${A_LAST})=0,"",MIN(D_Auslastung!$A${R0}:$A${A_LAST}))',
     "Datum", "Ältester Monat in D_Auslastung.", "Formel", "–", "mm/yyyy", "calc"),
    ("P_DatenBis", "→ Daten vorhanden bis", f'=IF(COUNT(D_Auslastung!$A${R0}:$A${A_LAST})=0,"",MAX(D_Auslastung!$A${R0}:$A${A_LAST}))',
     "Datum", "Neuester Monat in D_Auslastung.", "Formel", "–", "mm/yyyy", "calc"),
    ("P_MinStunden", "Mindestanzahl Stunden je Netzelement", 1, "Stunden",
     "Mit weniger Stunden steht ein Element auf 'zu wenig Daten' statt auf einer vorschnellen Ampel. Für den Dauerbetrieb "
     "empfohlen: 168 (eine Woche). Aktuell 1, damit die Ein-Stunden-Testdaten ausgewertet werden.",
     "Neu in V3 – Vorschlag, bitte abstimmen", "Leitungen, Trafos", "0", "input"),
    ("§", "B  Netzkapazität (Technik)"),
    ("P_ImaxAnsatz", "Imax-Ansatz bei saisonalen bzw. DLR-Werten", "MIN", "MIN / MAX",
     "MIN = vorsichtig (kleinster saisonaler Wert bzw. kleinster DLR-Wert), MAX = optimistisch (größter Wert).",
     "Annahme Projektteam (V2)", "Leitungen, Trafos", "@", "input"),
    ("P_CosPhi", "Leistungsfaktor cos φ (Umrechnung MVA → MW)", 1, "–",
     "Im Übertragungsnetz näherungsweise 1. Dann gilt: Wirkleistung in MW = Scheinleistung in MVA.",
     "Annahme Projektteam (V2)", "Leitungen", "0.00", "input"),
    ("P_NichtBegrenzend", "Wert für 'nicht begrenzend' (Trafos)", 99999, "A",
     "Trafos mit Imax 99999 A sind laut Daten in der Kapazitätsberechnung nicht begrenzend und werden nicht als Kapazität gewertet.",
     "Laut JAO-Daten (V2)", "Trafos", "0", "input"),
    ("§", "C  Ampel-Regeln Netzauslastung (JAO)"),
    ("P_SchwelleHoch", "Schwelle 'hoch belastet': höchste Auslastung ab", 0.8, "% der Grenze",
     "Ab dieser Auslastung im ungünstigsten Stundenwert des Fensters gilt ein Element als 'hoch belastet' (Gelb).",
     "Annahme Projektteam (V2)", "Leitungen, Trafos, Umspannwerke", "0%", "input"),
    ("P_SchwelleEngpass", "Schwelle 'Engpass': Anteil Überlast-Stunden ab", 0.05, "% der Stunden",
     "Überlast-Stunde = Stunde mit Auslastung ab 100 % im ungünstigsten Ausfallfall. Ab diesem Anteil gilt ein Element als 'Engpass' (Rot).",
     "Annahme Projektteam (V2)", "Leitungen, Trafos, Umspannwerke", "0%", "input"),
    ("§", "D  Ampel-Regeln Anschlusskapazität (ÜNB-Karten)"),
    ("P_FreiGruenAb", "Grün, wenn freie Kapazität (untere Bandgrenze) mindestens", 100, "MW",
     "Grün nur, wenn selbst die untere Bandgrenze mindestens so hoch ist. Rot, wenn die obere Bandgrenze 0 MW ist. Dazwischen Gelb. "
     "Beispiel aus dem Auftakt: 120 MW frei = grün.",
     "Neu in V3 – Vorschlag, bitte abstimmen", "Umspannwerke", "0", "input"),
    ("P_NichtAusgewiesen", "Umspannwerk fehlt in D_Netzanschluss (obwohl die Tabelle Einträge hat)", "GRAU", "GRAU / ROT",
     "GRAU = keine Aussage (Standard). ROT = 'nicht als frei ausgewiesen' gilt als ausgeschöpft (nach Angabe der ÜNB sind nahezu alle "
     "freien Anschlusspunkte für Jahre reserviert). Wirkt nur, wenn D_Netzanschluss befüllt ist.",
     "Neu in V3 – Vorschlag, bitte abstimmen", "Umspannwerke", "@", "input"),
]


def build_parameter(wb):
    ws = wb["Parameter"]
    ws["A1"] = "Parameter und Annahmen"
    ws["A1"].font = font(bold=True, size=14, color=NAVY)
    ws["A2"] = ("Blaue Werte auf hellblauem Grund sind Annahmen: hier ändern, alle Blätter rechnen neu. "
                "Grau hinterlegte Zeilen mit '→' sind berechnet.")
    ws["A2"].font = font(italic=True, size=9, color="595959")
    for j, h in enumerate(["", "Parameter", "Wert", "Einheit", "Was bedeutet das? (einfach erklärt)", "Herkunft", "Wirkt auf"], start=1):
        if h:
            header(ws.cell(4, j, h), "input")
    set_widths(ws, {"A": 2, "B": 52, "C": 14, "D": 14, "E": 88, "F": 34, "G": 30})
    r = R0
    dv_imax = DataValidation(type="list", formula1='"MIN,MAX"', allow_blank=False)
    dv_na = DataValidation(type="list", formula1='"GRAU,ROT"', allow_blank=False)
    ws.add_data_validation(dv_imax)
    ws.add_data_validation(dv_na)
    for p in PARAMS:
        if p[0] == "§":
            ws.cell(r, 2, p[1]).font = font(bold=True, color=NAVY)
            for j in range(2, 8):
                ws.cell(r, j).fill = fill("82BAFF")
            r += 1
            continue
        name, label, value, unit, expl, src, affects, fmt, kind = p
        ws.cell(r, 2, label).font = font(bold=(kind == "input"))
        vc = ws.cell(r, 3, value)
        vc.number_format = fmt
        vc.alignment = Alignment(horizontal="center", vertical="top")
        if kind == "input":
            vc.fill = fill(L_GOLD)
            vc.font = font(color="0B4BFF", bold=True)
            vc.border = BORDER
        else:
            vc.font = font(color="000F47" if "!" in str(value) else "000000")
            vc.fill = fill(L_GREY)
        ws.cell(r, 4, unit).font = font(color="595959")
        ws.cell(r, 5, expl).font = font(size=9)
        ws.cell(r, 6, src).font = font(size=9, italic=True, color="595959")
        ws.cell(r, 7, affects).font = font(size=9, color="595959")
        for j in (2, 4, 5, 6, 7):
            ws.cell(r, j).alignment = WRAP_TOP
        ws.row_dimensions[r].height = row_height(expl, 88 * 1.05, line_pt=12)
        wb.defined_names[name] = DefinedName(name, attr_text=f"Parameter!$C${r}")
        if name == "P_ImaxAnsatz":
            dv_imax.add(vc)
        if name == "P_NichtAusgewiesen":
            dv_na.add(vc)
        r += 1
    # Abgleich mit der Pipeline
    r += 1
    ws.cell(r, 2, "E  Abgleich mit der Pipeline").font = font(bold=True, color=NAVY)
    for j in range(2, 8):
        ws.cell(r, j).fill = fill("82BAFF")
    r += 1
    ws.cell(r, 2, "Schwelle 'hoch' in der Pipeline (Spalte 'Stunden ab hoch')").font = font()
    c = ws.cell(r, 3, f'=IFERROR(VALUE(SUBSTITUTE({meta_ref("schwelle_hoch")},".",MID(1/2,2,1))),"")')   # locale-sicher (Dezimalpunkt -> Dezimalzeichen des PCs)
    c.number_format = "0%"
    c.fill = fill(L_GREY)
    c.alignment = Alignment(horizontal="center")
    color_formula(c)
    ws.cell(r, 5, "Die Spalte 'Stunden ab hoch' in D_Auslastung wird in der Pipeline mit dieser festen Schwelle gezählt (config/config.yml). "
                  "Der Status in Excel nutzt dagegen die höchste Auslastung und die Schwelle oben.").font = font(size=9)
    ws.cell(r, 5).alignment = WRAP_TOP
    ws.row_dimensions[r].height = 28
    wb.defined_names["P_PipelineHoch"] = DefinedName("P_PipelineHoch", attr_text=f"Parameter!$C${r}")
    ws.freeze_panes = "A5"
    ws.sheet_view.zoomScale = 100
    _fit_width(ws)


def _fit_width(ws, landscape=True, title_rows=None):
    """Druck: A4, eine Seite breit, optional Kopfzeile auf jeder Seite wiederholen."""
    from openpyxl.worksheet.properties import PageSetupProperties
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = ws.sheet_properties.pageSetUpPr or PageSetupProperties()
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    if title_rows:
        ws.print_title_rows = title_rows


# --------------------------------------------------------------------------- Hilfen für Rechenblätter
def DA(c: str) -> str:
    return f"D_Auslastung!${c}${R0}:${c}${A_LAST}"


def band(ws, first, last, row, text, kind):
    ws.merge_cells(f"{first}{row}:{last}{row}")
    c = ws[f"{first}{row}"]
    c.value = text
    for cc in ws[f"{first}{row}:{last}{row}"][0]:
        header(cc, kind)


def add_ampel_cf(ws, rng):
    for name, (bg, fg) in AMPEL.items():
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{name}"'], fill=fill(bg),
                                                      font=Font(name="Arial", color=fg, bold=True)))


def add_status_cf(ws, rng):
    for text, bg in (("Engpass", "F6D3D2"), ("hoch belastet", "FBEDBF"), ("unkritisch", "DDF0CC"),
                     ("zu wenig Daten", "EDEDED"), ("keine Daten", "EDEDED")):
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{text}"'], fill=fill(bg)))


def add_util_cf(ws, rng, first_cell):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({first_cell}),{first_cell}>=1)"], fill=fill("F6D3D2")))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({first_cell}),{first_cell}>=P_SchwelleHoch)"], fill=fill("FBEDBF")))


def write_formulas(ws, r, formulas: dict, fmts: dict):
    for letter, formula in formulas.items():
        c = ws[f"{letter}{r}"]
        c.value = formula
        color_formula(c)
        if letter in fmts:
            c.number_format = fmts[letter]


def _calc_headers(ws, headers: dict, widths: dict, height=48):
    for letter, text in headers.items():
        header(ws[f"{letter}4"], "calc")
        ws[f"{letter}4"].value = text
    ws.row_dimensions[4].height = height
    set_widths(ws, widths)


def _jao_block(key_col, name_col, eic_col, r, key_cell):
    """Formeln des JAO-Blocks (identisch für Leitungen und Trafos); Rückgabe als Textbausteine."""
    crit = f'{DA("B")},{key_cell},{DA("A")},">="&P_Von,{DA("A")},"<="&P_Bis'
    return crit


# --------------------------------------------------------------------------- Leitungen
def build_leitungen(wb, n_data):
    ws = wb["Leitungen"]
    ws["A1"] = "Leitungen – Kapazität, Auslastung und Status je Stromkreis (Höchstspannung)"
    ws["A1"].font = font(bold=True, size=14, color=NAVY)
    ws["A2"] = (f"Zeile n in diesem Blatt = Zeile n in D_SGM_Leitungen (Platz für {N_LINES} Stromkreise). Alle Spalten sind Formeln. "
                "Auslastung = Werte aus D_Auslastung im Auswertungsfenster (Blatt Parameter).")
    ws["A2"].font = font(italic=True, size=9, color="595959")
    ws.merge_cells("A2:L2")
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[2].height = 28
    band(ws, "A", "I", 3, "STAMMDATEN (Kopie aus D_SGM_Leitungen)", "data")
    band(ws, "J", "M", 3, "KAPAZITÄT (berechnet)", "calc")
    band(ws, "N", "U", 3, "AUSLASTUNG aus D_Auslastung (Fenster laut Parameter)", "calc")
    band(ws, "V", "X", 3, "ERGEBNIS", "result")
    band(ws, "Y", "Z", 3, "KONTROLLE", "data")
    heads = {"A": "Name des Stromkreises", "B": "EIC-Code", "C": "ÜNB", "D": "Umspannwerk 1", "E": "Umspannwerk 2",
             "F": "Spannung (kV)", "G": "Typ", "H": "Nachbarland", "I": "Länge (km)",
             "J": "Imax Basis (A)", "K": "Imax Quelle", "L": "Smax (MVA)", "M": "Pmax (MW)",
             "N": "Schlüssel (EIC oder NAME|Name)", "O": "Stunden mit Daten", "P": "Auslastung Ø", "Q": "Auslastung max",
             "R": "Stunden ab 'hoch'", "S": "Stunden ab 'Überlast'", "T": "Anteil Überlast-Stunden", "U": "Fmax laut JAO (MW)",
             "V": "Rest-Kapazität (MW)", "W": "Status", "X": "Ampel", "Y": "Mögliche Dublette?", "Z": "Bewertet (1/0)"}
    widths = {"A": 44, "B": 20, "C": 14, "D": 18, "E": 18, "F": 9, "G": 26, "H": 10, "I": 9, "J": 10, "K": 10, "L": 10,
              "M": 10, "N": 26, "O": 10, "P": 10, "Q": 10, "R": 10, "S": 10, "T": 11, "U": 11, "V": 12, "W": 14, "X": 9,
              "Y": 11, "Z": 9}
    _calc_headers(ws, heads, widths)
    fm = {"F": "0", "I": "0.0", "J": "#,##0", "L": "#,##0", "M": "#,##0", "O": "#,##0", "P": "0.0%", "Q": "0.0%",
          "R": "#,##0", "S": "#,##0", "T": "0.0%", "U": "#,##0", "V": "#,##0"}
    d = "D_SGM_Leitungen"
    for r in range(R0, L_LAST + 1):
        crit = _jao_block("N", "A", "B", r, f"$N{r}")
        f = {
            "A": f'=IF({d}!A{r}="","",{d}!A{r})',
            "B": f'=IF($A{r}="","",{d}!B{r}&"")',
            "C": f'=IF($A{r}="","",{d}!C{r}&"")',
            "D": f'=IF($A{r}="","",{d}!D{r}&"")',
            "E": f'=IF($A{r}="","",{d}!E{r}&"")',
            "F": f'=IF(OR($A{r}="",{d}!F{r}=""),"",{d}!F{r})',
            "G": f'=IF($A{r}="","",{d}!U{r}&"")',
            "H": f'=IF($A{r}="","",{d}!V{r}&"")',
            "I": f'=IF(OR($A{r}="",{d}!S{r}=""),"",{d}!S{r})',
            "J": (f'=IF($A{r}="","",IF(ISNUMBER({d}!M{r}),{d}!M{r},IF(COUNT({d}!G{r}:L{r})>0,'
                  f'IF(P_ImaxAnsatz="MAX",MAX({d}!G{r}:L{r}),MIN({d}!G{r}:L{r})),'
                  f'IF(AND(P_ImaxAnsatz="MAX",ISNUMBER({d}!O{r})),{d}!O{r},IF(ISNUMBER({d}!N{r}),{d}!N{r},'
                  f'IF(ISNUMBER({d}!O{r}),{d}!O{r},""))))))'),
            "K": (f'=IF($A{r}="","",IF(ISNUMBER({d}!M{r}),"fix",IF(COUNT({d}!G{r}:L{r})>0,"saisonal",'
                  f'IF(OR(ISNUMBER({d}!N{r}),ISNUMBER({d}!O{r})),"DLR","fehlt"))))'),
            "L": f'=IF(OR($A{r}="",$J{r}="",$F{r}=""),"",SQRT(3)*$F{r}*$J{r}/1000)',
            "M": f'=IF($L{r}="","",$L{r}*P_CosPhi)',
            "N": (f'=IF($A{r}="","",IF(OR($B{r}="",ISNUMBER(SEARCH("NOTAVAILABLE",$B{r}))),'
                  f'"NAME|"&TRIM(SUBSTITUTE($A{r},CHAR(160)," ")),TRIM($B{r})))'),
            "O": f'=IF($A{r}="","",SUMIFS({DA("H")},{crit}))',
            "P": f'=IF(OR($A{r}="",$O{r}=0),"",SUMIFS({DA("I")},{crit})/$O{r})',
            "Q": f'=IF(OR($A{r}="",$O{r}=0),"",_xlfn.MAXIFS({DA("J")},{crit}))',
            "R": f'=IF($A{r}="","",SUMIFS({DA("K")},{crit}))',
            "S": f'=IF($A{r}="","",SUMIFS({DA("L")},{crit}))',
            "T": f'=IF(OR($A{r}="",$O{r}=0),"",$S{r}/$O{r})',
            "U": f'=IF(OR($A{r}="",$O{r}=0),"",_xlfn.MINIFS({DA("G")},{crit}))',
            "V": f'=IF(OR($A{r}="",$O{r}=0),"",$U{r}*MAX(0,1-$Q{r}))',
            "W": (f'=IF($A{r}="","",IF($O{r}=0,"keine Daten",IF($O{r}<P_MinStunden,"zu wenig Daten",'
                  f'IF($T{r}>=P_SchwelleEngpass,"Engpass",IF($Q{r}>=P_SchwelleHoch,"hoch belastet","unkritisch")))))'),
            "X": f'=IF($A{r}="","",IF($W{r}="Engpass","Rot",IF($W{r}="hoch belastet","Gelb",IF($W{r}="unkritisch","Grün","Grau"))))',
            "Y": (f'=IF($A{r}="","",IF(OR($B{r}="",ISNUMBER(SEARCH("NOTAVAILABLE",$B{r}))),"",'
                  f'IF(COUNTIFS($B${R0}:$B${L_LAST},$B{r},$A${R0}:$A${L_LAST},$A{r})>1,"prüfen","")))'),
            "Z": f'=IF($A{r}="","",IF(OR($W{r}="Engpass",$W{r}="hoch belastet",$W{r}="unkritisch"),1,0))',
        }
        write_formulas(ws, r, f, fm)
    add_status_cf(ws, f"W{R0}:W{L_LAST}")
    add_ampel_cf(ws, f"X{R0}:X{L_LAST}")
    add_util_cf(ws, f"Q{R0}:Q{L_LAST}", f"Q{R0}")
    ws.conditional_formatting.add(f"Y{R0}:Y{L_LAST}", CellIsRule(operator="equal", formula=['"prüfen"'], fill=fill("FBEDBF")))
    ws.freeze_panes = "B5"
    ws.auto_filter.ref = f"A4:Z{L_LAST}"
    ws.sheet_view.zoomScale = 90
    _fit_width(ws, title_rows="4:4")


# --------------------------------------------------------------------------- Trafos
def build_trafos(wb):
    ws = wb["Trafos"]
    ws["A1"] = "Trafos – Kapazität, Auslastung und Status je Transformator / Phasenschieber (380/220 kV)"
    ws["A1"].font = font(bold=True, size=14, color=NAVY)
    ws["A2"] = (f"Zeile n in diesem Blatt = Zeile n in D_SGM_Trafos (Platz für {N_TRAFOS}). Kapazität bezogen auf die Primärseite. "
                "Hinweis: 380/110-kV-Trafos (Schnittstelle zur Hochspannung) fehlen im JAO-Netzmodell.")
    ws["A2"].font = font(italic=True, size=9, color="595959")
    ws.merge_cells("A2:L2")
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[2].height = 28
    band(ws, "A", "G", 3, "STAMMDATEN (Kopie aus D_SGM_Trafos)", "data")
    band(ws, "H", "J", 3, "KAPAZITÄT (berechnet)", "calc")
    band(ws, "K", "R", 3, "AUSLASTUNG aus D_Auslastung (Fenster laut Parameter)", "calc")
    band(ws, "S", "U", 3, "ERGEBNIS", "result")
    band(ws, "V", "V", 3, "KONTROLLE", "data")
    heads = {"A": "Name", "B": "EIC-Code", "C": "ÜNB", "D": "Umspannwerk", "E": "Typ", "F": "Spannung primär (kV)",
             "G": "Spannung sekundär (kV)", "H": "Imax Basis (A)", "I": "Imax Quelle", "J": "Snenn (MVA)",
             "K": "Schlüssel (EIC oder NAME|Name)", "L": "Stunden mit Daten", "M": "Auslastung Ø", "N": "Auslastung max",
             "O": "Stunden ab 'hoch'", "P": "Stunden ab 'Überlast'", "Q": "Anteil Überlast-Stunden", "R": "Fmax laut JAO (MW)",
             "S": "Rest-Kapazität (MW)", "T": "Status", "U": "Ampel", "V": "Bewertet (1/0)"}
    widths = {"A": 36, "B": 20, "C": 14, "D": 18, "E": 16, "F": 10, "G": 10, "H": 10, "I": 14, "J": 10, "K": 26, "L": 10,
              "M": 10, "N": 10, "O": 10, "P": 10, "Q": 11, "R": 11, "S": 12, "T": 14, "U": 9, "V": 9}
    _calc_headers(ws, heads, widths)
    fm = {"F": "0", "G": "0", "H": "#,##0", "J": "#,##0", "L": "#,##0", "M": "0.0%", "N": "0.0%", "O": "#,##0", "P": "#,##0",
          "Q": "0.0%", "R": "#,##0", "S": "#,##0"}
    d = "D_SGM_Trafos"
    for r in range(R0, T_LAST + 1):
        crit = _jao_block("K", "A", "B", r, f"$K{r}")
        f = {
            "A": f'=IF({d}!A{r}="","",{d}!A{r})',
            "B": f'=IF($A{r}="","",{d}!B{r}&"")',
            "C": f'=IF($A{r}="","",{d}!C{r}&"")',
            "D": f'=IF($A{r}="","",{d}!S{r}&"")',
            "E": f'=IF($A{r}="","",{d}!T{r}&"")',
            "F": f'=IF(OR($A{r}="",{d}!G{r}=""),"",{d}!G{r})',
            "G": f'=IF(OR($A{r}="",{d}!H{r}=""),"",{d}!H{r})',
            "H": (f'=IF($A{r}="","",IF(AND(ISNUMBER({d}!F{r}),{d}!F{r}<P_NichtBegrenzend),{d}!F{r},'
                  f'IF(AND(ISNUMBER({d}!F{r}),{d}!F{r}>=P_NichtBegrenzend),"",IF(COUNT({d}!D{r}:E{r})>0,'
                  f'IF(P_ImaxAnsatz="MAX",MAX({d}!D{r}:E{r}),MIN({d}!D{r}:E{r})),""))))'),
            "I": (f'=IF($A{r}="","",IF(AND(ISNUMBER({d}!F{r}),{d}!F{r}>=P_NichtBegrenzend),"nicht begrenzend",'
                  f'IF(ISNUMBER({d}!F{r}),"fix",IF(COUNT({d}!D{r}:E{r})>0,"min/max","fehlt"))))'),
            "J": f'=IF(OR($A{r}="",$H{r}="",$F{r}=""),"",SQRT(3)*$F{r}*$H{r}/1000)',
            "K": (f'=IF($A{r}="","",IF(OR($B{r}="",ISNUMBER(SEARCH("NOTAVAILABLE",$B{r}))),'
                  f'"NAME|"&TRIM(SUBSTITUTE($A{r},CHAR(160)," ")),TRIM($B{r})))'),
            "L": f'=IF($A{r}="","",SUMIFS({DA("H")},{crit}))',
            "M": f'=IF(OR($A{r}="",$L{r}=0),"",SUMIFS({DA("I")},{crit})/$L{r})',
            "N": f'=IF(OR($A{r}="",$L{r}=0),"",_xlfn.MAXIFS({DA("J")},{crit}))',
            "O": f'=IF($A{r}="","",SUMIFS({DA("K")},{crit}))',
            "P": f'=IF($A{r}="","",SUMIFS({DA("L")},{crit}))',
            "Q": f'=IF(OR($A{r}="",$L{r}=0),"",$P{r}/$L{r})',
            "R": f'=IF(OR($A{r}="",$L{r}=0),"",_xlfn.MINIFS({DA("G")},{crit}))',
            "S": f'=IF(OR($A{r}="",$L{r}=0),"",$R{r}*MAX(0,1-$N{r}))',
            "T": (f'=IF($A{r}="","",IF($L{r}=0,"keine Daten",IF($L{r}<P_MinStunden,"zu wenig Daten",'
                  f'IF($Q{r}>=P_SchwelleEngpass,"Engpass",IF($N{r}>=P_SchwelleHoch,"hoch belastet","unkritisch")))))'),
            "U": f'=IF($A{r}="","",IF($T{r}="Engpass","Rot",IF($T{r}="hoch belastet","Gelb",IF($T{r}="unkritisch","Grün","Grau"))))',
            "V": f'=IF($A{r}="","",IF(OR($T{r}="Engpass",$T{r}="hoch belastet",$T{r}="unkritisch"),1,0))',
        }
        write_formulas(ws, r, f, fm)
    add_status_cf(ws, f"T{R0}:T{T_LAST}")
    add_ampel_cf(ws, f"U{R0}:U{T_LAST}")
    add_util_cf(ws, f"N{R0}:N{T_LAST}", f"N{R0}")
    ws.freeze_panes = "B5"
    ws.auto_filter.ref = f"A4:V{T_LAST}"
    ws.sheet_view.zoomScale = 90
    _fit_width(ws, title_rows="4:4")


# --------------------------------------------------------------------------- Umspannwerke (Ergebnis)
def build_umspannwerke(wb):
    ws = wb["Umspannwerke"]
    ws["A1"] = "Umspannwerke – Ampel je Netzknoten (Einspeisung und Entnahme)"
    ws["A1"].font = font(bold=True, size=14, color=NAVY)
    ws["A2"] = ("So lesen: Zuerst die grünen Spalten D–H (Ergebnis und Begründung). Rechts davon steht, wie das Ergebnis zustande kommt. "
                "Das ungünstigste angeschlossene Netzelement bestimmt die Netz-Ampel; liegen ÜNB-Angaben vor, zählt je Richtung der "
                "ungünstigere Wert. Filter 'Knotentyp = Netzknoten DE' zeigt nur echte deutsche Umspannwerke.")
    ws["A2"].font = font(italic=True, size=9, color="595959")
    ws.merge_cells("A2:H2")
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[2].height = 40
    band(ws, "A", "C", 3, "UMSPANNWERK", "data")
    band(ws, "D", "H", 3, "ERGEBNIS – zuerst lesen", "result")
    band(ws, "I", "P", 3, "NETZ: Auslastung der angeschlossenen Elemente (JAO)", "calc")
    band(ws, "Q", "S", 3, "TECHNISCHE KAPAZITÄT (Summen)", "calc")
    band(ws, "T", "Z", 3, "ANSCHLUSS: freie Kapazität laut ÜNB-Karten", "calc")
    band(ws, "AA", "AD", 3, "BELEGUNG (MaStR, Phase 2)", "calc")
    band(ws, "AE", "AH", 3, "HILFSSPALTEN", "data")
    heads = {"A": "Umspannwerk / Netzknoten", "B": "ÜNB", "C": "Knotentyp",
             "D": "AMPEL Einspeisung", "E": "AMPEL Entnahme", "F": "Höchste Auslastung (Netz)", "G": "Datengrundlage",
             "H": "Begründung", "I": "Anzahl Leitungen (Inland)", "J": "Anzahl Grenzkuppel-leitungen", "K": "Anzahl Trafos",
             "L": "Netzelemente mit JAO-Daten", "M": "davon Engpass", "N": "davon hoch belastet",
             "O": "Rest-Kapazität engstes Element (MW)", "P": "Ampel Netz", "Q": "Σ Smax Leitungen (MVA)",
             "R": "Σ Smax Grenzkuppel (MVA)", "S": "Σ Snenn Trafos (MVA)", "T": "Einspeisung frei von (MW)",
             "U": "Einspeisung frei bis (MW)", "V": "Last frei von (MW)", "W": "Last frei bis (MW)",
             "X": "Ampel Anschluss Einspeisung", "Y": "Ampel Anschluss Entnahme", "Z": "Stand ÜNB-Karte",
             "AA": "belegt Einspeisung (MW)", "AB": "belegt Last (MW)", "AC": "beantragt Einspeisung (MW)",
             "AD": "beantragt Last (MW)", "AE": "Sortwert", "AF": "Rang Auslastung", "AG": "Zeile ÜNB-Karte", "AH": "Zeile Belegung"}
    widths = {"A": 28, "B": 18, "C": 32, "D": 12, "E": 12, "F": 12, "G": 15, "H": 64, "I": 9, "J": 10, "K": 8, "L": 11,
              "M": 9, "N": 10, "O": 14, "P": 9, "Q": 12, "R": 12, "S": 12, "T": 11, "U": 11, "V": 11, "W": 11, "X": 12,
              "Y": 12, "Z": 12, "AA": 11, "AB": 11, "AC": 11, "AD": 11, "AE": 11, "AF": 9, "AG": 9, "AH": 9}
    _calc_headers(ws, heads, widths, height=60)
    fm = {"F": "0.0%", "O": "#,##0", "Q": "#,##0", "R": "#,##0", "S": "#,##0", "T": "#,##0", "U": "#,##0", "V": "#,##0",
          "W": "#,##0", "AA": "#,##0", "AB": "#,##0", "AC": "#,##0", "AD": "#,##0"}
    LT = lambda c: f"Leitungen!${c}${R0}:${c}${L_LAST}"
    TR = lambda c: f"Trafos!${c}${R0}:${c}${T_LAST}"
    NA = lambda c: f"D_Netzanschluss!${c}${R0}:${c}${NA_LAST}"
    BE = lambda c: f"D_Belegung!${c}${R0}:${c}${BEL_LAST}"
    for r in range(R0, U_LAST + 1):
        a = f"$A{r}"
        f = {
            "A": f'=IF(D_SGM_Umspannwerke!A{r}="","",D_SGM_Umspannwerke!A{r})',
            "B": (f'=IF({a}="","",IFERROR(INDEX({LT("C")},MATCH({a},{LT("D")},0)),IFERROR(INDEX({LT("C")},MATCH({a},{LT("E")},0)),'
                  f'IFERROR(INDEX({TR("C")},MATCH({a},{TR("D")},0)),""))))'),
            "C": f'=IF({a}="","",IF(LEFT({a},2)="Y-","Abzweig (Y-Knoten)",IF($I{r}+$K{r}>0,"Netzknoten DE","nur Kuppelleitung (meist Ausland)")))',
            "D": (f'=IF({a}="","",IF(OR($P{r}="Rot",$X{r}="Rot"),"Rot",IF(OR($P{r}="Gelb",$X{r}="Gelb"),"Gelb",'
                  f'IF(OR($P{r}="Grün",$X{r}="Grün"),"Grün","Grau"))))'),
            "E": (f'=IF({a}="","",IF(OR($P{r}="Rot",$Y{r}="Rot"),"Rot",IF(OR($P{r}="Gelb",$Y{r}="Gelb"),"Gelb",'
                  f'IF(OR($P{r}="Grün",$Y{r}="Grün"),"Grün","Grau"))))'),
            "F": (f'=IF(OR({a}="",$L{r}=0),"",MAX(_xlfn.MAXIFS({LT("Q")},{LT("D")},{a},{LT("Z")},1),'
                  f'_xlfn.MAXIFS({LT("Q")},{LT("E")},{a},{LT("Z")},1),_xlfn.MAXIFS({TR("N")},{TR("D")},{a},{TR("V")},1)))'),
            "G": (f'=IF({a}="","",IF(AND($P{r}<>"Grau",OR($X{r}<>"Grau",$Y{r}<>"Grau")),"Netz + Anschluss",'
                  f'IF($P{r}<>"Grau","nur Netz",IF(OR($X{r}<>"Grau",$Y{r}<>"Grau"),"nur Anschluss","keine Daten"))))'),
            "H": (f'=IF({a}="","",IF($P{r}="Rot",$M{r}&" Netzelement(e) im Engpass (≥ "&TEXT(P_SchwelleEngpass,"0%")&" der Stunden über 100 %)",'
                  f'IF($P{r}="Gelb",$N{r}&" Netzelement(e) mit Auslastung ≥ "&TEXT(P_SchwelleHoch,"0%"),'
                  f'IF($P{r}="Grün","alle "&$L{r}&" bewerteten Netzelemente unter "&TEXT(P_SchwelleHoch,"0%"),"keine Auslastungsdaten")))'
                  f'&IF($X{r}="Rot","; ÜNB: Einspeisung nicht frei",IF($X{r}="Gelb","; ÜNB: Einspeisung "&$T{r}&"–"&$U{r}&" MW",'
                  f'IF($X{r}="Grün","; ÜNB: Einspeisung ab "&$T{r}&" MW frei","")))'
                  f'&IF($Y{r}="Rot","; ÜNB: Entnahme nicht frei",IF($Y{r}="Gelb","; ÜNB: Entnahme "&$V{r}&"–"&$W{r}&" MW",'
                  f'IF($Y{r}="Grün","; ÜNB: Entnahme ab "&$V{r}&" MW frei",""))))'),
            "I": (f'=IF({a}="","",COUNTIFS({LT("D")},{a},{LT("G")},"<>Grenzkuppelleitung")+'
                  f'COUNTIFS({LT("E")},{a},{LT("G")},"<>Grenzkuppelleitung")-'
                  f'COUNTIFS({LT("D")},{a},{LT("E")},{a},{LT("G")},"<>Grenzkuppelleitung"))'),
            "J": (f'=IF({a}="","",COUNTIFS({LT("D")},{a},{LT("G")},"Grenzkuppelleitung")+'
                  f'COUNTIFS({LT("E")},{a},{LT("G")},"Grenzkuppelleitung")-'
                  f'COUNTIFS({LT("D")},{a},{LT("E")},{a},{LT("G")},"Grenzkuppelleitung"))'),
            "K": f'=IF({a}="","",COUNTIFS({TR("D")},{a}))',
            "L": (f'=IF({a}="","",SUMIFS({LT("Z")},{LT("D")},{a})+SUMIFS({LT("Z")},{LT("E")},{a})-'
                  f'SUMIFS({LT("Z")},{LT("D")},{a},{LT("E")},{a})+SUMIFS({TR("V")},{TR("D")},{a}))'),
            "M": (f'=IF({a}="","",COUNTIFS({LT("D")},{a},{LT("W")},"Engpass")+COUNTIFS({LT("E")},{a},{LT("W")},"Engpass")-'
                  f'COUNTIFS({LT("D")},{a},{LT("E")},{a},{LT("W")},"Engpass")+COUNTIFS({TR("D")},{a},{TR("T")},"Engpass"))'),
            "N": (f'=IF({a}="","",COUNTIFS({LT("D")},{a},{LT("W")},"hoch belastet")+COUNTIFS({LT("E")},{a},{LT("W")},"hoch belastet")-'
                  f'COUNTIFS({LT("D")},{a},{LT("E")},{a},{LT("W")},"hoch belastet")+COUNTIFS({TR("D")},{a},{TR("T")},"hoch belastet"))'),
            "O": (f'=IF(OR({a}="",$L{r}=0),"",MIN('
                  f'IF(COUNTIFS({LT("D")},{a},{LT("Z")},1)>0,_xlfn.MINIFS({LT("V")},{LT("D")},{a},{LT("Z")},1),1E+99),'
                  f'IF(COUNTIFS({LT("E")},{a},{LT("Z")},1)>0,_xlfn.MINIFS({LT("V")},{LT("E")},{a},{LT("Z")},1),1E+99),'
                  f'IF(COUNTIFS({TR("D")},{a},{TR("V")},1)>0,_xlfn.MINIFS({TR("S")},{TR("D")},{a},{TR("V")},1),1E+99)))'),
            "P": f'=IF({a}="","",IF($L{r}=0,"Grau",IF($M{r}>0,"Rot",IF($N{r}>0,"Gelb","Grün"))))',
            "Q": (f'=IF({a}="","",SUMIFS({LT("L")},{LT("D")},{a},{LT("G")},"<>Grenzkuppelleitung")+'
                  f'SUMIFS({LT("L")},{LT("E")},{a},{LT("G")},"<>Grenzkuppelleitung")-'
                  f'SUMIFS({LT("L")},{LT("D")},{a},{LT("E")},{a},{LT("G")},"<>Grenzkuppelleitung"))'),
            "R": (f'=IF({a}="","",SUMIFS({LT("L")},{LT("D")},{a},{LT("G")},"Grenzkuppelleitung")+'
                  f'SUMIFS({LT("L")},{LT("E")},{a},{LT("G")},"Grenzkuppelleitung")-'
                  f'SUMIFS({LT("L")},{LT("D")},{a},{LT("E")},{a},{LT("G")},"Grenzkuppelleitung"))'),
            "S": f'=IF({a}="","",SUMIFS({TR("J")},{TR("D")},{a}))',
            "T": f'=IF(OR({a}="",$AG{r}=""),"",IF(INDEX({NA("D")},$AG{r})="","",INDEX({NA("D")},$AG{r})))',
            "U": f'=IF(OR({a}="",$AG{r}=""),"",IF(INDEX({NA("E")},$AG{r})="","",INDEX({NA("E")},$AG{r})))',
            "V": f'=IF(OR({a}="",$AG{r}=""),"",IF(INDEX({NA("F")},$AG{r})="","",INDEX({NA("F")},$AG{r})))',
            "W": f'=IF(OR({a}="",$AG{r}=""),"",IF(INDEX({NA("G")},$AG{r})="","",INDEX({NA("G")},$AG{r})))',
            "X": (f'=IF({a}="","",IF($AG{r}="",IF(AND(P_NichtAusgewiesen="ROT",COUNTA({NA("A")})>0),"Rot","Grau"),'
                  f'IF(OR($T{r}="",$U{r}=""),"Grau",IF($U{r}<=0,"Rot",IF($T{r}>=P_FreiGruenAb,"Grün","Gelb")))))'),
            "Y": (f'=IF({a}="","",IF($AG{r}="",IF(AND(P_NichtAusgewiesen="ROT",COUNTA({NA("A")})>0),"Rot","Grau"),'
                  f'IF(OR($V{r}="",$W{r}=""),"Grau",IF($W{r}<=0,"Rot",IF($V{r}>=P_FreiGruenAb,"Grün","Gelb")))))'),
            "Z": f'=IF(OR({a}="",$AG{r}=""),"",INDEX({NA("I")},$AG{r})&"")',
            "AA": f'=IF(OR({a}="",$AH{r}=""),"",IF(INDEX({BE("B")},$AH{r})="","",INDEX({BE("B")},$AH{r})))',
            "AB": f'=IF(OR({a}="",$AH{r}=""),"",IF(INDEX({BE("C")},$AH{r})="","",INDEX({BE("C")},$AH{r})))',
            "AC": f'=IF(OR({a}="",$AH{r}=""),"",IF(INDEX({BE("D")},$AH{r})="","",INDEX({BE("D")},$AH{r})))',
            "AD": f'=IF(OR({a}="",$AH{r}=""),"",IF(INDEX({BE("E")},$AH{r})="","",INDEX({BE("E")},$AH{r})))',
            "AE": f'=IF(OR({a}="",$F{r}="",$C{r}<>"Netzknoten DE"),"",$F{r}+ROW()/1000000000)',
            "AF": f'=IF($AE{r}="","",RANK($AE{r},$AE${R0}:$AE${U_LAST}))',
            "AG": f'=IF({a}="","",IFERROR(MATCH({a},{NA("A")},0),""))',
            "AH": f'=IF({a}="","",IFERROR(MATCH({a},{BE("A")},0),""))',
        }
        write_formulas(ws, r, f, fm)
        ws[f"H{r}"].alignment = Alignment(wrap_text=False, vertical="center")
    for rng in (f"D{R0}:E{U_LAST}", f"P{R0}:P{U_LAST}", f"X{R0}:Y{U_LAST}"):
        add_ampel_cf(ws, rng)
    add_util_cf(ws, f"F{R0}:F{U_LAST}", f"F{R0}")
    ws.freeze_panes = "D5"
    ws.auto_filter.ref = f"A4:AH{U_LAST}"
    ws.column_dimensions.group("AE", "AH", hidden=True)
    ws.sheet_view.zoomScale = 90
    _fit_width(ws, title_rows="3:4")


# --------------------------------------------------------------------------- Dashboard
def _tbl_header(ws, row, labels, kind="result", first_col=1):
    for j, t in enumerate(labels, start=first_col):
        header(ws.cell(row, j, t), kind)
    ws.row_dimensions[row].height = 32


def _sum_row(ws, row, first, last, cols, label="Summe"):
    ws.cell(row, 1, label).font = font(bold=True)
    for c in cols:
        cell = ws[f"{c}{row}"]
        cell.value = f"=SUM({c}{first}:{c}{last})"
        cell.font = font(bold=True)
        cell.number_format = "#,##0"
        cell.border = BORDER
    ws.cell(row, 1).border = BORDER


def build_dashboard(wb, data):
    ws = wb["Dashboard"]
    set_widths(ws, {"A": 22, "B": 22, "C": 17, "D": 11, "E": 12, "F": 12, "G": 12, "H": 12, "I": 12, "J": 28, "K": 3})
    ws["A1"] = "Dashboard – Wie voll ist das Netz? Ampel je Umspannwerk"
    ws["A1"].font = font(bold=True, size=14, color=NAVY)
    ws["A2"] = ("Alle Zahlen sind Formeln aus den Blättern Umspannwerke, Leitungen und Trafos. "
                "Auswertungsfenster und Schwellen stehen im Blatt Parameter.")
    ws["A2"].font = font(italic=True, size=9, color="595959")
    UW = lambda c: f"Umspannwerke!${c}${R0}:${c}${U_LAST}"
    LT = lambda c: f"Leitungen!${c}${R0}:${c}${L_LAST}"
    TR = lambda c: f"Trafos!${c}${R0}:${c}${T_LAST}"

    # Datenbasis (eine Zeile) und Kennzahlen-Kacheln
    ws["A2"] = "Quelle der Zahlen: Blätter Umspannwerke, Leitungen und Trafos. Auswertungsfenster und Schwellen: Blatt Parameter."
    ws["A2"].font = font(size=9, color="7B7974")
    for lc, label, vc, formula, fmt in (("A", "Auswertungsfenster", "B", "=P_Von", "mm/yyyy"), ("C", "bis", "D", "=P_Bis", "mm/yyyy"),
                                        ("E", "Stand (UTC)", "F", f'=IFERROR({meta_ref("erstellt_utc")},"–")', None),
                                        ("H", "Netzmodell", "I", f'=IFERROR({meta_ref("sgm_release")},"–")', None)):
        ws[f"{lc}3"] = label
        ws[f"{lc}3"].font = font(size=8, color="7B7974")
        ws[f"{vc}3"] = formula
        ws[f"{vc}3"].font = font(size=9, bold=True, color=NAVY)
        ws[f"{vc}3"].alignment = Alignment(horizontal="left")
        if fmt:
            ws[f"{vc}3"].number_format = fmt
    ws.merge_cells("F3:G3")
    ws.merge_cells("I3:J3")
    ws.row_dimensions[3].height = 18
    tiles = [("A", "A", "Netzknoten Deutschland", f'=COUNTIF({UW("C")},"Netzknoten DE")', "Umspannwerke im Modell", None, "#,##0"),
             ("B", "B", '="Rot · "&K_Richtung', '=IF(K_Richtung="Entnahme",C12,B12)', "ausgeschöpft / Engpass", "Rot", "0"),
             ("C", "D", '="Gelb · "&K_Richtung', '=IF(K_Richtung="Entnahme",C13,B13)', "knapp / hoch belastet", "Gelb", "0"),
             ("E", "F", '="Grün · "&K_Richtung', '=IF(K_Richtung="Entnahme",C14,B14)', "frei / unkritisch", "Grün", "0"),
             ("G", "H", '="Grau · "&K_Richtung', '=IF(K_Richtung="Entnahme",C15,B15)', "keine Daten", "Grau", "0"),
             ("I", "J", "Höchste Auslastung", f'=IFERROR(MAX({UW("F")}),"")', f'=IFERROR(INDEX({UW("A")},MATCH(MAX({UW("F")}),{UW("F")},0)),"")', None, "0%")]
    white = Side(style="thick", color="FFFFFF")
    for c1, c2, cap, val, sub, ampel, fmt in tiles:
        i1, i2 = column_index_from_string(c1), column_index_from_string(c2)
        for rr in (4, 5, 6, 7):
            for jj in range(i1, i2 + 1):
                cell = ws.cell(rr, jj)
                cell.fill = fill(L_BLUE)
                cell.border = Border(left=white, right=white, top=Side(style="thick", color=AMPEL[ampel][0]) if (ampel and rr == 4) else None)
        ws.cell(4, i1, cap)
        ws.cell(5, i1, val)
        ws.cell(7, i1, sub)
        ws.cell(4, i1).font = font(bold=True, size=9, color=NAVY)
        ws.cell(5, i1).font = font(bold=True, size=24, color=NAVY)
        ws.cell(5, i1).number_format = fmt
        ws.cell(7, i1).font = font(size=8, color="595959")
        for rr in (4, 5, 7):
            ws.cell(rr, i1).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for r1, r2 in ((4, 4), (5, 6), (7, 7)):
            if i2 > i1 or r2 > r1:
                ws.merge_cells(start_row=r1, start_column=i1, end_row=r2, end_column=i2)
    for rr, hh in ((4, 20), (5, 22), (6, 22), (7, 18)):
        ws.row_dimensions[rr].height = hh
    ws["A8"] = f'=IFERROR(IF({meta_ref("hinweis")}="","","Hinweis: "&{meta_ref("hinweis")}),"")'
    ws["A8"].font = font(bold=True, color="C53532", size=9)
    ws.merge_cells("A8:J8")
    ws["A8"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[8].height = 24

    # 2 Ampel-Überblick
    band(ws, "A", "J", 10, "Ampel-Überblick: Umspannwerke / Netzknoten in Deutschland", "result")
    _tbl_header(ws, 11, ["Ampel", "Einspeisung (Anzahl)", "Entnahme (Anzahl)", "nur Netz (Anzahl)", "Anteil Einspeisung"])
    for i, name in enumerate(["Rot", "Gelb", "Grün", "Grau"]):
        r = 12 + i
        ws.cell(r, 1, name).alignment = Alignment(horizontal="center")
        for c, src in (("B", "D"), ("C", "E"), ("D", "P")):
            cell = ws[f"{c}{r}"]
            cell.value = f'=COUNTIFS({UW(src)},$A{r},{UW("C")},"Netzknoten DE")'
            color_formula(cell)
            cell.border = BORDER
        ws[f"E{r}"] = f'=IF(B$16=0,"",B{r}/B$16)'
        ws[f"E{r}"].number_format = "0%"
        color_formula(ws[f"E{r}"])
        ws[f"E{r}"].border = BORDER
    _sum_row(ws, 16, 12, 15, ["B", "C", "D"])
    add_ampel_cf(ws, "A12:A15")
    ws["A17"] = "Rot = ausgeschöpft/Engpass · Gelb = knapp/hoch belastet · Grün = frei/unkritisch · Grau = keine Daten (nicht 'unkritisch')"
    ws["A17"].font = font(size=8, italic=True, color="595959")

    # 3 je ÜNB
    band(ws, "A", "F", 18, "Ampel je Übertragungsnetzbetreiber (Einspeisung)", "result")
    _tbl_header(ws, 19, ["ÜNB", "Rot", "Gelb", "Grün", "Grau", "Summe"])
    uenbs = sorted({v for v in data["sgm_leitungen"]["UENB"] if v})
    for i, name in enumerate(uenbs):
        r = 20 + i
        ws.cell(r, 1, name).font = font(color=NAVY)
        for c in "BCDE":
            cell = ws[f"{c}{r}"]
            cell.value = f'=COUNTIFS({UW("B")},$A{r},{UW("D")},{c}$19,{UW("C")},"Netzknoten DE")'
            color_formula(cell)
            cell.border = BORDER
        ws[f"F{r}"] = f"=SUM(B{r}:E{r})"
        ws[f"F{r}"].border = BORDER
    last_u = 19 + len(uenbs)
    _sum_row(ws, last_u + 1, 20, last_u, ["B", "C", "D", "E", "F"])

    ch = BarChart()
    ch.type = "col"
    ch.grouping = "stacked"
    ch.overlap = 100
    ch.title = "Umspannwerke nach Ampel je ÜNB (Einspeisung)"
    ch.y_axis.title = "Anzahl Umspannwerke"
    ch.add_data(Reference(ws, min_col=2, max_col=5, min_row=19, max_row=last_u), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=20, max_row=last_u))
    for s, colr in zip(ch.series, ("C53532", "F4B400", "6ABF30", "B9B6B1")):
        s.graphicalProperties.solidFill = colr
        s.graphicalProperties.line.solidFill = colr
    ch.roundedCorners = False
    ch.legend.position = "b"
    ch.legend.overlay = False
    ch.title.overlay = False
    ch.y_axis.title.overlay = False
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.dataLabels = None            # Zahlen stehen in der Tabelle links daneben
    ch.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))
    ch.height, ch.width = 7.4, 11.8
    ws.add_chart(ch, "G11")

    # 4 Top 15
    r4 = last_u + 4
    band(ws, "A", "J", r4, "Die 15 am stärksten ausgelasteten Umspannwerke (nur Netzknoten DE mit JAO-Daten)", "result")
    _tbl_header(ws, r4 + 1, ["Rang", "Umspannwerk", "ÜNB", "Höchste Auslastung", "Ampel Einspeisung", "Ampel Entnahme", "Begründung"])
    ws.merge_cells(f"G{r4 + 1}:J{r4 + 1}")
    for k in range(1, 16):
        r = r4 + 1 + k
        ws.cell(r, 1, k).alignment = Alignment(horizontal="center")
        ws.cell(r, 1).font = font(color=NAVY)
        pos = f'MATCH($A{r},{UW("AF")},0)'
        for c, src, fmt in (("B", "A", None), ("C", "B", None), ("D", "F", "0.0%"), ("E", "D", None), ("F", "E", None), ("G", "H", None)):
            cell = ws[f"{c}{r}"]
            cell.value = f'=IFERROR(INDEX({UW(src)},{pos}),"")'
            color_formula(cell)
            cell.border = BORDER
            if fmt:
                cell.number_format = fmt
        ws.merge_cells(f"G{r}:J{r}")
        ws[f"G{r}"].alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[r].height = 26
    add_ampel_cf(ws, f"E{r4 + 2}:F{r4 + 16}")
    add_util_cf(ws, f"D{r4 + 2}:D{r4 + 16}", f"D{r4 + 2}")

    # 5 Status
    r5 = r4 + 19
    band(ws, "A", "J", r5, "Status der Stromkreise und Trafos", "result")
    _tbl_header(ws, r5 + 1, ["Status", "Leitungen (Inland)", "Grenzkuppel-leitungen", "Trafos / Phasenschieber"])
    for i, name in enumerate(["Engpass", "hoch belastet", "unkritisch", "zu wenig Daten", "keine Daten"]):
        r = r5 + 2 + i
        ws.cell(r, 1, name).border = BORDER
        ws[f"B{r}"] = f'=COUNTIFS({LT("W")},$A{r},{LT("G")},"<>Grenzkuppelleitung")'
        ws[f"C{r}"] = f'=COUNTIFS({LT("W")},$A{r},{LT("G")},"Grenzkuppelleitung")'
        ws[f"D{r}"] = f'=COUNTIF({TR("T")},$A{r})'
        for c in "BCD":
            color_formula(ws[f"{c}{r}"])
            ws[f"{c}{r}"].border = BORDER
    _sum_row(ws, r5 + 7, r5 + 2, r5 + 6, ["B", "C", "D"])
    add_status_cf(ws, f"A{r5 + 2}:A{r5 + 6}")

    # 6 Netz je ÜNB
    r6 = r5 + 10
    band(ws, "A", "J", r6, "Netz je Übertragungsnetzbetreiber (Höchstspannung, Stromkreise im Inland)", "result")
    _tbl_header(ws, r6 + 1, ["ÜNB", "Stromkreise", "davon 380 kV", "davon 220 kV", "Länge (km)", "ohne Imax",
                             "Ø Smax je Stromkreis (MVA)", "Umspannwerke", "Trafos"])
    inland = f'{LT("G")},"<>Grenzkuppelleitung"'
    for i, name in enumerate(uenbs):
        r = r6 + 2 + i
        ws.cell(r, 1, name).font = font(color=NAVY)
        fs = {"B": f'=COUNTIFS({LT("C")},$A{r},{inland})',
              "C": f'=COUNTIFS({LT("C")},$A{r},{LT("F")},380,{inland})',
              "D": f'=COUNTIFS({LT("C")},$A{r},{LT("F")},220,{inland})',
              "E": f'=SUMIFS({LT("I")},{LT("C")},$A{r},{inland})',
              "F": f'=COUNTIFS({LT("C")},$A{r},{LT("K")},"fehlt",{inland})',
              "G": f'=IFERROR(AVERAGEIFS({LT("L")},{LT("C")},$A{r},{inland}),"")',
              "H": f'=COUNTIFS({UW("B")},$A{r},{UW("C")},"Netzknoten DE")',
              "I": f'=COUNTIF({TR("C")},$A{r})'}
        for c, fml in fs.items():
            ws[f"{c}{r}"] = fml
            color_formula(ws[f"{c}{r}"])
            ws[f"{c}{r}"].border = BORDER
            ws[f"{c}{r}"].number_format = "#,##0"
    lr = r6 + 1 + len(uenbs)
    _sum_row(ws, lr + 1, r6 + 2, lr, ["B", "C", "D", "E", "F", "H", "I"])
    ws[f"G{lr + 1}"] = f'=IFERROR(AVERAGEIFS({LT("L")},{inland}),"")'
    ws[f"G{lr + 1}"].number_format = "#,##0"
    ws[f"G{lr + 1}"].font = font(bold=True, color="000F47")

    # 7 Grenzkuppel
    r7 = lr + 4
    band(ws, "A", "J", r7, "Grenzkuppelleitungen je Nachbarland (thermische Obergrenze)", "result")
    _tbl_header(ws, r7 + 1, ["Nachbarland", "Stromkreise", "Σ Smax (MVA)", "Σ Pmax (MW)", "Höchste Auslastung", "Stromkreise im Engpass"])
    gk = f'{LT("G")},"Grenzkuppelleitung"'
    laender = sorted({v for v in data["sgm_leitungen"]["Nachbarland"] if v})
    for i, name in enumerate(laender):
        r = r7 + 2 + i
        ws.cell(r, 1, name).font = font(color=NAVY)
        fs = {"B": f'=COUNTIFS({LT("H")},$A{r},{gk})',
              "C": f'=SUMIFS({LT("L")},{LT("H")},$A{r},{gk})',
              "D": f'=SUMIFS({LT("M")},{LT("H")},$A{r},{gk})',
              "E": f'=IF(COUNTIFS({LT("H")},$A{r},{LT("Z")},1)=0,"",_xlfn.MAXIFS({LT("Q")},{LT("H")},$A{r},{LT("Z")},1))',
              "F": f'=COUNTIFS({LT("H")},$A{r},{LT("W")},"Engpass")'}
        for c, fml in fs.items():
            ws[f"{c}{r}"] = fml
            color_formula(ws[f"{c}{r}"])
            ws[f"{c}{r}"].border = BORDER
            ws[f"{c}{r}"].number_format = "0.0%" if c == "E" else "#,##0"
    lr7 = r7 + 1 + len(laender)
    _sum_row(ws, lr7 + 1, r7 + 2, lr7, ["B", "C", "D", "F"])
    ws[f"A{lr7 + 2}"] = ("Hinweis: Die Summe der thermischen Grenzen je Grenze ist eine Obergrenze – die tatsächliche Handelskapazität "
                         "(NTC / flussbasiert) ist deutlich geringer.")
    ws[f"A{lr7 + 2}"].font = font(size=8, italic=True, color="595959")
    ws[f"A{lr7 + 4}"] = ("Quelle: JAO Publication Tool (Final Computation), JAO Static Grid Model, Netzanschlussdaten der Übertragungsnetzbetreiber; "
                         "eigene Berechnung. Lage der Umspannwerke auf der Karte geschätzt.")
    ws[f"A{lr7 + 4}"].font = font(size=8, color="7B7974")
    ws.freeze_panes = "A4"
    add_map(wb, ws)
    _fit_width(ws)
    ws.sheet_view.zoomScale = 85


# --------------------------------------------------------------------------- Datenstand
def build_datenstand(wb, base_url=None):
    ws = wb["Datenstand"]
    set_widths(ws, {"A": 46, "B": 40, "C": 18, "D": 16, "E": 22, "F": 16, "G": 70})
    ws["A1"] = "Datenstand und Kontrollen"
    ws["A1"].font = font(bold=True, size=14, color=NAVY)
    ws["A2"] = "Aktualität der Rohdaten und Ergebnis der Modellkontrollen. Nach jeder Aktualisierung prüfen."
    ws["A2"].font = font(italic=True, size=9, color="595959")
    band(ws, "A", "G", 4, "A  Rohdatenblätter: Umfang und Aufbau", "input")
    _tbl_header(ws, 5, ["Blatt", "Inhalt", "Zeilen", "Platz im Modell", "Spalten wie erwartet?", "Platz reicht?", "Hinweis"], "input")
    spec = [("D_SGM_Leitungen", "Stromkreise (Netzmodell)", H_LINES, N_LINES, "Rechenblatt Leitungen"),
            ("D_SGM_Trafos", "Trafos / Phasenschieber", H_TRAFOS, N_TRAFOS, "Rechenblatt Trafos"),
            ("D_SGM_Umspannwerke", "Umspannwerke / Netzknoten", ["Umspannwerk"], N_UW, "Rechenblatt Umspannwerke"),
            ("D_Auslastung", "Auslastung je Element und Monat", H_AUSL, A_LAST - R0 + 1, "Formeln lesen bis Zeile %d" % A_LAST),
            ("D_Netzanschluss", "Freie Anschlusskapazität (ÜNB)", H_NA, N_NA, "darf leer sein"),
            ("D_Belegung", "Belegung / Pipeline (MaStR)", H_BEL, N_BEL, "darf leer sein (Phase 2)"),
            ("D_Koordinaten", "Lage der Umspannwerke (Karte)", H_KO, 996, "Formeln lesen bis Zeile 1000"),
            ("D_Meta", "Datenstand der Pipeline", ["Schluessel", "Wert"], 56, "")]
    n = len(spec)
    last_spec = 5 + n
    for i, (sheet, what, heads, cap, hint) in enumerate(spec):
        r = 6 + i
        ws.cell(r, 1, sheet).font = font(bold=True)
        ws.cell(r, 2, what)
        ws.cell(r, 3, f"=COUNTA({sheet}!$A$5:$A$40000)").number_format = "#,##0"
        ws.cell(r, 4, cap).number_format = "#,##0"
        ws.cell(r, 4).font = BLUE_FONT
        cond = ",".join(f'{sheet}!{col(j)}4="{h}"' for j, h in enumerate(heads, start=1))
        ws.cell(r, 5, f'=IF(AND({cond}),"OK","ABWEICHUNG")')
        ws.cell(r, 6, f'=IF(C{r}<=D{r},"OK","ZU KLEIN")')
        ws.cell(r, 7, hint).font = font(size=9, color="595959")
        for j in (3, 5, 6):
            color_formula(ws.cell(r, j))
            ws.cell(r, j).alignment = Alignment(horizontal="center")
        for j in range(1, 8):
            ws.cell(r, j).border = BORDER
    rc = last_spec + 2
    ws.cell(rc, 1, "D_Auslastung deckt ab:").font = font(bold=True)
    ws.cell(rc, 2, "=P_DatenVon").number_format = "mm/yyyy"
    ws.cell(rc, 3, "bis").alignment = Alignment(horizontal="center")
    ws.cell(rc, 4, "=P_DatenBis").number_format = "mm/yyyy"
    for c_ in (ws.cell(rc, 2), ws.cell(rc, 4)):
        c_.font = font(color="000F47", bold=True)
        c_.alignment = Alignment(horizontal="left")
    ws.cell(rc + 1, 1, "Datenquelle für 'Alle aktualisieren':").font = font(bold=True)
    src = base_url if base_url else "keine Verbindungen eingebaut"
    ws.cell(rc + 1, 2, src).font = BLUE_FONT
    ws.merge_cells(start_row=rc + 1, start_column=2, end_row=rc + 1, end_column=7)
    src_row = rc + 1

    rb = rc + 3
    band(ws, "A", "G", rb, "B  Datenstand aus der Pipeline (Blatt D_Meta)", "input")
    meta_rows = [("erstellt_utc", "Erstellt (UTC)"), ("git_commit", "Code-Stand (Commit)"), ("pipeline_version", "Pipeline-Version"),
                 ("daten_von_tag", "Auslastungsdaten von (Tag)"), ("daten_bis_tag", "Auslastungsdaten bis (Tag)"),
                 ("tage_mit_daten", "Tage mit Daten"), ("tage_vollstaendig", "davon vollständig (24 Stunden)"),
                 ("sgm_release", "Netzmodell (SGM-Release)"), ("schwelle_hoch", "Schwelle 'hoch' in der Pipeline"),
                 ("schwelle_ueberlast", "Schwelle 'Überlast' in der Pipeline"), ("zuordnung_alias", "JAO-Elemente per Alias zugeordnet"),
                 ("zuordnung_eic", "JAO-Elemente per EIC zugeordnet"), ("zuordnung_name", "JAO-Elemente per Name zugeordnet"),
                 ("zuordnung_offen", "JAO-Elemente ohne Zuordnung"), ("hinweis", "Hinweis")]
    for i, (key, label) in enumerate(meta_rows):
        r = rb + 1 + i
        ws.cell(r, 1, label)
        c = ws.cell(r, 2, f'=IFERROR({meta_ref(key)},"–")')
        color_formula(c)
        c.alignment = Alignment(horizontal="left", wrap_text=True)
        ws.cell(r, 1).border = BORDER
        c.border = BORDER
    hint_row = rb + len(meta_rows)
    ws.merge_cells(f"B{hint_row}:G{hint_row}")
    ws.row_dimensions[hint_row].height = 30

    r0 = rb + len(meta_rows) + 2
    band(ws, "A", "G", r0, "C  Kontrollen", "input")
    _tbl_header(ws, r0 + 1, ["Kontrolle", "Ergebnis", "Status", "Was bedeutet das? Was ist zu tun?", "", "", ""], "input")
    ws.merge_cells(f"D{r0 + 1}:G{r0 + 1}")
    LTr = lambda c: f"Leitungen!${c}${R0}:${c}${L_LAST}"
    KD = lambda c: f"Karte_Daten!${c}${R0}:${c}${U_LAST}"
    checks = [
        ("Spaltenreihenfolge der Rohdatenblätter", f'=COUNTIF(E6:E{last_spec},"OK")&" von {n} OK"', f'=IF(COUNTIF(E6:E{last_spec},"OK")={n},"OK","PRÜFEN")',
         "Excel liest die Rohdaten nach Spaltenposition. Weicht ein Kopf ab, hat sich die Lieferung geändert: Rohdatenblatt und Formeln prüfen."),
        ("Platz in den Rechenblättern", f'=COUNTIF(F6:F{last_spec},"OK")&" von {n} OK"', f'=IF(COUNTIF(F6:F{last_spec},"OK")={n},"OK","PRÜFEN")',
         "Die Rechenblätter haben eine feste Zahl vorbereiteter Zeilen. 'ZU KLEIN' heißt: Formeln im Rechenblatt nach unten kopieren."),
        ("Datenquelle für 'Alle aktualisieren' eingerichtet", f"=B{src_row}",
         f'=IF(B{{r}}="keine Verbindungen eingebaut","INFO",IF(ISNUMBER(SEARCH("{conn.PLACEHOLDER}",B{{r}})),"PRÜFEN","OK"))',
         "Steht hier 'ORG/REPO', ist die Adresse noch ein Platzhalter: Datei über die GitHub-Aktion 'Excel-Modell bauen' erzeugen oder mit --base-url bauen. "
         "Das Repository muss öffentlich lesbar sein (sonst Power Query mit Token, Blatt Automatisierung)."),
        ("Auswertungsfenster enthält Daten", f'=IF(SUM({LTr("O")})>0,"ja","nein")', '=IF(B{r}="ja","OK","PRÜFEN")',
         "'nein' heißt: Im gewählten Zeitraum (Parameter) gibt es keine Auslastungsdaten. Fenster oder D_Auslastung prüfen."),
        ("Stromkreise mit JAO-Daten (Anteil)", f'=IF(COUNTIF({LTr("A")},"?*")=0,"",SUM({LTr("Z")})/COUNTIF({LTr("A")},"?*"))', '="INFO"',
         "JAO veröffentlicht nur handelsrelevante Elemente, der Rest steht auf 'keine Daten'. Über längere Zeiträume steigt der Anteil."),
        ("JAO-Elemente ohne Zuordnung (letzter Monat im Fenster)", '=COUNTIFS(D_Auslastung!$C$5:$C$20004,"nicht zugeordnet",D_Auslastung!$A$5:$A$20004,P_Bis)',
         '=IF(B{r}=0,"OK","HINWEIS")',
         "Diese JAO-Elemente finden im Netzmodell keinen Partner (Name oder EIC weichen ab) und fehlen deshalb in den Ampeln. "
         "Vorschläge: data/manual/alias_vorschlaege.csv im Repository, bestätigte Zuordnungen in alias_jao_sgm.csv."),
        ("Stromkreise ohne Spannung (Smax bleibt leer)", f'=COUNTIFS({LTr("A")},"?*",{LTr("F")},"")', '=IF(B{r}=0,"OK","HINWEIS")',
         "Lücke in den Quelldaten. Die Kapazität bleibt leer statt 0. Bei der nächsten SGM-Release prüfen."),
        ("Stromkreise ohne Imax (Quelle 'fehlt')", f'=COUNTIF({LTr("K")},"fehlt")', '=IF(B{r}=0,"OK","HINWEIS")',
         "Ohne Grenzstrom keine Kapazität. Diese Stromkreise zählen nicht in Summen."),
        ("Mögliche Dubletten (gleiche EIC und gleicher Name)", f'=COUNTIF({LTr("Y")},"prüfen")', '=IF(B{r}=0,"OK","HINWEIS")',
         "Können echte Doppelungen sein (z. B. Beznau–Tiengen als Leitung und Grenzkuppel) oder zwei Pole/Stränge (NordLink). "
         "Fachlich prüfen, nicht automatisch löschen."),
        ("Umspannwerke ohne Lage auf der Karte (Netzknoten DE)",
         f'=COUNTIF({KD("B")},"Netzknoten DE")-SUMPRODUCT(({KD("B")}="Netzknoten DE")*ISNUMBER({KD("G")}))', '=IF(B{r}=0,"OK","HINWEIS")',
         "Für diese Umspannwerke gibt es keine Koordinate: Der Name verbindet vermutlich mehrere Standorte (gleicher Name bei mehreren Betreibern) "
         "oder der Knoten hängt nicht am zusammenhängenden Netz. Liste: data/manual/koordinaten_pruefbericht.csv. Ein gleichnamiger Knoten wird im Modell "
         "als EIN Umspannwerk gezählt: auch die Ampeln dieser Namen sind mit Vorsicht zu lesen."),
        ("Umspannwerke mit nur geschätzter Lage", '=COUNTIF(D_Koordinaten!$D$5:$D$1000,"geschätzt*")', '="INFO"',
         "Die Lage ist aus Leitungslängen und Ankerpunkten berechnet (Median-Abweichung rund 16 km, einzelne Knoten über 50 km). Exakte Werte in "
         "data/manual/koordinaten_manuell.csv überschreiben die Schätzung."),
        ("Namen in D_Netzanschluss ohne Treffer im Modell",
         f'=SUMPRODUCT((D_Netzanschluss!$A$5:$A${NA_LAST}<>"")*(COUNTIF(Umspannwerke!$A$5:$A${U_LAST},D_Netzanschluss!$A$5:$A${NA_LAST})=0))',
         '=IF(B{r}=0,"OK","PRÜFEN")', "Schreibweise der ÜNB-Karte stimmt nicht mit dem Blatt Umspannwerke überein: Name anpassen, sonst bleibt die Zeile wirkungslos."),
        ("Namen in D_Belegung ohne Treffer im Modell",
         f'=SUMPRODUCT((D_Belegung!$A$5:$A${BEL_LAST}<>"")*(COUNTIF(Umspannwerke!$A$5:$A${U_LAST},D_Belegung!$A$5:$A${BEL_LAST})=0))',
         '=IF(B{r}=0,"OK","PRÜFEN")', "Wie oben, für die Belegungsdaten."),
        ("Schwelle 'hoch': Pipeline = Excel", '=IF(P_PipelineHoch="","–",P_PipelineHoch)',
         '=IF(P_PipelineHoch="","INFO",IF(P_PipelineHoch=P_SchwelleHoch,"OK","HINWEIS"))',
         "Nur die Spalte 'Stunden ab hoch' nutzt die Pipeline-Schwelle. Der Status in Excel nutzt die Schwelle im Blatt Parameter."),
        ("Hinweis der Pipeline", f'=IFERROR(IF({meta_ref("hinweis")}="","keiner",{meta_ref("hinweis")}),"keiner")',
         '=IF(B{r}="keiner","OK","HINWEIS")', "Zum Beispiel: Es liegen nur Beispieldaten vor."),
    ]
    for i, (label, result, status, expl) in enumerate(checks):
        r = r0 + 2 + i
        ws.cell(r, 1, label).alignment = WRAP_TOP
        rc_ = ws.cell(r, 2, result)
        color_formula(rc_)
        rc_.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        if "Anteil" in label:
            rc_.number_format = "0%"
        sc = ws.cell(r, 3, status.replace("{r}", str(r)))
        sc.alignment = Alignment(horizontal="center", vertical="top")
        sc.font = font(bold=True)
        ws.merge_cells(f"D{r}:G{r}")
        ws.cell(r, 4, expl).font = font(size=9)
        ws.cell(r, 4).alignment = WRAP_TOP
        ws.row_dimensions[r].height = max(30, row_height(expl, 120, line_pt=12), 44 if label.startswith(("Hinweis der", "Datenquelle")) else 0)
        for j in range(1, 8):
            ws.cell(r, j).border = BORDER
    last = r0 + 1 + len(checks)
    for text, bg in (("OK", "DDF0CC"), ("PRÜFEN", "F6D3D2"), ("ABWEICHUNG", "F6D3D2"), ("ZU KLEIN", "F6D3D2"),
                     ("HINWEIS", "FBEDBF"), ("INFO", "CEECFF")):
        ws.conditional_formatting.add(f"C{r0 + 2}:C{last}", CellIsRule(operator="equal", formula=[f'"{text}"'], fill=fill(bg)))
        ws.conditional_formatting.add(f"E6:F{last_spec}", CellIsRule(operator="equal", formula=[f'"{text}"'], fill=fill(bg)))
    _fit_width(ws)


# --------------------------------------------------------------------------- Textblätter
def _section(ws, r, text, c1="B", c2="G", kind="info"):
    band(ws, c1, c2, r, text, kind)
    ws.row_dimensions[r].height = 22


def _para(ws, r, text, c1="B", c2="G", width_chars=150, bold=False, size=10, color="000000", bg=None, italic=False, min_h=15):
    ws.merge_cells(f"{c1}{r}:{c2}{r}")
    c = ws[f"{c1}{r}"]
    c.value = text
    c.font = font(bold=bold, size=size, color=color, italic=italic)
    c.alignment = WRAP_TOP
    if bg:
        for cell in ws[f"{c1}{r}:{c2}{r}"][0]:
            cell.fill = fill(bg)
    est = text if not str(text).startswith("=") else "x" * 130
    ws.row_dimensions[r].height = row_height(est, width_chars, minimum=min_h)


def _link(cell, sheet):
    cell.value = sheet                      # sonst zeigt Excel den Link-Zielort statt des Blattnamens
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"'{sheet}'!A1", display=sheet)
    cell.font = font(color="0563C1", underline="single", bold=True)


def build_lies_mich(wb):
    ws = wb["Lies_mich"]
    set_widths(ws, {"A": 2, "B": 25, "C": 25, "D": 25, "E": 25, "F": 25, "G": 25})
    ws.sheet_view.showGridLines = False
    ws["B1"] = C.TITLE
    ws["B1"].font = font(bold=True, size=17, color=NAVY)
    ws.row_dimensions[1].height = 28
    _para(ws, 2, C.SUBTITLE, italic=True, color="595959")
    for lc, label, vc, formula, fmt in (("B", "Auswertungszeitraum von", "C", "=P_Von", "mm/yyyy"), ("D", "bis", "E", "=P_Bis", "mm/yyyy"),
                                        ("F", "Pipeline-Stand (UTC)", "G", f'=IFERROR({meta_ref("erstellt_utc")},"–")', None)):
        ws[f"{lc}3"] = label
        ws[f"{lc}3"].font = font(size=9, color="595959")
        ws[f"{vc}3"] = formula
        ws[f"{vc}3"].font = font(bold=True, color="000F47")
        ws[f"{vc}3"].alignment = Alignment(horizontal="left")
        if fmt:
            ws[f"{vc}3"].number_format = fmt
    _para(ws, 4, f'=IFERROR(IF({meta_ref("hinweis")}="","","Hinweis: "&{meta_ref("hinweis")}),"")', bold=True, color="C53532")

    r = 6
    _section(ws, r, "1  Worum geht es?")
    for t in C.INTRO:
        r += 1
        _para(ws, r, "•  " + t)
    r += 2
    _section(ws, r, "2  Lesehilfe: die Ampel")
    r += 1
    for rng, text in (("B", "Farbe"), ("C", "Bedeutung"), ("D", "Regel für das Netz (Auslastung)"), ("F", "Regel für den Anschluss (freie Kapazität)")):
        header(ws[f"{rng}{r}"], "info")
        ws[f"{rng}{r}"].value = text
    ws.merge_cells(f"D{r}:E{r}")
    ws.merge_cells(f"F{r}:G{r}")
    for rng in ("E", "G"):
        header(ws[f"{rng}{r}"], "info")
    for name, meaning, netz, anschluss in C.AMPEL_LEGENDE:
        r += 1
        bg, fg = AMPEL[name]
        ws[f"B{r}"] = name
        ws[f"B{r}"].fill = fill(bg)
        ws[f"B{r}"].font = font(bold=True, color=fg, size=11)
        ws[f"B{r}"].alignment = CENTER
        ws[f"C{r}"] = meaning
        ws[f"C{r}"].font = font(bold=True)
        ws[f"C{r}"].alignment = WRAP_TOP
        for c1, c2, val in (("D", "E", netz), ("F", "G", anschluss)):
            ws.merge_cells(f"{c1}{r}:{c2}{r}")
            ws[f"{c1}{r}"] = val
            ws[f"{c1}{r}"].alignment = WRAP_TOP
            ws[f"{c1}{r}"].font = font(size=9)
            color_formula(ws[f"{c1}{r}"])
        for j in range(2, 8):
            ws.cell(r, j).border = BORDER
        ws.row_dimensions[r].height = 40
    r += 1
    _para(ws, r, C.AMPEL_HINWEIS, bold=True, color=NAVY, bg=L_BLUE)

    r += 2
    _section(ws, r, "3  Landkarte der Blätter (Namen anklicken)")
    r += 1
    for rng, text in (("B", "Blatt"), ("C", "Farbe des Reiters"), ("D", "Rolle"), ("E", "Beantwortet die Frage")):
        header(ws[f"{rng}{r}"], "info")
        ws[f"{rng}{r}"].value = text
    ws.merge_cells(f"E{r}:G{r}")
    for rng in ("F", "G"):
        header(ws[f"{rng}{r}"], "info")
    tab_fill = {"hellblau": "CEECFF", "navy": "000F47", "blau": "82BAFF", "leuchtblau": "0B4BFF", "grau": GREY}
    for sheet, farbe, rolle, frage in C.BLATT_KARTE:
        r += 1
        _link(ws[f"B{r}"], sheet)
        ws[f"C{r}"] = farbe
        ws[f"C{r}"].fill = fill(tab_fill[farbe])
        ws[f"C{r}"].font = font(color="FFFFFF" if farbe in ("navy", "leuchtblau", "grau") else NAVY, bold=True)
        ws[f"C{r}"].alignment = CENTER
        ws[f"D{r}"] = rolle
        ws.merge_cells(f"E{r}:G{r}")
        ws[f"E{r}"] = frage
        for c in (f"D{r}", f"E{r}"):
            ws[c].alignment = WRAP_TOP
        for j in range(2, 8):
            ws.cell(r, j).border = BORDER

    r += 2
    _section(ws, r, "4  So hängt alles zusammen (von links nach rechts)")
    r += 1
    for j, (label, text, bg) in enumerate(C.FLOW):
        c = ws.cell(r, 2 + j, f"{label}\n{text}")
        c.fill = fill(bg)
        c.font = font(size=9)
        c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="center")
        c.border = BORDER
    ws.row_dimensions[r].height = 78
    r += 1
    _para(ws, r, "Ein Rechenbeispiel mit echten Zahlen aus diesem Modell steht im Blatt Modelllogik.", italic=True, color="595959", size=9)

    r += 2
    _section(ws, r, "5  Farbcode")
    styles = {"blue": dict(f=font(color="0B4BFF", bold=True), bg=None), "black": dict(f=font(bold=True), bg=None),
              "green": dict(f=font(color="000F47", bold=True), bg=None), "yellow": dict(f=font(color="0B4BFF", bold=True), bg=L_GOLD),
              "grey": dict(f=font(color="FFFFFF", bold=True), bg=GREY), "navy": dict(f=font(color="FFFFFF", bold=True), bg=NAVY),
              "green_h": dict(f=font(color="FFFFFF", bold=True), bg=GREEN_D)}
    for label, desc, key in C.FARBCODE:
        r += 1
        ws[f"B{r}"] = label
        ws[f"B{r}"].font = styles[key]["f"]
        if styles[key]["bg"]:
            ws[f"B{r}"].fill = fill(styles[key]["bg"])
        ws[f"B{r}"].border = BORDER
        _para(ws, r, desc, c1="C", c2="G", width_chars=125)

    r += 2
    _section(ws, r, "6  Daten aktualisieren")
    for title, text in C.AKTUALISIEREN:
        r += 1
        ws[f"B{r}"] = title
        ws[f"B{r}"].font = font(bold=True)
        ws[f"B{r}"].alignment = WRAP_TOP
        _para(ws, r, text, c1="C", c2="G", width_chars=125)
        ws.row_dimensions[r].height = max(ws.row_dimensions[r].height, 30)

    r += 2
    _section(ws, r, "7  Was das Modell (noch) nicht kann", kind="result")
    for t in C.GRENZEN:
        r += 1
        _para(ws, r, "•  " + t)

    r += 2
    _section(ws, r, "8  Neu in Version 3 gegenüber Version 2")
    for t in C.NEU_IN_V3:
        r += 1
        _para(ws, r, "•  " + t)
    _fit_width(ws, landscape=False)


def _pick_examples(data):
    """Beispiel für das Rechenblatt: ein Stromkreis im Engpass (falls vorhanden) und sein Umspannwerk."""
    from ..normalize import element_key
    mon = data["auslastung_monatlich"]
    if mon.empty:
        ln = data["sgm_leitungen"].iloc[0]
        return ln["NE_name"], ln["Umspannwerk_1"]
    agg = mon.groupby("Schluessel").agg(h=("Stunden", lambda s: s.astype(float).sum()), ue=("Stunden_Ueberlast", lambda s: s.astype(float).sum()),
                                        mx=("Auslastung_Max", lambda s: s.astype(float).max()))
    best, best_score = None, -1
    for _, ln in data["sgm_leitungen"].iterrows():
        if ln["Typ"] == "Grenzkuppelleitung":
            continue
        k = element_key(ln["EIC_Code"], ln["NE_name"])
        if k in agg.index:
            a = agg.loc[k]
            score = (1 if a["ue"] / a["h"] >= 0.05 else 0) * 10 + a["mx"]
            if score > best_score:
                best, best_score = ln, score
    ln = best if best is not None else data["sgm_leitungen"].iloc[0]
    return ln["NE_name"], ln["Umspannwerk_1"]


def build_modelllogik(wb, data):
    ws = wb["Modelllogik"]
    set_widths(ws, {"A": 2, "B": 34, "C": 62, "D": 62})
    ws.sheet_view.showGridLines = False
    ws["B1"] = "Modelllogik – so entsteht eine Ampel"
    ws["B1"].font = font(bold=True, size=17, color=NAVY)
    ws.row_dimensions[1].height = 28
    _para(ws, 2, "Sechs Schritte von den Rohdaten zur Ampel, mit zwei Rechenbeispielen auf Basis der Daten dieses Modells.", c2="D", width_chars=158, italic=True, color="595959")
    r = 4
    _section(ws, r, "Die 6 Schritte im Überblick", "B", "D")
    r += 1
    for rng, t in (("B", "Schritt"), ("C", "Was passiert?"), ("D", "Wo im Modell")):
        header(ws[f"{rng}{r}"], "info")
        ws[f"{rng}{r}"].value = t
    for title, blatt, text in C.SCHRITTE:
        r += 1
        ws[f"B{r}"] = title
        ws[f"B{r}"].font = font(bold=True)
        ws[f"C{r}"] = text
        ws[f"C{r}"].font = font(size=9)
        color_formula(ws[f"C{r}"])
        ws[f"D{r}"] = blatt
        for c in "BCD":
            ws[f"{c}{r}"].alignment = WRAP_TOP
            ws[f"{c}{r}"].border = BORDER
        ws.row_dimensions[r].height = max(30, row_height(text if not text.startswith("=") else "x" * 150, 62, line_pt=12))

    ex_line, ex_uw = _pick_examples(data)
    # ---- Beispiel 1: Stromkreis
    r += 2
    _section(ws, r, "Rechenbeispiel 1: ein Stromkreis (live)", "B", "D", kind="calc")
    r += 1
    ws[f"B{r}"] = "Stromkreis wählen ▶"
    ws[f"B{r}"].font = font(bold=True)
    ws[f"C{r}"] = ex_line
    ws[f"C{r}"].fill = fill(L_GOLD)
    ws[f"C{r}"].font = font(color="0B4BFF", bold=True)
    ws[f"C{r}"].border = BORDER
    ws[f"D{r}"] = "Wähle in der gelben Zelle einen anderen Stromkreis aus der Liste. Alles darunter rechnet neu."
    ws[f"D{r}"].font = font(size=9, italic=True, color="595959")
    sel1 = f"$C${r}"
    dv1 = DataValidation(type="list", formula1=f"=Leitungen!$A${R0}:$A${L_LAST}", allow_blank=False)
    ws.add_data_validation(dv1)
    dv1.add(ws[f"C{r}"])
    r += 1
    ws[f"B{r}"] = "Position in der Liste (Hilfswert)"
    ws[f"B{r}"].font = font(size=9, color="595959")
    ws[f"C{r}"] = f'=IFERROR(MATCH({sel1},Leitungen!$A${R0}:$A${L_LAST},0),"nicht gefunden")'
    color_formula(ws[f"C{r}"])
    ws[f"C{r}"].alignment = Alignment(horizontal="left")
    pos1 = f"$C${r}"
    r += 1
    for rng, t in (("B", "Schritt"), ("C", "Ergebnis (live)"), ("D", "So wird gerechnet")):
        header(ws[f"{rng}{r}"], "calc")
        ws[f"{rng}{r}"].value = t
    ix = lambda c: f"INDEX(Leitungen!${c}${R0}:${c}${L_LAST},{pos1})"
    steps = [
        ("1  Rohdaten (D_SGM_Leitungen)",
         f'=IFERROR({ix("C")}&" · "&{ix("F")}&" kV · "&{ix("G")}&IF({ix("H")}<>""," · Nachbarland "&{ix("H")},"")&" · "&IF({ix("I")}="","Länge unbekannt",ROUND({ix("I")},1)&" km"),"")',
         "Name, Netzbetreiber, Spannung und Typ des Stromkreises stehen im Rohdatenblatt."),
        ("2  Grenzstrom Imax",
         f'=IFERROR(IF({ix("J")}="","kein Grenzstrom vorhanden",{ix("J")}&" A  (Quelle: "&{ix("K")}&")"),"")',
         "Hat der Stromkreis einen festen Grenzstrom (fix), gilt er. Sonst gilt bei saisonalen Werten der kleinste (Annahme MIN) bzw. größte (MAX), bei DLR-Werten der DLR-Mindestwert."),
        ("3  Kapazität Smax",
         f'=IFERROR(IF({ix("L")}="","nicht berechenbar (Spannung oder Imax fehlt)","√3 × "&{ix("F")}&" kV × "&{ix("J")}&" A ÷ 1.000 = "&ROUND({ix("L")},0)&" MVA  (≈ "&ROUND({ix("M")},0)&" MW)"),"")',
         "Smax = √3 × Spannung × Imax ÷ 1.000. Mit cos φ = 1 entspricht die Zahl in MW der Zahl in MVA (Annahme im Blatt Parameter)."),
        ("4  Auslastung laut JAO",
         f'=IFERROR(IF({ix("O")}=0,"keine JAO-Daten im Auswertungsfenster",{ix("O")}&" Stunden · Ø "&ROUND({ix("P")}*100,1)&" % · höchste "&ROUND({ix("Q")}*100,1)&" % · Stunden ab 100 %: "&{ix("S")}),"")',
         f'="Verknüpfung über den Schlüssel: "&IFERROR({ix("N")},"")&". Es werden alle Monate im Fenster aus Blatt Parameter addiert."'),
        ("5  Status", f'=IFERROR({ix("W")},"")',
         '="Engpass, wenn in ≥ "&TEXT(P_SchwelleEngpass,"0%")&" der Stunden Überlast. Hoch belastet ab "&TEXT(P_SchwelleHoch,"0%")&" höchster Auslastung. Sonst unkritisch."'),
        ("6  Rest-Kapazität",
         f'=IFERROR(IF({ix("V")}="","–","ca. "&ROUND({ix("V")},0)&" MW"),"")',
         "Rechnung: Fmax laut JAO × (1 − höchste Auslastung). Eine Näherung dafür, wie viel Leistung im ungünstigsten Fall noch Platz hätte. Sie ersetzt keine Netzanschlussprüfung."),
        ("Ampel", f'=IFERROR({ix("X")},"")', "Rot = Engpass, Gelb = hoch belastet, Grün = unkritisch, Grau = keine oder zu wenig Daten."),
    ]
    ampel_cell = None
    for title, live, how in steps:
        r += 1
        ws[f"B{r}"] = title
        ws[f"B{r}"].font = font(bold=True)
        ws[f"C{r}"] = live
        color_formula(ws[f"C{r}"])
        ws[f"C{r}"].font = font(bold=True, color=ws[f"C{r}"].font.color.rgb[-6:])
        ws[f"D{r}"] = how
        ws[f"D{r}"].font = font(size=9)
        color_formula(ws[f"D{r}"])
        ws[f"D{r}"].font = font(size=9, color=ws[f"D{r}"].font.color.rgb[-6:] if ws[f"D{r}"].font.color is not None and str(ws[f"D{r}"].value).startswith("=") else "000000")
        for c in "BCD":
            ws[f"{c}{r}"].alignment = WRAP_TOP
            ws[f"{c}{r}"].border = BORDER
        ws.row_dimensions[r].height = max(30, row_height(how if not how.startswith("=") else "x" * 120, 62, line_pt=12))
        if title == "Ampel":
            ampel_cell = f"C{r}"
        if title == "5  Status":
            add_status_cf(ws, f"C{r}")
    add_ampel_cf(ws, ampel_cell)

    # ---- Beispiel 2: Umspannwerk
    r += 2
    _section(ws, r, "Rechenbeispiel 2: ein Umspannwerk (live)", "B", "D", kind="result")
    r += 1
    ws[f"B{r}"] = "Umspannwerk wählen ▶"
    ws[f"B{r}"].font = font(bold=True)
    ws[f"C{r}"] = ex_uw
    ws[f"C{r}"].fill = fill(L_GOLD)
    ws[f"C{r}"].font = font(color="0B4BFF", bold=True)
    ws[f"C{r}"].border = BORDER
    ws[f"D{r}"] = "Wähle ein anderes Umspannwerk. Die Zeile stammt aus dem Blatt Umspannwerke."
    ws[f"D{r}"].font = font(size=9, italic=True, color="595959")
    sel2 = f"$C${r}"
    dv2 = DataValidation(type="list", formula1=f"=Umspannwerke!$A${R0}:$A${U_LAST}", allow_blank=False)
    ws.add_data_validation(dv2)
    dv2.add(ws[f"C{r}"])
    r += 1
    ws[f"B{r}"] = "Position in der Liste (Hilfswert)"
    ws[f"B{r}"].font = font(size=9, color="595959")
    ws[f"C{r}"] = f'=IFERROR(MATCH({sel2},Umspannwerke!$A${R0}:$A${U_LAST},0),"nicht gefunden")'
    color_formula(ws[f"C{r}"])
    ws[f"C{r}"].alignment = Alignment(horizontal="left")
    pos2 = f"$C${r}"
    r += 1
    for rng, t in (("B", "Kennzahl"), ("C", "Ergebnis (live)"), ("D", "Bedeutung")):
        header(ws[f"{rng}{r}"], "result")
        ws[f"{rng}{r}"].value = t
    ux = lambda c: f"INDEX(Umspannwerke!${c}${R0}:${c}${U_LAST},{pos2})"
    ust = [
        ("Netzbetreiber · Knotentyp", f'=IFERROR({ux("B")}&" · "&{ux("C")},"")', "'Netzknoten DE' = echtes deutsches Umspannwerk, 'Abzweig (Y-Knoten)' = Verzweigungspunkt einer Leitung."),
        ("Angeschlossene Elemente", f'=IFERROR({ux("I")}&" Leitungen, "&{ux("J")}&" Grenzkuppelleitungen, "&{ux("K")}&" Trafos","")', "Alle Stromkreise und Trafos, die an diesem Knoten enden."),
        ("Davon mit JAO-Daten", f'=IFERROR({ux("L")}&"  (Engpass: "&{ux("M")}&", hoch belastet: "&{ux("N")}&")","")', "Nur diese Elemente fließen in die Netz-Ampel ein."),
        ("Höchste Auslastung", f'=IFERROR({ux("F")},"")', "Der ungünstigste Wert aller bewerteten Elemente am Knoten."),
        ("Ampel Netz", f'=IFERROR({ux("P")},"")', "Rot, sobald ein Element im Engpass ist. Gelb, sobald eines hoch belastet ist. Grün, wenn alle bewerteten Elemente unkritisch sind."),
        ("Angabe der ÜNB zur freien Kapazität",
         f'=IFERROR(IF({ux("AG")}="","keine Angabe in D_Netzanschluss","Einspeisung "&{ux("T")}&"–"&{ux("U")}&" MW, Entnahme "&{ux("V")}&"–"&{ux("W")}&" MW"),"")',
         "Bandbreiten laut Netzanschlusskarten der Übertragungsnetzbetreiber."),
        ("AMPEL Einspeisung", f'=IFERROR({ux("D")},"")', "Der ungünstigere Wert aus Netz-Ampel und Anschluss-Ampel Einspeisung."),
        ("AMPEL Entnahme", f'=IFERROR({ux("E")},"")', "Der ungünstigere Wert aus Netz-Ampel und Anschluss-Ampel Entnahme."),
        ("Datengrundlage", f'=IFERROR({ux("G")},"")', "Worauf die Ampel beruht: Netz, Anschluss, beides oder keine Daten."),
        ("Begründung", f'=IFERROR({ux("H")},"")', "Die Begründung steht im Blatt Umspannwerke in Spalte H."),
    ]
    for title, live, how in ust:
        r += 1
        ws[f"B{r}"] = title
        ws[f"B{r}"].font = font(bold=True)
        ws[f"C{r}"] = live
        color_formula(ws[f"C{r}"])
        ws[f"C{r}"].font = font(bold=True, color="000F47")
        ws[f"D{r}"] = how
        ws[f"D{r}"].font = font(size=9)
        for c in "BCD":
            ws[f"{c}{r}"].alignment = WRAP_TOP
            ws[f"{c}{r}"].border = BORDER
        ws.row_dimensions[r].height = max(30, row_height(how, 62, line_pt=12))
        if title == "Höchste Auslastung":
            ws[f"C{r}"].number_format = "0.0%"
            ws[f"C{r}"].alignment = Alignment(horizontal="left", vertical="top")
        if title in ("Ampel Netz", "AMPEL Einspeisung", "AMPEL Entnahme"):
            add_ampel_cf(ws, f"C{r}")
    r += 2
    _section(ws, r, "Warum bestimmt der schwächste Punkt die Farbe?", "B", "D")
    r += 1
    _para(ws, r, "Ein Umspannwerk ist nur so aufnahmefähig wie sein engster Punkt: Ist eine angeschlossene Leitung am Limit, hilft es nichts, "
                 "dass die anderen frei sind. Deshalb bildet das Modell nie einen Durchschnitt, sondern nimmt den ungünstigsten Wert "
                 "(Rot vor Gelb vor Grün). Grau erscheint nur, wenn gar keine Information vorliegt. Grau bedeutet nicht 'in Ordnung'.",
          c2="D", width_chars=158)
    _fit_width(ws, landscape=False)


def build_glossar(wb, data):
    ws = wb["Glossar"]
    rows = data["glossar"]
    ws["A1"] = "Glossar – Begriffe einfach erklärt"
    ws["A1"].font = font(bold=True, size=14, color=NAVY)
    ws["A2"] = "Für alle, die den Stromsektor nicht kennen. Die Liste lässt sich nach Kategorie filtern (Pfeil in der Kopfzeile)."
    ws["A2"].font = font(italic=True, size=9, color="595959")
    set_widths(ws, {"A": 34, "B": 12, "C": 84, "D": 60, "E": 30})
    for j, h in enumerate(rows[0], start=1):
        header(ws.cell(4, j, h), "info")
    for i, row in enumerate(rows[1:], start=R0):
        for j, v in enumerate(row, start=1):
            c = ws.cell(i, j, v)
            c.alignment = WRAP_TOP
            c.border = BORDER
            c.font = font(bold=(j == 1), size=10 if j < 3 else 9)
        ws.row_dimensions[i].height = max(28, row_height(row[2], 84, line_pt=12), row_height(row[3], 60, line_pt=12))
    ws.freeze_panes = "B5"
    ws.auto_filter.ref = f"A4:E{R0 + len(rows) - 2}"
    _fit_width(ws)


def build_quellen(wb, data):
    ws = wb["Quellen"]
    rows = data["quellen"]
    ws["A1"] = "Quellenregister – woher kommen die Daten?"
    ws["A1"].font = font(bold=True, size=14, color=NAVY)
    ws["A2"] = ("Links geprüft am 18.11.2026. Im Auftakt waren die meisten Links nur Startseiten: hier stehen die Seiten der Datensätze. "
                "Automatisierung: ●●● Schnittstelle oder Download · ●●○ teilweise manuell · ●○○ Webkarte oder PDF. "
                "Rolle: Kern / Einfluss / 2. Schritt wie im Auftakt (Folie Scope).")
    ws["A2"].font = font(italic=True, size=9, color="595959")
    widths = [6, 38, 22, 46, 22, 24, 24, 10, 10, 12, 26, 28, 12, 22, 28, 24, 30, 70]
    for j, w in enumerate(widths, start=1):
        ws.column_dimensions[col(j)].width = w
    for j, h in enumerate(rows[0], start=1):
        header(ws.cell(4, j, h), "info")
    ws.row_dimensions[4].height = 34
    for i, row in enumerate(rows[1:], start=R0):
        for j, v in enumerate(row, start=1):
            c = ws.cell(i, j, v)
            c.alignment = WRAP_TOP
            c.border = BORDER
            c.font = font(size=9, bold=(j == 2))
            if j == 4 and v.startswith("http"):
                c.hyperlink = v.split()[0]
                c.font = font(size=9, color="0563C1", underline="single")
        status = row[13]
        bg = {"aktiv": "DDF0CC", "vorbereitet": "FBEDBF"}.get(status.split(" (")[0], None)
        if bg:
            ws.cell(i, 14).fill = fill(bg)
        h = max(row_height(v, w, line_pt=12) for v, w in zip(row, widths))
        ws.row_dimensions[i].height = max(28, h)
    ws.freeze_panes = "C5"
    ws.auto_filter.ref = f"A4:{col(len(widths))}{R0 + len(rows) - 2}"
    _fit_width(ws)


def build_automatisierung(wb):
    from .content import M_QUERIES, m_code
    ws = wb["Automatisierung"]
    set_widths(ws, {"A": 2, "B": 44, "C": 36, "D": 46, "E": 64})
    ws.sheet_view.showGridLines = False
    ws["B1"] = "Automatisierung – so bleibt das Modell aktuell"
    ws["B1"].font = font(bold=True, size=17, color=NAVY)
    ws.row_dimensions[1].height = 28
    _para(ws, 2, C.AUTO_SATZ, c2="E", width_chars=186, bold=True, color=NAVY, bg=L_BLUE)
    r = 4
    _section(ws, r, "Wer holt welche Daten, wie oft?", "B", "E")
    r += 1
    for rng, t in (("B", "Quelle"), ("C", "Abruf"), ("D", "Auswertung im Modell"), ("E", "Warum so?")):
        header(ws[f"{rng}{r}"], "info")
        ws[f"{rng}{r}"].value = t
    for a, b, c_, d in C.AUTO_RHYTHMUS:
        r += 1
        for j, v in enumerate((a, b, c_, d)):
            cell = ws.cell(r, 2 + j, v)
            cell.alignment = WRAP_TOP
            cell.border = BORDER
            cell.font = font(size=9, bold=(j == 0))
        ws.row_dimensions[r].height = max(28, row_height(d, 64, line_pt=12))
    r += 2
    _section(ws, r, "Der Weg der Daten", "B", "E")
    r += 1
    flow_cols = ["B", "C", "D", "E"]
    flow = [("GitHub-Aktion (täglich)", "Python-Skript holt JAO-Daten, bereinigt Namen, verdichtet und speichert in data/interim", "CEECFF"),
            ("Repository (data/processed)", "kleine CSV-Dateien: Auslastung je Monat, Netzstammdaten, Datenstand", "CEECFF"),
            ("Excel: Daten → Alle aktualisieren", "die eingebauten Datenverbindungen laden die CSV in die grauen D_-Blätter", "D1CEC9"),
            ("Modell rechnet", "Leitungen/Trafos → Umspannwerke → Dashboard, ohne Zutun", "EBE7E2")]
    for cl, (t, d, bg) in zip(flow_cols, flow):
        c = ws[f"{cl}{r}"]
        c.value = f"{t}\n{d}"
        c.fill = fill(bg)
        c.font = font(size=9)
        c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="center")
        c.border = BORDER
    ws.row_dimensions[r].height = 62
    r += 2
    _section(ws, r, "Aktuelle Daten laden (Verbindungen sind eingebaut)", "B", "E")
    for i, t in enumerate(C.AUTO_EINGEBAUT, start=1):
        r += 1
        _para(ws, r, f"{i}.  {t}", c1="B", c2="E", width_chars=186)
    r += 2
    _section(ws, r, "Alternative für ein privates Repository: Power Query einrichten (etwa 10 Minuten)", "B", "E")
    for i, t in enumerate(C.AUTO_SCHRITTE, start=1):
        r += 1
        _para(ws, r, f"{i}.  {t}", c1="B", c2="E", width_chars=186)
    r += 2
    _section(ws, r, "Wichtige Hinweise", "B", "E", kind="result")
    for t in C.AUTO_HINWEISE:
        r += 1
        _para(ws, r, "•  " + t, c1="B", c2="E", width_chars=186)
    r += 2
    _section(ws, r, "M-Code zum Kopieren (eine Zeile je Zelle: Spalte markieren, kopieren, im Erweiterten Editor einfügen)", "B", "E", kind="calc")
    r += 1
    ws[f"B{r}"] = "Parameter GitHubBaseUrl (Typ Text), Wert:"
    ws[f"B{r}"].font = font(bold=True)
    r += 1
    mono = font(name="Consolas", size=9)
    ws[f"B{r}"] = "https://raw.githubusercontent.com/<ORG>/<REPO>/main/data/processed/"
    ws[f"B{r}"].font = mono
    ws[f"B{r}"].fill = fill(L_GREY)
    for name, filename, types in M_QUERIES:
        r += 2
        ws[f"B{r}"] = f"Abfrage {name}  →  lädt {filename}  →  Ziel: Blatt {name}, Zelle A4"
        ws[f"B{r}"].font = font(bold=True, color=NAVY)
        for line in m_code(name, filename, types):
            r += 1
            c = ws[f"B{r}"]
            c.value = line
            c.font = mono
            c.fill = fill(L_GREY)
            c.alignment = Alignment(wrap_text=False)
    _fit_width(ws, landscape=False)


# --------------------------------------------------------------------------- Zusammenbau
def write_m_files(root: Path) -> list[Path]:
    """Power-Query-Code als .m-Dateien ins Repository schreiben (excel/power_query)."""
    out_dir = root / "excel" / "power_query"
    out_dir.mkdir(parents=True, exist_ok=True)
    paths = []
    for name, filename, types in C.M_QUERIES:
        p = out_dir / f"{name}.m"
        p.write_text("// Power-Query-Abfrage " + name + " – lädt " + filename + "\n" + "\n".join(C.m_code(name, filename, types)) + "\n",
                     encoding="utf-8")
        paths.append(p)
    param = out_dir / "GitHubBaseUrl.m"
    param.write_text('// Parameter (Typ Text): Basis-Link der CSV-Dateien im Repository (mit Schrägstrich am Ende)\n'
                     '"https://raw.githubusercontent.com/<ORG>/<REPO>/main/data/processed/" '
                     'meta [IsParameterQuery = true, Type = "Text", IsParameterQueryRequired = true]\n', encoding="utf-8")
    paths.append(param)
    return paths


# --------------------------------------------------------------------------- Netzkarte (schematisch)
K_BREITE = 51.0
KL_SEG = N_LINES                         # Stromkreis-Plätze
KL_LAST = R0 + 3 * KL_SEG - 1            # je Stromkreis 3 Zeilen: Anfang, Ende, Trennzeile


def build_karte_daten(wb):
    ws = wb["Karte_Daten"]
    ws["A1"] = "Karte_Daten – Hilfsblatt: Lage und Ampel je Umspannwerk für die Netzkarte im Dashboard"
    ws["A1"].font = font(bold=True, size=14, color=NAVY)
    ws["A2"] = ("Zeile n = Zeile n im Blatt Umspannwerke. Die Lage kommt aus D_Koordinaten. X wird mit cos(Breite) verkürzt, damit Deutschland nicht gestreckt wirkt. "
                "#NV in den Ampelspalten ist gewollt: Das Diagramm zeichnet solche Punkte nicht.")
    ws["A2"].font = font(italic=True, size=9, color="595959")
    ws.merge_cells("A2:L2")
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[2].height = 28
    band(ws, "A", "G", 3, "UMSPANNWERK UND LAGE", "data")
    band(ws, "H", "L", 3, "AMPEL (Richtung laut Dashboard)", "calc")
    band(ws, "N", "Q", 3, "UMRISS DEUTSCHLANDS (fest)", "data")
    band(ws, "S", "S", 3, "BESCHRIFTUNG", "data")
    heads = {"A": "Umspannwerk", "B": "Knotentyp", "C": "Breitengrad", "D": "Längengrad", "E": "Genauigkeit der Lage", "F": "Karte X", "G": "Karte Y",
             "H": "Ampel", "I": "Punkt Rot", "J": "Punkt Gelb", "K": "Punkt Grün", "L": "Punkt Grau",
             "N": "Längengrad", "O": "Breitengrad", "P": "Karte X", "Q": "Karte Y", "S": "Legendentext / Zähler"}
    _calc_headers(ws, heads, {"A": 28, "B": 20, "C": 11, "D": 11, "E": 36, "F": 9, "G": 9, "H": 8, "I": 9, "J": 9, "K": 9, "L": 9, "M": 3,
                              "N": 10, "O": 10, "P": 9, "Q": 9, "R": 3, "S": 42}, height=34)
    KO = lambda c: f"D_Koordinaten!${c}$5:${c}$1000"
    for r in range(R0, U_LAST + 1):
        pos = f"MATCH($A{r},{KO('A')},0)"
        f = {
            "A": f'=IF(Umspannwerke!A{r}="","",Umspannwerke!A{r})',
            "B": f'=IF($A{r}="","",Umspannwerke!C{r})',
            "C": f'=IF($A{r}="","",IFERROR(IF(INDEX({KO("B")},{pos})="","",INDEX({KO("B")},{pos})),""))',
            "D": f'=IF($A{r}="","",IFERROR(IF(INDEX({KO("C")},{pos})="","",INDEX({KO("C")},{pos})),""))',
            "E": f'=IF($A{r}="","",IFERROR(INDEX({KO("D")},{pos})&"",""))',
            "F": f'=IF(OR($C{r}="",$D{r}=""),NA(),$D{r}*COS(RADIANS(K_Breite)))',
            "G": f'=IF(OR($C{r}="",$D{r}=""),NA(),$C{r})',
            "H": f'=IF($A{r}="","",IF(K_Richtung="Entnahme",Umspannwerke!E{r},Umspannwerke!D{r}))',
        }
        for letter, name in (("I", "Rot"), ("J", "Gelb"), ("K", "Grün"), ("L", "Grau")):
            f[letter] = f'=IF(AND($B{r}="Netzknoten DE",$H{r}="{name}",ISNUMBER($G{r})),$G{r},NA())'
        write_formulas(ws, r, f, {"C": "0.0000", "D": "0.0000", "F": "0.000", "G": "0.000", "I": "0.000", "J": "0.000", "K": "0.000", "L": "0.000"})
    for i, (lon, lat) in enumerate(OUTLINE):
        r = R0 + i
        ws[f"N{r}"], ws[f"O{r}"] = lon, lat
        ws[f"N{r}"].font = ws[f"O{r}"].font = BLUE_FONT
        write_formulas(ws, r, {"P": f"=N{r}*COS(RADIANS(K_Breite))", "Q": f"=O{r}"}, {"P": "0.000", "Q": "0.000"})
    labels = [("S5", f'="Rot – ausgeschöpft / Engpass ("&COUNT($I${R0}:$I${U_LAST})&")"'),
              ("S6", f'="Gelb – knapp / hoch belastet ("&COUNT($J${R0}:$J${U_LAST})&")"'),
              ("S7", f'="Grün – frei / unkritisch ("&COUNT($K${R0}:$K${U_LAST})&")"'),
              ("S8", f'="Grau – keine Daten ("&COUNT($L${R0}:$L${U_LAST})&")"')]
    for ref, fml in labels:
        ws[ref] = fml
        color_formula(ws[ref])
    ws["S10"] = "Breite für die Kartenverkürzung (°):"
    ws["S11"] = K_BREITE
    ws["S11"].font = BLUE_FONT
    ws["S11"].fill = fill(L_GOLD)
    wb.defined_names["K_Breite"] = DefinedName("K_Breite", attr_text="Karte_Daten!$S$11")
    ws.freeze_panes = "B5"
    _fit_width(ws, title_rows="3:4")


def build_karte_linien(wb):
    ws = wb["Karte_Linien"]
    ws["A1"] = "Karte_Linien – Hilfsblatt: Netzlinien für die Karte (je Stromkreis drei Zeilen: Anfang, Ende, Lücke)"
    ws["A1"].font = font(bold=True, size=14, color=NAVY)
    ws["A2"] = ("Jeder Stromkreis aus dem Blatt Leitungen wird als Strecke zwischen den Lagen seiner beiden Umspannwerke gezeichnet. Die dritte Zeile (#NV) "
                "trennt die Strecken. Grau = alle Stromkreise, Gelb = hoch belastet, Rot = Engpass.")
    ws["A2"].font = font(italic=True, size=9, color="595959")
    ws.merge_cells("A2:H2")
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[2].height = 28
    band(ws, "A", "E", 3, "ENDPUNKT UND LAGE", "data")
    band(ws, "F", "H", 3, "LINIENFARBE (Werte für das Diagramm)", "calc")
    _calc_headers(ws, {"A": "Umspannwerk (Ende)", "B": "Position in Karte_Daten", "C": "Karte X", "D": "Karte Y", "E": "Status des Stromkreises",
                       "F": "Y grau (alle)", "G": "Y gelb", "H": "Y rot"},
                  {"A": 28, "B": 11, "C": 9, "D": 9, "E": 16, "F": 11, "G": 9, "H": 9}, height=34)
    for i in range(KL_SEG):
        rl = R0 + i                      # Zeile im Blatt Leitungen
        base = R0 + 3 * i
        for k, src in enumerate(("D", "E", None)):
            r = base + k
            if src:
                a = f'=IF(Leitungen!${src}{rl}="","",Leitungen!${src}{rl})'
                e = f'=IF($A{r}="","",Leitungen!$W{rl})'
            else:
                a, e = '=""', '=""'
            write_formulas(ws, r, {
                "A": a,
                "B": f'=IF($A{r}="","",IFERROR(MATCH($A{r},Karte_Daten!$A${R0}:$A${U_LAST},0),""))',
                "C": f'=IF($B{r}="",NA(),INDEX(Karte_Daten!$F${R0}:$F${U_LAST},$B{r}))',
                "D": f'=IF($B{r}="",NA(),INDEX(Karte_Daten!$G${R0}:$G${U_LAST},$B{r}))',
                "E": e,
                "F": f'=IF(ISNUMBER($D{r}),$D{r},NA())',
                "G": f'=IF(AND(ISNUMBER($D{r}),$E{r}="hoch belastet"),$D{r},NA())',
                "H": f'=IF(AND(ISNUMBER($D{r}),$E{r}="Engpass"),$D{r},NA())',
            }, {"C": "0.000", "D": "0.000", "F": "0.000", "G": "0.000", "H": "0.000"})
    ws.freeze_panes = "A5"
    _fit_width(ws, title_rows="3:4")


def _axis_hidden(ax, lo, hi):
    ax.scaling.min, ax.scaling.max = lo, hi
    ax.delete = False
    ax.majorGridlines = None
    ax.tickLblPos = "none"
    ax.majorTickMark = "none"
    ax.minorTickMark = "none"
    ax.number_format = ";;;"          # auch dort keine Zahlen, wo tickLblPos ignoriert wird
    ax.spPr = GraphicalProperties(ln=LineProperties(noFill=True))


def add_map(wb, ws):
    """Bedienelemente und Streudiagramm (Netzkarte) rechts neben den Dashboard-Tabellen."""
    ws.column_dimensions["L"].width = 16
    ws["L2"] = "Netzkarte Deutschland (schematisch)"
    ws["L2"].font = font(bold=True, size=12, color=GREEN_D)
    ws["L3"] = "Ampel für:"
    ws["L3"].font = font(bold=True)
    ws["L3"].alignment = Alignment(horizontal="right")
    ws["M3"] = "Einspeisung"
    ws["M3"].fill = fill(L_GOLD)
    ws["M3"].font = font(color="0B4BFF", bold=True)
    ws["M3"].border = BORDER
    ws["M3"].alignment = Alignment(horizontal="center")
    ws.column_dimensions["M"].width = 14
    dv = DataValidation(type="list", formula1='"Einspeisung,Entnahme"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(ws["M3"])
    wb.defined_names["K_Richtung"] = DefinedName("K_Richtung", attr_text="Dashboard!$M$3")
    ws["N3"] = "← Einspeisung oder Entnahme wählen"
    ws["N3"].font = font(size=9, italic=True, color="595959")
    ws["L4"] = ("Die Lage der Umspannwerke ist aus Leitungslängen und einigen Ankerpunkten geschätzt (Median-Abweichung rund 16 km, einzelne Knoten über 50 km): "
                "eine Orientierung, kein exakter Lageplan. Genauigkeit je Umspannwerk: Blatt D_Koordinaten.")
    ws["L4"].font = font(size=8, italic=True, color="595959")
    ws.merge_cells("L4:S5")
    ws["L4"].alignment = Alignment(wrap_text=True, vertical="top")

    kd, kl = wb["Karte_Daten"], wb["Karte_Linien"]
    ref = lambda sh, c, r1, r2: Reference(sh, min_col=c, min_row=r1, max_row=r2)
    ch = ScatterChart()
    ch.scatterStyle = "lineMarker"
    ch.title = "Ampel je Umspannwerk"
    ch.title.overlay = False
    ch.roundedCorners = False
    ch.height, ch.width = 18.6, 14.4

    def line_series(xs, ys, color, width, name):
        s = Series(ys, xs)
        s.tx = SeriesLabel(v=name)
        s.marker.symbol = "none"
        s.graphicalProperties.line.solidFill = color
        s.graphicalProperties.line.width = width
        s.smooth = False
        return s

    def point_series(c_y, color, label_cell):
        s = Series(ref(kd, c_y, R0, U_LAST), ref(kd, 6, R0, U_LAST))
        s.tx = SeriesLabel(strRef=StrRef(f"'Karte_Daten'!{label_cell}"))
        s.marker.symbol = "circle"
        s.marker.size = 7
        s.marker.graphicalProperties = GraphicalProperties(solidFill=color)
        s.marker.graphicalProperties.line.solidFill = "FFFFFF"
        s.graphicalProperties.line.noFill = True
        s.smooth = False
        return s

    lx = ref(kl, 3, R0, KL_LAST)
    ch.series.append(line_series(ref(kd, 16, R0, R0 + len(OUTLINE) - 1), ref(kd, 17, R0, R0 + len(OUTLINE) - 1), "000F47", 15875, "Umriss"))
    ch.series.append(line_series(lx, ref(kl, 6, R0, KL_LAST), "B9B6B1", 6350, "Netzlinien"))
    ch.series.append(line_series(lx, ref(kl, 7, R0, KL_LAST), "F4B400", 15875, "Netzlinien hoch belastet"))
    ch.series.append(line_series(lx, ref(kl, 8, R0, KL_LAST), "C53532", 15875, "Netzlinien Engpass"))
    ch.series.append(point_series(12, "B9B6B1", "$S$8"))
    ch.series.append(point_series(11, "6ABF30", "$S$7"))
    ch.series.append(point_series(10, "F4B400", "$S$6"))
    ch.series.append(point_series(9, "C53532", "$S$5"))
    # Seitenverhältnis: Länge 5,4-15,4 °O (mal cos 51°) zu Breite 46,9-55,3 °N
    c51 = math.cos(math.radians(K_BREITE))
    _axis_hidden(ch.x_axis, round(5.4 * c51, 3), round(15.4 * c51, 3))
    _axis_hidden(ch.y_axis, 46.9, 55.3)
    ch.legend.position = "b"
    ch.legend.overlay = False
    ch.legend.legendEntry = [LegendEntry(idx=i, delete=True) for i in range(4)]
    ch.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))
    ch.plot_area.graphicalProperties = GraphicalProperties(solidFill="EAF5FF", ln=LineProperties(noFill=True))
    ws.add_chart(ch, "L6")


def apply_ow_theme(wb):
    """Feinarbeit am Erscheinungsbild: Titel nach dem Muster 'Thema | Aussage' (Georgia), keine Gitternetzlinien auf Anzeigeblättern."""
    for ws in wb.worksheets:
        if not ws.title.startswith(("D_", "Karte_")):
            ws.sheet_view.showGridLines = False
        for ref in ("A1", "B1"):
            c = ws[ref]
            v = c.value
            if isinstance(v, str) and not v.startswith("=") and c.font is not None and c.font.b and (c.font.sz or 10) >= 14:
                if " – " in v:
                    topic, rest = v.split(" – ", 1)
                    c.value = CellRichText(TextBlock(InlineFont(rFont=FONT_TITLE, b=True, sz=18, color=NAVY), topic + " | "),
                                           TextBlock(InlineFont(rFont=FONT_TITLE, sz=18, color=NAVY), rest))
                else:
                    c.font = font(bold=True, size=18, color=NAVY, name=FONT_TITLE)
                ws.row_dimensions[1].height = 34
                break


# --------------------------------------------------------------------------- Zusammenbau
def build_workbook(root: Path, out_path: Path, base_url: str | None = None, connections: bool = True) -> Path:
    """Modell erzeugen. connections=True baut je Rohdatenblatt eine Datenverbindung ein (Adresse: base_url)."""
    data = load_inputs(root)
    _STATE["connections"], _STATE["specs"] = connections, []
    base_url = base_url or f"https://raw.githubusercontent.com/{conn.PLACEHOLDER}/main/data/processed/"
    wb = Workbook()
    wb.remove(wb.active)
    for name in SHEETS:
        ws = wb.create_sheet(name)
        if name in TAB:
            ws.sheet_properties.tabColor = TAB[name]
    build_landing_sheets(wb, data)
    build_parameter(wb)
    build_leitungen(wb, len(data["sgm_leitungen"]))
    build_trafos(wb)
    build_umspannwerke(wb)
    build_karte_daten(wb)
    build_karte_linien(wb)
    build_dashboard(wb, data)
    build_datenstand(wb, base_url if connections else None)
    build_lies_mich(wb)
    build_modelllogik(wb, data)
    build_glossar(wb, data)
    build_quellen(wb, data)
    build_automatisierung(wb)
    apply_ow_theme(wb)
    wb.properties.title = "Kapazitätsmodell Deutschland (HöS/HS) – Version 3"
    wb.properties.subject = "Ampelmodell Netzauslastung und Anschlusskapazität auf Basis öffentlicher Daten"
    wb.properties.creator = "Projektteam Kapazitätsmodell"
    wb.calculation.fullCalcOnLoad = True
    wb.active = 0
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    if connections:
        conn.inject(str(out_path), _STATE["specs"], base_url)
    conn.patch_scatter_na_as_blank(str(out_path))
    return out_path


