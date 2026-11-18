#!/usr/bin/env python3
"""Excel-Modell erzeugen (aus data/processed und config) und die Power-Query-Dateien aktualisieren.

  python scripts/build_excel_model.py --out build/Kapazitaetsmodell_DE_V3.xlsx --base-url https://raw.githubusercontent.com/<ORG>/<REPO>/main/data/processed/
  python scripts/build_excel_model.py --out build/Kapazitaetsmodell_DE_V3_ohne_Verbindungen.xlsx --no-connections
"""
import argparse
import sys
from pathlib import Path

import _common as common
from kapmodell.xl.build import build_workbook, write_m_files


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default="build/Kapazitaetsmodell_DE_V3.xlsx")
    ap.add_argument("--base-url", default=None, help="Adresse der CSV-Dateien (mit Schrägstrich am Ende). Ohne Angabe: Platzhalter ORG/REPO")
    ap.add_argument("--no-connections", action="store_true", help="keine Datenverbindungen einbauen (Daten nur als Momentaufnahme)")
    a = ap.parse_args()
    out = Path(a.out)
    if not out.is_absolute():
        out = common.ROOT / out
    build_workbook(common.ROOT, out, base_url=a.base_url, connections=not a.no_connections)
    n = len(write_m_files(common.ROOT))
    kind = "ohne Datenverbindungen" if a.no_connections else f"mit Datenverbindungen ({a.base_url or 'Platzhalter ORG/REPO: bitte --base-url angeben'})"
    print(f"Modell geschrieben: {out}  ({out.stat().st_size / 1e6:.1f} MB), {kind}; Power-Query-Dateien: {n}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
