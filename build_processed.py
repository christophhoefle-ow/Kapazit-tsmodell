#!/usr/bin/env python3
"""Aus den Tagesdateien die Excel-Lieferdateien erzeugen (data/processed/).

Ergebnis:
    auslastung_monatlich.csv   -> Excel-Blatt D_Auslastung
    meta.csv                   -> Excel-Blatt D_Meta (Datenstand, Schwellen, Release)
"""
from __future__ import annotations

import logging
import os
import sys
from datetime import datetime, timezone

import pandas as pd

import _common as common
from kapmodell import __version__
from kapmodell.mapping import MONTHLY_COLUMNS, monthly_from_daily, sgm_elements
from kapmodell.storage import read_daily, read_log

log = logging.getLogger("build_processed")


def main() -> int:
    common.setup_logging()
    cfg = common.load_config()
    proc = common.path(cfg, "processed_dir")
    daily = read_daily(common.path(cfg, "daily_dir"))

    lines = pd.read_csv(proc / "sgm_leitungen.csv", dtype=str, keep_default_na=False)
    trafos = pd.read_csv(proc / "sgm_trafos.csv", dtype=str, keep_default_na=False)
    alias_file = common.ROOT / "data" / "manual" / "alias_jao_sgm.csv"
    alias = pd.read_csv(alias_file, dtype=str, keep_default_na=False) if alias_file.exists() else None
    monthly = monthly_from_daily(daily, sgm_elements(lines, trafos), alias)
    monthly.to_csv(proc / "auslastung_monatlich.csv", index=False, lineterminator="\n", encoding="utf-8")
    log.info("auslastung_monatlich.csv: %d Zeilen, %d Monate", len(monthly), monthly["Monat"].nunique())

    protokoll = read_log(common.path(cfg, "fetch_log"))
    n_voll = int((protokoll["Status"] == "vollstaendig").sum()) if not protokoll.empty else 0
    release_file = common.ROOT / "data" / "manual" / "sgm_release.txt"
    release = release_file.read_text(encoding="utf-8").strip() if release_file.exists() else "unbekannt"
    zuordnung = monthly["Zuordnung"].value_counts().to_dict() if not monthly.empty else {}
    meta = [
        ("erstellt_utc", datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M")),
        ("pipeline_version", __version__),
        ("git_commit", os.environ.get("GITHUB_SHA", "lokal")[:7]),
        ("daten_von_tag", daily["Tag"].min() if not daily.empty else ""),
        ("daten_bis_tag", daily["Tag"].max() if not daily.empty else ""),
        ("tage_mit_daten", daily["Tag"].nunique() if not daily.empty else 0),
        ("tage_vollstaendig", n_voll),
        ("schwelle_hoch", cfg["auswertung"]["schwelle_hoch"]),
        ("schwelle_ueberlast", cfg["auswertung"]["schwelle_ueberlast"]),
        ("jao_datensatz", cfg["jao"]["dataset"]),
        ("sgm_release", release),
        ("zuordnung_alias", zuordnung.get("Alias", 0)),
        ("zuordnung_eic", zuordnung.get("EIC", 0)),
        ("zuordnung_name", zuordnung.get("Name", 0)),
        ("zuordnung_offen", zuordnung.get("nicht zugeordnet", 0)),
    ]
    extra = common.ROOT / "data" / "manual" / "meta_hinweis.txt"
    if extra.exists() and n_voll == 0:      # Beispiel-Hinweis verschwindet, sobald der erste Tag vollständig geladen ist
        meta.append(("hinweis", extra.read_text(encoding="utf-8").strip()))
    pd.DataFrame(meta, columns=["Schluessel", "Wert"]).to_csv(
        proc / "meta.csv", index=False, lineterminator="\n", encoding="utf-8")
    log.info("meta.csv geschrieben")
    return 0


if __name__ == "__main__":
    sys.exit(main())
