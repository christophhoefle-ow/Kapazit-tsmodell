"""Ablage der Tages-, Monats- und Protokolldateien (CSV, UTF-8, Punkt als Dezimalzeichen).

Aufbau: data/interim/daily/JJJJ-MM.csv  eine Datei je Monat, eine Zeile je Netzelement und Tag.
Eine Tagesaktualisierung ändert nur die Datei des laufenden Monats – das hält die Git-Historie klein.
"""
from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

import pandas as pd

from .aggregate import DAILY_COLUMNS

LOG_COLUMNS = ["Tag", "Stunden_erwartet", "Stunden_geliefert", "Zeilen_DE", "Status", "Abruf_UTC"]
_TEXT_COLS = {"Tag": str, "Schluessel_JAO": str, "UENB": str, "Netzelement": str, "Namen": str, "EIC": str}


def _write_csv(df: pd.DataFrame, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(path, index=False, lineterminator="\n", encoding="utf-8")


def _ensure_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Ältere Tagesdateien ohne neue Spalten (z. B. Stunden_bindend) mit 0 auffüllen."""
    for c in DAILY_COLUMNS:
        if c not in df.columns:
            df[c] = "" if c in _TEXT_COLS else 0
    return df[DAILY_COLUMNS]


def read_daily(daily_dir) -> pd.DataFrame:
    files = sorted(Path(daily_dir).glob("????-??.csv"))
    if not files:
        return pd.DataFrame(columns=DAILY_COLUMNS)
    parts = [_ensure_columns(pd.read_csv(f, dtype=_TEXT_COLS, keep_default_na=False)) for f in files]
    return pd.concat(parts, ignore_index=True)


def write_day(daily_dir, tag: str, df: pd.DataFrame) -> Path:
    """Zeilen eines Tages in die Monatsdatei schreiben; vorhandene Zeilen desselben Tages werden ersetzt."""
    path = Path(daily_dir) / f"{tag[:7]}.csv"
    if path.exists():
        old = _ensure_columns(pd.read_csv(path, dtype=_TEXT_COLS, keep_default_na=False))
        old = old[old["Tag"] != tag]
        df = pd.concat([old, df], ignore_index=True)
    df = df[DAILY_COLUMNS].copy()
    df["Fmax_MW"] = pd.to_numeric(df["Fmax_MW"]).round(3)
    df["Auslastung_Summe"] = pd.to_numeric(df["Auslastung_Summe"]).round(6)
    df["Auslastung_Max"] = pd.to_numeric(df["Auslastung_Max"]).round(6)
    _write_csv(df.sort_values(["Tag", "Schluessel_JAO"]), path)
    return path


def read_log(path) -> pd.DataFrame:
    p = Path(path)
    if not p.exists():
        return pd.DataFrame(columns=LOG_COLUMNS)
    return pd.read_csv(p, dtype={"Tag": str, "Status": str, "Abruf_UTC": str}, keep_default_na=False)


def update_log(path, entry: dict) -> None:
    log = read_log(path)
    log = log[log["Tag"] != entry["Tag"]]
    log = pd.concat([log, pd.DataFrame([entry])], ignore_index=True)[LOG_COLUMNS]
    _write_csv(log.sort_values("Tag"), Path(path))


def days_to_fetch(today: str, lookback_days: int, log: pd.DataFrame) -> list[str]:
    """Tage der letzten `lookback_days` Tage (bis einschließlich heute), die noch nicht vollständig vorliegen."""
    end = date.fromisoformat(today)
    # 'fallback' = JAO hat das Backup-Verfahren genutzt (keine Domain-Daten): nicht erneut versuchen
    complete = set(log.loc[log["Status"].isin(["vollstaendig", "fallback"]), "Tag"]) if not log.empty else set()
    days = [(end - timedelta(days=i)).isoformat() for i in range(lookback_days, -1, -1)]
    return [d for d in days if d not in complete]
