#!/usr/bin/env python3
"""Erzeugt die Umspannwerks-Liste (data/processed/sgm_umspannwerke.csv) aus Leitungen und Trafos.

Regel (wie im Modell V2): Jeder Name, der als Umspannwerk 1/2 einer Leitung oder als Umspannwerk
eines Trafos vorkommt, ist ein Netzknoten. Namen werden bereinigt und alphabetisch sortiert.
"""
from __future__ import annotations

import sys

import pandas as pd

import _common as common
from kapmodell.normalize import norm_text


def derive(lines: pd.DataFrame, trafos: pd.DataFrame) -> list[str]:
    names = set()
    for col in ("Umspannwerk_1", "Umspannwerk_2"):
        names |= {norm_text(v) for v in lines[col]}
    names |= {norm_text(v) for v in trafos["Umspannwerk"]}
    return sorted((n for n in names if n), key=str.casefold)


def main() -> int:
    proc = common.ROOT / "data" / "processed"
    lines = pd.read_csv(proc / "sgm_leitungen.csv", dtype=str, keep_default_na=False)
    trafos = pd.read_csv(proc / "sgm_trafos.csv", dtype=str, keep_default_na=False)
    names = derive(lines, trafos)
    pd.DataFrame({"Umspannwerk": names}).to_csv(proc / "sgm_umspannwerke.csv", index=False,
                                                lineterminator="\n", encoding="utf-8")
    print(f"{len(names)} Umspannwerke / Netzknoten geschrieben")
    return 0


if __name__ == "__main__":
    sys.exit(main())
