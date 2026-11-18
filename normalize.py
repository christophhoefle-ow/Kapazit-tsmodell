"""Text- und Schlüssel-Normalisierung.

Warum das nötig ist: Die JAO-Rohdaten enthalten Netzelement-Namen mit Leerzeichen am Ende
(z. B. 'Diefflen - Ensdorf ENSDF N '). Ohne Bereinigung findet ein Namensvergleich diese
Elemente nicht – im Test fehlten dadurch 12 von 905 Leitungen.
"""
from __future__ import annotations

import re

_WS = re.compile(r"\s+")
_PLACEHOLDER_EIC = {"", "nan", "none", "ohne eic", "-"}


def norm_text(value) -> str:
    """Geschützte Leerzeichen -> normale, mehrfache Leerzeichen -> eins, Rand abschneiden."""
    if value is None:
        return ""
    try:
        if value != value:  # NaN
            return ""
    except Exception:  # pragma: no cover
        pass
    return _WS.sub(" ", str(value).replace("\xa0", " ")).strip()


def norm_eic(value) -> str:
    """EIC bereinigen; Platzhalter ('ohne EIC', '11TDNOTAVAILABLE' u. ä.) werden zu leerem Text."""
    text = norm_text(value)
    low = text.lower()
    return "" if (low in _PLACEHOLDER_EIC or "notavailable" in low) else text


def name_key(name) -> str:
    """Vergleichsform eines Namens (ohne Groß-/Kleinschreibung)."""
    return norm_text(name).casefold()


def element_key(eic, name) -> str:
    """Schlüssel eines Netzelements: EIC, ersatzweise 'NAME|<Name>'.

    Dieselbe Regel gilt im Excel-Modell (Blatt 'Leitungen', Spalte 'Schlüssel').
    """
    eic_n = norm_eic(eic)
    return eic_n if eic_n else "NAME|" + norm_text(name)
