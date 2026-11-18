#!/usr/bin/env python3
"""JAO-Auslastungsdaten abrufen und je Tag verdichtet ablegen.

Beispiele
---------
Automatischer Betrieb (fehlende Tage der letzten 7 Tage bis heute):
    python scripts/fetch_jao.py

Zeitraum nachholen (das ist die Stelle, an der Start- und Enddatum festgelegt werden):
    python scripts/fetch_jao.py --from 2026-11-01 --to 2026-11-15

Nur anzeigen, was passieren würde:
    python scripts/fetch_jao.py --dry-run

Hinweis: Große Nachholläufe (mehrere Monate) bitte vorher mit JAO abstimmen und außerhalb der
Veröffentlichungszeit (morgens ca. 09-11 Uhr) laufen lassen.
"""
from __future__ import annotations

import argparse
import logging
import os
import sys
from datetime import datetime, timezone

import _common as common
from kapmodell.aggregate import daily_aggregate
from kapmodell.jao_client import JaoClient, JaoError
from kapmodell.storage import days_to_fetch, read_log, update_log, write_day
from kapmodell.timeutil import day_range, today_local

log = logging.getLogger("fetch_jao")


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--from", dest="first", help="erster Tag (JJJJ-MM-TT); ohne Angabe: automatisch")
    ap.add_argument("--to", dest="last", help="letzter Tag (JJJJ-MM-TT); Standard: heute")
    ap.add_argument("--force", action="store_true", help="auch bereits vollständige Tage neu abrufen")
    ap.add_argument("--max-days", type=int, default=31, help="Schutz vor versehentlich riesigen Läufen")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args()
    common.setup_logging(args.verbose)

    cfg = common.load_config()
    jao, ausw = cfg["jao"], cfg["auswertung"]
    daily_dir, log_path = common.path(cfg, "daily_dir"), common.path(cfg, "fetch_log")
    today = today_local(ausw["timezone"])

    if args.first:
        days = day_range(args.first, args.last or today)
        if not args.force:
            done = read_log(log_path)
            complete = set(done.loc[done["Status"] == "vollstaendig", "Tag"]) if not done.empty else set()
            days = [d for d in days if d not in complete]
    else:
        days = days_to_fetch(today, int(ausw["lookback_tage"]), read_log(log_path))

    if len(days) > args.max_days:
        log.error("%d Tage angefordert, erlaubt sind %d (--max-days). Abbruch.", len(days), args.max_days)
        return 2
    log.info("Zu bearbeitende Tage: %s", ", ".join(days) if days else "keine")
    if args.dry_run or not days:
        return 0

    base_url = os.environ.get("JAO_BASE_URL", jao["base_url"])    # Überschreibung z. B. für Tests
    client = JaoClient(base_url, jao["dataset"], jao["pause_seconds"], jao["timeout_seconds"],
                       jao["max_retries"], jao["take"], token=os.environ.get("JAO_API_TOKEN") or None)
    errors = 0
    for tag in days:
        try:
            if jao.get("require_monitoring_received", True):
                status = client.monitoring_status(tag)
                now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
                if status == "fallback":
                    log.info("%s: JAO meldet Spanning/DFP (Backup-Verfahren) – keine Domain-Daten, wird nicht erneut versucht", tag)
                    update_log(log_path, {"Tag": tag, "Stunden_erwartet": "", "Stunden_geliefert": 0, "Zeilen_DE": 0,
                                          "Status": "fallback", "Abruf_UTC": now})
                    continue
                if status in ("expected", "pending"):
                    log.info("%s: JAO meldet '%s' – wird beim nächsten Lauf versucht", tag, status)
                    update_log(log_path, {"Tag": tag, "Stunden_erwartet": "", "Stunden_geliefert": 0, "Zeilen_DE": 0,
                                          "Status": "wartet", "Abruf_UTC": now})
                    continue
            frame, n_expected, n_hours = client.fetch_day(tag, jao["tsos"], ausw["timezone"])
            agg = daily_aggregate(frame, tag, ausw["schwelle_hoch"], ausw["schwelle_ueberlast"])
            if not agg.empty:
                write_day(daily_dir, tag, agg)
            status = "vollstaendig" if (n_hours == n_expected and not agg.empty) else "unvollstaendig"
            update_log(log_path, {"Tag": tag, "Stunden_erwartet": n_expected, "Stunden_geliefert": n_hours,
                                  "Zeilen_DE": len(frame), "Status": status,
                                  "Abruf_UTC": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")})
            log.info("%s: %s (%d/%d Stunden, %d Zeilen DE, %d Netzelemente)", tag, status, n_hours,
                     n_expected, len(frame), len(agg))
        except JaoError as exc:
            errors += 1
            log.error("%s: Abruf fehlgeschlagen: %s", tag, exc)
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
