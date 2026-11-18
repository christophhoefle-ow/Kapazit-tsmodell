# Betriebshandbuch

## Täglicher Betrieb
Die Action *JAO-Auslastung abrufen* läuft täglich um 20:15 UTC. Sie führt die Tests aus, holt fehlende Tage der letzten 7 Tage, baut
`data/processed/*.csv` und committet die Änderungen. Schlägt ein Tag fehl, werden die übrigen trotzdem gespeichert und der Lauf wird rot.

### Protokoll lesen (`data/interim/fetch_log.csv`)
| Status | Bedeutung | Maßnahme |
|---|---|---|
| `vollstaendig` | alle Stunden des Geschäftstags geliefert | keine |
| `unvollstaendig` | weniger Stunden als erwartet | wird bis zum Ende des 7-Tage-Fensters erneut versucht; bleibt es so, hat JAO einzelne Stunden überbrückt (Spanning) |
| `wartet` | JAO meldet Expected oder Pending | wird beim nächsten Lauf versucht |
| `fallback` | JAO hat das Backup-Verfahren genutzt (Spanning/DFP), keine Domain-Daten | nicht erneut versuchen |

### Typische Fehler
* **HTTP 400:** Zeitraum zu lang oder Datum fehlt. Nicht unverändert wiederholen.
* **HTTP 429:** zu viele Anfragen (Limit 100/Minute) oder mehrere gleichzeitig. Der Client wartet und versucht es erneut.
* **`rejected = true`:** JAO lehnt ab (auch bei HTTP 200 möglich). Meldung im Protokoll lesen.
* **Seit 19.11.2026:** Anfragen mit `Take` über 100.000 werden abgelehnt (der Client begrenzt auf 100.000).

## Zeitraum nachholen
*Actions → JAO-Auslastung abrufen → Run workflow*, Start- und Enddatum eintragen. Pro Lauf höchstens 31 Tage (Schutz). Für Monate bis Jahre
**vorher JAO ansprechen** (Leitfaden Kap. 5) und außerhalb der Veröffentlichungszeit (ca. 09–11 Uhr) laufen lassen.

## Netzmodell (SGM) aktualisieren
1. Der Monatsjob legt bei einer neuen Release ein Issue an.
2. ZIP laden, Excel-Datei ablegen, Stammdaten-CSV erzeugen (Einlese-Skript noch offen), `python scripts/validate_sgm.py`.
3. `python scripts/derive_umspannwerke.py`, `data/manual/sgm_release.txt` anpassen.
4. `python scripts/build_processed.py`, Modell bauen, `python scripts/check_model.py <berechnete Datei>`.
5. Namens- oder EIC-Änderungen: offene Elemente im Blatt *Datenstand* ansehen.

## JAO-Elemente ohne Zuordnung
`python scripts/suggest_aliases.py` schreibt `data/manual/alias_vorschlaege.csv`. **Nie automatisch übernehmen**: Die Ähnlichkeit täuscht
(Beispiel: "HALTRN W" und "HALTERN O" sind verschiedene Stränge). Bestätigte Paare in `data/manual/alias_jao_sgm.csv` eintragen
(Spalten `Schluessel_JAO`, `Schluessel_SGM`, `Bemerkung`) und `build_processed.py` ausführen.

## Zugangsdaten
Nie in Dateien, nur als *Secrets*: `JAO_API_TOKEN` (optional), später `ENTSOE_API_TOKEN` und für die Netztransparenz-API Client-ID und Secret.
Für Excel bei privatem Repository: fein abgestuftes Token mit Nur-Lese-Recht, nicht in der weitergegebenen Datei speichern.

## Schwellen ändern
`config/config.yml` (`schwelle_hoch`, `schwelle_ueberlast`). Das Excel-Blatt *Parameter* hat eigene Schwellen für Status und Ampel; das Blatt
*Datenstand* meldet, wenn beide auseinanderlaufen.

## Excel-Verbindungen ("Daten → Alle aktualisieren")
Die Datei `Kapazitaetsmodell_DE_V3_aktualisierbar.xlsx` enthält je grauem Rohdatenblatt eine Verbindung zur CSV-Datei im Repository (klassische Text-/CSV-Abfrage).
Beim ersten Öffnen: *Bearbeitung aktivieren*, dann *Inhalt aktivieren* (Datenverbindungen). Danach *Daten → Alle aktualisieren*; Kontrolle im Blatt *Datenstand*.

| Symptom | Ursache und Maßnahme |
|---|---|
| Datenstand zeigt "PRÜFEN" bei der Datenquelle | Adresse ist noch der Platzhalter ORG/REPO: `python scripts/set_base_url.py <Datei> <https://raw.githubusercontent.com/ORG/REPO/main/data/processed/>` oder die Datei über die Aktion *Excel-Modell bauen* erzeugen |
| Fehler beim Aktualisieren (404 oder "Datei nicht gefunden") | Repository ist privat (klassische Textabfragen können sich nicht anmelden) oder die Adresse stimmt nicht. Öffentlich machen oder Power Query mit Token nutzen (`docs/POWER_QUERY.md`) |
| Excel meldet beim Öffnen "Inhalte beschädigt, Reparatur?" | Die eingebauten Verbindungen wurden nicht in Excel getestet (die erste Fassung mit Tabellenbindung wurde bereits als beschädigt gemeldet). Datei `Kapazitaetsmodell_DE_V3.xlsx` (ohne Verbindungen) nehmen, Power Query einrichten und die Meldung im Wortlaut ans Projektteam geben |
| Zahlen falsch (Dezimalpunkt) | Verbindung → Eigenschaften → Definition: Dezimaltrennzeichen "." sicherstellen; Spaltenreihenfolge im Blatt Datenstand prüfen |

Wert-Vorschau: Die Aktion erzeugt Dateien ohne gespeicherte Rechenergebnisse (Excel rechnet beim Öffnen selbst; in der geschützten Ansicht erscheinen Zellen leer). Mit installiertem LibreOffice lassen sich Ergebnisse vorab eintragen:
`python scripts/add_cached_values.py ORIGINAL.xlsx VON_LIBREOFFICE_BERECHNET.xlsx AUSGABE.xlsx`.

## Koordinaten und Netzkarte
* `python scripts/estimate_coordinates.py` erzeugt `data/processed/umspannwerke_koordinaten.csv` und `data/manual/koordinaten_pruefbericht.csv`. Nach jeder neuen SGM-Release ausführen.
* Rangfolge: `koordinaten_manuell.csv` (exakt, z. B. aus OpenStreetMap) vor `koordinaten_anker.csv` (ungefähre Ortslage) vor Schätzung aus Leitungslängen.
* Die Anker wurden aus allgemeinem Kartenwissen gesetzt, nicht aus einer geprüften Quelle. Das Skript entfernt Anker, die zu den Leitungslängen nicht passen, und listet sie im Prüfbericht. **Bitte Prüfbericht lesen**: Dort stehen auch Umspannwerksnamen, die vermutlich mehrere Standorte verbinden.
* Gemessene Genauigkeit (Kreuzvalidierung): Median 16 km, 88 % unter 50 km, 98 % unter 100 km, Maximum 232 km.
