# Kapazitätsmodell Deutschland (Höchst- und Hochspannung)

Ein Ampelmodell für das deutsche Stromnetz auf Basis **öffentlicher** Daten: Wie stark sind Leitungen, Trafos und Umspannwerke ausgelastet,
und wie viel Anschlusskapazität ist laut Netzbetreibern noch frei? Dieses Repository enthält die **Daten-Pipeline** (Python, GitHub Actions)
und den **Generator des Excel-Modells**. Das Modell erklärt sich in den Blättern *Lies_mich*, *Modelllogik* und *Glossar*; das Dashboard zeigt neben den Tabellen
eine **schematische Netzkarte** (Ampel je Umspannwerk, Netzlinien nach Status).

## Grundriss in einem Bild

```
 ÖFFENTLICHE QUELLEN            GITHUB (dieses Repository)                       EXCEL
 ───────────────────            ──────────────────────────────────────           ──────────────────────────────
 JAO Publication Tool  ───►     Action "JAO-Auslastung abrufen"  (täglich)       Daten → Alle aktualisieren
   (Final Computation)            24 stündliche Abrufe je Tag                         │  eingebaute Datenverbindungen
 JAO Static Grid Model ───►     Action "SGM-Release prüfen" (monatlich,             ▼  (Alternative: Power Query)
   (Netzstammdaten)               legt nur ein Issue an)                         graue D_-Blätter  (Rohdaten, 8 Stück)
 später: SMARD, MaStR,          scripts/ + src/kapmodell  (Python, getestet)        ▼
 ÜNB-Netzanschlusskarten,             │                                          dunkelblaue Blätter (Rechnung)
 VNBdigital / gridAhead,              ▼                                             ▼
 OpenStreetMap            data/interim/daily/    Tagesaggregate je Netzelement   grüne Blätter (Ergebnis, Ampel, Karte)
                          data/processed/        kleine CSV für Excel
```

**Leitgedanke:** Excel bekommt nie Rohdaten. Rund 165.000 Zeilen je Tag (nur die vier deutschen ÜNB) passen nicht in eine Tabelle, API-Zugangsdaten gehören nicht in
eine Datei, und das Dezimalpunkt-Problem beim CSV-Import (Faktor 10^10 in den Testdaten) entfällt, weil das Skript JSON liest. Jede Aktualisierung ist ein Commit.

## "Daten → Alle aktualisieren": was eingebaut ist und was nicht

Einrichtung Schritt für Schritt, auch für Nicht-Techniker: `docs/EINRICHTUNG.md`.

`Kapazitaetsmodell_DE_V3_aktualisierbar.xlsx` enthält in jedem der acht grauen Rohdatenblätter eine **Datenverbindung** zur passenden CSV-Datei dieses Repositories
(klassische Text-/CSV-Abfrage als Bereichsabfrage ab Zeile 4, keine Excel-Tabelle). Ablauf: Datei öffnen, *Bearbeitung aktivieren*, *Inhalt aktivieren*, *Daten → Alle aktualisieren*.

* **Das Repository muss öffentlich lesbar sein.** Klassische Textabfragen können sich nicht anmelden. Bei privatem Repository: Power Query mit Token (`docs/POWER_QUERY.md`).
* **Die Adresse steht fest in den Verbindungen.** Die Aktion *Excel-Modell bauen* trägt automatisch die Adresse dieses Repositories ein. Von Hand:
  `python scripts/set_base_url.py <Datei> https://raw.githubusercontent.com/<ORG>/<REPO>/main/data/processed/`. Das Blatt *Datenstand* meldet "PRÜFEN", solange noch ORG/REPO dasteht.
* **Nicht in Excel getestet.** Die erste Fassung band die Abfragen an Excel-Tabellen und wurde von Excel als beschädigt gemeldet (Excel legt Text-/Webabfragen nie als Tabelle an). Diese Fassung nutzt Bereichsabfragen wie Excel selbst; die XML-Teile sind gegen das Office-Schema (ISO/IEC 29500) geprüft, ob Excel sie annimmt, ist offen.
  Zum Vergleich liegt `Kapazitaetsmodell_DE_V3.xlsx` ohne Verbindungen bei (Momentaufnahme, Daten per Power Query oder von Hand einspielen).

## Wie wird Datum und Uhrzeit festgelegt?

Gar nicht von Hand. `scripts/fetch_jao.py` rechnet den deutschen Geschäftstag (23, 24 oder 25 Stunden) in UTC-Stundenfenster um und ruft sie nacheinander ab
(so verlangt es der JAO-Leitfaden). Ohne Angabe holt es alle **fehlenden** Tage der letzten 7 Tage. Zum Nachholen: *Actions → JAO-Auslastung abrufen → Run workflow*
mit Start- und Enddatum (oder lokal `python scripts/fetch_jao.py --from 2026-11-01 --to 2026-11-15`). Große Historien (Monate) vorher mit JAO abstimmen.

## Rhythmus je Quelle

| Quelle | Abruf | Auswertung im Modell |
|---|---|---|
| JAO Final Computation (Auslastung) | täglich, 20:15 UTC | Monatsverdichtung, Fenster frei wählbar (Standard 3 Monate) |
| JAO Static Grid Model (Netz) | monatlich prüfen, halbjährlich laden | nach fachlicher Prüfung; danach `estimate_coordinates.py` |
| SMARD / Redispatch, MaStR | monatlich (Phase 2) | Quartal bzw. je Umspannwerk |
| ÜNB-Netzanschlusskarten | je Verfahrensrunde | Bandbreiten, manuell in `D_Netzanschluss` |
| VNB-Netzausbaupläne (VNBdigital) | alle zwei Jahre | Hochspannung, Phase 2 |

Eine einzelne Stunde zeigt nur 0 oder 100 % Überlast. Sinnvoll sind mindestens ein Quartal und ergänzend ein 12-Monats-Bild.

## Netzkarte und Koordinaten

Das Netzmodell enthält **keine Koordinaten**. Die Karte nutzt deshalb `data/processed/umspannwerke_koordinaten.csv` (Blatt `D_Koordinaten`), erzeugt von `scripts/estimate_coordinates.py`:

1. **Anker:** `data/manual/koordinaten_anker.csv` (rund 270 Umspannwerke mit ungefährer Ortslage der gleichnamigen Gemeinde). Diese Werte stammen aus allgemeinem Kartenwissen, **nicht aus einer geprüften Quelle**.
2. **Prüfung gegen das Netz:** Die Luftlinie zwischen zwei Knoten kann nie länger sein als der Leitungsweg. Anker, die das mehrfach verletzen (falsche Koordinate oder Namensverwechslung), werden entfernt.
3. **Schätzung:** Aus Leitungslängen (kürzeste Wege durch das Netz) und den verbleibenden Ankern wird die Lage aller übrigen Knoten per Stress-Minimierung berechnet (`src/kapmodell/geo.py`).
4. **Exakte Werte gewinnen:** `data/manual/koordinaten_manuell.csv` (z. B. aus OpenStreetMap) überschreibt Anker und Schätzung.

Ergebnis für 497 Knoten: 230 Anker, 233 geschätzt, 34 ohne Lage (Name verbindet vermutlich mehrere Standorte oder Knoten ohne Anschluss ans Hauptnetz).
**Gemessene Genauigkeit** (Kreuzvalidierung: Anker weglassen, schätzen, vergleichen; 444 Messungen): Median 16 km, 88 % unter 50 km, 98 % unter 100 km, Maximum 232 km.
Mit Anker höchstens 25 km Netzabstand entfernt: Median 10 km. Spalte *Genauigkeit* in `D_Koordinaten` zeigt die Klasse je Knoten. Die Karte ist eine Orientierung, kein Lageplan.

**Gestaltung:** Farben und Schriften folgen den Oliver-Wyman-Folien (Navy `#000F47`, Hellblau `#CEECFF`, warme Grautöne, Ampel in gedämpftem Rot/Gelb/Grün; Arial, Titel in Georgia als Ersatz für die Hausschrift Marsh Serif).

**Befund zur Modellqualität:** Gleichnamige Umspannwerke verschiedener Betreiber werden im Modell zu EINEM Knoten zusammengefasst. Sehr wahrscheinlich ist das bei "Buers" so (Amprion in NRW, TransnetBW in Vorarlberg).
`data/manual/koordinaten_pruefbericht.csv` listet 15 Verdachtsfälle; einzelne davon sind vermutlich Fehlalarme durch falsche Anker. Eine saubere Lösung braucht eine Standort-ID statt des Namens.

## Ordner

| Pfad | Inhalt |
|---|---|
| `config/config.yml` | Schwellen, Zeitzone, Look-back, JAO-Einstellungen |
| `config/quellen.csv`, `glossar.csv` | Texte der Blätter *Quellen* und *Glossar* (ohne Code pflegbar) |
| `src/kapmodell/` | Bibliothek: JAO-Client, Aggregation, Zuordnung, Ablage, Lageschätzung (`geo.py`), Excel-Generator (`xl/`) |
| `scripts/` | Kommandozeilen-Werkzeuge: Abruf, Aufbereitung, Koordinaten, Prüfungen, Modell bauen, Adresse ändern, Werte eintragen |
| `data/interim/` | Tagesaggregate und Abrufprotokoll (Quelle der Wahrheit, wird von der Action gepflegt) |
| `data/processed/` | Lieferdateien für Excel (CSV) |
| `data/manual/` | Von Hand gepflegt: SGM-Release, Alias-Zuordnungen, Koordinaten-Anker und -Korrekturen, Prüfbericht |
| `excel/power_query/` | M-Code der acht Abfragen und des Parameters `GitHubBaseUrl` (Alternative zu den eingebauten Verbindungen) |
| `.github/workflows/` | `jao-daily`, `sgm-release-check`, `build-model` |
| `tests/` | 34 Tests (inkl. Schein-Server für die JAO-API und Gitternetz für die Lageschätzung) |

## In 15 Minuten einrichten

1. Neues Repository anlegen (für die eingebauten Verbindungen **öffentlich**, sonst privat mit Power Query) und den Inhalt dieses Ordners hochladen.
2. *Settings → Actions → General → Workflow permissions:* **Read and write permissions** aktivieren.
3. Optional *Settings → Secrets → Actions:* `JAO_API_TOKEN` (nur falls JAO künftig ein Token verlangt).
4. *Actions → JAO-Auslastung abrufen → Run workflow* mit Start- und Enddatum **eines** Tages (z. B. gestern).
5. Protokoll lesen. Beim ersten Lauf prüfen: JSON-Feldnamen (`fmax`, `fref`, `tso`, `cneName`, `cneEic`, `presolved`, `imax`), Antwort des Monitoring-Endpunkts, `totalRows`.
   Der Client verträgt Groß-/Kleinschreibung und unbekannte Monitoring-Antworten, ist aber nur gegen einen Schein-Server getestet. Filter-Syntax: `https://publicationtool.jao.eu/core/api` (Request-Tab).
6. *Actions → Excel-Modell bauen* starten, Artefakt herunterladen (zwei Dateien), `Kapazitaetsmodell_DE_V3_aktualisierbar.xlsx` öffnen, Inhalt aktivieren, *Daten → Alle aktualisieren*.

## Qualitätssicherung

* `python -m pytest -q` (oder `python tests/run_all.py`): Normalisierung, Aggregation, Zuordnung, Zeitfenster, Ablage, JAO-Client, Lageschätzung, Adressänderung.
* `python scripts/validate_sgm.py`: Plausibilität der Netzstammdaten.
* `python scripts/check_model.py <berechnete Datei>`: rechnet Umspannwerke, Leitungen und Trafos unabhängig in pandas nach und vergleicht mit Excel
  (alle 497 Umspannwerke in 13 Spalten stimmen mit den Testdaten vom 16.11.2026 überein). Die Statuszahlen aus Modell V2 werden reproduziert.
* Die Office-XML-Teile (Verbindungen, Abfragetabellen, Tabellen, Diagramme, Blätter) wurden gegen das Schema ISO/IEC 29500 geprüft.

## Datenmodell

`data/interim/daily/JJJJ-MM.csv` (eine Zeile je Netzelement und Tag): `Tag, Schluessel_JAO, UENB, Netzelement, Namen, EIC, Fmax_MW, Stunden, Auslastung_Summe, Auslastung_Max,
Stunden_hoch, Stunden_Ueberlast, Stunden_bindend`. `data/processed/auslastung_monatlich.csv` (Lieferung an Excel): dieselben Kennzahlen je Monat und **Excel-Schlüssel**
(EIC, ersatzweise `NAME|<Name>`), plus `Zuordnung`. `umspannwerke_koordinaten.csv`: `Umspannwerk, Breitengrad, Laengengrad, Genauigkeit, Quelle`.

Regeln: Auslastung je Zeile = |Fref| / Fmax; je Stunde zählt der **ungünstigste** Fall eines Elements (beide Richtungen, alle Ausfallfälle). Namen werden bereinigt
(die JAO-Rohdaten haben Leerzeichen am Ende). Zeilen mit Imax = 9999 und technische Randbedingungen werden laut JAO-Handbuch ausgeschlossen.

## Aus dem JAO-Handbuch v2.0 und zu gridAhead

* Umgesetzt: Fallback-Tage (`Spanning/DFP`) werden protokolliert und nicht erneut abgerufen, Seitennamen wie "Final Computation" werden erkannt, optionales Token (`JAO_API_TOKEN`),
  Hilfszeilen ausgeschlossen, **bindende Stunden** (`Presolved`) mitgezählt. Das Handbuch zeigt noch das ältere URL-Muster `/core/api/core/<Datensatz>/index?date=…`; umgesetzt ist das
  Muster des Leitfadens v1.0 (09/2026): `/core/api/data/<Datensatz>?FromUtc=…&ToUtc=…`.
* **Bindend ist nicht gleich überlastet:** Im Testlauf sind 19 Elemente bindend, und nur eines der drei Elemente mit Auslastung ab 100 % ist bindend. Entscheidung für das Team:
  Soll "bindend" die Auslastung als Engpass-Kriterium ergänzen oder ersetzen?
* **gridAhead (Reiner Lemoine Institut, Goal100; Start 17.11.2026)** will die Netzausbaupläne der Verteilnetzbetreiber geokodieren und offen veröffentlichen. Stand der Projekt-Doku:
  nur Pipelines, **noch keine Daten**. Quelle dahinter: VNBdigital (2. Runde seit Ende Oktober 2026, über 33.000 Maßnahmen, ohne Koordinaten, Lizenz meist ungeklärt). Siehe `docs/DATENQUELLEN.md` (Q14, Q26).

## Roadmap bis Februar (Vorschlag)

| Zeitraum | Ziel |
|---|---|
| bis Ende November | Repository anlegen, erster echter JAO-Lauf mit einem Tag, Feldnamen prüfen, Excel-Verbindungen ausprobieren, JAO wegen Historie ansprechen |
| Dezember | Tagesbetrieb über mindestens vier Wochen, Mindeststunden auf 168 setzen, Alias-Liste prüfen, Entscheidung zu "bindend", Prüfbericht der Koordinaten fachlich durchsehen |
| Januar | ÜNB-Netzanschlusskarten als Bandbreiten erfassen, SGM-Einlesen aus der Roh-Excel fertigstellen, exakte Koordinaten aus OpenStreetMap, Prototyp MaStR-Belegung |
| Februar | Qualitätsprüfung, Hochspannung als Struktur mit ersten Daten, Standort-ID statt Namen als Schlüssel, Betriebshandbuch, Übergabe |

## Offene Punkte

* Einlesen der SGM-Roh-Excel (`scripts/ingest_sgm.py`) fehlt: die Datei `202603_Core_Static_Grid_Model_External.xlsx` lag nicht vor.
* Mögliche Dubletten (10 Zeilen), Stromkreise ohne Spannung (1) und gleichnamige Umspannwerke sind markiert, aber nicht bereinigt.
* Nicht in Excel getestet: eingebaute Verbindungen, Power-Query-Code, Darstellung der Netzkarte. Alles lief bisher in LibreOffice und gegen einen Schein-Server, nicht gegen die echte JAO-API.

## Hinweise zu Nutzung und Weitergabe

Die verwendeten Daten sind öffentlich zugänglich, tragen aber unterschiedliche Nutzungsbedingungen (siehe `docs/DATENQUELLEN.md`). Ein öffentliches Repository veröffentlicht auch die abgeleiteten CSV-Dateien:
vorher mit der Rechtsabteilung klären. Die Netzausbaupläne der VNB stehen laut gridAhead überwiegend ohne Lizenz online, ihre Weitergabe ist nicht gesichert (nicht ins öffentliche Repository legen).
MaStR-Daten verlangen die Quellenangabe (DL-DE-BY-2.0), SMARD-Daten "Bundesnetzagentur | SMARD.de" (CC BY 4.0), OpenStreetMap-Daten "© OpenStreetMap-Mitwirkende" (ODbL).
