# Einrichtung in einfachen Schritten

## Was ist ein Repository, und warum braucht das Modell eines?

Ein **Repository** ist ein Ordner im Internet (bei GitHub), in dem zwei Dinge liegen:

1. die Programme, die täglich die JAO-Daten holen und auf das Wesentliche eindampfen (aus rund 165.000 Zeilen je Tag werden wenige kleine Tabellen), und
2. diese kleinen Tabellen als CSV-Dateien.

Excel kann die JAO-Rohdaten nicht selbst verarbeiten (zu groß, besonderes Format). Excel holt sich bei *Daten → Alle aktualisieren* deshalb nur die fertigen CSV-Dateien aus dem Repository.
**Ohne Repository gibt es keine Adresse, von der Excel Daten laden könnte.** Die Excel-Datei lässt sich trotzdem öffnen, zeigt dann aber nur die Beispieldaten.

## Einmalige Einrichtung (etwa 20 Minuten, dafür genügt ein Kollege mit GitHub-Zugang)

1. Auf **github.com** anmelden (Firmenkonto oder neues Konto) und oben rechts *+ → New repository* wählen. Name z. B. `kapazitaetsmodell`, Sichtbarkeit **Public**, dann *Create repository*.
   *Public* ist nötig, weil die eingebauten Excel-Verbindungen sich nicht anmelden können. Vorher klären, dass abgeleitete JAO-Daten und der Programmcode öffentlich sein dürfen.
2. `kapazitaetsmodell-upload.zip` ins Repository bringen. Darin liegen direkt die Ordner `.github`, `config`, `data`, `docs`, `excel`, `scripts`, `src`, `tests` und die Dateien `README.md`, `requirements.txt`, `.gitignore`. Zwei Wege:
   * **Variante A, im Browser:** ZIP entpacken (Rechtsklick, *Alle extrahieren*), den entpackten Ordner öffnen, mit `Strg + A` alles markieren und in das Upload-Feld von *Add file → Upload files* ziehen (`https://github.com/<Konto>/<Repository>/upload/main`), dann *Commit changes*. Nur Ziehen funktioniert für Ordner, der Knopf *choose your files* kann keine Ordner auswählen.
   * **Variante B, ohne Ziehen von Ordnern (am sichersten):** Auf der Repository-Seite *Create a codespace* wählen. Im Editor die ZIP-Datei in die linke Dateiliste ziehen, unter *Terminal → New Terminal* den Befehl `unzip -o kapazitaetsmodell-upload.zip && rm kapazitaetsmodell-upload.zip` ausführen, links im Reiter *Source Control* eine Beschreibung eintragen, *Commit* und danach *Sync Changes* klicken.
   Kontrolle: Auf der Startseite des Repositories müssen alle Ordner stehen, auch `.github`.
3. *Settings → Actions → General → Workflow permissions* → **Read and write permissions** → *Save*.
4. Reiter **Actions** → links *JAO-Auslastung abrufen* → *Run workflow*. Beim ersten Mal Start- und Enddatum auf **gestern** setzen (Format `2026-11-17`), damit nur ein Tag geholt wird. Nach etwa einer Minute erscheint ein grüner Haken.
   Wird er rot: Lauf öffnen, das Protokoll kopieren und an das Projektteam schicken (die Verbindung zur JAO-Schnittstelle ist bisher nur gegen einen Testserver geprüft).
5. Reiter **Actions** → *Excel-Modell bauen* → *Run workflow*. Danach den Lauf öffnen und unten im Abschnitt *Artifacts* **Kapazitaetsmodell** herunterladen und entpacken. Darin liegen zwei Dateien mit der Adresse **dieses** Repositorys:
   * `Kapazitaetsmodell_DE_V3_aktualisierbar.xlsx` (mit Datenverbindungen)
   * `Kapazitaetsmodell_DE_V3.xlsx` (ohne Verbindungen, Momentaufnahme)
6. Die aktualisierbare Datei in Excel öffnen, *Bearbeitung aktivieren*, *Inhalt aktivieren*, dann **Daten → Alle aktualisieren**. Im Blatt *Datenstand* sollten alle Zeilen auf OK stehen.

## Danach

* Täglich holt GitHub die neuen Daten von selbst. In Excel genügt *Daten → Alle aktualisieren*.
* Soll das Modell beim Öffnen automatisch aktualisieren: *Daten → Abfragen und Verbindungen → Verbindung → Eigenschaften → Daten beim Öffnen der Datei aktualisieren*.
* Neue Netzmodell-Release oder neue Modellversion: Aktion *Excel-Modell bauen* erneut ausführen.

## Wenn etwas nicht klappt

| Problem | Maßnahme |
|---|---|
| Excel meldet die Datei als beschädigt | `Kapazitaetsmodell_DE_V3.xlsx` (ohne Verbindungen) öffnen und die Daten per Power Query einspielen (`docs/POWER_QUERY.md`). Bitte die Meldung im Wortlaut an das Projektteam schicken |
| Beim Aktualisieren Fehler 404 oder "Datei nicht gefunden" | Repository ist nicht öffentlich, oder die Adresse stimmt nicht (Blatt *Datenstand*, Zeile *Datenquelle*) |
| Datenstand zeigt bei *Datenquelle* "PRÜFEN" | In der Datei steht noch der Platzhalter ORG/REPO: Datei über die Aktion *Excel-Modell bauen* erzeugen oder `python scripts/set_base_url.py` verwenden |
| Privates Repository gewünscht | Dann sind die eingebauten Verbindungen nicht nutzbar. Power Query mit Token einrichten (`docs/POWER_QUERY.md`) |
