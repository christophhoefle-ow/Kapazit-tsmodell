# Power Query: Alternative für ein privates Repository

Die Datei `Kapazitaetsmodell_DE_V3_aktualisierbar.xlsx` enthält bereits Datenverbindungen, die nur mit einem **öffentlich lesbaren** Repository funktionieren.
Bei einem privaten Repository (oder wenn Excel die eingebauten Verbindungen nicht annimmt) richtest du Power Query ein: `Kapazitaetsmodell_DE_V3.xlsx` verwenden.
Der M-Code liegt in `excel/power_query/*.m` und im Modell im Blatt *Automatisierung* (eine Zeile je Zelle).

1. Parameter `GitHubBaseUrl` anlegen (Text): `https://raw.githubusercontent.com/<ORG>/<REPO>/main/data/processed/`
2. Je Datei eine leere Abfrage anlegen, Erweiterten Editor öffnen, M-Code einfügen, Abfrage wie die Datei im Blatt benennen:
   `D_SGM_Leitungen`, `D_SGM_Trafos`, `D_SGM_Umspannwerke`, `D_Auslastung`, `D_Netzanschluss`, `D_Belegung`, `D_Koordinaten`, `D_Meta`.
3. Beispieldaten im jeweiligen D_-Blatt ab Zeile 4 löschen, dann *Laden in … → Tabelle → Vorhandenes Arbeitsblatt → A4*.
4. Danach genügt *Daten → Alle aktualisieren*. Blatt *Datenstand*: alle Kontrollen sollten auf OK stehen.

Zahlen werden mit der Kultur `en-US` gelesen (Dezimalpunkt). Leere Felder werden vor der Typumwandlung zu `null`.
Privates Repository: nach `RelativePath` ein Feld `Headers = [Authorization = "token " & GitHubToken]` ergänzen (Token nie in der weitergegebenen Datei speichern).
**Nicht in Excel getestet.**
