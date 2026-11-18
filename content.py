"""Texte der erklärenden Blätter. Hier ändern, dann Modell neu erzeugen."""

TITLE = "Kapazitätsmodell Deutschland – Höchst- und Hochspannung (Version 3)"
SUBTITLE = ("Wo im deutschen Stromnetz ist noch Platz für neue Erzeuger, Speicher und Verbraucher? "
            "Ein Ampelmodell auf Basis öffentlicher Daten.")

INTRO = [
    "Strom wird dort erzeugt, wo der Wind weht, und dort gebraucht, wo Industrie und Städte sind. Das Netz verbindet "
    "beides – aber nicht überall ist noch Platz für Neues.",
    "Dieses Modell zeigt für jedes Umspannwerk (= Knotenpunkt im Netz) eine Ampel: Wie voll ist das Netz dort, und – "
    "soweit öffentlich bekannt – wie viel Anschlusskapazität ist noch frei?",
    "Es nutzt ausschließlich öffentliche Daten (vor allem JAO und die Netzbetreiber) und lässt sich über "
    "Daten → Alle aktualisieren auf den neuesten Stand bringen.",
    "Ziele laut Auftaktgespräch: A) Engpässe erkennen: Wo muss das Netz ausgebaut werden? "
    "B) Freie Anschlusskapazitäten finden: Wo ist Platz für neue Erzeuger und Verbraucher?",
]

# (Farbe, Bedeutung, Regel Netz (Formel), Regel Anschluss (Formel))
AMPEL_LEGENDE = [
    ("Rot", "ausgeschöpft / Engpass",
     '="Mindestens ein Netzelement (Leitung oder Trafo) war in ≥ "&TEXT(P_SchwelleEngpass,"0%")&" der Stunden über 100 % ausgelastet."',
     "Laut Übertragungsnetzbetreiber ist keine Kapazität frei (obere Bandgrenze = 0 MW)."),
    ("Gelb", "knapp / hoch belastet / unsicher",
     '="Mindestens ein Netzelement erreicht ≥ "&TEXT(P_SchwelleHoch,"0%")&" Auslastung."',
     '="Freie Kapazität nur teilweise gesichert (untere Bandgrenze unter "&P_FreiGruenAb&" MW)."'),
    ("Grün", "frei / unkritisch",
     '="Alle bewerteten Netzelemente liegen unter "&TEXT(P_SchwelleHoch,"0%")&" Auslastung."',
     '="Laut Übertragungsnetzbetreiber mindestens "&P_FreiGruenAb&" MW frei (untere Bandgrenze)."'),
    ("Grau", "keine Daten",
     "Für das Umspannwerk liegen keine Auslastungsdaten vor. Grau heißt NICHT 'unkritisch'.",
     "Keine Angabe der Übertragungsnetzbetreiber zur freien Kapazität."),
]
AMPEL_HINWEIS = ("Es zählt immer der ungünstigste Einzelwert: Der schwächste Punkt bestimmt die Farbe. Die Ampel zeigt, "
                 "was die öffentlichen Daten hergeben. Sie ersetzt keine Netzanschlussprüfung durch den Netzbetreiber.")

# (Blatt, Farbe, Rolle, Frage)
BLATT_KARTE = [
    ("Lies_mich", "hellblau", "Orientierung", "Worum geht es, und wie lese ich das Ergebnis?"),
    ("Modelllogik", "hellblau", "Orientierung", "Wie entsteht eine Ampel? Schritt für Schritt mit Rechenbeispiel."),
    ("Dashboard", "navy", "Ergebnis", "Wie sieht die Lage insgesamt aus?"),
    ("Umspannwerke", "navy", "Ergebnis", "Welche Ampel hat welches Umspannwerk, und warum?"),
    ("Leitungen", "blau", "Berechnung", "Wie viel kann jeder Stromkreis tragen, wie stark ist er ausgelastet?"),
    ("Trafos", "blau", "Berechnung", "Dasselbe für Transformatoren und Phasenschieber."),
    ("Karte_Daten", "blau", "Berechnung", "Hilfsblatt: Lage und Ampel je Umspannwerk für die Netzkarte im Dashboard."),
    ("Karte_Linien", "blau", "Berechnung", "Hilfsblatt: Netzlinien für die Karte (je Stromkreis eine Strecke)."),
    ("Parameter", "leuchtblau", "Steuerung", "Welche Annahmen und Schwellen gelten? Änderungen hier vornehmen."),
    ("Datenstand", "leuchtblau", "Kontrolle", "Wie aktuell sind die Daten, und stimmt alles?"),
    ("Glossar", "hellblau", "Nachschlagen", "Was bedeutet ein Begriff?"),
    ("Quellen", "hellblau", "Nachschlagen", "Woher kommen die Daten, wie oft aktualisiert, mit welchem Link?"),
    ("Automatisierung", "hellblau", "Anleitung", "Wie aktualisiere ich das Modell per GitHub und Power Query?"),
    ("D_SGM_Leitungen", "grau", "Rohdaten", "Stromkreise des Netzmodells (JAO Static Grid Model)."),
    ("D_SGM_Trafos", "grau", "Rohdaten", "Transformatoren und Phasenschieber (JAO Static Grid Model)."),
    ("D_SGM_Umspannwerke", "grau", "Rohdaten", "Liste aller Umspannwerke / Netzknoten."),
    ("D_Auslastung", "grau", "Rohdaten", "Auslastung je Netzelement und Monat (JAO Final Computation)."),
    ("D_Netzanschluss", "grau", "Rohdaten", "Freie Anschlusskapazität laut Netzanschlusskarten der ÜNB (Bandbreiten)."),
    ("D_Belegung", "grau", "Rohdaten", "Belegung und Pipeline je Umspannwerk (MaStR, Phase 2)."),
    ("D_Koordinaten", "grau", "Rohdaten", "Lage der Umspannwerke für die Netzkarte (exakt oder geschätzt, mit Angabe der Genauigkeit)."),
    ("D_Meta", "grau", "Rohdaten", "Datenstand, Schwellen und Release der Pipeline."),
]

FLOW = [
    ("① Quellen", "JAO, Netzbetreiber, Register im Internet", "CEECFF"),
    ("② GitHub", "Skript holt täglich Daten und verdichtet sie zu kleinen CSV-Dateien", "CEECFF"),
    ("③ Rohdaten", "Excel lädt die CSV in die grauen D_-Blätter (Daten → Alle aktualisieren)", "D1CEC9"),
    ("④ Berechnung", "Leitungen und Trafos: Kapazität, Auslastung, Status", "82BAFF"),
    ("⑤ Zusammenfassung", "Umspannwerke: Ampel je Knoten und Richtung", "EBE7E2"),
    ("⑥ Ergebnis", "Dashboard: Lage auf einen Blick", "EBE7E2"),
]

FARBCODE = [
    ("Blaue Schrift", "Eingegebene oder importierte Werte (keine Formel).", "blue"),
    ("Dunkle Schrift", "Formel.", "black"),
    ("Hellblauer Hintergrund", "Annahme oder Eingabefeld: hier darf geändert werden.", "yellow"),
    ("Graue Kopfzeile", "Rohdaten, nicht von Hand ändern.", "grey"),
    ("Dunkelblaue Kopfzeile", "Berechnung und Ergebnis.", "navy"),
]

AKTUALISIEREN = [
    ("Weg 1 – automatisch (empfohlen)",
     "Daten → Alle aktualisieren. Die Verbindungen zu den acht CSV-Dateien sind in diese Datei eingebaut (Adresse: Blatt Datenstand). "
     "Beim ersten Öffnen 'Bearbeitung aktivieren' und 'Inhalt aktivieren' (Datenverbindungen) wählen. Voraussetzung: Die Adresse ist kein Platzhalter "
     "(ORG/REPO) und das Repository ist öffentlich lesbar. Sonst: Weg über Power Query im Blatt Automatisierung."),
    ("Weg 2 – von Hand",
     "CSV-Datei aus dem GitHub-Ordner data/processed öffnen und den Inhalt ab Zeile 4 in das passende graue D_-Blatt "
     "kopieren (Einfügen als Werte). Die Spaltenreihenfolge muss unverändert bleiben (prüft das Blatt Datenstand)."),
    ("Weg 3 – komplett neu erzeugen",
     "Die GitHub-Aktion 'Excel-Modell bauen' erzeugt aus den aktuellen Daten eine fertige Datei zum Herunterladen, mit der Adresse "
     "deines Repositories bereits eingetragen. Dann ist in Excel nichts einzurichten."),
]

GRENZEN = [
    "Nur die Höchstspannung (380/220 kV) ist mit Daten hinterlegt. Die Hochspannung (110 kV) ist als Struktur vorbereitet, "
    "aber noch ohne öffentliche Kapazitätsdaten.",
    "Die Auslastung ist ein berechneter Referenzwert für den Folgetag (JAO), keine Messung.",
    "JAO veröffentlicht nur handelsrelevante Netzelemente. Ein Stromkreis ohne Daten ist nicht 'unkritisch', sondern 'keine Daten'.",
    "Die Übertragungsnetzbetreiber nennen freie Anschlusskapazität nur in Bandbreiten, nur für wenige Umspannwerke und nur je "
    "Verfahrensrunde. Solange D_Netzanschluss leer ist, beruht die Ampel allein auf dem Netz.",
    "Belegung durch bestehende Anlagen (MaStR) und Pipeline sind als Struktur vorhanden, aber noch nicht befüllt.",
    "Keine Lastflussrechnung (dafür fehlen öffentliche Netzdaten) und keine Wirtschaftlichkeitsbetrachtung.",
    "Die 380/110-kV-Trafos fehlen im JAO-Netzmodell: Die Schnittstelle zur Hochspannung ist noch eine Datenlücke.",
    "'Engpass' im Modell heißt: Referenzfluss ab Grenzwert. JAO kennzeichnet zusätzlich 'bindende' Elemente (Presolved), die den Stromhandel "
    "tatsächlich begrenzen. Beides deckt sich nur teilweise (Test: 1 von 3 Elementen). Die Pipeline sammelt bindende Stunden schon mit, "
    "die Aufnahme ins Modell ist eine offene Entscheidung.",
    "Die Netzkarte ist schematisch: Die Lage der Umspannwerke ist aus Leitungslängen und einigen Ankerpunkten geschätzt (Median-Abweichung rund 16 km, "
    "einzelne Knoten über 50 km, bei 34 Knoten gar keine Lage). Gleichnamige Umspannwerke verschiedener Betreiber (z. B. Buers, St. Peter) werden im "
    "Modell zu EINEM Knoten zusammengefasst: Ihre Ampeln sind mit Vorsicht zu lesen. Liste: data/manual/koordinaten_pruefbericht.csv.",
    "Für die Hochspannung sind die Netzausbaupläne der Verteilnetzbetreiber (VNBdigital, 2. Runde seit Ende Oktober 2026) und das Projekt gridAhead "
    "vorgemerkt. Heute enthalten die Pläne Maßnahmen ohne Koordinaten, und gridAhead hat noch keine Daten veröffentlicht (nur Pipelines).",
    "Namen im Netzmodell und in den JAO-Daten weichen teils ab. Nicht zugeordnete JAO-Elemente zählt das Blatt Datenstand; "
    "bestätigte Zuordnungen liegen im Repository in data/manual/alias_jao_sgm.csv.",
]

NEU_IN_V3 = [
    "Aufbau in drei Schichten: Rohdaten (grau, D_-Blätter) → Berechnung (blaue Reiter) → Ergebnis (Dashboard, Umspannwerke).",
    "Leitungen und Grenzkuppelstellen stehen gemeinsam im Blatt Leitungen (Spalte 'Typ'). Das Blatt API_Auslastung ist durch D_Auslastung ersetzt.",
    "Auslastung aus Monatsdaten mit frei wählbarem Auswertungsfenster (Parameter) statt eines festen Zeitraums.",
    "Ampel je Umspannwerk, getrennt für Einspeisung und Entnahme, mit Begründung und Angabe der Datengrundlage.",
    "Trafos und Phasenschieber werden ebenfalls mit JAO-Auslastung bewertet.",
    "Zuordnung Netzmodell ↔ JAO erfolgt im Skript: erst EIC, dann Name; bei mehreren Treffern gilt der ungünstigste Wert. "
    "In V2 zählte nur die erste Fundstelle.",
    "Freie Anschlusskapazität als Bandbreite (von/bis), so wie die Netzbetreiber sie veröffentlichen.",
    "Neue Blätter: Lies_mich, Modelllogik (mit Rechenbeispiel), Glossar, Quellen, Automatisierung, Datenstand.",
    "Korrekturen gegenüber V2: Fixierung bei T361 im Blatt Leitungen entfernt, Titel angepasst, Stromkreise ohne Spannung ergeben keine "
    "Kapazität von 0 mehr, sondern bleiben leer, mögliche Dubletten werden markiert. Stromkreise mit demselben Umspannwerk an beiden "
    "Enden (8, z. B. Krümmel–Krümmel) zählen am Knoten einmal; V2 zählte sie doppelt.",
    "Unverändert: Rechenregeln für Imax, Smax und Status. Mit den Testdaten vom 16.11.2026 liefert V3 dieselben Statuszahlen wie V2.",
]

SCHRITTE = [
    ("1  Daten holen", "GitHub → D_-Blätter",
     "Ein Skript lädt täglich die JAO-Daten, bereinigt Namen, verdichtet sie je Monat und legt kleine CSV-Dateien ab. "
     "Excel lädt sie in die grauen D_-Blätter."),
    ("2  Kapazität rechnen", "Leitungen, Trafos",
     "Aus Spannung und Grenzstrom wird die maximale Leistung: Smax = √3 × U × Imax ÷ 1.000 (in MVA). "
     "Bei mehreren möglichen Grenzströmen gilt die Annahme MIN oder MAX (Parameter)."),
    ("3  Auslastung zuordnen", "Leitungen, Trafos",
     "Zu jedem Element werden die JAO-Monatswerte im gewählten Zeitfenster addiert: Stunden, Mittelwert, Maximum und "
     "Stunden mit Überlast. Die Verknüpfung läuft über den Schlüssel (EIC-Code oder Name)."),
    ("4  Status je Element", "Leitungen, Trafos",
     '="Engpass, wenn in ≥ "&TEXT(P_SchwelleEngpass,"0%")&" der Stunden Überlast herrschte. Hoch belastet ab "&TEXT(P_SchwelleHoch,"0%")&" Auslastung. Sonst unkritisch. Ohne JAO-Daten: keine Daten."'),
    ("5  Je Umspannwerk zusammenfassen", "Umspannwerke",
     "Das ungünstigste angeschlossene Element bestimmt die Netz-Ampel. Kommt eine Angabe der ÜNB zur freien Kapazität dazu, "
     "zählt je Richtung (Einspeisung, Entnahme) der ungünstigere von beiden Werten."),
    ("6  Ergebnis zeigen", "Dashboard",
     "Zählt die Ampeln, zeigt die am stärksten belasteten Umspannwerke und fasst Netz und Grenzkuppelstellen zusammen."),
]

# ---------------------------------------------------------------- Automatisierung
AUTO_SATZ = ("Ein Skript auf GitHub holt täglich die öffentlichen Daten, verdichtet sie zu kleinen CSV-Dateien und legt sie im "
             "Repository ab. Excel lädt diese Dateien mit einem Klick: Daten → Alle aktualisieren. Rohdaten gelangen nie nach Excel.")

AUTO_RHYTHMUS = [
    ("JAO Final Computation (Auslastung)", "täglich abends (nicht 09–11 Uhr)", "Monatsverdichtung; im Modell frei wählbares Fenster, Standard 3 Monate",
     "Eine Stunde allein zeigt nur 0 oder 100 % Überlast. Ab einem Quartal wird das Bild stabil, 12 Monate zeigen die Jahreszeiten."),
    ("JAO Static Grid Model (Netzstammdaten)", "monatlich prüfen, halbjährlich neu laden", "nach fachlicher Prüfung",
     "Erscheint im März und September. Eine neue Release kann Spalten oder Namen ändern, deshalb prüft ein Mensch."),
    ("SMARD / Redispatch", "monatlich", "Quartal", "Die BNetzA veröffentlicht Quartalsberichte (Phase 2)."),
    ("MaStR (Belegung)", "monatlich", "je Umspannwerk", "Großer Datenabzug (rund 1,4 GB), daher nur monatlich (Phase 2)."),
    ("ÜNB-Netzanschlusskarten", "je Verfahrensrunde", "Bandbreiten in D_Netzanschluss", "Die ÜNB aktualisieren zyklisch, nicht täglich."),
]

AUTO_EINGEBAUT = [
    "GitHub-Repository anlegen und den Inhalt des Pakets hochladen. Das Repository muss öffentlich lesbar sein (die Daten sind öffentlich; Nutzungsbedingungen der Quellen vorher prüfen). Privates Repository: Alternative weiter unten.",
    "Actions → 'JAO-Auslastung abrufen' einmal starten (Start- und Enddatum eines Tages), danach 'Excel-Modell bauen'. Die erzeugte Datei hat die Adresse deines Repositories schon eingetragen.",
    "Datei herunterladen und in Excel öffnen. In der gelben Leiste 'Bearbeitung aktivieren', dann 'Inhalt aktivieren' (Datenverbindungen) wählen.",
    "Daten → Alle aktualisieren. Die acht grauen D_-Blätter laden ihre CSV-Datei aus dem Internet, danach rechnet das Modell. Kontrolle: Blatt Datenstand, alle Zeilen 'OK'.",
    "Ab dann genügt Schritt 4. Automatisch beim Öffnen aktualisieren: Daten → Abfragen und Verbindungen → Verbindung → Eigenschaften → 'Daten beim Öffnen der Datei aktualisieren'.",
]

AUTO_SCHRITTE = [
    "GitHub-Repository öffnen und den Basis-Link der CSV-Dateien notieren: https://raw.githubusercontent.com/<ORG>/<REPO>/main/data/processed/  (privates Repository: siehe Hinweis unten).",
    "In Excel: Daten → Daten abrufen → Aus anderen Quellen → Leere Abfrage. Im Power-Query-Editor: Start → Parameter verwalten → Neuer Parameter: Name GitHubBaseUrl, Typ Text, Wert = Basis-Link aus Schritt 1 (mit Schrägstrich am Ende).",
    "Für jede der 8 Abfragen unten: Neue leere Abfrage → Start → Erweiterter Editor → vorhandenen Text ersetzen durch den M-Code → Fertig. Abfrage genau so benennen wie angegeben (D_Auslastung usw.).",
    "Vor dem Laden die Beispieldaten im jeweiligen D_-Blatt löschen (Zeile 4 bis zum Ende markieren → Zellen löschen). Die Zeilen 1 bis 3 bleiben stehen.",
    "Start → Schließen und laden in … → Tabelle → Vorhandenes Arbeitsblatt → Zelle A4 des passenden D_-Blatts (z. B. =D_Auslastung!$A$4). Der Tabellenkopf landet in Zeile 4, die Daten ab Zeile 5.",
    "Fertig. Ab jetzt reicht Daten → Alle aktualisieren. Kontrolle im Blatt Datenstand: alle Zeilen 'OK'.",
]

AUTO_HINWEISE = [
    "Privates Repository: In jeder Abfrage nach RelativePath das Feld Headers = [Authorization = \"token \" & GitHubToken] ergänzen und einen Parameter GitHubToken (fein abgestufter Token, nur Lesezugriff auf Inhalte) anlegen. Das Token niemals in der Datei speichern, die du weitergibst.",
    "Zahlen werden mit der Kultur \"en-US\" gelesen. Das verhindert den bekannten Fehler, dass ein Dezimalpunkt als Tausendertrenner gedeutet wird (Faktor 10^10 in den JAO-Testdaten).",
    "Fehlermeldung 'Formula.Firewall': In Excel unter Daten → Abfrageoptionen → Datenschutz die Einstellung 'Datenschutzstufen ignorieren' wählen oder die Quelle als 'Organisation' einstufen.",
    "Die eingebauten Verbindungen sind klassische Text-/CSV-Abfragen. Sie sind gegen das Dateiformat-Schema (ISO/IEC 29500) geprüft, aber NICHT in Excel ausprobiert. Bietet Excel beim Öffnen eine Reparatur an oder schlägt das Aktualisieren fehl, die Datei 'ohne Verbindungen' nutzen und Power Query wie unten einrichten.",
    "Klassische Textabfragen können keine Zugangsdaten mitschicken: Ein privates GitHub-Repository lässt sich damit nicht lesen (Fehler 404). Dafür ist Power Query mit Token nötig (Abschnitt unten). Der M-Code unten wurde ebenfalls nicht in Excel ausgeführt.",
    "Ohne Power Query geht es auch: CSV aus GitHub öffnen und ab Zeile 4 in das D_-Blatt einfügen (Weg 2 im Blatt Lies_mich).",
]

# M-Code je Abfrage: (Abfragename, Datei, Typen {Spalte: m-Typ})
NUM = "type number"
M_QUERIES = [
    ("D_SGM_Leitungen", "sgm_leitungen.csv",
     {c: NUM for c in ["Spannung_kV", "Imax_Periode_1_A", "Imax_Periode_2_A", "Imax_Periode_3_A", "Imax_Periode_4_A",
                       "Imax_Periode_5_A", "Imax_Periode_6_A", "Imax_fix_A", "DLR_min_A", "DLR_max_A", "R_Ohm", "X_Ohm",
                       "B_uS", "Laenge_km"]}),
    ("D_SGM_Trafos", "sgm_trafos.csv",
     {c: NUM for c in ["Imax_min_A", "Imax_max_A", "Imax_fix_A", "U_primaer_kV", "U_sekundaer_kV", "R_Ohm", "X_Ohm", "B_uS",
                       "G_uS", "Winkel_Grad", "du_Phase_Prozent", "du_Winkel_Prozent"]}),
    ("D_SGM_Umspannwerke", "sgm_umspannwerke.csv", {}),
    ("D_Auslastung", "auslastung_monatlich.csv",
     {"Monat": "type date", "Fmax_MW": NUM, "Stunden": "Int64.Type", "Auslastung_Summe": NUM, "Auslastung_Max": NUM,
      "Stunden_hoch": "Int64.Type", "Stunden_Ueberlast": "Int64.Type"}),
    ("D_Netzanschluss", "netzanschluss_uenb.csv",
     {"Einspeisung_frei_von_MW": NUM, "Einspeisung_frei_bis_MW": NUM, "Last_frei_von_MW": NUM, "Last_frei_bis_MW": NUM,
      "Verfuegbar_ab_Jahr": "Int64.Type"}),
    ("D_Belegung", "belegung_mastr.csv",
     {"Belegt_Einspeisung_MW": NUM, "Belegt_Last_MW": NUM, "Beantragt_Einspeisung_MW": NUM, "Beantragt_Last_MW": NUM}),
    ("D_Koordinaten", "umspannwerke_koordinaten.csv", {"Breitengrad": NUM, "Laengengrad": NUM}),
    ("D_Meta", "meta.csv", {}),
]


def m_code(name: str, filename: str, types: dict) -> list[str]:
    """M-Code einer Abfrage als Liste einzelner Zeilen (eine Zeile je Zelle: sauber kopierbar)."""
    lines = [
        "let",
        f'    Quelle = Csv.Document(Web.Contents(GitHubBaseUrl, [RelativePath = "{filename}"]),',
        '        [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),',
        "    Kopf = Table.PromoteHeaders(Quelle, [PromoteAllScalars = true]),",
        "    Leer = Table.ReplaceValue(Kopf, \"\", null, Replacer.ReplaceValue, Table.ColumnNames(Kopf)),",
    ]
    if types:
        pairs = ", ".join("{" + f'"{c}", {t}' + "}" for c, t in types.items())
        lines += [f"    Typen = Table.TransformColumnTypes(Leer, {{{pairs}}}, \"en-US\")", "in", "    Typen"]
    else:
        lines += ["in", "    Leer"]
    return lines
