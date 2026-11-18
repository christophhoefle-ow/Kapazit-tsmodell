// Power-Query-Abfrage D_Auslastung – lädt auslastung_monatlich.csv
let
    Quelle = Csv.Document(Web.Contents(GitHubBaseUrl, [RelativePath = "auslastung_monatlich.csv"]),
        [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Kopf = Table.PromoteHeaders(Quelle, [PromoteAllScalars = true]),
    Leer = Table.ReplaceValue(Kopf, "", null, Replacer.ReplaceValue, Table.ColumnNames(Kopf)),
    Typen = Table.TransformColumnTypes(Leer, {{"Monat", type date}, {"Fmax_MW", type number}, {"Stunden", Int64.Type}, {"Auslastung_Summe", type number}, {"Auslastung_Max", type number}, {"Stunden_hoch", Int64.Type}, {"Stunden_Ueberlast", Int64.Type}}, "en-US")
in
    Typen
