// Power-Query-Abfrage D_Netzanschluss – lädt netzanschluss_uenb.csv
let
    Quelle = Csv.Document(Web.Contents(GitHubBaseUrl, [RelativePath = "netzanschluss_uenb.csv"]),
        [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Kopf = Table.PromoteHeaders(Quelle, [PromoteAllScalars = true]),
    Leer = Table.ReplaceValue(Kopf, "", null, Replacer.ReplaceValue, Table.ColumnNames(Kopf)),
    Typen = Table.TransformColumnTypes(Leer, {{"Einspeisung_frei_von_MW", type number}, {"Einspeisung_frei_bis_MW", type number}, {"Last_frei_von_MW", type number}, {"Last_frei_bis_MW", type number}, {"Verfuegbar_ab_Jahr", Int64.Type}}, "en-US")
in
    Typen
