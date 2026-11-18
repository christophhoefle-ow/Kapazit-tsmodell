// Power-Query-Abfrage D_Koordinaten – lädt umspannwerke_koordinaten.csv
let
    Quelle = Csv.Document(Web.Contents(GitHubBaseUrl, [RelativePath = "umspannwerke_koordinaten.csv"]),
        [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Kopf = Table.PromoteHeaders(Quelle, [PromoteAllScalars = true]),
    Leer = Table.ReplaceValue(Kopf, "", null, Replacer.ReplaceValue, Table.ColumnNames(Kopf)),
    Typen = Table.TransformColumnTypes(Leer, {{"Breitengrad", type number}, {"Laengengrad", type number}}, "en-US")
in
    Typen
