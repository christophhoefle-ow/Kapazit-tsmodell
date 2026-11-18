// Power-Query-Abfrage D_Belegung – lädt belegung_mastr.csv
let
    Quelle = Csv.Document(Web.Contents(GitHubBaseUrl, [RelativePath = "belegung_mastr.csv"]),
        [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Kopf = Table.PromoteHeaders(Quelle, [PromoteAllScalars = true]),
    Leer = Table.ReplaceValue(Kopf, "", null, Replacer.ReplaceValue, Table.ColumnNames(Kopf)),
    Typen = Table.TransformColumnTypes(Leer, {{"Belegt_Einspeisung_MW", type number}, {"Belegt_Last_MW", type number}, {"Beantragt_Einspeisung_MW", type number}, {"Beantragt_Last_MW", type number}}, "en-US")
in
    Typen
