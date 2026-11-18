// Power-Query-Abfrage D_SGM_Leitungen – lädt sgm_leitungen.csv
let
    Quelle = Csv.Document(Web.Contents(GitHubBaseUrl, [RelativePath = "sgm_leitungen.csv"]),
        [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Kopf = Table.PromoteHeaders(Quelle, [PromoteAllScalars = true]),
    Leer = Table.ReplaceValue(Kopf, "", null, Replacer.ReplaceValue, Table.ColumnNames(Kopf)),
    Typen = Table.TransformColumnTypes(Leer, {{"Spannung_kV", type number}, {"Imax_Periode_1_A", type number}, {"Imax_Periode_2_A", type number}, {"Imax_Periode_3_A", type number}, {"Imax_Periode_4_A", type number}, {"Imax_Periode_5_A", type number}, {"Imax_Periode_6_A", type number}, {"Imax_fix_A", type number}, {"DLR_min_A", type number}, {"DLR_max_A", type number}, {"R_Ohm", type number}, {"X_Ohm", type number}, {"B_uS", type number}, {"Laenge_km", type number}}, "en-US")
in
    Typen
