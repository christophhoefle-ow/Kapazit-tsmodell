// Power-Query-Abfrage D_SGM_Trafos – lädt sgm_trafos.csv
let
    Quelle = Csv.Document(Web.Contents(GitHubBaseUrl, [RelativePath = "sgm_trafos.csv"]),
        [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Kopf = Table.PromoteHeaders(Quelle, [PromoteAllScalars = true]),
    Leer = Table.ReplaceValue(Kopf, "", null, Replacer.ReplaceValue, Table.ColumnNames(Kopf)),
    Typen = Table.TransformColumnTypes(Leer, {{"Imax_min_A", type number}, {"Imax_max_A", type number}, {"Imax_fix_A", type number}, {"U_primaer_kV", type number}, {"U_sekundaer_kV", type number}, {"R_Ohm", type number}, {"X_Ohm", type number}, {"B_uS", type number}, {"G_uS", type number}, {"Winkel_Grad", type number}, {"du_Phase_Prozent", type number}, {"du_Winkel_Prozent", type number}}, "en-US")
in
    Typen
