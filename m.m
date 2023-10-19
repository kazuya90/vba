let
    ソース = Csv.Document(
        File.Contents("C:\Users\user\プロジェクト\vba\テストデータ - TM-WebTools.csv"),
        [
            Delimiter = ",",
            Columns = 2,
            Encoding = 65001,
            QuoteStyle = QuoteStyle.None
        ]
    ),
    変更された型 = Table.TransformColumnTypes(ソース, {{"Column1", type text}, {"Column2", type text}}),
    昇格されたヘッダー数 = Table.PromoteHeaders(変更された型, [PromoteAllScalars = true]),
    変更された型1 = Table.TransformColumnTypes(昇格されたヘッダー数, {{"性", type text}, {"名", type text}}),
    ReplacedValue = Table.ReplaceValue(変更された型1, "〃〃", each null, Replacer.ReplaceValue, {"性", "名"}),
    FilledDown = Table.FillDown(#"ReplacedValue", {"性", "名"})
in
    FilledDown
