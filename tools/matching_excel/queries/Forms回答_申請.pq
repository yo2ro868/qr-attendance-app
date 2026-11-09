// 休暇申請フォームの回答の Excel を読む（接続のみ）。[パス, 表（読めなければ null）] を返す。
// 他のクエリを参照しない（Power Query は、データの読み取りと他のクエリの参照を同じクエリに置くと、設定によって止める）。
let
    結果 =
        let
            値 = try Excel.CurrentWorkbook(){[Name = "Forms申請ファイル"]}[Content]{0}[Column1] otherwise null,
            パス = if 値 = null then "" else Text.Trim(Text.Replace(Text.From(値), """", "")),
            ブック = if パス = "" then null else try Table.Buffer(Excel.Workbook(File.Contents(パス), null, true)) otherwise null,
            // 最初の表（なければ最初のシートを見出し付きで）
            表 = if ブック = null then null
                else let
                    表たち = Table.SelectRows(ブック, each [Kind] = "Table"),
                    シート = Table.SelectRows(ブック, each [Kind] = "Sheet")
                in
                    try (if not Table.IsEmpty(表たち) then 表たち{0}[Data] else Table.PromoteHeaders(シート{0}[Data], [PromoteAllScalars = true])) otherwise null
        in
            [パス = パス, 表 = 表]
in
    結果
