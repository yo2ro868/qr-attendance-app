// 出勤管理アプリ（QR）の［Excel出力］のファイル（attendance.xlsx など）を置いたフォルダ（名前 QR出勤フォルダ ＝ 出力データ貼付シートの H3）を読む（接続のみ）。
// フォルダの直下の .xlsx を1ファイル1行で：ファイル名・更新日時・行（シート「出勤記録」の表。見出しは 出勤用ID・QR設置場所・出勤日・出勤時刻・登録日時）・読み取り。
// 他のクエリを参照しない（Power Query は、データの読み取りと他のクエリの参照を同じクエリに置くと、設定によって止める）。
let
    // 設定のフォルダとその中のファイルを、他のステップを参照せず1つのステップで読む
    読み取り結果 =
        let
            値 = try Excel.CurrentWorkbook(){[Name = "QR出勤フォルダ"]}[Content]{0}[Column1] otherwise null,
            場所 = if 値 = null then "" else Text.Trim(Text.Replace(Text.From(値), """", ""))   // ［パスとしてコピー］で付く " を取る
        in
            [フォルダ = 場所, ファイル = if 場所 = "" then null else try Table.Buffer(Folder.Files(場所)) otherwise null],
    フォルダ = 読み取り結果[フォルダ],
    ファイル = 読み取り結果[ファイル],
    場所 = (p as text) => Text.Lower(Text.TrimEnd(p, "\")),
    見出し = {"出勤用ID", "QR設置場所", "出勤日", "出勤時刻", "登録日時"},
    空 = #table(type table [ファイル名 = text, 更新日時 = nullable datetime, 行 = any, 読み取り = text], {}),
    // 1ファイル：シート「出勤記録」（なければ最初のシート）を、1行目を見出しにして読む
    読む = (name as text, content as binary) as table =>
        let
            ブック = Excel.Workbook(content, null, true),
            シート = Table.SelectRows(ブック, each [Kind] = "Sheet"),
            対象 = let s = Table.SelectRows(シート, each [Item] = "出勤記録") in if Table.IsEmpty(s) then シート else s,
            表 = Table.PromoteHeaders(対象{0}[Data], [PromoteAllScalars = true]),
            足りない = List.Difference(見出し, Table.ColumnNames(表))
        in
            if not List.IsEmpty(足りない) then error ("見出しが違います（" & Text.Combine(足りない, "・") & " がありません）")
            else Table.SelectColumns(表, 見出し),
    対象 = if ファイル = null then null
        else Table.SelectRows(ファイル, each Text.Lower([Extension]) = ".xlsx" and not Text.StartsWith([Name], "~$")
            and 場所([Folder Path]) = 場所(フォルダ)),
    行 = if フォルダ <> "" and ファイル = null then
            {[ファイル名 = フォルダ, 更新日時 = null, 行 = null, 読み取り = "フォルダが見つかりません。場所を確かめてください"]}
        else if 対象 = null then {}
        else List.Transform(Table.ToRecords(対象), (f) =>
            let r = try 読む(f[Name], f[Content])
            in [ファイル名 = f[Name], 更新日時 = f[Date modified],
                行 = if r[HasError] then null else Table.Buffer(r[Value]),
                読み取り = if r[HasError] then "読み取れませんでした：" & r[Error][Message] else "OK"])
in
    if List.IsEmpty(行) then 空 else Table.FromRecords(行, Value.Type(空))
