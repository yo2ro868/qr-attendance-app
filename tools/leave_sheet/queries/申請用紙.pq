// 「フォルダ」にある 休暇申請_出勤用ID_氏名.xlsx を1ファイル1行で読む（接続のみ）。
// 出勤用ID と氏名はファイル名から取る（申請用紙には見せる欄がないため）。
// 列：ファイル名・出勤用ID・氏名・申請（その人の申請の表）・読み取り
let
    設定値 = (名前 as text) =>
        let v = try Excel.CurrentWorkbook(){[Name = 名前]}[Content]{0}[Column1] otherwise null
        in if v = null then "" else Text.Trim(Text.From(v)),
    フォルダ = 設定値("フォルダ"),
    空 = #table(type table [ファイル名 = text, 出勤用ID = text, 氏名 = text, 申請 = table, 読み取り = text], {}),
    // 休暇申請_1234_山田太郎.xlsx → {"1234", "山田太郎"}
    名前の部分 = (名前 as text) =>
        let 部分 = Text.Split(Text.BeforeDelimiter(名前, ".", {0, RelativePosition.FromEnd}), "_")
        in {if List.Count(部分) > 1 then 部分{1} else "", if List.Count(部分) > 2 then 部分{2} else ""},
    読む = (名前 as text, 中身 as binary) as record =>
        let
            ブック = Excel.Workbook(中身, null, true),
            申請 = ブック{[Item = "申請", Kind = "Table"]}[Data],
            部分 = 名前の部分(名前)
        in
            [ファイル名 = 名前, 出勤用ID = 部分{0}, 氏名 = 部分{1}, 申請 = 申請,
             読み取り = if 部分{0} = "" then "ファイル名から出勤用ID が読めません（休暇申請_出勤用ID_氏名.xlsx にしてください）" else "OK"],
    ファイル = if フォルダ = "" then null else try Folder.Files(フォルダ) otherwise null,
    対象 = if ファイル = null then null
        else Table.SelectRows(ファイル, each Text.StartsWith([Name], "休暇申請_") and Text.Lower([Extension]) = ".xlsx"),
    行 = if 対象 = null then {}
        else List.Transform(Table.ToRecords(対象), (f) =>
            try 読む(f[Name], f[Content])
            otherwise [ファイル名 = f[Name], 出勤用ID = "", 氏名 = "", 申請 = #table({}, {}),
                       読み取り = "このファイルを読み取れませんでした（「申請」の表の名前や見出しが変わっていないか確認してください）"]),
    結果 = if List.IsEmpty(行) then 空 else Table.FromRecords(行, Value.Type(空))
in
    結果
