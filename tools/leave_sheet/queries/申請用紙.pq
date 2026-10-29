// 設定シートの「フォルダ」にある 休暇申請_*.xlsx を1ファイル1行で読む（接続のみ）。
// 列：ファイル名・出勤用ID・氏名・所属長・年・導入前の使用（分）・申請（その人の申請の表）・読み取り
let
    設定値 = (名前 as text) =>
        let v = try Excel.CurrentWorkbook(){[Name = 名前]}[Content]{0}[Column1] otherwise null
        in if v = null then "" else Text.Trim(Text.From(v)),
    フォルダ = 設定値("フォルダ"),
    空 = #table(type table [ファイル名 = text, 出勤用ID = text, 氏名 = text, 所属長 = text, 年 = number,
        導入前の使用 = number, 申請 = table, 読み取り = text], {}),
    数 = (v) => if v = null then 0 else try Number.From(v) otherwise 0,
    文字 = (v) => if v = null then "" else Text.Trim(Text.From(v)),
    // 時間のセル（[h]:mm）は time・duration・数のどれで来ても「日」の数にする
    日の数 = (v) => if v = null then 0
        else if v is time then Number.From(v)
        else if v is duration then Duration.TotalDays(v)
        else 数(v),
    読む = (名前 as text, 中身 as binary) as record =>
        let
            ブック = Excel.Workbook(中身, null, true),
            基本 = ブック{[Item = "基本", Kind = "Table"]}[Data]{0},
            申請 = ブック{[Item = "申請", Kind = "Table"]}[Data]
        in
            [ファイル名 = 名前, 出勤用ID = 文字(基本[出勤用ID]), 氏名 = 文字(基本[氏名]), 所属長 = 文字(基本[所属長]),
             年 = 数(基本[年]),
             導入前の使用 = Number.Round(数(基本[#"導入前の使用（日）"]) * 465 + 日の数(基本[#"導入前の使用（時間）"]) * 1440),
             申請 = 申請, 読み取り = "OK"],
    ファイル = if フォルダ = "" then null else try Folder.Files(フォルダ) otherwise null,
    対象 = if ファイル = null then null
        else Table.SelectRows(ファイル, each Text.StartsWith([Name], "休暇申請_") and Text.Lower([Extension]) = ".xlsx"),
    行 = if 対象 = null then {}
        else List.Transform(Table.ToRecords(対象), (f) =>
            try 読む(f[Name], f[Content])
            otherwise [ファイル名 = f[Name], 出勤用ID = "", 氏名 = "", 所属長 = "", 年 = 0, 導入前の使用 = 0,
                       申請 = #table({}, {}), 読み取り = "このファイルを読み取れませんでした（表の名前や見出しが変わっていないか確認してください）"]),
    結果 = if List.IsEmpty(行) then 空 else Table.FromRecords(行, Value.Type(空))
in
    結果
