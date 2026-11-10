// 「フォルダ」にある 休暇申請シート　出勤用ID　氏名.xlsx を1ファイル1行で読む（接続のみ）。
// 出勤用ID と氏名はファイル名から取る（休暇申請シートには見せる欄がないため）。
// 列：ファイル名・出勤用ID・氏名・申請（その人の申請の表）・超勤（超過勤務等命令簿の月のシートの表「超勤日別」をつないだもの。なければ null）・読み取り
let
    // 設定のフォルダとその中のファイルを、他のステップを参照せず1つのステップで読む
    // （Power Query は、データの読み取りと他のステップの参照を同じステップに置くと、設定によって止める）
    読み取り結果 =
        let
            値 = try Excel.CurrentWorkbook(){[Name = "フォルダ"]}[Content]{0}[Column1] otherwise null,
            場所 = if 値 = null then "" else Text.Trim(Text.Replace(Text.From(値), """", ""))   // ［パスとしてコピー］で付く " を取る
        in
            // Folder.Files は読むまでエラーにならないので、ここで読んでおく
            [フォルダ = 場所, ファイル = if 場所 = "" then null else try Table.Buffer(Folder.Files(場所)) otherwise null],
    フォルダ = 読み取り結果[フォルダ],
    空 = #table(type table [ファイル名 = text, 出勤用ID = text, 氏名 = text, 申請 = table, 超勤 = nullable table, 読み取り = text], {}),
    // 区切りは全角スペース。半角スペースや「_」、前の名前（休暇申請_1234_山田太郎.xlsx）でも読む
    区切りをそろえる = (名前 as text) => Text.Replace(Text.Replace(名前, "_", "　"), " ", "　"),
    // 休暇申請シート　1234　山田太郎.xlsx → {"1234", "山田太郎"}
    名前の部分 = (名前 as text) =>
        let 部分 = List.Select(Text.Split(区切りをそろえる(Text.BeforeDelimiter(名前, ".", {0, RelativePosition.FromEnd})), "　"), each _ <> "")
        in {if List.Count(部分) > 1 then 部分{1} else "", if List.Count(部分) > 2 then 部分{2} else ""},
    読む = (名前 as text, 中身 as binary) as record =>
        let
            ブック = Excel.Workbook(中身, null, true),
            申請 = Table.Buffer(ブック{[Item = "申請", Kind = "Table"]}[Data]),
            // 月のシートを写すと表の名前は 超勤日別2・超勤日別3… になる（超勤ひな形 の表は年月が空なので、使う側で除く）
            超勤の表 = Table.SelectRows(ブック, each [Kind] = "Table" and Text.StartsWith([Item], "超勤日別")),
            超勤 = if Table.IsEmpty(超勤の表) then null else Table.Buffer(Table.Combine(超勤の表[Data])),
            部分 = 名前の部分(名前)
        in
            [ファイル名 = 名前, 出勤用ID = 部分{0}, 氏名 = 部分{1}, 申請 = 申請, 超勤 = 超勤,
             読み取り = if 部分{0} = "" then "ファイル名から出勤用ID が読めません（休暇申請シート　出勤用ID　氏名.xlsx にしてください）" else "OK"],
    ファイル = 読み取り結果[ファイル],
    // フォルダの直下だけを読む（Folder.Files は下のフォルダも返すので、控えのフォルダのファイルを二重に数えないように）
    場所 = (p as text) => Text.Lower(Text.TrimEnd(p, "\")),
    対象 = if ファイル = null then null
        else Table.SelectRows(ファイル, each (Text.StartsWith(区切りをそろえる([Name]), "休暇申請シート　") or Text.StartsWith(区切りをそろえる([Name]), "休暇申請　"))
            and Text.Lower([Extension]) = ".xlsx"
            and 場所([Folder Path]) = 場所(フォルダ)),
    行 = if フォルダ <> "" and ファイル = null then
            {[ファイル名 = フォルダ, 出勤用ID = "", 氏名 = "", 申請 = #table({}, {}), 超勤 = null,
              読み取り = "休暇申請シートのフォルダが見つかりません。場所を確かめてください"]}
        else if 対象 = null then {}
        else List.Transform(Table.ToRecords(対象), (f) =>
            try 読む(f[Name], f[Content])
            otherwise [ファイル名 = f[Name], 出勤用ID = "", 氏名 = "", 申請 = #table({}, {}), 超勤 = null,
                       読み取り = "このファイルを読み取れませんでした（「申請」の表の名前や見出しが変わっていないか確認してください）"]),
    // ファイルは1回だけ読む（使う側が何度もたどると、そのたびにファイルを読み直して遅くなるため）
    結果 = if List.IsEmpty(行) then 空 else Table.Buffer(Table.FromRecords(行, Value.Type(空)))
in
    結果
