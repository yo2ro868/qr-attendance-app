// 「フォルダ」にある 申請・報告シート　出勤用ID　氏名.xlsx を1ファイル1行で読む（接続のみ）。
// 出勤用ID と氏名はファイル名から取る（申請・報告シートには見せる欄がないため）。
// 列：ファイル名・出勤用ID・氏名・申請（その人の申請の表）・超勤（超過勤務等命令簿の月のシートの表「超勤日別」をつないだもの。なければ null）・
//     フレックス日・フレックス期間（フレックスタイム制のシートの表「フレックス日別」「フレックス期間」。なければ null）・
//     テレワーク報告（テレワーク実施報告の表「テレワーク報告」と、その人の所属。なければ null）・読み取り
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
    空 = #table(type table [ファイル名 = text, 出勤用ID = text, 氏名 = text, 申請 = table, 超勤 = nullable table,
        フレックス日 = nullable table, フレックス期間 = nullable table, テレワーク報告 = nullable table, 報告の所属 = text, 読み取り = text], {}),
    // 区切りは全角スペース。半角スペースや「_」、前の名前（休暇申請_1234_山田太郎.xlsx）でも読む
    区切りをそろえる = (名前 as text) => Text.Replace(Text.Replace(名前, "_", "　"), " ", "　"),
    // 申請・報告シート　1234　山田太郎.xlsx → {"1234", "山田太郎"}
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
            // テレワーク報告シートの B3（所属。自由に書く1つの欄）
            所属 = (b as table) => let s = Table.SelectRows(b, each [Kind] = "Sheet" and [Item] = "テレワーク報告")
                in if Table.IsEmpty(s) then "" else let d = s{0}[Data], 値 = (r, c) => try Text.Trim(Text.From(Record.Field(d{r}, c))) otherwise ""
                    in 値(2, "Column2"),
            表 = (名 as text) => let t = Table.SelectRows(ブック, each [Kind] = "Table" and [Item] = 名)
                in if Table.IsEmpty(t) then null else Table.Buffer(t{0}[Data]),
            // 見出しの改行（「開始」と「時刻」の間など）を除いた列名にする（古い様式の見出しと同じ名前で読める）
            改行なし = (t) => if t = null then null else Table.TransformColumnNames(t, each Text.Remove(_, {"#(lf)", "#(cr)"})),
            部分 = 名前の部分(名前)
        in
            [ファイル名 = 名前, 出勤用ID = 部分{0}, 氏名 = 部分{1}, 申請 = 申請, 超勤 = 超勤,
             フレックス日 = 表("フレックス日別"), フレックス期間 = 表("フレックス期間"), テレワーク報告 = 改行なし(表("テレワーク報告")),
             報告の所属 = 所属(ブック),
             読み取り = if 部分{0} = "" then "ファイル名から出勤用ID が読めません（申請・報告シート　出勤用ID　氏名.xlsx にしてください）" else "OK"],
    ファイル = 読み取り結果[ファイル],
    // フォルダの直下だけを読む（Folder.Files は下のフォルダも返すので、控えのフォルダのファイルを二重に数えないように）
    場所 = (p as text) => Text.Lower(Text.TrimEnd(p, "\")),
    対象 = if ファイル = null then null
        else Table.SelectRows(ファイル, each (Text.StartsWith(区切りをそろえる([Name]), "申請・報告シート　") or Text.StartsWith(区切りをそろえる([Name]), "休暇申請シート　") or Text.StartsWith(区切りをそろえる([Name]), "休暇申請　"))
            and Text.Lower([Extension]) = ".xlsx"
            and 場所([Folder Path]) = 場所(フォルダ)),
    行 = if フォルダ <> "" and ファイル = null then
            {[ファイル名 = フォルダ, 出勤用ID = "", 氏名 = "", 申請 = #table({}, {}), 超勤 = null, フレックス日 = null, フレックス期間 = null, テレワーク報告 = null, 報告の所属 = "",
              読み取り = "申請・報告シートのフォルダが見つかりません。場所を確かめてください"]}
        else if 対象 = null then {}
        else List.Transform(Table.ToRecords(対象), (f) =>
            try 読む(f[Name], f[Content])
            otherwise [ファイル名 = f[Name], 出勤用ID = "", 氏名 = "", 申請 = #table({}, {}), 超勤 = null, フレックス日 = null, フレックス期間 = null, テレワーク報告 = null, 報告の所属 = "",
                       読み取り = "このファイルを読み取れませんでした（「申請」の表の名前や見出しが変わっていないか確認してください）"]),
    // ファイルは1回だけ読む（使う側が何度もたどると、そのたびにファイルを読み直して遅くなるため）
    結果 = if List.IsEmpty(行) then 空 else Table.Buffer(Table.FromRecords(行, Value.Type(空)))
in
    結果
