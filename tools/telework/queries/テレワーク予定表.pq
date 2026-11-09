// 「テレワークフォルダ」にある テレワーク予定表　所属名.xlsx の「予定」の表を、予定1件1行で読む（接続のみ）。
// 所属はファイル名から取る。読めないファイルやフォルダは、読み取りの列に理由を書いた1行にする。
// 列：所属・出勤用ID・氏名・日付・区分・開始時刻・終了時刻・備考・読み取り・ファイル名
let
    // 設定のフォルダとその中のファイルを、他のステップを参照せず1つのステップで読む
    // （Power Query は、データの読み取りと他のステップの参照を同じステップに置くと、設定によって止める）
    読み取り結果 =
        let
            値 = try Excel.CurrentWorkbook(){[Name = "テレワークフォルダ"]}[Content]{0}[Column1] otherwise null,
            場所 = if 値 = null then "" else Text.Trim(Text.Replace(Text.From(値), """", ""))   // ［パスとしてコピー］で付く " を取る
        in
            [フォルダ = 場所, ファイル = if 場所 = "" then null else try Table.Buffer(Folder.Files(場所)) otherwise null],
    フォルダ = 読み取り結果[フォルダ],
    型 = type table [所属 = text, 出勤用ID = text, 氏名 = text, 日付 = nullable date, 区分 = text,
        開始時刻 = nullable time, 終了時刻 = nullable time, 備考 = text, 読み取り = text, ファイル名 = text],
    文字 = (v) => if v = null then "" else Text.Trim(Text.From(v)),
    正規ID = (v) => let t = 文字(v) in if t = "" then "" else try Text.From(Number.From(t)) otherwise t,
    日付にする = (v) => if v = null then null else if v is date then v else if v is datetime then DateTime.Date(v)
        else if v is number then Date.From(v) else try Date.From(Text.Trim(Text.From(v))) otherwise null,
    時刻にする = (v) => if v = null then null else if v is time then v else if v is datetime then DateTime.Time(v)
        else if v is duration then Time.From(v) else if v is number then Time.From(Number.Mod(v, 1))
        else try Time.From(Text.Trim(Text.From(v))) otherwise null,
    // 区切りは全角スペース。半角スペースや「_」でも読む
    区切りをそろえる = (名前 as text) => Text.Replace(Text.Replace(名前, "_", "　"), " ", "　"),
    所属名 = (名前 as text) =>
        let 部分 = List.Select(Text.Split(区切りをそろえる(Text.BeforeDelimiter(名前, ".", {0, RelativePosition.FromEnd})), "　"), each _ <> "")
        in if List.Count(部分) > 1 then Text.Combine(List.Skip(部分, 1), "　") else "",
    問題 = (名前 as text, 理由 as text) => [所属 = "", 出勤用ID = "", 氏名 = "", 日付 = null, 区分 = "",
        開始時刻 = null, 終了時刻 = null, 備考 = "", 読み取り = 理由, ファイル名 = 名前],
    読む = (名前 as text, 中身 as binary) as list =>
        let
            表 = Excel.Workbook(中身, null, true){[Item = "予定", Kind = "Table"]}[Data],
            所属 = 所属名(名前),
            行 = List.Select(Table.ToRecords(表), (r) =>
                List.AnyTrue(List.Transform({"出勤用ID", "日付", "区分", "開始時刻", "終了時刻", "備考"},
                    each 文字(Record.FieldOrDefault(r, _, null)) <> "")))
        in
            List.Transform(行, (r) => [所属 = 所属, 出勤用ID = 正規ID(Record.FieldOrDefault(r, "出勤用ID", null)),
                氏名 = 文字(Record.FieldOrDefault(r, "氏名（確認用）", null)),
                日付 = 日付にする(Record.FieldOrDefault(r, "日付", null)), 区分 = 文字(Record.FieldOrDefault(r, "区分", null)),
                開始時刻 = 時刻にする(Record.FieldOrDefault(r, "開始時刻", null)), 終了時刻 = 時刻にする(Record.FieldOrDefault(r, "終了時刻", null)),
                備考 = 文字(Record.FieldOrDefault(r, "備考", null)),
                読み取り = if 所属 = "" then "ファイル名から所属が読めません（テレワーク予定表　所属名.xlsx にしてください）" else "OK",
                ファイル名 = 名前]),
    ファイル = 読み取り結果[ファイル],
    // フォルダの直下だけを読む（控えのフォルダのファイルを二重に数えないように）
    場所 = (p as text) => Text.Lower(Text.TrimEnd(p, "\")),
    対象 = if ファイル = null then null
        else Table.SelectRows(ファイル, each Text.StartsWith(区切りをそろえる([Name]), "テレワーク予定表　")
            and Text.Lower([Extension]) = ".xlsx"
            and 場所([Folder Path]) = 場所(フォルダ)),
    行 = if フォルダ <> "" and ファイル = null then {問題(フォルダ, "テレワーク予定表のフォルダが見つかりません。場所を確かめてください")}
        else if 対象 = null then {}
        else List.Combine(List.Transform(Table.ToRecords(対象), (f) =>
            try 読む(f[Name], f[Content])
            otherwise {問題(f[Name], "このファイルを読み取れませんでした（「予定」の表の名前や見出しが変わっていないか確認してください）")})),
    結果 = if List.IsEmpty(行) then #table(型, {}) else Table.FromRecords(行, 型)
in
    結果
