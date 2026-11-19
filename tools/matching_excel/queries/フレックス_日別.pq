// 申請・報告シートのフレックスタイム制の表（フレックス日別・フレックス期間。tools/flex）を、今の内容で1日1行にする。
// フレックスを使っている人（申告年月日を入れた単位期間がある人）の、単位期間に入る日だけ。時刻・時間は Excel の時間（1日＝1）にする。
// 「フレックスの読み込み」シートに出し、［最新を読み込む］のマクロが、所属長の承認済みになった単位期間をここから「フレックスの日別」に写す。
let
    正規ID = (id) => let t = Text.Trim(if id = null then "" else Text.From(id))
        in if t = "" then "" else try Text.From(Number.From(t)) otherwise t,
    文字 = (v) => if v = null then "" else Text.Trim(Text.From(v)),
    数 = (v) => if 文字(v) = "" then null else try Number.From(v) otherwise null,
    日付にする = (v) => if 文字(v) = "" then null else try Date.FromText(文字(v), [Format = "yyyy-MM-dd"]) otherwise null,
    時間 = (v) => let m = 数(v) in if m = null then null else m / 1440,
    ファイル = Table.SelectRows(休暇申請シート, each [フレックス日] <> null and [フレックス期間] <> null and [出勤用ID] <> ""),
    人 = (f as record) as list =>
        let
            期間 = List.Buffer(Table.ToRecords(f[フレックス期間])),
            使う = List.AnyTrue(List.Transform(期間, each 文字(Record.FieldOrDefault(_, "申告年月日")) <> "")),
            期間の情報 = (k) => let hit = List.Select(期間, each 数([期間]) = k) in if List.IsEmpty(hit) then null else hit{0},
            p値 = (p, 名) => if p = null then null else Record.FieldOrDefault(p, 名)
        in
            // 単位期間に入らない日（期間番号が空）は読まない
            if not 使う then {} else List.Transform(List.Select(Table.ToRecords(f[フレックス日]), each 数([期間]) <> null), (d) =>
                let k = 数(d[期間]), p = 期間の情報(k)
                in [出勤用ID = 正規ID(f[出勤用ID]), 氏名 = f[氏名], 所属 = 文字(p値(p, "所属")),
                    期間 = k, 期間の開始日 = 日付にする(p値(p, "開始日")), 日付 = 日付にする(d[日付]), 区分 = 文字(d[区分]),
                    勤務開始 = 時間(d[勤務開始]), 勤務終了 = 時間(d[勤務終了]), 休憩 = 時間(d[休憩]), 勤務時間 = 時間(d[勤務時間]),
                    休憩の時刻 = 文字(d[休憩の時刻]), 備考 = 文字(d[備考]), 確認 = 文字(d[確認]),
                    申告年月日 = 日付にする(p値(p, "申告年月日")), 割振年月日 = 日付にする(p値(p, "割振年月日")),
                    状態 = 文字(p値(p, "状態")), ファイル名 = f[ファイル名]]),
    行 = List.Combine(List.Transform(Table.ToRecords(ファイル), (f) => try 人(f) otherwise {})),
    型 = type table [出勤用ID = text, 氏名 = text, 所属 = text, 期間 = nullable number, 期間の開始日 = nullable date, 日付 = nullable date,
                    区分 = text, 勤務開始 = nullable number, 勤務終了 = nullable number, 休憩 = nullable number, 勤務時間 = nullable number,
                    休憩の時刻 = text, 備考 = text, 確認 = text, 申告年月日 = nullable date, 割振年月日 = nullable date,
                    状態 = text, ファイル名 = text],
    表 = if List.IsEmpty(行) then #table(型, {}) else Table.FromRecords(行, 型)
in
    Table.Sort(Table.SelectRows(表, each [日付] <> null), {{"出勤用ID", Order.Ascending}, {"日付", Order.Ascending}})
