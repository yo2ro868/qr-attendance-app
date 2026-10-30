// 全員の申請を1件1行にする（接続のみ）。日付は date、取得時間は「分」に直す。
// 分：1日単位は 日数×465（1日＝7時間45分）、時間単位は 時間×1440
let
    数 = (v) => if v = null then 0 else try Number.From(v) otherwise 0,
    文字 = (v) => if v = null then "" else Text.Trim(Text.From(v)),
    日付 = (v) => if v = null then null
        else if v is date then v
        else if v is datetime then DateTime.Date(v)
        else if v is number then Date.From(v)
        else try Date.From(v) otherwise null,
    日の数 = (v) => if v = null then 0
        else if v is time then Number.From(v)
        else if v is duration then Duration.TotalDays(v)
        else 数(v),
    行にする = (f as record) as list =>
        let
            申請 = Table.ToRecords(f[申請]),
            書いた行 = List.Select(申請, each 文字(Record.FieldOrDefault(_, "休暇の種類")) <> "")
        in
            List.Transform(書いた行, (r) =>
                let
                    // レコードの中では同じ名前のフィールドが優先されるので、フィールドと違う名前にしておく
                    開始 = 日付(r[開始日]),
                    日数の値 = 数(r[日数]),
                    時間の値 = 日の数(r[時間])
                in
                    [ファイル名 = f[ファイル名], 出勤用ID = f[出勤用ID], 氏名 = f[氏名],
                     申請日 = 日付(r[申請日]), 休暇の種類 = 文字(r[休暇の種類]), #"種別・内容" = 文字(r[#"種別・内容"]),
                     単位 = 文字(r[単位]), 開始日 = 開始, 終了日 = 日付(r[終了日]),
                     日数 = if 日数の値 = 0 then null else 日数の値,
                     時間 = if 時間の値 = 0 then null else 時間の値,
                     時刻 = 文字(r[時刻]), 事由 = 文字(r[事由]),
                     #"承認（所属長）" = 文字(r[#"承認（所属長）"]), 承認日 = 日付(r[承認日]),
                     所属長コメント = 文字(r[所属長コメント]), 確認 = 文字(r[確認]),
                     分 = Number.Round(日数の値 * 465 + 時間の値 * 1440),
                     年 = if 開始 = null then null else Date.Year(開始),
                     月 = if 開始 = null then null else Date.Month(開始)]),
    型 = type table [ファイル名 = text, 出勤用ID = text, 氏名 = text, 申請日 = nullable date,
        休暇の種類 = text, #"種別・内容" = text, 単位 = text, 開始日 = nullable date, 終了日 = nullable date,
        日数 = nullable number, 時間 = nullable number, 時刻 = text, 事由 = text, #"承認（所属長）" = text,
        承認日 = nullable date, 所属長コメント = text, 確認 = text, 分 = number, 年 = nullable number, 月 = nullable number],
    読めた = Table.SelectRows(申請用紙, each [読み取り] = "OK"),
    行 = List.Combine(List.Transform(Table.ToRecords(読めた), 行にする))
in
    Table.FromRecords(行, 型)
