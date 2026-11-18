// 「フレックス」シートに出す：フレックスを使っている人の単位期間ごとの状況（今の休暇申請シートの内容）と、管理者の取り込み、注意。
// 取り込み：所属長の承認済み（状態が「承認済み」）で、まだ取り込んでいないか、割振年月日が取り込んだときより新しい（承認し直した）単位期間は
//   「取り込む」。［最新を読み込む］のマクロがその単位期間を「フレックスの日別」に写す（写した内容は、あとで申告が変わっても変わらない）。
// 注意：取り込んだ後に申告が変わったもの（承認時の内容＝取り込んだ内容を使う）、始まった（1週間以内に始まる）のに承認されていない単位期間など。
// 照合の注意：取り込んだ日ごとの行（フレックスの日別）の「注意」がある日（出勤の記録がない勤務日など）。注意の文ごとに日付を8日まで、ほかは日数。
let
    文字 = (v) => if v = null then "" else Text.Trim(Text.From(v)),
    日付にする = (v) => if v = null then null else if v is date then v else if v is datetime then DateTime.Date(v)
        else if v is number then Date.From(v) else try Date.From(v) otherwise null,
    分 = (v) => if v = null or 文字(v) = "" then "" else try Text.From(Number.Round(Number.From(v) * 1440)) otherwise "",
    今日 = Date.From(DateTime.LocalNow()),
    キー = (r as record) => 文字(r[出勤用ID]) & "|" & 文字(r[期間]) & "|" & (let d = 日付にする(r[期間の開始日]) in if d = null then "" else Date.ToText(d, "yyyy-MM-dd")),
    // 中身の比べ方：日ごとの 勤務開始・勤務終了・休憩（分）をつないだ文字
    指紋 = (rows as list) => Text.Combine(List.Transform(List.Sort(rows, (a, b) => Value.Compare(日付にする(a[日付]), 日付にする(b[日付]))),
        each (let d = 日付にする([日付]) in if d = null then "" else Date.ToText(d, "yyyyMMdd")) & ":" & 分([勤務開始]) & "-" & 分([勤務終了]) & "/" & 分([休憩])), ","),
    // 今の内容（単位期間ごと）
    今 = Table.Group(Table.AddColumn(フレックス_日別, "キー", each キー(_), type text), {"キー"}, {{"行", each Table.ToRecords(_)}}),
    // 取り込んだ内容（単位期間ごと）
    確定表 = if フレックス_確定 = null then null else Table.SelectRows(フレックス_確定, each 文字([出勤用ID]) <> ""),
    確定 = if 確定表 = null or Table.IsEmpty(確定表) then #table({"キー", "行"}, {})
        else Table.Group(Table.AddColumn(確定表, "キー", each キー(_), type text), {"キー"}, {{"行", each Table.ToRecords(_)}}),
    確定の一覧 = List.Buffer(Table.ToRecords(確定)),
    確定を引く = (k as text) => let hit = List.Select(確定の一覧, each [キー] = k) in if List.IsEmpty(hit) then null else hit{0}[行],
    一つ = (k as text, rows as list) as record =>
        let
            a = rows{0},
            日付たち = List.RemoveNulls(List.Transform(rows, each 日付にする([日付]))),
            開始 = 日付にする(a[期間の開始日]),
            終了 = if List.IsEmpty(日付たち) then null else List.Max(日付たち),
            v合計 = List.Sum(List.Transform(rows, each if [勤務時間] = null then 0 else [勤務時間])),
            v状態 = 文字(a[状態]),
            割振 = 日付にする(a[割振年月日]),
            確 = 確定を引く(k),
            確の割振 = if 確 = null then null else 日付にする(確{0}[割振年月日]),
            確の日時 = if 確 = null then null else List.Max(List.Transform(確, each [確定日時])),
            注意の日 = if 確 = null then {} else List.Sort(List.Select(確, each 文字(Record.FieldOrDefault(_, "注意", null)) <> ""),
                (x, y) => Value.Compare(日付にする(x[日付]), 日付にする(y[日付]))),
            月日 = (r) => let d = 日付にする(r[日付]) in if d = null then "" else Text.From(Date.Month(d)) & "/" & Text.From(Date.Day(d)),
            文の一覧 = List.Distinct(List.Transform(注意の日, each Text.TrimEnd(文字([注意]), "。"))),
            v照合 = Text.Combine(List.Transform(文の一覧, (msg) =>
                let ds = List.Select(注意の日, each Text.TrimEnd(文字([注意]), "。") = msg)
                in msg & "（" & Text.Combine(List.Transform(List.FirstN(ds, 8), 月日), "・")
                    & (if List.Count(ds) > 8 then " ほか " & Text.From(List.Count(ds) - 8) & " 日" else "") & "）"), "／"),
            取り込む = v状態 = "承認済み" and (確 = null or (割振 <> null and (確の割振 = null or 割振 > 確の割振))),
            始まる = 開始 <> null and 開始 <= Date.AddDays(今日, 7) and (終了 = null or 終了 >= 今日),
            v注意 = if 取り込む then ""
                else if 確 <> null and Text.StartsWith(v状態, "承認後") then "承認後に申告が変更されました。所属長の承認し直し待ちです（取り込んだ承認時の内容を使っています）。"
                else if 確 <> null and v状態 <> "承認済み" then "所属長の承認が外れています（" & v状態 & "）。取り込んだ承認時の内容を使っています。"
                else if 確 <> null and 指紋(rows) <> 指紋(確) then "取り込んだ後に申告が書き換えられています（申告年月日が変わっていません）。取り込んだ承認時の内容を使っています。"
                else if 確 = null and 始まる and v状態 = "未申告" then "単位期間が始まる（始まっている）のに申告がありません。"
                else if 確 = null and 始まる and v状態 <> "承認済み" then "単位期間が始まる（始まっている）のに承認されていません（" & v状態 & "）。"
                else ""
        in
            [出勤用ID = 文字(a[出勤用ID]), 氏名 = 文字(a[氏名]), 所属 = 文字(a[所属]), 期間 = a[期間], 開始日 = 開始, 終了日 = 終了,
             合計 = v合計, 申告年月日 = 日付にする(a[申告年月日]), 割振年月日 = 割振, 状態 = v状態,
             取り込み = if 取り込む then "取り込む" else if 確 <> null then "取り込み済み" else "",
             取り込んだ日時 = 確の日時, 注意 = v注意, 照合の注意 = v照合, ファイル名 = 文字(a[ファイル名]), 表示する = 確 <> null or v状態 <> "未申告" or 始まる],
    行 = List.Transform(Table.ToRecords(今), each 一つ([キー], [行])),
    型 = type table [出勤用ID = text, 氏名 = text, 所属 = text, 期間 = nullable number, 開始日 = nullable date, 終了日 = nullable date,
                    合計 = number, 申告年月日 = nullable date, 割振年月日 = nullable date, 状態 = text, 取り込み = text,
                    取り込んだ日時 = nullable datetime, 注意 = text, 照合の注意 = text, ファイル名 = text, 表示する = logical],
    表 = if List.IsEmpty(行) then #table(型, {}) else Table.FromRecords(行, 型),
    出す = Table.RemoveColumns(Table.SelectRows(表, each [表示する]), {"表示する"})
in
    Table.Sort(出す, {{"出勤用ID", Order.Ascending}, {"開始日", Order.Ascending}})
