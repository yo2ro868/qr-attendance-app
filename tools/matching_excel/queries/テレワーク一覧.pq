// 全所属のテレワークの予定（テレワーク予定表）と注意。「テレワーク一覧」シートに出す。
// 注意：名簿にない出勤用ID、日付や区分の誤り、土日・祝日・年末年始、時刻の誤り、同じ職員・同じ日の重複、
//       承認済みの休暇と日時が重なる予定（休暇申請シートの申請データ。時間どうしは時刻が重なるときだけ）、読めないファイル。
// 区分が「取消」の行（所属長が取りやめた予定）は一覧に残すが、重複・休暇の注意は出さない（テレワーク_日別は終日・時間だけを使う）。
// 時間は開始時刻・終了時刻から12:00〜13:00の休憩を除いて計算する（テレワーク予定表の時間の欄と同じ）。
let
    予定 = テレワーク予定表,
    名簿 = Table.Buffer(職員名簿_メール),
    正規ID = (v) => let t = if v = null then "" else Text.Trim(Text.From(v))
        in if t = "" then "" else try Text.From(Number.From(t)) otherwise t,
    // 古い Excel でも動くように ?? 演算子は使わない
    氏名の表 = Table.FromRecords(List.Transform(Table.ToRecords(名簿), each
        [出勤用ID = 正規ID(Record.FieldOrDefault(_, "出勤用ID", null)),
         氏名 = let v = Record.FieldOrDefault(_, "氏名", null) in if v = null then "" else Text.From(v)]),
        type table [出勤用ID = text, 氏名 = text]),
    氏名を引く = (id as text) => let t = Table.SelectRows(氏名の表, each [出勤用ID] = id) in if Table.IsEmpty(t) then null else t{0}[氏名],
    休み = (d as date) => Date.DayOfWeek(d, Day.Monday) >= 5 or List.Contains(祝日の日付, d)
        or (Date.Month(d) = 12 and Date.Day(d) >= 29) or (Date.Month(d) = 1 and Date.Day(d) <= 3),
    分 = (t as time) => Time.Hour(t) * 60 + Time.Minute(t),
    時間の分 = (s, e) => if s = null or e = null or e <= s then null
        else 分(e) - 分(s) - List.Max({0, List.Min({分(e), 13 * 60}) - List.Max({分(s), 12 * 60})}),
    キー = (r as record) => if r[出勤用ID] = "" or r[日付] = null then "" else r[出勤用ID] & "|" & Date.ToText(r[日付], "yyyy-MM-dd"),
    有効 = Table.SelectRows(予定, each [読み取り] = "OK" and [区分] <> "取消"),
    同じ日の数 = Table.Group(Table.AddColumn(有効, "キー", キー, type text), {"キー"}, {{"数", Table.RowCount, Int64.Type}}),
    数を引く = (k as text) => let t = Table.SelectRows(同じ日の数, each [キー] = k) in if k = "" or Table.IsEmpty(t) then 0 else t{0}[数],
    // 承認済みの休暇（申請データ）。日単位は開始日〜終了日、時間・分単位は開始日（日数が2以上で終了日があれば開始日〜終了日）の、時刻「9:00〜12:00」
    時刻を読む = (t as text) => let p = Text.Split(t, "〜") in
        if List.Count(p) <> 2 then null else try {分(Time.FromText(Text.Trim(p{0}))), 分(Time.FromText(Text.Trim(p{1})))} otherwise null,
    休暇 = Table.Buffer(Table.FromRecords(List.Transform(Table.ToRecords(
            Table.SelectRows(申請データ, each [#"承認（所属長）"] = "承認" and [開始日] <> null)), (r) =>
        let
            時間単位 = Text.Contains(r[単位], "時間"),
            続けて = 時間単位 and r[日数] <> null and r[日数] > 1 and r[終了日] <> null,
            終わり = if (時間単位 and not 続けて) or r[終了日] = null or r[終了日] < r[開始日] then r[開始日] else r[終了日]
        in
            [出勤用ID = 正規ID(r[出勤用ID]), 開始日 = r[開始日], 終了日 = 終わり,
             時刻 = if 時間単位 then 時刻を読む(r[時刻]) else null, 時間単位 = 時間単位,
             名前 = r[休暇の種類] & (if 時間単位 then " " & r[時刻] else " 1日")]),
        type table [出勤用ID = text, 開始日 = date, 終了日 = date, 時刻 = any, 時間単位 = logical, 名前 = text])),
    // テレワークの予定と重なる休暇の名前（終日のテレワークや1日の休暇は、その日なら重なる。時間どうしは時刻が重なるときだけ）
    重なる休暇 = (r as record) as list =>
        if r[出勤用ID] = "" or r[日付] = null then {} else
        let
            id = 正規ID(r[出勤用ID]),
            その日 = Table.SelectRows(休暇, each [出勤用ID] = id and [開始日] <= r[日付] and [終了日] >= r[日付]),
            テレ = if r[区分] = "時間" and r[開始時刻] <> null and r[終了時刻] <> null then {分(r[開始時刻]), 分(r[終了時刻])} else null,
            重なる = Table.SelectRows(その日, each not [時間単位] or [時刻] = null or テレ = null
                or ([時刻]{0} < テレ{1} and [時刻]{1} > テレ{0}))
        in
            重なる[名前],
    行 = List.Transform(Table.ToRecords(予定), (r) =>
        let
            k = キー(r),
            氏名 = if r[出勤用ID] = "" then null else 氏名を引く(r[出勤用ID]),
            時間 = if r[区分] = "時間" then 時間の分(r[開始時刻], r[終了時刻]) else null,
            取消 = r[区分] = "取消",
            休暇の名前 = if 取消 or r[読み取り] <> "OK" then {} else 重なる休暇(r),
            注意 = if r[読み取り] <> "OK" then {r[読み取り]} else List.RemoveNulls({
                if r[出勤用ID] = "" then "出勤用ID がありません" else if 氏名 = null then "職員名簿にない出勤用ID です" else null,
                if r[日付] = null then "日付がない（または日付の形でない）" else if 休み(r[日付]) then "土日・祝日・年末年始です" else null,
                if r[区分] <> "終日" and r[区分] <> "時間" and not 取消 then "区分は「終日」「時間」「取消」です" else null,
                if r[区分] = "時間" and (r[開始時刻] = null or r[終了時刻] = null) then "開始時刻・終了時刻がありません"
                    else if r[開始時刻] <> null and r[終了時刻] <> null and r[終了時刻] <= r[開始時刻] then "終了時刻が開始時刻より前です" else null,
                if not 取消 and 数を引く(k) > 1 then "同じ職員・同じ日の予定がほかにもあります" else null,
                if List.IsEmpty(休暇の名前) then null else "承認済みの休暇と重なります（" & Text.Combine(休暇の名前, "、") & "）"})
        in
            [所属 = r[所属], 出勤用ID = r[出勤用ID], 氏名 = if 氏名 = null then r[氏名] else 氏名, 日付 = r[日付], 区分 = r[区分],
             開始時刻 = r[開始時刻], 終了時刻 = r[終了時刻],
             時間 = if 時間 = null then null else #duration(0, 0, 時間, 0),
             備考 = r[備考], 注意 = Text.Combine(注意, "・"), ファイル名 = r[ファイル名]]),
    型 = type table [所属 = text, 出勤用ID = text, 氏名 = text, 日付 = nullable date, 区分 = text,
        開始時刻 = nullable time, 終了時刻 = nullable time, 時間 = nullable duration, 備考 = text, 注意 = text, ファイル名 = text],
    表 = if List.IsEmpty(行) then #table(型, {}) else Table.FromRecords(行, 型),
    // 読めないファイルなどを先頭に、あとは日付・所属・出勤用ID の順
    並べた = Table.Sort(Table.AddColumn(表, "順", each if [日付] = null then 0 else 1, Int64.Type),
        {{"順", Order.Ascending}, {"日付", Order.Ascending}, {"所属", Order.Ascending}, {"出勤用ID", Order.Ascending}})
in
    Table.RemoveColumns(並べた, {"順"})
