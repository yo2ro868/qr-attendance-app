// 「利用予定」シート（フレックス／テレワーク利用予定の一覧）の日ごとの文字：キー「出勤用ID|yyyy-mm-dd」と表示。
//   1日の休暇（承認済み）… 年休・病休・夏休・特休 など（これがある日はこれだけ）
//   テレワークの予定 … テレ
//   フレックス … 勤務開始の時刻（9:15 など）。勤務を割り振らない平日は「なし」。
//     所属長が承認して取り込んだ単位期間はその内容（確定_フレックス日別）、まだの単位期間は今の申告（フレックス_日別。申告年月日のあるもの）
//   時間単位の休暇（承認済み）… 午前／午後＋時間＋略称（例 午後4H年休）
let
    文字にする = (v) => if v = null then "" else Text.Trim(Text.From(v)),
    正規ID = (id) => let t = 文字にする(id) in if t = "" then "" else try Text.From(Number.From(t)) otherwise t,
    日付にする = (v) => if v = null then null else if v is date then v else if v is datetime then DateTime.Date(v)
        else if v is number then Date.From(v) else try Date.From(v) otherwise null,
    分 = (v) => if v = null or 文字にする(v) = "" then null else try Number.Round(Number.From(v) * 1440) otherwise null,
    時刻 = (m as number) => Text.From(Number.IntegerDivide(m, 60)) & ":" & Text.PadStart(Text.From(Number.Mod(m, 60)), 2, "0"),
    キーを作る = (id, d) => 正規ID(id) & "|" & Date.ToText(d, "yyyy-MM-dd"),
    休み = (d as date) => Date.DayOfWeek(d, Day.Monday) >= 5 or List.Contains(祝日の日付, d)
        or (Date.Month(d) = 12 and Date.Day(d) >= 29) or (Date.Month(d) = 1 and Date.Day(d) <= 3),
    略称 = (種類 as text, 内容 as text) =>
        if 種類 = "年次休暇" then "年休" else if 種類 = "病気休暇" then "病休"
        else if 種類 = "看護休暇（子の看護等）" then "看護"
        else if 種類 = "特別休暇（看護休暇以外）" then (if 内容 = "夏季休暇" then "夏休" else "特休")
        else if 種類 = "介護休暇" then "介護" else if 種類 = "介護時間" then "介時" else if 種類 = "育児時間" then "育時"
        else if 種類 = "職務専念義務の免除" then "職免" else 種類,

    // ----- 休暇（承認済み）：[キー, 種, 文字]（種 1＝1日、3＝時間）-----
    承認 = Table.ToRecords(Table.SelectRows(申請データ, each [#"承認（所属長）"] = "承認" and [開始日] <> null and [出勤用ID] <> "")),
    休暇の日 = List.Combine(List.Transform(承認, (r) =>
        let
            時間単位 = Text.Contains(文字にする(r[単位]), "時間"),
            終わり = if r[終了日] = null or r[終了日] < r[開始日] then r[開始日] else r[終了日],
            期間 = List.Dates(r[開始日], Duration.Days(終わり - r[開始日]) + 1, #duration(1, 0, 0, 0)),
            勤務日 = List.Select(期間, each not 休み(_)),
            続けて = 時間単位 and r[日数] <> null and r[日数] > 1 and r[終了日] <> null,
            日たち = if 時間単位 and not 続けて then {r[開始日]} else if List.IsEmpty(勤務日) then {r[開始日]} else 勤務日,
            名 = 略称(文字にする(r[休暇の種類]), 文字にする(r[#"種別・内容"])),
            // 時刻「9:00〜12:00」から午前・午後
            時刻の文字 = 文字にする(r[時刻]),
            始め = try Time.FromText(Text.BeforeDelimiter(時刻の文字, "〜")) otherwise null,
            終り = try Time.FromText(Text.AfterDelimiter(時刻の文字, "〜")) otherwise null,
            午 = if 始め <> null and Time.Hour(始め) >= 12 then "午後" else if 終り <> null and (Time.Hour(終り) * 60 + Time.Minute(終り)) <= 13 * 60 then "午前" else "",
            時間数 = if r[時間] = null then 0 else Number.Round(r[時間] * 24 * 4) / 4,
            時間の文字 = 午 & (if 時間数 > 0 then Text.From(時間数) & "H" else "") & 名
        in
            List.Transform(日たち, (d) => [キー = キーを作る(r[出勤用ID], d), 種 = if 時間単位 then 3 else 1, 文字 = if 時間単位 then 時間の文字 else 名]))),

    // ----- テレワーク：[キー, 種 2, テレ] -----
    テレの日 = List.Transform(Table.ToRecords(テレワーク_日別), each [キー = [キー], 種 = 2, 文字 = "テレ"]),

    // ----- フレックス：取り込んだ単位期間はその内容、まだの単位期間は今の申告 -----
    期間キー = (r as record) => 正規ID(r[出勤用ID]) & "|" & 文字にする(r[期間]) & "|" &
        (let d = 日付にする(r[期間の開始日]) in if d = null then "" else Date.ToText(d, "yyyy-MM-dd")),
    確定 = if フレックス_確定 = null then {} else List.Select(Table.ToRecords(フレックス_確定), each 文字にする([出勤用ID]) <> ""),
    確定の期間 = List.Buffer(List.Distinct(List.Transform(確定, 期間キー))),
    // 今の申告は、申告年月日を入れた単位期間だけ（未申告の単位期間の既定の時刻は出さない）
    今 = List.Select(Table.ToRecords(フレックス_日別), each [状態] <> "未申告" and not List.Contains(確定の期間, 期間キー(_))),
    フレックスの日 = List.RemoveNulls(List.Transform(確定 & 今, (r) =>
        let d = 日付にする(r[日付]), s = 分(r[勤務開始]), 区分 = 文字にする(r[区分])
        in if d = null or 区分 = "休日" then null
           else if s <> null then [キー = キーを作る(r[出勤用ID], d), 種 = 2.5, 文字 = 時刻(s)]
           else if 区分 = "平日" then [キー = キーを作る(r[出勤用ID], d), 種 = 2.5, 文字 = "なし"]
           else null)),

    全部 = Table.FromRecords(休暇の日 & テレの日 & フレックスの日, type table [キー = text, 種 = number, 文字 = text]),
    まとめ = Table.Group(全部, {"キー"}, {{"表示", each
        let 行 = Table.Sort(_, {{"種", Order.Ascending}}), 一日 = Table.SelectRows(行, each [種] = 1)
        in if not Table.IsEmpty(一日) then Text.Combine(List.Distinct(一日[文字]), "・")
           else Text.Combine(List.Distinct(行[文字]), ""), type text}})
in
    Table.Sort(まとめ, {{"キー", Order.Ascending}})
