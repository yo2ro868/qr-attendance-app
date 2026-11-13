// 出勤簿の備考に出すフレックスの割振り（取り込んだ承認済みの内容。確定_フレックス日別）：キー「出勤用ID|yyyy-mm-dd」と表示の文字。
// 平日で勤務を割り振らない日は「割振なし」、既定（9:00〜17:45）と違う勤務の日は「フ10:00-19:00」のように出す。
let
    文字 = (v) => if v = null then "" else Text.Trim(Text.From(v)),
    日付にする = (v) => if v = null then null else if v is date then v else if v is datetime then DateTime.Date(v)
        else if v is number then Date.From(v) else try Date.From(v) otherwise null,
    分 = (v) => if v = null or 文字(v) = "" then null else try Number.Round(Number.From(v) * 1440) otherwise null,
    時刻 = (m as number) => Text.From(Number.IntegerDivide(m, 60)) & ":" & Text.PadStart(Text.From(Number.Mod(m, 60)), 2, "0"),
    行 = if フレックス_確定 = null then {} else Table.ToRecords(フレックス_確定),
    表示 = List.Transform(行, (r) =>
        let s = 分(r[勤務開始]), e = 分(r[勤務終了]), d = 日付にする(r[日付]), 区分 = 文字(r[区分])
        in [キー = 文字(r[出勤用ID]) & "|" & (if d = null then "" else Date.ToText(d, "yyyy-MM-dd")),
            表示 = if 区分 = "平日" and s = null then "割振なし"
                else if s <> null and e <> null and (s <> 540 or e <> 1065) then "フ" & 時刻(s) & "-" & 時刻(e)
                else ""]),
    出す = List.Select(表示, each [表示] <> "" and not Text.EndsWith([キー], "|"))
in
    if List.IsEmpty(出す) then #table(type table [キー = text, 表示 = text], {})
    else Table.Distinct(Table.FromRecords(出す, type table [キー = text, 表示 = text]), {"キー"})
