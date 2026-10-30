// 申請用紙の承認済みの申請を「出勤用ID|yyyy-mm-dd」の1日1行にして、出勤簿の日付の表の欄の文字を作る。
// 1日単位は 開始日〜終了日 の勤務日（土日・祝日シートの日・12/29〜1/3 を除く）に「1日」、時間単位は開始日に「2:15」の形。
// 看護休暇は特別休暇（種別「子の看護」）として出す。育児時間は出勤簿に欄がないので出さない。
let
    正規ID = (id) => let t = Text.Trim(if id = null then "" else Text.From(id))
        in if t = "" then "" else try Text.From(Number.From(t)) otherwise t,
    祝日 = let t = try Excel.CurrentWorkbook(){[Name = "祝日一覧"]}[Content] otherwise null
        in if t = null then {} else List.RemoveNulls(List.Transform(t[Column1], each
            if _ is datetime then DateTime.Date(_) else if _ is date then _ else if _ is number then Date.From(_) else null)),
    休み = (d as date) => Date.DayOfWeek(d, Day.Monday) >= 5 or List.Contains(祝日, d)
        or (Date.Month(d) = 12 and Date.Day(d) >= 29) or (Date.Month(d) = 1 and Date.Day(d) <= 3),
    時分 = (m as number) => Text.From(Number.IntegerDivide(m, 60)) & ":" & Text.PadStart(Text.From(Number.Mod(m, 60)), 2, "0"),
    // 休暇の種類 → 出勤簿の欄（時間の欄, 種別の欄, 種別の文字）
    欄 = (r as record) =>
        if r[休暇の種類] = "年次休暇" then {"年次休暇", null, null}
        else if r[休暇の種類] = "病気休暇" then {"病気休暇", null, null}
        else if r[休暇の種類] = "看護休暇（子の看護等）" then {"特別休暇（時間）", "特別休暇（種別）", "子の看護"}
        else if r[休暇の種類] = "特別休暇（看護休暇以外）" then
            {"特別休暇（時間）", "特別休暇（種別）", if r[#"種別・内容"] = "" then "特別休暇" else r[#"種別・内容"]}
        else if r[休暇の種類] = "介護休暇" then {"介護休暇", null, null}
        else if r[休暇の種類] = "介護時間" then {"介護時間", null, null}
        else if r[休暇の種類] = "職務専念義務の免除" then
            {"職専免（時間）", "職専免（種別）", if r[#"種別・内容"] = "" then "職専免" else r[#"種別・内容"]}
        else null,
    対象 = Table.SelectRows(申請データ, each [#"承認（所属長）"] = "承認" and [開始日] <> null),
    // 申請1件 → 1日ごとのレコード [キー, 欄, 値の分（1日単位は null）, 種別の欄, 種別]
    日にする = (r as record) as list =>
        let
            e = 欄(r),
            id = 正規ID(r[出勤用ID]),
            時間単位 = r[単位] = "時間単位",
            終わり = if r[終了日] = null or r[終了日] < r[開始日] then r[開始日] else r[終了日],
            期間 = List.Dates(r[開始日], Duration.Days(終わり - r[開始日]) + 1, #duration(1, 0, 0, 0)),
            勤務日 = List.Select(期間, each not 休み(_)),
            日付 = if 時間単位 then {r[開始日]} else if List.IsEmpty(勤務日) then {r[開始日]} else 勤務日
        in
            if e = null or id = "" then {}
            else List.Transform(日付, (d) =>
                [キー = id & "|" & Date.ToText(d, "yyyy-MM-dd"), 欄 = e{0}, 分 = if 時間単位 then r[分] else null,
                 種別の欄 = e{1}, 種別 = e{2}]),
    日 = List.Combine(List.Transform(Table.ToRecords(対象), 日にする)),
    表 = Table.FromRecords(日, type table [キー = text, 欄 = text, 分 = nullable number, 種別の欄 = nullable text, 種別 = nullable text]),
    // 同じ日・同じ欄に2件以上あれば、1日単位があれば「1日」、なければ時間を足す
    値 = Table.Group(表, {"キー", "欄"}, {{"値", each
        if List.Contains([分], null) then "1日" else 時分(List.Sum([分])), type text}}),
    種別の表 = Table.SelectRows(表, each [種別の欄] <> null),
    種別 = Table.Group(種別の表, {"キー", "種別の欄"}, {{"値", each Text.Combine(List.Distinct([種別]), "・"), type text}}),
    縦 = Table.Combine({値, Table.RenameColumns(種別, {{"種別の欄", "欄"}})}),
    欄名 = {"年次休暇", "病気休暇", "特別休暇（種別）", "特別休暇（時間）", "介護休暇", "介護時間", "職専免（種別）", "職専免（時間）"},
    横 = Table.Pivot(縦, 欄名, "欄", "値", each if List.IsEmpty(_) then null else _{0}),
    並べた = Table.SelectColumns(横, {"キー"} & 欄名, MissingField.UseNull)
in
    Table.Sort(並べた, {{"キー", Order.Ascending}})
