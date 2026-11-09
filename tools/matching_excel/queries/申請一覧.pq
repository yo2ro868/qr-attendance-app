// 「申請一覧」シートに出す：休暇申請シートと Forms の全員の申請（出勤用ID・開始日の順）。
// 先頭に、読み取れなかった休暇申請シートと、対応する申請が見つからない承認フォームの回答を出す。
let
    出典 = (f as text) => if Text.StartsWith(f, "Forms") then "Forms" else "休暇申請シート",
    正規ID = (id) => let t = Text.Trim(if id = null then "" else Text.From(id))
        in if t = "" then "" else try Text.From(Number.From(t)) otherwise t,
    並べた = Table.Sort(申請データ, {{"出勤用ID", Order.Ascending}, {"開始日", Order.Ascending}}),
    状況 = Table.AddColumn(並べた, "状況", each
        if [#"承認（所属長）"] = "" then "承認待ち" else [#"承認（所属長）"], type text),
    出典つき = Table.AddColumn(状況, "出典", each 出典([ファイル名]), type text),
    列名 = {"出典", "出勤用ID", "氏名", "申請日", "休暇の種類", "種別・内容", "単位", "開始日", "終了日", "日数", "時間",
        "時刻", "事由", "状況", "承認日", "所属長コメント", "確認", "ファイル名"},
    // 日数と、期間の勤務日（土日・祝日シートの日・12/29〜1/3 を除く）の数が合わないときは注意を足す
    休み = (d as date) => Date.DayOfWeek(d, Day.Monday) >= 5 or List.Contains(祝日の日付, d)
        or (Date.Month(d) = 12 and Date.Day(d) >= 29) or (Date.Month(d) = 1 and Date.Day(d) <= 3),
    勤務日数 = (開始 as date, 終了 as date) =>
        List.Count(List.Select(List.Dates(開始, Duration.Days(終了 - 開始) + 1, #duration(1, 0, 0, 0)), each not 休み(_))),
    合わない = (r as record) as text =>
        if r[開始日] = null or r[日数] = null or (r[終了日] <> null and r[終了日] < r[開始日]) then ""
        else if Text.Contains(r[単位], "時間") and r[終了日] = null then ""
        else if r[終了日] = null and 休み(r[開始日]) then ""     // 開始日の注意だけにする
        else
            let n = 勤務日数(r[開始日], if r[終了日] = null then r[開始日] else r[終了日])
            in if n = r[日数] then "" else "・日数（" & Text.From(r[日数]) & "）と期間の勤務日（" & Text.From(n) & "）が合いません",
    // Forms は入力時に止められないので、開始日・終了日が休みの日のときはここで注意を出す（休暇申請シートは「確認」の欄に出る）
    休みに開始 = (r as record) as text =>
        if not Text.StartsWith(r[ファイル名], "Forms") then ""
        else (if r[開始日] <> null and 休み(r[開始日]) then "・開始日が土日・祝日・年末年始です" else "")
            & (if r[終了日] <> null and 休み(r[終了日]) then "・終了日が土日・祝日・年末年始です" else ""),
    注意つき = Table.ReplaceValue(Table.AddColumn(出典つき, "日数の注意", each 休みに開始(_) & 合わない(_), type text), each [確認], each [確認] & [日数の注意], Replacer.ReplaceValue, {"確認"}),
    列 = Table.SelectColumns(注意つき, 列名),
    空の行 = (確認 as text, ファイル as text, 出典の値 as text) =>
        [出典 = 出典の値, 出勤用ID = "", 氏名 = "", 申請日 = null, 休暇の種類 = "", #"種別・内容" = "", 単位 = "",
         開始日 = null, 終了日 = null, 日数 = null, 時間 = null, 時刻 = "", 事由 = "", 状況 = "",
         承認日 = null, 所属長コメント = "", 確認 = 確認, ファイル名 = ファイル],
    読めない = List.Transform(Table.ToRecords(Table.SelectRows(休暇申請シート, each [読み取り] <> "OK")),
        each 空の行([読み取り], [ファイル名], "休暇申請シート")),
    // Forms の申請のキー（出勤用ID|開始日|休暇の種類）にない承認の回答
    Forms申請のキー = List.Transform(Table.ToRecords(Forms申請), each
        正規ID([出勤用ID]) & "|" & (if [開始日] = null then "" else Date.ToText([開始日], "yyyy-MM-dd")) & "|" & [休暇の種類]),
    宙に浮いた承認 = List.Transform(Table.ToRecords(Table.SelectRows(Forms承認, each not List.Contains(Forms申請のキー, [キー]))),
        each 空の行("承認フォームの回答（" & [キー] & "、" & [結果] & "、" & [承認者] & "）に合う Forms の申請がありません。出勤用ID・開始日・休暇の種類を確かめてください", "", "Forms")),
    // Forms の回答ファイルの場所が違うとき
    読めないForms = List.Transform(List.Select({{"Forms申請ファイル", Forms回答_申請}, {"Forms承認ファイル", Forms回答_承認}},
            each _{1}[パス] <> "" and _{1}[表] = null), each
        空の行("Forms の回答ファイルが見つかりません。場所を確かめてください（" & _{0} & "）", _{1}[パス], "Forms")),
    先頭 = Table.FromRecords(読めない & 読めないForms & 宙に浮いた承認, Value.Type(列))
in
    Table.Combine({先頭, 列})
