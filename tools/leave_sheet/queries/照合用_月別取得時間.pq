// 照合用Excel の「月別取得時間」シートと同じ列：出勤用ID・年（西暦）・区分・1月〜12月（時間:分）
// 設定シートの「集計年」の、承認済みの病気休暇・看護休暇を、開始日の月ごとに足す（1日単位は 1日＝7時間45分）
let
    年 = let v = try Excel.CurrentWorkbook(){[Name = "集計年"]}[Content]{0}[Column1] otherwise null
         in if v = null then 0 else Number.From(v),
    月名 = List.Transform({1..12}, each Text.From(_) & "月"),
    対象 = Table.SelectRows(申請データ, each [#"承認（所属長）"] = "承認" and [年] = 年
        and List.Contains({"病気休暇", "看護休暇（子の看護等）"}, [休暇の種類])),
    区分 = Table.AddColumn(対象, "区分", each if [休暇の種類] = "病気休暇" then "病気休暇" else "看護休暇", type text),
    月 = Table.AddColumn(区分, "月名", each Text.From([月]) & "月", type text),
    まとめ = Table.Group(月, {"出勤用ID", "区分", "月名"}, {{"分", each List.Sum([分]), type number}}),
    横 = Table.Pivot(まとめ, 月名, "月名", "分", List.Sum),
    足りない列 = List.Difference(月名, Table.ColumnNames(横)),
    全部の月 = List.Accumulate(足りない列, 横, (t, c) => Table.AddColumn(t, c, each null)),
    時間 = Table.TransformColumns(全部の月, List.Transform(月名, (c) => {c, each if _ = null or _ = 0 then null else _ / 1440, type nullable number})),
    年の列 = Table.AddColumn(時間, "年（西暦）", each 年, type number),
    並べた = Table.SelectColumns(年の列, {"出勤用ID", "年（西暦）", "区分"} & 月名)
in
    Table.Sort(並べた, {{"出勤用ID", Order.Ascending}, {"区分", Order.Descending}})
