// 休暇申請シートの病気休暇・看護休暇（承認済み）を「出勤用ID|年|区分」ごとに、開始日の月で足す。
// 1月〜12月の列は「日」の数（時間:分 の書式で表示する値）。1日単位は 1日＝7時間45分。
// 出勤・休暇管理データ の出勤簿が、月別取得時間シートの手入力（調整）にこの値を足して表示する。
let
    正規ID = (id) => let t = Text.Trim(if id = null then "" else Text.From(id))
        in if t = "" then "" else try Text.From(Number.From(t)) otherwise t,
    月名 = List.Transform({1..12}, each Text.From(_) & "月"),
    対象 = Table.SelectRows(申請データ, each [#"承認（所属長）"] = "承認" and [年] <> null
        and List.Contains({"病気休暇", "看護休暇（子の看護等）"}, [休暇の種類])),
    行 = List.Transform(Table.ToRecords(対象), each
        [キー = 正規ID([出勤用ID]) & "|" & Text.From([年]) & "|" & (if [休暇の種類] = "病気休暇" then "病気休暇" else "看護休暇"),
         月名 = Text.From([月]) & "月", 分 = [分]]),
    表 = Table.FromRecords(行, type table [キー = text, 月名 = text, 分 = number]),
    まとめ = Table.Group(表, {"キー", "月名"}, {{"分", each List.Sum([分]), type number}}),
    横 = Table.Pivot(まとめ, 月名, "月名", "分", List.Sum),
    足りない列 = List.Difference(月名, Table.ColumnNames(横)),
    全部の月 = List.Accumulate(足りない列, 横, (t, c) => Table.AddColumn(t, c, each null)),
    日の数 = Table.TransformColumns(全部の月,
        List.Transform(月名, (c) => {c, each if _ = null then 0 else _ / 1440, type number})),
    並べた = Table.SelectColumns(日の数, {"キー"} & 月名)
in
    Table.Sort(並べた, {{"キー", Order.Ascending}})
