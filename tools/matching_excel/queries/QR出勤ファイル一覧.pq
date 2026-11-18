// 「QR出勤の読み込み」シートに出す：フォルダで読んだ出勤記録のファイルと、その行数・読み取りの結果（読めなかったファイルの理由）。
let
    表 = Table.AddColumn(QR出勤_ファイル, "行数", each if [行] = null then null else Table.RowCount([行]), type nullable number),
    並べた = Table.Sort(Table.SelectColumns(表, {"ファイル名", "更新日時", "行数", "読み取り"}), {{"更新日時", Order.Descending}})
in
    // 更新日時は文字で（列の幅が狭くても ### にならないように）
    Table.TransformColumns(並べた, {{"更新日時", each if _ = null then "" else DateTime.ToText(_, "yyyy/MM/dd HH:mm"), type text}})
