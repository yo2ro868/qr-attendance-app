// 「申請一覧」シートに出す：申請用紙と Forms の全員の申請（出勤用ID・開始日の順）。
// 先頭に、読み取れなかった申請用紙と、対応する申請が見つからない承認フォームの回答を出す。
let
    出典 = (f as text) => if Text.StartsWith(f, "Forms") then "Forms" else "申請用紙",
    正規ID = (id) => let t = Text.Trim(if id = null then "" else Text.From(id))
        in if t = "" then "" else try Text.From(Number.From(t)) otherwise t,
    並べた = Table.Sort(申請データ, {{"出勤用ID", Order.Ascending}, {"開始日", Order.Ascending}}),
    状況 = Table.AddColumn(並べた, "状況", each
        if [#"承認（所属長）"] = "" then "承認待ち" else [#"承認（所属長）"], type text),
    出典つき = Table.AddColumn(状況, "出典", each 出典([ファイル名]), type text),
    列名 = {"出典", "出勤用ID", "氏名", "申請日", "休暇の種類", "種別・内容", "単位", "開始日", "終了日", "日数", "時間",
        "時刻", "事由", "状況", "承認日", "所属長コメント", "確認", "ファイル名"},
    列 = Table.SelectColumns(出典つき, 列名),
    空の行 = (確認 as text, ファイル as text, 出典の値 as text) =>
        [出典 = 出典の値, 出勤用ID = "", 氏名 = "", 申請日 = null, 休暇の種類 = "", #"種別・内容" = "", 単位 = "",
         開始日 = null, 終了日 = null, 日数 = null, 時間 = null, 時刻 = "", 事由 = "", 状況 = "",
         承認日 = null, 所属長コメント = "", 確認 = 確認, ファイル名 = ファイル],
    読めない = List.Transform(Table.ToRecords(Table.SelectRows(申請用紙, each [読み取り] <> "OK")),
        each 空の行([読み取り], [ファイル名], "申請用紙")),
    // Forms の申請のキー（出勤用ID|開始日|休暇の種類）にない承認の回答
    Forms申請のキー = List.Transform(Table.ToRecords(Forms申請), each
        正規ID([出勤用ID]) & "|" & (if [開始日] = null then "" else Date.ToText([開始日], "yyyy-MM-dd")) & "|" & [休暇の種類]),
    宙に浮いた承認 = List.Transform(Table.ToRecords(Table.SelectRows(Forms承認, each not List.Contains(Forms申請のキー, [キー]))),
        each 空の行("承認フォームの回答（" & [キー] & "、" & [結果] & "、" & [承認者] & "）に合う Forms の申請がありません。出勤用ID・開始日・休暇の種類を確かめてください", "", "Forms")),
    先頭 = Table.FromRecords(読めない & 宙に浮いた承認, Value.Type(列))
in
    Table.Combine({先頭, 列})
