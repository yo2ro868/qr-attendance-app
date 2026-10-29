// 「申請一覧」シートに出す：全員の申請（出勤用ID・開始日の順）と、読み取れなかったファイル
let
    並べた = Table.Sort(申請データ, {{"出勤用ID", Order.Ascending}, {"開始日", Order.Ascending}}),
    状況 = Table.AddColumn(並べた, "状況", each
        if [#"承認（所属長）"] = "" then "承認待ち" else [#"承認（所属長）"], type text),
    列 = Table.SelectColumns(状況, {"出勤用ID", "氏名", "所属長", "申請日", "休暇の種類", "種別・内容", "単位",
        "開始日", "終了日", "日数", "時間", "時刻", "事由", "状況", "承認日", "所属長コメント", "確認", "ファイル名"}),
    読めない = Table.SelectRows(申請用紙, each [読み取り] <> "OK"),
    読めない行 = Table.FromRecords(List.Transform(Table.ToRecords(読めない), each
        [出勤用ID = "", 氏名 = "", 所属長 = "", 申請日 = null, 休暇の種類 = "", #"種別・内容" = "", 単位 = "",
         開始日 = null, 終了日 = null, 日数 = null, 時間 = null, 時刻 = "", 事由 = "", 状況 = "",
         承認日 = null, 所属長コメント = "", 確認 = [読み取り], ファイル名 = [ファイル名]]), Value.Type(列))
in
    Table.Combine({読めない行, 列})
