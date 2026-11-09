// テレワークの予定を「出勤用ID|yyyy-mm-dd」の1日1行にして、出勤簿に出す文字を作る。
// 表示：出勤欄に足す「テレ」。備考：時間のテレワークの時刻（例「テレ13:00〜17:00」。終日の日は空欄）。
// テレワーク一覧のうち、出勤用ID・日付があり区分が正しい行だけを使う（名簿にない ID などの注意はテレワーク一覧で見る）。
let
    対象 = Table.SelectRows(テレワーク一覧, each [出勤用ID] <> "" and [日付] <> null and ([区分] = "終日" or [区分] = "時間")),
    時刻 = (t) => if t = null then "" else Time.ToText(t, "H:mm"),
    付けた = Table.AddColumn(対象, "キー", each [出勤用ID] & "|" & Date.ToText([日付], "yyyy-MM-dd"), type text),
    まとめた = Table.Group(付けた, {"キー"}, {
        {"表示", each "テレ", type text},
        {"備考", each if List.Contains([区分], "終日") then ""
            else "テレ" & Text.Combine(List.Transform(Table.ToRecords(Table.Sort(_, {{"開始時刻", Order.Ascending}})),
                (r) => 時刻(r[開始時刻]) & "〜" & 時刻(r[終了時刻])), "・"), type text}})
in
    Table.Sort(まとめた, {{"キー", Order.Ascending}})
