// 承認済みの年次休暇（分）を「出勤用ID|年」（開始日の年）ごとに足す。休暇申請シートと Forms の合計と、そのうち Forms の分。
// 出勤・休暇管理データ の計算用シートが、休暇等シートの本年使用の調整（導入前の使用など。手入力）に合計を足して本年使用にする。
// Forms の分は、休暇申請シートの作成.xlsm が休暇申請シートの「年次休暇の残り」に入れるために使う（休暇申請シートは自分の行しか数えないため）。
let
    // 出勤・休暇管理データ の照合キーと同じ：数だけの ID は先頭の0を取る
    正規ID = (id) => let t = Text.Trim(if id = null then "" else Text.From(id))
        in if t = "" then "" else try Text.From(Number.From(t)) otherwise t,
    申請 = Table.SelectRows(申請データ, each [休暇の種類] = "年次休暇" and [#"承認（所属長）"] = "承認"
        and [年] <> null and [出勤用ID] <> ""),
    申請の行 = List.Transform(Table.ToRecords(申請), each
        [キー = 正規ID([出勤用ID]) & "|" & Text.From([年]), 分 = [分],
         Forms分 = if Text.StartsWith([ファイル名], "Forms") then [分] else 0]),
    全部 = Table.FromRecords(申請の行, type table [キー = text, 分 = number, Forms分 = number]),
    まとめ = Table.Group(全部, {"キー"}, {{"分", each List.Sum([分]), type number}, {"Forms分", each List.Sum([Forms分]), type number}})
in
    Table.Sort(まとめ, {{"キー", Order.Ascending}})
