// 申請用紙の年次休暇の使用（分）を「出勤用ID|年」ごとに足す。
// 承認済みの年次休暇（開始日の年）と、申請用紙の「導入前の使用」（申請用紙の基本の年）を合わせる。
// 照合用Excel の計算用シートが、年次休暇シートの手入力（調整）にこの値を足して本年使用にする。
let
    // 照合用Excel の照合キーと同じ：数だけの ID は先頭の0を取る
    正規ID = (id) => let t = Text.Trim(if id = null then "" else Text.From(id))
        in if t = "" then "" else try Text.From(Number.From(t)) otherwise t,
    申請 = Table.SelectRows(申請データ, each [休暇の種類] = "年次休暇" and [#"承認（所属長）"] = "承認" and [年] <> null),
    申請の行 = List.Transform(Table.ToRecords(申請), each [キー = 正規ID([出勤用ID]) & "|" & Text.From([年]), 分 = [分]]),
    導入前 = Table.SelectRows(申請用紙, each [読み取り] = "OK" and [出勤用ID] <> "" and [年] > 0 and [導入前の使用] > 0),
    導入前の行 = List.Transform(Table.ToRecords(導入前), each [キー = 正規ID([出勤用ID]) & "|" & Text.From([年]), 分 = [導入前の使用]]),
    全部 = Table.FromRecords(申請の行 & 導入前の行, type table [キー = text, 分 = number]),
    まとめ = Table.Group(全部, {"キー"}, {{"分", each List.Sum([分]), type number}})
in
    Table.Sort(まとめ, {{"キー", Order.Ascending}})
