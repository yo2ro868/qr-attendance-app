// 照合用Excel の「年次休暇」シートと同じ列：出勤用ID・年（西暦）・本年分（日）（空欄）・本年使用（日）・本年使用（時間）
// 設定シートの「集計年」の、承認済みの年次休暇と、申請用紙の「導入前の使用」（申請用紙の年が集計年のとき）を足す
let
    年 = let v = try Excel.CurrentWorkbook(){[Name = "集計年"]}[Content]{0}[Column1] otherwise null
         in if v = null then 0 else Number.From(v),
    承認済み = Table.SelectRows(申請データ, each [休暇の種類] = "年次休暇" and [#"承認（所属長）"] = "承認" and [年] = 年),
    人 = Table.SelectRows(申請用紙, each [読み取り] = "OK" and [出勤用ID] <> ""),
    行 = List.Transform(Table.ToRecords(人), (p) =>
        let
            自分 = Table.SelectRows(承認済み, each [ファイル名] = p[ファイル名]),
            日 = List.Sum({0} & List.Transform(Table.ToRecords(Table.SelectRows(自分, each [単位] <> "時間単位")),
                each if [日数] = null then 0 else [日数])),
            分 = List.Sum({0} & List.Transform(Table.ToRecords(Table.SelectRows(自分, each [単位] = "時間単位")), each [分])),
            // 導入前の使用（分）は 7時間45分ごとに日へ、残りを時間へ
            導入前 = if p[年] = 年 then p[導入前の使用] else 0
        in
            [出勤用ID = p[出勤用ID], #"年（西暦）" = 年, #"本年分（日）" = null,
             #"本年使用（日）" = 日 + Number.IntegerDivide(導入前, 465),
             #"本年使用（時間）" = (分 + Number.Mod(導入前, 465)) / 1440]),
    型 = type table [出勤用ID = text, #"年（西暦）" = number, #"本年分（日）" = nullable number,
        #"本年使用（日）" = number, #"本年使用（時間）" = number]
in
    Table.Sort(Table.FromRecords(行, 型), {{"出勤用ID", Order.Ascending}})
