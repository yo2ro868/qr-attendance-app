// 祝日シートの日付の一覧（接続のみ）。
// 他のクエリを参照しない（Power Query は、データの読み取りと他のクエリの参照を同じクエリに置くと、設定によって止める）。
let
    結果 =
        let t = try Excel.CurrentWorkbook(){[Name = "祝日一覧"]}[Content] otherwise null
        in if t = null then {} else List.RemoveNulls(List.Transform(t[Column1], each
            if _ is datetime then DateTime.Date(_) else if _ is date then _ else if _ is number then Date.From(_) else null))
in
    結果
