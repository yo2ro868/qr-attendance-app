// 職員名簿（接続のみ）。Forms の回答者のメールアドレスから出勤用ID を引くのに使う。
// 他のクエリを参照しない（Power Query は、データの読み取りと他のクエリの参照を同じクエリに置くと、設定によって止める）。
let
    結果 = try Table.PromoteHeaders(Excel.CurrentWorkbook(){[Name = "職員名簿表"]}[Content], [PromoteAllScalars = true])
        otherwise #table({"出勤用ID", "氏名", "メールアドレス"}, {})
in
    結果
