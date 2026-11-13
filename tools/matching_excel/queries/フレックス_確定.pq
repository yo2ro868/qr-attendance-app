// 「フレックスの日別」シートの表 確定_フレックス日別（［最新を読み込む］のマクロが、承認済みの単位期間を写したもの）を読む（接続のみ）。表がなければ null。
// 他のクエリを参照しない（Power Query は、データの読み取りと他のクエリの参照を同じクエリに置くと、設定によって止める）。
let
    結果 = try Table.Buffer(Excel.CurrentWorkbook(){[Name = "確定_フレックス日別"]}[Content]) otherwise null
in
    結果
