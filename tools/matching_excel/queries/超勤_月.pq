// 休暇申請シートの超過勤務等命令簿の月のシートごとに1行（1人1か月）：所属長の確認、超過勤務のある日数、支給割合ごとの月の計と手当対象時間。
// 手当対象時間は命令簿の「手当対象時間（30分未満切捨て）」と同じく、月の計の30分未満を切り捨て、30分以上を1時間に切り上げる。
// 「超過勤務の読み込み」シートに出す。所属長の確認日のない月は「確認待ち」で、［超過勤務を転記］では写さない。
let
    正規ID = (id) => let t = Text.Trim(if id = null then "" else Text.From(id))
        in if t = "" then "" else try Text.From(Number.From(t)) otherwise t,
    文字にする = (v) => if v = null then "" else Text.Trim(Text.From(v)),
    分にする = (v) => if v = null or 文字にする(v) = "" then 0 else try Number.From(v) otherwise 0,
    日付にする = (v) => if 文字にする(v) = "" then null else try Date.FromText(文字にする(v), [Format = "yyyy-MM-dd"]) otherwise null,
    時間の列 = {"125", "150", "150（60H前）", "150（+25）", "150（+15）", "135", "160", "175", "175（+25）", "175（+15）",
              "休日勤務135", "超勤代休時間の勤務", "超勤代休（+25）", "超勤代休（+15）"},
    手当の列 = {"125", "150", "135", "160", "175", "休日勤務135"},      // 命令簿の手当対象時間の欄
    手当 = (m as number) => Number.RoundDown((m + 30) / 60) * 60,
    ファイル = Table.SelectRows(休暇申請シート, each [超勤] <> null and [出勤用ID] <> ""),
    月にする = (f as record) as list =>
        let
            t = List.Buffer(List.Select(Table.ToRecords(Table.Buffer(f[超勤])), each 文字にする(Record.FieldOrDefault(_, "年月")) <> "" and 日付にする(Record.FieldOrDefault(_, "日付")) <> null)),
            月 = List.Distinct(List.Transform(t, each 文字にする(_[年月]))),
            一つ = (ym as text) as record =>
                let
                    r = List.Buffer(List.Select(t, each 文字にする(_[年月]) = ym)),
                    計 = List.Transform(時間の列, (c) => List.Sum(List.Transform(r, each 分にする(Record.FieldOrDefault(_, c))))),
                    計の記録 = Record.FromList(List.Transform(計, each _ / 1440), 時間の列),
                    手当の記録 = Record.FromList(List.Transform(手当の列, (c) => 手当(List.Sum(List.Transform(r, each 分にする(Record.FieldOrDefault(_, c))))) / 1440),
                        List.Transform(手当の列, each "手当対象 " & _)),
                    確認の日 = 日付にする(Record.FieldOrDefault(r{0}, "確認日")),
                    日数 = List.Count(List.Select(r, (x) => 文字にする(Record.FieldOrDefault(x, "登庁時刻")) <> "" or 文字にする(Record.FieldOrDefault(x, "退庁時刻")) <> ""
                        or List.Sum(List.Transform(時間の列, (c) => 分にする(Record.FieldOrDefault(x, c)))) > 0)),
                    超60 = [#"60時間超（+25）" = List.Sum(List.Transform(r, each 分にする(_[#"150（+25）"]) + 分にする(_[#"175（+25）"]))) / 1440,
                           #"60時間超（+15）" = List.Sum(List.Transform(r, each 分にする(_[#"150（+15）"]) + 分にする(_[#"175（+15）"]))) / 1440]
                in
                    Record.Combine({[年月 = ym, 出勤用ID = 正規ID(f[出勤用ID]), 氏名 = f[氏名], 所属部課 = 文字にする(Record.FieldOrDefault(r{0}, "所属部課")),
                                     状況 = if 確認の日 = null then "確認待ち" else "確認済み", 確認日 = 確認の日, 超過勤務の日数 = 日数],
                                    計の記録, 手当の記録, 超60, [ファイル名 = f[ファイル名]]})
        in
            List.Transform(月, 一つ),
    行 = List.Combine(List.Transform(Table.ToRecords(ファイル), (f) => try 月にする(f) otherwise {})),
    型 = type table [年月 = text, 出勤用ID = text, 氏名 = text, 所属部課 = text, 状況 = text, 確認日 = nullable date, 超過勤務の日数 = number,
                    #"125" = number, #"150" = number, #"150（60H前）" = number, #"150（+25）" = number, #"150（+15）" = number,
                    #"135" = number, #"160" = number, #"175" = number, #"175（+25）" = number, #"175（+15）" = number,
                    休日勤務135 = number, 超勤代休時間の勤務 = number, #"超勤代休（+25）" = number, #"超勤代休（+15）" = number,
                    #"手当対象 125" = number, #"手当対象 150" = number, #"手当対象 135" = number, #"手当対象 160" = number,
                    #"手当対象 175" = number, #"手当対象 休日勤務135" = number,
                    #"60時間超（+25）" = number, #"60時間超（+15）" = number, ファイル名 = text],
    表 = if List.IsEmpty(行) then #table(型, {}) else Table.FromRecords(行, 型)
in
    Table.Sort(表, {{"年月", Order.Descending}, {"出勤用ID", Order.Ascending}})
