// 休暇申請シートの超過勤務等命令簿（月のシートの表「超勤日別」。tools/overtime）を、登庁・退庁の時刻か超過勤務の時間がある日の1日1行にする。
// 表の時刻・時間は分なので、Excel の時間（1日＝1）にする（25:00 は 1.0417）。支給割合ごとの時間は命令簿の式で計算した値をそのまま使う。
// 「超過勤務の読み込み」シートに出し、［超過勤務を転記］のマクロが所属長の確認済みの月をここから写す。
let
    正規ID = (id) => let t = Text.Trim(if id = null then "" else Text.From(id))
        in if t = "" then "" else try Text.From(Number.From(t)) otherwise t,
    文字にする = (v) => if v = null then "" else Text.Trim(Text.From(v)),
    分にする = (v) => if v = null or 文字にする(v) = "" then null else try Number.From(v) otherwise null,
    時間にする = (v) => let m = 分にする(v) in if m = null then null else m / 1440,
    日付にする = (v) => if 文字にする(v) = "" then null else try Date.FromText(文字にする(v), [Format = "yyyy-MM-dd"]) otherwise null,
    時刻の列 = {"登庁時刻", "退庁時刻", "休憩時間"},
    印の列 = {"振替日（1日）", "振替日（1日・日）", "振替日（午前）", "振替日（午後）", "休日・代休日"},
    時間の列 = {"125", "150", "150（60H前）", "150（+25）", "150（+15）", "135", "160", "175", "175（+25）", "175（+15）",
              "休日勤務135", "超勤代休時間の勤務", "超勤代休（+25）", "超勤代休（+15）"},
    ファイル = Table.SelectRows(休暇申請シート, each [超勤] <> null and [出勤用ID] <> ""),
    // 1人の表 → 日のレコード
    日にする = (f as record) as list =>
        let
            t = List.Buffer(Table.ToRecords(Table.Buffer(f[超勤]))),
            月の行 = List.Select(t, each 文字にする(Record.FieldOrDefault(_, "年月")) <> "" and 日付にする(Record.FieldOrDefault(_, "日付")) <> null),
            一つ = (r as record) as record =>
                let
                    基本 = [年月 = 文字にする(r[年月]), 出勤用ID = 正規ID(f[出勤用ID]), 氏名 = f[氏名],
                           所属部課 = 文字にする(Record.FieldOrDefault(r, "所属部課")),
                           確認日 = 日付にする(Record.FieldOrDefault(r, "確認日")), 日付 = 日付にする(r[日付])],
                    時刻 = Record.FromList(List.Transform(時刻の列, each 時間にする(Record.FieldOrDefault(r, _))), 時刻の列),
                    印 = Record.FromList(List.Transform(印の列, each 文字にする(Record.FieldOrDefault(r, _))), 印の列),
                    量 = Record.FromList(List.Transform(時間の列, each let m = 分にする(Record.FieldOrDefault(r, _)) in if m = null then 0 else m / 1440), 時間の列),
                    残り = [備考 = 文字にする(Record.FieldOrDefault(r, "備考")), ファイル名 = f[ファイル名]]
                in
                    Record.Combine({基本, 時刻, 印, 量, 残り}),
            ある日 = List.Select(List.Transform(月の行, 一つ), (r) =>
                r[登庁時刻] <> null or r[退庁時刻] <> null or List.Sum(Record.FieldValues(Record.SelectFields(r, 時間の列))) > 0)
        in
            ある日,
    行 = List.Combine(List.Transform(Table.ToRecords(ファイル), (f) => try 日にする(f) otherwise {})),
    型 = type table [年月 = text, 出勤用ID = text, 氏名 = text, 所属部課 = text, 確認日 = nullable date, 日付 = date,
                    登庁時刻 = nullable number, 退庁時刻 = nullable number, 休憩時間 = nullable number,
                    #"振替日（1日）" = text, #"振替日（1日・日）" = text, #"振替日（午前）" = text, #"振替日（午後）" = text, #"休日・代休日" = text,
                    #"125" = number, #"150" = number, #"150（60H前）" = number, #"150（+25）" = number, #"150（+15）" = number,
                    #"135" = number, #"160" = number, #"175" = number, #"175（+25）" = number, #"175（+15）" = number,
                    休日勤務135 = number, 超勤代休時間の勤務 = number, #"超勤代休（+25）" = number, #"超勤代休（+15）" = number,
                    備考 = text, ファイル名 = text],
    表 = if List.IsEmpty(行) then #table(型, {}) else Table.FromRecords(行, 型)
in
    Table.Sort(表, {{"年月", Order.Descending}, {"出勤用ID", Order.Ascending}, {"日付", Order.Ascending}})
