// 「QR出勤の読み込み」シートに出す：フォルダの出勤記録のファイルを全部つなぎ、同じ行（同じ出勤用ID・出勤日・出勤時刻・登録日時）を1つにしたもの。
// ［Excel出力］はいつも全件なので、何回分のファイルを置いても重ならない。出勤日・出勤時刻の古い順。
// 値は文字（出勤日 yyyy-MM-dd、出勤時刻 HH:mm:ss、登録日時 yyyy-MM-ddTHH:mm:ss）。［最新を読み込む］のマクロが「出力データ貼付」に写す。
let
    文字 = (v) => if v = null then "" else Text.Trim(Text.From(v)),
    日 = (v) => if v = null then "" else if v is date then Date.ToText(v, "yyyy-MM-dd")
        else if v is datetime then Date.ToText(DateTime.Date(v), "yyyy-MM-dd")
        else if v is number then Date.ToText(Date.From(v), "yyyy-MM-dd") else 文字(v),
    時 = (v) => if v = null then "" else if v is time then Time.ToText(v, "HH:mm:ss")
        else if v is datetime then Time.ToText(DateTime.Time(v), "HH:mm:ss")
        else if v is number then Time.ToText(Time.From(v - Number.RoundDown(v)), "HH:mm:ss") else 文字(v),
    日時 = (v) => if v = null then "" else if v is datetime then DateTime.ToText(v, "yyyy-MM-ddTHH:mm:ss")
        else if v is datetimezone then DateTimeZone.ToText(v, "yyyy-MM-ddTHH:mm:ssZ")
        else if v is number then DateTime.ToText(DateTime.From(v), "yyyy-MM-ddTHH:mm:ss") else 文字(v),
    読めた = Table.SelectRows(QR出勤_ファイル, each [読み取り] = "OK" and [行] <> null),
    型 = type table [出勤用ID = text, QR設置場所 = text, 出勤日 = text, 出勤時刻 = text, 登録日時 = text],
    行 = List.Combine(List.Transform(読めた[行], (t) => List.Transform(Table.ToRecords(t), (r) =>
        [出勤用ID = 文字(r[出勤用ID]), QR設置場所 = 文字(r[QR設置場所]), 出勤日 = 日(r[出勤日]), 出勤時刻 = 時(r[出勤時刻]), 登録日時 = 日時(r[登録日時])]))),
    使う = List.Select(行, each [出勤用ID] <> "" and [出勤日] <> ""),
    表 = if List.IsEmpty(使う) then #table(型, {}) else Table.Distinct(Table.FromRecords(使う, 型))
in
    Table.Sort(表, {{"出勤日", Order.Ascending}, {"出勤時刻", Order.Ascending}, {"出勤用ID", Order.Ascending}})
