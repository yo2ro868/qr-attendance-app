// 「テレワーク報告」シートに出す：休暇申請シートのテレワーク実施報告（表「テレワーク報告」。tools/telework/report_sheets.py）の全員分。
// 状況：所属長の確認日があれば「確認済み」、なければ「確認待ち」。
// 注意：報告の入力の不備（休暇申請シートの「確認」の欄）、テレワーク予定表に予定がない日、提出から1週間たっても確認がないもの。
let
    文字にする = (v) => if v = null then "" else Text.Trim(Text.From(v)),
    正規ID = (id) => let t = 文字にする(id) in if t = "" then "" else try Text.From(Number.From(t)) otherwise t,
    日付にする = (v) => if v = null then null else if v is date then v else if v is datetime then DateTime.Date(v)
        else if v is number then Date.From(v) else try Date.From(v) otherwise null,
    時刻にする = (v) => if v = null then null else if v is time then v else if v is datetime then DateTime.Time(v)
        else if v is number then Time.From(v - Number.RoundDown(v)) else try Time.From(v) otherwise null,
    分 = (t as time) => Time.Hour(t) * 60 + Time.Minute(t),
    // 時間：開始〜終了から 12:00〜13:00 の休憩を除く（テレワーク予定表と同じ）
    時間の分 = (s, e) => if s = null or e = null or e <= s then null
        else 分(e) - 分(s) - List.Max({0, List.Min({分(e), 13 * 60}) - List.Max({分(s), 12 * 60})}),
    今日 = Date.From(DateTime.LocalNow()),
    予定のキー = List.Buffer(テレワーク_日別[キー]),
    ファイル = Table.SelectRows(休暇申請シート, each [テレワーク報告] <> null and [出勤用ID] <> ""),
    行 = List.Combine(List.Transform(Table.ToRecords(ファイル), (f) =>
        let
            書いた行 = List.Select(Table.ToRecords(f[テレワーク報告]), (r) =>
                List.AnyTrue(List.Transform({"提出日", "実施日", "開始時刻", "終了時刻", "業務内容", "備考"},
                    (c) => 文字にする(Record.FieldOrDefault(r, c)) <> "")))
        in
            List.Transform(書いた行, (r) =>
                let
                    id = 正規ID(f[出勤用ID]),
                    提出 = 日付にする(r[提出日]), 実施 = 日付にする(r[実施日]),
                    始め = 時刻にする(r[開始時刻]), 終り = 時刻にする(r[終了時刻]),
                    確認 = 日付にする(r[#"確認日（所属長）"]),
                    v時間 = 時間の分(始め, 終り),
                    予定あり = 実施 <> null and List.Contains(予定のキー, id & "|" & Date.ToText(実施, "yyyy-MM-dd")),
                    注意の文 = 文字にする(r[確認])
                        & (if 実施 <> null and not 予定あり then "・テレワーク予定表にこの日の予定がありません" else "")
                        & (if 確認 = null and 提出 <> null and Duration.Days(今日 - 提出) >= 7 then "・提出から1週間以上、所属長の確認がありません" else "")
                in
                    [出勤用ID = id, 氏名 = f[氏名], 所属 = f[報告の所属], 番号 = r[番号], 提出日 = 提出, 実施日 = 実施,
                     開始時刻 = 始め, 終了時刻 = 終り, 時間 = if v時間 = null then null else v時間 / 1440,
                     業務内容 = 文字にする(r[業務内容]), 備考 = 文字にする(r[備考]), 確認日 = 確認,
                     所属長コメント = 文字にする(r[所属長コメント]), 状況 = if 確認 = null then "確認待ち" else "確認済み",
                     注意 = 注意の文, ファイル名 = f[ファイル名]]))),
    型 = type table [出勤用ID = text, 氏名 = text, 所属 = text, 番号 = nullable number, 提出日 = nullable date, 実施日 = nullable date,
                    開始時刻 = nullable time, 終了時刻 = nullable time, 時間 = nullable number, 業務内容 = text, 備考 = text,
                    確認日 = nullable date, 所属長コメント = text, 状況 = text, 注意 = text, ファイル名 = text],
    表 = if List.IsEmpty(行) then #table(型, {}) else Table.FromRecords(行, 型)
in
    Table.Sort(表, {{"実施日", Order.Descending}, {"出勤用ID", Order.Ascending}})
