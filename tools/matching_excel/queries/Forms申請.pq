// 休暇申請フォームの回答（「Forms申請ファイル」の Excel）を読み、申請用紙の申請データと同じ列にする（接続のみ）。
// 申請者は回答者のメールアドレスを職員名簿（メールアドレスの列）で出勤用ID にする。
// 承認は休暇承認フォームの回答（Forms承認）から、キー（出勤用ID|開始日|休暇の種類）のいちばん新しい回答を使う。
let
    設定値 = (名前 as text) =>
        let v = try Excel.CurrentWorkbook(){[Name = 名前]}[Content]{0}[Column1] otherwise null
        in if v = null then "" else Text.Trim(Text.Replace(Text.From(v), """", "")),
    数 = (v) => if v = null then 0 else try Number.From(v) otherwise 0,
    文字 = (v) => if v = null then "" else Text.Trim(Text.From(v)),
    日付 = (v) => if v = null then null
        else if v is date then v
        else if v is datetime then DateTime.Date(v)
        else if v is number then Date.From(v)
        else try Date.From(v) otherwise null,
    時刻 = (v) => if v = null then null
        else if v is datetime then v
        else if v is number then DateTime.From(v)
        else try DateTime.From(v) otherwise null,
    正規ID = (id) => let t = Text.Trim(if id = null then "" else Text.From(id))
        in if t = "" then "" else try Text.From(Number.From(t)) otherwise t,
    回答の表 = (path as text) as table =>
        let
            ブック = Excel.Workbook(File.Contents(path), null, true),
            表 = Table.SelectRows(ブック, each [Kind] = "Table"),
            シート = Table.SelectRows(ブック, each [Kind] = "Sheet")
        in
            if not Table.IsEmpty(表) then 表{0}[Data] else Table.PromoteHeaders(シート{0}[Data], [PromoteAllScalars = true]),
    列名 = (t as table, 名前 as text) =>
        let
            一覧 = Table.ColumnNames(t),
            完全 = List.Select(一覧, each Text.Trim(_) = 名前),
            前方 = List.Select(一覧, each Text.StartsWith(Text.Trim(_), 名前))
        in
            if not List.IsEmpty(完全) then 完全{0} else if not List.IsEmpty(前方) then 前方{0} else null,
    値 = (r as record, c) => if c = null then null else Record.Field(r, c),

    // 職員名簿：メールアドレス（小文字）→ 出勤用ID・氏名
    名簿 = try Table.PromoteHeaders(Excel.CurrentWorkbook(){[Name = "職員名簿表"]}[Content], [PromoteAllScalars = true])
        otherwise #table({"出勤用ID", "氏名", "メールアドレス"}, {}),
    名簿の行 = List.Select(Table.ToRecords(名簿), each 文字(Record.FieldOrDefault(_, "メールアドレス")) <> ""),
    名簿を引く = (メール as text) =>
        let hit = List.Select(名簿の行, each Text.Lower(文字([メールアドレス])) = Text.Lower(メール))
        in if List.IsEmpty(hit) then null else hit{0},

    // 承認：キーごとにいちばん新しい回答
    承認 = Table.Sort(Forms承認, {{"回答の時刻", Order.Descending}}),
    承認を引く = (キー as text) =>
        let hit = Table.SelectRows(承認, each [キー] = キー)
        in if Table.IsEmpty(hit) then null else hit{0},

    型 = type table [ファイル名 = text, 出勤用ID = text, 氏名 = text, 申請日 = nullable date,
        休暇の種類 = text, #"種別・内容" = text, 単位 = text, 開始日 = nullable date, 終了日 = nullable date,
        日数 = nullable number, 時間 = nullable number, 時刻 = text, 事由 = text, #"承認（所属長）" = text,
        承認日 = nullable date, 所属長コメント = text, 確認 = text, 分 = number, 年 = nullable number, 月 = nullable number],
    ファイル = 設定値("Forms申請ファイル"),
    表 = if ファイル = "" then null else try 回答の表(ファイル) otherwise null,
    行 = if 表 = null then {} else
        let
            c = [回答ID = 列名(表, "ID"), 完了 = 列名(表, "完了時刻"), メール = 列名(表, "メール"), 名前 = 列名(表, "名前"),
                 種類 = 列名(表, "休暇の種類"), 特別 = 列名(表, "特別休暇の種別"), 免除 = 列名(表, "免除の内容"),
                 単位 = 列名(表, "取得の単位"), 開始 = 列名(表, "開始日"), 終了 = 列名(表, "終了日"), 日数 = 列名(表, "日数"),
                 取得日 = 列名(表, "取得日"), 時間 = 列名(表, "時間"), 分 = 列名(表, "分"), 時刻 = 列名(表, "時刻"),
                 事由 = 列名(表, "事由")]
        in
            List.Transform(Table.ToRecords(表), (r) =>
                let
                    // レコードの中では同じ名前のフィールドが優先されるので、フィールドと違う名前にしておく
                    メール = 文字(値(r, c[メール])),
                    職員 = 名簿を引く(メール),
                    id = if 職員 = null then "" else 正規ID(職員[出勤用ID]),   // 承認との照合用（表示は職員名簿の書き方のまま）
                    種類 = 文字(値(r, c[種類])),
                    単位の値 = 文字(値(r, c[単位])),
                    時間単位 = Text.Contains(単位の値, "時間"),
                    開始 = if 時間単位 then 日付(値(r, c[取得日])) else 日付(値(r, c[開始])),
                    終了の値 = if 時間単位 then null else 日付(値(r, c[終了])),
                    日数の値 = if 時間単位 then 0 else 数(値(r, c[日数])),
                    分の値 = if 時間単位 then Number.Round(数(値(r, c[時間])) * 60 + 数(値(r, c[分]))) else 0,
                    内容 = if 種類 = "特別休暇（看護休暇以外）" then 文字(値(r, c[特別]))
                        else if 種類 = "職務専念義務の免除" then 文字(値(r, c[免除])) else "",
                    キー = id & "|" & (if 開始 = null then "" else Date.ToText(開始, "yyyy-MM-dd")) & "|" & 種類,
                    承 = 承認を引く(キー),
                    完了 = 時刻(値(r, c[完了])),
                    注意 = List.Select({
                        if 職員 = null then "職員名簿にメールアドレス（" & メール & "）がありません" else "",
                        if 開始 = null then "日付がありません" else "",
                        if not 時間単位 and 日数の値 <= 0 then "日数がありません" else "",
                        if 時間単位 and 分の値 <= 0 then "時間がありません" else "",
                        if 終了の値 <> null and 開始 <> null and 終了の値 < 開始 then "終了日が開始日より前です" else "",
                        if 時間単位 and 種類 = "年次休暇" and Number.Mod(分の値, 15) <> 0 then "年次休暇の時間は15分単位です" else "",
                        if 時間単位 and List.Contains({"介護時間", "育児時間"}, 種類) and Number.Mod(分の値, 30) <> 0 then "30分単位です" else "",
                        if 終了の値 <> null and 開始 <> null and List.Contains({"年次休暇", "病気休暇", "看護休暇（子の看護等）"}, 種類)
                            and Date.ToText(開始, "yyyyMM") <> Date.ToText(終了の値, "yyyyMM") then "月ごとに分けて申請してください" else ""
                    }, each _ <> "")
                in
                    [ファイル名 = "Forms（回答 " & 文字(値(r, c[回答ID])) & "）", 出勤用ID = if 職員 = null then "" else 文字(職員[出勤用ID]),
                     氏名 = if 職員 = null then 文字(値(r, c[名前])) else 文字(職員[氏名]),
                     申請日 = if 完了 = null then null else DateTime.Date(完了),
                     休暇の種類 = 種類, #"種別・内容" = 内容, 単位 = if 時間単位 then "時間単位" else "1日単位",
                     開始日 = 開始, 終了日 = 終了の値,
                     日数 = if 日数の値 = 0 then null else 日数の値,
                     時間 = if 分の値 = 0 then null else 分の値 / 1440,
                     時刻 = 文字(値(r, c[時刻])), 事由 = 文字(値(r, c[事由])),
                     #"承認（所属長）" = if 承 = null then "" else 承[結果],
                     承認日 = if 承 = null then null else 承[承認日],
                     所属長コメント = if 承 = null then "" else Text.Combine(List.Select({承[承認者], 承[コメント]}, each _ <> ""), "："),
                     確認 = Text.Combine(List.Transform(注意, each "・" & _), ""),
                     分 = Number.Round(日数の値 * 465 + 分の値),
                     年 = if 開始 = null then null else Date.Year(開始),
                     月 = if 開始 = null then null else Date.Month(開始)])
in
    Table.FromRecords(行, 型)
