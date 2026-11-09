// 休暇承認フォームの回答（「Forms承認ファイル」の Excel）を読む（接続のみ）。
// 1回答1行：キー（出勤用ID|開始日 yyyy-MM-dd|休暇の種類）・結果（承認／却下／取消）・承認日・承認者・コメント・回答の時刻。
// 同じキーに何度か回答があれば、いちばん新しい回答が使われる（Forms申請 で）。
let
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
    // 見出しで列を探す（完全一致 → 前方一致）
    列名 = (t as table, 名前 as text) =>
        let
            一覧 = Table.ColumnNames(t),
            完全 = List.Select(一覧, each Text.Trim(_) = 名前),
            前方 = List.Select(一覧, each Text.StartsWith(Text.Trim(_), 名前))
        in
            if not List.IsEmpty(完全) then 完全{0} else if not List.IsEmpty(前方) then 前方{0} else null,
    値 = (r as record, c) => if c = null then null else Record.Field(r, c),
    型 = type table [キー = text, 結果 = text, 承認日 = nullable date, 承認者 = text, コメント = text,
        回答の時刻 = nullable datetime, 出勤用ID = text, 開始日 = nullable date, 休暇の種類 = text],
    表 = Forms回答_承認[表],
    行 = if 表 = null then {} else
        let
            c = [ID = 列名(表, "申請者の出勤用ID"), 開始 = 列名(表, "休暇の開始日"), 種類 = 列名(表, "休暇の種類"),
                 結果 = 列名(表, "結果"), コメント = 列名(表, "コメント"), 名前 = 列名(表, "名前"), 完了 = 列名(表, "完了時刻")]
        in
            List.Transform(Table.ToRecords(表), (r) =>
                let
                    id = 正規ID(値(r, c[ID])),
                    開始 = 日付(値(r, c[開始])),
                    種類 = 文字(値(r, c[種類])),
                    完了 = 時刻(値(r, c[完了]))
                in
                    [キー = id & "|" & (if 開始 = null then "" else Date.ToText(開始, "yyyy-MM-dd")) & "|" & 種類,
                     結果 = 文字(値(r, c[結果])), 承認日 = if 完了 = null then null else DateTime.Date(完了),
                     承認者 = 文字(値(r, c[名前])), コメント = 文字(値(r, c[コメント])), 回答の時刻 = 完了,
                     出勤用ID = id, 開始日 = 開始, 休暇の種類 = 種類])
in
    Table.FromRecords(行, 型)
