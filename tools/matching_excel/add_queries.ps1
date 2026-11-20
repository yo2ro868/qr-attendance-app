# 出勤・休暇管理データ に Power Query を入れる：申請・報告シート（tools/leave_sheet）のフォルダ・テレワーク予定表などを読み、
# 「申請・報告シートの読み込み」シートに 年休_自動・月別_自動・日別_自動 の表を、「申請一覧」シートに申請一覧の表を作る。開いたときに読み直す設定にする。
# 使い方：add_queries.ps1 -Path 出勤・休暇管理データ.xlsx [-Folder 申請・報告シートのフォルダ]
#   -Folder などを付けると、それを読んだ状態で保存する（確認用）。付けなければ設定欄は空欄のまま。
#   -TeleworkFolder はテレワーク予定表（tools/telework）のフォルダ。テレワーク一覧シートと、出勤簿の「テレ」に使う。
param([Parameter(Mandatory = $true)][string]$Path, [string]$Folder = '', [string]$TeleworkFolder = '')

$ErrorActionPreference = 'Stop'
$Path = [IO.Path]::GetFullPath($Path)
$sheetQueries = Join-Path $PSScriptRoot '..\leave_sheet\queries'   # 申請・報告シートを読むクエリ（休暇申請シート・申請データ）は tools/leave_sheet/queries に置く
$teleworkQueries = Join-Path $PSScriptRoot '..\telework\queries'    # テレワーク予定表を読むクエリ
$queries = Join-Path $PSScriptRoot 'queries'
# 接続のみのクエリ（ほかのクエリから使う）。名前 → ファイル
$helpers = [ordered]@{
    '休暇申請シート'       = Join-Path $sheetQueries '休暇申請シート.pq'
    '祝日の日付'     = Join-Path $queries '祝日の日付.pq'
    '職員名簿_メール' = Join-Path $queries '職員名簿_メール.pq'
    '申請データ_用紙' = Join-Path $sheetQueries '申請データ.pq'
    '申請データ'     = Join-Path $queries '申請データ.pq'
    'テレワーク予定表' = Join-Path $teleworkQueries 'テレワーク予定表.pq'
    'フレックス_確定' = Join-Path $queries 'フレックス_確定.pq'     # 表 確定_フレックス日別（下で作る）を読む
    'QR出勤_ファイル' = Join-Path $queries 'QR出勤_ファイル.pq'       # 出力データ貼付 H3 のフォルダの attendance.xlsx
}
# 超過勤務の支給割合ごとの時間の列（超勤_日別.pq・超勤_月.pq と同じ）と、手当対象時間の列
$otBuckets = @('125', '150', '150（60H前）', '150（+25）', '150（+15）', '135', '160', '175', '175（+25）', '175（+15）',
               '休日勤務135', '超勤代休時間の勤務', '超勤代休（+25）', '超勤代休（+15）')
$otAllowance = @('125', '150', '135', '160', '175', '休日勤務135') | ForEach-Object { "手当対象 $_" }
# シートに出す表：シート・置く場所（見出しの行）・行数を書くセル・時間:分／日付の書式を付ける列
$loaded = [ordered]@{
    '年休_自動' = @{ Sheet = '申請・報告シートの読み込み'; At = 'A8'; Count = 'A6'; Time = @(); Date = @() }
    '月別_自動' = @{ Sheet = '申請・報告シートの読み込み'; At = 'E8'; Count = 'E6'; Time = (1..12 | ForEach-Object { "${_}月" }); Date = @() }
    '日別_自動' = @{ Sheet = '申請・報告シートの読み込み'; At = 'S8'; Count = 'S6'; Time = @(); Date = @() }
    '申請一覧'  = @{ Sheet = '申請一覧'; At = 'A3'; Count = ''; Time = @('時間'); Date = @('申請日', '開始日', '終了日', '承認日') }
    # テレワーク：一覧は注意の確認用、日別は出勤簿の出勤欄・備考用（一覧を参照するので一覧の後に作る）
    'テレワーク一覧'  = @{ Sheet = 'テレワーク一覧'; At = 'A5'; Count = ''; Time = @('開始時刻', '終了時刻', '時間'); Date = @('日付') }
    'テレワーク_日別' = @{ Sheet = '申請・報告シートの読み込み'; At = 'AD8'; Count = 'AD6'; Time = @(); Date = @() }
    # 超過勤務等命令簿（申請・報告シートの月のシート）：1人1か月の計と確認、日ごとの行
    '超勤_月'   = @{ Sheet = '超過勤務の読み込み'; At = 'A3'; Count = ''; Time = ($otBuckets + $otAllowance + @('60時間超（+25）', '60時間超（+15）')); Date = @('確認日') }
    '超勤_日別' = @{ Sheet = '超過勤務の読み込み'; At = 'AG3'; Count = ''; Time = (@('登庁時刻', '退庁時刻', '休憩時間') + $otBuckets); Date = @('確認日', '日付') }
    # フレックスタイム制（申請・報告シートのフレックス申告・割振り）：今の内容（日ごと）、単位期間ごとの状況と取り込み、出勤簿の備考
    'フレックス_日別' = @{ Sheet = 'フレックスの読み込み'; At = 'A3'; Count = ''; Time = @('勤務開始', '勤務終了', '休憩', '勤務時間'); Date = @('期間の開始日', '日付', '申告年月日', '割振年月日') }
    'フレックス一覧'  = @{ Sheet = 'フレックス'; At = 'A3'; Count = ''; Time = @('合計'); Date = @('開始日', '終了日', '申告年月日', '割振年月日') }
    'フレックス_出勤簿' = @{ Sheet = '申請・報告シートの読み込み'; At = 'AH8'; Count = 'AH6'; Time = @(); Date = @() }
    # 利用予定シート（フレックス／テレワーク利用予定）の日ごとの文字：休暇・テレワーク・フレックスの勤務開始
    # テレワーク実施報告（申請・報告シートのテレワーク報告）：全員の報告と所属長の確認・注意
    'テレワーク報告一覧' = @{ Sheet = 'テレワーク報告'; At = 'A3'; Count = ''; Time = @('開始時刻', '終了時刻', '時間'); Date = @('提出日', '実施日', '確認日') }
    '利用予定_日別' = @{ Sheet = '申請・報告シートの読み込み'; At = 'AK8'; Count = 'AK6'; Time = @(); Date = @() }
    # 出勤管理システムの［Excel出力］のファイル（フォルダで読む）と、つないだ出勤記録（マクロが出力データ貼付に写す）
    'QR出勤ファイル一覧' = @{ Sheet = 'QR出勤の読み込み'; At = 'A5'; Count = ''; Time = @(); Date = @() }
    'QR出勤' = @{ Sheet = 'QR出勤の読み込み'; At = 'G5'; Count = ''; Time = @(); Date = @() }
}
$settings = @{ 'B2' = $Folder }
# 折り返して表示する長い文の列と、その幅
$wrapWidth = @{ '注意' = 50; '確認' = 40; '業務内容' = 40; '備考' = 24; '所属長コメント' = 20; '事由' = 20; '休憩の時刻' = 22; '状態' = 22 }

$xl = New-Object -ComObject Excel.Application
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($Path)
    # プライバシー レベルを無視する設定（FastCombine）はほかのPCに引き継がれないので使わない。クエリの作りで止まらないようにしてある
    foreach ($name in $helpers.Keys) {
        [void]$wb.Queries.Add($name, [IO.File]::ReadAllText($helpers[$name], [Text.Encoding]::UTF8))
    }
    foreach ($name in $loaded.Keys) {
        [void]$wb.Queries.Add($name, [IO.File]::ReadAllText((Join-Path $queries "$name.pq"), [Text.Encoding]::UTF8))
    }
    $cfg = $wb.Worksheets.Item('申請・報告シートの読み込み')
    foreach ($k in $settings.Keys) { if ($settings[$k]) { $cfg.Range($k).Value2 = $settings[$k] } }
    $tw = $wb.Worksheets.Item('テレワーク一覧')
    if ($TeleworkFolder) { $tw.Range('B2').Value2 = $TeleworkFolder }
    foreach ($name in $loaded.Keys) {
        $spec = $loaded[$name]
        $ws = $wb.Worksheets.Item($spec.Sheet)
        # 表を作る前の列の幅（表を作ると見出しに合わせて狭くなることがあるので、あとでこれより狭くしない）
        $c0 = $ws.Range($spec.At).Column
        $widthBefore = @{}
        for ($k = 0; $k -lt 60; $k++) { $widthBefore[$c0 + $k] = [double]$ws.Columns.Item($c0 + $k).ColumnWidth }
        $conn = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=`$Workbook`$;Location=`"$name`";Extended Properties=`"`""
        $lo = $ws.ListObjects.Add(0, $conn, $true, 1, $ws.Range($spec.At))
        $lo.Name = $name
        $lo.TableStyle = 'TableStyleLight9'
        $qt = $lo.QueryTable
        $qt.CommandType = 2
        $qt.CommandText = "SELECT * FROM [$name]"
        $qt.BackgroundQuery = $false
        $qt.PreserveFormatting = $true
        $qt.AdjustColumnWidth = ($name -eq '申請一覧' -or $name -eq 'テレワーク報告一覧' -or $name -eq 'テレワーク一覧' -or $name -like '超勤_*' -or ($name -like 'フレックス*' -and $name -ne 'フレックス_出勤簿'))
        "読み込み中: $name " + (Get-Date -Format 'HH:mm:ss') | Out-Host      # どのクエリで止まったか分かるように
        $qt.Refresh($false) | Out-Null
        # 長い文の列は幅を決めて折り返す（読み直すたびに列の幅を変えない。行の高さは［最新を読み込む］のマクロが合わせる）
        $qt.AdjustColumnWidth = $false
        foreach ($lc in @($lo.ListColumns)) {
            # 読み込む前でも、名前・日付・ファイル名などが入る幅にしておく（読み込んだあとは［最新を読み込む］のマクロが中身に合わせる）
            $col = $ws.Columns.Item($lc.Range.Column)
            $min = if ($lc.Name -eq 'ファイル名') { 34 } elseif ($lc.Name -in @('氏名', '所属')) { 14 } else { 11 }
            if ($widthBefore.ContainsKey($lc.Range.Column) -and $widthBefore[$lc.Range.Column] -gt $min) { $min = $widthBefore[$lc.Range.Column] }
            if ($col.ColumnWidth -lt $min) { $col.ColumnWidth = [double]$min }
            if ($wrapWidth.ContainsKey($lc.Name)) {
                $ws.Columns.Item($lc.Range.Column).ColumnWidth = [double]$wrapWidth[$lc.Name]
                $ws.Columns.Item($lc.Range.Column).WrapText = $true
            }
        }
        # 開いたときには読み直さない（全部のクエリを読み直すと、読む場所が空でも1分ほどかかるため）。
        # 目次などの［最新を読み込む］（ImportRequests）を押したときに読み直す
        $lo.QueryTable.WorkbookConnection.OLEDBConnection.RefreshOnFileOpen = $false
        # 書式は列全体に付ける（読み直して増えた行にも付く。空のセルごとに書式を持たせるとファイルが何MBも大きくなる）
        foreach ($pair in @(@($spec.Time, '[h]:mm'), @($spec.Date, 'yyyy/m/d'))) {
            foreach ($col in $pair[0]) {
                $ws.Columns.Item($lo.ListColumns.Item($col).Range.Column).NumberFormat = $pair[1]
            }
        }
        # 行数（出勤簿などがこの範囲だけを探す）。表の名前で書くので、読み直して行が増減しても合う。列全体の書式のあとに付ける
        if ($spec.Count) {
            $ws.Range($spec.Count).Formula = "=COUNTA($name[キー])"
            $ws.Range($spec.Count).NumberFormat = '"（"0"行）"'
        }
    }
    # ----- ［超過勤務を転記］で写す先の表（値。照合の列だけは式）。見出しは読み込んだ表と同じ＋転記日時 -----
    $au = "'申請・報告シートの読み込み'"
    $match = { param($key, $col) "IFERROR(MATCH($key,INDEX($au!`$${col}:`$${col},9):INDEX($au!`$${col}:`$${col},8+MAX(1,$au!`$${col}`$6)),0)+8,"""")" }
    $dayKey = '[@出勤用ID]&"|"&TEXT([@日付],"yyyy-mm-dd")'
    $qr = 'COUNTIFS(照合結果!$G$2:$G$5001,[@出勤用ID]&"",照合結果!$H$2:$H$5001,TEXT([@日付],"yyyy-mm-dd"))'
    $leaveRow = & $match $dayKey 'S'
    $teleRow = & $match $dayKey 'AD'
    $checks = [ordered]@{
        '出勤の記録' = "=IF([@出勤用ID]="""","""",IF($qr>0,""QR"","""")&IF($teleRow="""","""",""テレ""))"
        '休暇'       = "=IF([@出勤用ID]="""","""",IF($leaveRow="""","""",IF(COUNTIF(INDEX($au!`$T:`$AA,$leaveRow,0),""1日"")>0,""1日"",""時間"")))"
        '注意'       = '=IF([@出勤用ID]="","",IF(AND([@出勤の記録]="",[@登庁時刻]<>""),"出勤の記録（QR・テレワーク）がありません。","")&IF([@休暇]="1日","1日の休暇（承認済み）の日です。",IF([@休暇]="時間","時間単位の休暇（承認済み）があります（時刻を確かめてください）。","")))'
    }
    # フレックス：承認済みの単位期間を写した日ごとの行の照合（今日より前の日だけ）
    $flexChecks = [ordered]@{
        '出勤の記録' = $checks['出勤の記録']
        '休暇'       = $checks['休暇']
        '注意'       = '=IF(OR([@出勤用ID]="",[@日付]="",[@日付]>=TODAY()),"",IF(AND([@区分]="平日",N([@勤務時間])>0,[@出勤の記録]="",[@休暇]<>"1日"),"勤務を割り振った日ですが、出勤の記録（QR・テレワーク）も1日の休暇もありません。","")&IF(AND([@区分]="平日",N([@勤務時間])=0,ISNUMBER(SEARCH("QR",[@出勤の記録]))),"勤務時間を割り振らない日に QR の出勤登録があります。","")&IF(AND([@区分]="週休日",N([@勤務時間])>0),"週休日に勤務を割り振っています。",""))'
    }
    foreach ($dst in @(@{ Sheet = '超過勤務'; At = 'A4'; From = '超勤_月'; Name = '転記_月計'; Checks = $null; Stamp = '転記日時' },
                       @{ Sheet = '超過勤務の日別'; At = 'A3'; From = '超勤_日別'; Name = '転記_日別'; Checks = $checks; Stamp = '転記日時' },
                       @{ Sheet = 'フレックスの日別'; At = 'A3'; From = 'フレックス_日別'; Name = '確定_フレックス日別'; Checks = $flexChecks; Stamp = '確定日時' })) {
        $ws = $wb.Worksheets.Item($dst.Sheet)
        $srcSheet = $loaded[$dst.From].Sheet
        $srcLo = $wb.Worksheets.Item($srcSheet).ListObjects.Item($dst.From)
        $heads = @($srcLo.HeaderRowRange.Cells | ForEach-Object { [string]$_.Value2 }) + @($dst.Stamp)
        if ($dst.Checks) { $heads += @($dst.Checks.Keys) }
        $top = $ws.Range($dst.At)
        for ($i = 0; $i -lt $heads.Count; $i++) { $ws.Cells.Item($top.Row, $top.Column + $i).Value2 = [string]$heads[$i] }
        $rng = $ws.Range($top, $ws.Cells.Item($top.Row + 1, $top.Column + $heads.Count - 1))
        $lo = $ws.ListObjects.Add(1, $rng, $null, 1)
        $lo.Name = $dst.Name
        $lo.TableStyle = 'TableStyleLight9'
        if ($dst.Checks) { foreach ($k in $dst.Checks.Keys) { $lo.ListColumns.Item($k).DataBodyRange.Formula = $dst.Checks[$k] } }
        # 書式：読み込んだ表と同じ列に時間・日付、転記日時。出勤用ID・年月は文字
        $first = $top.Row + 1
        $spec = $loaded[$dst.From]
        foreach ($pair in @(@($spec.Time, '[h]:mm'), @($spec.Date, 'yyyy/m/d'), @(@($dst.Stamp), 'yyyy/m/d h:mm'), @(@('出勤用ID', '年月'), '@'))) {
            foreach ($col in $pair[0]) {
                if (-not ($heads -contains $col)) { continue }
                $c = $lo.ListColumns.Item($col).Range.Column
                $ws.Range($ws.Cells.Item($top.Row, $c), $ws.Cells.Item($top.Row + 30000, $c)).NumberFormat = $pair[1]
            }
        }
        # 見出しが入る幅に広げる（もとの幅より狭くはしない）
        foreach ($lc in @($lo.ListColumns)) {
            $col = $ws.Columns.Item($lc.Range.Column)
            $before = $col.ColumnWidth
            [void]$lc.Range.Columns.AutoFit()
            if ($col.ColumnWidth -lt $before) { $col.ColumnWidth = $before }
        }
        if ($dst.Checks) { $lo.ListColumns.Item('注意').Range.ColumnWidth = 50; $lo.ListColumns.Item('注意').Range.WrapText = $true }
        foreach ($col in @($spec.Date)) { $r0 = $lo.ListColumns.Item($col).Range; if ($r0.ColumnWidth -lt 11) { $r0.ColumnWidth = 11 } }
        $lo.ListColumns.Item($dst.Stamp).Range.ColumnWidth = 16
        $lo.ListColumns.Item('ファイル名').Range.ColumnWidth = 34
    }
    # 超過勤務シート（転記_月計）に、超過勤務の日別（転記_日別）の注意をまとめる（2026-11-20。日別のシートはふだん非表示）
    $ws = $wb.Worksheets.Item('超過勤務')
    $lo = $ws.ListObjects.Item('転記_月計')
    $nc = $lo.ListColumns.Add(); $nc.Name = '注意のある日'
    $nc.DataBodyRange.Formula = '=IF([@出勤用ID]="","",COUNTIFS(転記_日別[出勤用ID],[@出勤用ID],転記_日別[年月],[@年月],転記_日別[注意],"?*"))'
    $nc.Range.ColumnWidth = 14
    $ws.Range($nc.DataBodyRange.Cells(1), $ws.Cells.Item($nc.DataBodyRange.Row + 3000, $nc.Range.Column)).NumberFormat = '0'     # 隣の列の時間の書式を引き継がないように
    $nc = $lo.ListColumns.Add(); $nc.Name = '注意'
    $nc.DataBodyRange.Formula2 = '=IF([@出勤用ID]="","",TEXTJOIN("／",TRUE,IF((転記_日別[出勤用ID]=[@出勤用ID])*(転記_日別[年月]=[@年月])*(転記_日別[注意]<>""),TEXT(転記_日別[日付],"m/d")&" "&転記_日別[注意],"")))'
    $nc.Range.ColumnWidth = 60
    $nc.Range.WrapText = $true
    # 注意のある行を薄い赤に
    $body = $lo.DataBodyRange
    $col = $lo.ListColumns.Item('注意のある日').Range.Column
    $fc = $ws.Range($ws.Cells.Item($body.Row, 1), $ws.Cells.Item($body.Row + 3000, $lo.Range.Column + $lo.ListColumns.Count - 1)).FormatConditions.Add(2, 0, ('=N($' + (($ws.Cells.Item(1, $col).Address($false, $false)) -replace '\d', '') + $body.Row + ')>0'))
    $fc.Interior.Color = 0xD9E5FC
    # H2：B2 の月の確認待ちの人（所属部課 氏名）と、ほかの月の確認待ちの数
    $ws.Range('H2').Formula2 = '=IF(COUNTIFS(超勤_月[年月],TEXT($B$2,"yyyy-mm"),超勤_月[状況],"確認待ち")=0,"なし",COUNTIFS(超勤_月[年月],TEXT($B$2,"yyyy-mm"),超勤_月[状況],"確認待ち")&"人（"&TEXTJOIN("、",TRUE,IF((超勤_月[年月]=TEXT($B$2,"yyyy-mm"))*(超勤_月[状況]="確認待ち"),TRIM(超勤_月[所属部課]&" "&超勤_月[氏名]),""))&"）")&IF(COUNTIF(超勤_月[状況],"確認待ち")>COUNTIFS(超勤_月[年月],TEXT($B$2,"yyyy-mm"),超勤_月[状況],"確認待ち"),"　／　ほかの月の確認待ち "&(COUNTIF(超勤_月[状況],"確認待ち")-COUNTIFS(超勤_月[年月],TEXT($B$2,"yyyy-mm"),超勤_月[状況],"確認待ち"))&" 人・月","")'
    $ws.Range('H2').Font.Bold = $true
    $ws.Range('H2').Font.Color = 0x0000C0
    $ws.Range('H2').VerticalAlignment = -4108
    foreach ($k in $settings.Keys) { if ($settings[$k]) { $cfg.Range($k).Value2 = '' } }
    if ($TeleworkFolder) { $tw.Range('B2').Value2 = '' }
    $wb.Save()
    $wb.Close($false)
} finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
'queries added'
