# 出勤・休暇管理データ に Power Query を入れる：休暇申請シート（tools/leave_sheet）のフォルダと、Forms の休暇申請・休暇承認の回答の Excel を読み、
# 「休暇申請シートの読み込み」シートに 年休_自動・月別_自動・日別_自動 の表を、「申請一覧」シートに申請一覧の表を作る。開いたときに読み直す設定にする。
# 使い方：add_queries.ps1 -Path 出勤・休暇管理データ.xlsx [-Folder 休暇申請シートのフォルダ] [-FormsRequest 申請の回答.xlsx] [-FormsApproval 承認の回答.xlsx]
#   -Folder などを付けると、それを読んだ状態で保存する（確認用）。付けなければ設定欄は空欄のまま。
#   -TeleworkFolder はテレワーク予定表（tools/telework）のフォルダ。テレワーク一覧シートと、出勤簿の「テレ」に使う。
param([Parameter(Mandatory = $true)][string]$Path, [string]$Folder = '', [string]$FormsRequest = '', [string]$FormsApproval = '', [string]$TeleworkFolder = '')

$ErrorActionPreference = 'Stop'
$Path = [IO.Path]::GetFullPath($Path)
$sheetQueries = Join-Path $PSScriptRoot '..\leave_sheet\queries'   # 休暇申請シートを読むクエリは休暇集計と同じもの
$teleworkQueries = Join-Path $PSScriptRoot '..\telework\queries'    # テレワーク予定表を読むクエリ
$queries = Join-Path $PSScriptRoot 'queries'
# 接続のみのクエリ（ほかのクエリから使う）。名前 → ファイル
$helpers = [ordered]@{
    '休暇申請シート'       = Join-Path $sheetQueries '休暇申請シート.pq'
    '祝日の日付'     = Join-Path $queries '祝日の日付.pq'
    '職員名簿_メール' = Join-Path $queries '職員名簿_メール.pq'
    'Forms回答_申請' = Join-Path $queries 'Forms回答_申請.pq'
    'Forms回答_承認' = Join-Path $queries 'Forms回答_承認.pq'
    '申請データ_用紙' = Join-Path $sheetQueries '申請データ.pq'
    'Forms承認'      = Join-Path $queries 'Forms承認.pq'
    'Forms申請'      = Join-Path $queries 'Forms申請.pq'
    '申請データ'     = Join-Path $queries '申請データ.pq'
    'テレワーク予定表' = Join-Path $teleworkQueries 'テレワーク予定表.pq'
}
# 超過勤務の支給割合ごとの時間の列（超勤_日別.pq・超勤_月.pq と同じ）と、手当対象時間の列
$otBuckets = @('125', '150', '150（60H前）', '150（+25）', '150（+15）', '135', '160', '175', '175（+25）', '175（+15）',
               '休日勤務135', '超勤代休時間の勤務', '超勤代休（+25）', '超勤代休（+15）')
$otAllowance = @('125', '150', '135', '160', '175', '休日勤務135') | ForEach-Object { "手当対象 $_" }
# シートに出す表：シート・置く場所（見出しの行）・行数を書くセル・時間:分／日付の書式を付ける列
$loaded = [ordered]@{
    '年休_自動' = @{ Sheet = '休暇申請シートの読み込み'; At = 'A8'; Count = 'A6'; Time = @(); Date = @() }
    '月別_自動' = @{ Sheet = '休暇申請シートの読み込み'; At = 'E8'; Count = 'E6'; Time = (1..12 | ForEach-Object { "${_}月" }); Date = @() }
    '日別_自動' = @{ Sheet = '休暇申請シートの読み込み'; At = 'S8'; Count = 'S6'; Time = @(); Date = @() }
    '申請一覧'  = @{ Sheet = '申請一覧'; At = 'A3'; Count = ''; Time = @('時間'); Date = @('申請日', '開始日', '終了日', '承認日') }
    # テレワーク：一覧は注意の確認用、日別は出勤簿の出勤欄・備考用（一覧を参照するので一覧の後に作る）
    'テレワーク一覧'  = @{ Sheet = 'テレワーク一覧'; At = 'A5'; Count = ''; Time = @('開始時刻', '終了時刻', '時間'); Date = @('日付') }
    'テレワーク_日別' = @{ Sheet = '休暇申請シートの読み込み'; At = 'AD8'; Count = 'AD6'; Time = @(); Date = @() }
    # 超過勤務等命令簿（休暇申請シートの月のシート）：1人1か月の計と確認、日ごとの行
    '超勤_月'   = @{ Sheet = '超過勤務の読み込み'; At = 'A3'; Count = ''; Time = ($otBuckets + $otAllowance + @('60時間超（+25）', '60時間超（+15）')); Date = @('確認日'); Whole = $true }
    '超勤_日別' = @{ Sheet = '超過勤務の読み込み'; At = 'AG3'; Count = ''; Time = (@('登庁時刻', '退庁時刻', '休憩時間') + $otBuckets); Date = @('確認日', '日付'); Whole = $true }
}
$settings = @{ 'B2' = $Folder; 'B3' = $FormsRequest; 'B4' = $FormsApproval }

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
    $cfg = $wb.Worksheets.Item('休暇申請シートの読み込み')
    foreach ($k in $settings.Keys) { if ($settings[$k]) { $cfg.Range($k).Value2 = $settings[$k] } }
    $tw = $wb.Worksheets.Item('テレワーク一覧')
    if ($TeleworkFolder) { $tw.Range('B2').Value2 = $TeleworkFolder }
    foreach ($name in $loaded.Keys) {
        $spec = $loaded[$name]
        $ws = $wb.Worksheets.Item($spec.Sheet)
        $conn = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=`$Workbook`$;Location=`"$name`";Extended Properties=`"`""
        $lo = $ws.ListObjects.Add(0, $conn, $true, 1, $ws.Range($spec.At))
        $lo.Name = $name
        $lo.TableStyle = 'TableStyleLight9'
        $qt = $lo.QueryTable
        $qt.CommandType = 2
        $qt.CommandText = "SELECT * FROM [$name]"
        $qt.BackgroundQuery = $false
        $qt.PreserveFormatting = $true
        $qt.AdjustColumnWidth = ($name -eq '申請一覧' -or $spec.Whole)
        $qt.Refresh($false) | Out-Null
        # 出勤・休暇管理データ を開いたときに読み直す
        $lo.QueryTable.WorkbookConnection.OLEDBConnection.RefreshOnFileOpen = $true
        # 行数（出勤簿などがこの範囲だけを探す）。表の名前で書くので、読み直して行が増減しても合う
        if ($spec.Count) {
            $ws.Range($spec.Count).Formula = "=COUNTA($name[キー])"
            $ws.Range($spec.Count).NumberFormat = '"（"0"行）"'
        }
        # 書式は見出しより下の列全体に付ける（空の表でも、読み直して増えた行に書式が付くように）
        $first = $ws.Range($spec.At).Row + 1
        foreach ($pair in @(@($spec.Time, '[h]:mm'), @($spec.Date, 'yyyy/m/d'))) {
            foreach ($col in $pair[0]) {
                $c = $lo.ListColumns.Item($col).Range.Column
                if ($spec.Whole) { $ws.Columns.Item($c).NumberFormat = $pair[1] }   # 列全体（空のセルごとの書式を持たないので、ファイルが大きくならない）
                else { $ws.Range($ws.Cells.Item($first, $c), $ws.Cells.Item($first + 19999, $c)).NumberFormat = $pair[1] }
            }
        }
    }
    # ----- ［超過勤務を転記］で写す先の表（値。照合の列だけは式）。見出しは読み込んだ表と同じ＋転記日時 -----
    $au = "'休暇申請シートの読み込み'"
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
    foreach ($dst in @(@{ Sheet = '超過勤務'; At = 'A4'; From = '超勤_月'; Name = '転記_月計'; Checks = $false },
                       @{ Sheet = '超過勤務の日別'; At = 'A3'; From = '超勤_日別'; Name = '転記_日別'; Checks = $true })) {
        $ws = $wb.Worksheets.Item($dst.Sheet)
        $srcLo = $wb.Worksheets.Item('超過勤務の読み込み').ListObjects.Item($dst.From)
        $heads = @($srcLo.HeaderRowRange.Cells | ForEach-Object { [string]$_.Value2 }) + @('転記日時')
        if ($dst.Checks) { $heads += @($checks.Keys) }
        $top = $ws.Range($dst.At)
        for ($i = 0; $i -lt $heads.Count; $i++) { $ws.Cells.Item($top.Row, $top.Column + $i).Value2 = [string]$heads[$i] }
        $rng = $ws.Range($top, $ws.Cells.Item($top.Row + 1, $top.Column + $heads.Count - 1))
        $lo = $ws.ListObjects.Add(1, $rng, $null, 1)
        $lo.Name = $dst.Name
        $lo.TableStyle = 'TableStyleLight9'
        if ($dst.Checks) { foreach ($k in $checks.Keys) { $lo.ListColumns.Item($k).DataBodyRange.Formula = $checks[$k] } }
        # 書式：読み込んだ表と同じ列に時間・日付、転記日時。出勤用ID・年月は文字
        $first = $top.Row + 1
        $spec = $loaded[$dst.From]
        foreach ($pair in @(@($spec.Time, '[h]:mm'), @($spec.Date, 'yyyy/m/d'), @(@('転記日時'), 'yyyy/m/d h:mm'), @(@('出勤用ID', '年月'), '@'))) {
            foreach ($col in $pair[0]) {
                $c = $lo.ListColumns.Item($col).Range.Column
                $ws.Columns.Item($c).NumberFormat = $pair[1]
            }
        }
        if ($dst.Checks) { $lo.ListColumns.Item('注意').Range.ColumnWidth = 60 }
        foreach ($col in @($spec.Date)) { $lo.ListColumns.Item($col).Range.ColumnWidth = 11 }
        $lo.ListColumns.Item('転記日時').Range.ColumnWidth = 16
        $lo.ListColumns.Item('ファイル名').Range.ColumnWidth = 12
    }
    foreach ($k in $settings.Keys) { if ($settings[$k]) { $cfg.Range($k).Value2 = '' } }
    if ($TeleworkFolder) { $tw.Range('B2').Value2 = '' }
    $wb.Save()
    $wb.Close($false)
} finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
'queries added'
