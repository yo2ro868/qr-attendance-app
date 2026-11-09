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
# シートに出す表：シート・置く場所（見出しの行）・行数を書くセル・時間:分／日付の書式を付ける列
$loaded = [ordered]@{
    '年休_自動' = @{ Sheet = '休暇申請シートの読み込み'; At = 'A8'; Count = 'A6'; Time = @(); Date = @() }
    '月別_自動' = @{ Sheet = '休暇申請シートの読み込み'; At = 'E8'; Count = 'E6'; Time = (1..12 | ForEach-Object { "${_}月" }); Date = @() }
    '日別_自動' = @{ Sheet = '休暇申請シートの読み込み'; At = 'S8'; Count = 'S6'; Time = @(); Date = @() }
    '申請一覧'  = @{ Sheet = '申請一覧'; At = 'A3'; Count = ''; Time = @('時間'); Date = @('申請日', '開始日', '終了日', '承認日') }
    # テレワーク：一覧は注意の確認用、日別は出勤簿の出勤欄・備考用（一覧を参照するので一覧の後に作る）
    'テレワーク一覧'  = @{ Sheet = 'テレワーク一覧'; At = 'A5'; Count = ''; Time = @('開始時刻', '終了時刻', '時間'); Date = @('日付') }
    'テレワーク_日別' = @{ Sheet = '休暇申請シートの読み込み'; At = 'AD8'; Count = 'AD6'; Time = @(); Date = @() }
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
        $qt.AdjustColumnWidth = ($name -eq '申請一覧')
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
                $ws.Range($ws.Cells.Item($first, $c), $ws.Cells.Item($first + 19999, $c)).NumberFormat = $pair[1]
            }
        }
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
