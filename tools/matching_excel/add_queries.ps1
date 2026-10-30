# 照合用Excel に Power Query を入れる：休暇申請用紙（tools/leave_sheet）のフォルダを読み、
# 「申請用紙の読み込み」シートに 年休_自動・月別_自動・日別_自動 の表を作る。開いたときに読み直す設定にする。
# 使い方：add_queries.ps1 -Path 照合用Excel.xlsx [-Folder 試しに読むフォルダ]
#   -Folder を付けると、そのフォルダを読んだ状態で保存する（確認用）。付けなければフォルダは空欄のまま。
param([Parameter(Mandatory = $true)][string]$Path, [string]$Folder = '')

$ErrorActionPreference = 'Stop'
$Path = [IO.Path]::GetFullPath($Path)
$sheetQueries = Join-Path $PSScriptRoot '..\leave_sheet\queries'   # 申請用紙を読むクエリは休暇集計と同じもの
$helpers = [ordered]@{
    '申請用紙'   = Join-Path $sheetQueries '申請用紙.pq'
    '申請データ' = Join-Path $sheetQueries '申請データ.pq'
}
# 表の名前 → 置く場所（キーの列の5行目。見出しが5行目、データが6行目から）と、時間:分 の書式を付ける列
$loaded = [ordered]@{
    '年休_自動' = @{ At = 'A5'; Time = @() }
    '月別_自動' = @{ At = 'D5'; Time = (1..12 | ForEach-Object { "${_}月" }) }
    '日別_自動' = @{ At = 'R5'; Time = @() }
}

$xl = New-Object -ComObject Excel.Application
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($Path)
    # 設定の値（フォルダ・祝日）と申請用紙を組み合わせるので、プライバシー レベルの確認を省く
    $wb.Queries.FastCombine = $true
    foreach ($name in $helpers.Keys) {
        [void]$wb.Queries.Add($name, [IO.File]::ReadAllText($helpers[$name], [Text.Encoding]::UTF8))
    }
    foreach ($name in $loaded.Keys) {
        [void]$wb.Queries.Add($name, [IO.File]::ReadAllText((Join-Path $PSScriptRoot "queries\$name.pq"), [Text.Encoding]::UTF8))
    }
    $ws = $wb.Worksheets.Item('申請用紙の読み込み')
    if ($Folder) { $ws.Range('B2').Value2 = $Folder }
    foreach ($name in $loaded.Keys) {
        $conn = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=`$Workbook`$;Location=`"$name`";Extended Properties=`"`""
        $lo = $ws.ListObjects.Add(0, $conn, $true, 1, $ws.Range($loaded[$name].At))
        $lo.Name = $name
        $lo.TableStyle = 'TableStyleLight9'
        $qt = $lo.QueryTable
        $qt.CommandType = 2
        $qt.CommandText = "SELECT * FROM [$name]"
        $qt.BackgroundQuery = $false
        $qt.PreserveFormatting = $true
        $qt.AdjustColumnWidth = $false
        $qt.Refresh($false) | Out-Null
        # 照合用Excel を開いたときに読み直す
        $lo.QueryTable.WorkbookConnection.OLEDBConnection.RefreshOnFileOpen = $true
        # 3行目：表の行数（出勤簿などがこの範囲だけを探す）。表の名前で書くので、読み直して行が増減しても合う
        $ws.Range($loaded[$name].At).Offset(-2, 0).Formula = "=COUNTA($name[キー])"
        $ws.Range($loaded[$name].At).Offset(-2, 0).NumberFormat = '"（"0"行）"'
        foreach ($col in $loaded[$name].Time) {
            $c = $lo.ListColumns.Item($col).Range.Column
            $ws.Range($ws.Cells.Item(6, $c), $ws.Cells.Item(20005, $c)).NumberFormat = '[h]:mm'
        }
    }
    if ($Folder) { $ws.Range('B2').Value2 = '' }
    $wb.Save()
    $wb.Close($false)
} finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
'queries added'
