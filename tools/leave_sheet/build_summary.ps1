# 休暇集計.xlsx を作る：make_summary.py の土台に、Excel で Power Query（queries\*.pq）を入れて保存する。
# 使い方：powershell -ExecutionPolicy Bypass -File build_summary.ps1 -Out 出力先.xlsx [-Folder 試しに読むフォルダ]
#   -Folder を付けると、そのフォルダを読んで更新した状態で保存する（確認用）。付けなければ空の様式。
param([Parameter(Mandatory = $true)][string]$Out, [string]$Folder = '')

$ErrorActionPreference = 'Stop'
$Out = [IO.Path]::GetFullPath($Out)
$raw = Join-Path $env:TEMP 'leave_summary_raw.xlsx'
python (Join-Path $PSScriptRoot 'make_summary.py') $raw | Out-Null
if ($LASTEXITCODE -ne 0) { throw 'make_summary.py が失敗しました' }

# 接続のみのクエリ（ほかのクエリから使う）と、シートに出すクエリ
$helpers = @('申請用紙', '申請データ')
$loaded = [ordered]@{
    '申請一覧' = @{ Time = @('時間'); Date = @('申請日', '開始日', '終了日', '承認日') }
}

$xl = New-Object -ComObject Excel.Application
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($raw)
    # 設定シートの値（フォルダ）と申請用紙を組み合わせるので、プライバシー レベルの確認を省く
    $wb.Queries.FastCombine = $true
    foreach ($name in $helpers + @($loaded.Keys)) {
        $m = [IO.File]::ReadAllText((Join-Path $PSScriptRoot "queries\$name.pq"), [Text.Encoding]::UTF8)
        [void]$wb.Queries.Add($name, $m)
    }
    if ($Folder) { $wb.Worksheets.Item('設定').Range('B3').Value2 = $Folder }
    foreach ($name in $loaded.Keys) {
        $ws = $wb.Worksheets.Item($name)
        $conn = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=`$Workbook`$;Location=`"$name`";Extended Properties=`"`""
        $lo = $ws.ListObjects.Add(0, $conn, $true, 1, $ws.Range('A3'))
        $lo.Name = $name
        $lo.TableStyle = 'TableStyleLight9'
        $qt = $lo.QueryTable
        $qt.CommandType = 2
        $qt.CommandText = "SELECT * FROM [$name]"
        $qt.BackgroundQuery = $false
        $qt.PreserveFormatting = $true
        $qt.AdjustColumnWidth = $true
        $qt.Refresh($false) | Out-Null
        # 書式は見出しより下の列全体に付ける（空の表でも、更新で増えた行に書式が付くように）
        foreach ($pair in @(@($loaded[$name].Time, '[h]:mm'), @($loaded[$name].Date, 'yyyy/m/d'))) {
            foreach ($col in $pair[0]) {
                $c = $lo.ListColumns.Item($col).Range.Column
                $ws.Range($ws.Cells.Item(4, $c), $ws.Cells.Item(10000, $c)).NumberFormat = $pair[1]
            }
        }
        $ws.Activate()
        $xl.ActiveWindow.SplitRow = 3
        $xl.ActiveWindow.FreezePanes = $true
    }
    if ($Folder) { $wb.Worksheets.Item('設定').Range('B3').Value2 = '' }
    $wb.Worksheets.Item('設定').Activate()
    if (Test-Path $Out) { Remove-Item $Out }
    $wb.SaveAs($Out, 51)
    $wb.Close($false)
} finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Remove-Item $raw -ErrorAction SilentlyContinue
}
"作成しました: $Out"
