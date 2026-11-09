# 出勤・休暇管理データ に［申請を取り込む］ボタンを付け、マクロ有効ブック（.xlsm）として保存する。
# マクロ（VBA）はスクリプトから入れられない（Excel のセキュリティ設定のため）ので、macro\出勤・休暇管理データのマクロ.bas を
# 出力先と同じフォルダに置き、使う人が一度だけ［Alt+F11］→［ファイル］→［ファイルのインポート］で取り込む。
# 使い方：add_buttons.ps1 -Path 元.xlsx -Out 出力先.xlsm
param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$Out)

$ErrorActionPreference = 'Stop'
$Out = [IO.Path]::GetFullPath($Out)
$bas = Join-Path (Split-Path $Out) '出勤・休暇管理データのマクロ.bas'
$code = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'macro\出勤・休暇管理データのマクロ.bas'), [Text.Encoding]::UTF8)
[IO.File]::WriteAllText($bas, ($code -replace "`r?`n", "`r`n"), [Text.Encoding]::GetEncoding(932))

$xl = New-Object -ComObject Excel.Application
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($Path)
    # フォーム コントロールのボタン（0 = xlButtonControl）。シートの上のほう、表にかからない行に置く
    foreach ($place in @(@('休暇申請シートの読み込み', 'D1', 1), @('申請一覧', 'A2', 2))) {
        $ws = $wb.Worksheets.Item($place[0])
        $ws.Rows.Item($place[2]).RowHeight = 30
        $c = $ws.Range($place[1])
        $btn = $ws.Shapes.AddFormControl(0, $c.Left + 2, $c.Top + 2, 150, 26)
        $btn.OLEFormat.Object.Caption = '申請を取り込む'
        $btn.OnAction = 'ImportRequests'
        $btn.Placement = 3      # セルに合わせて動かさない
    }
    $wb.Worksheets.Item(1).Activate()
    if (Test-Path $Out) { Remove-Item $Out }
    $wb.SaveAs($Out, 52)   # 52 = マクロ有効ブック（.xlsm）
    $wb.Close($false)
} finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
'buttons added'
