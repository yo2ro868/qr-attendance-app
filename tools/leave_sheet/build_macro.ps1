# 申請用紙の作成.xlsm（設定シート・結果シート・ボタン）と、そこに取り込むマクロ 申請用紙の作成.bas を作る。
# 使い方：powershell -ExecutionPolicy Bypass -File build_macro.ps1 -Out 出力先.xlsm
# マクロ（VBA）はスクリプトから入れられない（Excel のセキュリティ設定のため）ので、.bas を同じフォルダに置き、
# 使う人が一度だけ［Alt+F11］→［ファイル］→［ファイルのインポート］で取り込む。
param([Parameter(Mandatory = $true)][string]$Out)

$ErrorActionPreference = 'Stop'
$Out = [IO.Path]::GetFullPath($Out)
# VBA の取り込みは Shift_JIS（日本語の Windows の文字コード）のファイルから行う
$bas = Join-Path (Split-Path $Out) '申請用紙の作成.bas'
$code = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'macro\申請用紙の作成.bas'), [Text.Encoding]::UTF8)
[IO.File]::WriteAllText($bas, ($code -replace "`r?`n", "`r`n"), [Text.Encoding]::GetEncoding(932))

$xl = New-Object -ComObject Excel.Application
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Add()
    while ($wb.Worksheets.Count -lt 2) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
    $cfg = $wb.Worksheets.Item(1); $cfg.Name = '設定'
    $res = $wb.Worksheets.Item(2); $res.Name = '結果'
    $wb.Styles.Item('Normal').Font.Name = 'ＭＳ Ｐゴシック'
    $wb.Styles.Item('Normal').Font.Size = 10

    $cfg.Range('A1').Value2 = '休暇申請用紙の作成'
    $cfg.Range('A1').Font.Size = 14; $cfg.Range('A1').Font.Bold = $true; $cfg.Range('A1').Font.Color = 0x5F3A1E
    $labels = @('照合用Excel のファイル', 'ひな形のファイル', '作成先のフォルダ', '年（西暦）', 'シート保護のパスワード')
    $hints = @(
        '照合用Excel.xlsx の場所（職員名簿の 出勤用ID・氏名 と、年次休暇の合計を読みます）',
        '休暇申請用紙（ひな形）.xlsx の場所',
        'OneDrive の「休暇申請」フォルダの場所（ここに 休暇申請_出勤用ID_氏名.xlsx を作ります）',
        '申請用紙の基本の「年」に入れる年',
        'ひな形のシート保護にパスワードをかけた場合だけ入れます（ふつうは空欄）')
    for ($i = 0; $i -lt $labels.Count; $i++) {
        $r = 3 + $i
        $cfg.Range("A$r").Value2 = $labels[$i]; $cfg.Range("A$r").Font.Bold = $true
        $cfg.Range("B$r").Interior.Color = 0xCCF2FF
        $cfg.Range("B$r").Borders.LineStyle = 1
        $cfg.Range("C$r").Value2 = $hints[$i]; $cfg.Range("C$r").Font.Size = 9; $cfg.Range("C$r").Font.Color = 0x555555
    }
    $cfg.Range('B6').Value2 = [double](Get-Date).Year
    $cfg.Range('B6').HorizontalAlignment = -4131
    $cfg.Columns.Item(1).ColumnWidth = 24
    $cfg.Columns.Item(2).ColumnWidth = 60
    $notes = @(
        '使い方',
        '0. はじめに一度だけ：［Alt+F11］で VBA の画面を開き、［ファイル］→［ファイルのインポート］で同じフォルダの 申請用紙の作成.bas を取り込み、VBA の画面を閉じて上書き保存します。',
        '1. B3〜B6 を入れます。場所は、エクスプローラーでファイル（フォルダ）を Shift キーを押しながら右クリック →［パスとしてコピー］して貼り付けます。',
        '2.［申請用紙を作成］を押します。照合用Excel の職員名簿の全員について、申請用紙を作成先のフォルダに作ります。',
        '   ・基本の欄に 出勤用ID・氏名・年 と、照合用Excel の年次休暇のその年の「合計」（持ち日数）を入れます。所属長は空欄です。',
        '   ・作成先に同じ出勤用ID の申請用紙があれば、作りません（上書きしません）。新しい職員が増えたときも、もう一度押せば増えた人の分だけ作ります。',
        '3. 結果シートで、作った申請用紙を確かめます。',
        '4. 作った申請用紙を開いて所属長を入れ（［校閲］→［シート保護の解除］→ 入力 →［シートの保護］）、OneDrive で本人と所属長に共有します。')
    for ($i = 0; $i -lt $notes.Count; $i++) {
        $c = $cfg.Range("A$(10 + $i)"); $c.Value2 = $notes[$i]
        if ($i -eq 0) { $c.Font.Bold = $true } else { $c.Font.Size = 9 }
    }
    # フォーム コントロールのボタン（0 = xlButtonControl）
    $btn = $cfg.Shapes.AddFormControl(0, $cfg.Range('B9').Left, $cfg.Range('B9').Top - 2, 150, 24)
    $btn.OLEFormat.Object.Caption = '申請用紙を作成'
    $btn.OnAction = 'CreateLeaveSheets'

    $heads = @('出勤用ID', '氏名', 'ファイル名', '結果')
    for ($i = 0; $i -lt $heads.Count; $i++) { $res.Cells.Item(1, $i + 1).Value2 = $heads[$i] }
    $res.Range('A1:D1').Font.Bold = $true
    $res.Range('A1:D1').Interior.Color = 0xF5EEE8
    $res.Columns.Item(1).ColumnWidth = 12; $res.Columns.Item(2).ColumnWidth = 16
    $res.Columns.Item(3).ColumnWidth = 34; $res.Columns.Item(4).ColumnWidth = 70

    $cfg.Activate()
    if (Test-Path $Out) { Remove-Item $Out }
    $wb.SaveAs($Out, 52)   # 52 = マクロ有効ブック（.xlsm）
    $wb.Close($false)
} finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
"作成しました: $Out"
