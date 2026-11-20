# 申請・報告シートの作成.xlsm（設定シート・結果シート・ボタン）と、そこに取り込むマクロ 申請・報告シートの作成.bas を作る。
# 使い方：powershell -ExecutionPolicy Bypass -File build_macro.ps1 -Out 出力先.xlsm
# マクロ（VBA）はスクリプトから入れられない（Excel のセキュリティ設定のため）ので、.bas を同じフォルダに置き、
# 使う人が一度だけ［Alt+F11］→［ファイル］→［ファイルのインポート］で取り込む。
param([Parameter(Mandatory = $true)][string]$Out)

$ErrorActionPreference = 'Stop'
$Out = [IO.Path]::GetFullPath($Out)
# VBA の取り込みは Shift_JIS（日本語の Windows の文字コード）のファイルから行う
$bas = Join-Path (Split-Path $Out) '申請・報告シートの作成.bas'
$code = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'macro\申請・報告シートの作成.bas'), [Text.Encoding]::UTF8)
[IO.File]::WriteAllText($bas, ($code -replace "`r?`n", "`r`n"), [Text.Encoding]::GetEncoding(932))

$xl = New-Object -ComObject Excel.Application
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Add()
    while ($wb.Worksheets.Count -lt 2) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
    $cfg = $wb.Worksheets.Item(1); $cfg.Name = '設定'
    $res = $wb.Worksheets.Item(2); $res.Name = '結果'
    $wb.Styles.Item('Normal').Font.Name = 'BIZ UDPゴシック'
    $wb.Styles.Item('Normal').Font.Size = 10

    $cfg.Range('A1').Value2 = '申請・報告シートの作成'
    $cfg.Range('A1').Font.Size = 14; $cfg.Range('A1').Font.Bold = $true; $cfg.Range('A1').Font.Color = 0x5F3A1E
    $labels = @('出勤・休暇管理データ のファイル', 'ひな形のファイル', '作成先のフォルダ', '年（西暦）', '超過勤務の月')
    $hints = @(
        '出勤・休暇管理データ.xlsm の場所（職員名簿の 出勤用ID・氏名 と、休暇等シートの合計・本年使用の調整を読みます）',
        '申請・報告シート（ひな形）.xlsx の場所（［申請・報告シートを作成］のときだけ使います）',
        'OneDrive の「休暇申請」フォルダの場所（ここに 申請・報告シート　出勤用ID　氏名.xlsx を作ります）',
        '申請・報告シートに設定する年（出勤・休暇管理データ の休暇等シートのこの年の行を読みます）',
        '［超過勤務の月を追加］で作る超過勤務等命令簿の月（例 2026/12。その月のシート「超勤 R8.12」を全員の申請・報告シートに作ります）')
    for ($i = 0; $i -lt $labels.Count; $i++) {
        $r = 3 + $i
        $cfg.Range("A$r").Value2 = $labels[$i]; $cfg.Range("A$r").Font.Bold = $true
        $cfg.Range("B$r").Interior.Color = 0xCCF2FF
        $cfg.Range("B$r").Borders.LineStyle = 1
        $cfg.Range("C$r").Value2 = $hints[$i]; $cfg.Range("C$r").Font.Size = 9; $cfg.Range("C$r").Font.Color = 0x555555
    }
    $cfg.Range('B6').Value2 = [double](Get-Date).Year
    $cfg.Range('B6').HorizontalAlignment = -4131
    $next = (Get-Date -Day 1).AddMonths(1)
    $cfg.Range('B7').Value2 = [double]$next.Date.ToOADate()
    $cfg.Range('B7').NumberFormat = 'yyyy/m'
    $cfg.Range('B7').HorizontalAlignment = -4131
    # A 列は 306 ピクセル（＝ 229.5 ポイント。画面の倍率 100%）。幅の数は標準のフォントの文字数なので、ポイントで合わせる
    $colA = $cfg.Columns.Item(1); $colA.ColumnWidth = 40
    for ($k = 0; $k -lt 4 -and [math]::Abs($colA.Width - 229.5) -gt 0.4; $k++) { $colA.ColumnWidth = [math]::Round($colA.ColumnWidth * 229.5 / $colA.Width, 2) }
    $cfg.Columns.Item(2).ColumnWidth = 60
    $notes = @(
        '使い方',
        '0. はじめに一度だけ：［Alt+F11］で VBA の画面を開き、［ファイル］→［ファイルのインポート］で同じフォルダの 申請・報告シートの作成.bas を取り込み、VBA の画面を閉じて上書き保存します。',
        '1. B3〜B6 を入れます。場所は、エクスプローラーでファイル（フォルダ）を Shift キーを押しながら右クリック →［パスとしてコピー］して貼り付けます。',
        '2.［申請・報告シートを作成］：出勤・休暇管理データ の職員名簿の全員について、まだない人の申請・報告シートを作成先のフォルダに作ります（あれば作らない・上書きしない）。',
        '   新しい職員が増えたときも、もう一度押せば増えた人の分だけ作ります。作ったら OneDrive で本人と所属長に共有します。',
        '3.［申請・報告シートを更新］：作成先にある全員の申請・報告シートに、出勤・休暇管理データ の休暇等シートの B6 の年の 合計・本年使用の調整 を書き込みます。',
        '   年が変わったとき、休暇等シートの本年分・繰越分・調整を変えたときに押します。申請・報告シートの「年次休暇の残り」に反映されます。',
        '   職員が申請・報告シートを開いて書いている最中に押すと、書き込みがぶつかることがあります。朝や夕方など、使われていない時間に押してください。',
        '4.［様式を移行］：ひな形を新しくしたあと、様式の古い申請・報告シートを、書いてある申請の行と超過勤務等命令簿の月のシートを残して新しい様式にします（共有はそのまま）。元のファイルは作成先の中の 旧様式_日時 のフォルダに残します。',
        '   B4 を新しいひな形にして、職員や所属長が開いていない時間に押します。',
        '5.［超過勤務の月を追加］：毎月の終わりごろに、B7 を翌月にして押します。全員の申請・報告シートに超過勤務等命令簿のその月のシート（超勤 R8.12 など）を作り、前月の最終日の退庁時刻（25:00 など）をつなぎます。あるシートは作り直しません。',
        '6. 結果シートで、職員ごとの結果を確かめます。',
        '7.［申告・割振り簿を見る］：B9 に出勤用ID を入れて押すと、その人の申請・報告シートを読み取り専用で開き、管理者だけが見るフレックスの「申告・割振り簿」（総務課の様式）を表示します。印刷したら保存しないで閉じます。',
        '8.［古い超過勤務の月を外す］：申請・報告シートは命令簿の月が増えるほど重くなります（1か月で約90KB）。年度の初めなどに、B10 を前の年度の4月にして押すと、それより前の月のシートを外します。',
        '   外すのは、所属長の確認済みで、出勤・休暇管理データ に転記済み、［命令簿をPDFで保存］で PDF にした（そのあと直されていない）月だけです。外したシートは戻せません（PDF が残ります）。',
        '※ 申請・報告シートでは、年・合計・調整は非表示の「管理」シートに入り、本人・所属長には見えません（見えるのは氏名と、計算した残り日数などだけです）。')
    # ［申告・割振り簿を見る］で開く人の出勤用ID（B9）
    $cfg.Range('A9').Value2 = '申告・割振り簿を見る人'; $cfg.Range('A9').Font.Bold = $true
    $cfg.Range('B9').Interior.Color = 0xCCF2FF; $cfg.Range('B9').Borders.LineStyle = 1; $cfg.Range('B9').NumberFormat = '@'
    $cfg.Range('B9').HorizontalAlignment = -4131
    $cfg.Range('C9').Value2 = '出勤用ID（フレックスの申告・割振り簿を、読み取り専用で開いて表示します）'
    $cfg.Range('C9').Font.Size = 9; $cfg.Range('C9').Font.Color = 0x555555
    # 古い超過勤務の月を外すときの境目（B10。この月より前を外す）
    $cfg.Range('A10').Value2 = '超過勤務の月を外す境目'; $cfg.Range('A10').Font.Bold = $true
    $cfg.Range('B10').Interior.Color = 0xCCF2FF; $cfg.Range('B10').Borders.LineStyle = 1; $cfg.Range('B10').NumberFormat = 'yyyy/m'
    $cfg.Range('B10').HorizontalAlignment = -4131
    $cfg.Range('C10').Value2 = '［古い超過勤務の月を外す］で、この月より前の月のシートを外します（例 2025/4。目安は前の年度の4月）'
    $cfg.Range('C10').Font.Size = 9; $cfg.Range('C10').Font.Color = 0x555555
    # ボタン：11行目の見出しの下に1行に1つ（入力欄にかからないように）。右の B 列にすることを書く。
    # 行の高さを先に決めてから置く（後で高さを変えると、ボタンが下の行にはみ出すため）
    $cfg.Range('A11').Value2 = '操作（ボタン）'; $cfg.Range('A11').Font.Bold = $true
    $buttons = @(
        @('申請・報告シートを作成', 'CreateLeaveSheets', '職員名簿の全員について、まだない人の申請・報告シートを作ります（B3〜B6）'),
        @('申請・報告シートを更新', 'UpdateLeaveSheets', '全員の申請・報告シートに、休暇等シートの B6 の年の 合計・本年使用の調整 などを書き込みます'),
        @('様式を移行', 'MigrateLeaveSheets', '様式の古い申請・報告シートを、書いた内容を残して B4 のひな形の様式にします'),
        @('超過勤務の月を追加', 'AddOvertimeMonths', '全員の申請・報告シートに、B7 の月の超過勤務等命令簿のシートを作ります'),
        @('申告・割振り簿を見る', 'ShowFlexBook', 'B9 の人の申請・報告シートを読み取り専用で開き、フレックスの申告・割振り簿を表示します'),
        @('古い超過勤務の月を外す', 'RemoveOldOvertimeMonths', '全員の申請・報告シートから、B10 より前の月のうち、確認済み・転記済み・PDF 保存済みの月のシートを外します（軽くするため）'))
    for ($i = 0; $i -lt $buttons.Count; $i++) {
        $r = 12 + $i
        $cfg.Rows.Item($r).RowHeight = 30
        $cell = $cfg.Range("A$r")
        $btn = $cfg.Shapes.AddFormControl(0, $cell.Left + 2, $cell.Top + 2, $cell.Width - 4, 26)
        $btn.OLEFormat.Object.Caption = $buttons[$i][0]
        $btn.OnAction = $buttons[$i][1]
        $btn.Placement = 2      # セルに合わせて動く
        $cfg.Range("B$r").Value2 = $buttons[$i][2]
        $cfg.Range("B$r").Font.Size = 9; $cfg.Range("B$r").VerticalAlignment = -4108
    }
    # 使い方（ボタンの下）
    for ($i = 0; $i -lt $notes.Count; $i++) {
        $c = $cfg.Range("A$(19 + $i)"); $c.Value2 = $notes[$i]
        if ($i -eq 0) { $c.Font.Bold = $true } else { $c.Font.Size = 9 }
    }
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
