# 出勤・休暇管理データ に［申請を取り込む］（申請一覧シート）などのボタンを付け、マクロ有効ブック（.xlsm）として保存する。
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
    # 読み込みのボタンは目次に集める（各シートには置かない。2026-11-20）
    # 目次：E 列の「[ボタン]マクロ名|表示」をボタンに、F 列の「[式]=…」を式にする（make_matching.py が書いておく）
    $ws = $wb.Worksheets.Item('目次')
    foreach ($c in $ws.Range('E1:F80').Cells) {
        $v = [string]$c.Value2
        if ($v.StartsWith('[ボタン]')) {
            $parts = $v.Substring(5).Split('|')
            $c.Value2 = [string]''
            $btn = $ws.Shapes.AddFormControl(0, $c.Left + 2, $c.Top + 2, $c.Width - 4, $c.Height - 4)
            $btn.OLEFormat.Object.Caption = $parts[1]
            $btn.OnAction = $parts[0]
            $btn.Placement = 3
        } elseif ($v.StartsWith('[式]')) {
            $c.Formula = $v.Substring(3)
        }
    }
    # ［超過勤務を転記］（超過勤務シートの B2 の月の、所属長の確認済みの超過勤務等命令簿を写す）
    $ws = $wb.Worksheets.Item('超過勤務')
    $ws.Rows.Item(2).RowHeight = 30
    $c = $ws.Range('C2')
    $btn = $ws.Shapes.AddFormControl(0, $c.Left + 6, $c.Top + 2, 130, 26)
    $btn.OLEFormat.Object.Caption = '超過勤務を転記'
    $btn.OnAction = 'TransferOvertime'
    $btn.Placement = 2      # セルに合わせて動く（列の幅が変わっても B2 に重ならない）
    # ［命令簿をPDFで保存］（命令簿の保存シートの2行目。保存先のフォルダは目次で入れる）
    $ws = $wb.Worksheets.Item('命令簿の保存')
    $c = $ws.Range('A2')
    $btn = $ws.Shapes.AddFormControl(0, $c.Left + 6, $c.Top + 2, 160, 26)
    $btn.OLEFormat.Object.Caption = '命令簿をPDFで保存'
    $btn.OnAction = 'SaveOvertimePdfs'
    $btn.Placement = 3
    # ［翌年分の行を追加］（休暇等シートの使い方の文の上。AQ1）
    $ws = $wb.Worksheets.Item('休暇等')
    $c = $ws.Range('AQ1')
    $btn = $ws.Shapes.AddFormControl(0, $c.Left + 2, $c.Top + 12, 170, 26)
    $btn.OLEFormat.Object.Caption = '翌年分の行を追加'
    $btn.OnAction = 'AddLeaveRows'
    $btn.Placement = 2
    # ［利用予定を書き出す］（利用予定シートの表を、今日の日付のシートに値で書き出す）
    $ws = $wb.Worksheets.Item('利用予定')
    $c = $ws.Range('AM7')      # 説明（AM3:AN6）の下
    $btn = $ws.Shapes.AddFormControl(0, $c.Left + 2, $c.Top + 2, 170, 26)
    $btn.OLEFormat.Object.Caption = '利用予定を書き出す'
    $btn.OnAction = 'MakeScheduleSheet'
    $btn.Placement = 3
    # ［申告・割振り簿を見る］（フレックスシートの2行目。表は3行目から）
    $ws = $wb.Worksheets.Item('フレックス')
    $ws.Rows.Item(2).RowHeight = 30
    $btn = $ws.Shapes.AddFormControl(0, $ws.Range('A2').Left + 2, $ws.Range('A2').Top + 2, 170, 26)
    $btn.OLEFormat.Object.Caption = '申告・割振り簿を見る'
    $btn.OnAction = 'ShowFlexPeriod'
    $btn.Placement = 3
    # ［実施報告書を見る］［実施報告書をPDFで保存］（テレワーク報告シートの2行目。表は3行目から）
    $ws = $wb.Worksheets.Item('テレワーク報告')
    $ws.Rows.Item(2).RowHeight = 30
    $x = $ws.Range('A2').Left + 2
    foreach ($b in @(@('実施報告書を見る', 'ShowTwReport', 150), @('実施報告書をPDFで保存', 'SaveTwReports', 180))) {
        $btn = $ws.Shapes.AddFormControl(0, $x, $ws.Range('A2').Top + 2, $b[2], 26)
        $btn.OLEFormat.Object.Caption = $b[0]
        $btn.OnAction = $b[1]
        $btn.Placement = 3
        $x += $b[2] + 8
    }
    # ［目次に戻る］：目次以外の見えるシートに置く（図形にリンクを付けるので、マクロを使わずに動く。印刷はしない。2026-11-20）
    # 置き場所は、そのシートの上の空いている所。'after:行' は、その行のボタンの右に並べる
    $back = [ordered]@{
        '出力データ貼付' = 'G1'; '照合結果' = 'J1'; '月次一覧' = 'A2'; '出勤簿' = 'AE1'; '一括印刷' = 'AE1';
        '申請一覧' = 'A2'; '休暇等' = 'after:1'; 'テレワーク一覧' = 'A3'; 'テレワーク報告' = 'after:2';
        '超過勤務' = 'after:2'; '命令簿の保存' = 'after:2'; 'フレックス' = 'after:2'; '利用予定' = 'A2';
        '職員名簿' = 'Q1'; '祝日' = 'D1'; '使い方' = 'B2'
    }
    foreach ($sn in $back.Keys) {
        $ws = $wb.Worksheets.Item($sn)
        $at = $back[$sn]
        if ($at.StartsWith('after:')) {
            $row = [int]$at.Substring(6)
            $x = 0; foreach ($sh in $ws.Shapes) { if ($sh.TopLeftCell.Row -eq $row) { $x = [math]::Max($x, $sh.Left + $sh.Width) } }
            $top = $ws.Rows.Item($row).Top
            $x += 8
        } else {
            $c = $ws.Range($at); $row = $c.Row
            $x = $c.Left + 4; $top = $c.Top
        }
        if ($ws.Rows.Item($row).RowHeight -lt 22) { $ws.Rows.Item($row).RowHeight = 22 }
        $h = [math]::Min(20, $ws.Rows.Item($row).RowHeight - 4)
        $y = $top + ($ws.Rows.Item($row).RowHeight - $h) / 2
        $shp = $ws.Shapes.AddShape(5, $x, $y, 78, $h)      # 5 = 角の丸い四角形
        $shp.Name = '目次に戻る'
        $shp.Fill.ForeColor.RGB = 0xF5EEE8                 # 薄い青（BGR）
        $shp.Line.ForeColor.RGB = 0x5F3A1E
        $shp.Line.Weight = 0.75
        $tf = $shp.TextFrame2
        $tf.MarginLeft = 2; $tf.MarginRight = 2; $tf.MarginTop = 0; $tf.MarginBottom = 0
        $tf.VerticalAnchor = 3                             # 上下の中央
        $tf.TextRange.Text = '目次に戻る'
        $tf.TextRange.Font.Size = 9
        $tf.TextRange.Font.Bold = -1
        $tf.TextRange.Font.Name = 'BIZ UDPゴシック'
        $tf.TextRange.Font.NameFarEast = 'BIZ UDPゴシック'
        $tf.TextRange.Font.Fill.ForeColor.RGB = 0x5F3A1E
        $tf.TextRange.ParagraphFormat.Alignment = 2        # 中央
        $shp.Placement = 3                                 # セルに合わせて動かさない
        $shp.DrawingObject.PrintObject = $false
        [void]$ws.Hyperlinks.Add($shp, '', "'目次'!A1", '目次のシートへ')
    }
    # ふだんは見ないシートは非表示（中身はそのまま。読み込み・照合は裏で動く。シート見出しを右クリック→［再表示］で見られる。2026-11-20）
    #   フレックスの日別：照合の注意はフレックスシートに出る。QR出勤・申請・報告シート・フレックスの読み込み：読んだそのままの表（結果は各一覧と目次に出る）
#   超過勤務の読み込み・日別：確認待ちは超過勤務シートの H2、日ごとの注意は超過勤務シートの「注意」の列に出る
    foreach ($n in @('フレックスの日別', 'QR出勤の読み込み', '申請・報告シートの読み込み', 'フレックスの読み込み', '超過勤務の読み込み', '超過勤務の日別')) { $wb.Worksheets.Item($n).Visible = 0 }
    # シートの名前を、まとまりと作業の順の名前に付け直す（2026-11-20。例 超過勤務の読み込み → 超勤確認状況）。
    # 式・定義した名前・入力規則・条件付き書式は Excel が自動で直す。目次のリンク（SubAddress）と表示、説明の文の「…」「…シート」はここで直す（長い名前から）
    $rename = @(
        @('申請・報告シートの読み込み', '休暇読み込み'),
        @('フレックスの読み込み', 'フレックス読み込み'),
        @('QR出勤の読み込み', '出勤QR読み込み'),
        @('超過勤務の読み込み', '超勤確認状況'),
        @('フレックスの日別', 'フレックス日別照合'),
        @('出力データ貼付', '出勤①データ貼付'),
        @('テレワーク一覧', 'テレワーク①予定一覧'),
        @('テレワーク報告', 'テレワーク②実施報告'),
        @('超過勤務の日別', '超勤日別照合'),
        @('命令簿の保存', '超勤②PDF保存'),
        @('フレックス', 'フレックス①承認状況'),
        @('照合結果', '出勤②照合結果'),
        @('月次一覧', '出勤③月次一覧'),
        @('一括印刷', '出勤⑤一括印刷'),
        @('申請一覧', '休暇①申請一覧'),
        @('超過勤務', '超勤①転記・月計'),
        @('利用予定', 'フレックス②利用予定'),
        @('職員名簿', '台帳①職員名簿'),
        @('出勤簿', '出勤④出勤簿'),
        @('休暇等', '休暇②個人別一覧'),
        @('祝日', '台帳②祝日'))
    foreach ($p in $rename) { $wb.Worksheets.Item($p[0]).Name = $p[1] }
    foreach ($ws in $wb.Worksheets) {
        foreach ($h in $ws.Hyperlinks) {
            if ($h.Type -ne 0) { continue }     # 0 = セルのリンク（図形の［目次に戻る］は飛ばす）
            $sa = [string]$h.SubAddress
            foreach ($p in $rename) { if ($sa.StartsWith("'" + $p[0] + "'!")) { $h.SubAddress = "'" + $p[1] + "'!" + $sa.Substring($p[0].Length + 3); break } }
            $tx = [string]$h.Range.Value2
            foreach ($p in $rename) { if ($tx -eq $p[0]) { $h.Range.Value2 = $p[1]; $h.TextToDisplay = $p[1]; break } }
        }
        # セルを1つずつ聞くと何時間もかかるので、使っている範囲の式・文字をまとめて読み、変わるセルだけ書く
        $used = $ws.UsedRange
        $all = $used.Formula
        if ($all -isnot [array]) { $one = New-Object 'object[,]' 1, 1; $one[0, 0] = $all; $all = $one }
        $r0 = $used.Row; $c0 = $used.Column
        $lb1 = $all.GetLowerBound(0); $lb2 = $all.GetLowerBound(1)
        for ($i = $lb1; $i -le $all.GetUpperBound(0); $i++) {
            for ($j = $lb2; $j -le $all.GetUpperBound(1); $j++) {
                $v = $all[$i, $j]
                if ($v -is [string] -and -not $v.StartsWith('=') -and ($v.Contains('「') -or $v.Contains('シート'))) {
                    $nv = $v
                    foreach ($p in $rename) { $nv = $nv.Replace('「' + $p[0] + '」', '「' + $p[1] + '」').Replace($p[0] + 'シート', $p[1] + 'シート').Replace($p[0] + ' シート', $p[1] + ' シート') }
                    if ($nv -ne $v) { $ws.Cells.Item($r0 + $i - $lb1, $c0 + $j - $lb2).Value2 = $nv }
                }
            }
        }
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
