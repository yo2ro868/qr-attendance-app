# 休暇申請シートのひな形に、様式どおりの完成版のシート「申告・割振り簿」を入れる。
# 総務課の様式の Excel（-Source。シートの保護を外したもの）の「フレックスタイム制」シートをそのまま写し、次のところだけ変える：
#   ・申告の欄（Z〜AS 列：コアタイム・勤務・休憩の時と分）：本人が「フレックス申告」シートに書いた値を表示する式
#   ・時間数（AT・AU 列）：フレックス申告シートで計算した時間数（空の欄でも計算が止まらないように）
#   ・割振りの欄（D〜Y 列）：様式の式のまま（申告の欄を写す）
#   ・1年間の始期（F4〜I4）：管理シートの始期。単位期間（F5〜I5）：今日を含む単位期間（書き換えると、その期間の合計と表示になる）
#   ・所属・氏名・所属長氏名・異動後の所属長氏名（AS3〜AS6）：フレックス申告・フレックス割振りシートから
#   ・備考（AV 列）：フレックス申告シートの備考。単位期間ごとの申告年月日（AW）はフレックス申告、割振年月日はフレックス割振りから
#   ・様式のブックの外を指す名前（月午前 など。使われていない）は消す。入力規則は消す
#   ・単位期間（F5〜I5）のほかは入力できないように保護する（パスワードなし）
# 使い方：add_flex_book.ps1 -Template 休暇申請シート（ひな形）.xlsx -Source 申告・割振り簿（原本）.xlsx
param([Parameter(Mandatory = $true)][string]$Template, [Parameter(Mandatory = $true)][string]$Source)

$ErrorActionPreference = 'Stop'
$Template = [IO.Path]::GetFullPath($Template)
$Source = [IO.Path]::GetFullPath($Source)
if (-not (Test-Path $Source)) { throw "申告・割振り簿の様式のファイルが見つかりません：$Source" }

$FIRST, $LAST, $PERIOD = 12, 403, 28
$SIN = "'フレックス申告'"
$WARI = "'フレックス割振り'"

$xl = New-Object -ComObject Excel.Application
$xl.DisplayAlerts = $false
$xl.Visible = $false
$src = $null
try {
    $src = $xl.Workbooks.Open($Source, 0, $true)
    $ws = $src.Worksheets.Item(1)
    if ($ws.ProtectContents) { throw '様式のシートが保護されています。保護を外した様式を使ってください（パスワードの確認で止まるため）' }
    # 列の幅（ポイント）。標準のフォントがひな形と違うので、写したあとで同じ幅に直す
    $widths = @{}
    for ($i = 1; $i -le 60; $i++) { $widths[$i] = [double]$ws.Columns.Item($i).Width }

    $wb = $xl.Workbooks.Open($Template)
    foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq '申告・割振り簿') { throw 'ひな形にもう 申告・割振り簿 シートがあります（make_form.py で作り直してから実行してください）' } }
    $ws.Copy([Type]::Missing, $wb.Worksheets.Item('フレックス割振り'))
    $t = $null
    foreach ($s in @($wb.Worksheets)) { if ($s.Name -like 'フレックスタイム制*') { $t = $s } }
    if (-not $t) { throw '写したシートが見つかりません' }
    $t.Name = '申告・割振り簿'
    $src.Close($false)
    $src = $null

    # 様式のブックを指す名前（使われていない）を消す
    foreach ($n in @($wb.Names)) {
        if ($n.RefersTo -match '\[' -or $n.RefersTo -match '#REF!') { $n.Delete() }
    }
    for ($i = 1; $i -le 60; $i++) {
        $col = $t.Columns.Item($i)
        if ($widths[$i] -gt 0 -and $col.Width -gt 0) {
            for ($k = 0; $k -lt 3 -and [math]::Abs($col.Width - $widths[$i]) -gt 0.5; $k++) {
                $col.ColumnWidth = [math]::Round($col.ColumnWidth * $widths[$i] / $col.Width, 2)
            }
        }
    }
    $t.Cells.Validation.Delete()
    $xl.Calculation = -4135    # 手動（式を入れるあいだ再計算しない）

    # 1年間の始期・単位期間
    $t.Range('F4').Formula = '=YEAR(フレックス始期)'
    $t.Range('H4').Formula = '=MONTH(フレックス始期)'
    $t.Range('I4').Formula = '=DAY(フレックス始期)'
    $now = "(フレックス始期+$PERIOD*(MIN(14,MAX(1,INT((TODAY()-フレックス始期)/$PERIOD)+1))-1))"
    $t.Range('F5').Formula = "=YEAR($now)"
    $t.Range('H5').Formula = "=MONTH($now)"
    $t.Range('I5').Formula = "=DAY($now)"
    # 所属・氏名・所属長氏名・異動後の所属長氏名
    $t.Range('AS3').Formula = "=IF($SIN!`$D`$3="""","""",$SIN!`$D`$3)"
    $t.Range('AS4').Formula = "='管理'!`$B`$2&"""""
    $t.Range('AS5').Formula = "=IF($WARI!`$D`$3="""","""",$WARI!`$D`$3)"
    $t.Range('AS6').Formula = "=IF($WARI!`$D`$4="""","""",$WARI!`$D`$4)"

    # 申告の欄（Z〜AS）← フレックス申告の D〜W、時間数（AT・AU）← X・Y（どれも同じ行）
    $t.Range("Z${FIRST}:AU${LAST}").Formula = "=IF($SIN!D${FIRST}="""","""",$SIN!D${FIRST})"
    # 備考と、単位期間ごとの申告年月日・割振年月日
    for ($r = $FIRST; $r -le $LAST; $r++) {
        $av = $t.Range("AV$r")
        $label = [string]$av.Value2
        $p = [math]::Floor(($r - $FIRST) / $PERIOD) + 1        # 単位期間
        $pr = $FIRST + $p - 1                                   # 単位期間の表の行
        if ($label -like '申告年月日*') {
            $t.Range("AW$r").Formula = "=IF(ISNUMBER($SIN!`$AN`$$pr),$SIN!`$AN`$$pr,"""")"
        } elseif ($label -like '割振年月日*') {
            $t.Range("AW$r").Formula = "=IF(ISNUMBER($WARI!`$AP`$$pr),$WARI!`$AP`$$pr,"""")"
        } else {
            $av.MergeArea.Cells.Item(1, 1).Formula = "=IF($SIN!`$Z$r="""","""",$SIN!`$Z$r)"
        }
    }
    $xl.Calculation = -4105    # 自動

    # 保護：単位期間（F5〜I5）だけ書ける
    $t.Cells.Locked = $true
    $t.Range('F5:I5').Locked = $false
    $t.Activate()
    $win = $xl.ActiveWindow
    $win.ScrollRow = 1; $win.ScrollColumn = 1
    [void]$t.Range('F5').Select()
    $t.Protect('', $true, $true, $false, $false, $true, $true, $true)   # 書式（行の高さ・列の幅）の変更はできる
    $wb.Worksheets.Item(1).Activate()
    $wb.Save()
    $wb.Close($false)
    'added 申告・割振り簿'
} catch {
    Write-Host $_.InvocationInfo.PositionMessage
    throw
} finally {
    if ($src) { $src.Close($false) }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
