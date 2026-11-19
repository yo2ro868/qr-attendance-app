# 申請・報告シートのひな形に、様式どおりの完成版のシート「申告・割振り簿」を入れる。
# 総務課の様式の Excel（-Source。シートの保護を外したもの）の「フレックスタイム制」シートをそのまま写し、次のところだけ変える：
#   ・申告の欄（Z〜AS 列：コアタイム・勤務・休憩の時と分）：本人が「フレックス（申告の内容）」シートに書いた値を表示する式
#   ・時間数（AT・AU 列）：フレックス（申告の内容）シートで計算した時間数（空の欄でも計算が止まらないように）
#   ・割振りの欄（D〜Y 列）：様式の式のまま（申告の欄を写す）
#   ・1年間の始期（F4〜I4）：管理シートの始期。単位期間（F5〜I5）：今日を含む単位期間（書き換えると、その期間の合計と表示になる）
#   ・所属・氏名・所属長氏名・異動後の所属長氏名（AS3〜AS6）：フレックス（単位期間・承認）シートの上の欄から
#   ・備考（AV 列）：フレックス（申告の内容）シートの備考。単位期間ごとの申告年月日・割振年月日（AW）はフレックス（単位期間・承認）の表から
#   ・様式のブックの外を指す名前（月午前 など。使われていない）は消す。入力規則は消す
#   ・単位期間（F5〜I5）のほかは入力できないように保護する（パスワードなし）
#   ・本人・所属長には見せない（veryHidden）。管理者は 申請・報告シートの作成.xlsm の［申告・割振り簿を見る］で見る
# 使い方：add_flex_book.ps1 -Template 申請・報告シート（ひな形）.xlsx -Source 申告・割振り簿（原本）.xlsx
param([Parameter(Mandatory = $true)][string]$Template, [Parameter(Mandatory = $true)][string]$Source)

$ErrorActionPreference = 'Stop'
$Template = [IO.Path]::GetFullPath($Template)
$Source = [IO.Path]::GetFullPath($Source)
if (-not (Test-Path $Source)) { throw "申告・割振り簿の様式のファイルが見つかりません：$Source" }

$FIRST, $LAST, $PERIOD = 12, 403, 28
$SIN = "'フレックス（申告の内容）'"
$PER = "'フレックス（単位期間・承認）'"     # 単位期間の表：9〜22 行目が期間 1〜14（B 開始日、F 申告年月日、H 割振年月日）。C3 所属、C5 所属長氏名、C6 異動後

# Excel が計算などで忙しいと COM の呼び出しが断られる（RPC_E_CALL_REJECTED）ので、少し待ってやり直す
function Retry([scriptblock]$sb) {
    for ($i = 0; $i -lt 50; $i++) {
        try { return & $sb } catch { if ($_.Exception.HResult -eq -2147418111) { Start-Sleep -Milliseconds 200 } else { throw } }
    }
    throw 'Excel が応答しません'
}

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

    # ひな形は式が多く、開くと再計算で Excel が忙しくなる。手動にしてから開き、保存の前に自動に戻す
    $xl.Calculation = -4135
    $wb = $xl.Workbooks.Open($Template)
    foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq '申告・割振り簿') { throw 'ひな形にもう 申告・割振り簿 シートがあります（make_form.py で作り直してから実行してください）' } }
    $ws.Copy([Type]::Missing, $wb.Worksheets.Item('フレックス（申告の内容）'))
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
    # 文字が欄に入りきらないところ（様式のまま）：曜日の欄と「異動後の所属長氏名」は縮めて入れ、P5・P6（合計の分。計算用）は表示しない
    $t.Range('J3:J6').ShrinkToFit = $true
    $t.Range('AN6').MergeArea.ShrinkToFit = $true
    $t.Range('P5:P6').NumberFormat = ';;;'
    $xl.Calculation = -4135    # 手動（式を入れるあいだ再計算しない）

    # 1年間の始期・単位期間
    # 1年間の始期：フレックス（申告の内容）の日ごとの表の最初の日（単位期間1の開始日。なければ管理シートの始期）。行の日がフレックス申告とそろう
    $t.Range('F4').Formula = "=YEAR($SIN!`$B`$$FIRST)"
    $t.Range('H4').Formula = "=MONTH($SIN!`$B`$$FIRST)"
    $t.Range('I4').Formula = "=DAY($SIN!`$B`$$FIRST)"
    $now = "IFERROR(LOOKUP(2,1/(($PER!`$B`$9:`$B`$22<>"""")*($PER!`$B`$9:`$B`$22<=TODAY())),$PER!`$B`$9:`$B`$22),$SIN!`$B`$$FIRST)"
    $t.Range('F5').Formula = "=YEAR($now)"
    $t.Range('H5').Formula = "=MONTH($now)"
    $t.Range('I5').Formula = "=DAY($now)"
    # 所属・氏名・所属長氏名・異動後の所属長氏名
    $t.Range('AS3').Formula = "=IF($PER!`$C`$3="""","""",$PER!`$C`$3)"
    $t.Range('AS4').Formula = "='管理'!`$B`$2&"""""
    $t.Range('AS5').Formula = "=IF($PER!`$C`$5="""","""",$PER!`$C`$5)"
    $t.Range('AS6').Formula = "=IF($PER!`$C`$6="""","""",$PER!`$C`$6)"

    # 申告の欄（Z〜AS）← フレックス（申告の内容）の D〜W、時間数（AT・AU）← X・Y（どれも同じ行）
    $t.Range("Z${FIRST}:AU${LAST}").Formula = "=IF($SIN!D${FIRST}="""","""",$SIN!D${FIRST})"
    # 備考と、単位期間ごとの申告年月日・割振年月日
    for ($r = $FIRST; $r -le $LAST; $r++) {
        Retry {
        $av = $t.Range("AV$r")
        $label = [string]$av.Value2
        $p = [math]::Floor(($r - $FIRST) / $PERIOD) + 1        # 単位期間
        $pr = $FIRST + $p - 1                                   # 単位期間の表の行
        # その行の日を含む単位期間（フレックス（申告の内容）の A 列の期間番号）の申告年月日・割振年月日
        $k = "$SIN!`$A$r"
        if ($label -like '申告年月日*') {
            $t.Range("AW$r").Formula = "=IFERROR(IF(ISNUMBER(INDEX($PER!`$F`$9:`$F`$22,$k)),INDEX($PER!`$F`$9:`$F`$22,$k),""""),"""")"
        } elseif ($label -like '割振年月日*') {
            $t.Range("AW$r").Formula = "=IFERROR(IF(ISNUMBER(INDEX($PER!`$H`$9:`$H`$22,$k)),INDEX($PER!`$H`$9:`$H`$22,$k),""""),"""")"
        } else {
            $av.MergeArea.Cells.Item(1, 1).Formula = "=IF($SIN!`$Z$r="""","""",$SIN!`$Z$r)"
        }
        }
    }

    # 保護：単位期間（F5〜I5）だけ書ける
    $t.Cells.Locked = $true
    $t.Range('F5:I5').Locked = $false
    $t.Activate()
    $win = $xl.ActiveWindow
    $win.ScrollRow = 1; $win.ScrollColumn = 1
    [void]$t.Range('F5').Select()
    $t.Protect('', $true, $true, $false, $false, $true, $true, $true)   # 書式（行の高さ・列の幅）の変更はできる
    $t.Tab.Color = 0xB5752E      # フレックスの色（青）
    $wb.Worksheets.Item(1).Activate()
    # 管理者だけが見るシート：本人・所属長には見せない（veryHidden。Excel の［再表示］にも出ない）。
    # 管理者は 申請・報告シートの作成.xlsm の［申告・割振り簿を見る］で、読み取り専用で開いて表示する
    $t.Visible = 2
    Retry { $xl.Calculation = -4105 } | Out-Null    # 自動に戻してから保存する（ファイルに計算方法が残るため）
    Retry { $wb.Save() } | Out-Null
    Retry { $wb.Close($false) } | Out-Null
    'added 申告・割振り簿'
} catch {
    Write-Host $_.InvocationInfo.PositionMessage
    throw
} finally {
    if ($src) { $src.Close($false) }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
