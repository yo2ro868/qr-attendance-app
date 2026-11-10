# 休暇申請シートのひな形に、超過勤務等命令簿の「超勤ひな形」シート（非表示）を入れる。
# 命令簿の様式と計算式は、いま使っている命令簿の Excel（-Source）の月のシートをそのまま写す（日ごとに手で直した式があるため、作り直さない）。
# 写すときに変えるのは次のところだけ：
#   ・W1（年月）：シート名から読む CELL("filename") の式をやめ、マクロ（休暇申請シートの作成.xlsm の［超過勤務の月を追加］）が月の1日を書く
#     （CELL("filename") は Web の Excel では使えないため）
#   ・前月のシートの最終日の退庁時刻（'R8.10'!L65 など）：このシートの BH1 にする。マクロが前月のシートの BH2（その月の最終日の退庁時刻）を指す式を書く
#   ・休日・代休日（H 列）：祝日・年末年始は自動で ○（名前「休日」＝管理シートの祝日と年末年始）。代休日などは本人が ○ を書く
#   ・登庁時刻・退庁時刻・○・超勤代休・備考・所属長の確認日は空にする。氏名は管理シートから
#   ・隠し列 BJ〜 に、出勤・休暇管理データ が Power Query で読む表「超勤日別」（1日1行、時間は分）を足す
# 使い方：add_overtime_sheet.ps1 -Template 休暇申請シート（ひな形）.xlsx -Source 超過勤務命令簿.xlsx [-Sheet R8.11]
#   -Sheet を省くと、R年.月 の名前のシートのうち、いちばん新しい月のシートを使う。
param([Parameter(Mandatory = $true)][string]$Template, [Parameter(Mandatory = $true)][string]$Source, [string]$Sheet = '')

$ErrorActionPreference = 'Stop'
$Template = [IO.Path]::GetFullPath($Template)
$Source = [IO.Path]::GetFullPath($Source)
if (-not (Test-Path $Source)) { throw "超過勤務等命令簿のファイルが見つかりません：$Source" }

$PREV = '$BH$1'      # 前月の最終日の退庁時刻（マクロが式を書く）
$table = 'BJ'        # 表「超勤日別」の左上（見出し）
$TOP, $BOTTOM = 5, 66    # 日ごとの行（1日に2行。d 日は 5+2(d-1) 行目から）

$xl = New-Object -ComObject Excel.Application
$xl.DisplayAlerts = $false
$xl.Visible = $false
try {
    $src = $xl.Workbooks.Open($Source, 0, $true)
    if (-not $Sheet) {
        $best = $null; $bestKey = -1
        foreach ($s in $src.Worksheets) {
            if ($s.Visible -eq -1 -and $s.Name -match '^R(\d+)\.(\d+)$') {
                $k = [int]$Matches[1] * 100 + [int]$Matches[2]
                if ($k -gt $bestKey) { $bestKey = $k; $best = $s.Name }
            }
        }
        if (-not $best) { throw '命令簿に R年.月 の名前のシートがありません（-Sheet で指定してください）' }
        $Sheet = $best
    }
    $ws = $src.Worksheets.Item($Sheet)

    # ----- 写す前に、命令簿のほうで直す（保存はしない） -----
    $ws.Unprotect()
    $ws.Range('W1').MergeArea.ClearContents()
    # 前月のシートを指す式（'R8.10'!L65 など）を BH1 に。ほかのシートを指す式がほかにあれば止める
    $formulas = $ws.UsedRange.SpecialCells(-4123)    # xlCellTypeFormulas
    $changed = 0
    foreach ($c in $formulas.Cells) {
        $f = [string]$c.Formula
        if ($f.Contains('!')) {
            $g = [regex]::Replace($f, "('[^']+'|[^\s'!=(),+\-*/<>&""]+)!\`$?L\`$?\d+", $PREV.Replace('$', '$$'))
            if ($g.Contains('!')) { throw "前月の退庁時刻のほかに、ほかのシートを指す式があります：$($c.Address(0,0)) $f" }
            $c.Formula = $g
            $changed++
        }
    }
    if ($changed -eq 0) { Write-Warning '前月のシートを指す式が見つかりませんでした' }
    # 入力欄を空にする（A・B の登庁・退庁の標準時刻と I の休憩時間は残す）
    foreach ($a in @("D${TOP}:G${BOTTOM}", "K${TOP}:L${BOTTOM}", "Y${TOP}:AA${BOTTOM}", "AC${TOP}:AE${BOTTOM}", 'L69:O71', 'AC69:AE69')) {
        $ws.Range($a).ClearContents()
    }
    # 休日・代休日：祝日・年末年始は自動で ○
    for ($r = $TOP; $r -le $BOTTOM; $r += 2) {
        $ws.Range("H$r").Formula = "=IF(COUNTIF(休日,J$r)>0,""○"","""")"
    }
    $ws.Range('AC69').NumberFormat = 'ggge"年"m"月"d"日"'
    # 列の幅（ポイント）。列の幅の数は標準のフォントの文字数なので、標準のフォントが違うひな形に写すと幅が変わる。写したあとで同じ幅に直す
    $widths = @{}
    for ($i = 1; $i -le 80; $i++) { $widths[$i] = [double]$ws.Columns.Item($i).Width }

    # ----- ひな形に写す -----
    $wb = $xl.Workbooks.Open($Template)
    foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq '超勤ひな形') { throw 'ひな形にもう 超勤ひな形 シートがあります（make_form.py で作り直してから実行してください）' } }
    $usage = $wb.Worksheets.Item('使い方')
    $ws.Copy($usage)
    $t = $wb.ActiveSheet     # 写したシート（Worksheets の番号は非表示のシートでずれることがあるので使わない）
    $t.Name = '超勤ひな形'
    $src.Close($false)
    $src = $null

    # 名前：写すと命令簿の「休日」「振替」がこのシートだけの名前として付いてくるので、ひな形の管理シート（祝日・年末年始、○）を指すように直す
    # （月のシートはこのシートを同じブックの中で写すので、名前もそのまま付いていく）
    $targets = @{ '休日' = "='管理'!`$D`$1:`$E`$200"; '振替' = "='管理'!`$X`$1:`$X`$2" }
    foreach ($n in @($wb.Names)) {
        $short = ($n.Name -split '!')[-1]
        if ($n.RefersTo -match '\[' -or $n.RefersTo -match '#REF!') {
            if ($targets.ContainsKey($short)) { $n.RefersTo = $targets[$short] }
            elseif ($short -ne 'Print_Area') { throw "命令簿のブックを指す名前があります：$($n.Name) $($n.RefersTo)" }
        }
    }
    foreach ($c in $t.UsedRange.SpecialCells(-4123).Cells) {
        if (([string]$c.Formula).Contains('[')) { throw "命令簿のブックを指す式が残っています：$($c.Address(0,0)) $($c.Formula)" }
    }

    for ($i = 1; $i -le 80; $i++) {
        $col = $t.Columns.Item($i)
        if ($widths[$i] -gt 0 -and $col.Width -gt 0) {
            for ($k = 0; $k -lt 3 -and [math]::Abs($col.Width - $widths[$i]) -gt 0.5; $k++) {
                $col.ColumnWidth = [math]::Round($col.ColumnWidth * $widths[$i] / $col.Width, 2)
            }
        }
    }
    $t.Range('R69').Formula = "='管理'!`$B`$2&"""""     # 氏名
    # 前月・当月の最終日の退庁時刻
    $t.Range('BG1').Value2 = '前月の最終日の退庁時刻'
    $t.Range('BG2').Value2 = 'この月の最終日の退庁時刻'
    $last = "INDEX(`$L`$${TOP}:`$L`$${BOTTOM},2*DAY(EOMONTH(`$W`$1,0))-1)"
    $t.Range('BH2').Formula = "=IF(`$W`$1="""","""",IF($last="""","""",$last))"
    $t.Range('BH1:BH2').NumberFormat = '[h]:mm'

    # ----- 表「超勤日別」：1日1行。時間・時刻は分（Power Query で型の違いが出ないように数で持つ）、日付は文字 -----
    $up = { param($c) "INDEX(`$${c}`$${TOP}:`$${c}`$${BOTTOM},2*[@日]-1)" }      # その日の上の行
    $down = { param($c) "INDEX(`$${c}`$${TOP}:`$${c}`$${BOTTOM},2*[@日])" }      # その日の下の行
    $minutes = { param($c) "=IF($(& $up $c)="""","""",ROUND($(& $up $c)*1440,0))" }
    $mark = { param($c) "=$(& $up $c)&""""" }
    $both = { param($c) "=ROUND((N($(& $up $c))+N($(& $down $c)))*1440,0)" }
    $cols = [ordered]@{
        '日'       = $null
        '年月'     = '=IF($W$1="","",TEXT($W$1,"yyyy-mm"))'
        '日付'     = '=IF(OR($W$1="",[@日]>DAY(EOMONTH($W$1,0))),"",TEXT($W$1+[@日]-1,"yyyy-mm-dd"))'
        '確認日'   = '=IF(ISNUMBER($AC$69),TEXT($AC$69,"yyyy-mm-dd"),"")'
        '所属部課' = '=$L$69&""'
        '登庁時刻' = & $minutes 'K'
        '退庁時刻' = & $minutes 'L'
        '休憩時間' = & $minutes 'I'
        '標準の登庁時刻' = & $minutes 'A'
        '標準の退庁時刻' = & $minutes 'B'
        '振替日（1日）' = & $mark 'D'
        '振替日（1日・日）' = & $mark 'E'
        '振替日（午前）' = & $mark 'F'
        '振替日（午後）' = & $mark 'G'
        '休日・代休日' = & $mark 'H'
        '125'      = & $both 'M'
        '150'      = & $both 'N'
        '150（60H前）' = & $both 'O'
        '150（+25）'  = & $both 'P'
        '150（+15）'  = & $both 'Q'
        '135'      = & $both 'R'
        '160'      = & $both 'S'
        '175'      = & $both 'T'
        '175（+25）'  = & $both 'U'
        '175（+15）'  = & $both 'V'
        '休日勤務135' = & $both 'X'
        '超勤代休時間の勤務' = & $both 'Y'
        '超勤代休（+25）' = & $both 'Z'
        '超勤代休（+15）' = & $both 'AA'
        '備考'     = "=$(& $up 'AC')&"""""
    }
    $c0 = $t.Range("${table}4").Column
    $i = 0
    foreach ($h in $cols.Keys) { $t.Cells.Item(4, $c0 + $i).Value2 = [string]$h; $i++ }
    for ($d = 1; $d -le 31; $d++) { $t.Cells.Item(4 + $d, $c0).Value2 = [double]$d }
    $rng = $t.Range($t.Cells.Item(4, $c0), $t.Cells.Item(35, $c0 + $cols.Count - 1))
    $lo = $t.ListObjects.Add(1, $rng, $null, 1)     # xlSrcRange, 見出しあり
    $lo.Name = '超勤日別'
    $lo.TableStyle = ''
    $i = 0
    foreach ($h in $cols.Keys) {
        if ($cols[$h]) { $lo.ListColumns.Item($i + 1).DataBodyRange.Formula = $cols[$h] }
        $i++
    }
    $t.Range($t.Cells.Item(1, $t.Range('BG1').Column), $t.Cells.Item(1, $c0 + $cols.Count - 1)).EntireColumn.Hidden = $true

    # ----- 入力欄の保護を外してシートを保護する（命令簿の保護の設定はそのまま。所属部課と所属長の確認欄・休日の欄を足す） -----
    foreach ($a in @('L69:O71', 'AC69:AE70', "H${TOP}:H${BOTTOM}")) { $t.Range($a).Locked = $false }
    $t.Activate()
    $win = $xl.ActiveWindow
    $win.FreezePanes = $false
    $win.ScrollRow = 1; $win.ScrollColumn = 1
    $win.SplitColumn = 0; $win.SplitRow = 4
    $win.FreezePanes = $true
    [void]$t.Range('K5').Select()
    $t.Protect('', $true, $true, $false, $false, $true, $true, $true)   # 書式（行の高さ・列の幅）の変更はできる
    $wb.Worksheets.Item(1).Activate()
    $t.Visible = 2      # xlSheetVeryHidden：本人・所属長には見えない。マクロがこれを写して月のシートを作る
    $wb.Save()
    $wb.Close($false)
    "added 超勤ひな形 from $Sheet (formulas changed: $changed)"
} catch {
    Write-Host $_.InvocationInfo.PositionMessage
    throw
} finally {
    if ($src) { $src.Close($false) }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
