# 休暇申請台帳の「通知」表のうち、まだ送っていない行を Notes のメールで送る。
# 組織のパソコン（Notes が入っていて、OneDrive で台帳を同期しているもの）で動かす。
#
#   send_notices.ps1             … 数分おきに確認して送り続ける（ウィンドウを閉じると止まる）
#   send_notices.ps1 -Once       … 1回だけ確認して送る
#   send_notices.ps1 -DryRun -Once … 送らずに、送る予定の通知を表示するだけ（送信済みにもしない）
#   send_notices.ps1 -Test 宛先  … Notes でテストメールを1通送るだけ（-AskTest なら宛先を聞く）
#   -Ledger 台帳のパス           … 台帳の場所（省略すると OneDrive の中から 休暇申請台帳.xlsx を探す）
#
# 送った通知の通知番号は、このファイルと同じフォルダの 送信済み.txt に記録する。
# 送信済み.txt がないとき（初めて動かしたとき）は、その時点で台帳にある通知は送らずに送信済みにする。
param(
    [string]$Ledger = '',
    [int]$Interval = 5,
    [switch]$Once,
    [switch]$DryRun,
    [string]$Test = '',
    [switch]$AskTest
)

$ErrorActionPreference = 'Stop'
$ScriptArgs = $PSBoundParameters   # 32 ビットで起動し直すときに同じ引数を渡す
$SentFile = Join-Path $PSScriptRoot '送信済み.txt'
$LogFile = Join-Path $PSScriptRoot '通知送信ログ.txt'
$NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

function Write-Log([string]$msg) {
    $line = '{0:yyyy/MM/dd HH:mm:ss}  {1}' -f (Get-Date), $msg
    Write-Host $line
    try { Add-Content -Path $LogFile -Value $line -Encoding UTF8 } catch {}
}

# ---------- Notes ----------

# Notes のオブジェクトは型情報がないことがあるので、メソッド・プロパティは InvokeMember で呼ぶ
function Invoke-Com($obj, [string]$name, [object[]]$params = @()) {
    [System.__ComObject].InvokeMember($name, [Reflection.BindingFlags]::InvokeMethod, $null, $obj, $params)
}

function Set-ComProperty($obj, [string]$name, $value) {
    [void][System.__ComObject].InvokeMember($name, [Reflection.BindingFlags]::SetProperty, $null, $obj, @($value))
}

function Get-NotesSession {
    # Notes.NotesSession … 起動中の Notes クライアントを使う（パスワードを聞かれない）
    try { return New-Object -ComObject Notes.NotesSession } catch {}
    # Lotus.NotesSession … Notes の部品を直接使う（Notes と同じ 32/64 ビットの PowerShell が必要）
    try {
        $s = New-Object -ComObject Lotus.NotesSession
        [void](Invoke-Com $s 'Initialize')
        return $s
    } catch {}
    return $null
}

function Get-MailDatabase($session) {
    try {
        $db = Invoke-Com $session 'GetDatabase' @('', '')
        [void](Invoke-Com $db 'OpenMail')
        return $db
    } catch {
        $dir = Invoke-Com $session 'GetDbDirectory' @('')
        return Invoke-Com $dir 'OpenMailDatabase'
    }
}

function Send-NotesMail($db, [string]$to, [string]$subject, [string]$body) {
    $doc = Invoke-Com $db 'CreateDocument'
    [void](Invoke-Com $doc 'ReplaceItemValue' @('Form', 'Memo'))
    [void](Invoke-Com $doc 'ReplaceItemValue' @('SendTo', $to))
    [void](Invoke-Com $doc 'ReplaceItemValue' @('Subject', $subject))
    $rt = Invoke-Com $doc 'CreateRichTextItem' @('Body')
    foreach ($line in ($body -split "\r?\n")) {
        [void](Invoke-Com $rt 'AppendText' @($line))
        [void](Invoke-Com $rt 'AddNewLine' @(1))
    }
    Set-ComProperty $doc 'SaveMessageOnSend' $true
    [void](Invoke-Com $doc 'Send' @($false))
}

function Open-Notes {
    $session = Get-NotesSession
    if (-not $session) {
        # 64 ビットの PowerShell で Notes を使えないときは、32 ビットの PowerShell でやり直す
        $ps32 = Join-Path $env:WINDIR 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
        if ([Environment]::Is64BitProcess -and (Test-Path $ps32)) {
            Write-Log '32 ビットの PowerShell で起動し直します'
            $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"")
            foreach ($k in $ScriptArgs.Keys) {
                $v = $ScriptArgs[$k]
                if ($v -is [switch]) { if ($v) { $argList += "-$k" } } else { $argList += "-$k"; $argList += "`"$v`"" }
            }
            & $ps32 @argList
            exit $LASTEXITCODE
        }
        throw 'Notes に接続できません。Notes を起動してサインインしてから、もう一度実行してください。'
    }
    return Get-MailDatabase $session
}

# ---------- 台帳（xlsx）を読む ----------

function Find-Ledger {
    if ($Ledger) {
        if (Test-Path -LiteralPath $Ledger) { return (Resolve-Path -LiteralPath $Ledger).Path }
        throw "台帳が見つかりません: $Ledger"
    }
    foreach ($root in @($env:OneDriveCommercial, $env:OneDrive) | Where-Object { $_ }) {
        $hit = Get-ChildItem -LiteralPath $root -Filter '休暇申請台帳.xlsx' -Recurse -File -ErrorAction SilentlyContinue |
            Select-Object -First 1
        if ($hit) { return $hit.FullName }
    }
    throw 'OneDrive の中に 休暇申請台帳.xlsx が見つかりません。OneDrive の同期を確認するか、-Ledger で場所を指定してください。'
}

function Read-ZipXml($zip, [string]$path) {
    $entry = $zip.GetEntry($path)
    if (-not $entry) { return $null }
    $reader = New-Object System.IO.StreamReader($entry.Open(), [Text.Encoding]::UTF8)
    try { $xml = New-Object System.Xml.XmlDocument; $xml.LoadXml($reader.ReadToEnd()); return $xml }
    finally { $reader.Close() }
}

function Get-CellText($node, $nsm) {
    $text = ''
    foreach ($t in $node.SelectNodes('.//m:t', $nsm)) { $text += $t.InnerText }
    return $text
}

# 「通知」シートを読み、見出しの名前をキーにした行の一覧を返す
function Read-Notices([string]$path) {
    Add-Type -AssemblyName System.IO.Compression
    # OneDrive や Excel が開いていても読めるように、共有モードで読み込んでから開く
    $fs = [System.IO.File]::Open($path, 'Open', 'Read', 'ReadWrite')
    $ms = New-Object System.IO.MemoryStream
    try { $fs.CopyTo($ms) } finally { $fs.Close() }
    $zip = New-Object System.IO.Compression.ZipArchive($ms, [System.IO.Compression.ZipArchiveMode]::Read)
    try {
        $nsm = New-Object System.Xml.XmlNamespaceManager((New-Object System.Xml.NameTable))
        $nsm.AddNamespace('m', $NS)
        $nsm.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $nsm.AddNamespace('p', 'http://schemas.openxmlformats.org/package/2006/relationships')

        $wb = Read-ZipXml $zip 'xl/workbook.xml'
        $sheet = $wb.SelectSingleNode("//m:sheets/m:sheet[@name='通知']", $nsm)
        if (-not $sheet) { throw '台帳に「通知」シートがありません（新しい様式の台帳か確認してください）' }
        $rid = $sheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $rels = Read-ZipXml $zip 'xl/_rels/workbook.xml.rels'
        $target = $rels.SelectSingleNode("//p:Relationship[@Id='$rid']", $nsm).GetAttribute('Target')
        $sheetPath = if ($target.StartsWith('/')) { $target.TrimStart('/') } else { 'xl/' + $target }

        $shared = @()
        $sst = Read-ZipXml $zip 'xl/sharedStrings.xml'
        if ($sst) { foreach ($si in $sst.SelectNodes('//m:si', $nsm)) { $shared += , (Get-CellText $si $nsm) } }

        $rows = @{}
        $ws = Read-ZipXml $zip $sheetPath
        foreach ($c in $ws.SelectNodes('//m:sheetData/m:row/m:c', $nsm)) {
            $ref = $c.GetAttribute('r')
            if ($ref -notmatch '^([A-Z]+)(\d+)$') { continue }
            $colIdx = 0
            foreach ($ch in $Matches[1].ToCharArray()) { $colIdx = $colIdx * 26 + ([int]$ch - 64) }
            $rowNo = [int]$Matches[2]
            $v = $c.SelectSingleNode('m:v', $nsm)
            $value = switch ($c.GetAttribute('t')) {
                's' { if ($v) { $shared[[int]$v.InnerText] } else { '' } }
                'inlineStr' { Get-CellText $c $nsm }
                default { if ($v) { $v.InnerText } else { '' } }
            }
            if (-not $rows.ContainsKey($rowNo)) { $rows[$rowNo] = @{} }
            $rows[$rowNo][$colIdx] = $value
        }
        if (-not $rows.ContainsKey(1)) { return @() }
        $header = $rows[1]
        $result = @()
        foreach ($rowNo in ($rows.Keys | Where-Object { $_ -gt 1 } | Sort-Object)) {
            $item = @{}
            foreach ($colIdx in $header.Keys) { $item[$header[$colIdx]] = [string]$rows[$rowNo][$colIdx] }
            if ($item['通知番号']) { $result += , $item }
        }
        return $result
    } finally {
        $zip.Dispose()
        $ms.Dispose()
    }
}

# ---------- 送信 ----------

function Get-SentIds {
    $set = New-Object 'System.Collections.Generic.HashSet[string]'
    if (Test-Path -LiteralPath $SentFile) {
        foreach ($line in (Get-Content -LiteralPath $SentFile -Encoding UTF8)) {
            $id = ($line -split "`t")[0].Trim()
            if ($id) { [void]$set.Add($id) }
        }
    }
    return , $set
}

function Add-Sent([string]$id, [string]$note) {
    Add-Content -LiteralPath $SentFile -Value ("{0}`t{1:yyyy/MM/dd HH:mm}`t{2}" -f $id, (Get-Date), $note) -Encoding UTF8
}

function Send-Pending([string]$path, $db) {
    $notices = @(Read-Notices $path)
    if (-not (Test-Path -LiteralPath $SentFile)) {
        if ($DryRun) {
            Write-Log "（初回）台帳にある通知 $($notices.Count) 件は、本番の初回実行では送らずに送信済みにします"
            return
        }
        Set-Content -LiteralPath $SentFile -Value "# 送った通知の通知番号（消さないでください）" -Encoding UTF8
        foreach ($n in $notices) { Add-Sent $n['通知番号'] '初回のため送らずに記録' }
        Write-Log "初回のため、台帳にある通知 $($notices.Count) 件は送らずに送信済みにしました"
        return
    }
    $sent = Get-SentIds
    $pending = @($notices | Where-Object { -not $sent.Contains($_['通知番号']) })
    foreach ($n in $pending) {
        $id, $to, $subject = $n['通知番号'], ([string]$n['宛先']).Trim(), [string]$n['件名']
        if ($DryRun) {
            Write-Log "（送る予定）$id → $to ： $subject"
            continue
        }
        if (-not $to) {
            Write-Log "宛先が空欄なので送りません: $id $subject （設定シートの B1 に管理者の通知先を入れてください）"
            Add-Sent $id '宛先なしで送らず'
            continue
        }
        try {
            Send-NotesMail $db $to $subject $n['本文']
            Add-Sent $id $to
            Write-Log "送信しました: $id → $to ： $subject"
        } catch {
            Write-Log "送信できませんでした（次の確認でもう一度送ります）: $id → $to ： $($_.Exception.Message)"
        }
    }
    if ($pending.Count -eq 0 -and ($Once -or $DryRun)) { Write-Log '送っていない通知はありません' }
}

# ---------- 本体 ----------

if ($AskTest -and -not $Test) {
    $Test = (Read-Host 'テストメールを送る Notes のアドレスを入力して Enter').Trim()
    if (-not $Test) { Write-Host '宛先が入力されなかったので終わります'; exit 1 }
    $ScriptArgs['Test'] = $Test
    [void]$ScriptArgs.Remove('AskTest')
}
if ($Test) {
    $db = Open-Notes
    Send-NotesMail $db $Test '【休暇申請】通知送信のテスト' "休暇申請の通知送信のテストメールです。`nこのメールが届けば、Notes からの送信はできています。"
    Write-Log "テストメールを送りました: $Test"
    exit 0
}

$path = Find-Ledger
Write-Log "台帳: $path"
if ($Once -or $DryRun) {
    $db = if ($DryRun) { $null } else { Open-Notes }
    Send-Pending $path $db
    exit 0
}
# パソコンの起動直後など、Notes がまだ使えないときは 1 分おきに待つ
$db = $null
while (-not $db) {
    try { $db = Open-Notes }
    catch { Write-Log "$($_.Exception.Message)（1 分後にもう一度試します）"; Start-Sleep -Seconds 60 }
}
Write-Log "$Interval 分おきに確認します。止めるときはこのウィンドウを閉じます。"
while ($true) {
    try { Send-Pending $path $db }
    catch { Write-Log "確認できませんでした（$Interval 分後にもう一度確認します）: $($_.Exception.Message)" }
    Start-Sleep -Seconds ($Interval * 60)
}
