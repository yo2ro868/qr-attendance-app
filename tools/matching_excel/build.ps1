# 照合用Excel.xlsx を作り直す。
#   1. make_matching.py      … openpyxl でブック（様式・計算式）を作る
#   2. add_ignored_errors.py … 入力欄の「数値が文字列として保存されています」警告を出さない設定を入れる
#   3. finalize.ps1          … Excel で再計算して保存する（開いたときに値が入った状態にする）
# 使い方：powershell -ExecutionPolicy Bypass -File tools\matching_excel\build.ps1 [-Out 出力先.xlsx]
param([string]$Out = (Join-Path (Resolve-Path (Join-Path $PSScriptRoot '..\..')) '照合用Excel.xlsx'))

$ErrorActionPreference = 'Stop'
$raw = Join-Path $env:TEMP 'matching_excel_raw.xlsx'
$tmp = Join-Path $env:TEMP 'matching_excel_build.xlsx'

python (Join-Path $PSScriptRoot 'make_matching.py') $raw
if ($LASTEXITCODE -ne 0) { throw 'make_matching.py が失敗しました' }
python (Join-Path $PSScriptRoot 'add_ignored_errors.py') $raw $tmp
if ($LASTEXITCODE -ne 0) { throw 'add_ignored_errors.py が失敗しました' }
& (Join-Path $PSScriptRoot 'finalize.ps1') -Path $tmp | Out-Null

Copy-Item $tmp $Out -Force
Remove-Item $raw, $tmp -ErrorAction SilentlyContinue
"作成しました: $Out"
