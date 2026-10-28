param([string]$Path)
# Recalculate in Excel and save, so the file opens with values already filled in
$x = New-Object -ComObject Excel.Application
$x.Visible = $false
$x.DisplayAlerts = $false
$wb = $x.Workbooks.Open($Path)
$x.CalculateFull()
$wb.Worksheets.Item(1).Activate()
$wb.Save()
$wb.Close($false)
$x.Quit()
[Runtime.InteropServices.Marshal]::ReleaseComObject($x) | Out-Null
'finalized'
