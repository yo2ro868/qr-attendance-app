Attribute VB_Name = "MatchingMacro"
Option Explicit
' 出勤・休暇管理データ の［申請を取り込む］ボタン：休暇申請シートと Forms の回答を読み直し、休暇等・出勤簿・申請一覧を最新にする。
' 読み込むのは保存された内容なので、同じ Excel で開いている休暇申請シートに未保存の変更があれば、先に保存するか聞く。

Public Sub ImportRequests()
    Dim wb As Workbook, unsaved As String
    For Each wb In Workbooks
        If (wb.Name Like "休暇申請シート[　 _]*" Or wb.Name Like "休暇申請_*") And Not wb.Saved Then unsaved = unsaved & vbCrLf & wb.Name
    Next wb
    If unsaved <> "" Then
        If MsgBox("保存されていない休暇申請シートがあります。保存してから取り込みますか？" & vbCrLf & unsaved & vbCrLf & vbCrLf & _
                  "［いいえ］を押すと、保存されている内容だけを取り込みます。", vbYesNo + vbQuestion) = vbYes Then
            For Each wb In Workbooks
                If (wb.Name Like "休暇申請シート[　 _]*" Or wb.Name Like "休暇申請_*") And Not wb.Saved Then wb.Save
            Next wb
        End If
    End If

    On Error GoTo Failed
    Application.Cursor = xlWait
    Application.StatusBar = "申請を取り込んでいます..."
    ThisWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone
    Application.Calculate
    ThisWorkbook.Worksheets("休暇申請シートの読み込み").Range("B5").Value = Now
    Application.StatusBar = False
    Application.Cursor = xlDefault

    Dim lo As ListObject, total As Long, waiting As Long, notes As Long
    Set lo = ThisWorkbook.Worksheets("申請一覧").ListObjects("申請一覧")
    If Not lo.DataBodyRange Is Nothing Then
        total = WorksheetFunction.CountIf(lo.ListColumns("休暇の種類").DataBodyRange, "?*")
        waiting = WorksheetFunction.CountIf(lo.ListColumns("状況").DataBodyRange, "承認待ち")
        notes = WorksheetFunction.CountIf(lo.ListColumns("確認").DataBodyRange, "?*")
    End If
    MsgBox "申請を取り込みました。" & vbCrLf & vbCrLf & _
           "申請 " & total & " 件（うち承認待ち " & waiting & " 件）" & vbCrLf & _
           "注意のある行 " & notes & " 件（申請一覧シートの「確認」の列）", vbInformation
    Exit Sub
Failed:
    Application.StatusBar = False
    Application.Cursor = xlDefault
    MsgBox "取り込めませんでした：" & Err.Description, vbExclamation
End Sub
