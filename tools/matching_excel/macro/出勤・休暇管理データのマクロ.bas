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

' ［超過勤務を転記］ボタン（超過勤務シート）：休暇申請シートの超過勤務等命令簿を読み直し、B2 の月の所属長の確認済みの分を
' 「超過勤務」シート（1人1か月の計）と「超過勤務の日別」シート（日ごとの行と照合）に値で写す。同じ人・同じ月を前に写していれば置き換える。
' 確認待ち（命令簿の所属長の確認日が空）の人は写さない。
Public Sub TransferOvertime()
    Dim ws As Worksheet, ym As String, target As Variant
    Set ws = ThisWorkbook.Worksheets("超過勤務")
    target = ws.Range("B2").Value
    If Not IsDate(target) Then
        MsgBox "転記する月を B2 に入れてください（例 2026/11）。", vbExclamation
        Exit Sub
    End If
    ym = Format(target, "yyyy-mm")

    On Error GoTo Failed
    Application.Cursor = xlWait
    Application.StatusBar = "超過勤務等命令簿を読み込んでいます..."
    ThisWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone
    Application.Calculate
    Application.StatusBar = False
    Application.Cursor = xlDefault

    Dim srcMonth As ListObject, srcDay As ListObject, dstMonth As ListObject, dstDay As ListObject
    Set srcMonth = ThisWorkbook.Worksheets("超過勤務の読み込み").ListObjects("超勤_月")
    Set srcDay = ThisWorkbook.Worksheets("超過勤務の読み込み").ListObjects("超勤_日別")
    Set dstMonth = ws.ListObjects("転記_月計")
    Set dstDay = ThisWorkbook.Worksheets("超過勤務の日別").ListObjects("転記_日別")

    ' その月の確認済み・確認待ちの人
    Dim done As Object, waiting As String, nWaiting As Long, r As Long, before As Long
    Set done = CreateObject("Scripting.Dictionary")
    If Not srcMonth.DataBodyRange Is Nothing Then
        For r = 1 To srcMonth.ListRows.Count
            If CellText(srcMonth, r, "年月") = ym Then
                If CellText(srcMonth, r, "状況") = "確認済み" Then
                    done(CellText(srcMonth, r, "出勤用ID")) = True
                Else
                    nWaiting = nWaiting + 1
                    If nWaiting <= 15 Then waiting = waiting & vbCrLf & "  " & CellText(srcMonth, r, "出勤用ID") & " " & CellText(srcMonth, r, "氏名")
                End If
            End If
        Next r
    End If
    If nWaiting > 15 Then waiting = waiting & vbCrLf & "  ほか " & (nWaiting - 15) & " 人"
    If done.Count = 0 Then
        MsgBox Format(target, "yyyy年m月") & "の所属長の確認済みの超過勤務等命令簿がありません。" & _
               IIf(nWaiting > 0, vbCrLf & vbCrLf & "確認待ち " & nWaiting & " 人：" & waiting, ""), vbInformation
        Exit Sub
    End If
    before = CountRows(dstMonth, ym, done)
    If MsgBox(Format(target, "yyyy年m月") & "の超過勤務を転記します。" & vbCrLf & vbCrLf & _
              "所属長の確認済み " & done.Count & " 人" & IIf(before > 0, "（うち前に転記した " & before & " 人は今の内容で置き換えます）", "") & vbCrLf & _
              IIf(nWaiting > 0, "確認待ち " & nWaiting & " 人は転記しません：" & waiting & vbCrLf, "") & vbCrLf & _
              "転記しますか？", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Dim stamp As Date
    stamp = Now
    ReplaceRows dstMonth, srcMonth, ym, done, stamp, Array("年月", "出勤用ID")
    ReplaceRows dstDay, srcDay, ym, done, stamp, Array("年月", "出勤用ID", "日付")
    Application.Calculate
    Application.ScreenUpdating = True

    ' 注意のある日（この月の転記分）
    Dim notes As Long
    For r = 1 To dstDay.ListRows.Count
        If CellText(dstDay, r, "年月") = ym And CellText(dstDay, r, "注意") <> "" Then notes = notes + 1
    Next r
    ws.Activate
    MsgBox Format(target, "yyyy年m月") & "の超過勤務を転記しました（" & done.Count & " 人）。" & vbCrLf & _
           "注意のある日 " & notes & " 日（「超過勤務の日別」シートの「注意」の列）", vbInformation
    Exit Sub
Failed:
    Application.StatusBar = False
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
    MsgBox "転記できませんでした：" & Err.Description, vbExclamation
End Sub

' 表の r 行目（1から）の列の値を文字で
Private Function CellText(lo As ListObject, r As Long, col As String) As String
    CellText = CStr(lo.ListColumns(col).DataBodyRange.Cells(r).Value)
End Function

' 写す先の表にある、その月・その人たちの行の人数
Private Function CountRows(lo As ListObject, ym As String, ids As Object) As Long
    Dim r As Long
    If lo.DataBodyRange Is Nothing Then Exit Function
    For r = 1 To lo.ListRows.Count
        If CellText(lo, r, "年月") = ym And ids.Exists(CellText(lo, r, "出勤用ID")) Then CountRows = CountRows + 1
    Next r
End Function

' 写す先の表 dst から、その月・その人たちの行を除き、読み込んだ表 src のその月・その人たちの行を足して、並べ替えて書き直す。
' 見出しが同じ名前の列を写す。dst にだけある列：転記日時（今の日時）、式の列（照合。全部の行に同じ式を入れ直す）
Private Sub ReplaceRows(dst As ListObject, src As ListObject, ym As String, ids As Object, stamp As Date, sortKeys As Variant)
    Dim nCols As Long, c As Long, r As Long, k As Long
    nCols = dst.ListColumns.Count
    ' 式の列（1行目の式を覚えておく）と、src の同じ列
    Dim formulas() As String, srcCol() As Long
    ReDim formulas(1 To nCols)
    ReDim srcCol(1 To nCols)
    For c = 1 To nCols
        If dst.ListColumns(c).DataBodyRange.Cells(1).HasFormula Then formulas(c) = dst.ListColumns(c).DataBodyRange.Cells(1).Formula
        srcCol(c) = 0
        For k = 1 To src.ListColumns.Count
            If src.ListColumns(k).Name = dst.ListColumns(c).Name Then srcCol(c) = k
        Next k
    Next c
    ' 残す行（値の列だけ）
    Dim keep As New Collection, rowVals() As Variant, idCol As Long, ymCol As Long
    idCol = dst.ListColumns("出勤用ID").Index
    ymCol = dst.ListColumns("年月").Index
    Dim oldRows As Long
    oldRows = dst.ListRows.Count
    For r = 1 To oldRows
        If CStr(dst.DataBodyRange.Cells(r, idCol).Value) <> "" Then
            If Not (CStr(dst.DataBodyRange.Cells(r, ymCol).Value) = ym And ids.Exists(CStr(dst.DataBodyRange.Cells(r, idCol).Value))) Then
                ReDim rowVals(1 To nCols)
                For c = 1 To nCols
                    If formulas(c) = "" Then rowVals(c) = dst.DataBodyRange.Cells(r, c).Value
                Next c
                keep.Add rowVals
            End If
        End If
    Next r
    ' 足す行
    Dim srcId As Long, srcYm As Long
    srcId = src.ListColumns("出勤用ID").Index
    srcYm = src.ListColumns("年月").Index
    If Not src.DataBodyRange Is Nothing Then
        For r = 1 To src.ListRows.Count
            If CStr(src.DataBodyRange.Cells(r, srcYm).Value) = ym And ids.Exists(CStr(src.DataBodyRange.Cells(r, srcId).Value)) Then
                ReDim rowVals(1 To nCols)
                For c = 1 To nCols
                    If srcCol(c) > 0 Then
                        rowVals(c) = src.DataBodyRange.Cells(r, srcCol(c)).Value
                    ElseIf dst.ListColumns(c).Name = "転記日時" Then
                        rowVals(c) = stamp
                    End If
                Next c
                keep.Add rowVals
            End If
        Next r
    End If
    ' 書き直す（表の大きさを合わせ、値の列だけ書き、式の列は全部の行に入れ直す）
    Dim n As Long, data() As Variant
    n = keep.Count
    If oldRows > 0 Then dst.DataBodyRange.ClearContents
    dst.Resize dst.Range.Resize(1 + IIf(n > 0, n, 1))
    If oldRows > IIf(n > 0, n, 1) Then dst.Range.Offset(1 + IIf(n > 0, n, 1)).Resize(oldRows - IIf(n > 0, n, 1)).ClearContents
    If n > 0 Then
        ReDim data(1 To n, 1 To 1)
        For c = 1 To nCols
            If formulas(c) = "" Then
                For r = 1 To n
                    data(r, 1) = keep(r)(c)
                Next r
                dst.ListColumns(c).DataBodyRange.Value = data
            End If
        Next c
    End If
    For c = 1 To nCols
        If formulas(c) <> "" Then dst.ListColumns(c).DataBodyRange.Formula = formulas(c)
    Next c
    ' 並べ替え：年月の新しい順、出勤用ID、日付
    If n > 1 Then
        With dst.Sort
            .SortFields.Clear
            .SortFields.Add Key:=dst.ListColumns(sortKeys(0)).DataBodyRange, Order:=xlDescending
            For k = 1 To UBound(sortKeys)
                .SortFields.Add Key:=dst.ListColumns(sortKeys(k)).DataBodyRange, Order:=xlAscending
            Next k
            .Header = xlYes
            .Apply
        End With
    End If
End Sub
