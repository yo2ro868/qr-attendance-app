Attribute VB_Name = "MatchingMacro"
Option Explicit
' 出勤・休暇管理データ の［最新を読み込む］（目次）・［申請を取り込む］ボタン：休暇申請シート・テレワーク予定表・
' 超過勤務等命令簿・フレックスの申告と割振りを読み直し、休暇等・出勤簿・申請一覧・テレワーク一覧・超過勤務の読み込み・フレックスを最新にする（開いたときには読み直さない）。
' フレックスは、所属長の承認済みになった単位期間（フレックス一覧の取り込みが「取り込む」）を「フレックスの日別」に写す（承認時の内容として残す）。
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
    Dim flexTaken As Long
    Application.StatusBar = "承認済みのフレックスを取り込んでいます..."
    flexTaken = TakeFlexSnapshot()
    ThisWorkbook.Worksheets("休暇申請シートの読み込み").Range("B5").Value = Now
    FitRows
    Application.StatusBar = False
    Application.Cursor = xlDefault

    Dim lo As ListObject, total As Long, waiting As Long, notes As Long
    Set lo = ThisWorkbook.Worksheets("申請一覧").ListObjects("申請一覧")
    If Not lo.DataBodyRange Is Nothing Then
        total = WorksheetFunction.CountIf(lo.ListColumns("休暇の種類").DataBodyRange, "?*")
        waiting = WorksheetFunction.CountIf(lo.ListColumns("状況").DataBodyRange, "承認待ち")
        notes = WorksheetFunction.CountIf(lo.ListColumns("確認").DataBodyRange, "?*")
    End If
    Dim tele As Long, otWaiting As Long, flexNotes As Long, flexDayNotes As Long
    flexNotes = CountColumn("フレックス", "フレックス一覧", "注意", "?*")
    flexDayNotes = CountColumn("フレックスの日別", "確定_フレックス日別", "注意", "?*")
    tele = CountColumn("テレワーク一覧", "テレワーク一覧", "注意", "?*")
    otWaiting = CountColumn("超過勤務の読み込み", "超勤_月", "状況", "確認待ち")
    MsgBox "最新の内容を読み込みました。" & vbCrLf & vbCrLf & _
           "休暇の申請 " & total & " 件（うち承認待ち " & waiting & " 件、注意のある行 " & notes & " 件）" & vbCrLf & _
           "テレワークの注意 " & tele & " 件" & vbCrLf & _
           "超過勤務等命令簿の所属長の確認待ち " & otWaiting & " 人・月" & vbCrLf & _
           "フレックス：今回取り込んだ単位期間 " & flexTaken & " 件、注意 " & flexNotes & " 件、日ごとの注意 " & flexDayNotes & " 日" & vbCrLf & vbCrLf & _
           "くわしくは目次シートの各行から確かめてください。", vbInformation
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

' フレックス一覧で取り込みが「取り込む」の単位期間（出勤用ID|期間|期間の開始日）を、フレックス_日別（今の内容）から
' 確定_フレックス日別に写す（同じ単位期間を前に写していれば置き換える）。写したら、写した内容を読む表を読み直す。写した単位期間の数を返す
Private Function TakeFlexSnapshot() As Long
    Dim list As ListObject, src As ListObject, dst As ListObject, keys As Object, r As Long
    On Error Resume Next
    Set list = ThisWorkbook.Worksheets("フレックス").ListObjects("フレックス一覧")
    Set src = ThisWorkbook.Worksheets("フレックスの読み込み").ListObjects("フレックス_日別")
    Set dst = ThisWorkbook.Worksheets("フレックスの日別").ListObjects("確定_フレックス日別")
    On Error GoTo 0
    If list Is Nothing Or src Is Nothing Or dst Is Nothing Then Exit Function
    If list.DataBodyRange Is Nothing Then Exit Function
    Set keys = CreateObject("Scripting.Dictionary")
    For r = 1 To list.ListRows.Count
        If CellText(list, r, "取り込み") = "取り込む" Then
            keys(CellText(list, r, "出勤用ID") & "|" & CellText(list, r, "期間") & "|" & DateKey(list.ListColumns("開始日").DataBodyRange.Cells(r).Value)) = True
        End If
    Next r
    If keys.Count = 0 Then Exit Function
    Application.ScreenUpdating = False
    ReplaceByKey dst, src, keys, Array("出勤用ID", "期間", "期間の開始日"), "確定日時", Now
    Application.ScreenUpdating = True
    ' 写した内容を読む表（一覧の取り込み・注意、出勤簿の備考）を読み直す
    list.QueryTable.Refresh False
    ThisWorkbook.Worksheets("休暇申請シートの読み込み").ListObjects("フレックス_出勤簿").QueryTable.Refresh False
    Application.Calculate
    TakeFlexSnapshot = keys.Count
End Function

Private Function DateKey(v As Variant) As String
    If IsDate(v) Then DateKey = Format(v, "yyyy-mm-dd") Else DateKey = CStr(v)
End Function

' 写す先の表 dst から、keys にあるキー（keyCols の列の値を | でつないだもの。日付は yyyy-mm-dd）の行を除き、
' 読み込んだ表 src のそのキーの行を足して書き直す。見出しが同じ名前の列を写す。stampCol の列には stamp、式の列は全部の行に入れ直す
Private Sub ReplaceByKey(dst As ListObject, src As ListObject, keys As Object, keyCols As Variant, stampCol As String, stamp As Date)
    Dim nCols As Long, c As Long, r As Long, k As Long, i As Long
    nCols = dst.ListColumns.Count
    Dim formulas() As String, srcCol() As Long, dKey() As Long, sKey() As Long
    ReDim formulas(1 To nCols)
    ReDim srcCol(1 To nCols)
    ReDim dKey(0 To UBound(keyCols))
    ReDim sKey(0 To UBound(keyCols))
    For c = 1 To nCols
        If Not dst.DataBodyRange Is Nothing Then
            If dst.ListColumns(c).DataBodyRange.Cells(1).HasFormula Then formulas(c) = dst.ListColumns(c).DataBodyRange.Cells(1).Formula
        End If
        For k = 1 To src.ListColumns.Count
            If src.ListColumns(k).Name = dst.ListColumns(c).Name Then srcCol(c) = k
        Next k
    Next c
    For i = 0 To UBound(keyCols)
        dKey(i) = dst.ListColumns(CStr(keyCols(i))).Index
        sKey(i) = src.ListColumns(CStr(keyCols(i))).Index
    Next i
    Dim keep As New Collection, rowVals() As Variant, dv As Variant, sv As Variant, oldRows As Long, kk As String
    oldRows = dst.ListRows.Count
    If Not dst.DataBodyRange Is Nothing Then
        dv = dst.DataBodyRange.Value
        For r = 1 To oldRows
            kk = RowKey(dv, r, dKey)
            If kk <> "" And Not keys.Exists(kk) Then
                ReDim rowVals(1 To nCols)
                For c = 1 To nCols
                    If formulas(c) = "" Then rowVals(c) = dv(r, c)
                Next c
                keep.Add rowVals
            End If
        Next r
    End If
    If Not src.DataBodyRange Is Nothing Then
        sv = src.DataBodyRange.Value
        For r = 1 To src.ListRows.Count
            If keys.Exists(RowKey(sv, r, sKey)) Then
                ReDim rowVals(1 To nCols)
                For c = 1 To nCols
                    If srcCol(c) > 0 Then
                        rowVals(c) = sv(r, srcCol(c))
                    ElseIf dst.ListColumns(c).Name = stampCol Then
                        rowVals(c) = stamp
                    End If
                Next c
                keep.Add rowVals
            End If
        Next r
    End If
    Dim n As Long, data() As Variant
    n = keep.Count
    If oldRows > 0 And Not dst.DataBodyRange Is Nothing Then dst.DataBodyRange.ClearContents
    dst.Resize dst.Range.Resize(1 + IIf(n > 0, n, 1))
    If oldRows > IIf(n > 0, n, 1) Then dst.Range.Offset(1 + IIf(n > 0, n, 1)).Resize(oldRows - IIf(n > 0, n, 1)).ClearContents
    If n > 0 Then
        For c = 1 To nCols
            If formulas(c) = "" Then
                ReDim data(1 To n, 1 To 1)
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
    If n > 1 Then
        With dst.Sort
            .SortFields.Clear
            .SortFields.Add Key:=dst.ListColumns("出勤用ID").DataBodyRange, Order:=xlAscending
            .SortFields.Add Key:=dst.ListColumns("日付").DataBodyRange, Order:=xlAscending
            .Header = xlYes
            .Apply
        End With
    End If
End Sub

' 2次元配列の r 行目の、cols の列の値を | でつないだキー（日付は yyyy-mm-dd、出勤用ID が空なら ""）
Private Function RowKey(v As Variant, r As Long, cols() As Long) As String
    Dim i As Long, s As String
    If CStr(v(r, cols(0))) = "" Then Exit Function
    For i = 0 To UBound(cols)
        If i > 0 Then s = s & "|"
        s = s & DateKey(v(r, cols(i)))
    Next i
    RowKey = s
End Function

' ［利用予定を書き出す］（利用予定シート）：利用予定シートの表を、今日の日付の名前のシート（例 11月16日）に値で書き出す（配る用）。
' 部・課は同じものが続く行を結合し、職員のいない行と、印刷範囲の外の列（隠し列・対象月の欄）とボタンは消す。前に書き出したシートは残す
Public Sub MakeScheduleSheet()
    Const FIRST_ROW As Long = 6, LAST_COL As String = "AH"
    Dim src As Worksheet, ws As Worksheet, nm As String, n As Long, last As Long, r As Long, e As Long, c As Long
    Set src = ThisWorkbook.Worksheets("利用予定")
    If Not IsDate(src.Range("AN1").Value) Then
        MsgBox "対象月を AN1 に入れてください（例 2026/11）。", vbExclamation
        Exit Sub
    End If
    Application.Calculate
    last = FIRST_ROW - 1 + WorksheetFunction.CountIf(src.Range("AK" & FIRST_ROW & ":AK" & (FIRST_ROW + 499)), "?*")
    If last < FIRST_ROW Then
        MsgBox "職員名簿に部が入っている人がいません（職員名簿の G 列に部、H 列に課を入れてください）。", vbExclamation
        Exit Sub
    End If
    nm = Month(Date) & "月" & Day(Date) & "日"
    n = 1
    Do While Not SheetOf(nm & IIf(n > 1, "（" & n & "）", "")) Is Nothing
        n = n + 1
    Loop
    If n > 1 Then nm = nm & "（" & n & "）"

    On Error GoTo Failed
    Application.ScreenUpdating = False
    src.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Set ws = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    ws.Name = nm
    ' 値にする
    With ws.Range("A1:" & LAST_COL & (FIRST_ROW + 499))
        .Copy
        .PasteSpecial Paste:=xlPasteValues
    End With
    Application.CutCopyMode = False
    ' 土日・祝日の列の色（条件付き書式は消すので、色を付けておく）
    For c = 4 To 34
        If IsDate(ws.Cells(4, c).Value) Then
            If Weekday(ws.Cells(4, c).Value) = 1 Or Weekday(ws.Cells(4, c).Value) = 7 Or _
               WorksheetFunction.CountIf(ThisWorkbook.Worksheets("祝日").Range("A2:A200"), ws.Cells(4, c).Value) > 0 Then
                ws.Range(ws.Cells(4, c), ws.Cells(last, c)).Interior.Color = RGB(231, 230, 230)
            End If
        End If
    Next c
    ws.Cells.FormatConditions.Delete
    ' 職員のいない行、隠し列と対象月の欄、ボタンを消す
    If last < FIRST_ROW + 499 Then ws.Rows((last + 1) & ":" & (FIRST_ROW + 499)).Delete
    ws.Range("AI:AN").Delete
    Do While ws.Shapes.Count > 0
        ws.Shapes(1).Delete
    Loop
    ' 部・課：同じものが続く行を結合
    Application.DisplayAlerts = False
    r = FIRST_ROW
    Do While r <= last
        e = r
        Do While e + 1 <= last And CStr(ws.Cells(e + 1, 1).Value) = CStr(ws.Cells(r, 1).Value)
            e = e + 1
        Loop
        MergeRun ws, 2, r, e
        If e > r Then ws.Range(ws.Cells(r, 1), ws.Cells(e, 1)).Merge
        r = e + 1
    Loop
    Application.DisplayAlerts = True
    ws.PageSetup.PrintArea = "$A$1:$" & LAST_COL & "$" & last
    ws.Activate
    ws.Range("A1").Select
    Application.ScreenUpdating = True
    MsgBox "「" & nm & "」シートに書き出しました（" & (last - FIRST_ROW + 1) & " 人）。", vbInformation
    Exit Sub
Failed:
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "書き出せませんでした：" & Err.Description, vbExclamation
End Sub

' r1〜r2 行の col 列で、同じ値が続く行を結合する（部の中の課）
Private Sub MergeRun(ws As Worksheet, col As Long, r1 As Long, r2 As Long)
    Dim r As Long, e As Long
    r = r1
    Do While r <= r2
        e = r
        Do While e + 1 <= r2 And CStr(ws.Cells(e + 1, col).Value) = CStr(ws.Cells(r, col).Value)
            e = e + 1
        Loop
        If e > r Then ws.Range(ws.Cells(r, col), ws.Cells(e, col)).Merge
        r = e + 1
    Loop
End Sub

Private Function SheetOf(nm As String) As Worksheet
    Dim s As Worksheet
    For Each s In ThisWorkbook.Worksheets
        If s.Name = nm Then Set SheetOf = s: Exit Function
    Next s
End Function

' 一覧のシートの表を見やすくする：折り返しの列（注意・業務内容など。列全体が折り返し）は幅をそのままにし、ほかの列は中身に合わせて広げ
' （狭くはしない。広すぎるときは60まで）、行の高さを折り返した文が見える高さに合わせる
Private Sub FitRows()
    Dim nm As Variant, ws As Worksheet, lo As ListObject, lc As ListColumn, w As Double
    On Error Resume Next
    For Each nm In Array("申請一覧", "テレワーク一覧", "テレワーク報告", "フレックス", "フレックスの日別", "フレックスの読み込み", _
                         "超過勤務", "超過勤務の日別", "超過勤務の読み込み")
        Set ws = Nothing
        Set ws = ThisWorkbook.Worksheets(CStr(nm))
        If Not ws Is Nothing Then
            For Each lo In ws.ListObjects
                For Each lc In lo.ListColumns
                    If Not (lc.Range.WrapText = True) Then
                        w = lc.Range.ColumnWidth
                        lc.Range.Columns.AutoFit
                        If lc.Range.ColumnWidth < w Then lc.Range.ColumnWidth = w
                        If lc.Range.ColumnWidth > 60 Then lc.Range.ColumnWidth = 60
                    End If
                Next lc
            Next lo
            ws.UsedRange.Rows.AutoFit
        End If
    Next nm
End Sub

' シートの表の列で、条件に合うセルの数（表が空・見つからなければ 0）
Private Function CountColumn(sheetName As String, tableName As String, col As String, crit As String) As Long
    On Error Resume Next
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    CountColumn = WorksheetFunction.CountIf(lo.ListColumns(col).DataBodyRange, crit)
End Function

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
