Attribute VB_Name = "MatchingMacro"
Option Explicit
' 出勤・休暇管理データ の［最新を読み込む］（目次）・［申請を取り込む］（申請一覧）ボタン：休暇申請シート・テレワーク予定表・
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
    Dim qrNote As String
    qrNote = FillQrData()           ' フォルダの QR出勤のファイルを「出力データ貼付」に（フォルダが空欄なら貼り付けた内容のまま）
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
           "フレックス：今回取り込んだ単位期間 " & flexTaken & " 件、注意 " & flexNotes & " 件、日ごとの注意 " & flexDayNotes & " 日" & vbCrLf & _
           IIf(qrNote <> "", qrNote & vbCrLf, "") & vbCrLf & _
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

' ［利用予定を書き出す］（利用予定シート）：利用予定シートの表を、AN9 に指定した別のファイル（なければ作る）の、
' 今日の日付の名前のシート（例 11月16日）に値で書き出して保存する（配る用。出勤・休暇管理データの中には作らない）。
' 部・課は同じものが続く行を結合し、職員のいない行と、印刷範囲の外の列（隠し列・対象月の欄）とボタンは消す。前に書き出したシートは残す
Public Sub MakeScheduleSheet()
    Const FIRST_ROW As Long = 6, LAST_COL As String = "AH"
    Dim src As Worksheet, ws As Worksheet, nm As String, n As Long, last As Long, r As Long, e As Long, c As Long
    Dim dest As Workbook, destPath As String, created As Boolean, blank As Worksheet
    Set src = ThisWorkbook.Worksheets("利用予定")
    If Not IsDate(src.Range("AN1").Value) Then
        MsgBox "対象月を AN1 に入れてください（例 2026/11）。", vbExclamation
        Exit Sub
    End If
    destPath = Trim(CStr(src.Range("AN9").Value))
    If destPath = "" Then
        MsgBox "書き出す先のファイルを AN9 に入れてください（例 C:\共有\フレックス・テレワーク利用予定.xlsx）。", vbExclamation
        Exit Sub
    End If
    If LCase(Right(destPath, 5)) <> ".xlsx" Then destPath = destPath & ".xlsx"
    If InStrRev(destPath, "\") = 0 Then
        MsgBox "AN9 には、フォルダも含めたファイルの場所を入れてください（例 C:\共有\フレックス・テレワーク利用予定.xlsx）。", vbExclamation
        Exit Sub
    End If
    If Dir(Left(destPath, InStrRev(destPath, "\") - 1), vbDirectory) = "" Then
        MsgBox "AN9 のフォルダがありません：" & vbCrLf & Left(destPath, InStrRev(destPath, "\") - 1), vbExclamation
        Exit Sub
    End If
    Application.Calculate
    last = FIRST_ROW - 1 + WorksheetFunction.CountIf(src.Range("AK" & FIRST_ROW & ":AK" & (FIRST_ROW + 499)), "?*")
    If last < FIRST_ROW Then
        MsgBox "職員名簿に部が入っている人がいません（職員名簿の G 列に部、H 列に課を入れてください）。", vbExclamation
        Exit Sub
    End If
    On Error GoTo Failed
    Application.ScreenUpdating = False
    ' 書き出す先のファイル（開いていればそれを、なければ開く。ファイルがなければ作る）
    On Error Resume Next
    Set dest = Workbooks(Mid(destPath, InStrRev(destPath, "\") + 1))
    On Error GoTo Failed
    If dest Is Nothing Then
        If Dir(destPath) <> "" Then
            Set dest = Workbooks.Open(Filename:=destPath, UpdateLinks:=0)
            If dest.ReadOnly Then
                dest.Close SaveChanges:=False
                Application.ScreenUpdating = True
                MsgBox "書き出す先のファイルをほかの人が開いているため、書き込めません：" & vbCrLf & destPath, vbExclamation
                Exit Sub
            End If
        Else
            Set dest = Workbooks.Add(xlWBATWorksheet)
            Set blank = dest.Worksheets(1)
            created = True
        End If
    End If
    nm = Month(Date) & "月" & Day(Date) & "日"
    n = 1
    Do While SheetExists(dest, nm & IIf(n > 1, "（" & n & "）", ""))
        n = n + 1
    Loop
    If n > 1 Then nm = nm & "（" & n & "）"
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
    ' 書き出す先のファイルに移して保存する
    ws.Move After:=dest.Worksheets(dest.Worksheets.Count)
    Set ws = dest.Worksheets(dest.Worksheets.Count)
    If created Then
        Application.DisplayAlerts = False
        blank.Delete
        Application.DisplayAlerts = True
        dest.SaveAs Filename:=destPath, FileFormat:=xlOpenXMLWorkbook
    Else
        dest.Save
    End If
    dest.Activate
    ws.Activate
    ws.Range("A1").Select
    Application.ScreenUpdating = True
    MsgBox destPath & " の「" & nm & "」シートに書き出して保存しました（" & (last - FIRST_ROW + 1) & " 人）。", vbInformation
    Exit Sub
Failed:
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "書き出せませんでした：" & Err.Description, vbExclamation
End Sub

' ブックに、その名前のシートがあるか
Private Function SheetExists(wb As Workbook, nm As String) As Boolean
    Dim ws As Object
    For Each ws In wb.Sheets
        If ws.Name = nm Then SheetExists = True: Exit Function
    Next ws
End Function

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

' ［名簿の人の行を足す］（休暇等シート・目次）：年を聞き、職員名簿の全員について、休暇等シートにその年の行（出勤用ID・年）が
' まだない人の分だけ、いちばん下の行の次から足す（ある人・例の行の人は足さない。手で入れた数字の行は動かさない）
Public Sub AddLeaveRows()
    Const LAST_ROW As Long = 1001, NAME_LAST As Long = 501
    Dim yr As String, ws As Worksheet, nm As Worksheet, have As Object, r As Long, last As Long, nextRow As Long
    Dim id As String, added As Long, full As Boolean
    yr = InputBox("職員名簿の全員について、休暇等シートにこの年の行を足します（すでにある人は足しません）。" & vbCrLf & vbCrLf & _
                  "年（西暦）：", "名簿の人の行を足す", CStr(Year(Date)))
    If yr = "" Then Exit Sub
    If Not IsNumeric(yr) Then MsgBox "年を数字で入れてください（例 2027）。", vbExclamation: Exit Sub
    If Val(yr) < 2000 Or Val(yr) > 2100 Then MsgBox "年を数字で入れてください（例 2027）。", vbExclamation: Exit Sub
    Set ws = ThisWorkbook.Worksheets("休暇等")
    Set nm = ThisWorkbook.Worksheets("職員名簿")
    Set have = CreateObject("Scripting.Dictionary")
    last = 1
    For r = 2 To LAST_ROW
        id = Trim$(CStr(ws.Cells(r, 1).Value))
        If id <> "" Then
            have(IdKey(id) & "|" & Trim$(CStr(ws.Cells(r, 3).Value))) = True
            last = r
        End If
    Next r
    nextRow = last + 1
    For r = 2 To NAME_LAST
        id = Trim$(CStr(nm.Cells(r, 1).Value))
        If id <> "" And InStr(CStr(nm.Cells(r, 2).Value), "（例）") = 0 Then
            If Not have.Exists(IdKey(id) & "|" & CStr(Val(yr))) Then
                If nextRow > LAST_ROW Then full = True: Exit For
                ws.Cells(nextRow, 1).NumberFormat = "@"
                ws.Cells(nextRow, 1).Value = id
                ws.Cells(nextRow, 3).Value = CLng(Val(yr))
                have(IdKey(id) & "|" & CStr(Val(yr))) = True
                nextRow = nextRow + 1
                added = added + 1
            End If
        End If
    Next r
    ws.Activate
    If added > 0 Then ws.Cells(nextRow - added, 1).Select
    MsgBox yr & " 年の行を " & added & " 人分足しました。" & _
           IIf(full, vbCrLf & "休暇等シートの行がいっぱいです（" & LAST_ROW & " 行目まで）。古い年の行を整理してください。", "") & _
           IIf(added > 0, vbCrLf & vbCrLf & "氏名・採用日・退職日は職員名簿から、繰越分は前の年の行から自動で入ります。使い始めの年の繰越分や、本年分を変えるときだけ入れてください。", ""), vbInformation
End Sub

' 出勤用ID の照合キー（数だけの ID は先頭の0を取る。出勤・休暇管理データ の照合キーと同じ）
Private Function IdKey(ByVal s As String) As String
    s = Trim$(s)
    If IsNumeric(s) Then IdKey = CStr(CDbl(s)) Else IdKey = s
End Function

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

' ===== テレワーク実施報告書（テレワーク報告シートの［実施報告書を見る］［実施報告書をPDFで保存］） =====
' 報告は、その職員の休暇申請シートの「テレワーク実施報告書」シート（別紙３の様式）で見る。休暇申請シートは読み取り専用で開き、
' 印刷する番号（M2）に報告の番号を入れる（保存しないので、休暇申請シートは変わらない）。
Private Const TW_PRINT As String = "テレワーク実施報告書"
Private Const TW_PDF_DIR As String = "テレワーク実施報告書"     ' PDF を置くフォルダ（休暇申請シートのフォルダと同じ場所に作る）

' ［実施報告書を見る］：選んでいる行の報告を、実施報告書の様式で表示する（印刷・PDF の保存は、表示した画面からもできる）
Public Sub ShowTwReport()
    Dim lo As ListObject, rows As Collection, wb As Workbook, opened As Boolean
    Set rows = TwSelectedRows(lo)
    If rows Is Nothing Then Exit Sub
    Set wb = OpenLeaveSheet(CellText(lo, CLng(rows(1)), "ファイル名"), opened)
    If wb Is Nothing Then Exit Sub
    If Not SetTwNumber(wb, lo, CLng(rows(1))) Then Exit Sub
    wb.Activate
    wb.Worksheets(TW_PRINT).Activate
    If rows.Count > 1 Then MsgBox "選んだ行のうち、いちばん上の報告を表示しました（何件かまとめて保存するときは［実施報告書をPDFで保存］）。", vbInformation
End Sub

' ［実施報告書をPDFで保存］：選んでいる行の報告を、1件1ファイルの PDF（実施報告書の様式）にする
Public Sub SaveTwReports()
    Dim lo As ListObject, rows As Collection, item As Variant, r As Long, wb As Workbook, opened As Boolean
    Dim outDir As String, fname As String, n As Long, mine As New Collection, skipped As String
    Set rows = TwSelectedRows(lo)
    If rows Is Nothing Then Exit Sub
    outDir = TwPdfDir()
    If outDir = "" Then Exit Sub
    Application.ScreenUpdating = False
    For Each item In rows
        r = CLng(item)
        Set wb = OpenLeaveSheet(CellText(lo, r, "ファイル名"), opened, True)
        If wb Is Nothing Then
            skipped = skipped & vbCrLf & "・" & CellText(lo, r, "ファイル名")
        Else
            If opened Then mine.Add wb
            If SetTwNumber(wb, lo, r) Then
                fname = "テレワーク実施報告書　" & CellText(lo, r, "出勤用ID") & "　" & CellText(lo, r, "氏名") & "　" & _
                        Format(lo.ListColumns("実施日").DataBodyRange.Cells(r).Value, "yyyy-mm-dd") & "（" & CellText(lo, r, "番号") & "）.pdf"
                wb.Worksheets(TW_PRINT).ExportAsFixedFormat Type:=xlTypePDF, Filename:=outDir & "\" & fname, _
                    Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=False
                n = n + 1
            End If
        End If
    Next item
    For Each item In mine
        item.Close SaveChanges:=False
    Next item
    ThisWorkbook.Activate
    Application.ScreenUpdating = True
    MsgBox n & " 件の実施報告書を PDF で保存しました。" & vbCrLf & outDir & _
           IIf(skipped <> "", vbCrLf & vbCrLf & "開けなかった休暇申請シート：" & skipped, ""), vbInformation
End Sub

' テレワーク報告シートで選んでいる、表「テレワーク報告一覧」の行（表の中の何行目か）。なければ説明を出して Nothing
Private Function TwSelectedRows(lo As ListObject) As Collection
    Dim ws As Worksheet, area As Range, cell As Range, seen As Object, r As Long, result As New Collection, hit As Range
    Set ws = ThisWorkbook.Worksheets("テレワーク報告")
    Set lo = ws.ListObjects("テレワーク報告一覧")
    If lo.DataBodyRange Is Nothing Then MsgBox "報告がありません（目次の［最新を読み込む］を押してください）。", vbExclamation: Exit Function
    If Not ActiveSheet Is ws Then MsgBox "テレワーク報告シートで、報告の行を選んでから押してください。", vbExclamation: Exit Function
    Set hit = Intersect(Selection, lo.DataBodyRange)
    If hit Is Nothing Then MsgBox "報告の行（表の中のセル）を選んでから押してください。何件か選ぶときは行をまとめて選びます。", vbExclamation: Exit Function
    Set seen = CreateObject("Scripting.Dictionary")
    For Each area In hit.Areas
        For Each cell In area.Columns(1).Cells
            r = cell.Row - lo.HeaderRowRange.Row
            If Not seen.Exists(r) Then seen.Add r, True: result.Add r
        Next cell
    Next area
    Set TwSelectedRows = result
End Function

' 休暇申請シートを読み取り専用で開く（同じ名前のブックを開いていればそれを使う）。opened は、このマクロが開いたら True。
' quiet が True のときは、見つからなくてもメッセージを出さない（まとめて保存するとき）
Private Function OpenLeaveSheet(ByVal fileName As String, opened As Boolean, Optional ByVal quiet As Boolean = False) As Workbook
    Dim folder As String, path As String
    opened = False
    On Error Resume Next
    Set OpenLeaveSheet = Workbooks(fileName)
    On Error GoTo 0
    If Not OpenLeaveSheet Is Nothing Then Exit Function
    folder = LeaveFolder()
    path = folder & "\" & fileName
    If folder = "" Or fileName = "" Or Dir(path) = "" Then
        If Not quiet Then MsgBox "休暇申請シートが見つかりません：" & vbCrLf & path & vbCrLf & "（休暇申請シートの読み込みシートの B2 のフォルダを確かめてください）", vbExclamation
        Exit Function
    End If
    Set OpenLeaveSheet = Workbooks.Open(Filename:=path, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)
    opened = True
End Function

' 休暇申請シートのフォルダ（休暇申請シートの読み込みシートの B2。最後の \ は除く）
Private Function LeaveFolder() As String
    LeaveFolder = Trim(CStr(ThisWorkbook.Worksheets("休暇申請シートの読み込み").Range("B2").Value))
    If Right(LeaveFolder, 1) = "\" Then LeaveFolder = Left(LeaveFolder, Len(LeaveFolder) - 1)
End Function

' 実施報告書のシートの印刷する番号（M2）に、表の r 行目の報告の番号を入れる
Private Function SetTwNumber(wb As Workbook, lo As ListObject, r As Long) As Boolean
    Dim ps As Worksheet
    On Error Resume Next
    Set ps = wb.Worksheets(TW_PRINT)
    On Error GoTo 0
    If ps Is Nothing Then
        MsgBox wb.Name & " に「" & TW_PRINT & "」がありません（様式が古いので、［様式を移行］をしてください）。", vbExclamation
        Exit Function
    End If
    ps.Range("M2").Value = CLng(lo.ListColumns("番号").DataBodyRange.Cells(r).Value)
    ps.Calculate
    SetTwNumber = True
End Function

' PDF を置くフォルダ（休暇申請シートのフォルダと同じ場所の「テレワーク実施報告書」。なければ作る）
Private Function TwPdfDir() As String
    Dim folder As String
    folder = LeaveFolder()
    If folder = "" Or InStrRev(folder, "\") = 0 Then
        MsgBox "休暇申請シートの読み込みシートの B2 に、休暇申請シートのフォルダを入れてください。", vbExclamation
        Exit Function
    End If
    TwPdfDir = Left(folder, InStrRev(folder, "\") - 1) & "\" & TW_PDF_DIR
    If Dir(TwPdfDir, vbDirectory) = "" Then MkDir TwPdfDir
End Function

' ===== ［命令簿をPDFで保存］（命令簿の保存シート・目次）：所属長の確認済みの超過勤務等命令簿を、1人1か月1ファイルの PDF で保管する =====
' 超過勤務の読み込みシートの表（超勤_月）で状況が「確認済み」の月を、休暇申請シートを読み取り専用で開いて、その月のシート（超勤 R8.11 など）を PDF にする。
' 保存した記録（命令簿の保存シートの6行目から）に、シートの内容から作った文字（内容の記録）を残す。同じ内容なら保存しない。
' 保存したあとで直された月は、前の PDF を残したまま「…（2）.pdf」「…（3）.pdf」の名前で保存する。
Private Const OT_LOG As String = "命令簿の保存"
Private Const OT_LOG_FIRST As Long = 6         ' 記録の最初の行（5行目は見出し）
Private Const OT_PDF_DIR As String = "超過勤務等命令簿"     ' 保存先が空欄のとき：休暇申請シートのフォルダと同じ場所に作るフォルダ

Public Sub SaveOvertimePdfs()
    Dim lo As ListObject, logWs As Worksheet, outDir As String
    Set lo = ThisWorkbook.Worksheets("超過勤務の読み込み").ListObjects("超勤_月")
    If lo.DataBodyRange Is Nothing Then MsgBox "命令簿がありません（目次の［最新を読み込む］を押してください）。", vbExclamation: Exit Sub
    Set logWs = ThisWorkbook.Worksheets(OT_LOG)
    outDir = OtPdfDir()
    If outDir = "" Then Exit Sub

    ' これまでの記録：出勤用ID|年月 → いちばん新しい内容の記録、保存した回数
    Dim lastHash As Object, times As Object, r As Long, last As Long, k As String
    Set lastHash = CreateObject("Scripting.Dictionary")
    Set times = CreateObject("Scripting.Dictionary")
    last = logWs.Cells(logWs.Rows.Count, 2).End(xlUp).Row
    For r = OT_LOG_FIRST To last
        k = IdKey(CStr(logWs.Cells(r, 3).Value)) & "|" & CStr(logWs.Cells(r, 2).Value)
        If k <> "|" Then
            lastHash(k) = CStr(logWs.Cells(r, 7).Value)
            times(k) = times(k) + 1
        End If
    Next r
    If last < OT_LOG_FIRST - 1 Then last = OT_LOG_FIRST - 1

    Dim i As Long, ym As String, id As String, nm As String, fileName As String, sheetName As String
    Dim books As Object, wb As Workbook, opened As Boolean, ws As Worksheet, h As String, n As Long, fname As String
    Dim saved As Long, revised As Long, same As Long, skipped As String, item As Variant, pair As Variant
    Set books = CreateObject("Scripting.Dictionary")
    Application.ScreenUpdating = False
    On Error GoTo Failed
    For i = 1 To lo.ListRows.Count
        If CellText(lo, i, "状況") = "確認済み" Then
            ym = CellText(lo, i, "年月")
            id = CellText(lo, i, "出勤用ID")
            nm = CellText(lo, i, "氏名")
            fileName = CellText(lo, i, "ファイル名")
            sheetName = OtSheetName(ym)
            ' 休暇申請シートは1回だけ開く（開けなかったものは Nothing を覚えておく）
            If Not books.Exists(fileName) Then
                Set wb = OpenLeaveSheet(fileName, opened, True)
                If wb Is Nothing Then
                    skipped = skipped & vbCrLf & "・" & fileName & "（開けませんでした）"
                    books.Add fileName, Nothing
                Else
                    books.Add fileName, Array(wb, opened)
                End If
            End If
            Set ws = Nothing
            If IsObject(books(fileName)) Then
                ' Nothing（開けなかった）
            Else
                pair = books(fileName)
                Set wb = pair(0)
                On Error Resume Next
                Set ws = wb.Worksheets(sheetName)
                On Error GoTo Failed
                If ws Is Nothing Then skipped = skipped & vbCrLf & "・" & fileName & "：「" & sheetName & "」がありません"
            End If
            If Not ws Is Nothing Then
                ' 確認日が今も入っているか（読み込んだあとで消されたら保存しない）
                If Not IsNumeric(ws.Range("AC69").Value2) Or CStr(ws.Range("AC69").Value2) = "" Then
                    skipped = skipped & vbCrLf & "・" & nm & "（" & sheetName & "）：確認日がありません（読み込んだあとで消されました）"
                Else
                    k = IdKey(id) & "|" & ym
                    h = SheetHash(ws)
                    If lastHash.Exists(k) And CStr(lastHash(k)) = h Then
                        same = same + 1
                    Else
                        n = 0
                        If times.Exists(k) Then n = times(k)
                        fname = "超過勤務等命令簿　" & Mid$(sheetName, 4) & "　" & id & "　" & nm & IIf(n = 0, "", "（" & (n + 1) & "）") & ".pdf"
                        ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=outDir & "\" & fname, _
                            Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=False
                        last = last + 1
                        logWs.Cells(last, 1).Value = Now
                        logWs.Cells(last, 2).NumberFormat = "@"
                        logWs.Cells(last, 2).Value = ym
                        logWs.Cells(last, 3).NumberFormat = "@"
                        logWs.Cells(last, 3).Value = id
                        logWs.Cells(last, 4).Value = nm
                        logWs.Cells(last, 5).Value = ws.Range("AC69").Value
                        logWs.Cells(last, 6).Value = fname
                        logWs.Cells(last, 7).NumberFormat = "@"
                        logWs.Cells(last, 7).Value = h
                        lastHash(k) = h
                        times(k) = n + 1
                        If n = 0 Then saved = saved + 1 Else revised = revised + 1
                    End If
                End If
            End If
        End If
    Next i
Done:
    On Error Resume Next
    For Each item In books.Items
        If Not IsObject(item) Then
            If item(1) Then item(0).Close SaveChanges:=False
        End If
    Next item
    ThisWorkbook.Activate
    Application.ScreenUpdating = True
    MsgBox "確認済みの命令簿を PDF で保存しました。" & vbCrLf & _
           "　新しく保存：" & saved & " 件" & vbCrLf & _
           "　直された月を別の名前で保存：" & revised & " 件" & vbCrLf & _
           "　前に保存した内容と同じなので保存しなかった：" & same & " 件" & vbCrLf & outDir & _
           IIf(skipped <> "", vbCrLf & vbCrLf & "保存できなかったもの：" & skipped, ""), vbInformation
    Exit Sub
Failed:
    MsgBox "PDF の保存の途中で止まりました：" & Err.Description & vbCrLf & "（ここまでに保存した分は、命令簿の保存シートに記録してあります）", vbExclamation
    Resume Done
End Sub

' 2026-11 → 超勤 R8.11（休暇申請シートの月のシートの名前）
Private Function OtSheetName(ByVal ym As String) As String
    OtSheetName = "超勤 R" & (CLng(Left$(ym, 4)) - 2018) & "." & CLng(Mid$(ym, 6, 2))
End Function

' 保存先：命令簿の保存シートの B2。空欄なら、休暇申請シートのフォルダと同じ場所の「超過勤務等命令簿」（なければ作る）
Private Function OtPdfDir() As String
    Dim folder As String
    folder = Trim$(CStr(ThisWorkbook.Worksheets(OT_LOG).Range("B2").Value))
    If Right$(folder, 1) = "\" Then folder = Left$(folder, Len(folder) - 1)
    If folder = "" Then
        folder = LeaveFolder()
        If folder = "" Or InStrRev(folder, "\") = 0 Then
            MsgBox "命令簿の保存シートの B2 に保存先のフォルダを入れてください（空欄なら、休暇申請シートの読み込みシートの B2 のフォルダの隣に作ります）。", vbExclamation
            Exit Function
        End If
        folder = Left$(folder, InStrRev(folder, "\") - 1) & "\" & OT_PDF_DIR
    End If
    If Dir(folder, vbDirectory) = "" Then
        On Error Resume Next
        MkDir folder
        On Error GoTo 0
        If Dir(folder, vbDirectory) = "" Then
            MsgBox "保存先のフォルダがなく、作れませんでした：" & vbCrLf & folder, vbExclamation
            Exit Function
        End If
    End If
    OtPdfDir = folder
End Function

' 月のシートの内容（A1:AF72 の表示している値）から作る短い文字。内容が変わると変わる（直されたかどうかを見るため）
Private Function SheetHash(ws As Worksheet) As String
    Dim v As Variant, r As Long, c As Long, s As String, j As Long, ch As Long
    Dim h1 As Double, h2 As Double
    Const P As Double = 2147483647#
    v = ws.Range("A1:AF72").Value2
    For r = 1 To UBound(v, 1)
        For c = 1 To UBound(v, 2)
            If IsError(v(r, c)) Then s = "#" Else s = CStr(v(r, c))
            s = s & Chr$(31)
            For j = 1 To Len(s)
                ch = AscW(Mid$(s, j, 1)) And &HFFFF&
                h1 = h1 * 31# + ch: h1 = h1 - Int(h1 / P) * P
                h2 = h2 * 131# + ch: h2 = h2 - Int(h2 / P) * P
            Next j
        Next c
    Next r
    SheetHash = Right$("0000000" & Hex$(CLng(h1)), 8) & Right$("0000000" & Hex$(CLng(h2)), 8)
End Function

' ===== ［申告・割振り簿を見る］（フレックスシート）：選んでいる行の単位期間を、その職員の休暇申請シートの「申告・割振り簿」（総務課の様式）で表示する =====
' 休暇申請シートは読み取り専用で開き、管理者だけが見るシート（ふだんは非表示）を表示して、単位期間（F5・H5・I5 の年・月・日）にその開始日を入れる。保存はしない
Private Const FLEX_BOOK As String = "申告・割振り簿"

Public Sub ShowFlexPeriod()
    Dim ws As Worksheet, lo As ListObject, hit As Range, r As Long, wb As Workbook, opened As Boolean, bk As Worksheet, d As Variant
    Set ws = ThisWorkbook.Worksheets("フレックス")
    Set lo = ws.ListObjects("フレックス一覧")
    If lo.DataBodyRange Is Nothing Then MsgBox "単位期間がありません（目次の［最新を読み込む］を押してください）。", vbExclamation: Exit Sub
    If Not ActiveSheet Is ws Then MsgBox "フレックスシートで、単位期間の行を選んでから押してください。", vbExclamation: Exit Sub
    Set hit = Intersect(ActiveCell, lo.DataBodyRange)
    If hit Is Nothing Then MsgBox "単位期間の行（表の中のセル）を選んでから押してください。", vbExclamation: Exit Sub
    r = hit.Row - lo.HeaderRowRange.Row
    Set wb = OpenLeaveSheet(CellText(lo, r, "ファイル名"), opened)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    Set bk = wb.Worksheets(FLEX_BOOK)
    On Error GoTo 0
    If bk Is Nothing Then
        MsgBox wb.Name & " に「" & FLEX_BOOK & "」がありません（様式が古いので、［様式を移行］をしてください）。", vbExclamation
        Exit Sub
    End If
    bk.Visible = xlSheetVisible
    d = lo.ListColumns("開始日").DataBodyRange.Cells(r).Value
    If IsDate(d) Then
        bk.Range("F5").Value = Year(d)
        bk.Range("H5").Value = Month(d)
        bk.Range("I5").Value = Day(d)
    End If
    bk.Calculate
    wb.Activate
    bk.Activate
    MsgBox wb.Name & " を読み取り専用で開き、「" & FLEX_BOOK & "」の単位期間 " & CellText(lo, r, "期間") & "（" & Format(d, "yyyy/m/d") & "〜）を表示しました。" & vbCrLf & _
           "印刷や PDF の保存はこの画面からできます。見終わったら、保存しないで閉じてください。", vbInformation
End Sub

' ===== QR出勤のフォルダ（出力データ貼付シートの H3） =====
' フォルダが入っていれば、「QR出勤の読み込み」シートの表 QR出勤（フォルダの attendance.xlsx をつなぎ、同じ記録を1つにしたもの）を、
' 「出力データ貼付」の A〜E 列（2行目から）に文字で入れ直す。入りきらないとき（5000件を超える）は新しい分だけ。空欄なら何もしない（貼り付けた内容を使う）
Private Function FillQrData() As String
    Const CAP As Long = 5000
    Dim dst As Worksheet, lo As ListObject, n As Long, first As Long, v As Variant, out() As Variant, i As Long, j As Long, files As Long, bad As Long
    Set dst = ThisWorkbook.Worksheets("出力データ貼付")
    If Trim(CStr(dst.Range("H3").Value)) = "" Then Exit Function
    files = CountColumn("QR出勤の読み込み", "QR出勤ファイル一覧", "読み取り", "OK")
    bad = CountColumn("QR出勤の読み込み", "QR出勤ファイル一覧", "読み取り", "<>OK") - CountColumn("QR出勤の読み込み", "QR出勤ファイル一覧", "読み取り", "")
    Set lo = ThisWorkbook.Worksheets("QR出勤の読み込み").ListObjects("QR出勤")
    If lo.DataBodyRange Is Nothing Then
        FillQrData = "QR出勤：フォルダに出勤記録のファイルがありません（「出力データ貼付」は前の内容のままです）" & _
                     IIf(bad > 0, "。読めなかったファイル " & bad & " 件（「QR出勤の読み込み」シート）", "")
        Exit Function
    End If
    n = lo.ListRows.Count
    first = 1
    If n > CAP Then first = n - CAP + 1          ' 表は古い順なので、新しい CAP 件
    v = lo.DataBodyRange.Value2
    ReDim out(1 To n - first + 1, 1 To 5)
    For i = first To n
        For j = 1 To 5
            out(i - first + 1, j) = CStr(v(i, j))
        Next j
    Next i
    With dst.Range("A2:E" & (CAP + 1))
        .ClearContents
        .NumberFormat = "@"
    End With
    dst.Range("A2").Resize(n - first + 1, 5).Value2 = out
    FillQrData = "QR出勤：" & files & " ファイルから " & (n - first + 1) & " 件を「出力データ貼付」に入れました" & _
                 IIf(n > CAP, "（" & CAP & " 件を超えたので、新しい分だけ）", "") & _
                 IIf(bad > 0, "。読めなかったファイル " & bad & " 件（「QR出勤の読み込み」シート）", "")
End Function
