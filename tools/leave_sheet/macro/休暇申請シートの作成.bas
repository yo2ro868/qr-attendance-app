Attribute VB_Name = "LeaveSheetMaker"
Option Explicit
' 休暇申請シートを職員の数だけ作る・全員分を更新する（管理者用。休暇申請シートの作成.xlsm の「設定」シートのボタンから実行する）。
' 出勤・休暇管理データ の「職員名簿」（出勤用ID・氏名）と「休暇等」シート（年ごとの合計・本年使用の調整）を読み、
' 各休暇申請シートの非表示の「管理」シートに 出勤用ID・氏名・年・合計（分）・本年使用の調整（分）を書き込む。
' 調整には、休暇等シートの本年使用の調整に、Forms で申請して承認された年次休暇（休暇申請シートには書かれない分）を足して書く。
'   ［休暇申請シートを作成］… 職員名簿の全員について、まだない人の休暇申請シートをひな形から作る（あれば作らない・上書きしない）
'   ［休暇申請シートを更新］… 作成先にある全員の休暇申請シートの「管理」シートを、休暇等シートの今の値に書き換える
'   ［様式を移行］…… 様式の古い休暇申請シートを、書いてある申請の行と超過勤務等命令簿の月のシートを残して新しいひな形の様式にする
'   ［超過勤務の月を追加］… 作成先にある全員の休暇申請シートに、B7 の月の超過勤務等命令簿のシート（超勤 R8.12 など）を作る

Private Const DAY_MIN As Double = 465     ' 年次休暇の1日＝7時間45分＝465分
Private Const STAFF_LAST As Long = 501    ' 出勤・休暇管理データ の職員名簿の最後の行
Private Const LEAVE_LAST As Long = 1001   ' 出勤・休暇管理データ の休暇等（計算用シート）の最後の行
' 超過勤務等命令簿の月のシート：本人・所属長が書く欄（様式を移行するとき、月のシートの値をこの範囲だけ写す）と、ひな形のシートの名前
Private Const OT_INPUTS As String = "A5:B66,D5:I66,K5:L66,Y5:AA66,AC5:AE66,L69:O71,AC69:AE70,BH1"
Private Const OT_TEMPLATE As String = "超勤ひな形"

' 出勤・休暇管理データ から読んだ値
Private staff As Variant      ' 職員名簿 A:B（出勤用ID・氏名）
Private keys As Variant       ' 計算用 H（出勤用ID|年）
Private totals As Variant     ' 計算用 K（合計・分）
Private adjust As Variant     ' 休暇等 G:H（本年使用の調整：日・時間）
Private formsMin As Variant   ' 計算用 W（Forms で承認された年次休暇・分）
Private holidays As Variant   ' 祝日 A（休暇申請シートの日数の計算で除く日）
Private formVersion As String ' ひな形の様式の版（管理シート B11）
Private carries As Variant    ' 計算用 J（繰越分・分）
Private grants As Variant     ' 計算用 P（本年分・日）

Public Sub CreateLeaveSheets()
    RunAll True, False
End Sub

Public Sub UpdateLeaveSheets()
    RunAll False, False
End Sub

Public Sub MigrateLeaveSheets()
    MigrateAll False
End Sub

Public Sub AddOvertimeMonths()
    AddMonthAll False
End Sub

' ===== 超過勤務等命令簿：月のシート =====
' 休暇申請シートの非表示の「超勤ひな形」シート（tools/overtime/add_overtime_sheet.ps1 が入れる）を写して「超勤 R8.12」のような月のシートを作る。
' W1 に月の1日、BH1 に前月のシートの BH2（前月の最終日の退庁時刻。25:00 のように日をまたぐ分が1日の欄に入る）を指す式を書き、
' 月にない日（30日の月の31日など）の行を隠す。前月のシートがなければ BH1 は空（必要なら管理者が手で時刻を入れる）。

Public Sub AddMonthAll(ByVal quiet As Boolean)
    Dim cfg As Worksheet, res As Worksheet
    Set cfg = ThisWorkbook.Worksheets("設定")
    Set res = ThisWorkbook.Worksheets("結果")
    Dim outDir As String, ym As Date, problem As String
    outDir = CleanPath(cfg.Range("B5").Value)
    If outDir = "" Or Dir(outDir, vbDirectory) = "" Then problem = "作成先のフォルダが見つかりません（B5）。"
    If problem = "" And Not IsDate(cfg.Range("B7").Value) Then problem = "超過勤務の月を入れてください（B7。例 2026/12）。"
    res.Range("A2:E" & res.Rows.Count).ClearContents
    If problem <> "" Then
        res.Range("D2").Value = problem
        If Not quiet Then MsgBox problem, vbExclamation
        Exit Sub
    End If
    ym = DateSerial(Year(cfg.Range("B7").Value), Month(cfg.Range("B7").Value), 1)
    If Right$(outDir, 1) <> "\" Then outDir = outDir & "\"
    If Not quiet Then
        If MsgBox("作成先の全員の休暇申請シートに、超過勤務等命令簿の「" & OvertimeSheetName(ym) & "」のシートを作ります。" & vbCrLf & _
                  "職員や所属長が休暇申請シートを開いていない時間に実行してください。" & vbCrLf & vbCrLf & _
                  "実行しますか？", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    End If
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

    Dim names As New Collection, f As String, item As Variant
    f = Dir(outDir & "休暇申請*.xlsx")
    Do While f <> ""
        If IsSheetFile(f) Then names.Add f
        f = Dir()
    Loop
    Dim row As Long, note As String, done As Long, other As Long, parts() As String, id As String, nm As String
    row = 2
    For Each item In names
        parts = NameParts(CStr(item))
        id = parts(1)
        nm = ""
        If UBound(parts) >= 2 Then nm = parts(2)
        note = AddMonthOne(outDir & item, ym)
        If Left$(note, 4) = "作りまし" Then done = done + 1 Else other = other + 1
        WriteResult res, row, id, nm, CStr(item), note
        row = row + 1
    Next item
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    res.Activate
    If Not quiet Then MsgBox "「" & OvertimeSheetName(ym) & "」を作ったもの " & done & " 件、そのほか " & other & " 件です（結果シートを確かめてください）。", vbInformation
End Sub

' 休暇申請シート1つに月のシートを作る。結果の文を返す
Private Function AddMonthOne(path As String, ym As Date) As String
    On Error GoTo Failed
    Dim wb As Workbook, note As String
    Set wb = Workbooks.Open(Filename:=path, UpdateLinks:=0)
    note = AddMonthSheet(wb, ym)
    If Left$(note, 4) = "作りまし" Then
        wb.Close SaveChanges:=True
    Else
        wb.Close SaveChanges:=False
    End If
    AddMonthOne = note
    Exit Function
Failed:
    AddMonthOne = "作れませんでした：" & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Function

' 開いているブックに月のシートを作る（保存はしない）。結果の文を返す
Private Function AddMonthSheet(wb As Workbook, ym As Date) As String
    Dim tpl As Worksheet, ws As Worksheet, s As Worksheet, anchor As Worksheet, nm As String
    nm = OvertimeSheetName(ym)
    Set tpl = SheetByName(wb, OT_TEMPLATE)
    If tpl Is Nothing Then
        AddMonthSheet = "様式が古いので作れませんでした（先に［様式を移行］してください）"
        Exit Function
    End If
    If Not SheetByName(wb, nm) Is Nothing Then
        AddMonthSheet = "すでにあるので作りませんでした"
        Exit Function
    End If
    ' 置く場所：この月より後の月のシートの前。なければ「使い方」の前（なければ最後）
    For Each s In wb.Worksheets
        If OvertimeMonth(s.Name) > ym And anchor Is Nothing Then Set anchor = s
    Next s
    If anchor Is Nothing Then Set anchor = SheetByName(wb, "使い方")
    ' 写したシートを名前の違いで探す（非表示のシートを写すと、写したシートも非表示で、作業中のシートにならないため）
    Dim before As Object
    Set before = CreateObject("Scripting.Dictionary")
    For Each s In wb.Worksheets
        before(s.Name) = True
    Next s
    tpl.Visible = xlSheetHidden          ' 完全に非表示（veryHidden）のシートは写せない
    If anchor Is Nothing Then
        tpl.Copy After:=wb.Worksheets(wb.Worksheets.Count)
    Else
        tpl.Copy Before:=anchor
    End If
    tpl.Visible = xlSheetVeryHidden
    For Each s In wb.Worksheets
        If Not before.Exists(s.Name) Then Set ws = s
    Next s
    ws.Name = nm
    ws.Visible = xlSheetVisible
    ws.Unprotect
    ws.Range("W1").Value = ym
    ' 前月の最終日の退庁時刻（前月のシートがあるときだけ。ないシートを指す式を書くとファイルを探す画面が出るため）
    Dim prevName As String
    prevName = OvertimeSheetName(DateAdd("m", -1, ym))
    If Not SheetByName(wb, prevName) Is Nothing Then ws.Range("BH1").Formula = "='" & prevName & "'!$BH$2"
    ' 月にない日の行を隠す
    Dim lastDay As Long, d As Long
    lastDay = Day(DateSerial(Year(ym), Month(ym) + 1, 0))
    ws.Rows("5:66").Hidden = False
    For d = lastDay + 1 To 31
        ws.Rows((5 + 2 * (d - 1)) & ":" & (6 + 2 * (d - 1))).Hidden = True
    Next d
    ProtectOvertime ws
    ' 翌月のシートが先にあれば、その BH1 をこのシートにつなぐ（手で時刻を入れてあれば変えない）
    Dim nextWs As Worksheet
    Set nextWs = SheetByName(wb, OvertimeSheetName(DateAdd("m", 1, ym)))
    If Not nextWs Is Nothing Then
        If Not nextWs.Range("BH1").HasFormula And CStr(nextWs.Range("BH1").Value) = "" Then
            nextWs.Unprotect
            nextWs.Range("BH1").Formula = "='" & nm & "'!$BH$2"
            ProtectOvertime nextWs
        End If
    End If
    AddMonthSheet = "作りました（" & nm & "）"
End Function

Private Sub ProtectOvertime(ws As Worksheet)
    ' 書式（行の高さ・列の幅）の変更はできる。add_overtime_sheet.ps1 と同じ設定
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=False, _
               AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
End Sub

' 2026/12/1 → 超勤 R8.12
Private Function OvertimeSheetName(ym As Date) As String
    OvertimeSheetName = "超勤 R" & (Year(ym) - 2018) & "." & Month(ym)
End Function

' 超勤 R8.12 → 2026/12/1。月のシートでなければ 0
Private Function OvertimeMonth(ByVal nm As String) As Date
    If Not nm Like "超勤 R*.*" Then Exit Function
    Dim parts() As String
    parts = Split(Mid$(nm, 5), ".")
    If UBound(parts) <> 1 Then Exit Function
    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Then Exit Function
    If Val(parts(1)) < 1 Or Val(parts(1)) > 12 Then Exit Function
    OvertimeMonth = DateSerial(2018 + Val(parts(0)), Val(parts(1)), 1)
End Function

Private Function SheetByName(wb As Workbook, nm As String) As Worksheet
    Dim s As Worksheet
    For Each s In wb.Worksheets
        If s.Name = nm Then Set SheetByName = s: Exit Function
    Next s
End Function

' 月のシートの書いた値（式でないセル）を「月 → (番地, 値) の一覧」で読む（様式を移行するとき用）
Private Function ReadOvertimeInputs(wb As Workbook) As Collection
    Dim result As New Collection, s As Worksheet, c As Range, items As Collection
    For Each s In wb.Worksheets
        If OvertimeMonth(s.Name) > 0 Then
            Set items = New Collection
            For Each c In s.Range(OT_INPUTS).Cells
                If Not c.HasFormula Then
                    If Not IsError(c.Value2) Then If CStr(c.Value2) <> "" Then items.Add Array(c.Address(False, False), c.Value2)
                End If
            Next c
            result.Add Array(OvertimeMonth(s.Name), items)
        End If
    Next s
    Set ReadOvertimeInputs = result
End Function

' 読んでおいた月のシートを、新しい様式のブックに作り直して値を書く。作った月の数を返す
Private Function WriteOvertimeInputs(wb As Workbook, months As Collection) As Long
    ' 前月のシートにつなぐため、古い月から作る
    Dim order() As Long, i As Long, j As Long, tmp As Long, n As Long
    n = months.Count
    If n = 0 Then Exit Function
    ReDim order(1 To n)
    For i = 1 To n
        order(i) = i
    Next i
    For i = 1 To n - 1
        For j = i + 1 To n
            If months(order(j))(0) < months(order(i))(0) Then tmp = order(i): order(i) = order(j): order(j) = tmp
        Next j
    Next i
    Dim note As String, ws As Worksheet, cell As Variant
    For i = 1 To n
        note = AddMonthSheet(wb, CDate(months(order(i))(0)))
        If Left$(note, 4) <> "作りまし" Then Err.Raise vbObjectError + 2, , "超過勤務のシート：" & note
        Set ws = SheetByName(wb, OvertimeSheetName(CDate(months(order(i))(0))))
        ws.Unprotect
        For Each cell In months(order(i))(1)
            ws.Range(cell(0)).Value2 = cell(1)
        Next cell
        ProtectOvertime ws
        WriteOvertimeInputs = WriteOvertimeInputs + 1
    Next i
End Function

' 様式の古い休暇申請シートを新しい様式にする。元のファイルは控えのフォルダに複製してから、同じファイルに上書きする
' （別のファイルに置き換えると OneDrive の共有リンクが切れるため）。quiet が True のときはメッセージを出さない
Public Sub MigrateAll(ByVal quiet As Boolean)
    Dim cfg As Worksheet, res As Worksheet
    Set cfg = ThisWorkbook.Worksheets("設定")
    Set res = ThisWorkbook.Worksheets("結果")
    Dim matchPath As String, tplPath As String, outDir As String, yr As Long, problem As String
    matchPath = CleanPath(cfg.Range("B3").Value)
    tplPath = CleanPath(cfg.Range("B4").Value)
    outDir = CleanPath(cfg.Range("B5").Value)
    yr = Val(cfg.Range("B6").Value)
    If Not FileExists(matchPath) Then problem = "出勤・休暇管理データ のファイルが見つかりません（B3）。"
    If problem = "" And Not FileExists(tplPath) Then problem = "ひな形のファイルが見つかりません（B4）。"
    If problem = "" And (outDir = "" Or Dir(outDir, vbDirectory) = "") Then problem = "作成先のフォルダが見つかりません（B5）。"
    If problem = "" And (yr < 2000 Or yr > 2100) Then problem = "年（西暦）を入れてください（B6）。"
    res.Range("A2:E" & res.Rows.Count).ClearContents
    If problem = "" Then
        If Right$(outDir, 1) <> "\" Then outDir = outDir & "\"
        formVersion = ReadFormVersion(tplPath)
        If formVersion = "" Then problem = "ひな形に様式の版がありません（新しいひな形か確かめてください）。"
    End If
    If problem <> "" Then
        res.Range("D2").Value = problem
        If Not quiet Then MsgBox problem, vbExclamation
        Exit Sub
    End If
    If Not quiet Then
        If MsgBox("様式の古い休暇申請シートを、新しいひな形の様式にします。" & vbCrLf & _
                  "職員や所属長が休暇申請シートを開いていない時間に実行してください。" & vbCrLf & _
                  "元のファイルは作成先の中の控えのフォルダに残します。" & vbCrLf & vbCrLf & _
                  "実行しますか？", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    ReadMatching matchPath

    ' 先にファイルの名前を集める（処理の途中で Dir の続きが使えなくなるため）
    Dim names As New Collection, f As String, item As Variant
    f = Dir(outDir & "休暇申請*.xlsx")
    Do While f <> ""
        If IsSheetFile(f) Then names.Add f
        f = Dir()
    Loop

    Dim backupDir As String, row As Long, note As String, done As Long, other As Long, parts() As String, id As String, nm As String
    backupDir = outDir & "旧様式_" & Format(Now, "yyyymmdd_hhnn") & "\"
    row = 2
    For Each item In names
        parts = NameParts(CStr(item))
        id = parts(1)
        nm = ""
        If UBound(parts) >= 2 Then nm = parts(2)
        note = MigrateOne(outDir & item, CStr(item), tplPath, backupDir, id, nm, yr)
        If Left$(note, 4) = "移行しま" Then done = done + 1 Else other = other + 1
        ' 前の名前（休暇申請_ID_氏名）のファイルは、今の名前に直す
        Dim shownName As String
        shownName = CStr(item)
        If nm <> "" And StandardName(id, nm) <> shownName Then
            If Dir(outDir & StandardName(id, nm)) = "" Then
                If RenameFile(outDir & shownName, outDir & StandardName(id, nm)) Then
                    shownName = StandardName(id, nm)
                    note = note & "（ファイル名を直しました）"
                Else
                    note = note & "（開かれているのでファイル名は直せませんでした）"
                End If
            End If
        End If
        WriteResult res, row, id, nm, shownName, note
        row = row + 1
    Next item

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    res.Activate
    If Not quiet Then MsgBox "移行 " & done & " 件、そのほか " & other & " 件です（結果シートを確かめてください）。", vbInformation
End Sub

' 休暇申請シート1つを移行する。結果の文を返す
Private Function MigrateOne(path As String, fileName As String, tplPath As String, backupDir As String, _
                            id As String, nm As String, yr As Long) As String
    ' 本人・所属長が書く列（日数と時間は自動）
    Dim cols As Variant, c As Long, r As Long, n As Long, overwritten As Boolean
    cols = Array("申請日", "休暇の種類", "種別・内容", "単位", "開始日", "終了日", "開始時刻", "終了時刻", _
                 "事由", "承認（所属長）", "承認日", "所属長コメント")
    Dim wb As Workbook, lo As ListObject, vals() As Variant, hasValue() As Boolean, otMonths As Collection, otCount As Long
    On Error GoTo Failed

    ' 古いファイルから読む（式の入ったセルは、新しい様式の式に任せるので写さない）
    Set wb = Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True)
    If CStr(SheetValue(wb, "管理", "B11")) = formVersion Then
        wb.Close SaveChanges:=False
        Set wb = Nothing
        MigrateOne = "新しい様式なので何もしませんでした"
        Exit Function
    End If
    Set lo = wb.Worksheets("申請").ListObjects("申請")
    n = lo.ListRows.Count
    ReDim vals(1 To n, 0 To UBound(cols))
    ReDim hasValue(1 To n, 0 To UBound(cols))
    For c = 0 To UBound(cols)
        If HasColumn(lo, CStr(cols(c))) Then
            For r = 1 To n
                With lo.ListColumns(CStr(cols(c))).DataBodyRange.Cells(r)
                    If Not .HasFormula And CStr(.Value2) <> "" Then
                        vals(r, c) = .Value2
                        hasValue(r, c) = True
                    End If
                End With
            Next r
        End If
    Next c
    ' 時間は写さない（時刻から計算する）。時刻がなく時間だけを打ち込んだ行は、あとで時刻を入れてもらう
    Dim noClock As Long
    If HasColumn(lo, "時間") Then
        For r = 1 To n
            With lo.ListColumns("時間").DataBodyRange.Cells(r)
                If Not .HasFormula And CStr(.Value2) <> "" And Not hasValue(r, 6) Then noClock = noClock + 1
            End With
        Next r
    End If
    Set otMonths = ReadOvertimeInputs(wb)      ' 超過勤務等命令簿の月のシートの書いた値
    wb.Close SaveChanges:=False
    Set wb = Nothing

    ' 控えを残して、同じファイルに新しい様式を上書きする
    If Dir(backupDir, vbDirectory) = "" Then MkDir backupDir
    FileCopy path, backupDir & fileName
    FileCopy tplPath, path
    overwritten = True
    Dim written As String
    written = UpdateOne(path, id, nm, yr)
    If Left$(written, 4) <> "更新しま" Then Err.Raise vbObjectError + 1, , written

    ' 申請の行を写す
    Dim rowCount As Long, target As Long, anyValue As Boolean, text As String
    Set wb = Workbooks.Open(Filename:=path, UpdateLinks:=0)
    Set lo = wb.Worksheets("申請").ListObjects("申請")
    Application.Calculation = xlCalculationManual     ' 1セルごとに再計算しないように
    target = 0
    For r = 1 To n
        anyValue = False
        For c = 0 To UBound(cols)
            If hasValue(r, c) Then anyValue = True
        Next c
        If anyValue Then
            target = target + 1
            For c = 0 To UBound(cols)
                If hasValue(r, c) Then
                    If cols(c) = "単位" Then
                        ' 古い表記を新しい表記に
                        text = CStr(vals(r, c))
                        If text = "1日単位" Then text = "日単位"
                        If text = "時間単位" Then text = "時間・分単位"
                        lo.ListColumns(CStr(cols(c))).DataBodyRange.Cells(target).Value2 = text
                    Else
                        lo.ListColumns(CStr(cols(c))).DataBodyRange.Cells(target).Value2 = vals(r, c)
                    End If
                End If
            Next c
        End If
    Next r
    rowCount = target
    otCount = WriteOvertimeInputs(wb, otMonths)
    Application.Calculation = xlCalculationAutomatic
    wb.Close SaveChanges:=True
    Set wb = Nothing
    MigrateOne = "移行しました（申請 " & rowCount & " 行、超過勤務の月 " & otCount & "）。元のファイルの控え：" & backupDir & fileName
    If noClock > 0 Then MigrateOne = MigrateOne & "【時刻がなく時間だけの行が " & noClock & " 行あります。開始時刻・終了時刻を入れてください（赤くなっています）】"
    Exit Function
Failed:
    Dim message As String
    message = Err.Description
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    If overwritten Then FileCopy backupDir & fileName, path    ' 途中で失敗したら元に戻す
    MigrateOne = "移行できませんでした（元のままです）：" & message
End Function

Private Function RenameFile(fromPath As String, toPath As String) As Boolean
    On Error GoTo Failed
    Name fromPath As toPath
    RenameFile = True
    Exit Function
Failed:
    RenameFile = False
End Function

' 個人別のファイル名：休暇申請シート　出勤用ID　氏名.xlsx
Private Function StandardName(id As String, nm As String) As String
    StandardName = "休暇申請シート　" & SafeName(id) & "　" & SafeName(nm) & ".xlsx"
End Function

' 個人別の休暇申請シートのファイルか（ひな形は違う）。区切りは全角スペース（半角スペース・「_」、前の名前の 休暇申請_ でも可）
Private Function IsSheetFile(f As String) As Boolean
    IsSheetFile = (f Like "休暇申請シート[　 _]*.xlsx") Or (f Like "休暇申請_*.xlsx")
End Function

' ファイル名を 休暇申請シート / 出勤用ID / 氏名 に分ける
Private Function NameParts(f As String) As String()
    Dim s As String
    s = Replace(Replace(Left$(f, Len(f) - 5), "_", "　"), " ", "　")
    Do While InStr(s, "　　") > 0
        s = Replace(s, "　　", "　")
    Loop
    NameParts = Split(s, "　")
End Function

Private Function HasColumn(lo As ListObject, name As String) As Boolean
    Dim col As ListColumn
    For Each col In lo.ListColumns
        If col.Name = name Then HasColumn = True: Exit Function
    Next col
End Function

' シートがなければ Empty を返す
Private Function SheetValue(wb As Workbook, sheetName As String, ref As String) As Variant
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Name = sheetName Then SheetValue = ws.Range(ref).Value: Exit Function
    Next ws
End Function

' create が True なら作成、False なら更新。quiet が True のときはメッセージを出さない（動作確認用）
Public Sub RunAll(ByVal create As Boolean, ByVal quiet As Boolean)
    Dim cfg As Worksheet, res As Worksheet
    Set cfg = ThisWorkbook.Worksheets("設定")
    Set res = ThisWorkbook.Worksheets("結果")

    Dim matchPath As String, tplPath As String, outDir As String, pwd As String, yr As Long
    matchPath = CleanPath(cfg.Range("B3").Value)
    tplPath = CleanPath(cfg.Range("B4").Value)
    outDir = CleanPath(cfg.Range("B5").Value)
    yr = Val(cfg.Range("B6").Value)
    Dim problem As String
    If Not FileExists(matchPath) Then problem = "出勤・休暇管理データ のファイルが見つかりません（B3）。"
    If problem = "" And create And Not FileExists(tplPath) Then problem = "ひな形のファイルが見つかりません（B4）。"
    If problem = "" And (outDir = "" Or Dir(outDir, vbDirectory) = "") Then problem = "作成先のフォルダが見つかりません（B5）。"
    If problem = "" And (yr < 2000 Or yr > 2100) Then problem = "年（西暦）を入れてください（B6）。"
    res.Range("A2:E" & res.Rows.Count).ClearContents
    If problem <> "" Then
        res.Range("D2").Value = problem
        If Not quiet Then MsgBox problem, vbExclamation
        Exit Sub
    End If
    If Right$(outDir, 1) <> "\" Then outDir = outDir & "\"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    ReadMatching matchPath
    formVersion = ReadFormVersion(tplPath)

    ' 作成先にある休暇申請シート（出勤用ID → ファイル名）。出勤用ID はファイル名の2つ目の部分
    Dim existing As Object, f As String
    Set existing = CreateObject("Scripting.Dictionary")
    f = Dir(outDir & "休暇申請*.xlsx")
    Do While f <> ""
        If IsSheetFile(f) Then existing(NormId(NameParts(f)(1))) = f
        f = Dir()
    Loop

    Dim row As Long, done As Long, other As Long, note As String
    row = 2
    If create Then
        Dim i As Long, id As String, nm As String, fileName As String
        For i = 1 To UBound(staff, 1)
            id = Trim$(CStr(staff(i, 1)))
            nm = Trim$(CStr(staff(i, 2)))
            If id <> "" And nm <> "" And InStr(nm, "（例）") = 0 Then
                If existing.Exists(NormId(id)) Then
                    fileName = existing(NormId(id))
                    note = "すでにあるので作りませんでした"
                    other = other + 1
                Else
                    fileName = StandardName(id, nm)
                    note = CreateOne(tplPath, outDir & fileName, id, nm, yr)
                    If Left$(note, 4) = "作成しま" Then done = done + 1: existing(NormId(id)) = fileName
                End If
                WriteResult res, row, id, nm, fileName, note
                row = row + 1
            End If
        Next i
    Else
        Dim k As Variant, parts() As String
        For Each k In existing.Keys
            fileName = existing(k)
            parts = NameParts(fileName)     ' 休暇申請シート / 出勤用ID / 氏名
            id = parts(1)
            nm = ""
            If UBound(parts) >= 2 Then nm = parts(2)
            note = UpdateOne(outDir & fileName, id, nm, yr)
            If Left$(note, 4) = "更新しま" Then done = done + 1 Else other = other + 1
            WriteResult res, row, id, nm, fileName, note
            row = row + 1
        Next k
    End If

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    res.Activate
    If quiet Then Exit Sub
    If create Then
        MsgBox "作成 " & done & " 件、すでにあったもの " & other & " 件です。" & vbCrLf & _
               "次に、作った休暇申請シートを OneDrive で本人と所属長に共有してください。", vbInformation
    Else
        MsgBox "更新 " & done & " 件、更新できなかったもの " & other & " 件です（結果シートを確かめてください）。", vbInformation
    End If
End Sub

' ひな形の様式の版を読む。読めなければ ""（古い様式の確認はしない）
Private Function ReadFormVersion(tplPath As String) As String
    On Error GoTo Failed
    If Not FileExists(tplPath) Then Exit Function
    Dim wb As Workbook
    Set wb = Workbooks.Open(Filename:=tplPath, UpdateLinks:=0, ReadOnly:=True)
    ReadFormVersion = CStr(wb.Worksheets("管理").Range("B11").Value)
    wb.Close SaveChanges:=False
    Exit Function
Failed:
    ReadFormVersion = ""
End Function

' 出勤・休暇管理データ を読み取り専用で開いて（すでに開いていればそれを使って）、必要な範囲の値を読む
Private Sub ReadMatching(matchPath As String)
    Dim src As Workbook, openedHere As Boolean
    Set src = FindOpenWorkbook(matchPath)
    If src Is Nothing Then
        Set src = Workbooks.Open(Filename:=matchPath, UpdateLinks:=0, ReadOnly:=True)
        openedHere = True
    End If
    staff = src.Worksheets("職員名簿").Range("A2:F" & STAFF_LAST).Value2     ' E・F は休憩の始め・終わり（短縮の特例の人だけ）
    keys = src.Worksheets("計算用").Range("H2:H" & LEAVE_LAST).Value
    totals = src.Worksheets("計算用").Range("K2:K" & LEAVE_LAST).Value
    adjust = src.Worksheets("休暇等").Range("G2:H" & LEAVE_LAST).Value
    formsMin = src.Worksheets("計算用").Range("W2:W" & LEAVE_LAST).Value
    holidays = src.Worksheets("祝日").Range("A2:A200").Value
    carries = src.Worksheets("計算用").Range("J2:J" & LEAVE_LAST).Value
    grants = src.Worksheets("計算用").Range("P2:P" & LEAVE_LAST).Value
    If openedHere Then src.Close SaveChanges:=False
End Sub

' ひな形をコピーして「管理」シートに書き込む。結果の文を返す
Private Function CreateOne(tplPath As String, dest As String, id As String, nm As String, yr As Long) As String
    On Error GoTo Failed
    FileCopy tplPath, dest
    CreateOne = Replace(UpdateOne(dest, id, nm, yr), "更新しました", "作成しました")
    Exit Function
Failed:
    CreateOne = "作れませんでした：" & Err.Description
End Function

' 休暇申請シートを開いて「管理」シートに書き込む。結果の文を返す
Private Function UpdateOne(path As String, id As String, nm As String, yr As Long) As String
    On Error GoTo Failed
    Dim pos As Long, hold As Variant, adj As Variant, carry As Variant, grant As Variant, wb As Workbook, ws As Worksheet
    pos = FindRow(NormId(id) & "|" & yr)
    If pos > 0 Then
        If IsNumeric(totals(pos, 1)) And CStr(totals(pos, 1)) <> "" Then hold = totals(pos, 1)
        If IsNumeric(carries(pos, 1)) And CStr(carries(pos, 1)) <> "" Then carry = carries(pos, 1)
        If IsNumeric(grants(pos, 1)) And CStr(grants(pos, 1)) <> "" Then grant = grants(pos, 1) * DAY_MIN
        adj = Round(Val(adjust(pos, 1)) * DAY_MIN + Val(adjust(pos, 2)) * 1440 + Val(formsMin(pos, 1)), 0)
    End If
    Set wb = Workbooks.Open(Filename:=path, UpdateLinks:=0)
    Set ws = wb.Worksheets("管理")
    Dim oldForm As Boolean
    oldForm = (formVersion <> "" And CStr(ws.Range("B11").Value) <> formVersion)
    ws.Range("B1").NumberFormat = "@"
    ws.Range("B1").Value = id
    ws.Range("B2").Value = nm
    ws.Range("B3").Value = yr
    ws.Range("B4").Value = hold
    ws.Range("B5").Value = adj
    ws.Range("B6").Value = Now
    ' 休憩時間：職員名簿に書いてあればそれ、なければ 12:00〜13:00
    Dim breakStart As Double, breakEnd As Double, k As Long
    breakStart = 12 / 24
    breakEnd = 13 / 24
    For k = 1 To UBound(staff, 1)
        If NormId(CStr(staff(k, 1))) = NormId(id) Then
            If IsNumeric(staff(k, 5)) And IsNumeric(staff(k, 6)) And CStr(staff(k, 5)) <> "" And CStr(staff(k, 6)) <> "" Then
                breakStart = staff(k, 5)
                breakEnd = staff(k, 6)
            End If
            Exit For
        End If
    Next k
    ws.Range("B9").Value = breakStart
    ws.Range("B10").Value = breakEnd
    ws.Range("B7").Value = carry
    ws.Range("B8").Value = grant
    ' 祝日（日付だけを書く。文字があると休暇申請シートの日数の計算がエラーになるため）
    Dim i As Long, n As Long
    ws.Range("D1:D200").ClearContents
    For i = 1 To UBound(holidays, 1)
        If IsDate(holidays(i, 1)) Then
            n = n + 1
            ws.Cells(n, 4).Value = CDate(holidays(i, 1))
        End If
    Next i
    wb.Close SaveChanges:=True
    If pos > 0 Then
        UpdateOne = "更新しました（繰越分 " & DaysText(carry) & "、本年分 " & DaysText(grant) & "、合計 " & DaysText(hold) & "、調整（Forms の分を含む） " & DaysText(adj) & "）"
    Else
        UpdateOne = "更新しました（合計は空欄：出勤・休暇管理データ の休暇等シートに " & yr & " 年の行がありません）"
    End If
    If oldForm Then UpdateOne = UpdateOne & "【様式が古いままです：計算式は新しくなっていません。ひな形から作り直してください】"
    Exit Function
Failed:
    UpdateOne = "書き込めませんでした：" & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Function

' 出勤・休暇管理データ の計算用シートの 年休キー（出勤用ID|年）の位置（1から）。なければ 0
Private Function FindRow(key As String) As Long
    Dim i As Long
    For i = 1 To UBound(keys, 1)
        If CStr(keys(i, 1)) = key Then FindRow = i: Exit Function
    Next i
End Function

Private Sub WriteResult(res As Worksheet, row As Long, id As String, nm As String, fileName As String, note As String)
    res.Cells(row, 1).NumberFormat = "@"
    res.Cells(row, 1).Value = id
    res.Cells(row, 2).Value = nm
    res.Cells(row, 3).Value = fileName
    res.Cells(row, 4).Value = note
End Sub

' 分を「15日3時間30分」の形に（1日＝7時間45分）
Private Function DaysText(v As Variant) As String
    If IsEmpty(v) Then DaysText = "なし": Exit Function
    Dim m As Long, rest As Long
    m = CLng(v)
    rest = m Mod CLng(DAY_MIN)
    DaysText = (m \ CLng(DAY_MIN)) & "日"
    If rest \ 60 > 0 Then DaysText = DaysText & (rest \ 60) & "時間"
    If rest Mod 60 > 0 Then DaysText = DaysText & (rest Mod 60) & "分"
End Function

' 出勤・休暇管理データ の照合キーと同じ：数だけの ID は先頭の0を取る
Private Function NormId(ByVal s As String) As String
    s = Trim$(s)
    If s <> "" And IsNumeric(s) Then NormId = CStr(CDbl(s)) Else NormId = s
End Function

' ファイル名に使えない文字と、区切りに使う空白・「_」を取り除く
Private Function SafeName(ByVal s As String) As String
    Dim c As Variant
    For Each c In Array("\", "/", ":", "*", "?", """", "<", ">", "|", " ", "　", "_")
        s = Replace(s, c, "")
    Next c
    SafeName = s
End Function

' セルに入れた場所の前後の空白と「"」（エクスプローラーの［パスとしてコピー］で付く）を取り除く
Private Function CleanPath(ByVal v As Variant) As String
    CleanPath = Trim$(Replace(CStr(v), """", ""))
End Function

Private Function FileExists(p As String) As Boolean
    If p = "" Then Exit Function
    FileExists = (Dir(p) <> "")
End Function

Private Function FindOpenWorkbook(p As String) As Workbook
    Dim wb As Workbook
    For Each wb In Workbooks
        If StrComp(wb.FullName, p, vbTextCompare) = 0 Then Set FindOpenWorkbook = wb: Exit Function
    Next wb
End Function
