Attribute VB_Name = "LeaveSheetMaker"
Option Explicit
' 休暇申請シートを職員の数だけ作る・全員分を更新する（管理者用。休暇申請シートの作成.xlsm の「設定」シートのボタンから実行する）。
' 出勤・休暇管理データ の「職員名簿」（出勤用ID・氏名）と「休暇等」シート（年ごとの合計・本年使用の調整）を読み、
' 各休暇申請シートの非表示の「管理」シートに 出勤用ID・氏名・年・合計（分）・本年使用の調整（分）を書き込む。
' 調整には、休暇等シートの本年使用の調整に、Forms で申請して承認された年次休暇（休暇申請シートには書かれない分）を足して書く。
'   ［休暇申請シートを作成］… 職員名簿の全員について、まだない人の休暇申請シートをひな形から作る（あれば作らない・上書きしない）
'   ［休暇申請シートを更新］… 作成先にある全員の休暇申請シートの「管理」シートを、休暇等シートの今の値に書き換える
'   ［様式を移行］…… 様式の古い休暇申請シートを、書いてある申請の行を残して新しいひな形の様式にする

Private Const DAY_MIN As Double = 465     ' 年次休暇の1日＝7時間45分＝465分
Private Const STAFF_LAST As Long = 501    ' 出勤・休暇管理データ の職員名簿の最後の行
Private Const LEAVE_LAST As Long = 1001   ' 出勤・休暇管理データ の休暇等（計算用シート）の最後の行

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
    Dim wb As Workbook, lo As ListObject, vals() As Variant, hasValue() As Boolean
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
    Application.Calculation = xlCalculationAutomatic
    wb.Close SaveChanges:=True
    Set wb = Nothing
    MigrateOne = "移行しました（申請 " & rowCount & " 行）。元のファイルの控え：" & backupDir & fileName
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
