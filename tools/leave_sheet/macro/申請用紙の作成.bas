Attribute VB_Name = "LeaveSheetMaker"
Option Explicit
' 休暇申請用紙を職員の数だけ作る（管理者用。申請用紙の作成.xlsm の「設定」シートのボタンから実行する）。
' 照合用Excel の「職員名簿」（出勤用ID・氏名）を読み、ひな形をコピーして 休暇申請_出勤用ID_氏名.xlsx を作り、
' 基本の欄に 出勤用ID・氏名・年 と、照合用Excel の年次休暇の合計（その年の行があれば）を入れる。
' 作成先にすでに同じ出勤用ID の申請用紙があれば作らない（上書きしない）。所属長は空欄のまま（手で入れる）。

Private Const DAY_MIN As Double = 465     ' 年次休暇の1日＝7時間45分＝465分
Private Const STAFF_LAST As Long = 501    ' 照合用Excel の職員名簿の最後の行
Private Const LEAVE_LAST As Long = 1001   ' 照合用Excel の年次休暇（計算用シート）の最後の行

Public Sub CreateLeaveSheets()
    MakeAll False
End Sub

' quiet が True のときはメッセージを出さない（動作確認用）
Public Sub MakeAll(ByVal quiet As Boolean)
    Dim cfg As Worksheet, res As Worksheet
    Set cfg = ThisWorkbook.Worksheets("設定")
    Set res = ThisWorkbook.Worksheets("結果")

    Dim matchPath As String, tplPath As String, outDir As String, pwd As String, yr As Long
    matchPath = CleanPath(cfg.Range("B3").Value)
    tplPath = CleanPath(cfg.Range("B4").Value)
    outDir = CleanPath(cfg.Range("B5").Value)
    yr = Val(cfg.Range("B6").Value)
    pwd = CStr(cfg.Range("B7").Value)
    Dim problem As String
    If Not FileExists(matchPath) Then problem = "照合用Excel のファイルが見つかりません（B3）。"
    If problem = "" And Not FileExists(tplPath) Then problem = "ひな形のファイルが見つかりません（B4）。"
    If problem = "" And (outDir = "" Or Dir(outDir, vbDirectory) = "") Then problem = "作成先のフォルダが見つかりません（B5）。"
    If problem = "" And (yr < 2000 Or yr > 2100) Then problem = "年（西暦）を入れてください（B6）。"
    If problem <> "" Then
        res.Range("A2:E" & res.Rows.Count).ClearContents
        res.Range("D2").Value = problem
        If Not quiet Then MsgBox problem, vbExclamation
        Exit Sub
    End If
    If Right$(outDir, 1) <> "\" Then outDir = outDir & "\"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

    ' 照合用Excel から 職員名簿 と 年次休暇の合計（計算用シートの 年休キー → 合計（分））を読む
    Dim src As Workbook, openedHere As Boolean
    Set src = FindOpenWorkbook(matchPath)
    If src Is Nothing Then
        Set src = Workbooks.Open(Filename:=matchPath, UpdateLinks:=0, ReadOnly:=True)
        openedHere = True
    End If
    Dim staff As Variant, keys As Variant, totals As Variant
    staff = src.Worksheets("職員名簿").Range("A2:B" & STAFF_LAST).Value
    keys = src.Worksheets("計算用").Range("H2:H" & LEAVE_LAST).Value
    totals = src.Worksheets("計算用").Range("K2:K" & LEAVE_LAST).Value
    If openedHere Then src.Close SaveChanges:=False

    ' 作成先にある申請用紙の出勤用ID（ファイル名の 休暇申請_ と次の _ の間）
    Dim existing As Object, f As String
    Set existing = CreateObject("Scripting.Dictionary")
    f = Dir(outDir & "休暇申請_*.xlsx")
    Do While f <> ""
        existing(NormId(Split(f, "_")(1))) = f
        f = Dir()
    Loop

    res.Range("A2:E" & res.Rows.Count).ClearContents
    Dim i As Long, row As Long, made As Long, skipped As Long
    Dim id As String, nm As String, key As String, fileName As String, note As String, days As Variant
    row = 2
    For i = 1 To UBound(staff, 1)
        id = Trim$(CStr(staff(i, 1)))
        nm = Trim$(CStr(staff(i, 2)))
        If id <> "" And nm <> "" And InStr(nm, "（例）") = 0 Then
            key = NormId(id)
            If existing.Exists(key) Then
                fileName = existing(key)
                note = "すでにあるので作りませんでした"
                skipped = skipped + 1
            Else
                fileName = "休暇申請_" & SafeName(id) & "_" & SafeName(nm) & ".xlsx"
                days = LookupDays(keys, totals, key & "|" & yr)
                note = MakeSheet(tplPath, outDir & fileName, id, nm, yr, days, pwd)
                If note = "" Then
                    note = "作成しました"
                    If IsEmpty(days) Then note = note & "（持ち日数は空欄：照合用Excel の年次休暇に " & yr & " 年の行がありません）"
                    made = made + 1
                    existing(key) = fileName
                End If
            End If
            res.Cells(row, 1).NumberFormat = "@"
            res.Cells(row, 1).Value = id
            res.Cells(row, 2).Value = nm
            res.Cells(row, 3).Value = fileName
            res.Cells(row, 4).Value = note
            row = row + 1
        End If
    Next i

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    res.Activate
    If Not quiet Then MsgBox "作成 " & made & " 件、すでにあったもの " & skipped & " 件です。" & vbCrLf & _
           "次に、作った申請用紙を OneDrive で本人と所属長に共有し、所属長の欄を入れてください。", vbInformation
End Sub

' ひな形をコピーして基本の欄を入れる。うまくいけば ""、だめなら理由を返す
Private Function MakeSheet(tplPath As String, dest As String, id As String, nm As String, yr As Long, _
                           days As Variant, pwd As String) As String
    On Error GoTo Failed
    FileCopy tplPath, dest
    Dim wb As Workbook, ws As Worksheet
    Set wb = Workbooks.Open(Filename:=dest, UpdateLinks:=0)
    Set ws = wb.Worksheets("申請")
    ws.Unprotect Password:=pwd
    ws.Range("A4").NumberFormat = "@"
    ws.Range("A4").Value = id
    ws.Range("B4").Value = nm
    ws.Range("D4").Value = yr
    If Not IsEmpty(days) Then ws.Range("E4").Value = days
    ws.Protect Password:=pwd, AllowFormattingColumns:=True, AllowFormattingRows:=True, _
               AllowSorting:=True, AllowFiltering:=True
    wb.Close SaveChanges:=True
    MakeSheet = ""
    Exit Function
Failed:
    MakeSheet = "作れませんでした：" & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Kill dest
End Function

' 照合用Excel の計算用シートの 年休キー（出勤用ID|年）から、合計（分）を日数にして返す。なければ Empty
Private Function LookupDays(keys As Variant, totals As Variant, key As String) As Variant
    Dim i As Long
    For i = 1 To UBound(keys, 1)
        If CStr(keys(i, 1)) = key Then
            If IsNumeric(totals(i, 1)) And CStr(totals(i, 1)) <> "" Then LookupDays = totals(i, 1) / DAY_MIN
            Exit Function
        End If
    Next i
End Function

' 照合用Excel の照合キーと同じ：数だけの ID は先頭の0を取る
Private Function NormId(ByVal s As String) As String
    s = Trim$(s)
    If s <> "" And IsNumeric(s) Then NormId = CStr(CDbl(s)) Else NormId = s
End Function

' ファイル名に使えない文字と空白を取り除く
Private Function SafeName(ByVal s As String) As String
    Dim c As Variant
    For Each c In Array("\", "/", ":", "*", "?", """", "<", ">", "|", " ", "　", "_")
        s = Replace(s, c, "")
    Next c
    SafeName = s
End Function

' セルに入れた場所の前後の空白と「"」（エクスプローラーの［パスのコピー］で付く）を取り除く
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
