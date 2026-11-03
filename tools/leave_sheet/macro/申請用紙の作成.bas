Attribute VB_Name = "LeaveSheetMaker"
Option Explicit
' 休暇申請用紙を職員の数だけ作る・全員分を更新する（管理者用。申請用紙の作成.xlsm の「設定」シートのボタンから実行する）。
' 照合用Excel の「職員名簿」（出勤用ID・氏名）と「休暇等」シート（年ごとの持ち日数・本年使用の調整）を読み、
' 各申請用紙の非表示の「管理」シートに 出勤用ID・氏名・年・持ち日数（分）・本年使用の調整（分）を書き込む。
' 調整には、休暇等シートの本年使用の調整に、Forms で申請して承認された年次休暇（申請用紙には書かれない分）を足して書く。
'   ［申請用紙を作成］… 職員名簿の全員について、まだない人の申請用紙をひな形から作る（あれば作らない・上書きしない）
'   ［申請用紙を更新］… 作成先にある全員の申請用紙の「管理」シートを、休暇等シートの今の値に書き換える

Private Const DAY_MIN As Double = 465     ' 年次休暇の1日＝7時間45分＝465分
Private Const STAFF_LAST As Long = 501    ' 照合用Excel の職員名簿の最後の行
Private Const LEAVE_LAST As Long = 1001   ' 照合用Excel の休暇等（計算用シート）の最後の行

' 照合用Excel から読んだ値
Private staff As Variant      ' 職員名簿 A:B（出勤用ID・氏名）
Private keys As Variant       ' 計算用 H（出勤用ID|年）
Private totals As Variant     ' 計算用 K（合計・分）
Private adjust As Variant     ' 休暇等 G:H（本年使用の調整：日・時間）
Private formsMin As Variant   ' 計算用 W（Forms で承認された年次休暇・分）

Public Sub CreateLeaveSheets()
    RunAll True, False
End Sub

Public Sub UpdateLeaveSheets()
    RunAll False, False
End Sub

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
    If Not FileExists(matchPath) Then problem = "照合用Excel のファイルが見つかりません（B3）。"
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

    ' 作成先にある申請用紙（出勤用ID → ファイル名）。出勤用ID はファイル名の 休暇申請_ と次の _ の間
    Dim existing As Object, f As String
    Set existing = CreateObject("Scripting.Dictionary")
    f = Dir(outDir & "休暇申請_*.xlsx")
    Do While f <> ""
        existing(NormId(Split(f, "_")(1))) = f
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
                    fileName = "休暇申請_" & SafeName(id) & "_" & SafeName(nm) & ".xlsx"
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
            parts = Split(Left$(fileName, Len(fileName) - 5), "_")     ' 休暇申請 / 出勤用ID / 氏名
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
               "次に、作った申請用紙を OneDrive で本人と所属長に共有してください。", vbInformation
    Else
        MsgBox "更新 " & done & " 件、更新できなかったもの " & other & " 件です（結果シートを確かめてください）。", vbInformation
    End If
End Sub

' 照合用Excel を読み取り専用で開いて（すでに開いていればそれを使って）、必要な範囲の値を読む
Private Sub ReadMatching(matchPath As String)
    Dim src As Workbook, openedHere As Boolean
    Set src = FindOpenWorkbook(matchPath)
    If src Is Nothing Then
        Set src = Workbooks.Open(Filename:=matchPath, UpdateLinks:=0, ReadOnly:=True)
        openedHere = True
    End If
    staff = src.Worksheets("職員名簿").Range("A2:B" & STAFF_LAST).Value
    keys = src.Worksheets("計算用").Range("H2:H" & LEAVE_LAST).Value
    totals = src.Worksheets("計算用").Range("K2:K" & LEAVE_LAST).Value
    adjust = src.Worksheets("休暇等").Range("G2:H" & LEAVE_LAST).Value
    formsMin = src.Worksheets("計算用").Range("W2:W" & LEAVE_LAST).Value
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

' 申請用紙を開いて「管理」シートに書き込む。結果の文を返す
Private Function UpdateOne(path As String, id As String, nm As String, yr As Long) As String
    On Error GoTo Failed
    Dim pos As Long, hold As Variant, adj As Variant, wb As Workbook, ws As Worksheet
    pos = FindRow(NormId(id) & "|" & yr)
    If pos > 0 Then
        If IsNumeric(totals(pos, 1)) And CStr(totals(pos, 1)) <> "" Then hold = totals(pos, 1)
        adj = Round(Val(adjust(pos, 1)) * DAY_MIN + Val(adjust(pos, 2)) * 1440 + Val(formsMin(pos, 1)), 0)
    End If
    Set wb = Workbooks.Open(Filename:=path, UpdateLinks:=0)
    Set ws = wb.Worksheets("管理")
    ws.Range("B1").NumberFormat = "@"
    ws.Range("B1").Value = id
    ws.Range("B2").Value = nm
    ws.Range("B3").Value = yr
    ws.Range("B4").Value = hold
    ws.Range("B5").Value = adj
    ws.Range("B6").Value = Now
    wb.Close SaveChanges:=True
    If pos > 0 Then
        UpdateOne = "更新しました（持ち日数 " & DaysText(hold) & "、調整（Forms の分を含む） " & DaysText(adj) & "）"
    Else
        UpdateOne = "更新しました（持ち日数は空欄：照合用Excel の休暇等シートに " & yr & " 年の行がありません）"
    End If
    Exit Function
Failed:
    UpdateOne = "書き込めませんでした：" & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Function

' 照合用Excel の計算用シートの 年休キー（出勤用ID|年）の位置（1から）。なければ 0
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

' 照合用Excel の照合キーと同じ：数だけの ID は先頭の0を取る
Private Function NormId(ByVal s As String) As String
    s = Trim$(s)
    If s <> "" And IsNumeric(s) Then NormId = CStr(CDbl(s)) Else NormId = s
End Function

' ファイル名に使えない文字と空白、区切りの「_」を取り除く
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
