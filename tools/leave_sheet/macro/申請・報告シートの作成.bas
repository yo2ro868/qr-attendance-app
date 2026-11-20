Attribute VB_Name = "LeaveSheetMaker"
Option Explicit
' 申請・報告シートを職員の数だけ作る・全員分を更新する（管理者用。申請・報告シートの作成.xlsm の「設定」シートのボタンから実行する）。
' 出勤・休暇管理データ の「台帳①職員名簿」（出勤用ID・氏名）と「休暇②個人別一覧」シート（年ごとの合計・本年使用の調整）を読み、
' 各申請・報告シートの非表示の「管理」シートに 出勤用ID・氏名・年・合計（分）・本年使用の調整（分）を書き込む。
'   ［申請・報告シートを作成］… 職員名簿の全員について、まだない人の申請・報告シートをひな形から作る（あれば作らない・上書きしない）
'   ［申請・報告シートを更新］… 作成先にある全員の申請・報告シートの「管理」シートを、休暇等シートの今の値に書き換える
'   ［様式を移行］…… 様式の古い申請・報告シートを、書いてある申請の行と超過勤務等命令簿の月のシート・フレックスの申告と割振りを残して新しいひな形の様式にする
'   ［超過勤務の月を追加］… 作成先にある全員の申請・報告シートに、B7 の月の超過勤務等命令簿のシート（超勤 R8.12 など）を作る
'   ［申告・割振り簿を見る］… B9 の出勤用ID の人の申請・報告シートを読み取り専用で開き、管理者だけが見る「申告・割振り簿」を表示する
'   ［古い超過勤務の月を外す］… 全員の申請・報告シートから、B10 の月より前の超過勤務等命令簿の月のシートのうち、確認済み・転記済み・PDF 保存済みのものを外す

Private Const DAY_MIN As Double = 465     ' 年次休暇の1日＝7時間45分＝465分
Private Const STAFF_LAST As Long = 501    ' 出勤・休暇管理データ の職員名簿の最後の行
Private Const LEAVE_LAST As Long = 1001   ' 出勤・休暇管理データ の休暇等（計算用シート）の最後の行
' 超過勤務等命令簿の月のシート：本人・所属長が書く欄（様式を移行するとき、月のシートの値をこの範囲だけ写す）と、ひな形のシートの名前
Private Const OT_INPUTS As String = "A5:B66,D5:I66,K5:L66,Y5:AA66,AC5:AE66,L69:O71,X69:Z71,AC69:AE71,BH1"
Private Const OT_TEMPLATE As String = "超勤ひな形"
' フレックスタイム制のシート：本人・所属長が書く欄（様式を移行するとき、式でないセルを空欄も含めて写す。
' 既定の時刻は式なので、本人が消した欄（勤務しない日）は空欄のまま写す）
Private Const FLEX_SIN As String = "フレックス（申告の内容）"     ' 日ごとの勤務（2026-11-19 から。前は フレックス申告）
Private Const FLEX_PER As String = "フレックス（単位期間・承認）"  ' 単位期間の表・承認・割振年月日（2026-11-19 から）
Private Const FLEX_PER_INPUTS As String = "C3,C5,C6,B9:B22,F9:H22"   ' 所属・所属長氏名・異動後・開始日・申告年月日・承認・割振年月日
Private Const FLEX_SIN_OLD As String = "フレックス申告"       ' 2026-11-19 までのシート名（移行のときに読む）
Private Const FLEX_DAYS As String = "H12:Z403"          ' 日ごとの欄（12 行目が日ごとの表の最初の日）
Private Const FLEX_FIRST_DAY As String = "#B12"         ' 読んだ様式の日ごとの表の最初の日（日ごとの欄をずらすのに使う）
Private Const FLEX_WARI As String = "フレックス割振り（所属長確認用）"   ' 2026-11-19 までのシート（移行のときに読む）
Private Const FLEX_WARI_OLD As String = "フレックス割振り"      ' 2026-11-13 までのシート名（移行のときに読む）
Private Const FLEX_BOOK As String = "申告・割振り簿"            ' 管理者だけが見る完成版（veryHidden）
' テレワーク実施報告：本人・所属長が書く欄（同じく様式を移行するときに写す）
Private Const TW_REPORT As String = "テレワーク報告"
Private Const TW_REPORT_INPUTS As String = "B3,B7:I206"      ' B3 所属（1つの欄）

' 出勤・休暇管理データ から読んだ値
Private staff As Variant      ' 職員名簿 A:B（出勤用ID・氏名）
Private keys As Variant       ' 計算用 H（出勤用ID|年）
Private totals As Variant     ' 計算用 K（合計・分）
Private adjust As Variant     ' 休暇等 H:I（本年使用（手入力）：日・時間）
Private holidays As Variant   ' 祝日 A（申請・報告シートの日数の計算で除く日）
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
' 申請・報告シートの非表示の「超勤ひな形」シート（tools/overtime/add_overtime_sheet.ps1 が入れる）を写して「超勤 R8.12」のような月のシートを作る。
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
        If MsgBox("作成先の全員の申請・報告シートに、超過勤務等命令簿の「" & OvertimeSheetName(ym) & "」のシートを作ります。" & vbCrLf & _
                  "職員や所属長が申請・報告シートを開いていない時間に実行してください。" & vbCrLf & vbCrLf & _
                  "実行しますか？", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    End If
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

    Dim names As New Collection, f As String, item As Variant
    f = Dir(outDir & "*.xlsx")     ' 申請・報告シート（今の名前）と 休暇申請シート・休暇申請_（前の名前）を IsSheetFile で選ぶ
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

' 申請・報告シート1つに月のシートを作る。結果の文を返す
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

' ===== 超過勤務等命令簿：古い月のシートを外す =====
' 申請・報告シートが月ごとに重くならないように、B10 の月より前の月のシートを外す。外すのは次の3つがそろった月だけ：
'   ・所属長の確認日（AC69）がある
'   ・出勤・休暇管理データ の「超勤①転記・月計」シート（表 転記_月計）に写してある（［超過勤務を転記］済み）
'   ・出勤・休暇管理データ の「超勤②PDF保存」シートに PDF の記録があり、その内容の記録が今のシートと同じ（PDF のあとで直されていない）
' 外す月の翌月のシートの BH1（前月の最終日の退庁時刻を指す式）は、外す前に値にする。外したシートは元に戻せない（PDF が保管用）。

Public Sub RemoveOldOvertimeMonths()
    Dim cfg As Worksheet, res As Worksheet
    Set cfg = ThisWorkbook.Worksheets("設定")
    Set res = ThisWorkbook.Worksheets("結果")
    Dim matchPath As String, outDir As String, cutoff As Date, latest As Date, problem As String
    matchPath = CleanPath(cfg.Range("B3").Value)
    outDir = CleanPath(cfg.Range("B5").Value)
    latest = DateSerial(Year(Date), Month(Date) - 2, 1)
    If Not FileExists(matchPath) Then problem = "出勤・休暇管理データ のファイルが見つかりません（B3）。"
    If problem = "" And (outDir = "" Or Dir(outDir, vbDirectory) = "") Then problem = "作成先のフォルダが見つかりません（B5）。"
    If problem = "" And Not IsDate(cfg.Range("B10").Value) Then problem = "外す月の境目を入れてください（B10。例 2025/4。この月より前の月を外します）。"
    If problem = "" Then
        cutoff = DateSerial(Year(cfg.Range("B10").Value), Month(cfg.Range("B10").Value), 1)
        If cutoff > latest Then problem = "B10 は " & Format(latest, "yyyy/m") & " より前にしてください（今月と前月のシートは外せません）。"
    End If
    res.Range("A2:E" & res.Rows.Count).ClearContents
    If problem <> "" Then
        res.Range("D2").Value = problem
        MsgBox problem, vbExclamation
        Exit Sub
    End If
    If Right$(outDir, 1) <> "\" Then outDir = outDir & "\"
    If MsgBox("全員の申請・報告シートから、「" & OvertimeSheetName(cutoff) & "」より前の超過勤務等命令簿の月のシートを外します。" & vbCrLf & vbCrLf & _
              "外すのは、所属長の確認済みで、出勤・休暇管理データ に転記済み、PDF で保存済み（そのあと直されていない）の月だけです。" & vbCrLf & _
              "外したシートは元に戻せません（保管用の PDF が残ります）。" & vbCrLf & _
              "先に 出勤・休暇管理データ で［命令簿をPDFで保存］を押して、上書き保存しておいてください。" & vbCrLf & _
              "職員や所属長が申請・報告シートを開いていない時間に実行してください。" & vbCrLf & vbCrLf & _
              "実行しますか？", vbYesNo + vbQuestion + vbDefaultButton2) <> vbYes Then Exit Sub

    Dim pdfs As Object, moved As Object
    If Not ReadOvertimeRecords(matchPath, pdfs, moved) Then Exit Sub
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

    Dim names As New Collection, f As String, item As Variant
    f = Dir(outDir & "*.xlsx")     ' 申請・報告シート（今の名前）と 休暇申請シート・休暇申請_（前の名前）を IsSheetFile で選ぶ
    Do While f <> ""
        If IsSheetFile(f) Then names.Add f
        f = Dir()
    Loop
    Dim row As Long, note As String, total As Long, parts() As String, id As String, nm As String
    row = 2
    For Each item In names
        parts = NameParts(CStr(item))
        id = parts(1)
        nm = ""
        If UBound(parts) >= 2 Then nm = parts(2)
        note = RemoveOldOne(outDir & item, id, cutoff, pdfs, moved, total)
        WriteResult res, row, id, nm, CStr(item), note
        row = row + 1
    Next item
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    res.Activate
    MsgBox "外した月のシートは全部で " & total & " 枚です。残した月とその理由は、結果シートで確かめてください。", vbInformation
End Sub

' 出勤・休暇管理データ のシート（2026-11-20 に名前を付け直した。前の名前の 出勤・休暇管理データ も読む）
Private Function AdminSheet(src As Workbook, newName As String, oldName As String) As Worksheet
    On Error Resume Next
    Set AdminSheet = src.Worksheets(newName)
    If AdminSheet Is Nothing Then Set AdminSheet = src.Worksheets(oldName)
    On Error GoTo 0
    If AdminSheet Is Nothing Then Err.Raise vbObjectError + 5, , "出勤・休暇管理データ に「" & newName & "」シートがありません"
End Function

' 出勤・休暇管理データ から、PDF の記録（出勤用ID|年月 → いちばん新しい内容の記録）と、転記済みの 出勤用ID|年月 を読む
Private Function ReadOvertimeRecords(matchPath As String, pdfs As Object, moved As Object) As Boolean
    Dim src As Workbook, openedHere As Boolean, ws As Worksheet, lo As ListObject, r As Long, last As Long
    Set pdfs = CreateObject("Scripting.Dictionary")
    Set moved = CreateObject("Scripting.Dictionary")
    Set src = FindOpenWorkbook(matchPath)
    If src Is Nothing Then
        Set src = Workbooks.Open(Filename:=matchPath, UpdateLinks:=0, ReadOnly:=True)
        openedHere = True
    End If
    On Error Resume Next
    Set ws = AdminSheet(src, "超勤②PDF保存", "命令簿の保存")
    Set lo = AdminSheet(src, "超勤①転記・月計", "超過勤務").ListObjects("転記_月計")
    On Error GoTo 0
    If ws Is Nothing Or lo Is Nothing Then
        MsgBox "出勤・休暇管理データ に「超勤②PDF保存」シートか、超過勤務シートの表がありません（新しい 出勤・休暇管理データ にしてください）。", vbExclamation
    Else
        last = ws.Cells(ws.Rows.Count, 2).End(xlUp).row
        For r = 6 To last      ' 下の行ほど新しい記録
            If CStr(ws.Cells(r, 2).Value) <> "" Then pdfs(NormId(CStr(ws.Cells(r, 3).Value)) & "|" & YmKey(ws.Cells(r, 2).Value)) = CStr(ws.Cells(r, 7).Value)
        Next r
        If Not lo.DataBodyRange Is Nothing Then
            For r = 1 To lo.ListRows.Count
                moved(NormId(CStr(lo.ListColumns("出勤用ID").DataBodyRange.Cells(r).Value)) & "|" & _
                      YmKey(lo.ListColumns("年月").DataBodyRange.Cells(r).Value)) = True
            Next r
        End If
        ReadOvertimeRecords = True
    End If
    If openedHere Then src.Close SaveChanges:=False
End Function

' 年月の値（"2026-11" の文字、または日付になったもの）→ "2026-11"
Private Function YmKey(v As Variant) As String
    If VarType(v) = vbDate Or (IsNumeric(v) And CStr(v) <> "") Then
        YmKey = Format(CDate(v), "yyyy-mm")
    Else
        YmKey = Trim$(CStr(v))
    End If
End Function

' 申請・報告シート1つから、外せる古い月のシートを外す。結果の文を返す（外した枚数を total に足す）
Private Function RemoveOldOne(path As String, id As String, cutoff As Date, pdfs As Object, moved As Object, total As Long) As String
    On Error GoTo Failed
    Dim wb As Workbook, s As Worksheet, ym As Date, k As String, removed As String, kept As String, n As Long
    Dim targets As New Collection, item As Variant, nxt As Worksheet
    Set wb = Workbooks.Open(Filename:=path, UpdateLinks:=0)
    For Each s In wb.Worksheets
        ym = OvertimeMonth(s.Name)
        If ym > 0 And ym < cutoff Then
            k = NormId(id) & "|" & Format(ym, "yyyy-mm")
            If Not IsNumeric(s.Range("AC69").Value2) Or CStr(s.Range("AC69").Value2) = "" Then
                kept = kept & "、" & Mid$(s.Name, 4) & "（確認待ち）"
            ElseIf Not moved.Exists(k) Then
                kept = kept & "、" & Mid$(s.Name, 4) & "（転記していない）"
            ElseIf Not pdfs.Exists(k) Then
                kept = kept & "、" & Mid$(s.Name, 4) & "（PDF がない）"
            ElseIf CStr(pdfs(k)) <> SheetHash(s) Then
                kept = kept & "、" & Mid$(s.Name, 4) & "（PDF のあとで直されている）"
            Else
                targets.Add s.Name
            End If
        End If
    Next s
    For Each item In targets     ' シートは月の順に並んでいる
        ' 翌月のシートの BH1（この月の最終日の退庁時刻を指す式）を値にしてから外す
        Set nxt = SheetByName(wb, OvertimeSheetName(DateAdd("m", 1, OvertimeMonth(CStr(item)))))
        If Not nxt Is Nothing Then
            If nxt.Range("BH1").HasFormula Then
                nxt.Unprotect
                nxt.Range("BH1").Value = nxt.Range("BH1").Value
                ProtectOvertime nxt
            End If
        End If
        wb.Worksheets(CStr(item)).Delete
        removed = removed & "、" & Mid$(CStr(item), 4)
        n = n + 1
    Next item
    If n > 0 Then wb.Close SaveChanges:=True Else wb.Close SaveChanges:=False
    total = total + n
    If n > 0 Then RemoveOldOne = "外しました：" & Mid$(removed, 2) Else RemoveOldOne = "外す月はありません"
    If kept <> "" Then RemoveOldOne = RemoveOldOne & "／残した月：" & Mid$(kept, 2)
    Exit Function
Failed:
    RemoveOldOne = "外せませんでした：" & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Function

' 月のシートの内容（A1:AF72 の値）から作る短い文字（出勤・休暇管理データのマクロの SheetHash と同じ計算。PDF のあとで直されたかを見る）
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

Private Sub ProtectOvertime(ws As Worksheet)
    ' 書式（行の高さ・列の幅）の変更はできる。add_overtime_sheet.ps1 と同じ設定
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=False, _
               AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
End Sub

' 超過勤務の月のシートと超勤ひな形の BH2（その月の最終日の退庁時刻。翌月の BH1 が指す）の式を、新しい式にする。
' 2026-11-19 までのひな形は、空欄の日を "" にしていたので、翌月の1日の欄が #VALUE! になった
Private Sub FixOvertimeBh2(wb As Workbook)
    Dim s As Worksheet, f As String
    f = "=IF($W$1="""",0,N(INDEX($L$5:$L$66,2*DAY(EOMONTH($W$1,0))-1)))"     ' tools/overtime/add_overtime_sheet.ps1 と同じ式
    For Each s In wb.Worksheets
        If OvertimeMonth(s.Name) > 0 Or s.Name = OT_TEMPLATE Then
            If s.Range("BH2").Formula <> f Then
                s.Unprotect
                s.Range("BH2").Formula = f
                ProtectOvertime s
            End If
        End If
    Next s
End Sub

' 超過勤務のシートで、確認日（AC69）が入っている月の、職員名簿・フレックスから引く欄（所属部課・勤務時間管理員・所属長の所属と氏名・登庁と退庁の標準時刻）の式を値にする
Private Sub FreezeBossNames(wb As Workbook)
    Dim s As Worksheet
    For Each s In wb.Worksheets
        If OvertimeMonth(s.Name) > 0 Then
            If IsNumeric(s.Range("AC69").Value2) And CStr(s.Range("AC69").Value2) <> "" Then
                Dim ref As Variant
                s.Unprotect
                For Each ref In Array("L69", "X70", "X71", "AC70", "AC71")
                    If s.Range(ref).HasFormula Then s.Range(ref).Value = s.Range(ref).Value
                Next ref
                ' 登庁・退庁の標準時刻（A・B 列。フレックスの申告から引く式）
                Dim cell As Range
                For Each cell In s.Range("A5:B66").Cells
                    If cell.HasFormula Then cell.Value = cell.Value
                Next cell
                ProtectOvertime s
            End If
        End If
    Next s
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

' ブックの中の表を名前で探す（なければエラー）
Private Function TableByName(wb As Workbook, nm As String) As ListObject
    Dim s As Worksheet, lo As ListObject
    For Each s In wb.Worksheets
        For Each lo In s.ListObjects
            If lo.Name = nm Then Set TableByName = lo: Exit Function
        Next lo
    Next s
    Err.Raise vbObjectError + 3, , "表「" & nm & "」が見つかりません"
End Function

' ［申告・割振り簿を見る］：B9 の出勤用ID の人の申請・報告シートを読み取り専用で開き、管理者だけが見る「申告・割振り簿」を表示する。
' 開いたファイルは保存しないで閉じる（読み取り専用なので、元のファイルは変わらない。本人・所属長には見えないまま）
Public Sub ShowFlexBook()
    Dim cfg As Worksheet, outDir As String, id As String, f As String, found As String, wb As Workbook, ws As Worksheet
    Set cfg = ThisWorkbook.Worksheets("設定")
    outDir = CleanPath(cfg.Range("B5").Value)
    id = Trim$(CStr(cfg.Range("B9").Value))
    If outDir = "" Or Dir(outDir, vbDirectory) = "" Then MsgBox "作成先のフォルダが見つかりません（B5）。", vbExclamation: Exit Sub
    If id = "" Then MsgBox "見る人の出勤用ID を B9 に入れてください。", vbExclamation: Exit Sub
    If Right$(outDir, 1) <> "\" Then outDir = outDir & "\"
    f = Dir(outDir & "*.xlsx")     ' 申請・報告シート（今の名前）と 休暇申請シート・休暇申請_（前の名前）を IsSheetFile で選ぶ
    Do While f <> ""
        If IsSheetFile(f) Then
            If NormId(NameParts(f)(1)) = NormId(id) Then found = f
        End If
        f = Dir()
    Loop
    If found = "" Then MsgBox "出勤用ID " & id & " の申請・報告シートが作成先にありません。", vbExclamation: Exit Sub
    Set wb = Workbooks.Open(Filename:=outDir & found, UpdateLinks:=0, ReadOnly:=True)
    Set ws = SheetByName(wb, FLEX_BOOK)
    If ws Is Nothing Then
        wb.Close SaveChanges:=False
        MsgBox "この申請・報告シートには「" & FLEX_BOOK & "」がありません（様式が古いので、［様式を移行］をしてください）。", vbExclamation
        Exit Sub
    End If
    ws.Visible = xlSheetVisible
    ws.Activate
    MsgBox found & " を読み取り専用で開き、「" & FLEX_BOOK & "」を表示しました。" & vbCrLf & _
           "F5〜I5 の単位期間を変えると、その期間が表示されます（保護を外す必要はありません）。" & vbCrLf & _
           "見終わったら、保存しないで閉じてください。", vbInformation
End Sub

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

' フレックスのシートとテレワーク報告の書いた欄を読む：(シート名, 番地, 式か値の2次元配列) の一覧。シートがなければ空。
' 2026-11-16 までの様式（単位期間の表が AJ11 から縦）は、新しい様式の右上の横の表の番地に並べ替えて読む
Private Function ReadFlexInputs(wb As Workbook) As Collection
    Dim result As New Collection, ws As Worksheet, wari As Worksheet, area As Range
    Set ws = SheetByName(wb, FLEX_PER)
    If Not ws Is Nothing Then
        ' 今の様式（2026-11-19 から）：フレックス（単位期間・承認）とフレックス（申告の内容）
        For Each area In ws.Range(FLEX_PER_INPUTS).Areas
            result.Add Array(FLEX_PER, area.Address(False, False), AreaInputs(area))
        Next area
        Set ws = SheetByName(wb, FLEX_SIN)
        If Not ws Is Nothing Then
            result.Add Array(FLEX_SIN, FLEX_FIRST_DAY, ws.Range("B12").Value2)
            result.Add Array(FLEX_SIN, FLEX_DAYS, AreaInputs(ws.Range(FLEX_DAYS)))
        End If
    Else
        ' 前の様式：フレックス申告（単位期間の表は右上に横。2026-11-16 までは右に縦で AJ11 が「期間」）と、フレックス割振り
        Set ws = SheetByName(wb, FLEX_SIN_OLD)
        If Not ws Is Nothing Then
            result.Add Array(FLEX_SIN, FLEX_FIRST_DAY, ws.Range("B12").Value2)
            result.Add Array(FLEX_PER, "C3", NoBlank(AreaInputs(ws.Range("D3"))))      ' 所属（空欄なら、職員名簿の部・課の式のまま）
            If CStr(ws.Range("AJ11").Value) = "期間" Then
                result.Add Array(FLEX_PER, "B9:B22", StartInputs(ws, ws.Range("AK12:AK25")))
                result.Add Array(FLEX_PER, "F9:F22", AreaInputs(ws.Range("AN12:AN25")))
            Else
                result.Add Array(FLEX_PER, "B9:B22", Across(AreaInputs(ws.Range("AK4:AX4"))))
                result.Add Array(FLEX_PER, "F9:F22", Across(AreaInputs(ws.Range("AK7:AX7"))))
            End If
            result.Add Array(FLEX_SIN, FLEX_DAYS, AreaInputs(ws.Range(FLEX_DAYS)))
        End If
        Set wari = SheetByName(wb, FLEX_WARI)
        If wari Is Nothing Then Set wari = SheetByName(wb, FLEX_WARI_OLD)     ' 2026-11-13 までのシート名
        If Not wari Is Nothing Then
            result.Add Array(FLEX_PER, "C5", NoBlank(AreaInputs(wari.Range("D3"))))   ' 所属長氏名（空欄なら、職員名簿の式のまま）
            result.Add Array(FLEX_PER, "C6", AreaInputs(wari.Range("D4")))
            If CStr(wari.Range("AJ11").Value) = "期間" Then
                result.Add Array(FLEX_PER, "G9:H22", AreaInputs(wari.Range("AO12:AP25")))
            Else
                result.Add Array(FLEX_PER, "G9:H22", Across(AreaInputs(wari.Range("AK8:AX9"))))
            End If
        End If
    End If
    ' テレワーク報告
    Set ws = SheetByName(wb, TW_REPORT)
    If Not ws Is Nothing Then
        If CStr(ws.Range("C3").Value) = "部" Then
            ' 2026-11-16 までの様式（B3 部・D3 課）：「○○部○○課」の1つの文にして所属の欄に写す
            result.Add Array(TW_REPORT, "B3", OneCell(OldSection(ws)))
            result.Add Array(TW_REPORT, "B7:I206", AreaInputs(ws.Range("B7:I206")))
        Else
            For Each area In ws.Range(TW_REPORT_INPUTS).Areas
                If area.Address(False, False) = "B3" Then
                    result.Add Array(TW_REPORT, "B3", NoBlank(AreaInputs(area)))     ' 所属（空欄なら、職員名簿の部・課の式のまま）
                Else
                    result.Add Array(TW_REPORT, area.Address(False, False), AreaInputs(area))
                End If
            Next area
        End If
    End If
    Set ReadFlexInputs = result
End Function

' 1つの欄の値で、空欄（""）なら書かない（Empty）。新しい様式の式（職員名簿から入る既定）を残すため
Private Function NoBlank(v As Variant) As Variant
    If CStr(v(1, 1)) = "" Then v(1, 1) = Empty
    NoBlank = v
End Function

' 前の様式のテレワーク報告の所属（B3 部・D3 課）を1つの文にする
Private Function OldSection(ws As Worksheet) As String
    Dim b As String, d As String
    b = Trim(CStr(ws.Range("B3").Value)): d = Trim(CStr(ws.Range("D3").Value))
    If b <> "" Then OldSection = b & "部"
    If d <> "" Then OldSection = OldSection & d & "課"
End Function

' 1つの値を 1×1 の2次元配列にする
Private Function OneCell(ByVal v As Variant) As Variant
    Dim out(1 To 1, 1 To 1) As Variant
    out(1, 1) = v
    OneCell = out
End Function

' 2次元配列の行と列を入れ替える（縦の単位期間の表 → 横の表）
Private Function Across(v As Variant) As Variant
    Dim out() As Variant, i As Long, j As Long
    ReDim out(1 To UBound(v, 2), 1 To UBound(v, 1))
    For i = 1 To UBound(v, 1)
        For j = 1 To UBound(v, 2)
            out(j, i) = v(i, j)
        Next j
    Next i
    Across = out
End Function

' 日ごとの欄を shift 日ぶんずらす（新しい様式の行 i に、読んだ様式の行 i + shift を入れる。外は Empty＝新しい様式の式のまま）
Private Function Shifted(v As Variant, ByVal shift As Long) As Variant
    Dim out() As Variant, i As Long, j As Long
    ReDim out(1 To UBound(v, 1), 1 To UBound(v, 2))
    For i = 1 To UBound(v, 1)
        If i + shift >= 1 And i + shift <= UBound(v, 1) Then
            For j = 1 To UBound(v, 2)
                out(i, j) = v(i + shift, j)
            Next j
        End If
    Next i
    Shifted = out
End Function

' 単位期間の開始日（AK12:AK25）：前の様式では式（決まった4週間ごと）だったので、申告年月日（AN）のある行は式でも値を写す
Private Function StartInputs(ws As Worksheet, area As Range) As Variant
    Dim out() As Variant, i As Long
    ReDim out(1 To area.Rows.Count, 1 To 1)
    For i = 1 To area.Rows.Count
        With area.Cells(i, 1)
            If .HasFormula Then
                If IsDate(ws.Range("AN" & .Row).Value) And IsNumeric(.Value2) Then out(i, 1) = .Value2 Else out(i, 1) = Empty
            ElseIf IsError(.Value2) Then
                out(i, 1) = Empty
            Else
                out(i, 1) = .Value2
                If IsEmpty(out(i, 1)) Then out(i, 1) = ""
            End If
        End With
    Next i
    StartInputs = out
End Function

' 範囲の各セル：式なら Empty（新しい様式の式を使う）、式でなければ値（空欄は ""）を、2次元配列で
Private Function AreaInputs(area As Range) As Variant
    Dim f As Variant, v As Variant, out() As Variant, i As Long, j As Long
    ReDim out(1 To area.Rows.Count, 1 To area.Columns.Count)
    For i = 1 To area.Rows.Count
        For j = 1 To area.Columns.Count
            With area.Cells(i, j)
                If .HasFormula Then
                    out(i, j) = Empty
                ElseIf IsError(.Value2) Then
                    out(i, j) = Empty
                Else
                    out(i, j) = .Value2
                    If IsEmpty(out(i, j)) Then out(i, j) = ""
                End If
            End With
        Next j
    Next i
    AreaInputs = out
End Function

' 読んでおいたフレックスの欄を新しい様式に書く（式でなかったセルだけ。空欄は空にする）。書いたセルの数を返す
Private Function WriteFlexInputs(wb As Workbook, items As Collection) As Long
    Dim item As Variant, ws As Worksheet, area As Range, f As Variant, i As Long, j As Long, vals As Variant, one As Variant
    Dim oldFirst As Variant
    For Each item In items
        Set ws = SheetByName(wb, CStr(item(0)))
        If item(1) = FLEX_FIRST_DAY Then
            oldFirst = item(2)
        ElseIf Not ws Is Nothing Then
            ws.Unprotect
            Set area = ws.Range(CStr(item(1)))
            vals = item(2)
            ' 日ごとの表の最初の日（単位期間1の開始日）が読んだ様式と違えば、日ごとの欄を同じ日の行にずらす
            If item(0) = FLEX_SIN And item(1) = FLEX_DAYS And IsNumeric(oldFirst) And Not IsEmpty(oldFirst) Then
                ws.Calculate
                If IsNumeric(ws.Range("B12").Value2) Then vals = Shifted(vals, CLng(ws.Range("B12").Value2) - CLng(oldFirst))
            End If
            If area.Cells.Count = 1 Then
                If Not IsEmpty(vals(1, 1)) Then area.Value2 = vals(1, 1): WriteFlexInputs = WriteFlexInputs + 1
            Else
                f = area.Formula
                For i = 1 To area.Rows.Count
                    For j = 1 To area.Columns.Count
                        If Not IsEmpty(vals(i, j)) Then f(i, j) = vals(i, j): WriteFlexInputs = WriteFlexInputs + 1
                    Next j
                Next i
                area.Formula = f
            End If
            ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
        End If
    Next item
End Function

' 様式の古い申請・報告シートを新しい様式にする。元のファイルは控えのフォルダに複製してから、同じファイルに上書きする
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
        If MsgBox("様式の古い申請・報告シートを、新しいひな形の様式にします。" & vbCrLf & _
                  "職員や所属長が申請・報告シートを開いていない時間に実行してください。" & vbCrLf & _
                  "元のファイルは作成先の中の控えのフォルダに残します。" & vbCrLf & vbCrLf & _
                  "実行しますか？", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    ReadMatching matchPath

    ' 先にファイルの名前を集める（処理の途中で Dir の続きが使えなくなるため）
    Dim names As New Collection, f As String, item As Variant
    f = Dir(outDir & "*.xlsx")     ' 申請・報告シート（今の名前）と 休暇申請シート・休暇申請_（前の名前）を IsSheetFile で選ぶ
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

' 申請・報告シート1つを移行する。結果の文を返す
Private Function MigrateOne(path As String, fileName As String, tplPath As String, backupDir As String, _
                            id As String, nm As String, yr As Long) As String
    ' 本人・所属長が書く列（日数と時間は自動）
    Dim cols As Variant, c As Long, r As Long, n As Long, overwritten As Boolean
    cols = Array("申請日", "休暇の種類", "種別・内容", "単位", "開始日", "終了日", "開始時刻", "終了時刻", _
                 "事由", "承認（所属長）", "承認日", "所属長コメント", "基準日")
    Dim wb As Workbook, lo As ListObject, vals() As Variant, hasValue() As Boolean, otMonths As Collection, otCount As Long
    Dim flexItems As Collection, flexCount As Long
    On Error GoTo Failed

    ' 古いファイルから読む（式の入ったセルは、新しい様式の式に任せるので写さない）
    Set wb = Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True)
    If CStr(SheetValue(wb, "管理", "B11")) = formVersion Then
        wb.Close SaveChanges:=False
        Set wb = Nothing
        MigrateOne = "新しい様式なので何もしませんでした"
        Exit Function
    End If
    Set lo = TableByName(wb, "申請")     ' 表「申請」（シート名は 休暇申請。前は 申請）
    n = lo.ListRows.Count
    ReDim vals(1 To n, 0 To UBound(cols))
    ReDim hasValue(1 To n, 0 To UBound(cols))
    For c = 0 To UBound(cols)
        If HasColumn(lo, CStr(cols(c))) Then
            For r = 1 To n
                With lo.ListColumns(ColumnName(lo, CStr(cols(c)))).DataBodyRange.Cells(r)
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
    Set flexItems = ReadFlexInputs(wb)         ' フレックスの申告・割振りの書いた欄
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
    Set lo = TableByName(wb, "申請")
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
                        lo.ListColumns(ColumnName(lo, CStr(cols(c)))).DataBodyRange.Cells(target).Value2 = text
                    Else
                        lo.ListColumns(ColumnName(lo, CStr(cols(c)))).DataBodyRange.Cells(target).Value2 = vals(r, c)
                    End If
                End If
            Next c
        End If
    Next r
    rowCount = target
    otCount = WriteOvertimeInputs(wb, otMonths)
    flexCount = WriteFlexInputs(wb, flexItems)
    Application.Calculation = xlCalculationAutomatic
    wb.Close SaveChanges:=True
    Set wb = Nothing
    MigrateOne = "移行しました（申請 " & rowCount & " 行、超過勤務の月 " & otCount & IIf(flexItems.Count > 0, "、フレックス・テレワーク報告の欄 " & flexCount, "") & "）。元のファイルの控え：" & backupDir & fileName
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

' 個人別のファイル名：申請・報告シート　出勤用ID　氏名.xlsx
Private Function StandardName(id As String, nm As String) As String
    StandardName = "申請・報告シート　" & SafeName(id) & "　" & SafeName(nm) & ".xlsx"
End Function

' 個人別の申請・報告シートのファイルか（ひな形は違う）。区切りは全角スペース（半角スペース・「_」も可）。前の名前（休暇申請シート　…、休暇申請_…）も選ぶ
Private Function IsSheetFile(f As String) As Boolean
    IsSheetFile = (f Like "申請・報告シート[　 _]*.xlsx") Or (f Like "休暇申請シート[　 _]*.xlsx") Or (f Like "休暇申請_*.xlsx")
End Function

' ファイル名を 申請・報告シート / 出勤用ID / 氏名 に分ける
Private Function NameParts(f As String) As String()
    Dim s As String
    s = Replace(Replace(Left$(f, Len(f) - 5), "_", "　"), " ", "　")
    Do While InStr(s, "　　") > 0
        s = Replace(s, "　　", "　")
    Loop
    NameParts = Split(s, "　")
End Function

' 表の列があるか（見出しの改行は除いて比べる：「承認」＋改行＋「（所属長）」と「承認（所属長）」は同じ列）
Private Function HasColumn(lo As ListObject, name As String) As Boolean
    HasColumn = (ColumnName(lo, name) <> "")
End Function

' 表の列の、本当の見出し（改行も含む）。見出しの改行を除いて name と同じ列。なければ ""
Private Function ColumnName(lo As ListObject, name As String) As String
    Dim col As ListColumn
    For Each col In lo.ListColumns
        If Replace(Replace(col.Name, vbLf, ""), vbCr, "") = Replace(Replace(name, vbLf, ""), vbCr, "") Then ColumnName = col.Name: Exit Function
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

    ' 作成先にある申請・報告シート（出勤用ID → ファイル名）。出勤用ID はファイル名の2つ目の部分
    Dim existing As Object, f As String
    Set existing = CreateObject("Scripting.Dictionary")
    f = Dir(outDir & "*.xlsx")     ' 申請・報告シート（今の名前）と 休暇申請シート・休暇申請_（前の名前）を IsSheetFile で選ぶ
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
            parts = NameParts(fileName)     ' 申請・報告シート / 出勤用ID / 氏名
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
               "次に、作った申請・報告シートを OneDrive で本人と所属長に共有してください。", vbInformation
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
    staff = AdminSheet(src, "台帳①職員名簿", "職員名簿").Range("A2:O" & STAFF_LAST).Value2     ' C・D は休憩の始め・終わり、J・K は養育する子の数・要介護者の数、L 所属長氏名、M 所属長の所属、N・O 勤務時間管理員の所属・氏名
    ' C・D は休憩の始め・終わり（短縮の特例の人だけ）
    keys = src.Worksheets("計算用").Range("H2:H" & LEAVE_LAST).Value
    totals = src.Worksheets("計算用").Range("K2:K" & LEAVE_LAST).Value
    adjust = AdminSheet(src, "休暇②個人別一覧", "休暇等").Range("H2:I" & LEAVE_LAST).Value     ' B 列は氏名（自動）
    holidays = AdminSheet(src, "台帳②祝日", "祝日").Range("A2:A200").Value
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

' 申請・報告シートを開いて「管理」シートに書き込む。結果の文を返す
Private Function UpdateOne(path As String, id As String, nm As String, yr As Long) As String
    On Error GoTo Failed
    Dim pos As Long, hold As Variant, adj As Variant, carry As Variant, grant As Variant, wb As Workbook, ws As Worksheet
    pos = FindRow(NormId(id) & "|" & yr)
    If pos > 0 Then
        If IsNumeric(totals(pos, 1)) And CStr(totals(pos, 1)) <> "" Then hold = totals(pos, 1)
        If IsNumeric(carries(pos, 1)) And CStr(carries(pos, 1)) <> "" Then carry = carries(pos, 1)
        If IsNumeric(grants(pos, 1)) And CStr(grants(pos, 1)) <> "" Then grant = grants(pos, 1) * DAY_MIN
        adj = Round(Val(adjust(pos, 1)) * DAY_MIN + Val(adjust(pos, 2)) * 1440, 0)
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
    Dim breakStart As Double, breakEnd As Double, k As Long, children As Variant, carees As Variant, boss As String, section As String
    Dim bossSection As String, clerkSection As String, clerkName As String
    breakStart = 12 / 24
    breakEnd = 13 / 24
    children = Empty: carees = Empty
    For k = 1 To UBound(staff, 1)
        If NormId(CStr(staff(k, 1))) = NormId(id) Then
            ' 養育する子の数・要介護者の数（看護休暇・短期介護の上限に使う。空欄は空欄のまま）
            If CStr(staff(k, 10)) <> "" And IsNumeric(staff(k, 10)) Then children = CLng(staff(k, 10))
            If CStr(staff(k, 11)) <> "" And IsNumeric(staff(k, 11)) Then carees = CLng(staff(k, 11))
            If Not IsError(staff(k, 12)) Then boss = Trim$(CStr(staff(k, 12)))
            If Not IsError(staff(k, 13)) Then bossSection = Trim$(CStr(staff(k, 13)))
            If Not IsError(staff(k, 14)) Then clerkSection = Trim$(CStr(staff(k, 14)))
            If Not IsError(staff(k, 15)) Then clerkName = Trim$(CStr(staff(k, 15)))
            ' 所属：部・課（E・F。部が空欄なら課だけ）。フレックス（単位期間・承認）の所属の既定
            If Not IsError(staff(k, 5)) And Not IsError(staff(k, 6)) Then section = Trim$(Trim$(CStr(staff(k, 5))) & " " & Trim$(CStr(staff(k, 6))))
            If IsNumeric(staff(k, 3)) And IsNumeric(staff(k, 4)) And CStr(staff(k, 3)) <> "" And CStr(staff(k, 4)) <> "" Then
                breakStart = staff(k, 3)
                breakEnd = staff(k, 4)
            End If
            Exit For
        End If
    Next k
    ws.Range("B13").Value = children
    ws.Range("B14").Value = carees
    ' 所属長氏名・所属・勤務時間管理員などは、超過勤務のシートの最下部の欄が引く。変わる前に、確認日のある月の欄を値にして残す（異動しても前の名前のまま）
    ' 確認済みの月は、名簿やフレックスの申告があとで変わっても変わらないように、引いている欄を値にする（更新のたびに）
    FreezeBossNames wb
    FixOvertimeBh2 wb
    ws.Range("B15").Value = boss
    ws.Range("B16").Value = section
    ws.Range("B17").Value = bossSection
    ws.Range("B18").Value = clerkSection
    ws.Range("B19").Value = clerkName
    ws.Range("B9").Value = breakStart
    ws.Range("B10").Value = breakEnd
    ws.Range("B7").Value = carry
    ws.Range("B8").Value = grant
    ' 祝日（日付だけを書く。文字があると申請・報告シートの日数の計算がエラーになるため）
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
        UpdateOne = "更新しました（繰越分 " & DaysText(carry) & "、本年分 " & DaysText(grant) & "、合計 " & DaysText(hold) & "、調整 " & DaysText(adj) & "）"
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
