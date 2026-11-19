"""申請・報告シートに、テレワーク実施報告のシートを足す（tools/leave_sheet/make_form.py から呼ぶ）。

シート：
  テレワーク報告 … 本人が1回1行で書く（提出日・実施日・開始時刻・終了時刻・業務内容・備考）。所属長は同じ行に確認日（とコメント）を入れる。
                   表「テレワーク報告」を 出勤・休暇管理データ が Power Query で読む。
  テレワーク実施報告書 … 別紙３の様式（A4縦・游明朝）どおりの印刷用。M2 の番号の行を表示する（式なので保護し、M2 だけ書ける）。
                         改ページプレビューで印刷の形のまま見せる。入る欄は薄い色（白黒印刷にするので印刷はされない）。
"""
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.page import PageMargins

SHEET, PRINT = "テレワーク報告", "テレワーク報告書（保存用）"     # 印刷用のシート（2026-11-19 までは「テレワーク実施報告書」）
NOWORK_OVERLAP = "・フレックスで割振り済みの「勤務時間を割り振らない日」です（先に確認・割振りされた方が優先）"
# 休暇との重なりは、テレワーク予定表の予定と照らす（出勤・休暇管理データ の「テレワーク一覧」）。この報告では見ない
HEAD, ROWS = 6, 200
FIRST, LAST = HEAD + 1, HEAD + ROWS

FONT = "ＭＳ Ｐゴシック"
f_base = Font(name=FONT, size=10)
f_bold = Font(name=FONT, size=10, bold=True)
f_head = Font(name=FONT, size=10, bold=True, color="FFFFFF")
f_title = Font(name=FONT, size=14, bold=True, color="1E3A5F")
f_note = Font(name=FONT, size=9, color="555555")
fill_head = PatternFill("solid", fgColor="1E3A5F")
fill_staff = PatternFill("solid", fgColor="FFF2CC")
fill_boss = PatternFill("solid", fgColor="DDEBF7")
fill_calc = PatternFill("solid", fgColor="F2F2F2")
fill_bad = PatternFill("solid", fgColor="F4B6B0", bgColor="F4B6B0")
font_bad = Font(name=FONT, color="9C0006", bold=True)
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left = Alignment(horizontal="left", vertical="center", wrap_text=True)
top_left = Alignment(horizontal="left", vertical="top", wrap_text=True)
UNLOCK = Protection(locked=False)

# 表の列：(見出し, 幅, 書式, 塗り)
COLS = [
    ("番号", 5, "0", fill_calc),
    ("提出日", 11, "yyyy/m/d", fill_staff),
    ("実施日", 13, "yyyy/m/d", fill_staff),      # 2026/12/31 のような8桁の日付が入る幅
    ("開始時刻", 8, "h:mm", fill_staff),
    ("終了時刻", 8, "h:mm", fill_staff),
    ("業務内容", 50, None, fill_staff),
    ("備考", 24, None, fill_staff),
    ("確認日（所属長）", 13, "yyyy/m/d", fill_boss),   # 8桁の日付が入る幅
    ("所属長コメント", 20, None, fill_boss),
    ("確認", 30, None, fill_calc),
]


def add(wb, MG, TIME_RANGE, OFF_DAYS, LEAVE, NOWORK=None):
    """MG は管理シートの名前、TIME_RANGE は時刻の一覧（15分おき）の範囲、OFF_DAYS は祝日・年末年始の範囲（土日とあわせて提出日・実施日・確認日にできない）"""
    ws = wb.create_sheet(SHEET)
    for i, (_, width, _, _) in enumerate(COLS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws["A1"] = "テレワーク実施報告（本人が入力）"
    ws["A1"].font = f_title
    ws["F1"] = "黄色：本人が入力　青：所属長が入力　灰色：自動　赤：直すところ（理由は「確認」の欄）"
    ws["F1"].font = f_note
    # 3行目：所属（B3:E3 を結合した1つの欄。職員名簿の部・課が自動で入る（管理シート B16）。違うときは自由に書き直せる）と氏名
    ws["A3"] = "所属"
    ws["A3"].font, ws["A3"].alignment = f_bold, center
    ws.merge_cells("B3:E3")
    for col in "BCDE":
        ws[f"{col}3"].fill, ws[f"{col}3"].border = fill_staff, border
    ws["B3"] = f'={MG}!$B$16&""'
    ws["B3"].protection = UNLOCK
    ws["B3"].font, ws["B3"].alignment = f_bold, Alignment(horizontal="left", vertical="center", shrink_to_fit=True)
    ws["F3"] = f"=\"氏名：\"&{MG}!$B$2"
    ws["F3"].font, ws["F3"].alignment = f_bold, left
    ws["A4"] = ("テレワークをしたら、1回1行で書き、所属長に Notes でこのファイルのリンクを送って知らせます。"
                "所属長は内容を確かめて確認日を入れます。印刷は「テレワーク報告書（保存用）」シート（別紙３の様式）で。")
    ws["A4"].font = f_note

    for i, (label, _, fmt, fill) in enumerate(COLS, start=1):
        h = ws.cell(row=HEAD, column=i, value=label)
        h.font, h.alignment, h.border, h.fill = f_head, center, border, fill_head
        for r in range(FIRST, LAST + 1):
            c = ws.cell(row=r, column=i)
            c.font, c.border, c.fill = f_base, border, fill
            c.alignment = left if label in ("業務内容", "備考", "所属長コメント", "確認") else center
            if fmt:
                c.number_format = fmt
            if fill is not fill_calc:
                c.protection = UNLOCK
    ws.row_dimensions[HEAD].height = 28
    for r in range(FIRST, LAST + 1):      # 入力の行は3行ぶんの高さ（業務内容などを折り返して書ける）
        ws.row_dimensions[r].height = 45
    # フレックスで割振り済みの「勤務時間を割り振らない日」（NOWORK＝(割振年月日の範囲, 日付の範囲)）
    nowork_hit = lambda c, h: (f'SUMPRODUCT(ISNUMBER({NOWORK[0]})*({NOWORK[1]}={c})*((((1-ISNUMBER({h}))+({NOWORK[0]}<=N({h})))>0)))')
    for r in range(FIRST, LAST + 1):
        a, b, c, d, e, f, g, h = (f"{col}{r}" for col in "ABCDEFGH")
        ws[a] = r - HEAD
        ws[f"J{r}"] = (
            f'=IF(COUNTA({b}:{g})=0,IF({h}<>"","・報告がないのに確認日があります",""),'
            f'IF({b}="","・提出日を入れてください","")'
            f'&IF({c}="","・実施日を入れてください","")'
            f'&IF(OR({d}="",{e}=""),"・開始時刻と終了時刻を入れてください","")'
            f'&IF(AND(ISNUMBER({d}),ISNUMBER({e})),IF({e}<={d},"・終了時刻が開始時刻より前です",""),"")'
            f'&IF({f}="","・業務内容を書いてください","")'
            f'&IF(AND(ISNUMBER({b}),ISNUMBER({c})),IF({c}>{b},"・実施日が提出日より後です",""),"")'
            f'&IF(AND(ISNUMBER({h}),ISNUMBER({b})),IF({h}<{b},"・確認日が提出日より前です（提出日と同じ日か、その後にします）",""),"")'
            + "".join(f'&IF(ISNUMBER({x}),IF(NETWORKDAYS({x},{x},{OFF_DAYS})=0,"・{name}が土日・祝日・年末年始です",""),"")'
                      for x, name in ((b, "提出日"), (c, "実施日"), (h, "確認日")))
            + (f'&IF(ISNUMBER({c}),IF({nowork_hit(c, h)}>0,"{NOWORK_OVERLAP}",""),"")' if NOWORK else "")
            + ")"
        )
        ws[f"J{r}"].font = Font(name=FONT, size=9, color="9C0006")
    any_ = f"COUNTA($B{FIRST}:$G{FIRST})>0"
    off = lambda col: f"AND(ISNUMBER(${col}{FIRST}),NETWORKDAYS(${col}{FIRST},${col}{FIRST},{OFF_DAYS})=0)"   # 土日・祝日・年末年始
    rules = {
        "B": f'OR(AND({any_},$B{FIRST}=""),{off("B")})',
        "C": f'OR(AND({any_},$C{FIRST}=""),{off("C")},AND(ISNUMBER($B{FIRST}),ISNUMBER($C{FIRST}),$C{FIRST}>$B{FIRST}))',
        "D": f'AND({any_},$D{FIRST}="")',
        "E": f'OR(AND({any_},$E{FIRST}=""),AND(ISNUMBER($D{FIRST}),ISNUMBER($E{FIRST}),$E{FIRST}<=$D{FIRST}))',
        "F": f'AND({any_},$F{FIRST}="")',
        "H": f'OR(AND($H{FIRST}<>"",NOT({any_})),{off("H")},AND(ISNUMBER($H{FIRST}),ISNUMBER($B{FIRST}),$H{FIRST}<$B{FIRST}))',
        "J": f'$J{FIRST}<>""',
    }
    for col, rule in rules.items():
        ws.conditional_formatting.add(f"{col}{FIRST}:{col}{LAST}", FormulaRule(formula=[rule], fill=fill_bad, font=font_bad))
    t = Table(displayName="テレワーク報告", ref=f"A{HEAD}:{get_column_letter(len(COLS))}{LAST}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
    ws.add_table(t)
    for col in "BCH":
        dv = DataValidation(type="date", operator="greaterThan", formula1="43831", allow_blank=True, imeMode="halfAlpha")
        dv.errorTitle, dv.error = "日付を入れてください", "例：11/13 または 2026/11/13"
        ws.add_data_validation(dv)
        dv.add(f"{col}{FIRST}:{col}{LAST}")
    dv = DataValidation(type="list", allow_blank=True, imeMode="halfAlpha", formula1=TIME_RANGE)
    dv.showErrorMessage = False
    ws.add_data_validation(dv)
    dv.add(f"D{FIRST}:E{LAST}")
    ws.freeze_panes = f"A{FIRST}"
    ws.protection.sheet = True
    ws.protection.formatColumns = False
    ws.protection.formatRows = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{HEAD}:{HEAD}"

    # ===== テレワーク実施報告書（別紙３の様式。印刷用） =====
    ps = wb.create_sheet(PRINT)
    ps.sheet_view.showGridLines = False
    MIN = "游明朝"
    f11 = Font(name=MIN, size=12)     # 本文（2026-11-16 ユーザー指示で 11→12）
    f14 = Font(name=MIN, size=16)     # 表題（14→16）
    for i in range(1, 11):     # A〜J の10列で本文の幅（A4縦・左右25mm）
        ps.column_dimensions[get_column_letter(i)].width = 8.6
    ps.column_dimensions["K"].width = 3
    ps.column_dimensions["L"].width = 16
    ps.column_dimensions["M"].width = 8
    T = f"'{SHEET}'"
    NO = "$M$2"
    pick = lambda col: f"INDEX({T}!${col}${FIRST}:${col}${LAST},{NO})"
    # 印刷する番号（印刷範囲の外）。既定は書いてある最後の行
    ps["L2"] = "印刷する番号"
    ps["L2"].font = f_bold
    ps["M2"] = f'=IFERROR(LOOKUP(2,1/({T}!$C${FIRST}:$C${LAST}<>""),{T}!$A${FIRST}:$A${LAST}),1)'
    ps["M2"].fill, ps["M2"].border, ps["M2"].font, ps["M2"].alignment = fill_staff, border, f_bold, center
    ps["M2"].protection = UNLOCK
    ps["L3"] = "← テレワーク報告シートの番号"
    ps["L3"].font = f_note
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2=str(ROWS), allow_blank=False, imeMode="halfAlpha")
    dv.errorTitle, dv.error = "番号", f"1〜{ROWS} の番号を入れてください。"
    ps.add_data_validation(dv)
    dv.add("M2")

    def line(r, value, font=f11, align=left, merge=("A", "J"), height=20):
        ps.merge_cells(f"{merge[0]}{r}:{merge[1]}{r}")
        c = ps[f"{merge[0]}{r}"]
        c.value, c.font, c.alignment = value, font, align
        ps.row_dimensions[r].height = height

    era = lambda d: f'IF(ISNUMBER({d}),TEXT({d},"ggge年m月d日"),"令和　　年　　月　　日")'
    hm = lambda t: f'IF(ISNUMBER({t}),TEXT({t},"h""時""mm""分"""),"　　時　　分")'
    line(1, None)
    line(2, f'="提出日："&{era(pick("B"))}', align=Alignment(horizontal="right", vertical="center"))
    line(3, None)
    line(4, "テレワーク実施報告書", font=f14, align=Alignment(horizontal="center", vertical="center"), height=28)
    line(5, None)
    # 所属・氏名は「：」の後を1字下げる
    line(6, f'="所 属：　"&IF({T}!$B$3="","　　　　　　",{T}!$B$3)')
    line(7, f'="氏 名：　"&{MG}!$B$2')
    line(8, None)
    line(9, "１　実施日時")
    line(10, f'="　　"&{era(pick("C"))}&"（"&IF(ISNUMBER({pick("C")}),TEXT({pick("C")},"aaa"),"　　")&"）"')
    line(11, f'="　　　　"&{hm(pick("D"))}&"　～　"&{hm(pick("E"))}')
    line(12, None)
    line(13, "２　業務内容")
    ps.merge_cells("B14:J31")
    ps["B14"] = f'={pick("F")}&""'
    ps["B14"].font, ps["B14"].alignment = f11, top_left
    for r in range(14, 32):      # 業務内容の欄（12pt で18行ほど。A4 1枚に縮めずに入る高さ）
        ps.row_dimensions[r].height = 16
    line(32, "３　備考")
    ps.merge_cells("B33:J36")
    ps["B33"] = f'={pick("G")}&""'
    ps["B33"].font, ps["B33"].alignment = f11, top_left
    for r in range(33, 37):
        ps.row_dimensions[r].height = 18
    ps.print_area = "A1:J36"
    ps.page_setup.paperSize = ps.PAPERSIZE_A4
    ps.page_setup.orientation = "portrait"
    ps.page_setup.fitToWidth = 1
    ps.page_setup.fitToHeight = 1
    ps.sheet_properties.pageSetUpPr.fitToPage = True
    ps.page_margins = PageMargins(left=0.98, right=0.98, top=0.98, bottom=0.98, header=0.24, footer=0.59)
    ps.oddHeader.left.text = "別紙３"
    ps.oddHeader.left.font = "游明朝"
    ps.oddHeader.left.size = 10.5
    # 記入欄（番号の行から入る欄）を薄い色に。白黒印刷にして、色は印刷しない
    field = PatternFill("solid", fgColor="FFF8E1")
    for rng in ("A2:J2", "A6:J6", "A7:J7", "A10:J10", "A11:J11", "B14:J31", "B33:J36"):
        for row in ps[rng]:
            for c in row:
                c.fill = field
    ps.page_setup.blackAndWhite = True
    # 改ページプレビューで、印刷の形のまま表示する。式を消さないよう保護する（M2 だけ書ける）
    ps.sheet_view.view = "pageBreakPreview"
    ps.sheet_view.zoomScale = 100
    ps.sheet_view.zoomScaleNormal = 100
    ps.sheet_view.zoomScalePageLayoutView = 100
    ps.sheet_view.zoomScaleSheetLayoutView = 100
    ps.protection.sheet = True
    return [SHEET, PRINT]
