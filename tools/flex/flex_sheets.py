"""申請・報告シートに、フレックスタイム制の申告・割振りのシートを足す（tools/leave_sheet/make_form.py から呼ぶ）。

シート：
  フレックス申告 … 本人が入力。1年間（始期から392日＝4週間×14の単位期間）の日ごとの勤務開始・終了・休憩（時と分を別の欄）と、
                   単位期間ごとの申告年月日。
  フレックス割振り（所属長確認用） … 所属長が確認する。申告をそのまま写した日ごとの表と、単位期間ごとの承認・割振年月日。
  フレックス表（非表示）… 出勤・休暇管理データ が Power Query で読む表「フレックス日別」「フレックス期間」（時刻・時間は分）。
様式どおりの完成版のシート「申告・割振り簿」は tools/flex/add_flex_book.ps1 が、総務課の様式の Excel から作って足す。

日ごとの行は、どのシートも 12 行目が始期の日（様式の申告・割振り簿と同じ行）。
"""
from copy import copy
from datetime import date
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.workbook.defined_name import DefinedName

FLEX_START = date(2026, 4, 1)   # 1年間の始期（管理シート B12。年度が変わったら変える）
DAYS, PERIOD = 392, 28          # 4週間×14
N_PERIODS = DAYS // PERIOD
FIRST = 12                      # 日ごとの行の最初（始期の日）
LAST = FIRST + DAYS - 1         # 403
TOTAL_MIN = 155 * 60            # 単位期間の勤務時間（4週間＝20日×7時間45分）
CORE = (10 * 60, 12 * 60)       # コアタイム 10:00〜12:00
WORK = (9 * 60, 17 * 60 + 45)   # 既定の勤務 9:00〜17:45
BREAK = (12 * 60, 13 * 60)      # 既定の休憩 12:00〜13:00
FLEXIBLE = (5 * 60, 22 * 60)    # 勤務を割り振れる時間帯 5:00〜22:00
EXTRA_BREAK = (5 * 60 + 30, 6 * 60 + 30, 19 * 60 + 30)   # 開始が5:30より前で勤務が6時間30分超、または終了が19:30より後なら休憩を追加
NO_WORK = "勤務時間を割り振らない日"
REMARKS = [NO_WORK, "年次休暇"]
APPROVE = "承認"

SIN, WARI, HYO = "フレックス申告", "フレックス割振り（所属長確認用）", "フレックス表"
S, W = f"'{SIN}'", f"'{WARI}'"

# 日ごとの表の列（フレックス申告・フレックス割振りで同じ）：A 期間、B 勤務日、C 曜日、D〜W 時刻（時・分の2列ずつ×10）、
# X・Y 時間数（時・分）、Z 備考、AA 確認。隠し列 AB〜AH は計算用
TIME_ITEMS = [("コアタイム", "開始"), ("コアタイム", "終了"), ("勤務", "開始"), ("勤務", "終了"),
              ("休憩①", "開始"), ("休憩①", "終了"), ("休憩②", "開始"), ("休憩②", "終了"), ("休憩③", "開始"), ("休憩③", "終了")]
TIME_COLS = [(get_column_letter(4 + 2 * i), get_column_letter(5 + 2 * i)) for i in range(10)]   # (時, 分)
(CS, CE, WS, WE, B1S, B1E, B2S, B2E, B3S, B3E) = TIME_COLS
HOURS, MINS, NOTE, CHECK = "X", "Y", "Z", "AA"
# 「勤務時間を割り振らない日」で割振り（承認）済みの日の割振年月日（フレックス表の AF 列。行はフレックス申告と同じ。ほかの日は空欄）。
# 休暇申請・テレワーク報告の重なりの確認に使う（先に承認・確認・割振りされた方が優先）
NOWORK_COL = "AF"
NOWORK_RANGE = f"'{HYO}'!${NOWORK_COL}${FIRST}:${NOWORK_COL}${LAST}"
DAY_RANGE = f"'{SIN}'!$B${FIRST}:$B${LAST}"
NOWORK_MARK = "※"      # 確認の欄の、承認を止めない注意（単位期間の「入力に誤りがあります」に数えない）
OFF, WEEKEND, T_START, T_END, T_BREAK, T_WORK, T_BREAKS = "AB", "AC", "AD", "AE", "AF", "AG", "AH"
# 単位期間の表（右上。日ごとの表を下にスクロールしても見えるよう、固定した 3〜11 行目に横に並べる）：AJ 列が見出し、AK〜AX 列が期間 1〜14。
# 単位期間は本人が開始日を入れ、終了日は開始日＋27日。日ごとの行の期間番号は、その日を含む単位期間。日ごとの表は単位期間1の開始日から始まる
PCOL = lambda k: get_column_letter(36 + k)       # 期間 k の列（AK〜AX）
PK1, PKN = "AK", "AX"
SR_NO, SR_START, SR_END, SR_TOTAL, SR_DECL, SR_STATUS = 3, 4, 5, 6, 7, 8      # 申告：8〜11 行目が状態
WR_APPROVE, WR_ALLOC, WR_STATUS = 8, 9, 10                                    # 割振り：3〜7 は申告と同じ、10〜11 行目が状態
prow = lambda row: f"${PK1}${row}:${PKN}${row}"

FONT = "ＭＳ Ｐゴシック"
f_base = Font(name=FONT, size=10)
f_small = Font(name=FONT, size=9)
f_bold = Font(name=FONT, size=10, bold=True)
f_head = Font(name=FONT, size=9, bold=True, color="FFFFFF")
f_title = Font(name=FONT, size=14, bold=True, color="1E3A5F")
f_note = Font(name=FONT, size=9, color="555555")
fill_head = PatternFill("solid", fgColor="1E3A5F")
fill_staff = PatternFill("solid", fgColor="FFF2CC")
fill_boss = PatternFill("solid", fgColor="DDEBF7")
fill_calc = PatternFill("solid", fgColor="F2F2F2")
fill_bad = PatternFill("solid", fgColor="F4B6B0", bgColor="F4B6B0")
font_bad = Font(name=FONT, color="9C0006", bold=True)
fill_off = PatternFill("solid", fgColor="D9D9D9", bgColor="D9D9D9")
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left = Alignment(horizontal="left", vertical="center", wrap_text=True)
UNLOCK = Protection(locked=False)


def hm(minutes):
    return f"{minutes // 60}:{minutes % 60:02d}"


def minutes_of(h, m, r):
    """時・分の欄の分（時が空なら ""）"""
    return f'IF(COUNT({h}{r})=0,"",{h}{r}*60+N({m}{r}))'


def add(wb, MG, OFF_DAYS, LEAVE, TW):
    """MG は管理シートの名前（'管理'）、OFF_DAYS は祝日・年末年始の範囲"""
    mg = wb[MG.strip("'")]
    mg["A12"], mg["B12"] = "フレックスの1年間の始期", FLEX_START
    mg["B12"].number_format = "yyyy/m/d"
    START = f"{MG}!$B$12"
    wb.defined_names["フレックス始期"] = DefinedName("フレックス始期", attr_text=START)

    sin = wb.create_sheet(SIN)
    wari = wb.create_sheet(WARI)
    hyo = wb.create_sheet(HYO)

    def notes(ws, texts):
        """6〜7 行目の説明（A〜AI を結合して折り返す）"""
        for r, text in zip((6, 7), texts):
            ws.merge_cells(f"A{r}:AI{r}")
            ws[f"A{r}"] = text
            ws[f"A{r}"].font = f_note
            ws[f"A{r}"].alignment = Alignment(wrap_text=True, vertical="center")
            ws.row_dimensions[r].height = 26

    def day_frame(ws, title, legend):
        ws["A1"] = title
        ws["A1"].font = f_title
        ws["H1"] = legend
        ws["H1"].font = f_note
        ws.column_dimensions["A"].width = 4.5
        ws.column_dimensions["B"].width = 10
        ws.column_dimensions["C"].width = 5
        for h, m in TIME_COLS:
            ws.column_dimensions[h].width = 3.6
            ws.column_dimensions[m].width = 3.6
        ws.column_dimensions[HOURS].width = 3.6
        ws.column_dimensions[MINS].width = 3.6
        ws.column_dimensions[NOTE].width = 22
        ws.column_dimensions[CHECK].width = 34
        for c in (OFF, WEEKEND, T_START, T_END, T_BREAK, T_WORK, T_BREAKS):
            ws.column_dimensions[c].hidden = True
        ws.column_dimensions["AI"].width = 2         # 日ごとの行では、その日の承認済みの休暇の行の番号（文字は見せない）
        # 見出し（8〜11行目）：8 グループ、9 項目、10 開始・終了、11 時・分
        def head(ref1, ref2, text):
            if ref1 != ref2:
                ws.merge_cells(f"{ref1}:{ref2}")
            ws[ref1] = text
            for row in ws[f"{ref1}:{ref2}"]:
                for c in row:
                    c.fill, c.border = fill_head, border
            ws[ref1].font, ws[ref1].alignment = f_head, center
        head("A8", "A11", "期間")
        head("B8", "C11", "勤務日")
        groups = [("コアタイム", 0, 1), ("勤務", 2, 3), ("休憩①", 4, 5), ("休憩②", 6, 7), ("休憩③", 8, 9)]
        for name, i, j in groups:
            head(f"{TIME_COLS[i][0]}8", f"{TIME_COLS[j][1]}9" if name.startswith("休憩") else f"{TIME_COLS[j][1]}8", name)
            if not name.startswith("休憩"):
                head(f"{TIME_COLS[i][0]}9", f"{TIME_COLS[i][1]}10", "開始")
                head(f"{TIME_COLS[j][0]}9", f"{TIME_COLS[j][1]}10", "終了")
            else:
                head(f"{TIME_COLS[i][0]}10", f"{TIME_COLS[i][1]}10", "開始")
                head(f"{TIME_COLS[j][0]}10", f"{TIME_COLS[j][1]}10", "終了")
        head(f"{HOURS}8", f"{MINS}10", "時間数")
        for h, m in TIME_COLS + [(HOURS, MINS)]:
            head(f"{h}11", f"{h}11", "時")
            head(f"{m}11", f"{m}11", "分")
        head(f"{NOTE}8", f"{NOTE}11", "備考")
        head(f"{CHECK}8", f"{CHECK}11", "確認")
        for r in range(8, 12):
            ws.row_dimensions[r].height = 15
        ws.freeze_panes = f"D{FIRST}"
        ws.sheet_view.zoomScale = 90

    # ===== フレックス申告（本人） =====
    day_frame(sin, "フレックスタイム制　申告（本人が入力）",
              "黄色：本人が入力（既定の 9:00〜17:45・休憩 12:00〜13:00 が入っています）　灰色：自動　赤：直すところ（理由は「確認」の欄）")
    for r, (label, value, fill) in enumerate([
        ("所属", None, fill_staff),
        ("氏名", f"={MG}!$B$2&\"\"", fill_calc),
        ("表の最初の日", f"=$B${FIRST}", fill_calc),
    ], start=3):
        sin[f"A{r}"] = label
        sin.merge_cells(f"A{r}:C{r}")
        sin.merge_cells(f"D{r}:M{r}")
        sin[f"A{r}"].font, sin[f"A{r}"].alignment = f_bold, center
        sin[f"D{r}"] = value
        for col in "ABCDEFGHIJKLM":
            sin[f"{col}{r}"].border = border
        for col in "DEFGHIJKLM":
            sin[f"{col}{r}"].fill = fill
        sin[f"D{r}"].font, sin[f"D{r}"].alignment = f_bold, left
    sin["D5"].number_format = 'yyyy"年"m"月"d"日"'
    for col in "DEFGHIJKLM":
        sin[f"{col}3"].protection = UNLOCK
    notes(sin, [
        "・右上の表に単位期間（4週間）の開始日を入れます（終了日は自動）。日ごとの表は単位期間1の開始日から始まります。勤務開始・終了と休憩を時と分に分けて書き、"
        "勤務しない日は時刻を消して備考に「勤務時間を割り振らない日」と書きます。",
        "・コアタイム（10:00〜12:00）と祝日・年末年始は既定のまま変えません。単位期間の合計を155時間にし、右上の申告年月日を入れて所属長に伝えます"
        "（承認の後で書き直したら、申告年月日も書き直した日にします）。コアタイム免除日など、入力できない欄を変えたいときは総務課へ。",
    ])

    for r in range(FIRST, LAST + 1):
        i = r - FIRST
        b = f"$B{r}"
        # その日を含む単位期間の番号（重なっているときは先の単位期間。どれにも入らなければ空欄）
        sin[f"A{r}"] = (f'=IFERROR(INDEX({prow(SR_NO)},MATCH(1,INDEX(({prow(SR_START)}<>"")'
                        f'*({prow(SR_START)}<={b})*({prow(SR_END)}>={b}),0),0)),"")')
        sin[f"B{r}"] = (f"=IF(ISNUMBER(${PCOL(1)}${SR_START}),${PCOL(1)}${SR_START},{START})" if i == 0 else f"=B{r - 1}+1")
        sin[f"C{r}"] = f'=TEXT({b},"(aaa)")'
        sin[f"B{r}"].number_format = 'm"月"d"日"'
        # 休日（祝日・年末年始で平日）と週休日（土日）
        sin[f"{WEEKEND}{r}"] = f"=IF(WEEKDAY({b},2)>=6,1,0)"
        sin[f"{OFF}{r}"] = f"=IF(AND({WEEKEND}{r}=0,COUNTIF({OFF_DAYS},{b})>0),1,0)"
        # その日の承認済みの休暇（休暇申請の表の何番目の行か）。土日・祝日は、土日祝でも取れる休暇だけ
        lv = lambda col: f"'{LEAVE[0]}'!${col}${LEAVE[1]}:${col}${LEAVE[2]}"
        sin[f"AI{r}"] = (f'=IFERROR(MATCH(1,INDEX(({lv("L")}="承認")*ISNUMBER({lv("E")})*({lv("E")}<={b})*({lv("W")}>={b})'
                         f'*((({lv("R")}>0)+(({WEEKEND}{r}+{OFF}{r})=0))>0),0),0),"")')
        sin[f"AI{r}"].number_format = ";;;"
        own = f"'{HYO}'!${NOWORK_COL}{r}"
        tw = lambda col: f"'{TW[0]}'!${col}${TW[1]}:${col}${TW[2]}"
        # 備考の既定：承認済みの休暇（種別があれば種別、時間・分単位なら時刻も）。承認を取り消すと消える
        ix = f"$AI{r}"
        sin[f"{NOTE}{r}"] = (f'=IF(ISNUMBER({ix}),IF(INDEX({lv("C")},{ix})<>"",INDEX({lv("C")},{ix}),INDEX({lv("B")},{ix}))'
                             f'&IF(INDEX({lv("Z")},{ix})=1,""," "&INDEX({lv("U")},{ix})),"")')
        wd = f"${WEEKEND}{r}=0"
        # コアタイム（変えない）と既定の時刻（本人が書き換える。土日は空欄）
        for (h, m), v in ((CS, CORE[0]), (CE, CORE[1]), (WS, WORK[0]), (WE, WORK[1]), (B1S, BREAK[0]), (B1E, BREAK[1])):
            sin[f"{h}{r}"] = f'=IF({wd},{v // 60},"")'
            sin[f"{m}{r}"] = f'=IF({wd},{v % 60},"")'
        # 分（計算用）
        sin[f"{T_START}{r}"] = "=" + minutes_of(WS[0], WS[1], r)
        sin[f"{T_END}{r}"] = "=" + minutes_of(WE[0], WE[1], r)
        brk = []
        for (s_, e_) in ((B1S, B1E), (B2S, B2E), (B3S, B3E)):
            brk.append(f"IF(AND(COUNT({s_[0]}{r})>0,COUNT({e_[0]}{r})>0),({e_[0]}{r}*60+N({e_[1]}{r}))-({s_[0]}{r}*60+N({s_[1]}{r})),0)")
        sin[f"{T_BREAK}{r}"] = "=" + "+".join(brk)
        sin[f"{T_BREAKS}{r}"] = "=" + "+".join(f"IF(COUNT({s_[0]}{r})>0,1,0)" for s_ in (B1S, B2S, B3S))
        sin[f"{T_WORK}{r}"] = f'=IF(OR({T_START}{r}="",{T_END}{r}=""),"",{T_END}{r}-{T_START}{r}-{T_BREAK}{r})'
        sin[f"{HOURS}{r}"] = f'=IF({T_WORK}{r}="","",IF({T_WORK}{r}<0,"−",INT({T_WORK}{r}/60)))'
        sin[f"{MINS}{r}"] = f'=IF({T_WORK}{r}="","",IF({T_WORK}{r}<0,"",MOD({T_WORK}{r},60)))'
        st, en, wk = f"{T_START}{r}", f"{T_END}{r}", f"{T_WORK}{r}"
        core_s = f"(N({CS[0]}{r})*60+N({CS[1]}{r}))"
        core_e = f"(N({CE[0]}{r})*60+N({CE[1]}{r}))"
        break_checks = []
        for (s_, e_) in ((B1S, B1E), (B2S, B2E), (B3S, B3E)):
            sm = f"(N({s_[0]}{r})*60+N({s_[1]}{r}))"
            em = f"(N({e_[0]}{r})*60+N({e_[1]}{r}))"
            has_s, has_e = f"COUNT({s_[0]}{r})>0", f"COUNT({e_[0]}{r})>0"
            break_checks.append(f"({has_s})<>({has_e})")
            break_checks.append(f"AND({has_s},{has_e},{em}<={sm})")
            break_checks.append(f'AND({has_s},{has_e},{st}<>"",{en}<>"",OR({sm}<{st},{em}>{en}))')
        need_extra = (f'AND({wk}<>"",OR(AND({st}<{EXTRA_BREAK[0]},{wk}>{EXTRA_BREAK[1]}),{en}>{EXTRA_BREAK[2]}),{T_BREAKS}{r}<2)')
        default_changed = (f'OR({st}<>{WORK[0]},{en}<>{WORK[1]},{T_BREAK}{r}<>{BREAK[1] - BREAK[0]},{T_BREAKS}{r}<>1)')
        sin[f"{CHECK}{r}"] = (
            f'=IF($A{r}="","",'
            f'IF(({st}="")<>({en}=""),"・勤務開始と勤務終了の両方を書いてください","")'
            f'&IF(AND({st}<>"",{en}<>""),IF({en}<={st},"・勤務終了が勤務開始より前です",""),"")'
            f'&IF(AND({st}<>"",{en}<>""),IF(OR({st}<{FLEXIBLE[0]},{en}>{FLEXIBLE[1]}),"・勤務は5:00〜22:00の間にしてください",""),"")'
            f'&IF(AND({st}<>"",{en}<>"",COUNT({CS[0]}{r})>0),IF(OR({st}>{core_s},{en}<{core_e}),"・コアタイムに勤務していません",""),"")'
            f'&IF(OR({",".join(break_checks)}),"・休憩の時刻を確かめてください","")'
            f'&IF({need_extra},"・休憩を追加してください（開始が5:30より前で6時間30分を超える、または終了が19:30より後）","")'
            f'&IF(AND({wk}<>"",N({wk})<0),"・時間数がマイナスです","")'
            f'&IF(AND({WEEKEND}{r}=1,{st}<>""),"・週休日（土日）に勤務を割り振っています","")'
            f'&IF(AND({OFF}{r}=1,OR({st}="",IF({st}="",TRUE,{default_changed}))),"・祝日・年末年始は既定の時間のままにしてください","")'
            f'&IF(AND({WEEKEND}{r}=0,{OFF}{r}=0,{st}="",{en}="",ISERROR(SEARCH("割り振らない",{NOTE}{r}))),'
            f'"・勤務しない日は備考に「{NO_WORK}」と書いてください","")'
            f'&IF(AND({st}<>"",ISNUMBER(SEARCH("割り振らない",{NOTE}{r}))),"・勤務を割り振った日です（備考を確かめてください）","")'
            # 割り振らない日に、先に承認された休暇・先に確認されたテレワーク（この単位期間が割振り前か、割振年月日が同じ日か後）。承認は止めない
            f'&IF(AND(ISNUMBER(SEARCH("割り振らない",{NOTE}{r})),'
            f'SUMPRODUCT(({lv("L")}="承認")*ISNUMBER({lv("E")})*({lv("E")}<={b})*({lv("W")}>={b})*((((1-ISNUMBER({own}))+({lv("M")}<=N({own})))>0)))'
            f'+SUMPRODUCT(ISNUMBER({tw("H")})*({tw("C")}={b})*((((1-ISNUMBER({own}))+({tw("H")}<=N({own})))>0)))>0),'
            f'"{NOWORK_MARK}承認済みの休暇・確認済みのテレワークがある日です（先に承認・確認・割振りされた方が優先）",""))'
        )
        for c in range(1, 28):
            cell = sin.cell(row=r, column=c)
            cell.font = f_small if c <= 26 else Font(name=FONT, size=8, color="9C0006")
            cell.border = border
            cell.alignment = center if c <= 25 else left
            cell.fill = fill_calc
        for h, m in TIME_COLS[2:]:
            for col in (h, m):
                sin[f"{col}{r}"].fill = fill_staff
                sin[f"{col}{r}"].protection = UNLOCK
        sin[f"{NOTE}{r}"].fill = fill_staff
        sin[f"{NOTE}{r}"].protection = UNLOCK
        sin[f"{NOTE}{r}"].font = Font(name=FONT, size=9, color="C00000")
        sin.row_dimensions[r].height = 15

    # 単位期間に入らない日は薄い灰色の文字、単位期間の最後の日の下に線（開始日によって変わるので条件付き書式で）
    from openpyxl.styles.differential import DifferentialStyle
    sin.conditional_formatting.add(f"A{FIRST}:{CHECK}{LAST}", FormulaRule(
        formula=[f'$A{FIRST}=""'], font=Font(color="A6A6A6")))
    sin.conditional_formatting.add(f"A{FIRST}:{CHECK}{LAST}", FormulaRule(
        formula=[f'AND($A{FIRST}<>"",$B{FIRST}=INDEX({prow(SR_END)},$A{FIRST}))'],
        border=Border(bottom=Side(style="thin", color="1E3A5F"))))
    # 赤：確認の文があれば、その行の勤務・休憩の欄と確認の欄
    sin.conditional_formatting.add(f"{WS[0]}{FIRST}:{B3E[1]}{LAST}",
                                   FormulaRule(formula=[f'AND(${CHECK}{FIRST}<>"",{WS[0]}{FIRST}&""<>"")'], fill=fill_bad, font=font_bad))
    sin.conditional_formatting.add(f"{CHECK}{FIRST}:{CHECK}{LAST}", FormulaRule(formula=[f'AND(${CHECK}{FIRST}<>"",LEFT(${CHECK}{FIRST},1)<>"{NOWORK_MARK}")'], fill=fill_bad, font=font_bad))
    # ※ の注意（承認を止めない）は薄い橙
    sin.conditional_formatting.add(f"{CHECK}{FIRST}:{CHECK}{LAST}", FormulaRule(formula=[f'LEFT(${CHECK}{FIRST},1)="{NOWORK_MARK}"'],
                                   fill=PatternFill("solid", fgColor="FCE4D6", bgColor="FCE4D6")))
    sin.conditional_formatting.add(f"{NOTE}{FIRST}:{NOTE}{LAST}",
                                   FormulaRule(formula=[f'ISNUMBER(SEARCH("備考",${CHECK}{FIRST}))'], fill=fill_bad, font=font_bad))
    # 土日・祝日・年末年始の行は灰色
    sin.conditional_formatting.add(f"A{FIRST}:{MINS}{LAST}",
                                   FormulaRule(formula=[f"OR(${WEEKEND}{FIRST}=1,${OFF}{FIRST}=1)"], fill=fill_off))

    # 入力の制限：時は 0〜23、分は 0〜59 の整数。祝日・年末年始の行は書き換えられない（貼り付け・削除ですり抜けたら赤）
    for idx, (h, m) in enumerate(TIME_COLS[2:]):
        for col, top in ((h, 23), (m, 59)):
            dv = DataValidation(type="custom", allow_blank=True, imeMode="halfAlpha",
                                formula1=f"AND(${OFF}{FIRST}=0,ISNUMBER({col}{FIRST}),{col}{FIRST}=INT({col}{FIRST}),{col}{FIRST}>=0,{col}{FIRST}<={top})")
            dv.errorTitle = "入力できません"
            dv.error = f"{'時は 0〜23' if top == 23 else '分は 0〜59'} の数を書いてください。祝日・年末年始の行は変えられません。"
            sin.add_data_validation(dv)
            dv.add(f"{col}{FIRST}:{col}{LAST}")
    dv = DataValidation(type="list", allow_blank=True, formula1='"' + ",".join(REMARKS) + '"')
    dv.showErrorMessage = False
    sin.add_data_validation(dv)
    dv.add(f"{NOTE}{FIRST}:{NOTE}{LAST}")

    # ===== フレックス表（非表示）：Power Query が読む表と、単位期間の計算 =====
    hyo.sheet_state = "veryHidden"
    day_heads = ["日付", "期間", "区分", "勤務開始", "勤務終了", "休憩", "勤務時間", "休憩の時刻", "備考", "確認"]
    for i, h in enumerate(day_heads, start=1):
        hyo.cell(row=1, column=i, value=h)
    for r in range(FIRST, LAST + 1):
        t = r - FIRST + 2      # フレックス表の行
        src = f"{S}!"
        brk_txt = "&".join(
            f'IF(COUNT({src}{s_[0]}{r})>0,TEXT(({src}{s_[0]}{r}*60+N({src}{s_[1]}{r}))/1440,"h:mm")&"-"&TEXT(({src}{e_[0]}{r}*60+N({src}{e_[1]}{r}))/1440,"h:mm")&" ","")'
            for s_, e_ in ((B1S, B1E), (B2S, B2E), (B3S, B3E)))
        vals = [
            f'=TEXT({src}$B{r},"yyyy-mm-dd")',
            f"={src}$A{r}",
            f'=IF({src}${WEEKEND}{r}=1,"週休日",IF({src}${OFF}{r}=1,"休日","平日"))',
            f'={src}{T_START}{r}',
            f'={src}{T_END}{r}',
            f'={src}{T_BREAK}{r}',
            f'={src}{T_WORK}{r}',
            f"=TRIM({brk_txt})",
            f'={src}{NOTE}{r}&""',
            f'={src}{CHECK}{r}&""',
        ]
        for i, v in enumerate(vals, start=1):
            hyo.cell(row=t, column=i, value=v)
    t = Table(displayName="フレックス日別", ref=f"A1:{get_column_letter(len(day_heads))}{DAYS + 1}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
    hyo.add_table(t)

    # 単位期間の表（L 列から）：期間・開始日・終了日・合計（分）・誤り・申告年月日・承認・割振年月日・状態・所属・所属長氏名
    per_heads = ["期間", "開始日", "終了日", "合計", "誤り", "申告年月日", "承認", "割振年月日", "状態", "所属", "所属長氏名", "異動後の所属長氏名"]
    # 割り振らない日で、単位期間が割振り済み（承認・割振年月日あり・申告年月日が割振年月日より後でない）なら、その割振年月日。
    # 単位期間の状態（確認の欄の誤りを数える）は使わない（確認の欄がこの列を使うので、循環しないように）
    for r in range(FIRST, LAST + 1):
        kk = f"{S}!$A{r}"
        alloc_k = f"INDEX({W}!${PK1}${WR_ALLOC}:${PKN}${WR_ALLOC},{kk})"
        hyo[f"{NOWORK_COL}{r}"] = (f'=IFERROR(IF(AND(ISNUMBER(SEARCH("割り振らない",{S}!${NOTE}{r})),INDEX({W}!${PK1}${WR_APPROVE}:${PKN}${WR_APPROVE},{kk})="{APPROVE}",'
                                    f'ISNUMBER({alloc_k}),N(INDEX({S}!${PK1}${SR_DECL}:${PKN}${SR_DECL},{kk}))<={alloc_k}),{alloc_k},""),"")')
    PL = 12   # L 列
    for i, h in enumerate(per_heads):
        hyo.cell(row=1, column=PL + i, value=h)
    for k in range(1, N_PERIODS + 1):
        t = k + 1
        pc = PCOL(k)            # 申告・割振りの単位期間の表の列
        c = lambda i: get_column_letter(PL + i)
        start, end = f"{S}!${pc}${SR_START}", f"{S}!${pc}${SR_END}"
        total = f"SUMIFS({S}!${T_WORK}${FIRST}:${T_WORK}${LAST},{S}!$A${FIRST}:$A${LAST},{k})"
        errors = (f'COUNTIFS({S}!$A${FIRST}:$A${LAST},{k},{S}!${CHECK}${FIRST}:${CHECK}${LAST},"?*")'
                  f'-COUNTIFS({S}!$A${FIRST}:$A${LAST},{k},{S}!${CHECK}${FIRST}:${CHECK}${LAST},"{NOWORK_MARK}*")')
        # 開始日の誤り：表の日の外、前の単位期間と重なる（開始日は前の単位期間の終了日より後）
        prev_end = f"{S}!${PCOL(k - 1)}${SR_END}" if k > 1 else '""'
        d1, dn = f"{S}!$B${FIRST}", f"{S}!$B${LAST}"     # 日ごとの表の最初の日・最後の日
        bad_start = (f'IF(OR({start}<{d1},{end}>{dn}),"開始日がこの表の日の外です（"&TEXT({d1},"yyyy/m/d")&"から"&TEXT({dn},"yyyy/m/d")&"）",'
                     f'IF(AND(ISNUMBER({prev_end}),{start}<={prev_end}),"開始日が前の単位期間と重なっています",""))')
        declared = f"{S}!${pc}${SR_DECL}"
        approved = f"{W}!${pc}${WR_APPROVE}"
        alloc = f"{W}!${pc}${WR_ALLOC}"
        status = (f'IF(NOT(ISNUMBER({start})),"",IF({bad_start}<>"",{bad_start},'
                  f'IF(NOT(ISNUMBER({declared})),"未申告",'
                  f'IF({c(3)}{t}<>{TOTAL_MIN},"合計が155時間ではありません",'
                  f'IF({c(4)}{t}>0,"入力に誤りがあります",'
                  f'IF({approved}<>"{APPROVE}","承認待ち",'
                  f'IF(NOT(ISNUMBER({alloc})),"割振年月日を入れてください",'
                  f'IF({declared}>{alloc},"承認後に申告が変更されました（承認し直してください）","承認済み"))))))))')
        vals = [k, f'=IF(ISNUMBER({start}),TEXT({start},"yyyy-mm-dd"),"")', f'=IF(ISNUMBER({end}),TEXT({end},"yyyy-mm-dd"),"")', f"={total}", f"={errors}",
                f'=IF(ISNUMBER({declared}),TEXT({declared},"yyyy-mm-dd"),"")', f'={approved}&""',
                f'=IF(ISNUMBER({alloc}),TEXT({alloc},"yyyy-mm-dd"),"")', "=" + status,
                f"={S}!$D$3&\"\"", f"={W}!$D$3&\"\"", f"={W}!$D$4&\"\""]
        for i, v in enumerate(vals):
            hyo.cell(row=t, column=PL + i, value=v)
    t = Table(displayName="フレックス期間", ref=f"L1:{get_column_letter(PL + len(per_heads) - 1)}{N_PERIODS + 1}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
    hyo.add_table(t)
    H = f"'{HYO}'"
    per_col = lambda i: f"{H}!${get_column_letter(PL + i)}$"

    # ===== 単位期間の表（申告・割振りの右上。固定した行に横に並べる） =====
    def period_table(ws, rows, values, inputs, input_fill, status_row):
        """rows：[(行, 見出し)]。values(k) は rows の順の値。inputs は入力する行。状態の行は status_row〜11 を結合して折り返す"""
        ws.column_dimensions["AJ"].width = 11
        ws[f"AJ2"] = "単位期間（4週間）"
        ws[f"AJ2"].font = f_bold
        for row, head_ in rows:
            c = ws[f"AJ{row}"]
            c.value, c.font, c.fill, c.border, c.alignment = head_, f_head, fill_head, border, center
        ws.merge_cells(f"AJ{status_row}:AJ11")
        for k in range(1, N_PERIODS + 1):
            col = PCOL(k)
            ws.column_dimensions[col].width = 10.5
            for (row, _), v in zip(rows, values(k)):
                c = ws[f"{col}{row}"]
                c.value = v
                c.font, c.border, c.alignment = f_small, border, center
                c.fill = input_fill if row in inputs else fill_calc
                if row in inputs:
                    c.protection = UNLOCK
                if row in (SR_START, SR_END, SR_DECL, WR_ALLOC) and row != WR_APPROVE:
                    c.number_format = "yyyy/m/d"
            ws.merge_cells(f"{col}{status_row}:{col}11")
            ws[f"{col}{status_row}"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            for r in range(status_row, 12):
                ws[f"{col}{r}"].border = border
        ws[f"{PCOL(1)}{SR_NO}"].number_format = "0"

    def total_text(k):
        tm = f"{per_col(3)}{k + 1}"
        return f'=INT({tm}/60)&"時間"&MOD({tm},60)&"分"'

    period_table(sin, [(SR_NO, "期間"), (SR_START, "開始日"), (SR_END, "終了日"), (SR_TOTAL, "合計"), (SR_DECL, "申告年月日"), (SR_STATUS, "状態")],
                 lambda k: [k, None, f'=IF(ISNUMBER({PCOL(k)}{SR_START}),{PCOL(k)}{SR_START}+{PERIOD - 1},"")',
                            f'=IF({PCOL(k)}{SR_START}="","",{total_text(k)[1:]})', None, f"={per_col(8)}{k + 1}"],
                 inputs={SR_START, SR_DECL}, input_fill=fill_staff, status_row=SR_STATUS)
    # 開始日・申告年月日は日付だけ
    dv = DataValidation(type="date", operator="greaterThan", formula1="43831", allow_blank=True, imeMode="halfAlpha")
    dv.errorTitle, dv.error = "日付を入れてください", "例：2026/4/1"
    sin.add_data_validation(dv)
    dv.add(f"{PK1}{SR_START}:{PKN}{SR_START}")
    dv.add(f"{PK1}{SR_DECL}:{PKN}{SR_DECL}")
    bad_status = lambda cell: f'AND({cell}<>"",{cell}<>"承認済み",{cell}<>"未申告",{cell}<>"承認待ち")'
    sin.conditional_formatting.add(f"{PK1}{SR_STATUS}:{PKN}{SR_STATUS}", FormulaRule(formula=[bad_status(f"{PK1}${SR_STATUS}")], fill=fill_bad, font=font_bad))
    sin.conditional_formatting.add(f"{PK1}{SR_TOTAL}:{PKN}{SR_TOTAL}", FormulaRule(
        formula=[f'AND({PK1}${SR_START}<>"",INDEX({H}!$O$2:$O${N_PERIODS + 1},{PK1}${SR_NO})<>{TOTAL_MIN})'], fill=fill_bad, font=font_bad))

    # ===== フレックス割振り（所属長） =====
    day_frame(wari, "フレックスタイム制　割振り（所属長が確認）",
              "本人の申告をそのまま写しています（ここでは書き換えられません）。青：所属長が入力　赤：確かめるところ")
    for r, (label, value, fill) in enumerate([
        ("所属長氏名", None, fill_boss),
        ("異動後の所属長氏名", None, fill_boss),
        ("氏名", f"={S}!D4&\"\"", fill_calc),
    ], start=3):
        wari[f"A{r}"] = label
        wari.merge_cells(f"A{r}:C{r}")
        wari.merge_cells(f"D{r}:M{r}")
        wari[f"A{r}"].font, wari[f"A{r}"].alignment = f_bold, center
        wari[f"D{r}"] = value
        for col in "ABCDEFGHIJKLM":
            wari[f"{col}{r}"].border = border
        for col in "DEFGHIJKLM":
            wari[f"{col}{r}"].fill = fill
            if fill is fill_boss:
                wari[f"{col}{r}"].protection = UNLOCK
        wari[f"D{r}"].font, wari[f"D{r}"].alignment = f_bold, left
    notes(wari, [
        "・右上の表で、単位期間ごとに内容を確かめ、承認の欄で「承認」を選び、割振年月日を入れます。状態が「承認済み」の単位期間が、管理者に取り込まれます。",
        "・本人が承認の後で書き直すと「承認後に申告が変更されました」と出ます。確かめて、割振年月日を確かめた日に書き直します。",
    ])
    for r in range(FIRST, LAST + 1):
        for c in range(1, 28):
            col = get_column_letter(c)
            src = f"{S}!{col}{r}"
            wari[f"{col}{r}"] = f'=IF({src}="","",{src})' if c != 2 else f"={src}"
            cell = wari[f"{col}{r}"]
            cell.font = f_small if c <= 26 else Font(name=FONT, size=8, color="9C0006")
            if c == 26:
                cell.font = Font(name=FONT, size=9, color="C00000")
            cell.border = copy(sin.cell(row=r, column=c).border)
            cell.alignment = center if c <= 25 else left
            cell.fill = fill_calc
        wari[f"B{r}"].number_format = 'm"月"d"日"'
        wari.row_dimensions[r].height = 15
        for col in (OFF, WEEKEND):
            wari[f"{col}{r}"] = f"={S}!{col}{r}"
    wari.conditional_formatting.add(f"{CHECK}{FIRST}:{CHECK}{LAST}", FormulaRule(formula=[f'AND(${CHECK}{FIRST}<>"",LEFT(${CHECK}{FIRST},1)<>"{NOWORK_MARK}")'], fill=fill_bad, font=font_bad))
    # ※ の注意（承認を止めない）は薄い橙
    wari.conditional_formatting.add(f"{CHECK}{FIRST}:{CHECK}{LAST}", FormulaRule(formula=[f'LEFT(${CHECK}{FIRST},1)="{NOWORK_MARK}"'],
                                   fill=PatternFill("solid", fgColor="FCE4D6", bgColor="FCE4D6")))
    wari.conditional_formatting.add(f"A{FIRST}:{CHECK}{LAST}", FormulaRule(formula=[f'$A{FIRST}=""'], font=Font(color="A6A6A6")))
    wari.conditional_formatting.add(f"A{FIRST}:{CHECK}{LAST}", FormulaRule(
        formula=[f'AND($A{FIRST}<>"",$B{FIRST}=INDEX({prow(SR_END)},$A{FIRST}))'],
        border=Border(bottom=Side(style="thin", color="1E3A5F"))))
    wari.conditional_formatting.add(f"A{FIRST}:{MINS}{LAST}",
                                    FormulaRule(formula=[f"OR(${WEEKEND}{FIRST}=1,${OFF}{FIRST}=1)"], fill=fill_off))
    ref = lambda k, row: f"{S}!${PCOL(k)}${row}"
    period_table(wari, [(SR_NO, "期間"), (SR_START, "開始日"), (SR_END, "終了日"), (SR_TOTAL, "合計"), (SR_DECL, "申告年月日"),
                        (WR_APPROVE, "承認"), (WR_ALLOC, "割振年月日"), (WR_STATUS, "状態")],
                 lambda k: [k, f'=IF(ISNUMBER({ref(k, SR_START)}),{ref(k, SR_START)},"")', f'=IF(ISNUMBER({ref(k, SR_END)}),{ref(k, SR_END)},"")',
                            f'=IF({ref(k, SR_START)}="","",{total_text(k)[1:]})', f'=IF(ISNUMBER({ref(k, SR_DECL)}),{ref(k, SR_DECL)},"")',
                            None, None, f"={per_col(8)}{k + 1}"],
                 inputs={WR_APPROVE, WR_ALLOC}, input_fill=fill_boss, status_row=WR_STATUS)
    dv = DataValidation(type="list", allow_blank=True, formula1=f'"{APPROVE}"')
    dv.errorTitle, dv.error = "承認", "「承認」を選んでください（取り消すときは消します）。"
    wari.add_data_validation(dv)
    dv.add(f"{PK1}{WR_APPROVE}:{PKN}{WR_APPROVE}")
    dv = DataValidation(type="date", operator="greaterThan", formula1="43831", allow_blank=True, imeMode="halfAlpha")
    dv.errorTitle, dv.error = "日付を入れてください", "例：2026/3/31"
    wari.add_data_validation(dv)
    dv.add(f"{PK1}{WR_ALLOC}:{PKN}{WR_ALLOC}")
    wari.conditional_formatting.add(f"{PK1}{WR_STATUS}:{PKN}{WR_STATUS}", FormulaRule(formula=[bad_status(f"{PK1}${WR_STATUS}")], fill=fill_bad, font=font_bad))
    st = f"{PK1}${WR_STATUS}"
    wari.conditional_formatting.add(f"{PK1}{WR_APPROVE}:{PKN}{WR_ALLOC}", FormulaRule(
        formula=[f'AND({PK1}${WR_APPROVE}="{APPROVE}",OR({PK1}${WR_ALLOC}="",LEFT({st},3)="承認後",{st}="合計が155時間ではありません",{st}="入力に誤りがあります",{st}="未申告",LEFT({st},3)="開始日"))'],
        fill=fill_bad, font=font_bad))

    for ws in (sin, wari):
        ws.protection.sheet = True
        ws.protection.formatColumns = False
        ws.protection.formatRows = False
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_title_rows = "8:11"
        ws.print_area = f"A1:{CHECK}{LAST}"
    return [SIN, WARI]
