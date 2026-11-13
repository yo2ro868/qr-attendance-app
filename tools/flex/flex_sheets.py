"""休暇申請シートに、フレックスタイム制の申告・割振りのシートを足す（tools/leave_sheet/make_form.py から呼ぶ）。

シート：
  フレックス申告 … 本人が入力。1年間（始期から392日＝4週間×14の単位期間）の日ごとの勤務開始・終了・休憩（時と分を別の欄）と、
                   単位期間ごとの申告年月日。
  フレックス割振り … 所属長が確認する。申告をそのまま写した日ごとの表と、単位期間ごとの承認・割振年月日。
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

SIN, WARI, HYO = "フレックス申告", "フレックス割振り", "フレックス表"
S, W = f"'{SIN}'", f"'{WARI}'"

# 日ごとの表の列（フレックス申告・フレックス割振りで同じ）：A 期間、B 勤務日、C 曜日、D〜W 時刻（時・分の2列ずつ×10）、
# X・Y 時間数（時・分）、Z 備考、AA 確認。隠し列 AB〜AH は計算用
TIME_ITEMS = [("コアタイム", "開始"), ("コアタイム", "終了"), ("勤務", "開始"), ("勤務", "終了"),
              ("休憩①", "開始"), ("休憩①", "終了"), ("休憩②", "開始"), ("休憩②", "終了"), ("休憩③", "開始"), ("休憩③", "終了")]
TIME_COLS = [(get_column_letter(4 + 2 * i), get_column_letter(5 + 2 * i)) for i in range(10)]   # (時, 分)
(CS, CE, WS, WE, B1S, B1E, B2S, B2E, B3S, B3E) = TIME_COLS
HOURS, MINS, NOTE, CHECK = "X", "Y", "Z", "AA"
OFF, WEEKEND, T_START, T_END, T_BREAK, T_WORK, T_BREAKS = "AB", "AC", "AD", "AE", "AF", "AG", "AH"
# 単位期間の表（右側。11 行目が見出し、12〜25 行目が期間 1〜14）
P_HEAD = 11
P_COLS = ["AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ"]

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


def add(wb, MG, OFF_DAYS):
    """MG は管理シートの名前（'管理'）、OFF_DAYS は祝日・年末年始の範囲"""
    mg = wb[MG.strip("'")]
    mg["A12"], mg["B12"] = "フレックスの1年間の始期", FLEX_START
    mg["B12"].number_format = "yyyy/m/d"
    START = f"{MG}!$B$12"
    wb.defined_names["フレックス始期"] = DefinedName("フレックス始期", attr_text=START)

    sin = wb.create_sheet(SIN)
    wari = wb.create_sheet(WARI)
    hyo = wb.create_sheet(HYO)

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
        ws.column_dimensions["AI"].width = 2
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
        ("1年間の始期", f"={START}", fill_calc),
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
    sin["O3"] = ("・単位期間（4週間）ごとに、勤務開始・終了と休憩を時と分に分けて書きます。勤務しない日は時刻を消し、備考に「勤務時間を割り振らない日」と書きます。")
    sin["O4"] = ("・コアタイム（10:00〜12:00）と、祝日・年末年始の行は入っています（変えません）。単位期間の合計が 155時間 になるようにします。")
    sin["O5"] = ("・書いたら、右の表の申告年月日を入れて、所属長に伝えます。承認の後で書き直したら、申告年月日を書き直した日にします（所属長が承認し直します）。")
    for ref in ("O3", "O4", "O5"):
        sin[ref].font = f_note
    sin["A6"] = "コアタイム免除日など、入力できない欄を変えたいときは総務課に連絡してください。"
    sin["A6"].font = f_note

    for r in range(FIRST, LAST + 1):
        i = r - FIRST
        b = f"$B{r}"
        sin[f"A{r}"] = f"=INT(({b}-{START})/{PERIOD})+1"
        sin[f"B{r}"] = f"={START}+{i}" if i == 0 else f"=B{r - 1}+1"
        sin[f"C{r}"] = f'=TEXT({b},"(aaa)")'
        sin[f"B{r}"].number_format = 'm"月"d"日"'
        # 休日（祝日・年末年始で平日）と週休日（土日）
        sin[f"{WEEKEND}{r}"] = f"=IF(WEEKDAY({b},2)>=6,1,0)"
        sin[f"{OFF}{r}"] = f"=IF(AND({WEEKEND}{r}=0,COUNTIF({OFF_DAYS},{b})>0),1,0)"
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
            f'=IF(({st}="")<>({en}=""),"・勤務開始と勤務終了の両方を書いてください","")'
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
        # 単位期間の最後の日の下に太い線
        if (i + 1) % PERIOD == 0:
            for c in range(1, 28):
                cell = sin.cell(row=r, column=c)
                cell.border = Border(left=thin, right=thin, top=thin, bottom=Side(style="medium", color="1E3A5F"))

    # 赤：確認の文があれば、その行の勤務・休憩の欄と確認の欄
    sin.conditional_formatting.add(f"{WS[0]}{FIRST}:{B3E[1]}{LAST}",
                                   FormulaRule(formula=[f'AND(${CHECK}{FIRST}<>"",{WS[0]}{FIRST}&""<>"")'], fill=fill_bad, font=font_bad))
    sin.conditional_formatting.add(f"{CHECK}{FIRST}:{CHECK}{LAST}", FormulaRule(formula=[f'${CHECK}{FIRST}<>""'], fill=fill_bad, font=font_bad))
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
    PL = 12   # L 列
    for i, h in enumerate(per_heads):
        hyo.cell(row=1, column=PL + i, value=h)
    for k in range(1, N_PERIODS + 1):
        t = k + 1
        r1, r2 = FIRST + (k - 1) * PERIOD, FIRST + k * PERIOD - 1
        pr = FIRST + k - 1      # 申告・割振りの単位期間の表の行
        c = lambda i: get_column_letter(PL + i)
        total = f"SUM({S}!${T_WORK}${r1}:${T_WORK}${r2})"
        errors = f'SUMPRODUCT(--({S}!${CHECK}${r1}:${CHECK}${r2}<>""))'
        declared = f"{S}!$AN${pr}"
        approved = f"{W}!$AO${pr}"
        alloc = f"{W}!$AP${pr}"
        status = (f'IF(NOT(ISNUMBER({declared})),"未申告",'
                  f'IF({c(3)}{t}<>{TOTAL_MIN},"合計が155時間ではありません",'
                  f'IF({c(4)}{t}>0,"入力に誤りがあります",'
                  f'IF({approved}<>"{APPROVE}","承認待ち",'
                  f'IF(NOT(ISNUMBER({alloc})),"割振年月日を入れてください",'
                  f'IF({declared}>{alloc},"承認後に申告が変更されました（承認し直してください）","承認済み"))))))')
        vals = [k, f'=TEXT({S}!$B${r1},"yyyy-mm-dd")', f'=TEXT({S}!$B${r2},"yyyy-mm-dd")', f"={total}", f"={errors}",
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

    # ===== 単位期間の表（申告・割振りの右側） =====
    def period_table(ws, heads, values, inputs, input_fill):
        for i, (head_, width) in enumerate(heads):
            col = P_COLS[i]
            ws.column_dimensions[col].width = width
            c = ws[f"{col}{P_HEAD}"]
            c.value, c.font, c.fill, c.border, c.alignment = head_, f_head, fill_head, border, center
        ws[f"{P_COLS[0]}{P_HEAD - 1}"] = "単位期間（4週間）"
        ws[f"{P_COLS[0]}{P_HEAD - 1}"].font = f_bold
        for k in range(1, N_PERIODS + 1):
            pr = FIRST + k - 1
            for i, v in enumerate(values(k, pr)):
                c = ws[f"{P_COLS[i]}{pr}"]
                c.value = v
                c.font, c.border, c.alignment = f_small, border, center
                c.fill = input_fill if i in inputs else fill_calc
                if i in inputs:
                    c.protection = UNLOCK
            ws[f"{P_COLS[1]}{pr}"].number_format = 'm"/"d'
            ws[f"{P_COLS[2]}{pr}"].number_format = 'm"/"d'

    def total_text(k):
        tm = f"{per_col(3)}{k + 1}"
        return f'=INT({tm}/60)&"時間"&MOD({tm},60)&"分"'

    period_table(sin, [("期間", 4), ("開始日", 6), ("終了日", 6), ("合計", 11), ("申告年月日", 11), ("状態", 30)],
                 lambda k, pr: [k, f"={S}!$B${FIRST + (k - 1) * PERIOD}", f"={S}!$B${FIRST + k * PERIOD - 1}",
                                total_text(k), None, f"={per_col(8)}{k + 1}"],
                 inputs={4}, input_fill=fill_staff)
    for k in range(1, N_PERIODS + 1):
        sin[f"AN{FIRST + k - 1}"].number_format = "yyyy/m/d"
    # 申告年月日は日付だけ
    dv = DataValidation(type="date", operator="greaterThan", formula1="43831", allow_blank=True, imeMode="halfAlpha")
    dv.errorTitle, dv.error = "日付を入れてください", "例：2026/3/31"
    sin.add_data_validation(dv)
    dv.add(f"AN{FIRST}:AN{FIRST + N_PERIODS - 1}")
    status_rules = [(f'AND($AO{FIRST}<>"承認済み",$AO{FIRST}<>"未申告",$AO{FIRST}<>"承認待ち")', fill_bad, font_bad)]
    for rng_, col in ((f"AO{FIRST}:AO{FIRST + N_PERIODS - 1}", "AO"),):
        for rule, fill, font in status_rules:
            sin.conditional_formatting.add(rng_, FormulaRule(formula=[rule], fill=fill, font=font))
    sin.conditional_formatting.add(f"AM{FIRST}:AM{FIRST + N_PERIODS - 1}",
                                   FormulaRule(formula=[f"{H}!$O2<>{TOTAL_MIN}"], fill=fill_bad, font=font_bad))

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
    wari["O3"] = "・右の表で、単位期間ごとに内容を確かめ、承認の欄で「承認」を選び、割振年月日を入れます。"
    wari["O4"] = "・本人が承認の後で書き直すと「承認後に申告が変更されました」と出ます。確かめて、割振年月日を確かめた日に書き直します。"
    wari["O5"] = "・状態が「承認済み」の単位期間が、管理者に取り込まれます。"
    for ref in ("O3", "O4", "O5"):
        wari[ref].font = f_note
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
    wari.conditional_formatting.add(f"{CHECK}{FIRST}:{CHECK}{LAST}", FormulaRule(formula=[f'${CHECK}{FIRST}<>""'], fill=fill_bad, font=font_bad))
    wari.conditional_formatting.add(f"A{FIRST}:{MINS}{LAST}",
                                    FormulaRule(formula=[f"OR(${WEEKEND}{FIRST}=1,${OFF}{FIRST}=1)"], fill=fill_off))
    period_table(wari, [("期間", 4), ("開始日", 6), ("終了日", 6), ("合計", 11), ("申告年月日", 11), ("承認", 7), ("割振年月日", 11), ("状態", 30)],
                 lambda k, pr: [k, f"={S}!$B${FIRST + (k - 1) * PERIOD}", f"={S}!$B${FIRST + k * PERIOD - 1}",
                                total_text(k), f'=IF(ISNUMBER({S}!$AN${pr}),{S}!$AN${pr},"")', None, None, f"={per_col(8)}{k + 1}"],
                 inputs={5, 6}, input_fill=fill_boss)
    for k in range(1, N_PERIODS + 1):
        wari[f"AN{FIRST + k - 1}"].number_format = "yyyy/m/d"
        wari[f"AP{FIRST + k - 1}"].number_format = "yyyy/m/d"
    # 割振り側の列：AJ 期間 AK 開始 AL 終了 AM 合計 AN 申告年月日 AO 承認 AP 割振年月日 AQ 状態 … 承認は AO、割振年月日は AP
    dv = DataValidation(type="list", allow_blank=True, formula1=f'"{APPROVE}"')
    dv.errorTitle, dv.error = "承認", "「承認」を選んでください（取り消すときは消します）。"
    wari.add_data_validation(dv)
    dv.add(f"AO{FIRST}:AO{FIRST + N_PERIODS - 1}")
    dv = DataValidation(type="date", operator="greaterThan", formula1="43831", allow_blank=True, imeMode="halfAlpha")
    dv.errorTitle, dv.error = "日付を入れてください", "例：2026/3/31"
    wari.add_data_validation(dv)
    dv.add(f"AP{FIRST}:AP{FIRST + N_PERIODS - 1}")
    last_p = FIRST + N_PERIODS - 1
    wari.conditional_formatting.add(f"AQ{FIRST}:AQ{last_p}", FormulaRule(
        formula=[f'AND($AQ{FIRST}<>"承認済み",$AQ{FIRST}<>"未申告",$AQ{FIRST}<>"承認待ち")'], fill=fill_bad, font=font_bad))
    wari.conditional_formatting.add(f"AO{FIRST}:AP{last_p}", FormulaRule(
        formula=[f'AND($AO{FIRST}="{APPROVE}",OR($AP{FIRST}="",LEFT($AQ{FIRST},3)="承認後",$AQ{FIRST}="合計が155時間ではありません",$AQ{FIRST}="入力に誤りがあります",$AQ{FIRST}="未申告"))'],
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
