"""申請・報告シートに、フレックスタイム制のシートを足す（tools/leave_sheet/make_form.py から呼ぶ）。

シート：
  フレックス（単位期間・承認） … 単位期間（4週間）ごとに1行の表。本人が開始日・申告年月日を、所属長が承認・割振年月日を入れる。
                                  上に所属（職員名簿の部・課）・氏名・所属長氏名（職員名簿）・異動後の所属長氏名。
  フレックス（申告の内容） … 本人が入力。1年間（始期から392日＝4週間×14の単位期間）の日ごとの勤務開始・終了・休憩（時と分を別の欄）。
                            上の固定した行に、表示する単位期間（空欄なら今日を含む期間）の合計と155時間との差・状態。
  フレックス表（非表示）… 出勤・休暇管理データ が Power Query で読む表「フレックス日別」「フレックス期間」（時刻・時間は分）。
様式どおりの完成版のシート「申告・割振り簿」は tools/flex/add_flex_book.ps1 が、総務課の様式の Excel から作って足す。
2026-11-19 までは「フレックス申告」（右上に横に長い単位期間の表）と「フレックス割振り（所属長確認用）」（申告を写した表と承認）の2枚だった。

日ごとの行は 12 行目が始期の日（様式の申告・割振り簿と同じ行）。
"""
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

PER, SIN, HYO = "フレックス（単位期間・承認）", "フレックス（申告の内容）", "フレックス表"
P, S = f"'{PER}'", f"'{SIN}'"

# 日ごとの表の列（フレックス（申告の内容））：A 期間、B 勤務日、C 曜日、D〜W 時刻（時・分の2列ずつ×10）、
# X・Y 時間数（時・分）、Z 備考、AA 確認。隠し列 AB〜AH は計算用
TIME_ITEMS = [("コアタイム", "開始"), ("コアタイム", "終了"), ("勤務", "開始"), ("勤務", "終了"),
              ("休憩①", "開始"), ("休憩①", "終了"), ("休憩②", "開始"), ("休憩②", "終了"), ("休憩③", "開始"), ("休憩③", "終了")]
TIME_COLS = [(get_column_letter(4 + 2 * i), get_column_letter(5 + 2 * i)) for i in range(10)]   # (時, 分)
(CS, CE, WS, WE, B1S, B1E, B2S, B2E, B3S, B3E) = TIME_COLS
HOURS, MINS, NOTE, CHECK = "X", "Y", "Z", "AA"
# 「勤務時間を割り振らない日」で割振り（承認）済みの日の割振年月日（フレックス表の AF 列。行は申告の内容と同じ。ほかの日は空欄）。
# 休暇申請・テレワーク報告の重なりの確認に使う（先に承認・確認・割振りされた方が優先）
NOWORK_COL = "AF"
NOWORK_RANGE = f"'{HYO}'!${NOWORK_COL}${FIRST}:${NOWORK_COL}${LAST}"
DAY_RANGE = f"'{SIN}'!$B${FIRST}:$B${LAST}"
NOWORK_MARK = "※"      # 確認の欄の、承認を止めない注意（単位期間の「入力に誤りがあります」に数えない）
OFF, WEEKEND, T_START, T_END, T_BREAK, T_WORK, T_BREAKS = "AB", "AC", "AD", "AE", "AF", "AG", "AH"

# 単位期間の表（フレックス（単位期間・承認）の 8 行目が見出し、9〜22 行目が期間 1〜14）。単位期間は本人が開始日を入れ、終了日は開始日＋27日。
# 日ごとの行の期間番号は、その日を含む単位期間。日ごとの表は単位期間1の開始日から始まる
P_HEAD, P_FIRST = 8, 9
P_LAST = P_FIRST + N_PERIODS - 1     # 22
PC_NO, PC_START, PC_END, PC_TOTAL, PC_DIFF, PC_DECL, PC_APPROVE, PC_ALLOC, PC_STATUS = "ABCDEFGHI"
prng = lambda col: f"{P}!${col}${P_FIRST}:${col}${P_LAST}"      # 単位期間の表の列（ほかのシートから）
pcell = lambda col, k: f"{P}!${col}${P_FIRST + k - 1}"         # 期間 k の欄（ほかのシートから）
# 上の欄（フレックス（単位期間・承認） 3〜6 行目の C 列）：所属・氏名・所属長氏名・異動後の所属長氏名
P_SECTION, P_NAME, P_BOSS, P_BOSS2 = "C3", "C4", "C5", "C6"

FONT = "BIZ UDPゴシック"     # 見やすい（ユニバーサルデザイン）。総務課の様式（命令簿・申告・割振り簿）は元のフォントのまま
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


def hm_text(expr):
    """分 → 「155時間0分」"""
    return f'INT({expr}/60)&"時間"&MOD({expr},60)&"分"'


def diff_text(expr):
    """合計（分）と155時間との差の文"""
    d = f"({expr}-{TOTAL_MIN})"
    return (f'IF({d}=0,"155時間ちょうど",IF({d}<0,"あと"&{hm_text(f"(-{d})")}&"足りません",'
            f'{hm_text(d)}&"多すぎます"))')


def add(wb, MG, OFF_DAYS, LEAVE, TW):
    """MG は管理シートの名前（'管理'）、OFF_DAYS は祝日・年末年始の範囲"""
    mg = wb[MG.strip("'")]
    mg["A12"], mg["B12"] = "フレックスの1年間の始期", FLEX_START
    mg["B12"].number_format = "yyyy/m/d"
    START = f"{MG}!$B$12"
    wb.defined_names["フレックス始期"] = DefinedName("フレックス始期", attr_text=START)

    per = wb.create_sheet(PER)
    sin = wb.create_sheet(SIN)
    hyo = wb.create_sheet(HYO)
    H = f"'{HYO}'"
    PL = 12   # フレックス表の単位期間の表は L 列から
    per_col = lambda i: f"{H}!${get_column_letter(PL + i)}$"     # フレックス表の単位期間の表の i 番目の列（0 から）

    def title(ws, text, legend):
        ws["A1"] = text
        ws["A1"].font = f_title
        ws["A2"] = legend
        ws["A2"].font = f_note

    # ===== フレックス（単位期間・承認） =====
    title(per, "フレックスタイム制　単位期間・承認",
          "黄色：本人が入力　青：所属長が入力　灰色：自動　赤：直すところ・確かめるところ")
    for col, width in zip("ABCDEFGHI", (8, 11, 11, 12, 20, 11, 8, 11, 34)):     # A+B に「異動後の所属長氏名」が入る
        per.column_dimensions[col].width = width
    # 上の欄：所属（職員名簿の部・課）・氏名・所属長氏名（職員名簿）は自動。違うときは上書きできる。異動後の所属長氏名は所属長が入れる
    for r, (label, value, fill, unlock) in enumerate([
        ("所属", f'={MG}!$B$16&""', fill_staff, True),
        ("氏名", f'={MG}!$B$2&""', fill_calc, False),
        ("所属長氏名", f'={MG}!$B$15&""', fill_boss, True),
        ("異動後の所属長氏名", None, fill_boss, True),
    ], start=3):
        per[f"A{r}"] = label
        per.merge_cells(f"A{r}:B{r}")
        per.merge_cells(f"C{r}:E{r}")
        per[f"A{r}"].font, per[f"A{r}"].alignment = f_bold, center
        per[f"C{r}"] = value
        for col in "ABCDE":
            per[f"{col}{r}"].border = border
        for col in "CDE":
            per[f"{col}{r}"].fill = fill
            if unlock:
                per[f"{col}{r}"].protection = UNLOCK
        per[f"C{r}"].font, per[f"C{r}"].alignment = f_bold, left
        per.row_dimensions[r].height = 18
    per.merge_cells("F3:I6")
    per["F3"] = ("・本人：単位期間（4週間）の開始日を入れ（終了日は自動）、「フレックス（申告の内容）」シートに日ごとの勤務を書きます。"
                 "合計を155時間にし、申告年月日を入れて所属長に伝えます（承認の後で書き直したら、申告年月日も書き直した日にします）。\n"
                 "・所属長：内容を確かめ、承認の欄で「承認」を選び、割振年月日を入れます。状態が「承認済み」の単位期間が、管理者に取り込まれます。"
                 "所属・所属長氏名は職員名簿から自動で入ります（違うときは上書きできます）。")
    per["F3"].font = f_note
    per["F3"].alignment = Alignment(wrap_text=True, vertical="top")
    heads = ["期間", "開始日", "終了日", "合計", "155時間との差", "申告年月日", "承認", "割振年月日", "状態"]
    for col, text in zip("ABCDEFGHI", heads):
        c = per[f"{col}{P_HEAD}"]
        c.value, c.font, c.fill, c.border, c.alignment = text, f_head, fill_head, border, center
    per.row_dimensions[P_HEAD].height = 22
    per[f"A{P_HEAD - 1}"] = "単位期間（4週間）ごとの申告と承認"
    per[f"A{P_HEAD - 1}"].font = f_bold
    for k in range(1, N_PERIODS + 1):
        r = P_FIRST + k - 1
        tm = f"{per_col(3)}{k + 1}"      # 合計（分）
        values = {
            PC_NO: k,
            PC_END: f'=IF(ISNUMBER({PC_START}{r}),{PC_START}{r}+{PERIOD - 1},"")',
            PC_TOTAL: f'=IF({PC_START}{r}="","",{hm_text(tm)})',
            PC_DIFF: f'=IF({PC_START}{r}="","",{diff_text(tm)})',
            PC_STATUS: f"={per_col(8)}{k + 1}",
        }
        for col in "ABCDEFGHI":
            c = per[f"{col}{r}"]
            if col in values:
                c.value = values[col]
            c.font, c.border = f_small, border
            c.alignment = left if col == PC_STATUS else center
            c.fill = fill_staff if col in (PC_START, PC_DECL) else fill_boss if col in (PC_APPROVE, PC_ALLOC) else fill_calc
            if col in (PC_START, PC_DECL, PC_APPROVE, PC_ALLOC):
                c.protection = UNLOCK
            if col in (PC_START, PC_END, PC_DECL, PC_ALLOC):
                c.number_format = "yyyy/m/d"
        per.row_dimensions[r].height = 20
    per.freeze_panes = f"A{P_FIRST}"
    # 開始日・申告年月日・割振年月日は日付だけ。承認は一覧から
    dv = DataValidation(type="date", operator="greaterThan", formula1="43831", allow_blank=True, imeMode="halfAlpha")
    dv.errorTitle, dv.error = "日付を入れてください", "例：2026/4/1"
    per.add_data_validation(dv)
    for col in (PC_START, PC_DECL, PC_ALLOC):
        dv.add(f"{col}{P_FIRST}:{col}{P_LAST}")
    dv = DataValidation(type="list", allow_blank=True, formula1=f'"{APPROVE}"')
    dv.errorTitle, dv.error = "承認", "「承認」を選んでください（取り消すときは消します）。"
    per.add_data_validation(dv)
    dv.add(f"{PC_APPROVE}{P_FIRST}:{PC_APPROVE}{P_LAST}")
    bad_status = lambda cell: f'AND({cell}<>"",{cell}<>"承認済み",{cell}<>"未申告",{cell}<>"承認待ち")'
    per.conditional_formatting.add(f"{PC_STATUS}{P_FIRST}:{PC_STATUS}{P_LAST}",
                                   FormulaRule(formula=[bad_status(f"${PC_STATUS}{P_FIRST}")], fill=fill_bad, font=font_bad))
    # 合計・差：開始日があって155時間でなければ赤
    per.conditional_formatting.add(f"{PC_TOTAL}{P_FIRST}:{PC_DIFF}{P_LAST}", FormulaRule(
        formula=[f'AND(${PC_START}{P_FIRST}<>"",INDEX({H}!$O$2:$O${N_PERIODS + 1},$A{P_FIRST})<>{TOTAL_MIN})'], fill=fill_bad, font=font_bad))
    # 承認したのに、割振年月日がない・申告が承認の後で変わった・155時間でないなど
    st = f"${PC_STATUS}{P_FIRST}"
    per.conditional_formatting.add(f"{PC_APPROVE}{P_FIRST}:{PC_ALLOC}{P_LAST}", FormulaRule(
        formula=[f'AND(${PC_APPROVE}{P_FIRST}="{APPROVE}",OR(${PC_ALLOC}{P_FIRST}="",LEFT({st},3)="承認後",{st}="合計が155時間ではありません",'
                 f'{st}="入力に誤りがあります",{st}="未申告",LEFT({st},3)="開始日"))'],
        fill=fill_bad, font=font_bad))
    # まだ使っていない単位期間（開始日がない行）は薄い灰色の文字
    per.conditional_formatting.add(f"A{P_FIRST}:{PC_STATUS}{P_LAST}", FormulaRule(formula=[f'${PC_START}{P_FIRST}=""'], font=Font(color="A6A6A6")))

    # ===== フレックス（申告の内容） =====
    title(sin, "フレックスタイム制　申告の内容（本人が入力）",
          "黄色：本人が入力（既定の 9:00〜17:45・休憩 12:00〜13:00 が入っています）　灰色：自動　赤：直すところ（理由は「確認」の欄）")
    sin.column_dimensions["A"].width = 4.5
    sin.column_dimensions["B"].width = 10
    sin.column_dimensions["C"].width = 5
    for h, m in TIME_COLS:
        sin.column_dimensions[h].width = 3.6
        sin.column_dimensions[m].width = 3.6
    sin.column_dimensions[HOURS].width = 3.6
    sin.column_dimensions[MINS].width = 3.6
    sin.column_dimensions[NOTE].width = 22
    sin.column_dimensions[CHECK].width = 17
    for c in (OFF, WEEKEND, T_START, T_END, T_BREAK, T_WORK, T_BREAKS):
        sin.column_dimensions[c].hidden = True
    sin.column_dimensions["AI"].width = 2         # 日ごとの行では、その日の承認済みの休暇の行の番号（文字は見せない）
    # 左上：所属・氏名・表の最初の日（自動）
    for r, (label, value) in enumerate([
        ("所属", f"={P}!${P_SECTION[0]}${P_SECTION[1:]}&\"\""),
        ("氏名", f"={MG}!$B$2&\"\""),
        ("表の最初の日", f"=$B${FIRST}"),
    ], start=3):
        sin[f"A{r}"] = label
        sin.merge_cells(f"A{r}:C{r}")
        sin.merge_cells(f"D{r}:M{r}")
        sin[f"A{r}"].font, sin[f"A{r}"].alignment = f_bold, center
        sin[f"D{r}"] = value
        for col in "ABCDEFGHIJKLM":
            sin[f"{col}{r}"].border = border
        for col in "DEFGHIJKLM":
            sin[f"{col}{r}"].fill = fill_calc
        sin[f"D{r}"].font, sin[f"D{r}"].alignment = f_bold, left
    sin["D5"].number_format = 'yyyy"年"m"月"d"日"'
    # 右上：表示する単位期間（空欄なら今日を含む期間）の合計・155時間との差・状態。日ごとの表を下にスクロールしても見える
    KSEL = "AB3"     # 表示する単位期間の番号（隠し列。T3 が空欄なら、今日までに始まったいちばん新しい単位期間。なければ 1）
    today_k = (f'IFERROR(LOOKUP(2,1/(({prng(PC_START)}<>"")*({prng(PC_START)}<=TODAY())),{prng(PC_NO)}),1)')
    sin[KSEL] = f'=IF(AND(ISNUMBER($T$3),$T$3>=1,$T$3<={N_PERIODS}),INT($T$3),{today_k})'
    k_start = f"INDEX({prng(PC_START)},{KSEL})"
    k_end = f"INDEX({prng(PC_END)},{KSEL})"
    k_min = f"INDEX({H}!$O$2:$O${N_PERIODS + 1},{KSEL})"
    k_status = f"INDEX({H}!$T$2:$T${N_PERIODS + 1},{KSEL})"
    summary = [
        (3, "表示する単位期間", None,
         f'="期間"&{KSEL}&IF(ISNUMBER({k_start}),"：　"&TEXT({k_start},"yyyy/m/d")&"〜"&TEXT({k_end},"yyyy/m/d"),"：　（開始日が入っていません）")'
         f'&IF(ISNUMBER($T$3),"","　（今日を含む期間）")'),
        (4, "合計", None, f'=IF(ISNUMBER({k_start}),{hm_text(k_min)},"")'),
        (5, "155時間との差", None, f'=IF(ISNUMBER({k_start}),{diff_text(k_min)},"")'),
        (6, "状態", None, f'=IF(ISNUMBER({k_start}),{k_status},"")'),
    ]
    for r, label, _, value in summary:
        sin.merge_cells(f"N{r}:S{r}")
        sin[f"N{r}"] = label
        sin[f"N{r}"].font, sin[f"N{r}"].alignment = f_bold, center
        val_from = "V" if r == 3 else "T"
        sin.merge_cells(f"{val_from}{r}:{CHECK}{r}")
        sin[f"{val_from}{r}"] = value
        sin[f"{val_from}{r}"].font, sin[f"{val_from}{r}"].alignment = f_bold, left
        for col in ["N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA"]:
            sin[f"{col}{r}"].border = border
            if col not in "NOPQRS":
                sin[f"{col}{r}"].fill = fill_calc
        sin.row_dimensions[r].height = 18
    sin.merge_cells("T3:U3")      # 表示する単位期間の番号（本人が入れる。空欄なら今日を含む期間）
    for col in ("T", "U"):
        sin[f"{col}3"].fill = fill_staff
        sin[f"{col}3"].protection = UNLOCK
    sin["T3"].alignment, sin["T3"].font = center, f_bold
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2=str(N_PERIODS), allow_blank=True, imeMode="halfAlpha")
    dv.errorTitle, dv.error = "期間", f"1〜{N_PERIODS} の番号を入れてください（空欄なら今日を含む単位期間）。"
    sin.add_data_validation(dv)
    dv.add("T3")
    # A6：表示する単位期間（T3）の最初の日の行へ移るリンク（ブックはマクロなしなので、T3 を変えても表は動かない。押すと移る）
    k_row = f"IFERROR(MATCH({k_start},$B${FIRST}:$B${LAST},0)+{FIRST - 1},\"\")"
    sin.merge_cells("A6:M6")
    sin["A6"] = (f'=IF(ISNUMBER({k_row}),HYPERLINK("#\'{SIN}\'!D"&{k_row},"▼ 期間"&{KSEL}&"の最初の日（"&TEXT({k_start},"m/d")&"）の行へ移る（ここを押す）"),"")')
    sin["A6"].font = Font(name=FONT, size=10, bold=True, color="0563C1", underline="single")
    sin["A6"].alignment = left
    sin.conditional_formatting.add("T5:AA5", FormulaRule(formula=[f'AND($T$5<>"",$T$5<>"155時間ちょうど")'], fill=fill_bad, font=font_bad))
    sin.conditional_formatting.add("T6:AA6", FormulaRule(formula=[bad_status("$T$6")], fill=fill_bad, font=font_bad))
    # 説明（7 行目）
    sin.merge_cells(f"A7:{CHECK}7")
    sin["A7"] = ("・単位期間（4週間）の開始日と申告年月日は「フレックス（単位期間・承認）」シートに入れます。日ごとに勤務開始・終了と休憩を時と分に分けて書き、"
                 "勤務しない日は時刻を消して備考に「勤務時間を割り振らない日」と書きます。コアタイム（10:00〜12:00）と祝日・年末年始は既定のまま変えません。"
                 "右上の T3 に期間の番号を入れると、その単位期間の合計が出ます（空欄なら今日を含む期間）。左の A6 を押すと、その期間の最初の日の行へ移ります。コアタイム免除日など、入力できない欄を変えたいときは総務課へ。")
    sin["A7"].font = f_note
    sin["A7"].alignment = Alignment(wrap_text=True, vertical="center")
    sin.row_dimensions[7].height = 36

    # 見出し（8〜11行目）：8 グループ、9 項目、10 開始・終了、11 時・分
    def head(ref1, ref2, text):
        if ref1 != ref2:
            sin.merge_cells(f"{ref1}:{ref2}")
        sin[ref1] = text
        for row in sin[f"{ref1}:{ref2}"]:
            for c in row:
                c.fill, c.border = fill_head, border
        sin[ref1].font, sin[ref1].alignment = f_head, center
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
        sin.row_dimensions[r].height = 15
    sin.freeze_panes = f"D{FIRST}"
    sin.sheet_view.zoomScale = 90
    per.sheet_view.zoomScale = 100

    for r in range(FIRST, LAST + 1):
        i = r - FIRST
        b = f"$B{r}"
        # その日を含む単位期間の番号（重なっているときは先の単位期間。どれにも入らなければ空欄）
        sin[f"A{r}"] = (f'=IFERROR(INDEX({prng(PC_NO)},MATCH(1,INDEX(({prng(PC_START)}<>"")'
                        f'*({prng(PC_START)}<={b})*({prng(PC_END)}>={b}),0),0)),"")')
        sin[f"B{r}"] = (f"=IF(ISNUMBER({pcell(PC_START, 1)}),{pcell(PC_START, 1)},{START})" if i == 0 else f"=B{r - 1}+1")
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
    sin.conditional_formatting.add(f"A{FIRST}:{CHECK}{LAST}", FormulaRule(
        formula=[f'$A{FIRST}=""'], font=Font(color="A6A6A6")))
    sin.conditional_formatting.add(f"A{FIRST}:{CHECK}{LAST}", FormulaRule(
        formula=[f'AND($A{FIRST}<>"",$B{FIRST}=INDEX({prng(PC_END)},$A{FIRST}))'],
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
        alloc_k = f"INDEX({prng(PC_ALLOC)},{kk})"
        hyo[f"{NOWORK_COL}{r}"] = (f'=IFERROR(IF(AND(ISNUMBER(SEARCH("割り振らない",{S}!${NOTE}{r})),INDEX({prng(PC_APPROVE)},{kk})="{APPROVE}",'
                                    f'ISNUMBER({alloc_k}),N(INDEX({prng(PC_DECL)},{kk}))<={alloc_k}),{alloc_k},""),"")')
    for i, h in enumerate(per_heads):
        hyo.cell(row=1, column=PL + i, value=h)
    for k in range(1, N_PERIODS + 1):
        t = k + 1
        c = lambda i: get_column_letter(PL + i)
        start, end = pcell(PC_START, k), pcell(PC_END, k)
        total = f"SUMIFS({S}!${T_WORK}${FIRST}:${T_WORK}${LAST},{S}!$A${FIRST}:$A${LAST},{k})"
        errors = (f'COUNTIFS({S}!$A${FIRST}:$A${LAST},{k},{S}!${CHECK}${FIRST}:${CHECK}${LAST},"?*")'
                  f'-COUNTIFS({S}!$A${FIRST}:$A${LAST},{k},{S}!${CHECK}${FIRST}:${CHECK}${LAST},"{NOWORK_MARK}*")')
        # 開始日の誤り：表の日の外、前の単位期間と重なる（開始日は前の単位期間の終了日より後）
        prev_end = pcell(PC_END, k - 1) if k > 1 else '""'
        d1, dn = f"{S}!$B${FIRST}", f"{S}!$B${LAST}"     # 日ごとの表の最初の日・最後の日
        bad_start = (f'IF(OR({start}<{d1},{end}>{dn}),"開始日がこの表の日の外です（"&TEXT({d1},"yyyy/m/d")&"から"&TEXT({dn},"yyyy/m/d")&"）",'
                     f'IF(AND(ISNUMBER({prev_end}),{start}<={prev_end}),"開始日が前の単位期間と重なっています",""))')
        declared = pcell(PC_DECL, k)
        approved = pcell(PC_APPROVE, k)
        alloc = pcell(PC_ALLOC, k)
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
                f"={P}!${P_SECTION[0]}${P_SECTION[1:]}&\"\"", f"={P}!${P_BOSS[0]}${P_BOSS[1:]}&\"\"", f"={P}!${P_BOSS2[0]}${P_BOSS2[1:]}&\"\""]
        for i, v in enumerate(vals):
            hyo.cell(row=t, column=PL + i, value=v)
    t = Table(displayName="フレックス期間", ref=f"L1:{get_column_letter(PL + len(per_heads) - 1)}{N_PERIODS + 1}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
    hyo.add_table(t)

    for ws, area in ((per, f"A1:{PC_STATUS}{P_LAST}"), (sin, f"A1:{CHECK}{LAST}")):
        ws.protection.sheet = True
        ws.protection.formatColumns = False
        ws.protection.formatRows = False
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_area = area
    sin.print_title_rows = "8:11"
    return [PER, SIN]
