"""照合用Excel.xlsx を作る（様式と計算式）。

使い方: python make_matching.py 出力先.xlsx
通常は同じフォルダの build.ps1 から実行する（Excel での再計算までまとめて行う）。
"""
import sys
from datetime import date, time
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.pagebreak import Break
from openpyxl.workbook.defined_name import DefinedName

OUT = sys.argv[1]
FONT = "ＭＳ Ｐゴシック"
STAFF_ROWS = 500      # 職員名簿の行数（2〜501行目）
DATA_ROWS = 5000      # 出力データの行数（2〜5001行目）
NAME_LAST = STAFF_ROWS + 1
DATA_LAST = DATA_ROWS + 1
BATCH = 200          # 一括印刷できる人数

f_base = Font(name=FONT, size=10)
f_bold = Font(name=FONT, size=10, bold=True)
f_head = Font(name=FONT, size=10, bold=True, color="FFFFFF")
f_title = Font(name=FONT, size=14, bold=True, color="1E3A5F")
f_note = Font(name=FONT, size=9, color="555555")
fill_head = PatternFill("solid", fgColor="1E3A5F")
fill_input = PatternFill("solid", fgColor="FFF2CC")
fill_calc = PatternFill("solid", fgColor="F2F2F2")
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center")
left = Alignment(horizontal="left", vertical="center")

# 入力シートには数式を置かない。参照は INDEX(列, ROW()) にして、
# 行ごと貼り付け・行の削除をしても参照がずれたり #REF! にならないようにする。
P = "'出力データ貼付'"
N = "'職員名簿'"
K = "'計算用'"
LV = "'年次休暇'"
LEAVE_LAST = 1001     # 年次休暇シートの行数（2〜1001行目）
CARRY_MAX = 20        # 翌年に繰り越せる日数の上限
GRANT_DEFAULT = 20    # 本年分が空欄のときの日数（通年で勤務する場合の付与日数）
# 年の中途採用・退職者の本年分（在職期間が 1月に達するまで, 1月を超え2月に達するまで, … 11月を超え1年未満）
PRORATED = [2, 3, 5, 7, 8, 10, 12, 13, 15, 17, 18, 20]
TOTAL_MAX = 40        # 1年に使える日数（繰越分＋本年分）の上限
DAY_MIN = 465         # 年次休暇の1日＝7時間45分＝465分（15分単位で取得できる）
MH = "'月別取得時間'"
HOL = "'祝日'!$A$2:$A$200"
MH_LAST = 1501        # 月別取得時間シートの行数（2〜1501行目）
HOUR_CATS = ["病気休暇", "看護休暇"]


def key(expr):
    # 文字列・数値・前後の空白・先頭の0の違いを吸収した照合キー
    return f'IF({expr}="","",IFERROR(VALUE(TRIM({expr}))&"",TRIM({expr})))'


def paste_col(col, row_expr="ROW()"):
    return f"INDEX({P}!${col}:${col},{row_expr})"


def name_col(col, row_expr="ROW()"):
    return f"INDEX({N}!${col}:${col},{row_expr})"


def off_day(d):
    """土曜・日曜・祝日（祝日シート）・年末年始（12/29〜1/3）なら TRUE になる式。"""
    return (f'OR(WEEKDAY({d})=1,WEEKDAY({d})=7,COUNTIF({HOL},{d})>0,'
            f'AND(MONTH({d})=12,DAY({d})>=29),AND(MONTH({d})=1,DAY({d})<=3))')


def leave_col(col, row_expr="ROW()"):
    return f"INDEX({LV}!${col}:${col},{row_expr})"


def days_text(m):
    """分（m はセル参照）を「15日3時間30分」の形の文字にする。1日＝7時間45分。"""
    rest = f"MOD({m},{DAY_MIN})"
    return (f'IF({m}="","",INT({m}/{DAY_MIN})&"日"'
            f'&IF(INT({rest}/60)>0,INT({rest}/60)&"時間","")'
            f'&IF(MOD({rest},60)>0,MOD({rest},60)&"分",""))')


wb = Workbook()


def header(ws, row, labels, widths=None, fill=fill_head):
    for i, label in enumerate(labels, start=1):
        c = ws.cell(row=row, column=i, value=label)
        c.font, c.fill, c.alignment, c.border = f_head, fill, center, border
        if widths:
            ws.column_dimensions[get_column_letter(i)].width = widths[i - 1]


def style_range(ws, r1, r2, c1, c2, fill=None, fmt=None, align=None):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = f_base
            cell.border = border
            if fill:
                cell.fill = fill
            if fmt:
                cell.number_format = fmt
            if align:
                cell.alignment = align


# ===== 使い方 =====
ws = wb.active
ws.title = "使い方"
ws.column_dimensions["A"].width = 4
ws.column_dimensions["B"].width = 96
ws["B1"] = "出勤記録 照合用Excel（管理者のみ保管）"
ws["B1"].font = f_title
lines = [
    ("手順", True),
    ("1. 「職員名簿」シートに、出勤用IDと氏名を入力します（初回のみ。異動時に更新）。", False),
    ("2. 出勤管理システムの管理者画面から「Excel出力」を押し、attendance.xlsx をダウンロードします。", False),
    ("3. attendance.xlsx のデータ（見出しを除く）をコピーし、「出力データ貼付」シートの A2 に貼り付けます。", False),
    ("    ※ セル範囲のコピーでも、行ごと（行番号を選んで）のコピーでも大丈夫です。", False),
    ("    ※ 前回のデータは、2行目から下を選んで Delete キーで消すか、行ごと削除してから貼り付けてください。", False),
    ("4. 「照合結果」シートに氏名つきの出勤記録が表示されます。見出しのフィルターで氏名・日付を絞り込めます。", False),
    ("5. 「月次一覧」シートの B1 に対象年月（例：2026/10）を入力すると、職員ごと・日ごとの出勤（○）が表示されます。", False),
    ("6. 「出勤簿」シートの右側（AF2・AF3）で対象年月と出勤用ID（一覧から選択）を指定すると、その職員の1か月分の出勤簿（A4縦）が表示され、そのまま印刷できます。", False),
    ("    ※ 上の欄の題名は「令和〇年出勤簿」と自動で表示されます。所属の欄は記入用の空欄です。", False),
    ("    ※ 病気休暇・看護休暇の各月の取得時間は「月別取得時間」シートの内容が表示されます。", False),
    ("    ※ 年次休暇の取得日数は「年次休暇」シートの内容が表示されます。", False),
    ("    ※ 出勤欄の○は自動で表示されます。年次休暇・病気休暇などの欄と備考は記入用の空欄です。", False),
    (f"7. 全員分をまとめて印刷するときは、「一括印刷」シートの AF2 に対象年月を入力し、そのまま印刷（Ctrl+P）します。職員名簿の全員分の出勤簿が1人1枚で印刷されます（{BATCH}人まで）。", False),
    ("", False),
    ("色の意味", True),
    ("黄色のセル … 入力欄（ここに入力・貼り付けします）", False),
    ("灰色のセル … 計算式（自動で表示されます。触らないでください）", False),
    ("", False),
    ("注意", True),
    ("・入力するシート（職員名簿・出力データ貼付）には計算式がないので、どんな貼り付け方をしても壊れません。", False),
    ("・出勤用IDは、文字列でも数値でも照合できます。前後の空白や、先頭の0が消えた場合（0425 → 425）も同じIDとして扱います。", False),
    ("・照合結果に「★名簿に未登録」と出た場合は、職員名簿にそのIDがありません。名簿を追加・修正してください。", False),
    ("・月次一覧の「重複」欄に表示が出た場合は、職員名簿に同じIDが2回以上登録されています。", False),
    (f"・職員名簿は{STAFF_ROWS}人分、出力データは{DATA_ROWS}件分まで対応しています。", False),
    ("・月次一覧・出勤簿では、土曜・日曜・祝日・年末年始（12/29〜1/3）は赤で表示されます。祝日は「祝日」シートの一覧を使います。", False),
    ("・「計算用」シートは非表示にしています。消したり変更したりしないでください。", False),
    ("・年次休暇シートには、職員ごと・年ごとに本年分と本年使用（日と時間。15分単位、1日＝7時間45分）を入力します。繰越分は前年の翌年繰越が自動で入ります（活用開始時のみ手入力）。翌年に繰り越せるのは本年分の残りだけで（一度繰り越した日数は更に繰り越せない）、上限は20日、1日未満の端数は切り捨てです。", False),
    ("・このファイルには氏名が含まれます。クラウドやメールで共有せず、管理者のPC内だけで保管してください。", False),
    ("・各シートの2行目に入っている（例）のデータは、使い始める前に削除してください。", False),
]
for i, (text, bold) in enumerate(lines, start=3):
    c = ws.cell(row=i, column=2, value=text)
    c.font = f_bold if bold else f_base
    if text.startswith("黄色"):
        c.fill = fill_input
    if text.startswith("灰色"):
        c.fill = fill_calc

# ===== 職員名簿（入力専用） =====
ws = wb.create_sheet("職員名簿")
header(ws, 1, ["出勤用ID", "氏名", "備考"], [12, 18, 30])
style_range(ws, 2, NAME_LAST, 1, 3, fill=fill_input)
style_range(ws, 2, NAME_LAST, 1, 1, fmt="@", align=center)
ws["A2"], ws["B2"], ws["C2"] = "1234", "（例）田中 太郎", "この行は例です。削除してください"
ws.freeze_panes = "A2"

# ===== 出力データ貼付（入力専用） =====
ws = wb.create_sheet("出力データ貼付")
header(ws, 1, ["出勤用ID", "QR設置場所", "出勤日", "出勤時刻", "登録日時"], [12, 14, 14, 12, 26])
style_range(ws, 2, DATA_LAST, 1, 5, fill=fill_input, fmt="@")
for col, v in zip("ABCDE", ["1234", "K", "2026-10-01", "09:02:15", "2026-10-01T00:02:15.000Z"]):
    ws[f"{col}2"] = v
ws["G1"] = "← attendance.xlsx のデータ（見出しを除く）をコピーして、このシートの A2 に貼り付けます（2行目は例です）"
ws["G1"].font = f_note
ws.freeze_panes = "A2"

# ===== 計算用（非表示）：職員名簿の照合キーと重複 =====
ws = wb.create_sheet("計算用")
header(ws, 1, ["名簿の照合キー", "重複", "", "", "", "名簿の連番", "",
                "年休キー（ID|年）", "前年の行", "繰越分（分）", "合計（分）", "翌年繰越（分）", "", "月別キー（ID|年|区分）",
                "本年使用（分）", "本年分（日）", "在職期間の始め", "在職期間の終わり", "在職期間の区分"],
       [16, 8, 4, 4, 4, 10, 4, 16, 8, 10, 10, 10, 4, 22, 10, 10, 12, 12, 10],
       fill=PatternFill("solid", fgColor="7F7F7F"))
for r in range(2, NAME_LAST + 1):
    ws[f"A{r}"] = "=" + key(name_col("A"))
    ws[f"B{r}"] = f'=IF(A{r}="","",IF(COUNTIF($A$2:$A${NAME_LAST},A{r})>1,"重複",""))'
    ws[f"F{r}"] = f'=IF(A{r}="","",COUNTIF($A$2:A{r},"?*"))'
# 年次休暇シートの各行（同じ行番号）について、繰越分・合計・翌年繰越を「分」で計算する（1日＝7時間45分＝465分）。
# 繰越分 = 前年の行の翌年繰越（前年の行がなければ、その行に手入力した繰越分）
# 合計 = 繰越分＋本年分（上限40日）
# 翌年繰越 = 残り（合計−本年使用）の1日未満を切り捨てた日数。上限20日（繰越分と本年分は区別しない）
for r in range(2, LEAVE_LAST + 1):
    ws[f"H{r}"] = f'=IF(OR({key(leave_col("A"))}="",{leave_col("B")}=""),"",{key(leave_col("A"))}&"|"&{leave_col("B")})'
    ws[f"I{r}"] = f'=IF(H{r}="","",IFERROR(MATCH({key(leave_col("A"))}&"|"&({leave_col("B")}-1),$H$2:$H${LEAVE_LAST},0),""))'
    y = leave_col("B")
    hire, leave = leave_col("G"), leave_col("H")
    ws[f"Q{r}"] = f'=IF(H{r}="","",IF({hire}="",DATE({y},1,1),MAX(DATE({y},1,1),{hire})))'
    ws[f"R{r}"] = f'=IF(H{r}="","",IF({leave}="",DATE({y},12,31),MIN(DATE({y},12,31),{leave})))'
    ws[f"S{r}"] = (f'=IF(H{r}="","",IF(R{r}<Q{r},-1,DATEDIF(Q{r},R{r}+1,"m")'
                   f'-IF(EDATE(Q{r},DATEDIF(Q{r},R{r}+1,"m"))=R{r}+1,1,0)))')
    ws[f"P{r}"] = (f'=IF(H{r}="","",IF({leave_col("C")}&""<>"",N({leave_col("C")}),'
                   f'IF(AND({hire}="",{leave}=""),{GRANT_DEFAULT},'
                   f'IF(S{r}<0,0,CHOOSE(MIN(S{r},11)+1,{",".join(str(d) for d in PRORATED)})))))')
    ws[f"O{r}"] = f'=IF(H{r}="","",ROUND(N({leave_col("D")})*{DAY_MIN}+N({leave_col("E")})*1440,0))'
    # 活用開始時に手入力する繰越分も、繰越のルールどおり1日未満は切り捨て・上限20日
    ws[f"J{r}"] = (f'=IF(H{r}="","",IF(I{r}<>"",INDEX($L$2:$L${LEAVE_LAST},I{r}),'
                   f'MIN({CARRY_MAX},INT(N({leave_col("F")})))*{DAY_MIN}))')
    ws[f"K{r}"] = f'=IF(H{r}="","",MIN({TOTAL_MAX * DAY_MIN},J{r}+P{r}*{DAY_MIN}))'
    ws[f"L{r}"] = (f'=IF(H{r}="","",MIN({CARRY_MAX * DAY_MIN},'
                   f'INT(MAX(0,K{r}-O{r})/{DAY_MIN})*{DAY_MIN}))')
for r in range(2, MH_LAST + 1):
    mid = f"INDEX({MH}!$A:$A,ROW())"
    ws[f"N{r}"] = (f'=IF(OR({key(mid)}="",INDEX({MH}!$B:$B,ROW())="",INDEX({MH}!$C:$C,ROW())=""),"",'
                   f'{key(mid)}&"|"&INDEX({MH}!$B:$B,ROW())&"|"&TRIM(INDEX({MH}!$C:$C,ROW())))')
ws.sheet_state = "hidden"

# ===== 照合結果 =====
# 表示：A No, B 出勤用ID, C 氏名, D QR設置場所, E 出勤日, F 出勤時刻
# 隠し：G 照合キー, H 出勤日キー（月次一覧・出勤簿が参照する）
ws = wb.create_sheet("照合結果")
header(ws, 1, ["No", "出勤用ID", "氏名", "QR設置場所", "出勤日", "出勤時刻", "照合キー", "出勤日キー"],
       [7, 12, 20, 14, 14, 12, 12, 12])
style_range(ws, 2, DATA_LAST, 1, 8, fill=fill_calc, align=center)
for r in range(2, DATA_LAST + 1):
    match = f"MATCH(G{r},{K}!$A$2:$A${NAME_LAST},0)"
    ws[f"G{r}"] = "=" + key(paste_col("A"))
    ws[f"H{r}"] = f'=IF(G{r}="","",TEXT({paste_col("C")},"yyyy-mm-dd"))'
    ws[f"A{r}"] = f'=IF(G{r}="","",ROW()-1)'
    ws[f"B{r}"] = f'=IF(G{r}="","",TRIM({paste_col("A")}&""))'
    ws[f"C{r}"] = f'=IF(G{r}="","",IFERROR(INDEX({N}!$B$2:$B${NAME_LAST},{match})&"","★名簿に未登録"))'
    ws[f"D{r}"] = f'=IF(G{r}="","",{paste_col("B")}&"")'
    ws[f"E{r}"] = f"=H{r}"
    ws[f"F{r}"] = f'=IF(G{r}="","",TEXT({paste_col("D")},"hh:mm:ss"))'
    ws[f"C{r}"].alignment = left
for col in "GH":
    ws.column_dimensions[col].hidden = True
ws.auto_filter.ref = f"A1:F{DATA_LAST}"
ws.freeze_panes = "A2"
ws.conditional_formatting.add(
    f"C2:C{DATA_LAST}",
    FormulaRule(formula=['C2="★名簿に未登録"'], font=Font(name=FONT, color="C00000", bold=True)),
)

# ===== 月次一覧 =====
ws = wb.create_sheet("月次一覧")
ws["A1"] = "対象年月"
ws["A1"].font = f_bold
ws["B1"] = date(2026, 10, 1)
ws["B1"].number_format = 'yyyy"年"m"月"'
ws["B1"].fill, ws["B1"].font, ws["B1"].border = fill_input, f_bold, border
ws["C1"] = "← 対象年月を入力（例：2026/10）"
ws["C1"].font = f_note
HOL = "'祝日'!$A$2:$A$200"
M1 = "DATE(YEAR($B$1),MONTH($B$1),1)"
header(ws, 3, ["出勤用ID", "氏名", "重複", "出勤日数"], [12, 18, 6, 9])
first_day_col = 5
for d in range(31):
    col = first_day_col + d
    L = get_column_letter(col)
    ws.column_dimensions[L].width = 4.2
    c = ws.cell(row=3, column=col, value=f'=IF(MONTH({M1}+{d})=MONTH({M1}),{M1}+{d},"")')
    c.number_format = "d"
    c.font, c.fill, c.alignment, c.border = f_head, fill_head, center, border
    w = ws.cell(row=4, column=col, value=f'=IF({L}3="","",IF(COUNTIF({HOL},{L}3)>0,"祝",TEXT({L}3,"aaa")))')
    w.font, w.alignment, w.border, w.fill = f_bold, center, border, fill_calc
for col in range(1, first_day_col):
    c = ws.cell(row=4, column=col)
    c.fill, c.border = fill_calc, border
first_L = get_column_letter(first_day_col)
last_L = get_column_letter(first_day_col + 30)
LAST = 4 + STAFF_ROWS
style_range(ws, 5, LAST, 1, first_day_col + 30, fill=fill_calc, align=center)
for r in range(5, LAST + 1):
    src = r - 3  # 職員名簿・計算用の行（5行目 → 2行目）
    row_expr = f"ROW()-3"
    ws[f"A{r}"] = f'=IF({K}!A{src}="","",TRIM({name_col("A", row_expr)}&""))'
    ws[f"B{r}"] = f'=IF(A{r}="","",{name_col("B", row_expr)}&"")'
    ws[f"C{r}"] = f"={K}!B{src}"
    ws[f"D{r}"] = f'=IF(A{r}="","",COUNTIF({first_L}{r}:{last_L}{r},"○"))'
    ws[f"B{r}"].alignment = left
    for d in range(31):
        L = get_column_letter(first_day_col + d)
        ws[f"{L}{r}"] = (
            f'=IF(OR($A{r}="",{L}$3=""),"",'
            f'IF(COUNTIFS(照合結果!$G$2:$G${DATA_LAST},{K}!$A{src},照合結果!$H$2:$H${DATA_LAST},TEXT({L}$3,"yyyy-mm-dd"))>0,"○",""))'
        )
# 土曜・日曜・祝日は同じ色
sun_hol = f'AND({first_L}$3<>"",{off_day(f"{first_L}$3")})'
ws.conditional_formatting.add(
    f"{first_L}3:{last_L}3",
    FormulaRule(formula=[sun_hol], fill=PatternFill("solid", fgColor="C0392B", bgColor="C0392B"), stopIfTrue=True),
)
ws.conditional_formatting.add(
    f"{first_L}4:{last_L}{LAST}",
    FormulaRule(formula=[sun_hol], fill=PatternFill("solid", fgColor="F8D7D3", bgColor="F8D7D3"),
                font=Font(name=FONT, color="9C2A1F"), stopIfTrue=True),
)
ws.conditional_formatting.add(
    f"C5:C{LAST}",
    FormulaRule(formula=['C5="重複"'], font=Font(name=FONT, color="C00000", bold=True)),
)
ws.freeze_panes = f"{first_L}5"
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_title_rows = "3:4"

# ===== 出勤簿・一括印刷（どちらも同じ様式。1人1か月・A4縦1枚） =====
# 列の作り：A=日付、B=曜日、C〜Z=細い列24本。日付の表の各列（出勤〜備考の12列）は細い列2本ぶん、
# 上の欄の育児時間（1月〜12月）は細い列1本ずつ（O〜Z）を使う。
# 日付の表の行はセルを結合せず、2本ごとに外枠を引いて1つの欄に見せる（一括印刷で結合が増えすぎないように）。
R = "照合結果"
BLOCK = 47            # 1人分の行数（改ページ単位）
TOP = 12              # 日付の表の見出し（下の段）の行（ブロック先頭からの位置）。上の段はその1行上
f_label = Font(name=FONT, size=10, bold=True, color="1E3A5F")
fill_label = PatternFill("solid", fgColor="E8EEF5")
fill_white = PatternFill(fill_type=None)
rng_key = f"{R}!$G$2:$G${DATA_LAST}"
rng_day = f"{R}!$H$2:$H${DATA_LAST}"
# 日付の表の列（見出し, 細い列の本数）。備考は広め（4本）にして、細い列は全部で24本（C〜Z）
TABLE = [("出勤", 2), ("年次\n休暇", 2), ("病気\n休暇", 2), ("特別休暇\n（種別）", 2), ("特別休暇\n（時間）", 2),
         ("介護\n休暇", 2), ("介護\n時間", 2), ("職専免\n（種別）", 2), ("職専免\n（時間）", 2), ("欠勤", 2), ("備考", 4)]
SUB = [get_column_letter(3 + i) for i in range(sum(w for _, w in TABLE))]   # C〜Z
SPANS = []                          # 日付の表の各列が使う細い列（最初, 最後）
_i = 0
for _, _w in TABLE:
    SPANS.append((SUB[_i], SUB[_i + _w - 1]))
    _i += _w
LAST_COL = SUB[-1]                                                 # Z
N_COLS = 2 + len(SUB)                                              # 26
SUB_W = 3.8
HID_COUNT, HID = "AB", "AC"     # 隠し列：AB=その日の登録件数、AC=照合キー・名簿の行位置・対象月の1日
CTRL_L, CTRL_V = "AE", "AF"     # 印刷範囲の外：対象年月・出勤用ID などの指定欄
wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)
thin_side = Side(style="thin", color="BFBFBF")
none_side = Side(style=None)


def setup_report_sheet(ws):
    ws.column_dimensions["A"].width = 6.5
    ws.column_dimensions["B"].width = 4.5
    for col in SUB:
        ws.column_dimensions[col].width = SUB_W
    ws.column_dimensions["AA"].width = 2
    for col in (HID_COUNT, HID):
        ws.column_dimensions[col].width = 8
        ws.column_dimensions[col].hidden = True
    ws.column_dimensions["AD"].width = 2
    ws.column_dimensions[CTRL_L].width = 12
    ws.column_dimensions[CTRL_V].width = 14
    ws.sheet_view.showGridLines = False
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "portrait"
    ws.print_options.horizontalCentered = True
    ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.5, bottom=0.4)


def box(ws, first, last, value=None, font=None, fill=fill_label, align=center):
    if first != last:
        ws.merge_cells(f"{first}:{last}")
    ws[first] = value
    for row in ws[f"{first}:{last}"]:
        for c in row:
            c.border = border
            if fill is not None:
                c.fill = fill
    ws[first].font = font or f_label
    ws[first].alignment = align


def span_cells(ws, r, k):
    """日付の表の k 列目にあたる細い列（2本または4本）。外枠だけ引いて1つの欄に見せる。"""
    first, last = SPANS[k]
    cells = [c for c in ws[f"{first}{r}:{last}{r}"][0]]
    for i, c in enumerate(cells):
        c.border = Border(left=thin_side if i == 0 else none_side,
                          right=thin_side if i == len(cells) - 1 else none_side,
                          top=thin_side, bottom=thin_side)
        c.font = f_base
    return cells


def report_block(ws, t, month_ref, key_formula, idx_formula):
    """t 行目から1人分の出勤簿を書く。month_ref は対象年月のセル（印刷範囲の外）。
    隠し列 AC：AC{t}=照合キー、AC{t+1}=名簿の行位置、AC{t+2}=対象月の1日"""
    P1, P2, P3 = f"${HID}${t}", f"${HID}${t + 1}", f"${HID}${t + 2}"
    ws[f"{HID}{t}"] = key_formula
    ws[f"{HID}{t + 1}"] = idx_formula
    ws[f"{HID}{t + 2}"] = f'=IF({month_ref}="","",DATE(YEAR({month_ref}),MONTH({month_ref}),1))'

    # 1行目：令和〇年出勤簿
    ws.merge_cells(f"A{t}:{LAST_COL}{t}")
    ws[f"A{t}"] = f'=IF({P3}="","出勤簿","令和"&IF(YEAR({P3})-2018=1,"元",YEAR({P3})-2018)&"年出勤簿")'
    ws[f"A{t}"].font = Font(name=FONT, size=18, bold=True, color="1E3A5F")
    ws[f"A{t}"].alignment = center
    ws.row_dimensions[t].height = 32

    # 2〜7行目：左に所属（空欄）とその下に氏名、右に各月の取得時間と累計（病気休暇・看護休暇 × 1月〜12月）
    # 取得時間は月別取得時間シートのその年・区分の行から表示する（AC{t+4}〜AC{t+5}=その行の位置）
    value_font = Font(name=FONT, size=11, bold=True)
    box(ws, f"A{t + 1}", f"B{t + 3}", "所属")
    box(ws, f"C{t + 1}", f"K{t + 3}", None, value_font, fill_white, left)
    box(ws, f"A{t + 4}", f"B{t + 6}", "氏名")
    box(ws, f"C{t + 4}", f"K{t + 6}",
        f'=IF({P1}="","",IF({P2}="","★名簿に未登録",INDEX({N}!$B$2:$B${NAME_LAST},{P2})&""))',
        value_font, fill_white, left)
    small_label = Font(name=FONT, size=8, bold=True, color="1E3A5F")
    hour_font = Font(name=FONT, size=8)
    # L列は2〜7行目を結合して縦書きの「取得時間」。M:N は区分名、O〜Z は月
    box(ws, f"L{t + 1}", f"L{t + 6}", "取得時間", f_label,
        align=Alignment(horizontal="center", vertical="center", text_rotation=255))
    # 見出し（1月〜12月）は2行分の高さ、その下に区分ごとに「その月」と「累計」の2行
    box(ws, f"M{t + 1}", f"N{t + 2}", None, small_label)
    for i, col in enumerate(SUB[12:]):          # O〜Z
        box(ws, f"{col}{t + 1}", f"{col}{t + 2}", f"{i + 1}月", small_label)
    for j, cat in enumerate(HOUR_CATS):
        r = t + 3 + j * 2                       # その月の取得時間
        rc = r + 1                              # 1月からその月までの累計
        MR = f"${HID}${t + 4 + j}"
        ws[f"{HID}{t + 4 + j}"] = (f'=IF(OR({P1}="",{P3}=""),"",'
                                   f'IFERROR(MATCH({P1}&"|"&YEAR({P3})&"|{cat}",{K}!$N$2:$N${MH_LAST},0),""))')
        box(ws, f"M{r}", f"N{r}", cat, small_label)
        box(ws, f"M{rc}", f"N{rc}", "（累計）", small_label)
        first_m = f"INDEX({MH}!$D$2:$D${MH_LAST},{MR})"
        for i, col in enumerate(SUB[12:]):
            mcol = get_column_letter(4 + i)     # 月別取得時間シートの i+1 月の列（D〜O）
            v = f"INDEX({MH}!${mcol}$2:${mcol}${MH_LAST},{MR})"
            box(ws, f"{col}{r}", f"{col}{r}", f'=IF({MR}="","",IF({v}&""="","",{v}))', hour_font, fill_white)
            # 累計は対象年月の月まで表示（それより後の月と、累計が0の月は空欄）
            total = f"SUM({first_m}:{v})"
            box(ws, f"{col}{rc}", f"{col}{rc}",
                f'=IF(OR({MR}="",{i + 1}>MONTH({P3})),"",IF({total}=0,"",{total}))', hour_font, fill_white)
            ws[f"{col}{r}"].number_format = ws[f"{col}{rc}"].number_format = "[h]:mm"   # 分単位なので「時間:分」
    for r in range(t + 1, t + 7):
        ws.row_dimensions[r].height = 18
    ws.row_dimensions[t + 7].height = 8

    # 9〜10行目：年次休暇の取得日数（年次休暇シートのその年の行から表示。AC{t+3}=その行の位置）
    r = t + 8
    LR = f"${HID}${t + 3}"
    ws[f"{HID}{t + 3}"] = (f'=IF(OR({P1}="",{P3}=""),"",'
                           f'IFERROR(MATCH({P1}&"|"&YEAR({P3}),{K}!$H$2:$H${LEAVE_LAST},0),""))')

    def leave_value(col):
        # 計算用シートの「分」の値を「15日3時間30分」の形で表示する
        return f'=IF({LR}="","",IFERROR({days_text(f"INDEX({K}!${col}$2:${col}${LEAVE_LAST},{LR})")},""))'

    granted = f"INDEX({K}!$P$2:$P${LEAVE_LAST},{LR})"
    values = [
        leave_value("J"),                                                       # 繰越分
        f'=IF({LR}="","",IF({granted}&""="","",{granted}&"日"))',             # 本年分
        leave_value("K"),                                                       # 合計
        leave_value("O"),                                                       # 本年使用
        leave_value("L"),                                                       # 翌年繰越
    ]
    box(ws, f"A{r}", f"B{r + 1}", "年次休暇\n取得日数", align=wrap)
    for i, text in enumerate(["繰越分", "本年分", "合計", "本年使用", "翌年繰越"]):
        c1, c2 = SUB[i * 4], SUB[i * 4 + 3]       # C:F, G:J, K:N, O:R, S:V
        box(ws, f"{c1}{r}", f"{c2}{r}", text)
        box(ws, f"{c1}{r + 1}", f"{c2}{r + 1}", values[i], Font(name=FONT, size=11), fill_white)
    ws.row_dimensions[r].height = 18
    ws.row_dimensions[r + 1].height = 24
    ws.row_dimensions[t + 10].height = 8

    # 日付の表（見出しは2段：上の段に「休暇等の日数」（年次休暇〜欠勤）、日付・曜日・出勤・備考は2段分）
    hr = t + TOP
    gr = hr - 1
    head_font = Font(name=FONT, size=9, bold=True, color="FFFFFF")
    for col, text in (("A", "日付"), ("B", "曜日")):
        box(ws, f"{col}{gr}", f"{col}{hr}", text, head_font, fill_head, wrap)
    for k, ((text, _), (first, last)) in enumerate(zip(TABLE, SPANS)):
        if k in (0, len(TABLE) - 1):            # 出勤・備考
            box(ws, f"{first}{gr}", f"{last}{hr}", text, head_font, fill_head, wrap)
        else:
            box(ws, f"{first}{hr}", f"{last}{hr}", text, head_font, fill_head, wrap)
    box(ws, f"{SPANS[1][0]}{gr}", f"{SPANS[-2][1]}{gr}", "休暇等の日数", head_font, fill_head, center)
    ws.row_dimensions[gr].height = 18
    ws.row_dimensions[hr].height = 30
    for d in range(31):
        r = hr + 1 + d
        ws[f"A{r}"] = f'=IF(OR({P1}="",{P3}=""),"",IF(MONTH({P3}+{d})=MONTH({P3}),{P3}+{d},""))'
        ws[f"A{r}"].number_format = "m/d"
        ws[f"B{r}"] = f'=IF(A{r}="","",IF(COUNTIF({HOL},A{r})>0,"祝",TEXT(A{r},"aaa")))'
        for col in "AB":
            ws[f"{col}{r}"].font, ws[f"{col}{r}"].border, ws[f"{col}{r}"].alignment = f_base, border, center
        ws[f"{HID_COUNT}{r}"] = f'=IF(A{r}="","",COUNTIFS({rng_key},{P1},{rng_day},TEXT(A{r},"yyyy-mm-dd")))'
        for k in range(len(TABLE)):
            cells = span_cells(ws, r, k)
            if k == 0:      # 出勤の○は2本の真ん中に（「選択範囲内で中央」は両方のセルに指定が必要）
                for c in cells:
                    c.alignment = Alignment(horizontal="centerContinuous", vertical="center")
            else:
                cells[0].alignment = left
        ws[f"C{r}"] = f'=IF(N({HID_COUNT}{r})=0,"","○")'
        ws.row_dimensions[r].height = 19
    ws[f"A{hr + 32}"] = "※ 出勤欄の○は、出勤管理システムの記録（QRコードの読み取りによる出勤登録）によります。"
    ws[f"A{hr + 32}"].font = f_note


def report_colors(ws, first_row, last_row):
    # 土曜・日曜・祝日の行は同じ色。A列が日付の行だけに効く（見出しなどの文字の行は対象外）
    a = f"$A{first_row}"
    rule = f'AND(ISNUMBER({a}),{off_day(a)})'
    ws.conditional_formatting.add(
        f"A{first_row}:{LAST_COL}{last_row}",
        FormulaRule(formula=[rule], fill=PatternFill("solid", fgColor="F8D7D3", bgColor="F8D7D3"),
                    font=Font(name=FONT, color="9C2A1F")),
    )


def control(ws, ref, value, font, fill, fmt=None):
    c = ws[ref]
    c.value, c.font, c.fill, c.border, c.alignment = value, font, fill, border, center
    if fmt:
        c.number_format = fmt


# --- 出勤簿：対象年月と出勤用ID（印刷範囲の外の AE〜AF 列）を指定して1人分 ---
ws = wb.create_sheet("出勤簿")
setup_report_sheet(ws)
control(ws, f"{CTRL_L}2", "対象年月", f_label, fill_label)
control(ws, f"{CTRL_V}2", date(2026, 10, 1), f_bold, fill_input, 'yyyy"年"m"月"')
control(ws, f"{CTRL_L}3", "出勤用ID", f_label, fill_label)
control(ws, f"{CTRL_V}3", "1234", f_bold, fill_input, "@")
ws[f"{CTRL_L}5"] = "← 黄色の欄（対象年月・出勤用ID）を指定すると、出勤簿が切り替わります。"
ws[f"{CTRL_L}6"] = "（この欄は印刷されません）"
ws[f"{CTRL_L}5"].font = ws[f"{CTRL_L}6"].font = f_note
report_block(
    ws, 1, f"${CTRL_V}$2",
    key_formula="=" + key(f"${CTRL_V}$3"),
    idx_formula=f'=IF({HID}1="","",IFERROR(MATCH({HID}1,{K}!$A$2:$A${NAME_LAST},0),""))',
)
report_colors(ws, 1 + TOP + 1, 1 + TOP + 31)
ws.conditional_formatting.add(
    "C5", FormulaRule(formula=['LEFT(C5,1)="★"'], font=Font(name=FONT, size=11, color="C00000", bold=True))
)
dv = DataValidation(type="list", formula1=f"={N}!$A$2:$A${NAME_LAST}", allow_blank=True)
dv.showErrorMessage = False
ws.add_data_validation(dv)
dv.add(f"{CTRL_V}3")
ws.print_area = f"A1:{LAST_COL}{TOP + 33}"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True

# --- 一括印刷：職員名簿の全員分の出勤簿を1人1ページで ---
ws = wb.create_sheet("一括印刷")
setup_report_sheet(ws)
control(ws, f"{CTRL_L}2", "対象年月", f_label, fill_label)
control(ws, f"{CTRL_V}2", date(2026, 10, 1), f_bold, fill_input, 'yyyy"年"m"月"')
control(ws, f"{CTRL_L}3", "印刷枚数", f_label, fill_label)
control(ws, f"{CTRL_V}3", f'=MAX(0,MIN({BATCH},MAX({K}!$F$2:$F${NAME_LAST})))', f_bold, fill_calc, '0"枚"')
ws[f"{CTRL_L}5"] = f"使い方：{CTRL_V}2 に対象年月を入力して、そのまま印刷（Ctrl+P）してください。"
ws[f"{CTRL_L}6"] = "職員名簿の全員分の出勤簿が、1人1枚（A4縦）で印刷されます。PDFとして保存することもできます。"
ws[f"{CTRL_L}7"] = f"（この欄は印刷されません。{BATCH}人まで対応）"
for i in (5, 6, 7):
    ws[f"{CTRL_L}{i}"].font = f_note
for k in range(BATCH):
    t = 1 + k * BLOCK
    report_block(
        ws, t, f"${CTRL_V}$2",
        key_formula=f'=IF(${HID}${t + 1}="","",INDEX({K}!$A$2:$A${NAME_LAST},${HID}${t + 1}))',
        idx_formula=f'=IFERROR(MATCH({k + 1},{K}!$F$2:$F${NAME_LAST},0),"")',
    )
    if k:
        ws.row_breaks.append(Break(id=t - 1))
report_colors(ws, 1, BATCH * BLOCK)
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
# 印刷範囲は名簿の人数分だけ（空のページは印刷しない）
ws.defined_names["_xlnm.Print_Area"] = DefinedName(
    "_xlnm.Print_Area",
    attr_text=f"OFFSET('一括印刷'!$A$1,0,0,MAX(1,'一括印刷'!${CTRL_V}$3)*{BLOCK},{N_COLS})",
)

# ===== 祝日（入力・追加できる） =====
HOLIDAYS = [
    ("2026-01-01", "元日"), ("2026-01-12", "成人の日"), ("2026-02-11", "建国記念の日"),
    ("2026-02-23", "天皇誕生日"), ("2026-03-20", "春分の日"), ("2026-04-29", "昭和の日"),
    ("2026-05-03", "憲法記念日"), ("2026-05-04", "みどりの日"), ("2026-05-05", "こどもの日"),
    ("2026-05-06", "振替休日"), ("2026-07-20", "海の日"), ("2026-08-11", "山の日"),
    ("2026-09-21", "敬老の日"), ("2026-09-22", "国民の休日"), ("2026-09-23", "秋分の日"),
    ("2026-10-12", "スポーツの日"), ("2026-11-03", "文化の日"), ("2026-11-23", "勤労感謝の日"),
    ("2027-01-01", "元日"), ("2027-01-11", "成人の日"), ("2027-02-11", "建国記念の日"),
    ("2027-02-23", "天皇誕生日"), ("2027-03-21", "春分の日（予定）"), ("2027-03-22", "振替休日（予定）"),
    ("2027-04-29", "昭和の日"), ("2027-05-03", "憲法記念日"), ("2027-05-04", "みどりの日"),
    ("2027-05-05", "こどもの日"), ("2027-07-19", "海の日"), ("2027-08-11", "山の日"),
    ("2027-09-20", "敬老の日"), ("2027-09-23", "秋分の日（予定）"), ("2027-10-11", "スポーツの日"),
    ("2027-11-03", "文化の日"), ("2027-11-23", "勤労感謝の日"),
]
ws = wb.create_sheet("祝日")
header(ws, 1, ["日付", "名称"], [14, 22])
style_range(ws, 2, 200, 1, 2, fill=fill_input)
style_range(ws, 2, 200, 1, 1, fmt="yyyy/m/d", align=center)
for i, (d, name) in enumerate(HOLIDAYS, start=2):
    ws[f"A{i}"] = date.fromisoformat(d)
    ws[f"B{i}"] = name
ws["D1"] = "月次一覧で赤く表示する日の一覧です。年末年始の休みなどもここに追加できます（200行目まで）。"
ws["D1"].font = f_note
ws["D2"] = "（予定）と書いた2027年の春分の日・秋分の日は、内閣府の「国民の祝日」のページで正式な日付を確認してください。"
ws["D2"].font = f_note
ws.freeze_panes = "A2"

# ===== 年次休暇（職員ごと・年ごとに1行。入力欄と、自動計算の表示） =====
ws = wb.create_sheet("年次休暇")
header(ws, 1, ["出勤用ID", "年（西暦）", "本年分（日）\n空欄は自動", "本年使用\n（日）", "本年使用\n（時間）",
               "繰越分（日）\n活用開始時のみ", "採用日\n（年の途中のみ）", "退職日\n（年の途中のみ）", "備考",
               "本年分\n（自動）", "繰越分\n（自動）", "合計\n（自動）", "翌年繰越\n（自動）"],
       [12, 10, 11, 9, 9, 13, 12, 12, 20, 9, 15, 15, 15])
ws.row_dimensions[1].height = 32
for c in ws[1]:
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
style_range(ws, 2, LEAVE_LAST, 1, 9, fill=fill_input, align=center)
style_range(ws, 2, LEAVE_LAST, 1, 1, fmt="@", align=center)
style_range(ws, 2, LEAVE_LAST, 5, 5, fill=fill_input, fmt="[h]:mm", align=center)   # 時間は「時間:分」（例 2:15）
style_range(ws, 2, LEAVE_LAST, 7, 8, fill=fill_input, fmt="yyyy/m/d", align=center)
style_range(ws, 2, LEAVE_LAST, 10, 13, fill=fill_calc, align=center)
for r in range(2, LEAVE_LAST + 1):
    ws[f"I{r}"].alignment = left
    ws[f"J{r}"] = f'=IF(INDEX({K}!$H:$H,ROW())="","",INDEX({K}!$P:$P,ROW())&"日")'
    for col, src in (("K", "J"), ("L", "K"), ("M", "L")):
        ws[f"{col}{r}"] = f'=IF(INDEX({K}!$H:$H,ROW())="","",{days_text(f"INDEX({K}!${src}:${src},ROW())")})'
ws["A2"], ws["B2"], ws["C2"], ws["D2"], ws["E2"] = "1234", 2026, 20, 5, time(2, 15)
ws["F2"], ws["I2"] = 10, "（例）この行は例です。削除してください"
ws.freeze_panes = "A2"
ws["O1"] = "使い方"
ws["O1"].font = f_bold
for i, text in enumerate([
    "・職員ごと・年ごとに1行入力します（例：出勤用ID 1234 の2026年分）。",
    f"・本年分（その年に付与された日数）は、通年で勤務する場合は空欄のままで構いません（{GRANT_DEFAULT}日として計算）。",
    "・年の途中で採用・退職した場合は、採用日・退職日を入力すると、在職期間に応じた本年分が自動で入ります（1月に達するまで2日〜11月を超え1年未満20日）。",
    "・本年分を入力した場合は、入力した日数が優先されます。",
    "・本年使用（その年に使った日数と時間）を入力します。",
    "・年次休暇は15分単位で取得できます。時間は「時間:分」で入力します（例：2時間15分 → 2:15）。",
    f"・1日は7時間45分として計算します。7時間45分以上の時間は自動で日数に繰り上がります。",
    "・繰越分は、前の年の行があれば、その行の「翌年繰越」が自動で入ります。",
    f"・活用開始時（前の年の行がない年）だけ、F列に繰越分の日数を手入力してください。繰越分は日単位なので、1日未満の端数は切り捨てます（上限{CARRY_MAX}日）。",
    f"・合計は繰越分＋本年分です（上限{TOTAL_MAX}日）。",
    f"・翌年繰越は、残り（合計−本年使用）の1日未満の端数を切り捨てた日数です（上限{CARRY_MAX}日）。繰越分と本年分は区別しません。",
    "・ここで入力した内容は、出勤簿の「年次休暇 取得日数」の欄に表示されます。",
], start=2):
    ws[f"O{i}"] = text
    ws[f"O{i}"].font = f_note

# ===== 月別取得時間（職員ごと・年ごと・区分ごとに1行。1月〜12月の取得時間を入力） =====
ws = wb.create_sheet("月別取得時間")
header(ws, 1, ["出勤用ID", "年（西暦）", "区分"] + [f"{m}月" for m in range(1, 13)] + ["備考"],
       [12, 10, 12] + [6] * 12 + [24])
style_range(ws, 2, MH_LAST, 1, 16, fill=fill_input, align=center)
style_range(ws, 2, MH_LAST, 1, 1, fmt="@", align=center)
for r in range(2, MH_LAST + 1):
    ws[f"P{r}"].alignment = left
dv_cat = DataValidation(type="list", formula1='"' + ",".join(HOUR_CATS) + '"', allow_blank=True)
dv_cat.error, dv_cat.errorTitle = "病気休暇・看護休暇から選んでください。", "区分"
ws.add_data_validation(dv_cat)
dv_cat.add(f"C2:C{MH_LAST}")
ws["A2"], ws["B2"], ws["C2"], ws["D2"], ws["P2"] = "1234", 2026, "病気休暇", time(3, 30), "（例）この行は例です。削除してください"
ws["A3"], ws["B3"], ws["C3"], ws["E3"], ws["P3"] = "1234", 2026, "看護休暇", time(0, 45), "（例）この行は例です。削除してください"
for ref in ("D2", "E3"):
    ws[ref].number_format = "[h]:mm"
ws.freeze_panes = "D2"
ws["R1"] = "使い方"
ws["R1"].font = f_bold
for i, text in enumerate([
    "・職員ごと・年ごと・区分（病気休暇・看護休暇）ごとに1行入力します。",
    "・1月〜12月の欄に、その月の取得時間を「時間:分」で入力します（例：3時間30分 → 3:30、45分 → 0:45）。",
    "・ここで入力した内容は、出勤簿の上の欄（取得時間）に表示されます。",
    "・区分は一覧から選んでください。同じID・年・区分の行が2つあると、上の行が使われます。",
], start=2):
    ws[f"R{i}"] = text
    ws[f"R{i}"].font = f_note

wb.calculation.calcMode = "auto"
wb.calculation.fullCalcOnLoad = True
wb.active = 0
wb.save(OUT)
print("saved", OUT)
