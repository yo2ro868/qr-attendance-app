"""テレワーク予定表（所属に1ファイル）のひな形を作る。

使い方: python make_form.py 出力先.xlsx
管理者がこのひな形を所属ごとにコピーし（ファイル名は テレワーク予定表　所属名.xlsx）、
管理者の OneDrive から所属長（編集）にだけ共有する。所属長は、配属の職員のテレワークの予定を1日1行で書く（承認はない）。
出勤・休暇管理データ が Power Query でフォルダごと読み、出勤簿の出勤欄に「テレ」を出す。
「予定」の表の名前と見出しは、出勤・休暇管理データ の Power Query（tools/telework/queries）が参照するので変えないこと。
"""
import sys
from datetime import date
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "matching_excel"))
from holidays import HOLIDAYS

OUT = sys.argv[1]
FONT = "ＭＳ Ｐゴシック"
ROWS = 500            # 予定を書ける行数（保護したシートでは表に行を足せないので、先に用意しておく）
HEAD = 5              # 予定の表の見出しの行
FIRST, LAST = HEAD + 1, HEAD + ROWS
KINDS = ["終日", "時間"]      # 区分：1日ずっとテレワーク／時刻を指定（時間単位）
BREAK_START, BREAK_END = 12 / 24, 13 / 24     # 時間の計算で除く休憩（12:00〜13:00）
DUPLICATE = "・同じ職員・同じ日の行がほかにもあります"

f_base = Font(name=FONT, size=10)
f_bold = Font(name=FONT, size=10, bold=True)
f_head = Font(name=FONT, size=10, bold=True, color="FFFFFF")
f_title = Font(name=FONT, size=14, bold=True, color="1E3A5F")
f_note = Font(name=FONT, size=9, color="555555")
fill_head = PatternFill("solid", fgColor="1E3A5F")
fill_boss = PatternFill("solid", fgColor="FFF2CC")     # 所属長が入力
fill_calc = PatternFill("solid", fgColor="F2F2F2")     # 自動
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left = Alignment(horizontal="left", vertical="center", wrap_text=True)

wb = Workbook()
ws = wb.active
ws.title = "予定"

# 予定の表の列：(見出し, 幅, 書式, 塗り)
COLS = [
    ("出勤用ID", 10, "@", fill_boss),
    ("氏名（確認用）", 14, None, fill_boss),   # 所属長の確認用。システムは使わない（出勤用ID で照合する）
    ("日付", 11, "yyyy/m/d", fill_boss),
    ("区分", 7, None, fill_boss),
    ("開始時刻", 8, "h:mm", fill_boss),
    ("終了時刻", 8, "h:mm", fill_boss),
    ("時間", 7, "[h]:mm", fill_calc),         # 開始時刻・終了時刻から休憩を除いて自動で計算（打ち込めない）
    ("備考", 24, None, fill_boss),
    ("確認", 34, None, fill_calc),
]
for i, (_, width, _, _) in enumerate(COLS, start=1):
    ws.column_dimensions[get_column_letter(i)].width = width

# ----- 管理（非表示）：土日以外で除く日（祝日・年末年始）と時刻の一覧 -----
MG = "'管理'"
mg = wb.create_sheet("管理")
for i, (d, _) in enumerate(HOLIDAYS, start=1):
    mg[f"A{i}"] = date.fromisoformat(d)
year_end = [date(y, 12, d) for y in range(2025, 2041) for d in (29, 30, 31)] + [date(y, 1, d) for y in range(2026, 2042) for d in (1, 2, 3)]
for i, d in enumerate(year_end, start=1):
    mg[f"B{i}"] = d
OFF_DAYS = f"{MG}!$A$1:$B$200"
TIMES = [(8 * 60 + 15 * i) / 1440 for i in range(45)]     # 8:00〜19:00 の15分おき（一覧にない時刻も書ける）
for i, v in enumerate(TIMES, start=1):
    mg[f"C{i}"] = v
    mg[f"C{i}"].number_format = "h:mm"
TIME_RANGE = f"{MG}!$C$1:$C${len(TIMES)}"
mg.sheet_state = "veryHidden"

# ----- 1〜4行目：題名と説明 -----
ws["A1"] = "テレワーク予定表"
ws["A1"].font = f_title
ws["C1"] = "黄色：所属長が入力　　灰色：自動　　赤：入力が足りない・誤りのある欄（理由は「確認」の欄）"
ws["C1"].font = f_note
for r, text in enumerate([
    "配属の職員のテレワークの予定を、1人1日1行で書きます（順番は自由。予定が変わったら行を直すか消してください）。",
    "区分：1日ずっとテレワークなら「終日」、一部の時間だけなら「時間」を選び、開始時刻・終了時刻を書きます（時間は12:00〜13:00の休憩を除いて自動で入ります）。",
    "土日・祝日・年末年始は書けません。書いた内容は、管理者の 出勤・休暇管理データ の出勤簿に「テレ」と表示されます。",
], start=2):
    ws[f"A{r}"] = text
    ws[f"A{r}"].font = f_note

# ----- 5行目から：予定の表 -----
for i, (label, _, fmt, fill) in enumerate(COLS, start=1):
    h = ws.cell(row=HEAD, column=i, value=label)
    h.font, h.alignment, h.border, h.fill = f_head, center, border, fill_head
    for r in range(FIRST, LAST + 1):
        c = ws.cell(row=r, column=i)
        c.font, c.border, c.fill = f_base, border, fill
        c.alignment = left if label in ("氏名（確認用）", "備考", "確認") else center
        if fmt:
            c.number_format = fmt
ws.row_dimensions[HEAD].height = 22
for r in range(FIRST, LAST + 1):
    a, b, c_, d, e, f, g, h = (f"{x}{r}" for x in "ABCDEFGH")
    # 時間：区分が「時間」で開始時刻・終了時刻があれば、その差から休憩（重なる分だけ）を引く
    ws[g] = (f'=IF(AND({d}="時間",ISNUMBER({e}),ISNUMBER({f})),IF({f}>{e},'
             f'ROUND(({f}-{e}-MAX(0,MIN({f},{BREAK_END})-MAX({e},{BREAK_START})))*1440,0)/1440,""),"")')
    ws[f"I{r}"] = (
        f'=IF(COUNTA({a}:{f},{h})=0,"",'
        f'IF({a}="","・出勤用ID を入れてください","")'
        f'&IF({c_}="","・日付を入れてください",IF(NOT(ISNUMBER({c_})),"・日付の形で入れてください",'
        f'IF(NETWORKDAYS({c_},{c_},{OFF_DAYS})=0,"・土日・祝日・年末年始です","")))'
        f'&IF({d}="","・区分を選んでください","")'
        f'&IF(AND({a}<>"",ISNUMBER({c_})),IF(COUNTIFS($A${FIRST}:$A${LAST},{a},$C${FIRST}:$C${LAST},{c_})>1,"{DUPLICATE}",""),"")'
        f'&IF(AND({d}="時間",OR({e}="",{f}="")),"・開始時刻と終了時刻を入れてください","")'
        f'&IF(AND({d}="終日",OR({e}<>"",{f}<>"")),"・終日に時刻は要りません","")'
        f'&IF(AND({e}<>"",{f}<>""),IF({f}<={e},"・終了時刻が開始時刻より前です",""),""))'
    )

# 入力が足りない・誤りのある欄を赤くする（Excel Online でも動くよう条件付き書式で）。条件は I 列の式とそろえる
R = FIRST
fill_bad = PatternFill("solid", fgColor="F4B6B0", bgColor="F4B6B0")
font_bad = Font(name=FONT, color="9C0006", bold=True)
used = f"COUNTA($A{R}:$F{R},$H{R})>0"
RED_RULES = {
    "A": f'AND({used},$A{R}="")',
    "C": (f'OR(AND({used},$C{R}=""),AND($C{R}<>"",NOT(ISNUMBER($C{R}))),'
          f'IF(ISNUMBER($C{R}),NETWORKDAYS($C{R},$C{R},{OFF_DAYS})=0,FALSE),'
          f'AND($A{R}<>"",ISNUMBER($C{R}),COUNTIFS($A${FIRST}:$A${LAST},$A{R},$C${FIRST}:$C${LAST},$C{R})>1))'),
    "D": f'AND({used},$D{R}="")',
    "E": f'OR(AND($D{R}="時間",$E{R}=""),AND($D{R}="終日",$E{R}<>""))',
    "F": f'OR(AND($D{R}="時間",$F{R}=""),AND($D{R}="終日",$F{R}<>""),AND($E{R}<>"",$F{R}<>"",$F{R}<=$E{R}))',
    "I": f'$I{R}<>""',
}
for col, rule in RED_RULES.items():
    ws.conditional_formatting.add(f"{col}{FIRST}:{col}{LAST}", FormulaRule(formula=[rule], fill=fill_bad, font=font_bad))
# 終日の行の時刻の欄は要らないので灰色に
ws.conditional_formatting.add(f"E{FIRST}:F{LAST}", FormulaRule(
    formula=[f'AND($D{R}="終日",$E{R}="",$F{R}="")'], fill=PatternFill("solid", fgColor="D9D9D9", bgColor="D9D9D9")))

t = Table(displayName="予定", ref=f"A{HEAD}:{get_column_letter(len(COLS))}{LAST}")
t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
ws.add_table(t)
ws.freeze_panes = f"A{FIRST}"


def col_cells(col):
    return f"{col}{FIRST}:{col}{LAST}"


dv = DataValidation(type="list", formula1='"' + ",".join(KINDS) + '"', allow_blank=True)
dv.errorTitle, dv.error = "区分", "「終日」か「時間」を選んでください。"
ws.add_data_validation(dv)
dv.add(col_cells("D"))
dv = DataValidation(type="list", allow_blank=True, imeMode="halfAlpha", formula1=TIME_RANGE)
dv.showErrorMessage = False      # 一覧にない時刻も書ける
ws.add_data_validation(dv)
dv.add(col_cells("E"))
dv.add(col_cells("F"))
# 日付は勤務日だけ（土日・祝日・年末年始は入力できない。貼り付けですり抜けたときは赤くなる）
dv = DataValidation(type="custom", allow_blank=True, imeMode="halfAlpha",
                    formula1=f"AND(ISNUMBER(C{FIRST}),NETWORKDAYS(C{FIRST},C{FIRST},{OFF_DAYS})=1)")
dv.errorTitle, dv.error = "日付", "土日・祝日・年末年始は書けません。日付は 11/5 または 2026/11/5 の形で入れてください。"
ws.add_data_validation(dv)
dv.add(col_cells("C"))
dv = DataValidation(type="custom", allow_blank=True, imeMode="halfAlpha", formula1=f"LEN(A{FIRST})>0")
ws.add_data_validation(dv)
dv.add(col_cells("A"))

# 所属長が入力する欄だけ保護を外す（時間・確認は自動なので保護のまま）。パスワードはかけない（誤操作を防ぐため）
for r in range(FIRST, LAST + 1):
    for c in range(1, 9):
        if c != 7:
            ws.cell(row=r, column=c).protection = Protection(locked=False)
ws.protection.sheet = True
ws.protection.formatColumns = False
ws.protection.formatRows = False
ws.protection.autoFilter = False
ws.protection.sort = False

ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_title_rows = f"{HEAD}:{HEAD}"

# ===== 使い方 =====
us = wb.create_sheet("使い方")
us.column_dimensions["A"].width = 110
us["A1"] = "テレワーク予定表の使い方"
us["A1"].font = f_title
for i, text in enumerate([
    "■ 所属長",
    "・職員からテレワークの利用の申し出（口頭やメール）を受けたら、この表に予定を書きます。承認の欄はありません。",
    "・1人1日1行です。出勤用ID・日付・区分を書きます。氏名は確認用で、書かなくてもかまいません（システムは出勤用ID で照合します）。",
    "・区分が「時間」のときは、開始時刻・終了時刻を書きます。時間は12:00〜13:00の休憩を除いて自動で入ります。",
    "・予定が変わったら、その行を書き直すか、行の中身を消してください（行ごと削除してもかまいません）。",
    "・足りない欄や誤りのある欄は赤くなり、理由が「確認」の欄に出ます。",
    "",
    "■ 管理者",
    "・このファイルは管理者の OneDrive にあり、所属長にだけ共有しています。ファイル名は「テレワーク予定表　所属名.xlsx」（全角スペース区切り）にします。",
    "・出勤・休暇管理データ が、保存された内容を開いたとき（または［申請を取り込む］）に読み込み、出勤簿の出勤欄に「テレ」、時間のテレワークは備考に時刻を表示します。",
    "・全員の予定と注意（名簿にない出勤用ID、休暇と重なる日など）は、出勤・休暇管理データ の「テレワーク一覧」シートで確かめます。",
], start=3):
    us[f"A{i}"] = text
    us[f"A{i}"].font = f_bold if text.startswith("■") else f_base

wb.active = 0
wb.save(OUT)
print("saved", OUT)
