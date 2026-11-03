"""休暇申請用紙（職員1人に1ファイル）のひな形を作る。

使い方: python make_form.py 出力先.xlsx
管理者がこのひな形を職員ごとにコピーし（ファイル名は 休暇申請_出勤用ID_氏名.xlsx。マクロ 申請用紙の作成.xlsm が作る）、
管理者の OneDrive から本人（編集）と所属長（編集）にだけ共有して使う。
出勤用ID・氏名はファイル名から読む。年・持ち日数・本年使用の調整は、照合用Excel の休暇等シートで一括管理し、
マクロが非表示の「管理」シートに書き込む（本人・所属長には入力欄を見せない）。
「申請」の表の名前と見出しは、照合用Excel・休暇集計の Power Query が参照するので変えないこと。
"""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

OUT = sys.argv[1]
FONT = "ＭＳ Ｐゴシック"
DAY_MIN = 465         # 1日＝7時間45分＝465分
ROWS = 300            # 申請を書ける行数（保護したシートでは表に行を足せないので、先に用意しておく）
HEAD = 11             # 申請の表の見出しの行
FIRST, LAST = HEAD + 1, HEAD + ROWS

KINDS = ["年次休暇", "病気休暇", "看護休暇（子の看護等）", "特別休暇（看護休暇以外）",
         "介護休暇", "介護時間", "育児時間", "職務専念義務の免除"]
SPECIAL = ["夏季休暇", "結婚", "忌引（親族の死亡）", "父母の追悼", "妻の出産", "育児参加", "短期介護",
           "不妊治療", "産前", "産後", "保育時間", "ボランティア", "骨髄等の提供", "選挙権の行使等",
           "官公署への出頭", "現住居の滅失等"]
UNITS = ["1日単位", "時間単位"]
APPROVALS = ["承認", "却下", "取消"]

f_base = Font(name=FONT, size=10)
f_bold = Font(name=FONT, size=10, bold=True)
f_head = Font(name=FONT, size=10, bold=True, color="FFFFFF")
f_title = Font(name=FONT, size=14, bold=True, color="1E3A5F")
f_big = Font(name=FONT, size=12, bold=True, color="1E3A5F")
f_note = Font(name=FONT, size=9, color="555555")
fill_head = PatternFill("solid", fgColor="1E3A5F")     # 見出し（白い文字）
fill_staff = PatternFill("solid", fgColor="FFF2CC")    # 本人が入力
fill_boss = PatternFill("solid", fgColor="DDEBF7")     # 所属長が入力
fill_calc = PatternFill("solid", fgColor="F2F2F2")     # 自動
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left = Alignment(horizontal="left", vertical="center", wrap_text=True)


def days_text(m):
    """分（m はセル参照）を「15日3時間30分」の形の文字にする。1日＝7時間45分。負の数は先頭に −。"""
    a = f"ABS({m})"
    rest = f"MOD({a},{DAY_MIN})"
    return (f'IF({m}<0,"−","")&INT({a}/{DAY_MIN})&"日"'
            f'&IF(INT({rest}/60)>0,INT({rest}/60)&"時間","")'
            f'&IF(MOD({rest},60)>0,MOD({rest},60)&"分","")')


def hm_text(m):
    return f'INT({m}/60)&"時間"&IF(MOD({m},60)>0,MOD({m},60)&"分","")'


wb = Workbook()
ws = wb.active
ws.title = "申請"

# 申請の表の列：(見出し, 幅, 書式, 塗り, 入力する人)
COLS = [
    ("申請日", 11, "yyyy/m/d", fill_staff),
    ("休暇の種類", 20, None, fill_staff),
    ("種別・内容", 16, None, fill_staff),
    ("単位", 9, None, fill_staff),
    ("開始日", 11, "yyyy/m/d", fill_staff),
    ("終了日", 11, "yyyy/m/d", fill_staff),
    ("日数", 6, "0.##", fill_staff),
    ("時間", 7, "[h]:mm", fill_staff),
    ("時刻", 13, None, fill_staff),
    ("事由", 22, None, fill_staff),
    ("承認（所属長）", 9, None, fill_boss),
    ("承認日", 11, "yyyy/m/d", fill_boss),
    ("所属長コメント", 18, None, fill_boss),
    ("確認", 30, None, fill_calc),
    ("分", 6, "0", fill_calc),     # 取得時間（分）。上の残日数の計算に使う（非表示）
    ("年", 6, "0", fill_calc),     # 開始日の年（非表示）
]
for i, (_, width, _, _) in enumerate(COLS, start=1):
    ws.column_dimensions[get_column_letter(i)].width = width

# ----- 1〜2行目：題名 -----
ws["A1"] = "休暇申請用紙"
ws["A1"].font = f_title
ws["D1"] = "黄色：本人が入力　　青：所属長が入力　　灰色：自動"
ws["D1"].font = f_note

# ----- 管理（非表示）：マクロ（申請用紙の作成.xlsm）が照合用Excel の休暇等シートから書き込む -----
MG = "'管理'"
YR, HOLD, ADJ = f"{MG}!$B$3", f"{MG}!$B$4", f"{MG}!$B$5"   # 年、持ち日数（分）、本年使用の調整（分）
mg = wb.create_sheet("管理")
for i, label in enumerate(["出勤用ID", "氏名", "年", "持ち日数（分）", "本年使用の調整と Forms の承認済み（分）", "更新日時"], start=1):
    mg[f"A{i}"] = label
mg["B1"].number_format = "@"
mg["C1"] = "このシートは マクロ（申請用紙の作成.xlsm）が書き込みます。手で変えないでください。"
mg.sheet_state = "veryHidden"     # 本人・所属長からは見えない（Excel の［再表示］にも出ない）

# ----- 3行目：氏名（管理シートから表示） -----
ws["A3"] = "氏名"
ws["A3"].font, ws["A3"].alignment, ws["A3"].border = f_bold, center, border
ws.merge_cells("B3:C3")
ws["B3"] = f"={MG}!$B$2&\"\""
for col in "BC":
    ws[f"{col}3"].font, ws[f"{col}3"].border, ws[f"{col}3"].fill = f_bold, border, fill_calc
ws["B3"].alignment = center
ws.row_dimensions[3].height = 22

# ----- 6〜9行目：年次休暇の残り（自動）。9行目は分の計算用で非表示 -----
rng = lambda c: f"${c}${FIRST}:${c}${LAST}"
KIND, MIN, YEAR, OK = rng("B"), rng("O"), rng("P"), rng("K")
ws["A6"] = f'=IF({YR}="","年次休暇の残り（管理者が設定すると表示されます）",{YR}&"年の年次休暇")'
ws["A6"].font = f_big
SUMMARY = [
    ("持ち日数", f"=N({HOLD})", days_text("A9")),
    # 使用済み＝ 照合用Excel の調整（導入前の使用など）と Forms で承認された分 ＋ この用紙の承認済みの年次休暇
    ("使用済み（承認済み）",
     f'=N({ADJ})+SUMIFS({MIN},{KIND},"年次休暇",{YEAR},{YR},{OK},"承認")', days_text("B9")),
    ("承認待ち", f'=SUMIFS({MIN},{KIND},"年次休暇",{YEAR},{YR},{OK},"")', days_text("C9")),
    ("残り", "=A9-B9", days_text("D9")),
    ("病気休暇（承認済み）", f'=SUMIFS({MIN},{KIND},"病気休暇",{YEAR},{YR},{OK},"承認")', hm_text("E9")),
    ("看護休暇（承認済み）", f'=SUMIFS({MIN},{KIND},"看護休暇（子の看護等）",{YEAR},{YR},{OK},"承認")', hm_text("F9")),
]
for i, (label, minutes, text) in enumerate(SUMMARY, start=1):
    col = get_column_letter(i)
    h = ws[f"{col}7"]
    h.value, h.font, h.alignment, h.border = label, f_bold, center, border
    v = ws[f"{col}8"]
    v.value, v.font, v.alignment, v.border, v.fill = "=" + text, f_bold, center, border, fill_calc
    ws[f"{col}9"] = minutes
ws.row_dimensions[7].height = 28
ws.row_dimensions[9].hidden = True
ws["H7"] = "1日＝7時間45分。承認待ちは、所属長の承認欄が空欄の申請です。"
ws["H7"].font = f_note

# ----- 11行目から：申請の表 -----
ws["A10"] = "申請（1件1行。上から順に書きます）"
ws["A10"].font = f_big
for i, (label, _, fmt, fill) in enumerate(COLS, start=1):
    h = ws.cell(row=HEAD, column=i, value=label)
    h.font, h.alignment, h.border, h.fill = f_head, center, border, fill_head
    for r in range(FIRST, LAST + 1):
        c = ws.cell(row=r, column=i)
        c.font, c.border, c.fill = f_base, border, fill
        c.alignment = left if label in ("種別・内容", "時刻", "事由", "所属長コメント", "確認") else center
        if fmt:
            c.number_format = fmt
ws.row_dimensions[HEAD].height = 28
for r in range(FIRST, LAST + 1):
    b, d, e, f, g, h = (f"{c}{r}" for c in "BDEFGH")
    hmin = f"ROUND(N({h})*1440,0)"
    ws[f"O{r}"] = f'=IF(OR({b}="",{e}=""),"",ROUND(N({g})*{DAY_MIN},0)+{hmin})'
    ws[f"P{r}"] = f'=IF({e}="","",YEAR({e}))'
    ws[f"N{r}"] = (
        f'=IF({b}="","",'
        f'IF({e}="","・開始日を入れてください","")'
        f'&IF(AND({d}<>"時間単位",N({g})<=0),"・日数を入れてください","")'
        f'&IF(AND({d}="時間単位",N({h})<=0),"・時間を入れてください","")'
        f'&IF(AND({f}<>"",{e}<>""),IF({f}<{e},"・終了日が開始日より前です",""),"")'
        f'&IF(AND({b}="年次休暇",N({h})>0),IF(MOD({hmin},15)<>0,"・年次休暇の時間は15分単位です",""),"")'
        f'&IF(AND(OR({b}="介護時間",{b}="育児時間"),N({h})>0),IF(MOD({hmin},30)<>0,"・30分単位です",""),"")'
        f'&IF(AND({f}<>"",{e}<>"",OR({b}="年次休暇",{b}="病気休暇",{b}="看護休暇（子の看護等）")),'
        f'IF(TEXT({e},"yyyymm")<>TEXT({f},"yyyymm"),"・月ごとに分けて書いてください",""),""))'
    )
t = Table(displayName="申請", ref=f"A{HEAD}:{get_column_letter(len(COLS))}{LAST}")
t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
ws.add_table(t)
ws.column_dimensions["O"].hidden = True
ws.column_dimensions["P"].hidden = True
ws.freeze_panes = f"A{FIRST}"


def add_list(cells, items, strict=True, prompt=None):
    dv = DataValidation(type="list", formula1='"' + ",".join(items) + '"', allow_blank=True)
    dv.showErrorMessage = strict
    if strict:
        dv.errorTitle, dv.error = "入力できない値です", "一覧から選んでください。"
    if prompt:
        dv.promptTitle, dv.prompt, dv.showInputMessage = prompt[0], prompt[1], True
    ws.add_data_validation(dv)
    dv.add(cells)


def col_cells(col):
    return f"{col}{FIRST}:{col}{LAST}"


add_list(col_cells("B"), KINDS)
add_list(col_cells("C"), SPECIAL, strict=False,
         prompt=("種別・内容", "特別休暇は種別を一覧から選ぶか書きます。職務専念義務の免除は内容（例：健康診断）を書きます。"))
add_list(col_cells("D"), UNITS)
add_list(col_cells("K"), APPROVALS, prompt=("所属長が入力", "承認・却下を選びます。取り消すときは「取消」。"))
for col in "AEFL":
    dv = DataValidation(type="date", operator="greaterThan", formula1="43831", allow_blank=True)   # 2020/1/1 より後
    dv.errorTitle, dv.error = "日付を入れてください", "例：11/5 または 2026/11/5"
    ws.add_data_validation(dv)
    dv.add(col_cells(col))
dv = DataValidation(type="decimal", operator="between", formula1="0", formula2="31", allow_blank=True)
dv.errorTitle, dv.error = "日数", "0〜31 の数を入れてください。"
ws.add_data_validation(dv)
dv.add(col_cells("G"))
dv = DataValidation(type="time", operator="between", formula1="0", formula2=str(DAY_MIN / 1440), allow_blank=True)
dv.errorTitle, dv.error = "時間", "「時間:分」で入れてください（例：2:15）。7:45 までです。"
dv.promptTitle, dv.prompt, dv.showInputMessage = "時間", "時間単位のときに「時間:分」で入れます（例：2:15）。", True
ws.add_data_validation(dv)
dv.add(col_cells("H"))

# 本人・所属長が入力する欄だけ保護を外す（基本・計算式は保護）。パスワードはかけない（誤操作を防ぐため）
for r in range(FIRST, LAST + 1):
    for c in range(1, 14):   # A〜M
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
us["A1"] = "休暇申請用紙の使い方"
us["A1"].font = f_title
for i, text in enumerate([
    "■ 本人（申請する人）",
    "・「申請」シートの黄色の欄に、1件1行で書きます。休暇の種類・単位などは一覧から選びます。",
    "・1日単位のときは 開始日・終了日（1日だけなら空欄）・日数（土日・祝日を除く）を、時間単位のときは 開始日・時間（例 2:15）・時刻（例 13:00〜15:15）を書きます。",
    "・年次休暇の時間は15分単位、介護時間・育児時間は30分単位です。月をまたぐときは月ごとに分けて書きます。",
    "・書いたら、所属長に申請したことを伝えます（Notes や口頭で）。",
    "・「確認」の欄に注意が出たら、書き直してください。",
    "・上の「年次休暇の残り」で、残日数と承認待ちが分かります（1日＝7時間45分）。",
    "・承認された行は書き換えないでください。取り消したいときは所属長に伝えます。",
    "",
    "■ 所属長",
    "・青の欄に、承認・却下（取り消すときは取消）と承認日を入れます。必要ならコメントを書きます。",
    "・承認欄が空欄の行は「承認待ち」として数えます。",
    "",
    "■ 管理者",
    "・年・持ち日数・本年使用の調整（導入前の使用など）は、照合用Excel の「休暇等」シートで全職員分を管理し、申請用紙の作成.xlsm の［申請用紙を更新］でこの用紙に書き込みます（非表示の「管理」シート）。",
    "・このファイルは管理者の OneDrive にあり、本人と所属長にだけ共有しています。ほかの人には見えません。",
], start=3):
    us[f"A{i}"] = text
    us[f"A{i}"].font = f_bold if text.startswith("■") else f_base

wb.active = 0
wb.save(OUT)
print("saved", OUT)
