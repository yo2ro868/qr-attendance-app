"""出勤・休暇管理データ.xlsx を作る（様式と計算式）。

使い方: python make_matching.py 出力先.xlsx
通常は同じフォルダの build.ps1 から実行する（Excel での再計算までまとめて行う）。
"""
import sys
from datetime import date, time
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.pagebreak import Break
from openpyxl.workbook.defined_name import DefinedName
from holidays import HOLIDAYS

OUT = sys.argv[1]
FONT = "ＭＳ Ｐゴシック"
STAFF_ROWS = 500      # 職員名簿の行数（2〜501行目）
DATA_ROWS = 5000      # 出力データの行数（2〜5001行目）
NAME_LAST = STAFF_ROWS + 1
DATA_LAST = DATA_ROWS + 1
BATCH = 80           # 一括印刷できる人数（前もって式を入れておく人数。多いとファイルが重くなる）

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
LV = "'休暇等'"
LEAVE_LAST = 1001     # 休暇等シートの行数（2〜1001行目）
CARRY_MAX = 20        # 翌年に繰り越せる日数の上限
GRANT_DEFAULT = 20    # 本年分が空欄のときの日数（通年で勤務する場合の付与日数）
# 年の中途採用・退職者の本年分（在職期間が 1月に達するまで, 1月を超え2月に達するまで, … 11月を超え1年未満）
PRORATED = [2, 3, 5, 7, 8, 10, 12, 13, 15, 17, 18, 20]
TOTAL_MAX = 40        # 1年に使える日数（繰越分＋本年分）の上限
DAY_MIN = 465         # 年次休暇の1日＝7時間45分＝465分（15分単位で取得できる）
HOL = "'祝日'!$A$2:$A$200"
HOUR_CATS = ["病気休暇", "看護休暇"]
# 申請・報告シートの読み込みシート：申請・報告シート（tools/leave_sheet）を Power Query で読んだ表を置く（add_queries.ps1 が作る）。
# 2行目が設定（申請・報告シートのフォルダ）、5行目が前回の取り込みの日時。表の見出しは8行目、データは9行目から。
# 6行目に各表の行数（照合を速くするため、その行数の範囲だけを探す。式は add_queries.ps1 が表を作ってから入れる）。
# 表は読み直すと行が増減して下のセルがずれるので、この表を指す式は列全体（$A:$A）と行番号で書く
AU = "'申請・報告シートの読み込み'"
AUTO_COUNT_ROW, AUTO_HEAD = 6, 8
AUTO_FIRST = AUTO_HEAD + 1
AU_Y, AU_M, AU_D = "A", "E", "S"     # 年休_自動（キー・分）、月別_自動（キー・1〜12月）、日別_自動（キー・8欄）のキーの列
AU_T = "AD"                           # テレワーク_日別（キー・表示・備考）のキーの列
AU_F = "AH"                           # フレックス_出勤簿（キー・表示）のキーの列
AU_U = "AK"                           # 利用予定_日別（キー・表示）のキーの列
AUTO_TABLES = {"年休_自動": AU_Y, "月別_自動": AU_M, "日別_自動": AU_D, "テレワーク_日別": AU_T, "フレックス_出勤簿": AU_F, "利用予定_日別": AU_U}
TW = "'テレワーク一覧'"                # テレワーク一覧シート：B2 がテレワーク予定表のフォルダ、表は TW_HEAD 行目から（add_queries.ps1 が作る）
TW_HEAD = 5


def col_after(col, n):
    """col から n 列右の列の文字"""
    return get_column_letter(column_index_from_string(col) + n)


def auto_col(col):
    return f"{AU}!${col}:${col}"


def auto_match(key_expr, key_col):
    """申請・報告シートの読み込みシートのキーの列から key_expr を探した行番号（なければ ""）。"""
    c = auto_col(key_col)
    return (f'IFERROR(MATCH({key_expr},INDEX({c},{AUTO_FIRST}):INDEX({c},{AUTO_FIRST - 1}+MAX(1,{AU}!${key_col}${AUTO_COUNT_ROW})),0)'
            f'+{AUTO_FIRST - 1},"")')


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
ws["B1"] = "出勤記録 出勤・休暇管理データ（管理者のみ保管）"
ws["B1"].font = f_title
lines = [
    ("まずは先頭の「目次」シートを見てください。作業の流れごとに、行くシートとボタンと今の状況（承認待ちの件数など）をまとめています。", True),
    ("", False),
    ("手順", True),
    ("1. 「職員名簿」シートに、出勤用IDと氏名を入力します（初回のみ。異動時に更新）。", False),
    ("2. 出勤管理システムの管理者画面から「Excel出力」を押し、attendance.xlsx をダウンロードします。", False),
    ("3. attendance.xlsx を、「出力データ貼付」シートの H3 に入れたフォルダに置き、目次の［最新を読み込む］を押します（フォルダの中のファイルを全部読み、A2 から下に入れ直します）。", False),
    ("    ※ フォルダを使わないときは、attendance.xlsx のデータ（見出しを除く）をコピーし、「出力データ貼付」シートの A2 に貼り付けます（H3 は空欄にします）。", False),
    ("    ※ セル範囲のコピーでも、行ごと（行番号を選んで）のコピーでも大丈夫です。", False),
    ("    ※ 前回のデータは、2行目から下を選んで Delete キーで消すか、行ごと削除してから貼り付けてください。", False),
    ("4. 「照合結果」シートに氏名つきの出勤記録が表示されます。見出しのフィルターで氏名・日付を絞り込めます。", False),
    ("5. 「月次一覧」シートの B1 に対象年月（例：2026/10）を入力すると、職員ごと・日ごとの出勤（○）が表示されます。", False),
    ("6. 「出勤簿」シートの右側（AF2・AF3）で対象年月と出勤用ID（一覧から選択）を指定すると、その職員の1か月分の出勤簿（A4縦）が表示され、そのまま印刷できます。", False),
    ("    ※ 上の欄の題名は「令和〇年出勤簿」と自動で表示されます。所属の欄は記入用の空欄です。", False),
    ("    ※ 病気休暇・看護休暇の各月の取得時間は、申請・報告シート（承認済み）から表示されます（「休暇等」シートでも確かめられます）。", False),
    ("    ※ 年次休暇の取得日数は「休暇等」シートの内容が表示されます（本年使用は 申請・報告シートの承認済み＋調整）。", False),
    ("    ※ 出勤欄の○は自動で表示されます。テレワークの予定がある日は「テレ」（出勤登録もあれば「○テレ」）、時間のテレワークは備考に時刻が出ます。年次休暇・病気休暇などの欄は、申請・報告シートの承認済みの申請から自動で表示されます（上書きして手で書くこともできます）。欠勤は記入用の空欄です。", False),
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
    ("・申請・報告シート：「申請・報告シートの読み込み」シートの B2 に申請・報告シートのフォルダを入れておくと、目次の［最新を読み込む］（または申請一覧シートの［申請を取り込む］）を押したときに、承認済みの申請を読み込みます（開いただけでは読み込みません。時間がかかるため）。「外部データ接続が無効」と出たら［コンテンツの有効化］を押してください。", False),
    ("・休暇等シートには、職員ごと・年ごとに1行、本年分などと、申請・報告シートに書かれていない年次休暇の使用があればその調整（日と時間。15分単位、1日＝7時間45分）を入力します。本年使用・病気休暇・看護休暇は申請・報告シートから自動で表示されます。繰越分は前年の翌年繰越が自動で入ります（活用開始時のみ手入力）。翌年に繰り越せるのは本年分の残りだけで（一度繰り越した日数は更に繰り越せない）、上限は20日、1日未満の端数は切り捨てです。", False),
    ("・テレワーク：「テレワーク一覧」シートの B2 に、所属長のテレワーク予定表（テレワーク予定表　所属名.xlsx）を置いたフォルダを入れておくと、［最新を読み込む］を押したときに読み込み、出勤簿に「テレ」と出します。全員の予定と注意（名簿にない ID、休暇と重なる日など）はテレワーク一覧シートで確かめます。", False),
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
header(ws, 1, ["出勤用ID", "氏名", "休憩の始め\n（短縮の人のみ）", "休憩の終わり\n（短縮の人のみ）",
                "部", "課", "並び順", "採用日\n（年の途中のみ）", "退職日\n（年の途中のみ）",
                "養育する子\nの数", "要介護者\nの数", "所属長氏名"], [12, 18, 13, 13, 12, 12, 8, 13, 13, 10, 10, 16])
ws.row_dimensions[1].height = 36
for c in ws[1]:
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
style_range(ws, 2, NAME_LAST, 1, 12, fill=fill_input)
style_range(ws, 2, NAME_LAST, 10, 11, fill=fill_input, align=center)
style_range(ws, 2, NAME_LAST, 7, 7, fill=fill_input, align=center)
style_range(ws, 2, NAME_LAST, 8, 9, fill=fill_input, fmt="yyyy/m/d", align=center)
style_range(ws, 2, NAME_LAST, 3, 4, fill=fill_input, fmt="h:mm", align=center)
style_range(ws, 2, NAME_LAST, 1, 1, fmt="@", align=center)
ws["A2"], ws["B2"] = "1234", "（例）田中 太郎"
ws["O1"] = "2行目は例です。使い始めるときに消してください（氏名に「（例）」とある行は、申請・報告シートを作りません）。"
ws["O2"] = "休憩時間の短縮の特例を使う人だけ、C・D 列に休憩の始めと終わりを入れます（例 12:00、12:30）。空欄の人は 12:00〜13:00 です。"
ws["O3"] = "変えたら、申請・報告シートの作成.xlsm の［申請・報告シートを更新］を押すと、申請・報告シートの時間の計算に反映されます。"
ws["O5"] = "採用日・退職日は、年の途中で採用・退職した人だけ入れます。休暇等シートに表示され、その年の本年分（年次休暇）が在職期間に応じた日数になります（採用日がなければ20日）。"
ws["O6"] = "養育する子の数・要介護者の数は、看護休暇（子の看護等）・短期介護（特別休暇）の1年の上限に使います（1人は5日、2人以上は10日）。申請・報告シートの作成.xlsm の［申請・報告シートを更新］で申請・報告シートに反映され、残りと、上限を超える申請の注意が出ます。"
ws["O4"] = "部・課・並び順は「利用予定」シート（フレックス／テレワーク利用予定）の並びに使います。並び順の小さい順（同じなら上の行から）。部・課は空欄でもかまいません（空欄の人も一覧に出ます）。"
ws["O7"] = "所属長氏名は、その職員の超過勤務等命令簿の「所属長 確認日及び氏名」に、所属長が確認日を入れると自動で出ます。異動したら書き換えて、申請・報告シートの作成.xlsm の［申請・報告シートを更新］を押します（確認済みの月の氏名はそのまま残ります）。"
for _r in range(1, 8):
    ws[f"O{_r}"].font = f_note
# 養育する子の数・要介護者の数は 0〜20 の整数
dv = DataValidation(type="whole", operator="between", formula1="0", formula2="20", allow_blank=True)
dv.errorTitle, dv.error = "人数", "0〜20 の数を入れてください。"
ws.add_data_validation(dv)
dv.add(f"J2:K{NAME_LAST}")
wb.defined_names["職員名簿表"] = DefinedName("職員名簿表", attr_text=f"{N}!$A$1:$B${NAME_LAST}")
ws.freeze_panes = "A2"

# ===== 出力データ貼付（入力専用） =====
ws = wb.create_sheet("出力データ貼付")
header(ws, 1, ["出勤用ID", "QR設置場所", "出勤日", "出勤時刻", "登録日時"], [12, 14, 14, 12, 26])
style_range(ws, 2, DATA_LAST, 1, 5, fill=fill_input, fmt="@")
for col, v in zip("ABCDE", ["1234", "K", "2026-10-01", "09:02:15", "2026-10-01T00:02:15.000Z"]):
    ws[f"{col}2"] = v
ws["G1"] = "← attendance.xlsx のデータ（見出しを除く）をコピーして、このシートの A2 に貼り付けます（2行目は例です）"
ws["G1"].font = f_note
# 貼り付けの代わりに、フォルダに置いた attendance.xlsx を［最新を読み込む］で読む（H3 が空欄なら、貼り付けた内容をそのまま使う）
ws["G3"] = "QR出勤のフォルダ"
ws["G3"].font = Font(name=FONT, size=10, bold=True, color="1E3A5F")
ws["G3"].fill, ws["G3"].border, ws["G3"].alignment = PatternFill("solid", fgColor="E8EEF5"), border, center
ws.merge_cells("H3:P3")
for col in "HIJKLMNOP":
    ws[f"{col}3"].fill, ws[f"{col}3"].border = fill_input, border
ws["H3"].number_format = "@"
ws["H3"].alignment = Alignment(horizontal="left", vertical="center", shrink_to_fit=True)
ws.column_dimensions["G"].width = 18
ws["G4"] = ("↑ 出勤管理システムの［Excel出力］でダウンロードした attendance.xlsx を置くフォルダ（例 C:\\共有\\QR出勤）。"
            "ここを入れておくと、目次の［最新を読み込む］で、フォルダの中のファイルを全部読み、同じ記録を1つにして A〜E 列に入れ直します（貼り付けは要りません）。")
ws["G5"] = "　 ファイルは何回分置いてもかまいません（名前が attendance (1).xlsx などでも読みます）。読んだファイルは「QR出勤の読み込み」シートで確かめられます。空欄なら、貼り付けた内容をそのまま使います。"
for ref in ("G4", "G5"):
    ws[ref].font = f_note
ws.freeze_panes = "A2"

# ===== 計算用（非表示）：職員名簿の照合キーと重複 =====
ws = wb.create_sheet("計算用")
header(ws, 1, ["名簿の照合キー", "重複", "", "", "", "名簿の連番", "",
                "年休キー（ID|年）", "前年の行", "繰越分（分）", "合計（分）", "翌年繰越（分）", "", "",
                "本年使用（分）", "本年分（日）", "在職期間の始め", "在職期間の終わり", "在職期間の区分",
                "申請の年休（分）", "病気休暇の行", "看護休暇の行", "", "利用予定の並び", "k番目の名簿の位置"],
       [16, 8, 4, 4, 4, 10, 4, 16, 8, 10, 10, 10, 4, 4, 10, 10, 12, 12, 10, 12, 10, 10, 14, 12, 12],
       fill=PatternFill("solid", fgColor="7F7F7F"))
for r in range(2, NAME_LAST + 1):
    ws[f"A{r}"] = "=" + key(name_col("A"))
    ws[f"B{r}"] = f'=IF(A{r}="","",IF(COUNTIF($A$2:$A${NAME_LAST},A{r})>1,"重複",""))'
    ws[f"F{r}"] = f'=IF(A{r}="","",COUNTIF($A$2:A{r},"?*"))'
    # 利用予定の並び：名簿の全員（例の行を除く。部は空欄でもよい）を 並び順（空欄は最後）→ 名簿の行 の順に。Y 列は k 番目の人の名簿の位置（1〜）
    ws[f"X{r}"] = (f'=IF(OR(A{r}="",ISNUMBER(SEARCH("（例）",{name_col("B")}))),"",'
                   f'IF(ISNUMBER({name_col("G")}),{name_col("G")},100000)*1000+ROW())')
    ws[f"Y{r}"] = f'=IFERROR(MATCH(SMALL($X$2:$X${NAME_LAST},ROW()-1),$X$2:$X${NAME_LAST},0),"")'
# 休暇等シートの各行（同じ行番号）について、繰越分・合計・翌年繰越を「分」で計算する（1日＝7時間45分＝465分）。
# 休暇等シートの列：A 出勤用ID, B 氏名（自動）, C 年, D 本年分, E 繰越分（活用開始時）, F 採用日, G 退職日, H・I 本年使用の調整（日・時間）
# 繰越分 = 前年の行の翌年繰越（前年の行がなければ、その行に手入力した繰越分）
# 合計 = 繰越分＋本年分（上限40日）
# 翌年繰越 = 残り（合計−本年使用）の1日未満を切り捨てた日数。上限20日（繰越分と本年分は区別しない）
for r in range(2, LEAVE_LAST + 1):
    ws[f"H{r}"] = f'=IF(OR({key(leave_col("A"))}="",{leave_col("C")}=""),"",{key(leave_col("A"))}&"|"&{leave_col("C")})'
    ws[f"I{r}"] = f'=IF(H{r}="","",IFERROR(MATCH({key(leave_col("A"))}&"|"&({leave_col("C")}-1),$H$2:$H${LEAVE_LAST},0),""))'
    y = leave_col("C")
    hire, leave = leave_col("F"), leave_col("G")
    ws[f"Q{r}"] = f'=IF(H{r}="","",IF({hire}="",DATE({y},1,1),MAX(DATE({y},1,1),{hire})))'
    ws[f"R{r}"] = f'=IF(H{r}="","",IF({leave}="",DATE({y},12,31),MIN(DATE({y},12,31),{leave})))'
    ws[f"S{r}"] = (f'=IF(H{r}="","",IF(R{r}<Q{r},-1,DATEDIF(Q{r},R{r}+1,"m")'
                   f'-IF(EDATE(Q{r},DATEDIF(Q{r},R{r}+1,"m"))=R{r}+1,1,0)))')
    ws[f"P{r}"] = (f'=IF(H{r}="","",IF({leave_col("D")}&""<>"",N({leave_col("D")}),'
                   f'IF(AND({hire}="",{leave}=""),{GRANT_DEFAULT},'
                   f'IF(S{r}<0,0,CHOOSE(MIN(S{r},11)+1,{",".join(str(d) for d in PRORATED)})))))')
    # 申請・報告シートの承認済みの年次休暇（分）と、病気休暇・看護休暇の月別_自動の行
    year_row = auto_match(f"H{r}", AU_Y)
    ws[f"T{r}"] = f'=IF(H{r}="","",IFERROR(INDEX({auto_col(col_after(AU_Y, 1))},{year_row}),0))'
    ws[f"U{r}"] = f'=IF(H{r}="","",{auto_match(f"H{r}&" + chr(34) + "|病気休暇" + chr(34), AU_M)})'
    ws[f"V{r}"] = f'=IF(H{r}="","",{auto_match(f"H{r}&" + chr(34) + "|看護休暇" + chr(34), AU_M)})'
    # 本年使用（分）＝ 休暇等シートの調整（手入力）＋ 申請・報告シートの分
    ws[f"O{r}"] = f'=IF(H{r}="","",ROUND(N({leave_col("H")})*{DAY_MIN}+N({leave_col("I")})*1440,0)+T{r})'
    # 活用開始時に手入力する繰越分も、繰越のルールどおり1日未満は切り捨て・上限20日
    ws[f"J{r}"] = (f'=IF(H{r}="","",IF(I{r}<>"",INDEX($L$2:$L${LEAVE_LAST},I{r}),'
                   f'MIN({CARRY_MAX},INT(N({leave_col("E")})))*{DAY_MIN}))')
    ws[f"K{r}"] = f'=IF(H{r}="","",MIN({TOTAL_MAX * DAY_MIN},J{r}+P{r}*{DAY_MIN}))'
    ws[f"L{r}"] = (f'=IF(H{r}="","",MIN({CARRY_MAX * DAY_MIN},'
                   f'INT(MAX(0,K{r}-O{r})/{DAY_MIN})*{DAY_MIN}))')
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
KIND_COLS = (3, 7)              # TABLE のうち特別休暇（種別）・職専免（種別）
HID_COUNT, HID = "AB", "AC"     # 隠し列：AB=その日の登録件数、AC=照合キー・名簿の行位置・対象月の1日
HID_DAY = "AD"                  # 隠し列：その日の申請（日別_自動）の行位置
HID_TELE = "AA"                 # 隠し列：その日のテレワークの予定（テレワーク_日別）の行位置
HID_FLEX = "AG"                 # 隠し列：その日のフレックスの割振り（フレックス_出勤簿）の行位置
MIN_FIRST_COL = 34              # 隠し列 AH〜AS：出勤簿の上の欄の各月・累計の分
CTRL_L, CTRL_V = "AE", "AF"     # 印刷範囲の外：対象年月・出勤用ID などの指定欄
wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)
thin_side = Side(style="thin", color="BFBFBF")
none_side = Side(style=None)


def setup_report_sheet(ws):
    ws.column_dimensions["A"].width = 6.5
    ws.column_dimensions["B"].width = 4.5
    for col in SUB:
        ws.column_dimensions[col].width = SUB_W
    for col in [HID_TELE, HID_COUNT, HID, HID_DAY, HID_FLEX] + [get_column_letter(MIN_FIRST_COL + i) for i in range(12)]:
        ws.column_dimensions[col].width = 8
        ws.column_dimensions[col].hidden = True
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


def day_hm_text(m):
    """分（m はセル参照）を「1日と2時間30分」の形の文字にする。1日（7時間45分）未満は「2時間30分」。m が空欄なら空欄。"""
    rest = f"MOD({m},{DAY_MIN})"
    return (f'=IF({m}="","",IF({m}>={DAY_MIN},INT({m}/{DAY_MIN})&"日"&IF({rest}>0,"と",""),"")'
            f'&IF(INT({rest}/60)>0,INT({rest}/60)&"時間","")&IF(MOD({rest},60)>0,MOD({rest},60)&"分",""))')


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
    # 取得時間は 申請・報告シートの読み込みシートの月別_自動の、その年・区分の行（AC{t+4}〜AC{t+5}=その行番号）から表示する
    value_font = Font(name=FONT, size=11, bold=True)
    box(ws, f"A{t + 1}", f"B{t + 3}", "所属")
    box(ws, f"C{t + 1}", f"K{t + 3}", None, value_font, fill_white, left)
    box(ws, f"A{t + 4}", f"B{t + 6}", "氏名")
    box(ws, f"C{t + 4}", f"K{t + 6}",
        f'=IF({P1}="","",IF({P2}="","★名簿に未登録",INDEX({N}!$B$2:$B${NAME_LAST},{P2})&""))',
        value_font, fill_white, left)
    small_label = Font(name=FONT, size=8, bold=True, color="1E3A5F")
    hour_font = Font(name=FONT, size=6)    # 「1日と2時間30分」を細い欄に2行で入れる
    minutes_cols = [get_column_letter(MIN_FIRST_COL + i) for i in range(12)]   # 隠し列：各月・累計の分
    # L列は2〜7行目を結合して縦書きの「取得時間」。M:N は区分名、O〜Z は月
    box(ws, f"L{t + 1}", f"L{t + 6}", "取得時間", f_label,
        align=Alignment(horizontal="center", vertical="center", text_rotation=255))
    # 見出し（1月〜12月）は2行分の高さ、その下に区分ごとに「その月」と「累計」の2行
    box(ws, f"M{t + 1}", f"N{t + 2}", None, small_label)
    for i, col in enumerate(SUB[12:]):          # O〜Z
        box(ws, f"{col}{t + 1}", f"{col}{t + 2}", f"{i + 1}月", small_label,
            align=Alignment(horizontal="center", vertical="center", shrink_to_fit=True))   # 10月〜12月は細い欄に縮めて入れる
    for j, cat in enumerate(HOUR_CATS):
        r = t + 3 + j * 2                       # その月の取得時間
        rc = r + 1                              # 1月からその月までの累計
        AR = f"${HID}${t + 4 + j}"
        month_key = P1 + '&"|"&YEAR(' + P3 + ')&"|' + cat + '"'
        ws[f"{HID}{t + 4 + j}"] = f'=IF(OR({P1}="",{P3}=""),"",{auto_match(month_key, AU_M)})'
        box(ws, f"M{r}", f"N{r}", cat, small_label)
        box(ws, f"M{rc}", f"N{rc}", "（累計）", small_label)
        first_a = f"INDEX({auto_col(col_after(AU_M, 1))},{AR})"
        for i, col in enumerate(SUB[12:]):
            acol = col_after(AU_M, 1 + i)       # 月別_自動の i+1 月の列
            a = f"INDEX({auto_col(acol)},{AR})"
            month = f'IF({AR}="",0,N({a}))'
            mc = minutes_cols[i]
            ws[f"{mc}{r}"] = f'=IF(OR({P1}="",{P3}=""),"",IF({month}=0,"",ROUND(({month})*1440,0)))'
            box(ws, f"{col}{r}", f"{col}{r}", day_hm_text(f"{mc}{r}"), hour_font, fill_white, wrap)
            # 累計は対象年月の月まで表示（それより後の月と、累計が0の月は空欄）
            total = f'IF({AR}="",0,SUM({first_a}:{a}))'
            ws[f"{mc}{rc}"] = (f'=IF(OR({P1}="",{P3}=""),"",IF({i + 1}>MONTH({P3}),"",'
                               f'IF({total}=0,"",ROUND(({total})*1440,0))))')
            box(ws, f"{col}{rc}", f"{col}{rc}", day_hm_text(f"{mc}{rc}"), hour_font, fill_white, wrap)
    # 見出し（2行）は低く、取得時間（4行）は「1日と3時間15分」が3行で入る高さに
    for r in range(t + 1, t + 3):
        ws.row_dimensions[r].height = 14
    for r in range(t + 3, t + 7):
        ws.row_dimensions[r].height = 25
    ws.row_dimensions[t + 7].height = 8

    # 9〜10行目：年次休暇の取得日数（休暇等シートのその年の行から表示。AC{t+3}=その行の位置）
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
            # 「特別休暇（種別）」「特別休暇（時間）」は 9pt だと3行になるので、8pt にして2行に収める
            font = Font(name=FONT, size=8, bold=True, color="FFFFFF") if text.startswith("特別休暇") else head_font
            box(ws, f"{first}{hr}", f"{last}{hr}", text, font, fill_head, wrap)
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
            if k in KIND_COLS:    # 特別休暇・職専免の種別：2本を結合して、小さめの文字で折り返す
                cells[0].border = border
                cells[0].font = Font(name=FONT, size=7)
                cells[0].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                ws.merge_cells(f"{SPANS[k][0]}{r}:{SPANS[k][1]}{r}")
            elif k <= 8:      # 出勤の○と、申請・報告シートから入る欄（年次休暇〜職専免（時間））は2本の真ん中に
                for c in cells:   # 「選択範囲内で中央」は両方のセルに指定が必要
                    c.alignment = Alignment(horizontal="centerContinuous", vertical="center")
                    if k:
                        c.font = Font(name=FONT, size=9)
            else:
                cells[0].alignment = left
        day_key = P1 + f'&"|"&TEXT(A{r},"yyyy-mm-dd")'
        # 出勤欄：QR の出勤登録があれば「○」、テレワークの予定（テレワーク_日別）があれば「テレ」（両方なら「○テレ」）。
        # 時間のテレワークは備考に時刻を出す（隠し列 AA がテレワーク_日別の行位置）
        ws[f"{HID_TELE}{r}"] = f'=IF(OR({P1}="",A{r}=""),"",{auto_match(day_key, AU_T)})'
        ws[f"C{r}"] = (f'=IF(N({HID_COUNT}{r})=0,"","○")'
                       f'&IF({HID_TELE}{r}="","",INDEX({auto_col(col_after(AU_T, 1))},{HID_TELE}{r})&"")')
        # 備考：時間のテレワークの時刻と、フレックスの割振り（取り込んだ承認済みの内容。隠し列 AG がフレックス_出勤簿の行位置）
        ws[f"{HID_FLEX}{r}"] = f'=IF(OR({P1}="",A{r}=""),"",{auto_match(day_key, AU_F)})'
        tele_note = f'IF({HID_TELE}{r}="","",INDEX({auto_col(col_after(AU_T, 2))},{HID_TELE}{r})&"")'
        flex_note = f'IF({HID_FLEX}{r}="","",INDEX({auto_col(col_after(AU_F, 1))},{HID_FLEX}{r})&"")'
        ws[f"{SPANS[-1][0]}{r}"] = f'=TRIM({flex_note}&" "&{tele_note})'
        ws[f"{SPANS[-1][0]}{r}"].font = Font(name=FONT, size=8)
        # 年次休暇〜職専免（時間）の欄（TABLE の 1〜8 番目）は、申請・報告シートの承認済みの申請から表示する
        ws[f"{HID_DAY}{r}"] = f'=IF(OR({P1}="",A{r}=""),"",{auto_match(day_key, AU_D)})'
        for k in range(1, 9):
            acol = col_after(AU_D, k)           # 日別_自動の k 番目の欄の列
            ws[f"{SPANS[k][0]}{r}"] = (f'=IF({HID_DAY}{r}="","",INDEX({auto_col(acol)},{HID_DAY}{r})&"")')
        ws.row_dimensions[r].height = 19


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

# ===== 休暇等（職員ごと・年ごとに1行。年次休暇の入力欄と、申請・報告シートから読んだ値・自動計算の表示） =====
ws = wb.create_sheet("休暇等")
INPUTS = [("出勤用ID", 12), ("氏名\n（名簿から）", 14), ("年（西暦）", 9), ("本年分（日）\n空欄は自動", 11), ("繰越分（日）\n前年の翌年繰越", 13),
          ("採用日\n（名簿から）", 12), ("退職日\n（名簿から）", 12),
          ("本年使用\n（手入力：日）", 11), ("本年使用\n（手入力：時間）", 12)]                    # A〜I（B 氏名は自動）
AUTOS = [("本年分\n（自動）", 9), ("繰越分\n（自動）", 15), ("合計\n（自動）", 15),
         ("本年使用\n（申請・報告シート）", 15), ("本年使用\n（合計）", 15), ("残り\n（自動）", 15),
         ("翌年繰越\n（自動）", 15)]           # J〜P 「28日3時間45分」が入る幅
MONTH_COLS = {cat: [get_column_letter(17 + j * 12 + i) for i in range(12)] for j, cat in enumerate(HOUR_CATS)}
NOTE_COL = get_column_letter(17 + 24)                                                             # AO 備考
labels = [x for x, _ in INPUTS + AUTOS]
widths = [w for _, w in INPUTS + AUTOS]
for cat in HOUR_CATS:
    labels += [f"{cat[:2]}\n{i + 1}月" for i in range(12)]
    widths += [6] * 12
labels.append("備考")
widths.append(24)
header(ws, 1, labels, widths)
ws.row_dimensions[1].height = 50
for c in ws[1]:
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
fill_month = PatternFill("solid", fgColor="5B4A7A")      # 病気休暇・看護休暇の見出しは色を変えて区切る
for cols in MONTH_COLS.values():
    for col in cols:
        ws[f"{col}1"].fill = fill_month
last_month = MONTH_COLS[HOUR_CATS[-1]][-1]
style_range(ws, 2, LEAVE_LAST, 1, 9, fill=fill_input, align=center)
style_range(ws, 2, LEAVE_LAST, 1, 1, fmt="@", align=center)
style_range(ws, 2, LEAVE_LAST, 2, 2, fill=fill_calc, align=left)
style_range(ws, 2, LEAVE_LAST, 6, 7, fill=fill_calc, fmt="yyyy/m/d", align=center)       # 採用日・退職日は職員名簿から
style_range(ws, 2, LEAVE_LAST, 9, 9, fill=fill_input, fmt="[h]:mm", align=center)   # 時間は「時間:分」（例 2:15）
style_range(ws, 2, LEAVE_LAST, 10, 16, fill=fill_calc, align=center)
style_range(ws, 2, LEAVE_LAST, 17, 40, fill=fill_calc, fmt="[h]:mm", align=center)
style_range(ws, 2, LEAVE_LAST, 41, 41, fill=fill_input)
row_key = f"INDEX({K}!$H:$H,ROW())"


def calc(col):
    return f"INDEX({K}!${col}:${col},ROW())"


for r in range(2, LEAVE_LAST + 1):
    ws[f"{NOTE_COL}{r}"].alignment = left
    # 氏名：職員名簿から（出勤用ID の照合キーで探す）
    ws[f"B{r}"] = (f'=IF(A{r}="","",IFERROR(INDEX({N}!$B$2:$B${NAME_LAST},MATCH({key(f"A{r}")},{K}!$A$2:$A${NAME_LAST},0))&"",'
                   f'"★名簿に未登録"))')
    # 採用日・退職日：職員名簿から（なければ空欄）
    pos = f'MATCH({key(f"A{r}")},{K}!$A$2:$A${NAME_LAST},0)'
    for col, src in (("F", "H"), ("G", "I")):
        v = f"INDEX({N}!${src}$2:${src}${NAME_LAST},{pos})"
        ws[f"{col}{r}"] = f'=IF(A{r}="","",IFERROR(IF({v}="","",{v}),""))'
    # 繰越分：前の年の行があれば、その行の翌年繰越（日）。なければ空欄（使い始めの年だけ、この式の上に手で入れる）
    ws[f"E{r}"] = f'=IF({calc("I")}="","",INDEX({K}!$L$2:$L${LEAVE_LAST},{calc("I")})/{DAY_MIN})'
    ws[f"J{r}"] = f'=IF({row_key}="","",{calc("P")}&"日")'
    for col, expr in (("K", calc("J")), ("L", calc("K")), ("M", calc("T")), ("N", calc("O")),
                      ("O", f'({calc("K")}-{calc("O")})'), ("P", calc("L"))):
        ws[f"{col}{r}"] = f'=IF({row_key}="","",{days_text(expr)})'
    # 病気休暇・看護休暇の1月〜12月：申請・報告シートの読み込みシートの月別_自動の行（計算用 U・V 列）から
    for cat, idx_col in zip(HOUR_CATS, ("U", "V")):
        idx = calc(idx_col)
        for i, col in enumerate(MONTH_COLS[cat]):
            v = f"INDEX({auto_col(col_after(AU_M, 1 + i))},{idx})"
            ws[f"{col}{r}"] = f'=IF({idx}="","",IF(N({v})=0,"",{v}))'
ws["A2"], ws["C2"], ws["D2"], ws["E2"] = "1234", 2026, 20, 10     # 例の行（繰越分は使い始めの年なので手入力）
ws[f"{NOTE_COL}2"] = "（例）この行は例です。削除してください"
ws.freeze_panes = "D2"
note_col = get_column_letter(17 + 24 + 2)    # AQ（使い方の文。1行目は［名簿の人の行を足す］のボタン）
for i, text in enumerate([
    "・職員ごと・年ごとに1行入力します（例：出勤用ID 1234 の2026年分）。黄色の欄が入力欄、灰色の欄は自動です。氏名は職員名簿から自動で入ります。",
    "・上の［名簿の人の行を足す］（目次にもあります）を押して年を入れると、職員名簿の全員について、その年の行（出勤用ID・年）がまだない人の分だけ下に足します。年が変わったときや、名簿に人を足したときに押します。",
    "・灰色の欄には計算式が入っています。行を消すときは行ごと削除し（行番号を右クリック →［削除］）、行を選んで Delete キーで消さないでください（計算式が消えます）。",
    f"・本年分（その年に付与された日数）は、通年で勤務する場合は空欄のままで構いません（{GRANT_DEFAULT}日として計算）。",
    "・採用日・退職日は職員名簿（J・K列）に入れると、ここに表示されます。その年の途中の採用・退職なら、在職期間に応じた本年分が自動で入ります（1月に達するまで2日〜11月を超え1年未満20日）。採用日がなければ20日です。",
    "・本年分を入力した場合は、入力した日数が優先されます。",
    "・繰越分（E列）は、前の年の行があれば、その行の「翌年繰越」（P列）の日数が自動で入ります。",
    f"・使い始めの年（前の年の行がない年）だけ、E列に繰越分の日数を手入力してください（式の上に書きます）。繰越分は日単位なので、1日未満の端数は切り捨てます（上限{CARRY_MAX}日）。",
    "・本年使用（申請・報告シート）は、申請・報告シートの承認済みの年次休暇です（「申請・報告シートの読み込み」シートの B2 のフォルダから、目次の［最新を読み込む］を押したときに読み込みます）。",
    "・申請・報告シートに書かれていない年次休暇の使用（この仕組みを使い始める前の使用など）があるときだけ、H・I列（本年使用（手入力））に日数と時間（「時間:分」、例 2:15）を入力します。申請・報告シートの分に足されます。普段は空欄です。",
    "・年次休暇は15分単位で、1日は7時間45分として計算します。7時間45分以上の時間は自動で日数に繰り上がります。",
    f"・合計は繰越分＋本年分です（上限{TOTAL_MAX}日）。残りは 合計−本年使用（合計） です。",
    f"・翌年繰越は、残りの1日未満の端数を切り捨てた日数です（上限{CARRY_MAX}日）。繰越分と本年分は区別しません。",
    "・病気休暇・看護休暇の1月〜12月は、申請・報告シートの承認済みの申請を月ごとに足した時間です（開始日の月。日単位は1日＝7時間45分）。直したいときは申請・報告シートを直します。",
    "・この行の内容は、出勤簿の「年次休暇 取得日数」と「取得時間」の欄に表示されます（取得時間は、この行がなくても申請・報告シートから表示されます）。",
], start=2):
    ws[f"{note_col}{i}"] = text
    ws[f"{note_col}{i}"].font = f_note

# ===== 申請・報告シートの読み込み（申請・報告シートを Power Query で読んだ表。表は add_queries.ps1 が作る） =====
ws = wb.create_sheet("申請・報告シートの読み込み")
ws["A1"] = "休暇の申請の読み込み"
ws["A1"].font = f_title
for r, (label, note) in enumerate([
    ("申請・報告シートのフォルダ", "← 申請・報告シート（申請・報告シート　出勤用ID　氏名.xlsx）を置いた OneDrive の同期フォルダの場所（使わなければ空欄）"),
], start=2):
    ws[f"A{r}"] = label
    ws[f"A{r}"].font = f_bold
    ws[f"B{r}"].fill = fill_input
    ws[f"B{r}"].border = border
    ws.merge_cells(f"B{r}:J{r}")
    ws[f"L{r}"] = note
    ws[f"L{r}"].font = f_note
ws["A5"] = "前回の取り込み"
ws["A5"].font = f_bold
ws["B5"].number_format = "yyyy/m/d h:mm"
ws["B5"].alignment = left
ws["L5"] = "目次の［最新を読み込む］（または申請一覧シートの［申請を取り込む］）を押したときに読み込みます（1分ほどかかります。開いただけでは読み込みません）。読むのは保存された内容です。6行目は読み込んだ行数です。"
ws["L5"].font = f_note
ws.column_dimensions["B"].width = 19
for col, label in zip(AUTO_TABLES.values(), ["年次休暇（出勤用ID|年 ごとの使用・分）", "病気休暇・看護休暇（月ごと）", "出勤簿の日ごとの欄",
                                              "テレワーク（出勤簿の出勤欄・備考）", "フレックス（出勤簿の備考）", "利用予定（日ごとの文字）"]):
    ws[f"{col}{AUTO_HEAD - 1}"] = label
    ws[f"{col}{AUTO_HEAD - 1}"].font = f_bold
    # 6行目：表の行数（出勤簿などがこの行数の範囲だけを探す。式は add_queries.ps1 が入れる）
    ws[f"{col}{AUTO_COUNT_ROW}"].font = f_note
    ws[f"{col}{AUTO_COUNT_ROW}"].number_format = '"（"0"行）"'
ws.column_dimensions["A"].width = 18
ws.column_dimensions[AU_M].width = 22
ws.column_dimensions[AU_D].width = 18
for _c in range(1, 9):     # 日別_自動の8欄（見出しが入る幅）
    ws.column_dimensions[col_after(AU_D, _c)].width = 16
ws.column_dimensions[AU_T].width = 18
ws.column_dimensions[AU_F].width = 18
ws.column_dimensions[AU_U].width = 18
ws.column_dimensions[col_after(AU_F, 1)].width = 16
ws.column_dimensions[col_after(AU_U, 1)].width = 18
wb.defined_names["フォルダ"] = DefinedName("フォルダ", attr_text=f"{AU}!$B$2")
wb.defined_names["QR出勤フォルダ"] = DefinedName("QR出勤フォルダ", attr_text="'出力データ貼付'!$H$3")
# QR出勤の読み込み：出力データ貼付 H3 のフォルダで読んだファイル（表 QR出勤ファイル一覧、A5 から）と、つないだ出勤記録（表 QR出勤、G5 から）。
# 表は add_queries.ps1 が作る。［最新を読み込む］のマクロが 表 QR出勤 を「出力データ貼付」に写す
ws = wb.create_sheet("QR出勤の読み込み")
ws["A1"] = ("出力データ貼付シートの H3 のフォルダで読んだ、出勤管理システムの［Excel出力］のファイル（左）と、つないだ出勤記録（右。同じ記録は1つ）です。"
            "目次の［最新を読み込む］で読み直し、右の表を「出力データ貼付」の A〜E 列に入れます。")
ws["A1"].font = f_note
ws["A3"] = "読んだファイル"
ws["G3"] = "つないだ出勤記録"
ws["A3"].font = ws["G3"].font = f_bold

# ===== 申請一覧（申請・報告シートの全員の申請と状況。表は add_queries.ps1 が作る） =====
ws = wb.create_sheet("申請一覧")
ws["A1"] = "全員の休暇の申請（申請・報告シート）。目次の［最新を読み込む］（またはこのシートの［申請を取り込む］）を押すと最新になります。見出しの ▼ で絞り込めます。"
ws["A1"].font = f_note
wb.defined_names["祝日一覧"] = DefinedName("祝日一覧", attr_text="'祝日'!$A$2:$A$200")

# ===== テレワーク一覧（所属長のテレワーク予定表を読んだ全員の予定と注意。表は add_queries.ps1 が作る） =====
ws = wb.create_sheet("テレワーク一覧")
ws["A1"] = "全所属のテレワークの予定（所属長のテレワーク予定表）。目次の［最新を読み込む］を押すと最新になります。「注意」の列に確かめることが出ます。"
ws["A1"].font = f_note
ws["A2"] = "テレワーク予定表のフォルダ"
ws["A2"].font = f_bold
ws["B2"].fill = fill_input
ws["B2"].border = border
ws.merge_cells("B2:H2")
ws["J2"] = "← テレワーク予定表（テレワーク予定表　所属名.xlsx）を置いた OneDrive の同期フォルダの場所（使わなければ空欄）"
ws["J2"].font = f_note
ws.column_dimensions["A"].width = 38     # A2「テレワーク予定表のフォルダ」が入る幅
wb.defined_names["テレワークフォルダ"] = DefinedName("テレワークフォルダ", attr_text=f"{TW}!$B$2")

# ===== 超過勤務（申請・報告シートの超過勤務等命令簿。tools/overtime） =====
# 超過勤務：B2 の月の所属長の確認済みの分を［超過勤務を転記］で写した、1人1か月の計（値。あとで命令簿が変わっても変わらない）
# 超過勤務の日別：同じく写した日ごとの行と、出勤の記録・休暇との照合（照合の列だけは式）
# 超過勤務の読み込み：命令簿を Power Query で読んだ今の内容（確認待ちも含む）。表はどれも add_queries.ps1 が作る
ws = wb.create_sheet("超過勤務")
ws["A1"] = "超過勤務等命令簿（申請・報告シートの月のシート）の、所属長の確認済みの月を写したものです（1人1か月で1行）。日ごとの行と照合は「超過勤務の日別」シート。"
ws["A1"].font = f_note
ws["A2"] = "転記する月"
ws["A2"].font = f_bold
ws["B2"] = date(2026, 11, 1)
ws["B2"].number_format = 'yyyy"年"m"月"'
ws["B2"].fill, ws["B2"].font, ws["B2"].border = fill_input, f_bold, border
ws["B2"].alignment = left
ws["A3"] = ("↑ 月を入れて［超過勤務を転記］を押すと、その月の所属長の確認済みの命令簿を写します（同じ人・同じ月を前に写していれば置き換えます）。"
            "確認待ちの人は写しません（「超過勤務の読み込み」シートで確かめられます）。")
ws["A3"].font = f_note
ws.column_dimensions["A"].width = 12
ws.column_dimensions["B"].width = 16
ws = wb.create_sheet("超過勤務の日別")
ws["A1"] = ("［超過勤務を転記］で写した日ごとの行です。「注意」の列に、出勤の記録（QR の出勤登録・テレワークの予定）がない日、"
            "承認済みの休暇がある日が出ます（出勤記録の貼り付けや休暇の承認が変われば、注意も変わります）。")
ws["A1"].font = f_note
# 命令簿の保存：所属長の確認済みの命令簿を［命令簿をPDFで保存］で PDF にした記録（マクロが6行目から足す）。B2 は保存先のフォルダ
ws = wb.create_sheet("命令簿の保存")
ws["A1"] = ("所属長の確認済みの超過勤務等命令簿を、1人1か月1ファイルの PDF で保管します。［命令簿をPDFで保存］を押すと、"
            "「超過勤務の読み込み」シートで確認済みの月のうち、まだ保存していない月を保存します（先に目次の［最新を読み込む］を押してください）。")
ws["A1"].font = f_note
ws["A2"] = "保存先のフォルダ"
ws["A2"].font = f_bold
ws.merge_cells("B2:F2")
ws["B2"].fill, ws["B2"].border, ws["B2"].alignment = fill_input, border, left
for _col in "CDEF":
    ws[f"{_col}2"].border = border
ws["A3"] = ("↑ 空欄なら、申請・報告シートのフォルダ（申請・報告シートの読み込みシートの B2）と同じ場所に「超過勤務等命令簿」フォルダを作って保存します。"
            "保存したあとで命令簿が直されたときは、前の PDF を残したまま「…（2）.pdf」の名前で保存します。")
ws["A3"].font = f_note
ws.row_dimensions[2].height = 30
header(ws, 5, ["保存日時", "年月", "出勤用ID", "氏名", "確認日", "ファイル名", "内容の記録"], [17, 9, 10, 14, 12, 52, 18])
ws.column_dimensions["A"].width = 17
for _r in range(6, 2006):
    ws[f"A{_r}"].number_format = "yyyy/m/d h:mm"
    ws[f"E{_r}"].number_format = "yyyy/m/d"
    ws[f"G{_r}"].font = f_note
ws["H5"] = "内容の記録：保存したときの命令簿の内容から作った文字です。次に押したとき、これと違えば直されたと見て、別の名前で保存します。この表の行は消さないでください。"
ws["H5"].font = f_note
ws.freeze_panes = "A6"
ws = wb.create_sheet("超過勤務の読み込み")
ws["A1"] = ("申請・報告シートの超過勤務等命令簿を読んだ今の内容です（目次の［最新を読み込む］、または［超過勤務を転記］で最新になります）。"
            "左が1人1か月の計と所属長の確認、右（AG 列から）が日ごとの行です。")
ws["A1"].font = f_note

# ===== フレックスタイム制（申請・報告シートのフレックス申告・割振り。tools/flex） =====
# フレックス：単位期間ごとの状況（今の内容）と取り込み・注意（Power Query の表 フレックス一覧）
# フレックスの日別：所属長の承認済みの単位期間を［最新を読み込む］で写した日ごとの行（値。照合の列だけは式）
# フレックスの読み込み：申請・報告シートのフレックスの表を読んだ今の内容（承認待ちも含む）。表はどれも add_queries.ps1 が作る
ws = wb.create_sheet("フレックス")
ws["A1"] = ("フレックスタイム制で勤務する人の単位期間（4週間）ごとの状況です（申請・報告シートの今の内容）。所属長が承認した単位期間は、"
            "［最新を読み込む］で取り込みます（取り込み：取り込み済み）。「注意」の列に確かめること、「照合の注意」の列に取り込んだ日ごとの照合の注意"
            "（出勤の記録がない勤務日など）が出ます。行を選んで［申告・割振り簿を見る］を押すと、その単位期間を総務課の様式（申告・割振り簿）で表示します。"
            "日ごとの行は「フレックスの日別」シート（ふだんは非表示。シート見出しを右クリック→［再表示］）にあります。")
ws["A1"].font = f_note
ws = wb.create_sheet("フレックスの日別")
# ふだんは見ない（照合の注意はフレックスシートに出る）ので、add_buttons.ps1 が最後に非表示にする（クエリを作る間は表示しておく）
ws["A1"] = ("取り込んだ承認済みの単位期間の日ごとの割振りです（あとで申告が書き換えられても変わりません。承認し直すと取り込み直します）。"
            "「注意」の列に、今日より前の日で、出勤の記録（QR・テレワーク）がない勤務日などが出ます。")
ws["A1"].font = f_note
ws = wb.create_sheet("フレックスの読み込み")
ws["A1"] = "申請・報告シートのフレックスの申告・割振りを読んだ今の内容です（承認待ちも含む。目次の［最新を読み込む］で最新になります）。"
ws["A1"].font = f_note

# ===== テレワーク報告（申請・報告シートのテレワーク実施報告の全員分。表 テレワーク報告一覧 は add_queries.ps1 が作る） =====
ws = wb.create_sheet("テレワーク報告")
ws["A1"] = ("全員のテレワーク実施報告（申請・報告シートの「テレワーク報告」）と所属長の確認です。目次の［最新を読み込む］で最新になります。"
            "「注意」の列に、入力の不備、テレワーク予定表に予定がない日、提出から1週間たっても確認がないものが出ます。"
            "実施日がテレワーク一覧（テレワーク予定表）にない報告は、状況が「予定なし」で赤くなります（予定と報告が合わないので、所属長に確かめます）。"
            "報告の行を選んで［実施報告書を見る］を押すと、その職員の申請・報告シートの実施報告書（別紙３の様式）で表示します。"
            "［実施報告書をPDFで保存］は、選んだ行（何行でも）を1件1ファイルの PDF にして、申請・報告シートのフォルダと同じ場所の「テレワーク実施報告書」フォルダに置きます。")
ws["A1"].font = f_note
# 状況（N 列）が「予定なし」の行を赤く（表 テレワーク報告一覧 は A3 から。列は テレワーク報告一覧.pq の順）
ws.conditional_formatting.add("A4:P5000", FormulaRule(formula=['$N4="予定なし"'], fill=PatternFill("solid", fgColor="F4B6B0", bgColor="F4B6B0"),
                                                      font=Font(name=FONT, color="9C0006")))

# ===== 利用予定（フレックス／テレワーク利用予定の一覧：部・課・職員 × その月の日付。配る用の表は、マクロ［その日のシートに書き出す］が値で作る） =====
# 行は職員名簿の全員（部は空欄でもよい）を 並び順 → 名簿の行 の順（計算用 X・Y 列）。日付の欄は 利用予定_日別（休暇・テレ・フレックスの勤務開始）。
SCHED, S_FIRST = "利用予定", 6
S_LAST = S_FIRST + STAFF_ROWS - 1
S_DAYS = [get_column_letter(4 + i) for i in range(31)]          # D〜AH
S_HID_POS, S_HID_KEY = "AJ", "AK"                                # 隠し列：k 番目の人の名簿の位置、照合キー
S_CTRL_L, S_CTRL_V = "AM", "AN"                                  # 印刷範囲の外：対象月・情報の日時
ws = wb.create_sheet(SCHED)
ws.sheet_view.showGridLines = False
for col, width in (("A", 8), ("B", 10), ("C", 14)):
    ws.column_dimensions[col].width = width
for col in S_DAYS:
    ws.column_dimensions[col].width = 5.2
ws.column_dimensions["AI"].width = 2
for col in (S_HID_POS, S_HID_KEY):
    ws.column_dimensions[col].hidden = True
ws.column_dimensions[S_CTRL_L].width = 12
ws.column_dimensions[S_CTRL_V].width = 16
control(ws, f"{S_CTRL_L}1", "対象月", f_label, fill_label)
control(ws, f"{S_CTRL_V}1", date(2026, 11, 1), f_bold, fill_input, 'yyyy"年"m"月"')
control(ws, f"{S_CTRL_L}2", "情報の日時", f_label, fill_label)
control(ws, f"{S_CTRL_V}2", f"={AU}!$B$5", f_base, fill_calc, "m/d h:mm")
MONTH1 = f"${S_CTRL_V}$1"
ws[f"{S_CTRL_L}3"] = ("↑ AN1 に対象月を入れ、目次の［最新を読み込む］のあとで、下のボタンを押すと、AN9 のファイル（なければ作ります）の"
                      "今日の日付のシート（例 11月16日）にこの表を値で書き出して保存します（配る用。このブックの中には作りません）。")
ws[f"{S_CTRL_L}3"].font = f_note
ws[f"{S_CTRL_L}3"].alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells(f"{S_CTRL_L}3:{S_CTRL_V}6")      # 説明はボタン（7行目）の上
# 書き出す先のファイル（フォルダも含めた場所。例 C:\共有\フレックス・テレワーク利用予定.xlsx）
control(ws, f"{S_CTRL_L}9", "書き出す先", f_label, fill_label)
control(ws, f"{S_CTRL_V}9", None, f_base, fill_input, "@")
ws.merge_cells(f"{S_CTRL_V}9:AU9")
ws[f"{S_CTRL_V}9"].alignment = Alignment(horizontal="left", vertical="center", shrink_to_fit=True)
ws[f"{S_CTRL_L}10"] = "↑ 書き出す先のファイルの場所（フォルダも含めて。例 C:\共有\フレックス・テレワーク利用予定.xlsx）。同じファイルに日付ごとのシートが増えていきます。"
ws[f"{S_CTRL_L}10"].font = f_note
ws.row_dimensions[3].height = 13
# 1〜2行目：題名と「令和8年11月16日9:00現在の情報を基に更新」
ws.merge_cells("A1:T2")
ws["A1"] = f'=IF({MONTH1}="","フレックス／テレワーク利用予定","令和"&(YEAR({MONTH1})-2018)&"年"&MONTH({MONTH1})&"月　フレックス／テレワーク利用予定")'
ws["A1"].font = Font(name=FONT, size=12, bold=True, color="1E3A5F")
ws["A1"].alignment = left
ws.merge_cells("V1:AH2")
ws["V1"] = (f'=IF(${S_CTRL_V}$2="","",IF(YEAR(${S_CTRL_V}$2)-2018=1,"令和元",'
            f'"令和"&(YEAR(${S_CTRL_V}$2)-2018))&"年"&TEXT(${S_CTRL_V}$2,"m月d日h:mm")&"現在の情報を基に更新")')
ws["V1"].font = f_note
ws["V1"].alignment = Alignment(horizontal="right", vertical="center")
# 3〜5行目：月・日・曜日
head_fill = PatternFill("solid", fgColor="DDEBF7")
ws.merge_cells("A3:C5")
ws["A3"] = "氏名"
ws.merge_cells(f"D3:{S_DAYS[-1]}3")
ws["D3"] = f'=IF({MONTH1}="","",MONTH({MONTH1})&"月")'
for i, col in enumerate(S_DAYS, start=1):
    ws[f"{col}4"] = f'=IF({MONTH1}="","",IF({i}<=DAY(EOMONTH({MONTH1},0)),DATE(YEAR({MONTH1}),MONTH({MONTH1}),{i}),""))'
    ws[f"{col}4"].number_format = "d"
    ws[f"{col}5"] = f'=IF({col}4="","",TEXT({col}4,"aaa"))'
for row in ws[f"A3:{S_DAYS[-1]}5"]:
    for c in row:
        c.font, c.fill, c.border, c.alignment = Font(name=FONT, size=9, bold=True), head_fill, border, center
small = Font(name=FONT, size=9)
day_align = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
for r in range(S_FIRST, S_LAST + 1):
    k = r - S_FIRST + 1
    ws[f"{S_HID_POS}{r}"] = f"=INDEX({K}!$Y$2:$Y${NAME_LAST},{k})"
    ws[f"{S_HID_KEY}{r}"] = f'=IF({S_HID_POS}{r}="","",INDEX({K}!$A$2:$A${NAME_LAST},{S_HID_POS}{r}))'
    for col, src in (("A", "E"), ("B", "F"), ("C", "B")):
        ws[f"{col}{r}"] = f'=IF(${S_HID_POS}{r}="","",INDEX({N}!${src}$2:${src}${NAME_LAST},${S_HID_POS}{r})&"")'
        ws[f"{col}{r}"].font, ws[f"{col}{r}"].border = small, border
        ws[f"{col}{r}"].alignment = Alignment(horizontal="left" if col == "C" else "center", vertical="center", wrap_text=True)
    day_key = f'${S_HID_KEY}{r}&"|"&TEXT({{c}}$4,"yyyy-mm-dd")'
    for col in S_DAYS:
        kk = day_key.replace("{c}", col)
        ws[f"{col}{r}"] = (f'=IF(OR(${S_HID_KEY}{r}="",{col}$4=""),"",'
                           f'IFERROR(INDEX({auto_col(col_after(AU_U, 1))},{auto_match(kk, AU_U)})&"",""))')
        ws[f"{col}{r}"].font, ws[f"{col}{r}"].border, ws[f"{col}{r}"].alignment = small, border, day_align
    ws.row_dimensions[r].height = 16.5
# 土日・祝日・年末年始の列は灰色。部・課は上の行と同じなら文字を見せない（書き出したシートでは結合する）
ws.conditional_formatting.add(f"D4:{S_DAYS[-1]}{S_LAST}", FormulaRule(
    formula=[f"AND(ISNUMBER(D$4),{off_day('D$4')})"], fill=PatternFill("solid", fgColor="E7E6E6", bgColor="E7E6E6")))
ws.conditional_formatting.add(f"A{S_FIRST + 1}:A{S_LAST}", FormulaRule(formula=[f'AND(A{S_FIRST + 1}<>"",A{S_FIRST + 1}=A{S_FIRST})'], font=Font(color="FFFFFF")))
ws.conditional_formatting.add(f"B{S_FIRST + 1}:B{S_LAST}", FormulaRule(
    formula=[f'AND(B{S_FIRST + 1}<>"",A{S_FIRST + 1}=A{S_FIRST},B{S_FIRST + 1}=B{S_FIRST})'], font=Font(color="FFFFFF")))
ws.freeze_panes = f"D{S_FIRST}"
# 開いたときは改ページプレビュー・85％（印刷の形のまま見る。書き出したシートも同じ）
ws.sheet_view.view = "pageBreakPreview"
ws.sheet_view.zoomScale = 85
ws.sheet_view.zoomScaleSheetLayoutView = 85
ws.sheet_view.zoomScaleNormal = 85
ws.page_setup.paperSize = ws.PAPERSIZE_A4
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_title_rows = "1:5"
ws.print_area = f"A1:{S_DAYS[-1]}{S_LAST}"

# ===== 目次（先頭のシート）：作業の流れごとに、行くシートとボタンと今の状況 =====
# E 列の「[ボタン]マクロ名|表示」は add_buttons.ps1 がボタンにする。F 列の「[式]=…」は表（Power Query の表など）を指す式なので、
# 表ができたあとで add_buttons.ps1 が式にする（openpyxl では表の名前を使った式を書けないため）
from openpyxl.worksheet.hyperlink import Hyperlink

ws = wb.create_sheet("目次", 0)
ws.sheet_view.showGridLines = False
for col, width in zip("ABCDEFG", [2, 6, 44, 22, 26, 24, 2]):
    ws.column_dimensions[col].width = width
ws["B1"] = "出勤・休暇管理データ　目次"
ws["B1"].font = f_title
ws["B2"] = "上から順に作業します。シート名を押すとそのシートへ移ります（Ctrl＋PageUp／PageDown でも移れます）。各シートの見出しの色は、下の作業のまとまりの色です。"
ws["B2"].font = f_note
GROUPS = [   # （まとまり, 見出しの色, [(番号, 作業, シート, 行き先のセル, ボタン, 状況の式)])
    ("最新の内容を読み込む（作業の前に）", "1E3A5F", [
        ("", "申請・報告シート・テレワーク予定表・超過勤務等命令簿・フレックス・テレワーク報告を読み直し、承認済みのフレックスを取り込みます（1分ほどかかります）", "", "", "ImportRequests|最新を読み込む",
         "='申請・報告シートの読み込み'!B5"),
    ]),
    ("毎月の出勤簿", "2E75B6", [
        ("1", "出勤管理システムの「Excel出力」を、出力データ貼付の H3 のフォルダに置く（または A2 に貼り付ける）", "出力データ貼付", "H3", "", ""),
        ("2", "照合結果を確かめる（★名簿に未登録 が出たら職員名簿を直す）", "照合結果", "A1", "", '=COUNTIF(照合結果!C2:C5001,"★名簿に未登録")'),
        ("3", "月次一覧で出勤日数を確かめる（B1 に年月）", "月次一覧", "B1", "", ""),
        ("4", "1人ずつ確かめる・印刷する（AF2 に年月、AF3 に出勤用ID）", "出勤簿", "AF2", "", ""),
        ("5", f"全員分を印刷する（AF2 に年月、{BATCH}人まで）", "一括印刷", "AF2", "", ""),
    ]),
    ("休暇", "548235", [
        ("1", "申請の状況を確かめる（承認待ち・注意）", "申請一覧", "A1", "", "[式]=COUNTIF(申請一覧[状況],\"承認待ち\")"),
        ("2", "年次休暇の残り・本年分・本年使用（手入力）を確かめる（年が変わったら、名簿の人の行を足す）", "休暇等", "A1", "AddLeaveRows|名簿の人の行を足す", ""),
        ("3", "申請・報告シートのフォルダの場所（最初に一度）", "申請・報告シートの読み込み", "B2", "", ""),
    ]),
    ("テレワーク", "00827F", [
        ("1", "テレワークの予定と注意を確かめる", "テレワーク一覧", "A5", "", "[式]=COUNTIF(テレワーク一覧[注意],\"?*\")"),
        ("", "テレワーク実施報告と所属長の確認を確かめる", "テレワーク報告", "A3", "", "[式]=COUNTIF(テレワーク報告一覧[状況],\"確認待ち\")+COUNTIF(テレワーク報告一覧[状況],\"予定なし\")"),
    ]),
    ("超過勤務", "C55A11", [
        ("1", "超過勤務等命令簿の所属長の確認の状況を見る", "超過勤務の読み込み", "A3", "", "[式]=COUNTIF(超勤_月[状況],\"確認待ち\")"),
        ("2", "月を B2 に入れて、確認済みの月を写す（月1回）", "超過勤務", "B2", "TransferOvertime|超過勤務を転記", ""),
        ("3", "写した日ごとの行の注意を確かめる", "超過勤務の日別", "A3", "", "[式]=COUNTIF(転記_日別[注意],\"?*\")"),
        ("4", "確認済みの命令簿を PDF で保管する（B2 に保存先のフォルダ）", "命令簿の保存", "B2", "SaveOvertimePdfs|命令簿をPDFで保存", ""),
    ]),
    ("フレックス", "7030A0", [
        ("1", "フレックスの単位期間の承認の状況と注意を確かめる", "フレックス", "A3", "", "[式]=COUNTIF(フレックス一覧[注意],\"?*\")"),
        ("2", "取り込んだフレックスの照合の注意（出勤の記録がない勤務日など）を確かめる", "フレックス", "A3", "", "[式]=COUNTIF(フレックス一覧[照合の注意],\"?*\")"),
        ("3", "フレックス／テレワーク利用予定を作る（AN1 に月、AN9 に書き出す先のファイル）", "利用予定", "AN1", "MakeScheduleSheet|利用予定を書き出す", ""),
    ]),
    ("台帳（変わったときに）", "7F7F7F", [
        ("1", "職員名簿（出勤用ID・氏名・休憩の特例・部課・所属長など）", "職員名簿", "A2", "", '=COUNTIF(計算用!B2:B501,"重複")'),
        ("2", "祝日の一覧", "祝日", "A2", "", ""),
        ("", "このブックの使い方（詳しい説明）", "使い方", "A1", "", ""),
    ]),
]
STATUS_LABEL = {"申請・報告シートの読み込み'!B5": "前回の読み込み", "★名簿に未登録": "名簿に未登録 {} 件",
                "承認待ち": "承認待ち {} 件", "テレワーク一覧[注意]": "注意 {} 件", "テレワーク報告一覧[状況]": "確認待ち・予定なし {} 件", "確認待ち": "確認待ち {} 人・月",
                "転記_日別[注意]": "注意 {} 日", "フレックス一覧[注意]": "注意 {} 件", "フレックス一覧[照合の注意]": "照合の注意 {} 件",
                "重複": "ID の重複 {} 件"}
TAB_COLORS = {}
r = 4
for title, color, items in GROUPS:
    c = ws.cell(row=r, column=2, value=title)
    c.font = Font(name=FONT, size=11, bold=True, color="FFFFFF")
    for col in range(2, 7):
        ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=color)
    ws.row_dimensions[r].height = 20
    r += 1
    for no, task, sheet, cell, button, status in items:
        ws.row_dimensions[r].height = max(30, 15 * -(-len(task) // 21))     # 長い説明は折り返して入る高さに
        ws.cell(row=r, column=2, value=no).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=r, column=3, value=task).alignment = Alignment(vertical="center", wrap_text=True)
        if sheet:
            TAB_COLORS[sheet] = color
            link = ws.cell(row=r, column=4, value=sheet)
            link.hyperlink = Hyperlink(ref=link.coordinate, location=f"'{sheet}'!{cell}", display=sheet)
            link.font = Font(name=FONT, size=10, color="0563C1", underline="single")
            link.alignment = Alignment(vertical="center")
        if button:
            ws.cell(row=r, column=5, value="[ボタン]" + button)
        if status:
            s = ws.cell(row=r, column=6)
            # 状況の表示：件数のときは「承認待ち 3 件」のように。0 件なら空
            label = next((v for k, v in STATUS_LABEL.items() if k in status), "{}")
            expr = status.split("=", 1)[1] if status.startswith("[式]") else status[1:]
            if "{}" in label and label != "{}":
                pre, post = label.split("{}")
                text = f'=IF(IFERROR({expr},0)=0,"",\"{pre}\"&{expr}&\"{post}\")'
            else:
                text = f'=IF({expr}="","まだ読み込んでいません","{label}："&TEXT({expr},"m/d h:mm"))'
            s.value = ("[式]" if status.startswith("[式]") else "") + text
            s.font = Font(name=FONT, size=10, bold=True, color="C00000")
            s.alignment = Alignment(vertical="center", wrap_text=True)
        ws.cell(row=r, column=2).font = ws.cell(row=r, column=3).font = f_base
        for col in range(2, 7):
            ws.cell(row=r, column=col).border = Border(bottom=Side(style="hair", color="BFBFBF"))
        r += 1
    r += 1
ws.cell(row=r, column=2, value="年に1回：休暇等シートの［名簿の人の行を足す］で新しい年の行を足し、申請・報告シートの作成.xlsm の B6 を新しい年にして［申請・報告シートを更新］を押します。"
        ).font = f_note
ws.cell(row=r + 1, column=2, value="月に1回：申請・報告シートの作成.xlsm の B7 を翌月にして［超過勤務の月を追加］を押します（全員の申請・報告シートに超過勤務等命令簿の月のシートを作る）。"
        ).font = f_note
ws.cell(row=r + 2, column=2, value="このファイルを開いただけでは読み込みません。作業の前に［最新を読み込む］を押してください（「申請・報告シートの読み込み」シートの B5 に前回の日時）。"
        ).font = f_note

# シートの並びと見出しの色：目次 → 毎月の出勤簿 → 休暇 → テレワーク → 超過勤務 → フレックス → 台帳 → 使い方（計算用は非表示）
ORDER = ["目次", "出力データ貼付", "照合結果", "月次一覧", "出勤簿", "一括印刷",
         "申請一覧", "休暇等", "申請・報告シートの読み込み",
         "テレワーク一覧", "テレワーク報告", "超過勤務の読み込み", "超過勤務", "超過勤務の日別", "命令簿の保存",
         "フレックス", "フレックスの日別", "フレックスの読み込み", "利用予定",
         "職員名簿", "祝日", "使い方", "計算用", "QR出勤の読み込み"]
assert sorted(ORDER) == sorted(s.title for s in wb.worksheets), [s.title for s in wb.worksheets]
wb._sheets = [wb[name] for name in ORDER]
wb["目次"].sheet_properties.tabColor = "1E3A5F"
TAB_COLORS.setdefault("フレックスの読み込み", TAB_COLORS["フレックス"])     # 目次にないフレックスのシートも同じ色
TAB_COLORS.setdefault("フレックスの日別", TAB_COLORS["フレックス"])
TAB_COLORS.setdefault("QR出勤の読み込み", TAB_COLORS["出力データ貼付"])
for name, color in TAB_COLORS.items():
    wb[name].sheet_properties.tabColor = color

wb.calculation.calcMode = "auto"
wb.calculation.fullCalcOnLoad = True
wb.active = 0
wb.save(OUT)
print("saved", OUT)
