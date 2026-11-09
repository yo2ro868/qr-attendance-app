"""休暇申請シート（職員1人に1ファイル）のひな形を作る。

使い方: python make_form.py 出力先.xlsx
管理者がこのひな形を職員ごとにコピーし（ファイル名は 休暇申請シート　出勤用ID　氏名.xlsx。マクロ 休暇申請シートの作成.xlsm が作る）、
管理者の OneDrive から本人（編集）と所属長（編集）にだけ共有して使う。
出勤用ID・氏名はファイル名から読む。年・合計・本年使用の調整は、出勤・休暇管理データ の休暇等シートで一括管理し、
マクロが非表示の「管理」シートに書き込む（本人・所属長には入力欄を見せない）。日数は開始日・終了日から自動で計算する。
「申請」の表の名前と見出しは、出勤・休暇管理データ・休暇集計の Power Query が参照するので変えないこと。
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
DAY_MIN = 465         # 1日＝7時間45分＝465分
CARRY_MAX = 20        # 翌年に繰り越せる日数の上限
ROWS = 300            # 申請を書ける行数（保護したシートでは表に行を足せないので、先に用意しておく）
HEAD = 6              # 申請の表の見出しの行（入力欄が画面の上のほうに来るように詰める）
FIRST, LAST = HEAD + 1, HEAD + ROWS

KINDS = ["年次休暇", "病気休暇", "看護休暇（子の看護等）", "特別休暇（看護休暇以外）",
         "介護休暇", "介護時間", "育児時間", "職務専念義務の免除"]
SPECIAL = ["夏季休暇", "結婚", "忌引（親族の死亡）", "父母の追悼", "妻の出産", "育児参加", "短期介護",
           "不妊治療", "産前", "産後", "保育時間", "ボランティア", "骨髄等の提供", "選挙権の行使等",
           "官公署への出頭", "現住居の滅失等", "その他"]
EXEMPT = ["健康診断", "研修"]     # 職務専念義務の免除の内容の例（一覧にないものも書ける）
KIND_SPECIAL, KIND_EXEMPT = "特別休暇（看護休暇以外）", "職務専念義務の免除"
KIND_ANNUAL = "年次休暇"      # 事由は書かない
NEED_APPROVAL_DATE = "・承認日を入れてください"
DUPLICATE = "・承認済みの申請と日時が重なっています（このままでは承認できません）"
# 事由の定型文（KINDS の順。一覧にないものも自由に書ける）
REASONS = {
    KIND_ANNUAL: [],
    "病気休暇": ["通院", "発熱・体調不良", "負傷", "入院"],
    "看護休暇（子の看護等）": ["子の通院", "子の発熱・体調不良", "子の予防接種・健康診断", "学級閉鎖等", "入学式・卒業式等"],
    KIND_SPECIAL: [],
    "介護休暇": ["家族の介護", "家族の通院の付き添い"],
    "介護時間": ["家族の介護", "家族の通院の付き添い"],
    "育児時間": ["保育所等への送迎", "子の世話"],
    KIND_EXEMPT: [],
}
assert list(REASONS) == KINDS
UNITS = ["日単位", "時間・分単位"]
# 単位が1つに決まる休暇（休暇の種類か種別の文字）。選ぶと単位の欄に自動で入り、一覧もそれだけになる。ほかの休暇は両方から選ぶ
UNIT_TIME_ONLY = ["介護時間", "育児時間"]     # 時間・分単位だけ（休暇制度の資料による）
UNIT_DAY_ONLY = []                            # 日単位だけ（該当なし。必要になったら足す）
APPROVALS = ["承認", "却下", "取消"]
# 様式の版。申請シートの式や列を変えたら変える（マクロが、ひな形と版の違う休暇申請シートを結果シートで知らせる）
FORM_VERSION = "2026-11-06d"

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
    ("日数", 6, "0", fill_calc),     # 開始日・終了日から自動で計算（土日・祝日・年末年始を除く）
    ("開始時刻", 8, "h:mm", fill_staff),
    ("終了時刻", 8, "h:mm", fill_staff),
    ("時間", 7, "[h]:mm", fill_calc),   # 開始時刻・終了時刻から休憩を除いて自動で計算（打ち込めない）
    ("事由", 22, None, fill_staff),
    ("承認（所属長）", 9, None, fill_boss),
    ("承認日", 11, "yyyy/m/d", fill_boss),
    ("所属長コメント", 18, None, fill_boss),
    ("確認", 30, None, fill_calc),
    ("分", 6, "0", fill_calc),     # 取得時間（分）。上の残日数の計算に使う（非表示）
    ("年", 6, "0", fill_calc),     # 開始日の年（非表示）
    ("休日可", 6, "0", fill_calc),  # 休暇の種類・種別が「土日祝でも可」の一覧にあれば 1 以上（非表示）
    ("単位の制限", 6, None, fill_calc),  # 単位が決まる休暇なら「日」か「時間」（非表示）
    ("事由の一覧", 6, "0", fill_calc),  # 休暇の種類の番号。事由の定型文の一覧を選ぶ（非表示）
    ("時刻", 13, None, fill_calc),      # 開始時刻〜終了時刻の文字。出勤・休暇管理データ が読む（非表示）
    ("重複", 6, "0", fill_calc),        # この行と日時が重なる、ほかの承認済みの行の数（非表示）
    ("期間の終わり", 6, "yyyy/m/d", fill_calc),  # 終了日（なければ開始日）（非表示）
    ("その日の始め", 6, "h:mm", fill_calc),      # 時刻（全日なら 0）（非表示）
    ("その日の終わり", 6, "h:mm", fill_calc),    # 時刻（全日なら 1）（非表示）
    ("全日", 6, "0", fill_calc),        # 時刻の指定がなければ 1（非表示）
]
for i, (_, width, _, _) in enumerate(COLS, start=1):
    ws.column_dimensions[get_column_letter(i)].width = width

# ----- 1行目：題名 -----
ws["A1"] = "休暇申請シート"
ws["A1"].font = f_title
ws["C1"] = "黄色：本人が入力　　青：所属長が入力　　灰色：自動　　赤：入力が足りない・誤りのある欄（理由は「確認」の欄）"
ws["C1"].font = f_note

# ----- 管理（非表示）：マクロ（休暇申請シートの作成.xlsm）が出勤・休暇管理データ の休暇等シートから書き込む -----
MG = "'管理'"
YR, HOLD, ADJ = f"{MG}!$B$3", f"{MG}!$B$4", f"{MG}!$B$5"   # 年、合計（分）、本年使用の調整（分）
BREAK_START, BREAK_END = f"{MG}!$B$9", f"{MG}!$B$10"        # 休憩の始め・終わり（短縮の特例の人はマクロが書き換える）
CARRY, GRANT = f"{MG}!$B$7", f"{MG}!$B$8"                   # 繰越分（前年から引き継ぐ未使用分）、本年分（新たに付与される分）（分）
mg = wb.create_sheet("管理")
for i, label in enumerate(["出勤用ID", "氏名", "年", "合計（分）", "本年使用の調整と Forms の承認済み（分）", "更新日時", "繰越分（分）", "本年分（分）", "休憩の始め", "休憩の終わり"], start=1):
    mg[f"A{i}"] = label
mg["B1"].number_format = "@"
mg["A11"], mg["B11"] = "様式の版", FORM_VERSION
mg["B11"].number_format = "@"
mg["B9"], mg["B10"] = 12 / 24, 13 / 24     # 12:00〜13:00
mg["B9"].number_format = mg["B10"].number_format = "h:mm"
# 承認の一覧：W列（日時が重なる行では「承認」を出さない）
for i, name in enumerate(APPROVALS, start=1):
    mg[f"W{i}"] = name
mg["C1"] = "このシートは マクロ（休暇申請シートの作成.xlsm）が書き込みます。手で変えないでください。"
# 日数の計算で除く日：D列 祝日（マクロが出勤・休暇管理データ の祝日シートで書き換える）、E列 年末年始（12/29〜1/3）
for i, (d, _) in enumerate(HOLIDAYS, start=1):
    mg[f"D{i}"] = date.fromisoformat(d)
year_end = [date(y, 12, d) for y in range(2025, 2041) for d in (29, 30, 31)] + [date(y, 1, d) for y in range(2026, 2042) for d in (1, 2, 3)]
for i, d in enumerate(year_end, start=1):
    mg[f"E{i}"] = d
OFF_DAYS = f"{MG}!$D$1:$E$200"
# 開始日・終了日に土日・祝日・年末年始を指定してよい休暇（休暇の種類か種別の文字）。F列に1行1つ。今は例外なし
OFF_DAY_OK = []
for i, name in enumerate(OFF_DAY_OK, start=1):
    mg[f"F{i}"] = name
OFF_DAY_LIST = f"{MG}!$F$1:$F$50"
# 単位の条件：G列 日単位だけの休暇、H列 時間・分単位だけの休暇、I列 単位の選択肢
for col, names in (("G", UNIT_DAY_ONLY), ("H", UNIT_TIME_ONLY), ("I", UNITS)):
    for i, name in enumerate(names, start=1):
        mg[f"{col}{i}"] = name
DAY_ONLY, TIME_ONLY = f"{MG}!$G$1:$G$50", f"{MG}!$H$1:$H$50"
# 種別・内容の一覧：J列 特別休暇の種別、K列 職務専念義務の免除の内容の例、L列 空（ほかの休暇は一覧なし）
for col, names in (("J", SPECIAL), ("K", EXEMPT)):
    for i, name in enumerate(names, start=1):
        mg[f"{col}{i}"] = name
SPECIAL_RANGE = f"{MG}!$J$1:$J${len(SPECIAL)}"
EXEMPT_RANGE = f"{MG}!$K$1:$K${len(EXEMPT)}"
# 事由の定型文：M列 休暇の種類、N列から 種類ごとの定型文（年次休暇は空）
for i, kind in enumerate(KINDS, start=1):
    mg[f"M{i}"] = kind
    for j, text in enumerate(REASONS[kind], start=1):
        mg[f"{get_column_letter(13 + i)}{j}"] = text
KIND_LIST = f"{MG}!$M$1:$M${len(KINDS)}"
REASON_RANGES = ",".join(
    f"{MG}!${get_column_letter(13 + i)}$1:${get_column_letter(13 + i)}${max(1, len(REASONS[kind]))}"
    for i, kind in enumerate(KINDS, start=1))
# 時刻の一覧：V列 8:00〜19:00 の15分おき（一覧にない時刻も書ける）
TIMES = [(8 * 60 + 15 * i) / 1440 for i in range(45)]
for i, v in enumerate(TIMES, start=1):
    mg[f"V{i}"] = v
    mg[f"V{i}"].number_format = "h:mm"
TIME_RANGE = f"{MG}!$V$1:$V${len(TIMES)}"
mg.sheet_state = "veryHidden"     # 本人・所属長からは見えない（Excel の［再表示］にも出ない）

# ----- 2〜3行目：氏名と年次休暇の残りなど（自動）。4行目は分の計算用で非表示 -----
rng = lambda c: f"${c}${FIRST}:${c}${LAST}"
KIND, MIN, YEAR, OK = rng("B"), rng("P"), rng("Q"), rng("L")
TOP = [
    ("氏名", "A", "A", None, f'{MG}!$B$2&""'),
    ("対象", "B", "B", None, f'IF({YR}="","（管理者が設定すると表示されます）",{YR}&"年の年次休暇")'),
    ("繰越分", "C", "C", f"=N({CARRY})", days_text("C4")),
    ("本年分", "D", "D", f"=N({GRANT})", days_text("D4")),
    ("合計", "E", "E", f"=N({HOLD})", days_text("E4")),
    # 本年使用＝ 出勤・休暇管理データ の調整（導入前の使用など）と Forms で承認された分 ＋ この用紙の承認済みの年次休暇
    ("本年使用（承認済み）", "F", "F",
     f'=N({ADJ})+SUMIFS({MIN},{KIND},"年次休暇",{YEAR},{YR},{OK},"承認")', days_text("F4")),
    ("承認待ち", "G", "H", f'=SUMIFS({MIN},{KIND},"年次休暇",{YEAR},{YR},{OK},"")', days_text("G4")),
    ("残り", "I", "J", "=E4-F4", days_text("I4")),
    # 翌年繰越：残りの1日未満を切り捨て、上限20日（出勤・休暇管理データ の計算用シートと同じ）
    ("翌年繰越", "K", "K", f"=MIN({CARRY_MAX * DAY_MIN},INT(MAX(0,I4)/{DAY_MIN})*{DAY_MIN})", days_text("K4")),
    ("病気休暇（承認済み）", "L", "M", f'=SUMIFS({MIN},{KIND},"病気休暇",{YEAR},{YR},{OK},"承認")', hm_text("L4")),
    ("看護休暇（承認済み）", "N", "N", f'=SUMIFS({MIN},{KIND},"看護休暇（子の看護等）",{YEAR},{YR},{OK},"承認")', hm_text("N4")),
]
for label, c1, c2, minutes, text in TOP:
    if c1 != c2:
        ws.merge_cells(f"{c1}2:{c2}2")
        ws.merge_cells(f"{c1}3:{c2}3")
    ws[f"{c1}2"] = label
    ws[f"{c1}3"] = "=" + text
    for col in {c1, c2}:
        ws[f"{col}2"].border = ws[f"{col}3"].border = border
        ws[f"{col}3"].fill = fill_calc
    ws[f"{c1}2"].font, ws[f"{c1}2"].alignment = f_bold, center
    ws[f"{c1}3"].font, ws[f"{c1}3"].alignment = f_bold, center
    if minutes:
        ws[f"{c1}4"] = minutes
ws.row_dimensions[2].height = 28
ws.row_dimensions[3].height = 30     # 「24日5時間45分」が折り返しても切れないように
ws.row_dimensions[4].hidden = True
ws["O2"] = "繰越分：前年の未使用分　本年分：新たに付与　翌年繰越：残りの1日未満切り捨て（上限20日）"
ws["O3"] = "1日＝7時間45分。承認待ちは承認欄が空欄の申請です。"
for ref in ("O2", "O3"):
    ws[ref].font = f_note
    ws[ref].alignment = Alignment(vertical="center")

# ----- 6行目から：申請の表 -----
ws["A5"] = ("申請は1件1行で上から順に。種別・内容、事由、時刻は一覧から選びます（一覧にないものも書けます。特別休暇の種別と年次休暇の事由は除く）。"
            "時間（1日あたり）は開始時刻・終了時刻から休憩を除いて自動で入ります。種別・内容を書いた行の事由は不要です。")
ws["A5"].font = f_note
for i, (label, _, fmt, fill) in enumerate(COLS, start=1):
    h = ws.cell(row=HEAD, column=i, value=label)
    h.font, h.alignment, h.border, h.fill = f_head, center, border, fill_head
    for r in range(FIRST, LAST + 1):
        c = ws.cell(row=r, column=i)
        c.font, c.border, c.fill = f_base, border, fill
        c.alignment = left if label in ("種別・内容", "事由", "所属長コメント", "確認") else center
        if fmt:
            c.number_format = fmt
ws.row_dimensions[HEAD].height = 28
for r in range(FIRST, LAST + 1):
    a, b, c_, d, e, f, g, st, en, h, j = (f"{c}{r}" for c in "ABCDEFGHIJK")
    hmin = f"ROUND(N({h})*1440,0)"
    # 取得時間（分）：時間・分単位は 1日あたりの時間×日数（日数が空欄なら1日）、日単位は 日数×7時間45分
    ws[f"P{r}"] = (f'=IF(OR({b}="",{e}=""),"",IF({d}="時間・分単位",{hmin}*MAX(1,N({g})),'
                   f'ROUND(N({g})*{DAY_MIN},0)+{hmin}))')
    ws[f"Q{r}"] = f'=IF({e}="","",YEAR({e}))'
    # 空欄で COUNTIF すると空のセルを数えてしまうので、書いてあるときだけ数える
    ws[f"R{r}"] = f'=IF({b}="",0,COUNTIF({OFF_DAY_LIST},{b}))+IF({c_}="",0,COUNTIF({OFF_DAY_LIST},{c_}))'
    ws[f"S{r}"] = (f'=IF({b}="","",IF(COUNTIF({DAY_ONLY},{b})+IF({c_}="",0,COUNTIF({DAY_ONLY},{c_}))>0,"日",'
                   f'IF(COUNTIF({TIME_ONLY},{b})+IF({c_}="",0,COUNTIF({TIME_ONLY},{c_}))>0,"時間","")))')
    ws[f"T{r}"] = f'=IF({b}="",1,IFERROR(MATCH({b},{KIND_LIST},0),1))'
    # 時間：時間・分単位で開始時刻・終了時刻があれば、その差から休憩（重なる分だけ）を引いて入る
    ws[h] = (f'=IF(AND({d}="時間・分単位",ISNUMBER({st}),ISNUMBER({en})),IF({en}>{st},'
             f'ROUND(({en}-{st}-MAX(0,MIN({en},{BREAK_END})-MAX({st},{BREAK_START})))*1440,0)/1440,""),"")')
    # 重複：開始日〜終了日と、その日の時刻（全日なら 0〜1）の両方が重なる、ほかの承認済みの行の数（却下・取消の行は見ない）
    ws[f"Z{r}"] = f'=IF({e}="",0,IF(AND({d}="時間・分単位",ISNUMBER({st}),ISNUMBER({en})),0,1))'
    ws[f"W{r}"] = f'=IF({e}="","",IF({f}="",{e},{f}))'
    ws[f"X{r}"] = f'=IF({e}="","",IF(Z{r}=1,0,{st}))'
    ws[f"Y{r}"] = f'=IF({e}="","",IF(Z{r}=1,1,{en}))'
    ws[f"V{r}"] = (f'=IF(OR({e}="",L{r}="却下",L{r}="取消"),0,SUMPRODUCT(($L${FIRST}:$L${LAST}="承認")*ISNUMBER($E${FIRST}:$E${LAST})'
                   f'*($E${FIRST}:$E${LAST}<=W{r})*($W${FIRST}:$W${LAST}>={e})'
                   f'*($X${FIRST}:$X${LAST}<Y{r})*($Y${FIRST}:$Y${LAST}>X{r}))-IF(L{r}="承認",1,0))')
    ws[f"U{r}"] = f'=IF(OR({st}="",{en}=""),"",TEXT({st},"h:mm")&"〜"&TEXT({en},"h:mm"))'
    # 単位：決まる休暇は自動で入る。両方取れる休暇は空欄で、本人が一覧から選ぶ（この式の上に書く）
    ws[d] = f'=IF(S{r}="日","{UNITS[0]}",IF(S{r}="時間","{UNITS[1]}",""))'
    # 日数：開始日だけなら 1、終了日もあれば期間の勤務日の数（終了日が前のときは空欄）
    ws[f"G{r}"] = f'=IF({e}="","",IF({f}="",1,IF({f}<{e},"",NETWORKDAYS({e},{f},{OFF_DAYS}))))'
    ws[f"O{r}"] = (
        f'=IF(COUNTA($A{r}:$C{r},$E{r}:$F{r},$H{r}:$I{r},$K{r})=0,"",IF({b}="","・休暇の種類を選んでください",'
        f'IF({a}="","・申請日を入れてください","")'
        f'&IF({d}="","・単位を選んでください","")'
        f'&IF(AND({b}<>"{KIND_ANNUAL}",{c_}="",{j}=""),"・事由を入れてください","")'
        f'&IF(V{r}>0,"{DUPLICATE}","")'
        f'&IF(AND({b}="{KIND_ANNUAL}",{j}<>""),"・年次休暇に事由は要りません","")'
        f'&IF(AND(L{r}="承認",M{r}=""),"{NEED_APPROVAL_DATE}","")'
        f'&IF(AND(S{r}="日",{d}="時間・分単位"),"・この休暇は日単位だけです","")'
        f'&IF(AND(S{r}="時間",{d}="日単位"),"・この休暇は時間・分単位だけです","")'
        f'&IF(AND(OR({b}="{KIND_SPECIAL}",{b}="{KIND_EXEMPT}"),{c_}=""),"・種別・内容を入れてください","")'
        f'&IF(AND({b}="{KIND_SPECIAL}",{c_}<>""),IF(COUNTIF({SPECIAL_RANGE},{c_})=0,"・特別休暇の種別は一覧から選んでください",""),"")'
        f'&IF(AND({b}<>"{KIND_SPECIAL}",{b}<>"{KIND_EXEMPT}",{c_}<>""),"・この休暇に種別・内容は要りません","")'
        f'&IF({e}="","・開始日を入れてください","")'
        f'&IF(AND(ISNUMBER({e}),R{r}=0),IF(NETWORKDAYS({e},{e},{OFF_DAYS})=0,"・開始日が土日・祝日・年末年始です",""),"")'
        f'&IF(AND(ISNUMBER({f}),R{r}=0),IF(NETWORKDAYS({f},{f},{OFF_DAYS})=0,"・終了日が土日・祝日・年末年始です",""),"")'
        f'&IF(AND({e}<>"",{f}<>"",{f}>={e},N({g})<=0),"・期間に勤務日がありません","")'
        f'&IF(AND({d}="時間・分単位",OR({st}="",{en}="")),"・開始時刻と終了時刻を入れてください","")'
        f'&IF(AND({d}="時間・分単位",{st}<>"",{en}<>"",N({h})<=0),IF({en}>{st},"・時刻が休憩時間の中です",""),"")'
        f'&IF(N({h})>{DAY_MIN}/1440,"・時間が7時間45分を超えています（時刻を確かめてください）","")'
        f'&IF(AND({st}<>"",{en}<>""),IF({en}<={st},"・終了時刻が開始時刻より前です",""),"")'
        f'&IF(AND({d}="日単位",OR({st}<>"",{en}<>"")),"・日単位に時刻は要りません","")'
        f'&IF(AND({f}<>"",{e}<>""),IF({f}<{e},"・終了日が開始日より前です",""),"")'
        f'&IF(AND({b}="年次休暇",N({h})>0),IF(MOD({hmin},15)<>0,"・年次休暇の時間は15分単位です",""),"")'
        f'&IF(AND(OR({b}="介護時間",{b}="育児時間"),N({h})>0),IF(MOD({hmin},30)<>0,"・30分単位です",""),"")'
        f'&IF(AND({f}<>"",{e}<>"",OR({b}="年次休暇",{b}="病気休暇",{b}="看護休暇（子の看護等）")),'
        f'IF(TEXT({e},"yyyymm")<>TEXT({f},"yyyymm"),"・月ごとに分けて書いてください",""),"")))'
    )
# 入力が足りない・誤りのある欄を赤くする（Excel Online でも動くよう、マクロではなく条件付き書式で）。
# 理由は「確認」の欄（O 列）に文章で出る。条件は O 列の式とそろえること
FIRST_ROW = FIRST
month_kinds = f'OR($B{FIRST_ROW}="年次休暇",$B{FIRST_ROW}="病気休暇",$B{FIRST_ROW}="看護休暇（子の看護等）")'
hmin1 = f"ROUND(N($J{FIRST_ROW})*1440,0)"
fill_bad = PatternFill("solid", fgColor="F4B6B0", bgColor="F4B6B0")
font_bad = Font(name=FONT, color="9C0006", bold=True)
RED_RULES = {
    "A": f'AND($B{FIRST_ROW}<>"",$A{FIRST_ROW}="")',
    "B": f'AND(COUNTA($A{FIRST_ROW}:$C{FIRST_ROW},$E{FIRST_ROW}:$F{FIRST_ROW},$H{FIRST_ROW}:$I{FIRST_ROW},$K{FIRST_ROW})>0,$B{FIRST_ROW}="")',
    "C": (f'OR(AND(OR($B{FIRST_ROW}="{KIND_SPECIAL}",$B{FIRST_ROW}="{KIND_EXEMPT}"),$C{FIRST_ROW}=""),'
          f'AND($B{FIRST_ROW}="{KIND_SPECIAL}",$C{FIRST_ROW}<>"",COUNTIF({SPECIAL_RANGE},$C{FIRST_ROW})=0),'
          f'AND($B{FIRST_ROW}<>"",$B{FIRST_ROW}<>"{KIND_SPECIAL}",$B{FIRST_ROW}<>"{KIND_EXEMPT}",$C{FIRST_ROW}<>""))'),
    "D": (f'OR(AND($B{FIRST_ROW}<>"",$D{FIRST_ROW}=""),AND($S{FIRST_ROW}="日",$D{FIRST_ROW}="時間・分単位"),'
          f'AND($S{FIRST_ROW}="時間",$D{FIRST_ROW}="日単位"))'),
    "E": (f'OR($V{FIRST_ROW}>0,AND($B{FIRST_ROW}<>"",$E{FIRST_ROW}=""),'
          f'IF(AND(ISNUMBER($E{FIRST_ROW}),$R{FIRST_ROW}=0),NETWORKDAYS($E{FIRST_ROW},$E{FIRST_ROW},{OFF_DAYS})=0,FALSE))'),
    "F": (f'OR(AND($F{FIRST_ROW}<>"",$E{FIRST_ROW}<>"",OR($F{FIRST_ROW}<$E{FIRST_ROW},'
          f'AND({month_kinds},TEXT($E{FIRST_ROW},"yyyymm")<>TEXT($F{FIRST_ROW},"yyyymm")))),'
          f'IF(AND(ISNUMBER($F{FIRST_ROW}),$R{FIRST_ROW}=0),NETWORKDAYS($F{FIRST_ROW},$F{FIRST_ROW},{OFF_DAYS})=0,FALSE))'),
    "G": f'AND($E{FIRST_ROW}<>"",$F{FIRST_ROW}<>"",$F{FIRST_ROW}>=$E{FIRST_ROW},N($G{FIRST_ROW})<=0)',
    "K": (f'OR(AND($B{FIRST_ROW}<>"",$B{FIRST_ROW}<>"{KIND_ANNUAL}",$C{FIRST_ROW}="",$K{FIRST_ROW}=""),'
          f'AND($B{FIRST_ROW}="{KIND_ANNUAL}",$K{FIRST_ROW}<>""))'),
    # H 開始時刻、I 終了時刻、J 時間（自動）
    "H": f'OR(AND($D{FIRST_ROW}="時間・分単位",$H{FIRST_ROW}=""),AND($D{FIRST_ROW}="日単位",$H{FIRST_ROW}<>""))',
    "I": (f'OR(AND($D{FIRST_ROW}="時間・分単位",$I{FIRST_ROW}=""),AND($D{FIRST_ROW}="日単位",$I{FIRST_ROW}<>""),'
          f'AND($H{FIRST_ROW}<>"",$I{FIRST_ROW}<>"",$I{FIRST_ROW}<=$H{FIRST_ROW}))'),
    "J": (f'OR(AND($D{FIRST_ROW}="時間・分単位",$H{FIRST_ROW}<>"",$I{FIRST_ROW}>$H{FIRST_ROW},N($J{FIRST_ROW})<=0),N($J{FIRST_ROW})>{DAY_MIN}/1440,'
          f'AND($B{FIRST_ROW}="年次休暇",N($J{FIRST_ROW})>0,MOD({hmin1},15)<>0),'
          f'AND(OR($B{FIRST_ROW}="介護時間",$B{FIRST_ROW}="育児時間"),N($J{FIRST_ROW})>0,MOD({hmin1},30)<>0))'),
    # 所属長が「承認」にしたのに、まだ不備が残っている行
    # （承認日がないことは承認日の欄を赤くするので、それ以外の不備で見る）
    "L": f'AND($L{FIRST_ROW}="承認",SUBSTITUTE($O{FIRST_ROW},"{NEED_APPROVAL_DATE}","")<>"")',
    "M": f'AND($L{FIRST_ROW}="承認",$M{FIRST_ROW}="")',
    "O": f'$O{FIRST_ROW}<>""',
}
for col, rule in RED_RULES.items():
    ws.conditional_formatting.add(f"{col}{FIRST}:{col}{LAST}", FormulaRule(formula=[rule], fill=fill_bad, font=font_bad))

# 年次休暇と、種別・内容を書いた行の事由は要らないので、欄を灰色にする
ws.conditional_formatting.add(f"K{FIRST}:K{LAST}", FormulaRule(
    formula=[f'AND(OR($B{FIRST_ROW}="{KIND_ANNUAL}",$C{FIRST_ROW}<>""),$K{FIRST_ROW}="")'], fill=PatternFill("solid", fgColor="D9D9D9", bgColor="D9D9D9")))

t = Table(displayName="申請", ref=f"A{HEAD}:{get_column_letter(len(COLS))}{LAST}")
t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
ws.add_table(t)
for col in "PQRSTUVWXYZ":
    ws.column_dimensions[col].hidden = True
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
# 種別・内容の一覧：休暇の種類で絞る。免除の内容は一覧にないものも書けるよう、ここでは止めない（誤りは赤くする）
dv = DataValidation(type="list", allow_blank=True,
                    formula1=f'IF($B{FIRST}="{KIND_SPECIAL}",{SPECIAL_RANGE},IF($B{FIRST}="{KIND_EXEMPT}",{EXEMPT_RANGE},{MG}!$L$1:$L$1))')
dv.showErrorMessage = False
ws.add_data_validation(dv)
dv.add(col_cells("C"))
# 事由の一覧：休暇の種類ごとの定型文。一覧にないものも自由に書ける
reason_formula = f"CHOOSE($T{FIRST},{REASON_RANGES})"
assert len(reason_formula) <= 255, len(reason_formula)
dv = DataValidation(type="list", allow_blank=True, formula1=reason_formula)
dv.showErrorMessage = False
ws.add_data_validation(dv)
dv.add(col_cells("K"))
# 時刻の一覧：15分おき。一覧にない時刻（分単位の病気休暇など）も書ける
dv = DataValidation(type="list", allow_blank=True, imeMode="halfAlpha", formula1=TIME_RANGE)
dv.showErrorMessage = False
ws.add_data_validation(dv)
dv.add(col_cells("H"))
dv.add(col_cells("I"))
# 単位の一覧：休暇の種類・種別によって絞る
dv = DataValidation(type="list", allow_blank=True,
                    formula1=f'IF($S{FIRST}="日",{MG}!$I$1:$I$1,IF($S{FIRST}="時間",{MG}!$I$2:$I$2,{MG}!$I$1:$I$2))')
dv.errorTitle, dv.error = "単位", "この休暇で選べる単位を一覧から選んでください。"
ws.add_data_validation(dv)
dv.add(col_cells("D"))
# 承認の一覧：承認済みの申請と日時が重なる行では「承認」を選べない
dv = DataValidation(type="list", allow_blank=True,
                    formula1=f"IF($V{FIRST}>0,{MG}!$W$2:$W$3,{MG}!$W$1:$W$3)")
dv.errorTitle, dv.error = "承認", "承認済みの申請と日時が重なっている行は承認できません。"
ws.add_data_validation(dv)
dv.add(col_cells("L"))
# 開始日・終了日は勤務日だけ（土日・祝日・年末年始は入力できない。ただし選んだ休暇が「土日祝でも可」の一覧にあれば入力できる）。
# 貼り付けですり抜けたときは赤くなる
for col, label in (("E", "開始日"), ("F", "終了日")):
    dv = DataValidation(type="custom", allow_blank=True, imeMode="halfAlpha",
                        formula1=(f"AND(ISNUMBER({col}{FIRST}),OR($R{FIRST}>0,NETWORKDAYS({col}{FIRST},{col}{FIRST},{OFF_DAYS})=1),"
                                  f'COUNTIFS($L${FIRST}:$L${LAST},"承認",$E${FIRST}:$E${LAST},"<="&{col}{FIRST},'
                                  f'$W${FIRST}:$W${LAST},">="&{col}{FIRST},$Z${FIRST}:$Z${LAST},1)=0)'))
    dv.errorTitle, dv.error = label, (f"この日は{label}にできません。土日・祝日・年末年始か、すでに承認された全日の休暇と重なっています。")
    ws.add_data_validation(dv)
    dv.add(col_cells(col))
for col in "AM":
    dv = DataValidation(type="date", operator="greaterThan", formula1="43831", allow_blank=True, imeMode="halfAlpha")   # 2020/1/1 より後
    dv.errorTitle, dv.error = "日付を入れてください", "例：11/5 または 2026/11/5"
    ws.add_data_validation(dv)
    dv.add(col_cells(col))

# 本人・所属長が入力する欄だけ保護を外す（基本・計算式は保護）。パスワードはかけない（誤操作を防ぐため）
for r in range(FIRST, LAST + 1):
    for c in range(1, 15):   # A〜N（G の日数と J の時間は自動なので保護のまま）
        if c not in (7, 10):
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
us["A1"] = "休暇申請シートの使い方"
us["A1"].font = f_title
for i, text in enumerate([
    "■ 本人（申請する人）",
    "・「申請」シートの黄色の欄に、1件1行で書きます。休暇の種類・単位などは一覧から選びます。",
    "・日単位のときは 開始日・終了日（1日だけなら空欄）を、時間・分単位のときは 開始日・開始時刻・終了時刻を書きます。時間（1日あたり）は、休憩時間（12:00〜13:00。短縮の特例の人はその時間）を除いて自動で入ります。",
    "・単位は、休暇の種類によって選べるものが絞られます。1つに決まる休暇（介護時間・育児時間など）は自動で入ります。",
    "・日数は自動で入ります。開始日だけなら 1、終了日も書くと土日・祝日・年末年始を除いた日数です。",
    "・土日・祝日・年末年始は開始日・終了日にできません。",
    "・時間・分単位で何日か続けて取るときは、終了日も書きます。例：11/6〜11/9、時間 1:00 なら、日数は 2 で、11/6 と 11/9 に1時間ずつ、合計2時間です。1日ずつ1行で書いてもかまいません。",
    "・年次休暇の時間は15分単位、介護時間・育児時間は30分単位です。月をまたぐときは月ごとに分けて書きます。",
    "・書いたら、所属長に申請したことを伝えます（Notes や口頭で）。",
    "・事由は、年次休暇と、種別・内容を書いた行では要りません（欄が灰色になります）。ほかの休暇では一覧の定型文から選ぶか、自由に書きます。",
    "・承認済みの申請と日時が重なると赤くなり、所属長は承認できません。日時を直してください。",
    "・足りない欄や誤りのある欄は赤くなり、理由が「確認」の欄に出ます。赤い欄がなくなるように書き直してください。",
    "・上の「年次休暇の残り」で、残日数と承認待ちが分かります（1日＝7時間45分）。",
    "・承認された行は書き換えないでください。取り消したいときは所属長に伝えます。",
    "",
    "■ 所属長",
    "・青の欄に、承認・却下（取り消すときは取消）を選びます。承認を選んだら承認日も必ず入れます（入れるまで赤くなります）。必要ならコメントを書きます。",
    "・承認欄が空欄の行は「承認待ち」として数えます。",
    "",
    "■ 管理者",
    "・年・合計・本年使用の調整（導入前の使用など）は、出勤・休暇管理データ の「休暇等」シートで全職員分を管理し、休暇申請シートの作成.xlsm の［休暇申請シートを更新］でこの用紙に書き込みます（非表示の「管理」シート）。",
    "・このファイルは管理者の OneDrive にあり、本人と所属長にだけ共有しています。ほかの人には見えません。",
], start=3):
    us[f"A{i}"] = text
    us[f"A{i}"].font = f_bold if text.startswith("■") else f_base

wb.active = 0
wb.save(OUT)
print("saved", OUT)
