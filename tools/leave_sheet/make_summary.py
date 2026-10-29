"""休暇集計.xlsx の土台（設定シートと出力先のシート）を作る。

使い方: python make_summary.py 出力先.xlsx
このあと build_summary.ps1 が Excel で Power Query（queries フォルダの .pq）を入れて、出力先のシートに表を作る。
"""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

OUT = sys.argv[1]
FONT = "ＭＳ Ｐゴシック"
f_base = Font(name=FONT, size=10)
f_bold = Font(name=FONT, size=10, bold=True)
f_title = Font(name=FONT, size=14, bold=True, color="1E3A5F")
f_note = Font(name=FONT, size=9, color="555555")
fill_input = PatternFill("solid", fgColor="FFF2CC")

wb = Workbook()
ws = wb.active
ws.title = "設定"
ws.column_dimensions["A"].width = 22
ws.column_dimensions["B"].width = 70
ws["A1"] = "休暇集計"
ws["A1"].font = f_title
ws["A3"], ws["A4"] = "申請用紙のフォルダ", "集計する年（西暦）"
ws["A3"].font = ws["A4"].font = f_bold
ws["B3"].fill = ws["B4"].fill = fill_input
ws["B3"].font = ws["B4"].font = f_base
ws["B4"] = 2026
wb.defined_names["フォルダ"] = DefinedName("フォルダ", attr_text="'設定'!$B$3")
wb.defined_names["集計年"] = DefinedName("集計年", attr_text="'設定'!$B$4")
for i, text in enumerate([
    "使い方",
    "・B3 に、申請用紙（休暇申請_出勤用ID_氏名.xlsx）を置いた OneDrive の同期フォルダの場所を入れます。",
    "    例：C:\\Users\\（ユーザー名）\\OneDrive - （組織名）\\休暇申請",
    "    エクスプローラーでフォルダを開き、上のアドレス欄をクリックして表示される場所をコピーして貼り付けます。",
    "・B4 に集計する年を入れます。",
    "・［データ］→［すべて更新］を押すと、各シートが最新になります（全員の申請用紙を読み直します）。",
    "",
    "・申請一覧：全員の申請。状況（承認待ち・承認・却下・取消）や確認の注意で絞り込めます。読み取れなかったファイルは先頭に出ます。",
    "・照合用_年次休暇：照合用Excel の「年次休暇」シートの 本年使用（日）・本年使用（時間）に写す値（承認済み＋導入前の使用）。",
    "・照合用_月別取得時間：照合用Excel の「月別取得時間」シートに行ごと貼り付ける値（承認済みの病気休暇・看護休暇。1日＝7時間45分）。",
], start=6):
    ws[f"A{i}"] = text
    ws[f"A{i}"].font = f_bold if text == "使い方" else f_note

for name, note in [
    ("申請一覧", "全員の申請（［データ］→［すべて更新］で最新になります）"),
    ("照合用_年次休暇", "照合用Excel の「年次休暇」シートの、同じ出勤用ID・年の行の 本年使用（日）・本年使用（時間）に写します"),
    ("照合用_月別取得時間", "見出しを除いてコピーし、照合用Excel の「月別取得時間」シートに値で貼り付けます（同じ ID・年・区分の古い行は消します）"),
]:
    s = wb.create_sheet(name)
    s["A1"] = note
    s["A1"].font = f_note

wb.save(OUT)
print("saved", OUT)
