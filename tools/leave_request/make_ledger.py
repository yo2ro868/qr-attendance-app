"""休暇申請台帳.xlsx を作る（Forms の休暇申請を記録するブックの空の様式）。

使い方: python make_ledger.py [出力先.xlsx]
出力先を省略するとプロジェクト直下に 休暇申請台帳.xlsx を作る（同じ名前のファイルは上書き）。
できたブックは SharePoint（または OneDrive）に置き、Power Automate と Office Script から使う。
表（テーブル）の名前・列名はスクリプトが参照するので変えないこと。
"""
import sys
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

OUT = sys.argv[1] if len(sys.argv) > 1 else str(Path(__file__).resolve().parents[2] / "休暇申請台帳.xlsx")
FONT = "ＭＳ Ｐゴシック"
f_base = Font(name=FONT, size=10)
f_bold = Font(name=FONT, size=10, bold=True)
f_title = Font(name=FONT, size=14, bold=True, color="1E3A5F")
f_note = Font(name=FONT, size=9, color="555555")
fill_input = PatternFill("solid", fgColor="FFF2CC")
wrap = Alignment(vertical="center", wrap_text=True)

STATUSES = ["承認待ち", "承認", "却下", "取消"]

wb = Workbook()


def table(ws, name, columns, rows=1, style="TableStyleMedium2"):
    """1行目に見出し、その下に空の行を rows 行持つ表（テーブル）を作る。"""
    for i, (label, width, fmt) in enumerate(columns, start=1):
        ws.cell(row=1, column=i, value=label).font = f_bold
        ws.column_dimensions[get_column_letter(i)].width = width
        for r in range(2, rows + 2):
            c = ws.cell(row=r, column=i)
            c.font = f_base
            if fmt:
                c.number_format = fmt
    ref = f"A1:{get_column_letter(len(columns))}{rows + 1}"
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name=style, showRowStripes=True)
    ws.add_table(t)
    ws.freeze_panes = "A2"


def notes(ws, cell, lines):
    col, row = cell[0], int(cell[1:])
    ws[cell] = "使い方"
    ws[cell].font = f_bold
    for i, text in enumerate(lines, start=1):
        ws[f"{col}{row + i}"] = text
        ws[f"{col}{row + i}"].font = f_note


# ===== 使い方 =====
ws = wb.active
ws.title = "使い方"
ws.column_dimensions["A"].width = 110
ws["A1"] = "休暇申請台帳"
ws["A1"].font = f_title
for i, text in enumerate([
    "Microsoft Forms の休暇申請を記録するブックです。Power Automate が申請を「申請」表に追加し、承認結果を書き込みます。",
    "",
    "■ 管理者が入力するもの",
    "・職員：Forms に回答するときのメールアドレス、出勤用ID、氏名、承認者のメールアドレス。",
    "・年次休暇の持ち日数：職員ごと・年ごとに、その年に使える日数（繰越分＋本年分）。照合用Excel の年次休暇シートの「合計」を写します。",
    "    この仕組みを使い始める前に、その年に使った日数・時間があれば「導入前の使用」に入れます。",
    "",
    "■ 自動で入るもの",
    "・申請：Forms の回答1件ごとに1行。状況は 承認待ち → 承認／却下。取り消すときは状況を「取消」に書き換えます（集計から外れます）。",
    "",
    "■ 照合用Excel への反映",
    "・「集計」シートの B1 に年（西暦）を入れ、［自動化］タブから「照合用Excel用に集計」スクリプトを実行します。",
    "・承認済みの申請（と導入前の使用）を、照合用Excel の「年次休暇」「月別取得時間」シートと同じ列の並びで出力します。",
    "",
    "■ 注意",
    "・表（テーブル）の名前と見出しは Office Script が参照しています。列の追加・名前の変更はしないでください。",
    "・氏名が入るため、保存場所の共有範囲は管理者に限ってください。",
], start=3):
    ws[f"A{i}"] = text
    ws[f"A{i}"].font = f_bold if text.startswith("■") else f_base

# ===== 職員 =====
ws = wb.create_sheet("職員")
table(ws, "職員", [
    ("メールアドレス", 32, "@"), ("出勤用ID", 12, "@"), ("氏名", 16, None),
    ("承認者メールアドレス", 32, "@"), ("備考", 24, None),
])
notes(ws, "G1", [
    "・Forms に回答するときのメールアドレス（組織のアカウント）で職員を見分けます。",
    "・出勤用ID は照合用Excel の職員名簿と同じ ID にします。",
    "・承認者メールアドレスの人に承認依頼が届きます。",
    "・行を足すときは表の最後の行の下に入力すると表が広がります。",
])

# ===== 年次休暇の持ち日数 =====
ws = wb.create_sheet("年次休暇の持ち日数")
table(ws, "持ち日数", [
    ("出勤用ID", 12, "@"), ("年", 8, "0"), ("合計（日）", 11, "0"),
    ("導入前の使用（日）", 16, "0"), ("導入前の使用（時間）", 18, "[h]:mm"), ("備考", 24, None),
])
notes(ws, "H1", [
    "・職員ごと・年ごとに1行入力します。",
    "・合計（日）は、その年に使える日数（繰越分＋本年分）です。照合用Excel の年次休暇シートの「合計（自動）」を写します。",
    "・導入前の使用は、この仕組みを使い始める前にその年に使った年次休暇です。時間は「時間:分」（例 2:15）。",
    "・残日数は 合計 − 導入前の使用 − 承認済みの申請 で計算します（1日＝7時間45分）。",
])

# ===== 申請 =====
ws = wb.create_sheet("申請")
table(ws, "申請", [
    ("申請番号", 9, "0"), ("受付日時", 16, "@"), ("メールアドレス", 28, "@"), ("出勤用ID", 10, "@"),
    ("氏名", 14, None), ("休暇の種類", 20, None), ("種別・内容", 22, None), ("単位", 9, None),
    ("開始日", 11, "yyyy/m/d"), ("終了日", 11, "yyyy/m/d"), ("日数", 6, "0.##"), ("時間", 7, "[h]:mm"),
    ("時刻", 14, None), ("事由", 24, None), ("状況", 9, None), ("承認者", 14, None),
    ("処理日時", 16, "@"), ("承認者コメント", 24, None), ("確認事項", 40, None),
], style="TableStyleLight9")
dv = DataValidation(type="list", formula1='"' + ",".join(STATUSES) + '"', allow_blank=True)
ws.add_data_validation(dv)
dv.add("O2:O5000")

# ===== 集計（照合用Excel に貼り付ける形） =====
ws = wb.create_sheet("集計")
ws["A1"], ws["B1"] = "年（西暦）", 2026
ws["A1"].font = f_bold
ws["B1"].fill = fill_input
ws["D1"] = "B1 に年を入れ、［自動化］タブから「照合用Excel用に集計」を実行すると、4行目から下に出力されます。"
ws["D1"].font = f_note
for col, width in zip("ABCDEFGHIJKLMNOPQRSTU", [12, 8, 10, 10, 10, 3, 12, 8, 10] + [6] * 12):
    ws.column_dimensions[col].width = width

wb.save(OUT)
print("saved", OUT)
