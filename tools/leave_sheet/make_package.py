"""別のPCに引き渡す「休暇申請一式」フォルダと zip を作る（Forms と休暇申請シートの両方）。

使い方: python make_package.py [出力先フォルダ]
出力先を省略するとプロジェクト直下に 休暇申請一式（フォルダ）と 休暇申請一式.zip を作る（同じ名前は作り直し）。
中身：手順書（Forms）.html、手順書.html（休暇申請シート）、休暇申請シート（ひな形）.xlsx、休暇申請シートの作成.xlsm と 休暇申請シートの作成.bas（職員分の休暇申請シートを作るマクロ）、
      休暇集計.xlsx（申請一覧）、出勤・休暇管理データ.xlsm と 出勤・休暇管理データのマクロ.bas（休暇申請シート・Forms を読み込む版の空の様式と、［申請を取り込む］のマクロ）、
      手順書（テレワーク）.html と テレワーク予定表（ひな形）.xlsx（所属長が出すテレワークの予定。tools/telework）、
      手順書（超過勤務）.html（休暇申請シートの超過勤務等命令簿。tools/overtime）、
      手順書（フレックス）.html（休暇申請シートのフレックスタイム制の申告・割振り。tools/flex）
休暇申請シートのひな形には、プロジェクト直下の 超過勤務等命令簿（原本）.xlsx（いま使っている命令簿。Git には入れない）の
いちばん新しい月のシートを「超勤ひな形」として入れる（tools/overtime/add_overtime_sheet.ps1）。
また、プロジェクト直下の 申告・割振り簿（原本）.xlsx（総務課の様式。シートの保護を外したもの。Git には入れない）を写して、
様式どおりの完成版のシート「申告・割振り簿」を入れる（tools/flex/add_flex_book.ps1）。
休暇集計・出勤・休暇管理データ の Power Query は Excel で入れるので、このPCに Excel が必要。
"""
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
OT_SOURCE = HERE.parents[1] / "超過勤務等命令簿（原本）.xlsx"
FLEX_SOURCE = HERE.parents[1] / "申告・割振り簿（原本）.xlsx"
OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE.parents[1] / "休暇申請一式"

if not OT_SOURCE.exists():
    sys.exit(f"超過勤務等命令簿の原本がありません：{OT_SOURCE}（いま使っている命令簿の Excel をこの名前で置いてください）")
if not FLEX_SOURCE.exists():
    sys.exit(f"申告・割振り簿の原本がありません：{FLEX_SOURCE}（総務課の様式の Excel を、シートの保護を外してこの名前で置いてください）")
if OUT.exists():
    shutil.rmtree(OUT)
OUT.mkdir(parents=True)


def ps(script, *args):
    subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script), *args], check=True)


shutil.copy(HERE / "手順書.html", OUT / "手順書.html")
shutil.copy(HERE / "手順書（Forms）.html", OUT / "手順書（Forms）.html")
subprocess.run([sys.executable, str(HERE / "make_form.py"), str(OUT / "休暇申請シート（ひな形）.xlsx")], check=True)
ps(HERE.parent / "overtime" / "add_overtime_sheet.ps1", "-Template", str(OUT / "休暇申請シート（ひな形）.xlsx"), "-Source", str(OT_SOURCE))
ps(HERE.parent / "flex" / "add_flex_book.ps1", "-Template", str(OUT / "休暇申請シート（ひな形）.xlsx"), "-Source", str(FLEX_SOURCE))
ps(HERE / "build_macro.ps1", "-Out", str(OUT / "休暇申請シートの作成.xlsm"))
ps(HERE / "build_summary.ps1", "-Out", str(OUT / "休暇集計.xlsx"))
ps(HERE.parent / "matching_excel" / "build.ps1", "-Out", str(OUT / "出勤・休暇管理データ.xlsm"))
TELEWORK = HERE.parent / "telework"
shutil.copy(TELEWORK / "手順書（テレワーク）.html", OUT / "手順書（テレワーク）.html")
subprocess.run([sys.executable, str(TELEWORK / "make_form.py"), str(OUT / "テレワーク予定表（ひな形）.xlsx")], check=True)

shutil.copy(HERE.parent / "overtime" / "手順書（超過勤務）.html", OUT / "手順書（超過勤務）.html")
shutil.copy(HERE.parent / "flex" / "手順書（フレックス）.html", OUT / "手順書（フレックス）.html")

zip_path = shutil.make_archive(str(OUT), "zip", root_dir=OUT.parent, base_dir=OUT.name)
print("作成しました:", OUT, zip_path)
