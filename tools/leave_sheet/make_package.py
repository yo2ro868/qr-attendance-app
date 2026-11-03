"""別のPCに引き渡す「休暇申請一式」フォルダと zip を作る（Forms と休暇申請用紙の両方）。

使い方: python make_package.py [出力先フォルダ]
出力先を省略するとプロジェクト直下に 休暇申請一式（フォルダ）と 休暇申請一式.zip を作る（同じ名前は作り直し）。
中身：手順書（Forms）.html、手順書.html（申請用紙）、休暇申請用紙（ひな形）.xlsx、申請用紙の作成.xlsm と 申請用紙の作成.bas（職員分の申請用紙を作るマクロ）、
      休暇集計.xlsx（申請一覧）、照合用Excel.xlsx（申請用紙を読み込む版の空の様式）
休暇集計・照合用Excel の Power Query は Excel で入れるので、このPCに Excel が必要。
"""
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE.parents[1] / "休暇申請一式"

if OUT.exists():
    shutil.rmtree(OUT)
OUT.mkdir(parents=True)


def ps(script, *args):
    subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script), *args], check=True)


shutil.copy(HERE / "手順書.html", OUT / "手順書.html")
shutil.copy(HERE / "手順書（Forms）.html", OUT / "手順書（Forms）.html")
subprocess.run([sys.executable, str(HERE / "make_form.py"), str(OUT / "休暇申請用紙（ひな形）.xlsx")], check=True)
ps(HERE / "build_macro.ps1", "-Out", str(OUT / "申請用紙の作成.xlsm"))
ps(HERE / "build_summary.ps1", "-Out", str(OUT / "休暇集計.xlsx"))
ps(HERE.parent / "matching_excel" / "build.ps1", "-Out", str(OUT / "照合用Excel.xlsx"))

zip_path = shutil.make_archive(str(OUT), "zip", root_dir=OUT.parent, base_dir=OUT.name)
print("作成しました:", OUT, zip_path)
