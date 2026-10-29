"""別のPCに引き渡す「休暇申請システム一式」フォルダと zip を作る。

使い方: python make_package.py [出力先フォルダ]
出力先を省略するとプロジェクト直下に 休暇申請システム一式（フォルダ）と 休暇申請システム一式.zip を作る（同じ名前は作り直し）。
中身：手順書.html、休暇申請台帳.xlsx（空の様式）、スクリプト フォルダの .txt（Office Script。メモ帳で開いてコピーする）
引き渡し先のPCでは Python などは使わず、ブラウザの Microsoft 365 だけで作業する。
"""
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE.parents[1] / "休暇申請システム一式"
SCRIPTS = ["申請を登録", "承認結果を記録", "残日数を照会", "照合用Excel用に集計"]

if OUT.exists():
    shutil.rmtree(OUT)
(OUT / "スクリプト").mkdir(parents=True)

shutil.copy(HERE / "手順書.html", OUT / "手順書.html")
subprocess.run([sys.executable, str(HERE / "make_ledger.py"), str(OUT / "休暇申請台帳.xlsx")], check=True)
for i, name in enumerate(SCRIPTS, start=1):
    text = (HERE / "office_scripts" / f"{name}.ts").read_text(encoding="utf-8")
    # メモ帳で文字化けしないように BOM 付き UTF-8・改行は CRLF
    (OUT / "スクリプト" / f"{i}_{name}.txt").write_text(text, encoding="utf-8-sig", newline="\r\n")

zip_path = shutil.make_archive(str(OUT), "zip", root_dir=OUT.parent, base_dir=OUT.name)
print("作成しました:", OUT, zip_path)
