"""別のPCに引き渡す「休暇申請一式」フォルダと zip を作る。

使い方: python make_package.py [出力先フォルダ]
出力先を省略するとプロジェクト直下に 休暇申請一式（フォルダ）と 休暇申請一式.zip を作る（同じ名前は作り直し）。
中身：手順書.html（休暇・フレックス・テレワーク・超過勤務をまとめた、目次つきの1つの手順書。管理者向け）、申請・報告シート（ひな形）.xlsx、
      申請・報告シートの作成.xlsm と 申請・報告シートの作成.bas（職員分の申請・報告シートを作るマクロ）、
      出勤・休暇管理データ.xlsm と 出勤・休暇管理データのマクロ.bas、テレワーク予定表（ひな形）.xlsx（tools/telework）、
      QRコード出勤管理 利用者マニュアル.docx（QR出勤管理アプリの職員向けの使い方。Git の外：プロジェクト直下か、前に作った一式のフォルダにあれば入れる）
職員・所属長に配る手順書（申請者向け・所属長向け）は、別の文書（claude.ai の文書。Word・PDF に書き出して配る）が正本なので、一式には入れない。
どの Excel のファイルも、最後に scrub_personal.py で作成者・最終更新者を消す（このPCの利用者の名前が入らないように）。
申請・報告シートのひな形には、プロジェクト直下の 超過勤務等命令簿（原本）.xlsx（いま使っている命令簿。Git には入れない）の
いちばん新しい月のシートを「超勤ひな形」として入れる（tools/overtime/add_overtime_sheet.ps1）。
また、プロジェクト直下の 申告・割振り簿（原本）.xlsx（総務課の様式。シートの保護を外したもの。Git には入れない）を写して、
様式どおりの完成版のシート「申告・割振り簿」を入れる（tools/flex/add_flex_book.ps1）。
出勤・休暇管理データ の Power Query は Excel で入れるので、このPCに Excel が必要。
"""
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
OT_SOURCE = HERE.parents[1] / "超過勤務等命令簿（原本）.xlsx"
FLEX_SOURCE = HERE.parents[1] / "申告・割振り簿（原本）.xlsx"
OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE.parents[1] / "休暇申請一式"

if not OT_SOURCE.exists():
    sys.exit(f"超過勤務等命令簿の原本がありません：{OT_SOURCE}（いま使っている命令簿の Excel をこの名前で置いてください）")
if not FLEX_SOURCE.exists():
    sys.exit(f"申告・割振り簿の原本がありません：{FLEX_SOURCE}（総務課の様式の Excel を、シートの保護を外してこの名前で置いてください）")
# 一式に入れる Git の外の文書（プロジェクト直下か、前に作った一式のフォルダにあるもの）。作り直しで消えないように、先に読んでおく
EXTRAS = ["QRコード出勤管理 利用者マニュアル.docx"]
extras = {}
for name in EXTRAS:
    for src in (HERE.parents[1] / name, OUT / name, HERE.parents[1] / "休暇申請一式" / name):
        if src.exists():
            extras[name] = src.read_bytes()
            break
if OUT.exists():
    shutil.rmtree(OUT)
OUT.mkdir(parents=True)


def ps(script, *args):
    subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script), *args], check=True)


shutil.copy(HERE / "手順書.html", OUT / "手順書.html")
for name, data in extras.items():
    (OUT / name).write_bytes(data)
# 申請・報告シートのひな形。式が多く Excel の再計算中に呼び出しを断られる（RPC_E_CALL_REJECTED）ことがあるので、失敗したら作り直す（3回まで）
for attempt in range(1, 4):
    try:
        subprocess.run([sys.executable, str(HERE / "make_form.py"), str(OUT / "申請・報告シート（ひな形）.xlsx")], check=True)
        ps(HERE.parent / "overtime" / "add_overtime_sheet.ps1", "-Template", str(OUT / "申請・報告シート（ひな形）.xlsx"), "-Source", str(OT_SOURCE))
        ps(HERE.parent / "flex" / "add_flex_book.ps1", "-Template", str(OUT / "申請・報告シート（ひな形）.xlsx"), "-Source", str(FLEX_SOURCE))
        break
    except subprocess.CalledProcessError:
        if attempt == 3:
            raise
        print(f"ひな形の作成に失敗しました。作り直します（{attempt + 1}回目）")
ps(HERE / "build_macro.ps1", "-Out", str(OUT / "申請・報告シートの作成.xlsm"))
ps(HERE.parent / "matching_excel" / "build.ps1", "-Out", str(OUT / "出勤・休暇管理データ.xlsm"))
TELEWORK = HERE.parent / "telework"
subprocess.run([sys.executable, str(TELEWORK / "make_form.py"), str(OUT / "テレワーク予定表（ひな形）.xlsx")], check=True)

# 作成者・最終更新者などの個人の情報を消し、保存時に個人情報を残さない設定にする（別のPCに渡すため）
subprocess.run([sys.executable, str(HERE / "scrub_personal.py"), str(OUT)], check=True)

# zip（Excel で開いているときにできるロックファイル ~$… は入れない）
zip_path = OUT.with_suffix(".zip")
with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
    for f in sorted(OUT.iterdir()):
        if not f.name.startswith("~$"):
            z.write(f, f"{OUT.name}/{f.name}")
print("作成しました:", OUT, zip_path)
