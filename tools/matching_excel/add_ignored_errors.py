"""Add <ignoredErrors numberStoredAsText> to the two input sheets of the workbook."""
import re
import shutil
import sys
import zipfile

SRC, DST = sys.argv[1], sys.argv[2]
TARGETS = {"職員名簿": "A2:A501", "出力データ貼付": "A2:E5001", "年次休暇": "A2:A1001", "月別取得時間": "A2:A1501"}

with zipfile.ZipFile(SRC) as z:
    files = {n: z.read(n) for n in z.namelist()}

wb_xml = files["xl/workbook.xml"].decode("utf-8")
rels = files["xl/_rels/workbook.xml.rels"].decode("utf-8")
for name, sqref in TARGETS.items():
    rid = re.search(r'<sheet[^>]*name="%s"[^>]*r:id="(rId\d+)"' % name, wb_xml).group(1)
    target = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid, rels) or re.search(r'Target="([^"]+)"[^>]*Id="%s"' % rid, rels)
    path = "xl/" + target.group(1).lstrip("/").removeprefix("xl/")
    xml = files[path].decode("utf-8")
    assert "<ignoredErrors" not in xml
    tag = f'<ignoredErrors><ignoredError sqref="{sqref}" numberStoredAsText="1"/></ignoredErrors>'
    # ignoredErrors comes after pageSetup/headerFooter and before drawing/tableParts/extLst
    m = re.search(r"<(drawing|legacyDrawing|tableParts|extLst)\b", xml)
    xml = xml[: m.start()] + tag + xml[m.start():] if m else xml.replace("</worksheet>", tag + "</worksheet>")
    files[path] = xml.encode("utf-8")
    print(name, "->", path, sqref)

with zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED) as z:
    for n, data in files.items():
        z.writestr(n, data)
