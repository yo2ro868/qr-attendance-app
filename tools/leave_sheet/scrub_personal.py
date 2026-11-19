"""Excel のファイル（.xlsx・.xlsm）から、作成者・最終更新者などの個人の情報を消す。

使い方: python scrub_personal.py ファイルまたはフォルダ [...]
- docProps/core.xml の作成者（dc:creator）・最終更新者（cp:lastModifiedBy）を空にする
- docProps/app.xml の会社名（Company）・管理者（Manager）を空にする
- xl/workbook.xml の workbookPr に filterPrivacy="1"（Excel の「保存時にファイルのプロパティから個人情報を削除する」）を入れる。
  このひな形から作った申請・報告シートなどを、あとで別のPCで保存しても、保存した人の名前が残らない
中身（シート・式・マクロ）は変えない。zip の中の、ほかのファイルはそのまま写す。
"""
import re
import sys
import zipfile
from pathlib import Path


def _blank(xml: str, tags) -> str:
    for tag in tags:
        xml = re.sub(rf"<{tag}>[^<]*</{tag}>", f"<{tag}></{tag}>", xml)
    return xml


def _privacy(xml: str) -> str:
    m = re.search(r"<workbookPr\b[^>]*?/?>", xml)
    if m is None:          # workbookPr がなければ、workbook の最初の子として入れる（fileVersion の後ろが決まりの順）
        return re.sub(r"(<fileVersion\b[^>]*/>)", r'\1<workbookPr filterPrivacy="1"/>', xml, count=1) if "<fileVersion" in xml \
            else re.sub(r"(<workbook\b[^>]*>)", r'\1<workbookPr filterPrivacy="1"/>', xml, count=1)
    tag = m.group(0)
    if "filterPrivacy=" in tag:
        new = re.sub(r'filterPrivacy="[^"]*"', 'filterPrivacy="1"', tag)
    else:
        new = tag.replace("<workbookPr", '<workbookPr filterPrivacy="1"', 1)
    return xml.replace(tag, new, 1)


def scrub(path: Path) -> None:
    src = zipfile.ZipFile(path)
    items = [(info, src.read(info.filename)) for info in src.infolist()]
    src.close()
    out = []
    for info, data in items:
        name = info.filename
        if name == "docProps/core.xml":
            data = _blank(data.decode("utf-8"), ["dc:creator", "cp:lastModifiedBy"]).encode("utf-8")
        elif name == "docProps/app.xml":
            data = _blank(data.decode("utf-8"), ["Company", "Manager"]).encode("utf-8")
        elif name == "xl/workbook.xml":
            data = _privacy(data.decode("utf-8")).encode("utf-8")
        out.append((info, data))
    tmp = path.with_name(path.name + ".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for info, data in out:
            z.writestr(info, data)
    tmp.replace(path)


def targets(args):
    for a in args:
        p = Path(a)
        if p.is_dir():
            yield from sorted(f for f in p.iterdir() if f.suffix.lower() in (".xlsx", ".xlsm") and not f.name.startswith("~$"))
        else:
            yield p


if __name__ == "__main__":
    for f in targets(sys.argv[1:]):
        scrub(f)
        print("個人の情報を消しました:", f)
