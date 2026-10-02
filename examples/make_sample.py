"""見本の Excel（架空の請求書発行資料）を作る。Excel やライブラリなしで、xlsx の中身を直接書く。

    python examples/make_sample.py examples/sample_2026-09.xlsx
"""

import sys
import zipfile
from pathlib import Path

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"""

WORKBOOK = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="請求書発行資料" sheetId="1" r:id="rId1"/></sheets></workbook>"""

WORKBOOK_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>"""

# 0: 標準 / 1: 太字で大きい見出し / 2: 日付（独自の表示形式 yyyy"年"m"月"d"日"）/ 3: 日付（組み込み 14）/ 4: 桁区切り
STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="176" formatCode="yyyy&quot;年&quot;m&quot;月&quot;d&quot;日&quot;"/></numFmts><fonts count="2"><font><sz val="11"/><name val="游ゴシック"/></font><font><b/><sz val="14"/><name val="游ゴシック"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="176" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>"""

# 0「2026年9月分 請求書発行資料」/ 1「対象期間」/ 2「作成日」/ 3「支払期日」/ 4「金額」/ 5「9月分の業務委託費」（ふりがな付き）
SHARED = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="6" uniqueCount="6"><si><t>2026年9月分 請求書発行資料</t></si><si><t>対象期間</t></si><si><t>作成日</t></si><si><t>支払期日</t></si><si><t>金額</t></si><si><t>9月分の業務委託費</t><rPh sb="0" eb="1"><t>ガツ</t></rPh></si></sst>"""

# 日付: 2026/9/30 = 46295、2026/10/5 = 46300、2026/10/31 = 46326
SHEET = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols><col min="1" max="1" width="18.5" customWidth="1"/><col min="2" max="2" width="24" customWidth="1"/></cols><sheetData><row r="1"><c r="A1" s="1" t="s"><v>0</v></c></row><row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2" t="inlineStr"><is><t>2026/09/01-2026/09/30</t></is></c></row><row r="3"><c r="A3" t="s"><v>2</v></c><c r="B3" s="2"><v>46300</v></c></row><row r="4"><c r="A4" t="s"><v>3</v></c><c r="B4" s="3"><v>46326</v></c></row><row r="5"><c r="A5" t="s"><v>5</v></c><c r="B5" s="4"><v>220000</v></c></row><row r="6"><c r="A6" t="s"><v>4</v></c><c r="B6" s="4"><f>B5*1.1</f><v>242000</v></c></row></sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="landscape"/></worksheet>"""


def make(path: Path) -> None:
    parts = {
        "[Content_Types].xml": CONTENT_TYPES,
        "_rels/.rels": ROOT_RELS,
        "xl/workbook.xml": WORKBOOK,
        "xl/_rels/workbook.xml.rels": WORKBOOK_RELS,
        "xl/styles.xml": STYLES,
        "xl/sharedStrings.xml": SHARED,
        "xl/worksheets/sheet1.xml": SHEET,
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, xml in parts.items():
            z.writestr(name, xml)


if __name__ == "__main__":
    make(Path(sys.argv[1] if len(sys.argv) > 1 else "sample_2026-09.xlsx"))
