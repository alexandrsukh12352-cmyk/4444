"""Быстрая запись простой книги Excel: шапка, строки, форматы чисел и дат, фильтр, закрепление.

Выгрузка цен по датам — сотни тысяч строк. openpyxl пишет такую книгу минуту и больше,
здесь XML листа собирается строками сразу. Строки текста — в общей таблице (sharedStrings):
поставщики, города и бренды повторяются, и файл получается в разы меньше.
Книга открывается в Excel 2016: только базовые части формата, без расширений.
"""

from __future__ import annotations

import datetime as dt
import re
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

MAX_ROWS = 1_048_576
_ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")
_EPOCH = dt.date(1899, 12, 30)

# Стили: 0 — обычный, 1 — шапка, 2 — дата, 3 — деньги, 4 — процент со знаком, 5 — целое.
FORMATS = {"date": 2, "money": 3, "pct": 4, "int": 5}

_STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="3"><numFmt numFmtId="164" formatCode="dd\\.mm\\.yyyy"/><numFmt numFmtId="165" formatCode="#,##0.00"/><numFmt numFmtId="166" formatCode="+0.0%;\\-0.0%;0.0%"/></numFmts>
<fonts count="2"><font><sz val="10"/><name val="Arial"/><family val="2"/></font><font><b/><sz val="10"/><name val="Arial"/><family val="2"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE8EEF2"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="6">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="1" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""


def _letters(n: int) -> list[str]:
    out = []
    for i in range(n):
        x, s = i + 1, ""
        while x:
            x, r = divmod(x - 1, 26)
            s = chr(65 + r) + s
        out.append(s)
    return out


def _text(value) -> str:
    return escape(_ILLEGAL.sub("", str(value)))


class _Strings:
    def __init__(self):
        self.index: dict[str, int] = {}
        self.count = 0

    def get(self, value: str) -> int:
        self.count += 1
        i = self.index.get(value)
        if i is None:
            i = self.index[value] = len(self.index)
        return i

    def xml(self) -> str:
        items = "".join(
            f'<si><t xml:space="preserve">{_text(s)}</t></si>' if s[:1].isspace() or s[-1:].isspace()
            else f"<si><t>{_text(s)}</t></si>"
            for s in self.index
        )
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                f'count="{self.count}" uniqueCount="{len(self.index)}">{items}</sst>')


def _sheet_xml(sheet: dict, strings: _Strings) -> tuple[str, int]:
    header = list(sheet.get("header") or [])
    formats = sheet.get("formats") or {}
    widths = sheet.get("widths") or []
    width = max(len(header), max((i + 1 for i in formats), default=0))
    rows = sheet.get("rows") or []
    letters = _letters(max(width, 1) + 64)
    parts = []
    n = 0
    if header:
        n = 1
        cells = "".join(f'<c r="{letters[i]}1" t="s" s="1"><v>{strings.get(str(h))}</v></c>' for i, h in enumerate(header))
        parts.append(f'<row r="1">{cells}</row>')
    for row in rows:
        if n >= MAX_ROWS:
            break
        n += 1
        cells = []
        for i, value in enumerate(row):
            if value is None or value == "":
                continue
            if i >= len(letters):
                letters = _letters(i + 64)
            ref = f"{letters[i]}{n}"
            style = formats.get(i, 0)
            if isinstance(value, bool):
                cells.append(f'<c r="{ref}" t="s"><v>{strings.get("да" if value else "нет")}</v></c>')
            elif isinstance(value, (int, float)):
                if value != value or value in (float("inf"), float("-inf")):
                    continue
                s = f' s="{style}"' if style else ""
                text = repr(float(value)) if isinstance(value, float) else str(value)
                if text.endswith(".0"):
                    text = text[:-2]
                cells.append(f'<c r="{ref}"{s}><v>{text}</v></c>')
            elif isinstance(value, (dt.date, dt.datetime)):
                day = value.date() if isinstance(value, dt.datetime) else value
                cells.append(f'<c r="{ref}" s="{FORMATS["date"]}"><v>{(day - _EPOCH).days}</v></c>')
            else:
                cells.append(f'<c r="{ref}" t="s"><v>{strings.get(str(value))}</v></c>')
        parts.append(f'<row r="{n}">{"".join(cells)}</row>')
        width = max(width, len(row))
    last = f"{_letters(max(width, 1))[-1]}{max(n, 1)}"
    cols = ""
    if widths:
        cols = "<cols>" + "".join(
            f'<col min="{i + 1}" max="{i + 1}" width="{w}" customWidth="1"/>' for i, w in enumerate(widths) if w
        ) + "</cols>"
    view = ('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" '
            'state="frozen"/><selection pane="bottomLeft"/></sheetView></sheetViews>'
            if header and sheet.get("freeze", True) else '<sheetViews><sheetView workbookViewId="0"/></sheetViews>')
    auto = f'<autoFilter ref="A1:{last}"/>' if header and sheet.get("filter", True) and n > 1 else ""
    xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
           f'<dimension ref="A1:{last}"/>{view}<sheetFormatPr defaultRowHeight="12.75"/>{cols}'
           f'<sheetData>{"".join(parts)}</sheetData>{auto}</worksheet>')
    return xml, n


def write_book(path: Path, sheets: list[dict]) -> list[int]:
    """sheets: [{"name", "header", "rows", "formats": {столбец: "date"|"money"|"pct"|"int"}, "widths"}].

    Возвращает число строк на каждом листе вместе с шапкой.
    """
    strings = _Strings()
    bodies = []
    counts = []
    for sheet in sheets:
        fmt = {i: FORMATS[f] for i, f in (sheet.get("formats") or {}).items() if f in FORMATS}
        xml, n = _sheet_xml({**sheet, "formats": fmt}, strings)
        bodies.append(xml)
        counts.append(n)
    names = [_text(str(s["name"])[:31]) for s in sheets]
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    defined = ""
    filters = [i for i, s in enumerate(sheets) if s.get("header") and s.get("filter", True) and counts[i] > 1]
    if filters:
        refs = []
        for i in filters:
            m = re.search(r'<dimension ref="A1:([A-Z]+)(\d+)"/>', bodies[i])
            refs.append(f'<definedName name="_xlnm._FilterDatabase" localSheetId="{i}" hidden="1">'
                        f"'{names[i]}'!$A$1:${m.group(1)}${m.group(2)}</definedName>")
        defined = "<definedNames>" + "".join(refs) + "</definedNames>"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED, compresslevel=5) as z:
        z.writestr("[Content_Types].xml", (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            + "".join(f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" '
                      'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                      for i in range(len(sheets)))
            + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
            '</Types>'))
        z.writestr("_rels/.rels", (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>'))
        z.writestr("xl/workbook.xml", (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<bookViews><workbookView/></bookViews><sheets>'
            + "".join(f'<sheet name="{names[i]}" sheetId="{i + 1}" r:id="rId{i + 1}"/>' for i in range(len(sheets)))
            + f"</sheets>{defined}</workbook>"))
        z.writestr("xl/_rels/workbook.xml.rels", (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(f'<Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                      f'Target="worksheets/sheet{i + 1}.xml"/>' for i in range(len(sheets)))
            + f'<Relationship Id="rId{len(sheets) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            + f'<Relationship Id="rId{len(sheets) + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
            '</Relationships>'))
        z.writestr("xl/styles.xml", _STYLES)
        for i, body in enumerate(bodies):
            z.writestr(f"xl/worksheets/sheet{i + 1}.xml", body)
        z.writestr("xl/sharedStrings.xml", strings.xml())
    tmp.replace(path)
    return counts
