"""Быстрое чтение .xlsx: значения ячеек потоком, без построения объектов openpyxl.

openpyxl в обычном режиме создаёт объект на каждую ячейку и тратит на прайс
в 150 тысяч строк 15–20 секунд. Здесь лист читается потоком из XML внутри zip,
а значения приводятся так же, как это делает openpyxl при data_only=True:
числа — int или float, даты по формату ячейки — datetime, логические — bool,
у формул берётся сохранённое значение. Сетка получается та же, что у
openpyxl.iter_rows(values_only=True): с первой строки и первого столбца до
последней ячейки, объединённые диапазоны тоже расширяют сетку.

Если файл устроен непривычно, read_xlsx бросает исключение, и вызывающий код
читает книгу через openpyxl, как раньше.
"""

from __future__ import annotations

import posixpath
import re
import zipfile
from pathlib import Path
from xml.etree.ElementTree import iterparse, parse

from openpyxl.cell.text import Text
from openpyxl.styles.stylesheet import Stylesheet
from openpyxl.utils.cell import range_boundaries
from openpyxl.utils.datetime import CALENDAR_MAC_1904, CALENDAR_WINDOWS_1900, from_excel, from_ISO8601

MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
WORKSHEET_TYPE = "/worksheet"

T_C, T_ROW, T_V, T_IS, T_SI, T_T = MAIN + "c", MAIN + "row", MAIN + "v", MAIN + "is", MAIN + "si", MAIN + "t"
T_MERGE = MAIN + "mergeCell"
_REF = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")


class XlsxFastError(Exception):
    pass


def _col_index(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _cast_number(value: str):
    if "." in value or "E" in value or "e" in value:
        return float(value)
    return int(value)


def _shared_strings(zf: zipfile.ZipFile, name: str) -> list[str]:
    if name not in zf.namelist():
        return []
    out: list[str] = []
    with zf.open(name) as fh:
        for _, node in iterparse(fh):
            if node.tag != T_SI:
                continue
            kids = list(node)
            if len(kids) == 1 and kids[0].tag == T_T and not len(kids[0]):
                text = kids[0].text or ""
            else:
                text = Text.from_tree(node).content
            out.append(text.replace("x005F_", ""))
            node.clear()
    return out


def _target(base: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base), target))


def _rels(zf: zipfile.ZipFile, part: str) -> dict[str, tuple[str, str]]:
    folder, name = posixpath.split(part)
    rel_name = posixpath.join(folder, "_rels", name + ".rels")
    if rel_name not in zf.namelist():
        return {}
    root = parse(zf.open(rel_name)).getroot()
    return {
        el.get("Id"): (el.get("Type") or "", _target(part, el.get("Target") or ""))
        for el in root.iter(PKG_REL + "Relationship")
    }


def _workbook_part(zf: zipfile.ZipFile) -> str:
    for rid, (kind, target) in _rels(zf, "").items():
        if kind.endswith("/officeDocument"):
            return target
    if "xl/workbook.xml" in zf.namelist():
        return "xl/workbook.xml"
    raise XlsxFastError("в файле нет книги")


def _styles(zf: zipfile.ZipFile) -> tuple[set[int], set[int]]:
    """Какие стили ячеек — даты. Как в openpyxl: только xl/styles.xml и только если стили в нём есть."""
    if "xl/styles.xml" not in zf.namelist():
        return set(), set()
    sheet = Stylesheet.from_tree(parse(zf.open("xl/styles.xml")).getroot())
    if not sheet.cell_styles:
        return set(), set()
    return set(sheet.date_formats), set(sheet.timedelta_formats)


def _inline_text(node) -> str:
    kids = list(node)
    if len(kids) == 1 and kids[0].tag == T_T and not len(kids[0]):
        return kids[0].text or ""
    return Text.from_tree(node).content


def _read_sheet(zf: zipfile.ZipFile, part: str, strings: list[str], dates: set[int], deltas: set[int], epoch):
    """Строки листа разбираются целиком в конце тега <row>: так нужны только события end."""
    rows: dict[int, dict[int, object]] = {}
    merges: list[tuple[int, int, int, int]] = []
    row_n = 0
    max_r = max_c = 0
    any_cell = False
    match = _REF.match
    with zf.open(part) as fh:
        for _, el in iterparse(fh):
            tag = el.tag
            if tag == T_ROW:
                r = el.get("r")
                if r is not None:
                    try:
                        row_n = int(r)
                    except ValueError:
                        val = float(r)
                        if not val.is_integer():
                            raise XlsxFastError(f"номер строки {r}")
                        row_n = int(val)
                else:
                    row_n += 1
                col_n = 0
                for c in el:
                    if c.tag != T_C:
                        continue
                    ref = c.get("r")
                    if ref:
                        m = match(ref)
                        if not m:
                            raise XlsxFastError(f"адрес ячейки {ref}")
                        letters = m.group(1)
                        if len(letters) == 1:
                            col_n = ord(letters.upper()) - 64
                        else:
                            col_n = _col_index(letters)
                        row = int(m.group(2))
                    else:
                        col_n += 1
                        row = row_n
                    kind = c.get("t", "n")
                    value = None
                    if kind == "inlineStr":
                        for ch in c:
                            if ch.tag == T_IS:
                                value = _inline_text(ch)
                                break
                    else:
                        for ch in c:
                            if ch.tag == T_V:
                                value = ch.text or None
                                break
                        if value is not None:
                            if kind == "n":
                                value = _cast_number(value)
                                style = c.get("s")
                                if style and dates and int(style) in dates:
                                    try:
                                        value = from_excel(value, epoch, timedelta=int(style) in deltas)
                                    except (OverflowError, ValueError):
                                        value = "#VALUE!"
                            elif kind == "s":
                                value = strings[int(value)]
                            elif kind == "b":
                                value = bool(int(value))
                            elif kind == "d":
                                value = from_ISO8601(value)
                    bucket = rows.get(row)
                    if bucket is None:
                        bucket = rows[row] = {}
                    bucket[col_n] = value
                    any_cell = True
                    if row > max_r:
                        max_r = row
                    if col_n > max_c:
                        max_c = col_n
                el.clear()
            elif tag == T_MERGE:
                ref = el.get("ref")
                if ref and ":" in ref:
                    c1, r1, c2, r2 = range_boundaries(ref)
                    merges.append((r1, r2, c1, c2))
    if not any_cell and not merges:
        return [], merges
    for r1, r2, c1, c2 in merges:
        # Объединённые ячейки openpyxl заводит как пустые — они тоже расширяют лист.
        if r2 > max_r:
            max_r = r2
        if c2 > max_c:
            max_c = c2
        for r in range(r1, r2 + 1):
            bucket = rows.get(r)
            if bucket is None:
                bucket = rows[r] = {}
            for c in range(c1, c2 + 1):
                if (r, c) != (r1, c1):
                    bucket[c] = None
    grid = []
    for r in range(1, max_r + 1):
        line = [None] * max_c
        bucket = rows.get(r)
        if bucket:
            for c, v in bucket.items():
                line[c - 1] = v
        grid.append(line)
    return grid, merges


def read_xlsx(path: Path) -> list[tuple[str, list[list], list[tuple[int, int, int, int]]]]:
    """Листы книги: (имя, сетка значений, объединённые диапазоны (r1, r2, c1, c2) с единицы)."""
    with zipfile.ZipFile(path) as zf:
        wb_part = _workbook_part(zf)
        rels = _rels(zf, wb_part)
        root = parse(zf.open(wb_part)).getroot()
        pr = root.find(MAIN + "workbookPr")
        epoch = CALENDAR_WINDOWS_1900
        if pr is not None and pr.get("date1904") in ("1", "true"):
            epoch = CALENDAR_MAC_1904
        strings_part = next((t for k, t in rels.values() if k.endswith("/sharedStrings")), "xl/sharedStrings.xml")
        strings = _shared_strings(zf, strings_part)
        dates, deltas = _styles(zf)
        sheets = []
        for sheet in root.iter(MAIN + "sheet"):
            rid = sheet.get(REL + "id")
            kind, part = rels.get(rid, ("", ""))
            if not kind.endswith(WORKSHEET_TYPE):
                continue
            if part not in zf.namelist():
                raise XlsxFastError(f"нет листа {part}")
            grid, merges = _read_sheet(zf, part, strings, dates, deltas, epoch)
            sheets.append((sheet.get("name") or "", grid, merges))
        return sheets
