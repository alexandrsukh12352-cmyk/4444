"""Книга итогового отчёта.

ПСТ, СБС и Продажи — отдельные листы, как выгрузка 1С. Числа с них стоят ссылками по коду
товара. СБС за штуку, закупочная цена, Кат. А, ВР, суммы и отклонения считаются в строке
отчёта. Курс, ГЗ и скидка записаны в формулах числами. Цена конкурента — число на момент
формирования. Суммы и количества — с разделением тысяч.

Полный отчёт: лист отчёта (или лист на подгруппу), лист «Кроссы», ПСТ, СБС и Продажи.
Быстрый отчёт: короткий лист, те же «Кроссы» и выгрузки.
Какие товары не попадают в файл, решают фильтры программы до сборки.
Строка «Показать / Скрыть» в лист не пишется: в Excel она не пересчитывается сама.
"""

from __future__ import annotations

import datetime as dt
import os
import re
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

# lxml пишет части книги в порядке, который Excel считает битым XML.
os.environ["OPENPYXL_LXML"] = "False"

from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import coordinate_to_tuple, get_column_letter

import openpyxl.cell._writer as _cell_writer

if _cell_writer.LXML:
    _cell_writer.LXML = False
    _cell_writer.write_cell = _cell_writer.etree_write_cell

from . import style
from .style import argb
from .prices import PriceFile
from .report import (
    CROSS_MODES, CROSS_SPLITS, FIELDS, FILTERS, PRICE_SOURCES, REPORT_KINDS, SOURCE_CHOICES, SUBGROUP_MODES,
    Options, Row, Sources, acp_crosses, competitor_order, prepare, result_bucket, result_folder_name,
    rows_for_workbook, segment_label, svod_comp, svod_for,
)
from .recommend import TARGETS, WEIGHT_CLASSES
from .settings import Settings
from .sources import KIND_TITLES, norm_article

HEAD_COLORS = {
    key: argb(color) for key, color in {
        "id": "0F4C5C", "cross": "5B3E96", "info": "52606D", "car": "2F6B4F", "buy": "52606D",
        "calc": "1F3A4D", "comp": "9A5B13", "pos": "7A3E65", "rec": "1D5FA8", "sister": "0F766E",
    }.items()
}
# Без длинного тире и кириллицы: Excel 2016 из‑за такого формата чисел выбрасывает весь styles.xml.
MONEY = '#,##0.00'
RUB = '#,##0'
INT = '#,##0'
PCT = '0.0%;-0.0%;0.0%'
PCT_SIGNED = '+0.0%;-0.0%;0.0%'
CHANGE = '+0.0%;-0.0%;""'  # цену не меняем — пусто, а не 0,0%
COEF = '0.000'
INK = Font(name=style.FONT, color=style.INK)
INPUT = Font(name=style.FONT, color=argb("1F4E9E"))  # синий — исходное значение, можно править
NO_FONT = Font(name=style.FONT, italic=True, color=argb("8A5A00"))
ACP_FONT = Font(name=style.FONT, color=argb("52606D"))
CAT_A_FILL = PatternFill("solid", fgColor=argb("FFFBE0"))
REC_FILL = PatternFill("solid", fgColor=argb("EAF2FD"))
COMP_TONES = (PatternFill("solid", fgColor=argb("FBF6EE")), PatternFill("solid", fgColor=argb("EEF3F6")))
SISTER_FILL = PatternFill("solid", fgColor=argb("D1FAE5"))
COMP_EDGE = Border(left=Side(style="medium", color=argb("C4A574")), bottom=Side(style="thin", color=argb("C9D1D9")))
ROW_BORDER = Border(bottom=Side(style="thin", color=argb("C9D1D9")))
# Как пишет сам Excel: в условном формате цвет лежит в bgColor, без patternType.
GREEN = (PatternFill(bgColor=argb("C6EFCE")), Font(color=argb("006100")))
RED = (PatternFill(bgColor=argb("FFC7CE")), Font(color=argb("9C0006")))
# Строка 1 — название, 2 — итоги по фильтру, 3 — надписи блоков, 4 — шапка.
TITLE_ROW = 1
TOTAL_ROW = 2
BAND_ROW = 3
HEADER_ROW = 4
FIRST = 5  # первая строка данных на листе «Товары»


def code_value(code: str):
    return int(code) if code.isdigit() and not code.startswith("0") else code


def q(name: str) -> str:
    return "'" + name.replace("'", "''") + "'"


def excel_num(value: float) -> str:
    text = f"{float(value):.10f}".rstrip("0").rstrip(".")
    return text if text not in ("", "-", "-0") else "0"


def took_own(mode: str, own, acp) -> bool:
    """Число взято из нашей выгрузки, а не из АЦПКД."""
    empty = lambda v: v is None or v == ""  # noqa: E731
    if mode == "own":
        return True
    if mode == "acp":
        return False
    if mode == "acp_own":
        return empty(acp)
    return not empty(own)


def index_match(sheet: str, value_col: str, key_col: str, key_cell: str) -> str:
    """Ссылка на ячейку выгрузки по коду товара. Лист можно заменить — формула найдёт строку."""
    return (
        f"IFERROR(INDEX({q(sheet)}!${value_col}:${value_col},"
        f"MATCH({key_cell},{q(sheet)}!${key_col}:${key_col},0)),\"\")"
    )


@dataclass
class Live:
    """Столбцы исходных листов ПСТ, СБС и Продажи — как в выгрузке 1С."""
    pst: object | None = None
    sbs: object | None = None
    sales: object | None = None
    names: dict[str, str] = field(default_factory=dict)

    def pull(self, kind: str, field: str, key_cell: str) -> str | None:
        loc = getattr(self, kind)
        if loc is None or kind not in self.names:
            return None
        value_idx = loc.cols.get(field)
        key_idx = loc.cols.get("code")
        if value_idx is None or key_idx is None:
            return None
        return index_match(self.names[kind], get_column_letter(value_idx + 1), get_column_letter(key_idx + 1), key_cell)


def write_grid(ws, grid: list[list]) -> None:
    """Лист как выгрузка: те же ячейки, без оформления. Дробные числа — с разделением тысяч."""
    ws.sheet_view.showGridLines = True
    for r, row in enumerate(grid, 1):
        for c, value in enumerate(row, 1):
            if value is None or value == "":
                continue
            cell = ws.cell(r, c, value)
            if isinstance(value, float) and not isinstance(value, bool):
                cell.number_format = MONEY


# ---------------------------------------------------------------- скрытые листы выгрузок: XML один раз


@dataclass
class RawTemplate:
    """Готовый XML ячеек листа-выгрузки. Номера стилей заменены метками по порядку появления."""
    sheet_data: str
    dimension: str
    probes: list[tuple[int, int]]  # первая ячейка каждого стиля — по ней узнаём номер стиля в книге


_RAW_TEMPLATES: dict[int, tuple[list, RawTemplate]] = {}
_STYLE_ATTR = re.compile(r'(<c r="[A-Z]+\d+") s="(\d+)"')
_SHEET_DATA = re.compile(r"<sheetData\s*/>|<sheetData>.*</sheetData>", re.S)
_DIMENSION = re.compile(r'<dimension ref="([^"]*)"\s*/>')
_MARK = "\x00{}\x00"


_ILLEGAL = re.compile(r"[\000-\010]|[\013-\014]|[\016-\037]")
_ERROR_CODES = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"}


def _raw_template_fast(grid: list[list]) -> RawTemplate | None:
    """Тот же XML, что пишет openpyxl, но без объектов ячеек. None — в листе есть то, что так не пишем
    (даты, формулы, коды ошибок, недопустимые символы), тогда лист собирает сам openpyxl."""
    from xml.etree.ElementTree import _escape_cdata as esc

    mark = _MARK.format(0)
    parts = ["<sheetData>"]
    probes: list[tuple[int, int]] = []
    min_r = min_c = None
    max_r = max_c = 0
    letters: dict[int, str] = {}
    for r, row in enumerate(grid, 1):
        cells = []
        for c, v in enumerate(row, 1):
            if v is None or v == "":
                continue
            t = type(v)
            col = letters.get(c)
            if col is None:
                col = letters[c] = get_column_letter(c)
            if t is str:
                if len(v) > 32767 or (len(v) > 1 and v[0] == "=") or v in _ERROR_CODES or _ILLEGAL.search(v):
                    return None
                stripped = v.strip()
                space = ' xml:space="preserve"' if stripped and v != stripped else ""
                cells.append(f'<c r="{col}{r}" t="inlineStr"><is><t{space}>{esc(v)}</t></is></c>')
            elif t is bool:
                cells.append(f'<c r="{col}{r}" t="b"><v>{"1" if v else "0"}</v></c>')
            elif t is int:
                cells.append(f'<c r="{col}{r}" t="n"><v>{"%.16g" % v}</v></c>')
            elif t is float:
                if v != v or v in (float("inf"), float("-inf")):
                    return None
                if not probes:
                    probes.append((r, c))
                cells.append(f'<c r="{col}{r}" s="{mark}" t="n"><v>{"%.16g" % v}</v></c>')
            else:
                return None
            if min_r is None:
                min_r = r
            min_c = c if min_c is None else min(min_c, c)
            max_r = r
            max_c = max(max_c, c)
        if cells:
            parts.append(f'<row r="{r}">{"".join(cells)}</row>')
    if min_r is None:
        return RawTemplate("<sheetData></sheetData>", "A1:A1", [])
    parts.append("</sheetData>")
    dimension = f"{get_column_letter(min_c)}{min_r}:{get_column_letter(max_c)}{max_r}"
    return RawTemplate("".join(parts), dimension, probes)


def raw_template(grid: list[list]) -> RawTemplate | None:
    """XML листа-выгрузки, как его пишет openpyxl. Один раз на выгрузку за сеанс программы.

    Общие ПСТ, СБС и Продажи одинаковы в файле каждой группы. Раньше каждая группа
    заново создавала сотни тысяч ячеек; теперь лист пишется один раз, а в следующие
    книги вставляется готовым. Содержимое и оформление листа те же.
    """
    hit = _RAW_TEMPLATES.get(id(grid))
    if hit is not None and hit[0] is grid:
        return hit[1]
    tpl = _raw_template_fast(grid)
    if tpl is not None:
        _RAW_TEMPLATES[id(grid)] = (grid, tpl)
        while len(_RAW_TEMPLATES) > 6:
            _RAW_TEMPLATES.pop(next(iter(_RAW_TEMPLATES)))
        return tpl
    return _raw_template_openpyxl(grid)


def _raw_template_openpyxl(grid: list[list]) -> RawTemplate | None:
    """Лист через сам openpyxl — для выгрузок с датами, формулами и прочим непростым."""
    try:
        from io import BytesIO

        from openpyxl import Workbook
        from openpyxl.worksheet._writer import WorksheetWriter

        scratch = Workbook()
        ws = scratch.active
        write_grid(ws, grid)
        order: dict[int, int] = {}
        probes: list[tuple[int, int]] = []
        for (r, c), cell in sorted(ws._cells.items()):
            if cell.has_style:
                sid = cell.style_id
                if sid not in order:
                    order[sid] = len(order)
                    probes.append((r, c))
        writer = WorksheetWriter(ws, out=BytesIO())
        writer.write()
        xml = writer.read().decode("utf-8")
        data = _SHEET_DATA.search(xml)
        dim = _DIMENSION.search(xml)
        if not data or not dim:
            return None
        body = _STYLE_ATTR.sub(lambda m: f'{m.group(1)} s="{_MARK.format(order[int(m.group(2))])}"', data.group(0))
        tpl = RawTemplate(body, dim.group(1), probes)
    except Exception:  # noqa: BLE001 — не вышло: лист пишется как раньше, по ячейкам
        return None
    _RAW_TEMPLATES[id(grid)] = (grid, tpl)
    while len(_RAW_TEMPLATES) > 6:
        _RAW_TEMPLATES.pop(next(iter(_RAW_TEMPLATES)))
    return tpl


def write_raw_sheet(ws, grid: list[list]) -> RawTemplate | None:
    """Лист-выгрузка: в книгу ставятся только ячейки-образцы стилей, остальное вставится в XML после сохранения."""
    tpl = raw_template(grid)
    if tpl is None:
        write_grid(ws, grid)
        return None
    ws.sheet_view.showGridLines = True
    for r, c in tpl.probes:
        value = grid[r - 1][c - 1]
        cell = ws.cell(r, c, value)
        if isinstance(value, float) and not isinstance(value, bool):
            cell.number_format = MONEY
    return tpl


def _inject_raw(text: str, tpl: RawTemplate) -> str:
    """Подставить готовые ячейки в XML листа. Номера стилей берутся у ячеек-образцов этой книги."""
    ids = []
    for r, c in tpl.probes:
        coord = f"{get_column_letter(c)}{r}"
        m = re.search(rf'<c r="{coord}" s="(\d+)"', text)
        if not m:
            raise ValueError(f"нет ячейки-образца {coord}")
        ids.append(m.group(1))
    body = tpl.sheet_data
    for k, sid in enumerate(ids):
        body = body.replace(_MARK.format(k), sid)
    if "\x00" in body or not _SHEET_DATA.search(text):
        raise ValueError("лист выгрузки не собрался")
    text = _SHEET_DATA.sub(lambda _m: body, text, count=1)
    return _DIMENSION.sub(lambda _m: f'<dimension ref="{tpl.dimension}" />', text, count=1)


# ---------------------------------------------------------------- ссылки на справочники и цены


@dataclass
class Refs:
    rates: dict[str, str] = field(default_factory=dict)  # валюта -> 'Данные'!$B$3
    gz: dict[str, str] = field(default_factory=dict)  # группа заказа
    discounts: dict[tuple[str, str, str], str] = field(default_factory=dict)
    percents: dict[tuple[str, str], str] = field(default_factory=dict)  # (подгруппа, конкурент)
    prices: dict | None = None  # {sheet, first, last, code, comp, price, use}


def write_reference_sheet(ws, rows: list[Row], settings: Settings, comps_by_key: dict, start_row: int = 1) -> tuple[Refs, int]:
    refs = Refs()
    name = q(ws.title)
    ws.sheet_view.showGridLines = True

    def block(start: int, title: str, headers: list[str], widths: list[int], data: list[list], fmts: list[str | None]) -> tuple[dict[int, str], int]:
        ws.cell(start, 1, title)
        head = start + 1
        for i, (h, w) in enumerate(zip(headers, widths)):
            ws.cell(head, i + 1, h)
            letter = get_column_letter(i + 1)
            ws.column_dimensions[letter].width = max(w, ws.column_dimensions[letter].width or 0)
        first = head + 1
        for n, values in enumerate(data):
            for i, v in enumerate(values):
                c = ws.cell(first + n, i + 1, v)
                if fmts[i]:
                    c.number_format = fmts[i]
        last_col = get_column_letter(len(headers))
        cells = {first + n: f"{name}!${last_col}${first + n}" for n in range(len(data))}
        return cells, first + max(len(data), 1) + 1

    row = start_row
    currencies = sorted({r.purchase_currency for r in rows if r.purchase_currency and r.purchase_price})
    cells, row = block(row, "Курсы", ["Валюта", "Курс, руб."], [18, 14],
                       [[c, settings.rate(c)] for c in currencies], [None, "0.000000"])
    refs.rates = {c: cells[n] for n, c in enumerate(currencies, start_row + 2)}

    gz_table = settings.gz
    groups = sorted({r.order_group for r in rows if r.order_group})
    cells, row = block(row, "ГЗ", ["Группа заказа", "Коэффициент"], [30, 14],
                       [[g, r_gz(rows, g, gz_table)] for g in groups], [None, COEF])
    refs.gz = {g: cells[n] for n, g in enumerate(groups, min(cells) if cells else row)}

    combos = sorted({(r.subgroup, r.order_group, r.brand) for r in rows if r.discount is not None})
    disc = {(r.subgroup, r.order_group, r.brand): r.discount for r in rows if r.discount is not None}
    cells, row = block(row, "Скидки", ["Подгруппа", "Группа заказа", "ТМ", "Скидка, %"], [30, 26, 14, 12],
                       [[*k, disc[k]] for k in combos], [None, None, None, "0.00"])
    refs.discounts = {k: cells[n] for n, k in enumerate(combos, min(cells) if cells else row)}

    pairs = sorted({(r.subgroup, k) for r in rows for k in r.percents if k in comps_by_key},
                   key=lambda p: (p[0], comps_by_key[p[1]].order))
    pct = {}
    for r in rows:
        for k, v in r.percents.items():
            pct.setdefault((r.subgroup, k), v)
    cells, row = block(row, "Позиционирование", ["Подгруппа", "Конкурент", "%"], [30, 16, 10],
                       [[s, comps_by_key[k].name, pct[(s, k)]] for s, k in pairs], [None, None, "0.0"])
    refs.percents = {p: cells[n] for n, p in enumerate(pairs, min(cells) if cells else row)}
    return refs, row


def r_gz(rows: list[Row], group: str, table: dict) -> float | None:
    return next((r.gz for r in rows if r.order_group == group and r.gz is not None), None)


@dataclass
class PriceLine:
    row: Row
    comp: object
    cross: str
    kind: str  # Да / Нет / АЦПКД
    offer: object
    use: bool


def price_lines(src: Sources, rows: list[Row], index) -> list[PriceLine]:
    """Все цены по кроссам товара: взятые в расчёт («да») и остальные кроссы этого конкурента («нет»)."""
    lines = []
    has_svod = bool(src.crosses)
    for r in rows:
        svod = svod_for(src, r)
        for comp in sorted(r.competitors, key=lambda c: c.order):
            if comp.key not in r.picks:
                continue
            picked = {norm_article(p.article) for p in r.picks[comp.key]}
            if has_svod:
                crosses = [(c.cross_article, c.cross_key, "Да" if c.yes else "Нет") for c in svod_comp(svod, comp.key)]
            else:
                crosses = [(c.article.upper(), c.article, "АЦПКД") for c in acp_crosses(r.acp, comp.key)]
            seen = set()
            for article, key, kind in crosses:
                if norm_article(article) in seen:
                    continue
                seen.add(norm_article(article))
                for o in sorted(index.get(comp.key, key), key=lambda o: o.price):
                    lines.append(PriceLine(r, comp, article, kind, o, norm_article(article) in picked))
    return lines


def write_prices_sheet(ws, lines: list[PriceLine], start_row: int = 1) -> tuple[dict, int]:
    headers = ["Кросс-артикул", "Код товара", "Наш артикул", "Подгруппа", "ABC", "Конкурент", "Кросс «Да»",
               "Поставщик / склад", "Источник", "Цена", "Наличие", "В расчёт", "Мин."]
    widths = [16, 12, 14, 26, 7, 13, 10, 36, 22, 12, 10, 10, 7]
    ws.cell(start_row, 1, "Цены конкурентов")
    header_row = start_row + 1
    for i, (h, w) in enumerate(zip(headers, widths), 1):
        ws.cell(header_row, i, h)
        letter = get_column_letter(i)
        ws.column_dimensions[letter].width = max(w, ws.column_dimensions[letter].width or 0)
    first = header_row + 1
    last = first + max(len(lines), 1) - 1
    rng = lambda col: f"${col}${first}:${col}${last}"  # noqa: E731
    for offset, ln in enumerate(lines):
        n = first + offset
        ws.cell(n, 1, ln.cross)
        ws.cell(n, 2, code_value(ln.row.code))
        ws.cell(n, 3, ln.row.article)
        ws.cell(n, 4, ln.row.subgroup)
        ws.cell(n, 5, ln.row.abc)
        ws.cell(n, 6, ln.comp.name)
        ws.cell(n, 7, ln.kind)
        ws.cell(n, 8, ln.offer.supplier)
        ws.cell(n, 9, ln.offer.source)
        price = ws.cell(n, 10, ln.offer.price)
        price.number_format = MONEY
        qty = ws.cell(n, 11, ln.offer.qty)
        qty.number_format = INT
        ws.cell(n, 12, "да" if ln.use else "нет")
        ws.cell(n, 13, f'=IF(L{n}<>"да","",IF(J{n}=_xlfn.AGGREGATE(15,6,{rng("J")}/(({rng("B")}=B{n})*({rng("F")}=F{n})*'
                f'({rng("L")}="да")),1),"да",""))')
    info = {"sheet": q(ws.title), "first": first, "last": last, "code": "B", "comp": "F", "price": "J", "use": "L"}
    return info, last + 2


def raw_table(ws, name: str, headers: list[str], data: list[list], widths: list[int],
              number_formats: dict[int, str] | None = None, inputs: set[int] | None = None, start_row: int = 1) -> int:
    """Кусок листа без оформления: название блока, шапка, строки. Возвращает строку следующего блока."""
    del inputs
    ws.sheet_view.showGridLines = True
    number_formats = number_formats or {}
    ws.cell(start_row, 1, name)
    head = start_row + 1
    for i, h in enumerate(headers, 1):
        ws.cell(head, i, h)
        if i - 1 < len(widths):
            letter = get_column_letter(i)
            ws.column_dimensions[letter].width = max(widths[i - 1], ws.column_dimensions[letter].width or 0)
    first = head + 1
    for offset, values in enumerate(data):
        for i, v in enumerate(values, 1):
            c = ws.cell(first + offset, i, v)
            if i in number_formats:
                c.number_format = number_formats[i]
    return first + max(len(data), 1) + 1


def plain_block(ws, title: str, headers: list[str], rows: list[list], start_row: int = 1, widths: list[int] | None = None) -> int:
    """Обычная таблица: название, шапка, строки. Возвращает строку, с которой можно писать следующий блок."""
    ws.sheet_view.showGridLines = True
    ws.cell(start_row, 1, title)
    head = start_row + 1
    for col, text in enumerate(headers, 1):
        ws.cell(head, col, text)
        if widths and col - 1 < len(widths):
            letter = get_column_letter(col)
            current = ws.column_dimensions[letter].width
            ws.column_dimensions[letter].width = max(widths[col - 1], current or 0)
    for i, values in enumerate(rows):
        for col, value in enumerate(values, 1):
            ws.cell(head + 1 + i, col, value)
    return head + max(len(rows), 1) + 2


def write_raw_sections(ws, src: Sources, rows: list[Row], start_row: int) -> int:
    codes = {r.code for r in rows}
    row = start_row
    data = [[p.article, code_value(p.code), p.name, p.brand, p.category, p.group, p.subgroup, p.order_group, p.variety,
             "да" if p.discontinued else "", p.overstock, p.abc, p.assortment, p.purchase_currency, p.purchase_price,
             p.mo_currency, p.mo_price, p.mo_rub] for p in (r.pst for r in rows) if p]
    row = raw_table(ws, "ПСТ", ["Артикул", "Код товара", "Наименование", "ТМ", "Категория", "Группа", "Подгруппа", "Группа заказа",
                                "Разновидность", "Вывод", "Оверсток", "ABC", "Тип ассорт.", "Закупка, валюта", "Закупка, цена",
                                "МО, валюта", "МО, цена", "МО, руб."], data,
                    [14, 12, 34, 12, 20, 16, 24, 22, 40, 8, 12, 7, 9, 10, 12, 9, 12, 12],
                    {15: MONEY, 17: MONEY, 18: MONEY}, start_row=row)

    by_code = {r.code: r for r in rows}
    if src.sbs:
        codes_sbs = [c for c in src.sbs if c in codes]
        first = row + 2
        data = []
        for i, code in enumerate(codes_sbs):
            qty, sbs = src.sbs[code]
            n = first + i
            data.append([by_code[code].article, code_value(code), qty, sbs * qty if sbs and qty else None,
                         f'=IF(AND(ISNUMBER(C{n}),ISNUMBER(D{n}),C{n}>0),D{n}/C{n},"")'])
        row = raw_table(ws, "СБС", ["Артикул", "Код товара", "Кол-во, кон. остаток", "Себестоимость, сумма", "СБС за шт"], data,
                        [14, 12, 14, 18, 14], {3: INT, 4: MONEY, 5: MONEY}, start_row=row)

    if src.sales:
        codes_sales = [c for c in {**src.sales, **src.sales_qty} if c in codes]
        first = row + 2
        data = []
        for i, code in enumerate(codes_sales):
            n = first + i
            data.append([by_code[code].article, code_value(code), src.sales_qty.get(code), src.sales.get(code),
                         f'=IF(AND(ISNUMBER(C{n}),ISNUMBER(D{n}),C{n}>0),D{n}/C{n},"")'])
        period = f" ({src.sales_period})" if src.sales_period else ""
        row = raw_table(ws, "Продажи", ["Артикул", "Код товара", f"Продажи{period}, шт", f"Продажи{period}, руб.", "Средняя цена"], data,
                        [14, 12, 14, 18, 14], {3: INT, 4: MONEY, 5: MONEY}, start_row=row)

    acp = [r.acp for r in rows if r.acp]
    if acp:
        data = [[a.article, code_value(a.code), a.name, a.subgroup, a.abc, a.assortment, a.mo_price, a.cat_a_price,
                 a.avg_cost, a.stock, a.sold_qty, a.avg_sale_price] for a in acp]
        row = raw_table(ws, "АЦПКД", ["Артикул", "Код товара", "Наименование", "Подгруппа", "ABC", "Тип ассорт.", "Цена МО",
                                      "Цена кат. А", "Средняя себест.", "Остаток", "Продано, шт", "Ср. цена продажи"], data,
                        [14, 12, 34, 24, 7, 9, 12, 12, 12, 10, 10, 12],
                        {7: MONEY, 8: MONEY, 9: MONEY, 10: INT, 11: INT, 12: MONEY}, start_row=row)

    if src.crosses:
        data = []
        for r in rows:
            for brand, crosses in svod_for(src, r).items():
                for c in crosses:
                    data.append([r.article, code_value(r.code), c.cross_brand or brand, c.cross_article, "Да" if c.yes else "Нет"])
        row = raw_table(ws, "Кроссы", ["Артикул", "Код товара", "Бренд кросса", "Кросс-артикул", "Использовать для сравнения"], data,
                        [14, 12, 16, 18, 14], start_row=row)
    return row


# ---------------------------------------------------------------- столбцы отчёта


@dataclass
class Column:
    title: str
    kind: str
    width: float
    value: object  # (row, n, L) -> значение или формула
    fmt: str | None = None
    key: str = ""
    outline: bool = False
    input: object = False  # bool или (row) -> bool: синий шрифт
    rule: str = ""  # dev | band | change | vr | bar
    tone: int | None = None  # чередование заливки у пары «цена / отклонение» конкурента
    edge: bool = False  # левая граница блока конкурента
    band: str = ""  # надпись над шапкой: ТОВАР, ПРОДАЖИ, НАША ЦЕНА, КОНКУРЕНТЫ, РЕШЕНИЕ


def short_reason(reason: str) -> str:
    """Причина в две-три слова для столбца «Причина»; полный текст — в свёрнутом «Подробно»."""
    text = (reason or "").strip()
    head = text.split(";")[0].strip()
    tail = text[len(head):]
    m = re.match(r"цену меняли (\d{2}\.\d{2})\.\d{4}.*не трогаем до (\d{2}\.\d{2})", text)
    if m:
        return f"меняли {m.group(1)}, ждём до {m.group(2)}"
    if head.startswith("нет продаж и нет остатка"):
        out = "нет продаж и остатка"
    elif head.startswith("нет продаж, а мы не дороже"):
        out = "нет продаж, цена не мешает"
    elif head.startswith("нет продаж, есть остаток"):
        who = head.rsplit("к цели", 1)[-1].strip()
        out = f"лежит, дороже {who} — снижаем".replace("  ", " ") if who else "лежит, дороже рынка — снижаем"
    elif head.startswith("нет продаж"):
        out = "нет продаж"
    elif head.startswith("в пределах позиционирования"):
        m = re.search(r"\(([^)]*)\)", head)
        out = f"у цели ({m.group(1)})" if m else "у цели"
    elif head.startswith("топ продаж, категория A"):
        out = "топ — не снижаем"
    elif head.startswith("снижение ограничено окупаемостью"):
        out = "снижать невыгодно"
    elif head.startswith("ВР ниже") and "поднимаем" in head:
        out = "мин. ВР — поднимаем"
    elif head.startswith("к цели"):
        out = "к " + head[len("к цели"):].split(",")[0].strip()
        if "ограничено" in head or "окупаемост" in tail:
            out += ", шаг ограничен"
    else:
        out = head
    if "не ниже ВР" in tail:
        out += ", мин. ВР"
    elif "вопрос к закупке" in tail:
        out += ", ВР низкая — к закупке"
    return out


def _oem_used(r) -> dict:
    return {k: v for k, v in (getattr(r, "oem", None) or {}).items() if v.get("used")}


def _oem_mark(r) -> str:
    """Цель взята по цене, найденной через ОЕМ, — так и пишем в причине."""
    used = _oem_used(r)
    if not used or not r.target_by:
        return ""
    names = {c.key: c.name for c in r.competitors}
    return " (по ОЕМ)" if any(names.get(k) == r.target_by for k in used) else ""


def _oem_notes(r) -> list[str]:
    from .report import oem_text

    names = {c.key: c.name for c in r.competitors}
    return [f"цена {names.get(k, k)} по ОЕМ: {oem_text(r, k)}" for k in _oem_used(r)]


def both(*cells: str) -> str:
    return "AND(" + ",".join(f"ISNUMBER({c})" for c in cells) + ")"


def report_columns(rows: list[Row], refs: Refs, opt: Options, period: str,
                   src: Sources | None = None, live: Live | None = None, settings: Settings | None = None) -> list[Column]:
    del period, refs
    ro = opt.recs
    live = live or Live()
    comps = competitor_order(rows)
    sales_title = "Продажи, шт"
    is_input = lambda r: True  # noqa: E731

    def key_cell(n, L) -> str:
        return f"${L['code']}{n}"

    def pull(kind: str, field: str, n, L) -> str | None:
        return live.pull(kind, field, key_cell(n, L))

    def linked(kind: str, field: str, n, L, plain):
        formula = pull(kind, field, n, L)
        return "=" + formula if formula else plain

    def rate_of(r) -> float | None:
        if not r.pst or not r.pst.purchase_currency or settings is None:
            return None
        return settings.rate(r.pst.purchase_currency)

    def fold(columns: list[Column]) -> list[Column]:
        for c in columns:
            c.outline = True
        return columns

    variety = Column("Разновидность", "info", 42, lambda r, n, L: linked("pst", "variety", n, L, r.variety) if r.pst and r.pst.variety else r.variety, key="variety")
    def pst_text(r, field: str):
        item = r.pst
        return bool(item and getattr(item, field))

    def mo_value(r, n, L):
        own = r.pst.mo_rub if r.pst else None
        acp = r.acp.mo_price if r.acp else None
        if took_own(opt.source("mo"), own, acp):
            return linked("pst", "mo_rub", n, L, r.mo)
        return r.mo

    def abc_value(r, n, L):
        own = r.pst.abc if r.pst else ""
        acp = r.acp.abc if r.acp else ""
        if took_own(opt.source("abc"), own, acp):
            return linked("pst", "abc", n, L, r.abc)
        return r.abc

    def stock_value(r, n, L):
        own = src.sbs[r.code][0] if src and r.code in src.sbs else None
        acp = r.acp.stock if r.acp else None
        if took_own(opt.source("stock"), own, acp):
            return linked("sbs", "qty", n, L, r.stock)
        return r.stock

    def sales_qty_value(r, n, L):
        loc = src.located("sales") if src else None
        has_qty = bool(loc and "qty" in loc.cols)
        own = src.sales_qty.get(r.code) if has_qty and r.code in src.sales_qty else None
        acp = r.acp.sold_qty if r.acp else None
        if has_qty and took_own(opt.source("sales_qty"), own, acp):
            return linked("sales", "qty", n, L, r.sales_qty)
        return r.sales_qty

    def sales_rub_value(r, n, L):
        own = src.sales.get(r.code) if src and src.located("sales") and r.code in src.sales else None
        acp = None
        if r.acp and r.acp.sold_qty is not None and r.acp.avg_sale_price:
            acp = r.acp.sold_qty * r.acp.avg_sale_price
        if r.acp and r.acp.sold_qty == 0:
            acp = 0.0
        if took_own(opt.source("sales"), own, acp):
            return linked("sales", "amount", n, L, r.sales)
        return r.sales

    def sbs_value(r, n, L):
        own = src.sbs[r.code][1] if src and r.code in src.sbs else None
        acp = r.acp.avg_cost if r.acp else None
        if took_own(opt.source("sbs"), own, acp):
            qty, cost = pull("sbs", "qty", n, L), pull("sbs", "cost", n, L)
            if qty and cost:
                return f'=IF(AND(ISNUMBER({qty}),ISNUMBER({cost}),{qty}>0),{cost}/{qty},"")'
        return r.sbs

    def buy_rub(r, n, L):
        rate = rate_of(r)
        if rate is None or "buy_cur" not in L:
            return r.purchase_rub
        cell = f'{L["buy_cur"]}{n}'
        return f'=IF(ISNUMBER({cell}),{cell}*{excel_num(rate)},"")'

    def sbs_plan(r, n, L):
        rate, gz = rate_of(r), r.gz
        if rate is None or gz is None or "buy_cur" not in L:
            return r.sbs_plan
        cell = f'{L["buy_cur"]}{n}'
        return f'=IF(ISNUMBER({cell}),{cell}*{excel_num(rate)}*{excel_num(gz)},"")'

    def subgroup_value(r, n, L):
        if pst_text(r, "subgroup"):
            return linked("pst", "subgroup", n, L, r.subgroup)
        return r.subgroup

    def cat_a(r, n, L):
        if r.cat_a_calc and r.discount is not None and r.mo is not None and "mo" in L:
            return f'=IF(ISNUMBER({L["mo"]}{n}),{L["mo"]}{n}*(1-{excel_num(r.discount)}/100),"")'
        return r.cat_a

    def sales_sum(price_key: str):
        def value(r, n, L):
            p = f'{L[price_key]}{n}'
            if opt.money_base == "potential":
                qty = L.get("potential") or L.get("stock")
                if qty:
                    return f'=IF(OR(NOT(ISNUMBER({p})),NOT(ISNUMBER({qty}{n}))),"",{qty}{n}*{p})'
            fallback = (f'IF({both(L["sales"] + str(n), p, L["cat_a"] + str(n))},{L["sales"]}{n}*{p}/{L["cat_a"]}{n},"")'
                        if "sales" in L else (f"{r.sales}*{p}/{L['cat_a']}{n}" if r.sales and r.cat_a and price_key != "cat_a"
                                              else (str(r.sales) if r.sales else '""')))
            if "sales" in L and price_key == "cat_a":
                fallback = f'IF(ISNUMBER({L["sales"]}{n}),{L["sales"]}{n},"")'
            return f'=IF(NOT(ISNUMBER({p})),"",IF(ISNUMBER({L["sales_qty"]}{n}),{L["sales_qty"]}{n}*{p},{fallback}))'
        return value

    def margin(price_key: str):
        """ВР от СБС. Если себестоимости нет, а СБС план есть — от плана."""
        def value(r, n, L):
            if price_key not in L or "sbs" not in L:
                return None
            price, sbs = f'{L[price_key]}{n}', f'{L["sbs"]}{n}'
            use = f'AND(ISNUMBER({sbs}),ISNUMBER({price}),{price}<>0),1-{sbs}/{price}'
            if "sbs_plan" in L:
                plan = f'{L["sbs_plan"]}{n}'
                return f'=IF({use},IF(AND(ISNUMBER({plan}),ISNUMBER({price}),{price}<>0),1-{plan}/{price},""))'
            if r.sbs_plan is not None:
                plan = excel_num(r.sbs_plan)
                return f'=IF({use},IF(AND(ISNUMBER({price}),{price}<>0),1-{plan}/{price},""))'
            return f'=IF({use},"")'
        return value

    def value_of(kind: str, field: str, plain_attr: str):
        return lambda r, n, L: linked(kind, field, n, L, getattr(r, plain_attr)) if pst_text(r, field) else getattr(r, plain_attr)

    # ---- ТОВАР: на виду код, артикул, разновидность и ABC; остальное свёрнуто под «+»
    item_fold = [
        Column("Наименование", "id", 34, lambda r, n, L: linked("pst", "name", n, L, r.name) if r.pst and r.pst.name else r.name, key="name"),
        Column("Группа заказа", "info", 20, value_of("pst", "order_group", "order_group")),
        Column("Категория NEW", "info", 22, value_of("pst", "category", "category")),
        Column("Группа NEW", "info", 16, value_of("pst", "group", "group")),
        Column("Подгруппа NEW", "info", 24, subgroup_value, key="subgroup"),
    ]
    if any(r.assortment for r in rows):
        item_fold.append(Column("Тип ассорт.", "info", 8, lambda r, n, L: linked("pst", "assortment", n, L, r.assortment)
                                if took_own(opt.source("assortment"), r.pst.assortment if r.pst else "", r.acp.assortment if r.acp else "")
                                else r.assortment))
    if any(r.overstock for r in rows):
        item_fold.append(Column("Оверсток", "info", 14, value_of("pst", "overstock", "overstock")))
    if any(r.discontinued for r in rows) and opt.filter("sale") != "exclude":
        def discontinued_value(r, n, L):
            formula = pull("pst", "discontinued", n, L) if r.pst else None
            if not formula:
                return "да" if r.discontinued else ""
            return f'=IF({formula}=1,"да","")'
        item_fold.append(Column("Вывод из ассорт.", "info", 9, discontinued_value))
    item_fold += [
        Column("Марки авто", "info", 28, lambda r, *_: r.car.label if r.car else ""),
        Column("Автопарк", "car", 16, lambda r, *_: r.car.fleet if r.car else ""),
        Column("Ценовой сегмент", "car", 12, lambda r, *_: r.car.segment if r.car else "", key="segment"),
    ]
    # Разновидность до свёрнутых столбцов и внутри закрепления. Иначе «+» стоит на ней,
    # и Excel 2016 при развороте группы внутри закреплённой области зависает.
    cols = [
        Column("Код", "id", 10, lambda r, *_: code_value(r.code), key="code"),
        Column("Артикул", "id", 14, lambda r, n, L: linked("pst", "article", n, L, r.article) if r.pst and r.pst.article else r.article, key="article"),
        variety,
    ] + fold(item_fold) + [Column("ABC", "calc", 6, abc_value, key="abc")]
    for c in cols:
        c.band = "ТОВАР"

    # ---- ПРОДАЖИ
    sales = [
        Column(sales_title, "calc", 10, sales_qty_value, INT, key="sales_qty", input=is_input, rule="bar"),
    ]
    if any(r.sales3 is not None for r in rows):
        # Продажи за последние месяцы — по ним правило «нет продаж — не поднимаем».
        qty3 = any(r.sales3_qty for r in rows)
        sales.append(Column("Продажи 3 мес, шт" if qty3 else "Продажи 3 мес, ₽", "calc", 10,
                            (lambda r, *_: r.sales3_qty) if qty3 else (lambda r, *_: r.sales3), INT, key="sales3"))
    sales += [
        Column("Продажи, ₽", "calc", 13, sales_rub_value, RUB, key="sales", input=is_input),
        Column("Остаток, шт", "calc", 9, stock_value, INT, key="stock", input=is_input),
    ]
    if any(r.in_transit is not None for r in rows):
        sales += [
            Column("В пути, шт", "calc", 9, lambda r, *_: r.in_transit, INT, key="transit", input=is_input),
            Column("Потенциал, шт", "calc", 11,
                   lambda r, n, L: f'=IF(COUNT({L["stock"]}{n},{L["transit"]}{n})=0,"",SUM({L["stock"]}{n},{L["transit"]}{n}))',
                   INT, key="potential"),
        ]
    for c in sales:
        c.band = "ПРОДАЖИ"

    # ---- НАША ЦЕНА: Кат. А и ВР на виду, из чего они складываются — под «+»
    price_fold = [
        Column("МО, ₽", "calc", 11, mo_value, RUB, key="mo", input=is_input),
        Column("Закупка, валюта", "buy", 9, value_of("pst", "purchase_currency", "purchase_currency"), key="cur"),
        Column("Закупка в валюте", "buy", 11, lambda r, n, L: linked("pst", "purchase_price", n, L, r.purchase_price) if r.pst and r.pst.purchase_price is not None else r.purchase_price, MONEY, key="buy_cur", input=is_input),
        Column("Закупка, ₽", "calc", 11, buy_rub, RUB, key="buy"),
        Column("СБС, ₽", "calc", 11, sbs_value, RUB, key="sbs", input=is_input),
        Column("СБС план, ₽", "calc", 11, sbs_plan, RUB, key="sbs_plan"),
        Column("Сумма потенциала, ₽" if opt.money_base == "potential" else "Сумма продажи, ₽", "calc", 14, sales_sum("cat_a"), RUB, key="sum_sale"),
        Column("Прибыль на потенциале, ₽" if opt.money_base == "potential" else "Сумма прибыли, ₽", "calc", 14,
               lambda r, n, L: f'=IF({both(L["sum_sale"] + str(n), L["vr"] + str(n))},{L["sum_sale"]}{n}*{L["vr"]}{n},"")', RUB, key="sum_profit"),
    ]
    ours = [
        Column("Кат. А, ₽", "calc", 11, cat_a, RUB, key="cat_a", input=lambda r: not r.cat_a_calc),
        Column("ВР", "calc", 7, margin("cat_a"), PCT, key="vr", rule="vr"),
    ] + fold(price_fold)
    for c in ours:
        c.band = "НАША ЦЕНА"
    cols += sales + ours

    # ---- КОНКУРЕНТЫ: пара «цена / откл.» на своей заливке, через одного
    market = []
    for i, c in enumerate(comps):
        tone = i % 2
        market.append(Column(f"{c.name}, ₽", "comp", 11, lambda r, *_ , comp=c: r.prices.get(comp.key), RUB,
                             key=f"price:{c.key}", tone=tone))
        market.append(Column("Откл.", "comp", 7, lambda r, n, L, k=c.key: f'=IF({both(L["price:" + k] + str(n), L["cat_a"] + str(n))},'
                             f'1-{L["cat_a"]}{n}/{L["price:" + k]}{n},"")', PCT_SIGNED, key=f"dev:{c.key}", rule="dev", tone=tone))
    sister = next((getattr(r, "sister_brand", "") for r in rows if getattr(r, "sister_brand", "")), "")
    if sister:
        from .sisters import sister_title
        market.append(Column(sister_title(sister), "sister", 14, lambda r, *_: r.sister_price, RUB, key="sister"))
    if any(r.target for r in rows):
        # Цель — снимок на момент формирования: цена конкурента × (1 + % позиционирования).
        market += [
            Column("Цель, ₽", "pos", 11, lambda r, *_: round(r.target, 2) if r.target else None, RUB, key="target"),
            Column("Откл. от цели", "pos", 8, lambda r, n, L: f'=IF({both(L["target"] + str(n), L["cat_a"] + str(n))},'
                   f'1-{L["cat_a"]}{n}/{L["target"]}{n},"")', PCT_SIGNED, key="gap", rule="band"),
        ] + fold([Column("По кому цель", "pos", 12, lambda r, *_: r.target_by if r.target else "", key="target_by")])
    for c in market:
        c.band = "КОНКУРЕНТЫ"
    cols += market

    # ---- РЕШЕНИЕ
    if ro.enabled:
        def rec_value(r, *_):
            return r.rec.price if r.rec else None

        decide = [
            Column("Новая цена, ₽", "rec", 11, rec_value, RUB, key="rec", input=is_input),
            Column("Изм.", "rec", 7, lambda r, n, L: f'=IF({both(L["rec"] + str(n), L["cat_a"] + str(n))},'
                   f'{L["rec"]}{n}/{L["cat_a"]}{n}-1,"")', CHANGE, key="change", rule="change"),
        ]
        if opt.reason_cols in ("all", "short"):
            decide.append(Column("Причина", "rec", 30, lambda r, *_: (short_reason(r.rec.reason) + _oem_mark(r)) if r.rec else "", key="reason"))
        extra = []
        if opt.reason_cols == "all":
            extra.append(Column("Подробно", "rec", 60, lambda r, *_: "; ".join(x for x in [r.rec.reason if r.rec else "", *_oem_notes(r)] if x), key="reason_full"))
        if any(getattr(r, "last_change", None) for r in rows):
            from .pricehistory import change_text

            extra.append(Column("Цену меняли", "info", 26, lambda r, *_: change_text(getattr(r, "last_change", None)), key="changed"))
        extra.append(Column("Новое МО, ₽", "rec", 11, lambda r, n, L: f'=IF(ISNUMBER({L["rec"]}{n}),{L["rec"]}{n}/(1-{excel_num(r.discount)}/100),"")'
                            if r.discount is not None else None, RUB, key="rec_mo"))
        extra.append(Column("ВР после", "rec", 7, margin("rec"), PCT, key="vr_rec", rule="vr"))
        extra += [
            Column("Сумма потенциала после, ₽" if opt.money_base == "potential" else "Сумма продажи после, ₽", "rec", 14, sales_sum("rec"), RUB, key="sum_sale_rec"),
            Column("Прибыль на потенциале после, ₽" if opt.money_base == "potential" else "Сумма прибыли после, ₽", "rec", 14,
                   lambda r, n, L: f'=IF({both(L["sum_sale_rec"] + str(n), L["vr_rec"] + str(n))},{L["sum_sale_rec"]}{n}*{L["vr_rec"]}{n},"")',
                   RUB, key="sum_profit_rec"),
        ]
        decide += fold(extra)
        for c in decide:
            c.band = "РЕШЕНИЕ"
        cols += decide
    elif any(getattr(r, "last_change", None) for r in rows):
        from .pricehistory import change_text

        cols.append(Column("Цену меняли", "info", 26, lambda r, *_: change_text(getattr(r, "last_change", None)), key="changed", band="РЕШЕНИЕ"))
    return cols


# ---------------------------------------------------------------- лист отчёта


@dataclass
class SheetInfo:
    title: str
    letters: dict[str, str]
    first: int
    last: int
    rows: list[Row]

    def rng(self, key: str) -> str:
        col = self.letters[key]
        return f"{q(self.title)}!${col}${self.first}:${col}${self.last}"


def write_report_sheet(ws, title: str, rows: list[Row], refs: Refs, opt: Options, period: str,
                       src: Sources | None = None, live: Live | None = None, settings: Settings | None = None) -> SheetInfo:
    """Лист «Товары»: название, строка итогов по фильтру, надписи блоков, шапка, товары.

    Оформление как у таблицы, сделанной руками: серая шапка, тонкие серые линии, заливкой выделены
    только наша цена, новая цена и конкуренты (через одного), отклонения — цветом текста.
    """
    cols = report_columns(rows, refs, opt, period, src, live, settings)
    L = {c.key: get_column_letter(i) for i, c in enumerate(cols, 1) if c.key}
    last = max(FIRST, FIRST + len(rows) - 1)
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90

    ink = Font(name=style.FONT, color=style.INK)
    bold = Font(name=style.FONT, bold=True, color=style.INK)
    ws.cell(TITLE_ROW, 1, title).font = Font(name=style.FONT, size=14, bold=True, color=style.INK)
    ws.row_dimensions[TITLE_ROW].height = 24

    # ---- итоги по тому, что видно после фильтра. 109 — без строк, которые фильтр скрыл
    tot = lambda key: f"SUBTOTAL(109,{L[key]}{FIRST}:{L[key]}{last})"  # noqa: E731
    t = TOTAL_ROW
    total_fill = PatternFill("solid", fgColor=style.XL_TOTAL)
    for i in range(1, len(cols) + 1):
        c = ws.cell(t, i)
        c.fill, c.border = total_fill, style.XL_CELL
    label = ws.cell(t, 1, f'="Итого (товаров: "&SUBTOTAL(103,{L["code"]}{FIRST}:{L["code"]}{last})&")"')
    label.font = bold
    total_cells = [
        ("vr", f'=IFERROR({L["sum_profit"]}{t}/{L["sum_sale"]}{t},"")', PCT),
        ("sum_sale", "=" + tot("sum_sale"), RUB),
        ("sum_profit", "=" + tot("sum_profit"), RUB),
        ("sales_qty", "=" + tot("sales_qty"), INT),
        ("stock", "=" + tot("stock"), INT),
    ]
    for key in ("sales", "sales3", "transit", "potential"):
        if key in L:
            total_cells.append((key, "=" + tot(key), RUB if key == "sales" else INT))
    if "rec" in L:
        total_cells += [
            ("vr_rec", f'=IFERROR({L["sum_profit_rec"]}{t}/{L["sum_sale_rec"]}{t},"")', PCT),
            ("sum_sale_rec", "=" + tot("sum_sale_rec"), RUB),
            ("sum_profit_rec", "=" + tot("sum_profit_rec"), RUB),
            ("change", f'=IFERROR({L["sum_sale_rec"]}{t}/{L["sum_sale"]}{t}-1,"")', PCT_SIGNED),
        ]
    for key, formula, fmt in total_cells:
        c = ws[f"{L[key]}{t}"]
        c.value, c.number_format, c.font = formula, fmt, bold

    # ---- надписи блоков над шапкой: «по центру выделения», без объединения ячеек
    band_font = Font(name=style.FONT, bold=True, size=9, color=style.XL_BAND)
    band_line = Border(bottom=Side(style="medium", color=argb("7F7F7F")))
    start = 0
    for i in range(1, len(cols) + 2):
        if i <= len(cols) and cols[i - 1].band == cols[start].band:
            continue
        band = cols[start].band
        if band:
            for j in range(start + 1, i):
                c = ws.cell(BAND_ROW, j)
                c.alignment = Alignment(horizontal="centerContinuous")
                c.border = band_line
                c.font = band_font
            ws.cell(BAND_ROW, start + 1, band)
        start = i - 1
    ws.row_dimensions[BAND_ROW].height = 15

    head_fill = PatternFill("solid", fgColor=style.XL_HEAD)
    for i, c in enumerate(cols, 1):
        cell = ws.cell(HEADER_ROW, i, c.title)
        cell.font = bold
        cell.fill = head_fill
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = style.XL_HEAD_BORDER
        dim = ws.column_dimensions[get_column_letter(i)]
        dim.width = c.width
        if c.outline:
            dim.outlineLevel = 1
            dim.hidden = True
    ws.row_dimensions[HEADER_ROW].height = 32

    fills = {
        "cat_a": PatternFill("solid", fgColor=style.XL_CAT_A),
        "rec": PatternFill("solid", fgColor=style.XL_NEW),
        "sister": PatternFill("solid", fgColor=style.XL_SISTER),
    }
    tones = [PatternFill("solid", fgColor=x) for x in style.XL_COMP]
    for key, fill in fills.items():
        if key in L:
            ws[f"{L[key]}{HEADER_ROW}"].fill = fill
    for i, c in enumerate(cols, 1):
        if c.tone is not None:
            ws.cell(HEADER_ROW, i).fill = tones[c.tone]
    small = Font(name=style.FONT, color=style.XL_MUTED, size=9)
    for n, r in enumerate(rows, FIRST):
        for i, c in enumerate(cols, 1):
            value = c.value(r, n, L)
            cell = ws.cell(n, i, value if value != "" else None)
            is_formula = isinstance(value, str) and value.startswith("=")
            inp = c.input(r) if callable(c.input) else c.input
            cell.font = INPUT if inp and not is_formula and value not in (None, "") else ink
            cell.border = style.XL_CELL
            if c.fmt:
                cell.number_format = c.fmt
            if c.key in fills:
                cell.fill = fills[c.key]
            elif c.tone is not None:
                cell.fill = tones[c.tone]
            if c.key == "article":
                cell.font = Font(name=style.FONT, bold=True, color=style.INK)
            elif c.key.startswith("price:") and c.key[6:] in _oem_used(r):
                cell.font = Font(name=style.FONT, italic=True, color=style.INK)  # цена найдена через ОЕМ
            elif c.key == "reason_full":
                cell.font = small

    if rows:
        add_rules(ws, cols, L, last, opt)
    anchor = next((i for i, c in enumerate(cols, 1) if c.key == "variety"), 3)
    freeze = f"{get_column_letter(anchor + 1)}{FIRST}"
    ws.freeze_panes = freeze
    # openpyxl ставит активную ячейку A1 во все области. Для правой нижней это клетка вне панели.
    freeze_row, freeze_col = coordinate_to_tuple(freeze)
    for sel in ws.sheet_view.selection:
        if sel.pane == "bottomRight":
            cell = freeze
        elif sel.pane == "topRight":
            cell = f"{get_column_letter(freeze_col)}1"
        elif sel.pane == "bottomLeft":
            cell = f"A{freeze_row}"
        else:
            continue
        sel.activeCell = cell
        sel.sqref = cell
    # sheetFormatPr пишется раньше, чем openpyxl узнает уровень группировки столбцов.
    if any(getattr(dim, "outlineLevel", None) for dim in ws.column_dimensions.values()):
        ws.column_dimensions.to_tree()
    last_col = get_column_letter(len(cols))
    ws.auto_filter.ref = f"A{HEADER_ROW}:{last_col}{last}"
    ws.sheet_properties.outlinePr.summaryRight = True
    style.print_wide(ws, f"{BAND_ROW}:{HEADER_ROW}")
    return SheetInfo(ws.title, L, FIRST, last, rows)


def add_rules(ws, cols: list[Column], L: dict, last: int, opt: Options) -> None:
    """Условное форматирование — только цвет текста: заливка в отчёте у столбцов, а не у отклонений."""
    ro = opt.recs
    tol = ro.tolerance / 100
    green = Font(color=style.XL_GREEN)
    red = Font(color=style.XL_RED)
    above, below = (green, red) if opt.green_when == "above" else (red, green)
    for c in cols:
        col = L.get(c.key)
        if not col:
            continue
        rng = f"{col}{FIRST}:{col}{last}"
        x = f"{col}{FIRST}"
        if c.rule == "dev":
            # Откл = 1 − наша / их: < 0 — мы дороже конкурента
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({x}),{x}<0)"], font=above))
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({x}),{x}>0)"], font=below))
        elif c.rule == "band":
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({x}),{x}<-{tol})"], font=above))
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({x}),{x}>{tol})"], font=below))
        elif c.rule == "change":
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({x}),{x}>0.0005)"], font=Font(bold=True)))
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({x}),{x}<-0.0005)"], font=Font(color=style.XL_RED, bold=True)))
        elif c.rule == "vr":
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({x}),{x}<{ro.min_vr / 100})"], font=Font(color=style.XL_RED, bold=True)))


def _flag_xf(tag: str) -> str:
    """Excel 2016 применяет шрифт и заливку, только если у стиля стоят applyFont / applyFill.

    Более новый Excel догадывается сам. Без этих флагов лист открывается как обычный текст.
    """
    def num(name: str) -> int:
        found = re.search(rf'{name}="(\d+)"', tag)
        return int(found.group(1)) if found else 0

    extra = []
    if num("fontId") and "applyFont=" not in tag:
        extra.append('applyFont="1"')
    if num("fillId") and "applyFill=" not in tag:
        extra.append('applyFill="1"')
    if num("borderId") and "applyBorder=" not in tag:
        extra.append('applyBorder="1"')
    if num("numFmtId") and "applyNumberFormat=" not in tag:
        extra.append('applyNumberFormat="1"')
    if not extra:
        return tag
    suffix = "/>" if tag.endswith("/>") else ">"
    return tag[: -len(suffix)].rstrip() + " " + " ".join(extra) + suffix


def _cache_show_formulas(blob: bytes, shared: list[str] | None = None) -> bytes:
    """Записать уже посчитанное «да» в скрытый столбец, чтобы фильтр сработал при открытии.

    Иначе Excel видит пустую формулу и прячет все строки или не прячет нужные.
    """
    text = blob.decode("utf-8")
    if '="Скрыть"' not in text:
        return blob
    # [^>/]* не захватывает самозакрывающийся <c ... />: иначе «да» попадает в числовую ячейку.
    cell_re = re.compile(r'<c r="([A-Z]+)(\d+)"([^>/]*)>(.*?)</c>', re.S)
    row1 = {}
    for match in cell_re.finditer(text):
        if match.group(2) != "1":
            continue
        inner = match.group(4)
        found = re.search(r"<t[^>]*>([^<]*)</t>", inner)
        if found:
            row1[match.group(1)] = found.group(1)
            continue
        attrs = match.group(3)
        if shared and 't="s"' in attrs:
            index = re.search(r"<v>(\d+)</v>", inner)
            if index and int(index.group(1)) < len(shared):
                row1[match.group(1)] = shared[int(index.group(1))]

    def shown(formula: str) -> bool:
        hide = False
        for part in re.finditer(r'AND\(\$([A-Z]+)\$1="Скрыть",(NOT\((\d+)\)|(\d+))\)', formula):
            choice = row1.get(part.group(1), "")
            if part.group(3) is not None:
                flag = int(part.group(3)) == 0
            else:
                flag = int(part.group(4)) != 0
            if choice == "Скрыть" and flag:
                hide = True
        return not hide

    def repl(match):
        inner = match.group(4)
        if "<f>" not in inner or '="Скрыть"' not in inner or "<v>" in inner:
            return match.group(0)
        formula = re.search(r"<f[^>]*>(.*?)</f>", inner, re.S)
        if not formula:
            return match.group(0)
        value = "да" if shown(formula.group(1)) else "нет"
        attrs = match.group(3)
        if 't="' not in attrs:
            attrs += ' t="str"'
        inner = inner.replace("</f>", f"</f><v>{value}</v>", 1)
        return f'<c r="{match.group(1)}{match.group(2)}"{attrs}>{inner}</c>'

    return cell_re.sub(repl, text).encode("utf-8")


# Excel читает <col> не как обычный XML: min и max должны идти первыми, outlineLevel — после них.
_COL_ATTR_ORDER = (
    "min", "max", "width", "style", "hidden", "bestFit", "customWidth", "phonetic", "outlineLevel", "collapsed",
)


def _repair_sheet_xml(text: str) -> str:
    """Разметка, из‑за которой Excel заменяет лист целиком: «ошибка XML, строка 1, столбец 0».

    Проверка данных, пустое значение формулы, группировка столбца без outlineLevelCol
    и активная ячейка A1 вне закреплённой области. Весь лист — одна строка, поэтому
    номер столбца в сообщении Excel всегда 0.
    """
    text = re.sub(r"<v\s*/>", "", text)
    text = text.replace("<v></v>", "")
    # Встроенный список «Показать,Скрыть» Excel считает битым XML и удаляет весь лист.
    # Список с ячеек листа «Списки» оставляем.
    if '"Показать,Скрыть"' in text or '"да,нет"' in text:
        text = re.sub(r"<dataValidations\b[^>]*/>", "", text)
        text = re.sub(r"<dataValidations\b.*?</dataValidations>", "", text, flags=re.S)
    text = _reorder_col_attrs(text)
    text = _set_outline_level_col(text)
    text = _fix_frozen_selection(text)
    text = text.replace("</row><row ", "</row>\n<row ")
    if not text.lstrip().startswith("<?xml"):
        text = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + text
    return text


def _reorder_col_attrs(text: str) -> str:
    def repl(match: re.Match) -> str:
        attrs = dict(re.findall(r'([\w:]+)="([^"]*)"', match.group(1)))
        ordered = [f'{key}="{attrs.pop(key)}"' for key in _COL_ATTR_ORDER if key in attrs]
        ordered += [f'{key}="{value}"' for key, value in attrs.items()]
        return "<col " + " ".join(ordered) + "/>"

    return re.sub(r"<col\b([^>]*)/>", repl, text)


def _set_outline_level_col(text: str) -> str:
    levels = [int(value) for value in re.findall(r'<col\b[^>]*\boutlineLevel="(\d+)"', text)]
    if not levels:
        return text
    level = str(max(levels))

    def repl(match: re.Match) -> str:
        body = match.group(1).rstrip()
        if "outlineLevelCol=" in body:
            body = re.sub(r'outlineLevelCol="\d+"', f'outlineLevelCol="{level}"', body)
        else:
            body += f' outlineLevelCol="{level}"'
        return "<sheetFormatPr" + body + "/>"

    return re.sub(r"<sheetFormatPr\b([^>]*?)\s*/?>", repl, text, count=1)


def _fix_frozen_selection(text: str) -> str:
    match = re.search(r"<pane\b([^>]*)/>", text)
    if not match:
        return text
    attrs = dict(re.findall(r'([\w:]+)="([^"]*)"', match.group(1)))
    top = attrs.get("topLeftCell")
    if attrs.get("state") != "frozen" or not top:
        return text
    ref = re.fullmatch(r"([A-Z]+)(\d+)", top)
    if not ref:
        return text
    col, row = ref.group(1), ref.group(2)
    has_x = attrs.get("xSplit") not in (None, "", "0")
    has_y = attrs.get("ySplit") not in (None, "", "0")
    parts = []
    if has_x and has_y:
        parts = [
            f'<selection pane="topRight" activeCell="{col}1" sqref="{col}1"/>',
            f'<selection pane="bottomLeft" activeCell="A{row}" sqref="A{row}"/>',
            f'<selection pane="bottomRight" activeCell="{top}" sqref="{top}"/>',
        ]
    elif has_y:
        parts = [f'<selection pane="bottomLeft" activeCell="A{row}" sqref="A{row}"/>']
    elif has_x:
        parts = [f'<selection pane="topRight" activeCell="{col}1" sqref="{col}1"/>']
    text = re.sub(r"<selection\b[^>]*/>", "", text)
    text = re.sub(r"<selection\b[^>]*>.*?</selection>", "", text, flags=re.S)
    return text.replace(match.group(0), match.group(0) + "".join(parts), 1)


def _hide_filtered_rows(text: str) -> str:
    """Скрыть строки, которые не прошли фильтр над таблицей. Без готового автофильтра."""
    hidden: set[str] = set()
    for match in re.finditer(r'<c r="[A-Z]+(\d+)"[^>]*>(.*?)</c>', text, flags=re.S):
        inner = match.group(2)
        if "<f>" not in inner or '="Скрыть"' not in inner:
            continue
        if "<v>да</v>" not in inner:
            hidden.add(match.group(1))

    def repl(match: re.Match) -> str:
        if match.group(1) not in hidden or "hidden=" in match.group(2):
            return match.group(0)
        return f'<row r="{match.group(1)}"{match.group(2)} hidden="1">'

    return re.sub(r'<row r="(\d+)"([^>]*)>', repl, text)


def _drop_filter_banner(text: str) -> str:
    """Строка «Показать / Скрыть» над таблицей. Excel из‑за неё удаляет лист «Отчет»."""
    def repl(match: re.Match) -> str:
        row = match.group(0)
        if "Скрыть" in row or "Показать" in row:
            return ""
        return row

    return re.sub(r'<row r="1"\b[^>]*>.*?</row>', repl, text, count=1, flags=re.S)


def _drop_preset_filter(text: str) -> str:
    text = re.sub(r"<filterColumn\b.*?</filterColumn>", "", text, flags=re.S)
    return re.sub(r"(<autoFilter\b[^>]*)>\s*</autoFilter>", r"\1/>", text)


# Порядок детей <font>, которого требует схема Excel. Иначе стили битые.
_FONT_ORDER = (
    "b", "i", "strike", "condense", "extend", "outline", "shadow", "u", "vertAlign",
    "sz", "color", "name", "family", "charset", "scheme",
)


def _reorder_fonts(text: str) -> str:
    def repl(match: re.Match) -> str:
        tags = re.findall(r"<[a-zA-Z]+\b[^>]*/>", match.group(1))
        if not tags:
            return match.group(0)

        def key(tag: str) -> int:
            name = re.match(r"<([a-zA-Z]+)", tag).group(1)
            return _FONT_ORDER.index(name) if name in _FONT_ORDER else len(_FONT_ORDER)

        return "<font>" + "".join(sorted(tags, key=key)) + "</font>"

    return re.sub(r"<font>(.*?)</font>", repl, text, flags=re.S)


def _sanitize_sheet(blob: bytes) -> bytes:
    """Привести лист к виду, который Excel открывает без замены sheet.xml."""
    return _repair_sheet_xml(blob.decode("utf-8")).encode("utf-8")


def _sheet_parts(blobs: dict) -> dict[str, str]:
    """Имя листа -> файл листа внутри книги."""
    import posixpath

    book = blobs.get("xl/workbook.xml", b"").decode("utf-8")
    rels = blobs.get("xl/_rels/workbook.xml.rels", b"").decode("utf-8")
    targets = {}
    for m in re.finditer(r"<Relationship\b[^>]*>", rels):
        tag = m.group(0)
        rid = re.search(r'Id="([^"]+)"', tag)
        target = re.search(r'Target="([^"]+)"', tag)
        if rid and target:
            t = target.group(1)
            targets[rid.group(1)] = t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join("xl", t))
    out = {}
    for m in re.finditer(r"<sheet\b[^>]*>", book):
        tag = m.group(0)
        name = re.search(r'name="([^"]*)"', tag)
        rid = re.search(r'r:id="([^"]+)"|\bid="([^"]+)"', tag)
        if name and rid:
            from xml.sax.saxutils import unescape

            out[unescape(name.group(1), {"&quot;": '"', "&apos;": "'"})] = targets.get(rid.group(1) or rid.group(2), "")
    return out


def _excel2016_stylesheet(path: Path, raw_sheets: dict | None = None) -> None:
    """Дописать флаги стилей, без которых Excel 2016 не показывает оформление.

    raw_sheets — листы-выгрузки, ячейки которых вставляются готовым XML (write_raw_sheet).
    """
    import zipfile
    from io import BytesIO

    src = zipfile.ZipFile(path)
    blobs = {item.filename: src.read(item.filename) for item in src.infolist()}
    if raw_sheets:
        parts = _sheet_parts(blobs)
        for title, tpl in raw_sheets.items():
            part = parts.get(title)
            if not part or part not in blobs:
                raise ValueError(f"не найден лист «{title}» в книге")
            blobs[part] = _inject_raw(blobs[part].decode("utf-8"), tpl).encode("utf-8")
    shared_xml = blobs.get("xl/sharedStrings.xml", b"").decode("utf-8")
    shared = re.findall(r"<t[^>]*>([^<]*)</t>", shared_xml)
    out = BytesIO()
    with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            blob = blobs[item.filename]
            if item.filename == "xl/styles.xml":
                text = blob.decode("utf-8")
                for piece in ("<colors />", "<colors/>", "<colors></colors>"):
                    text = text.replace(piece, "")
                text = re.sub(r"<xf\b[^>]*?/?>", lambda m: _flag_xf(m.group(0)), text)
                text = _reorder_fonts(text)
                blob = text.encode("utf-8")
            elif item.filename == "xl/workbook.xml":
                text = blob.decode("utf-8").replace("<workbookProtection />", "").replace("<workbookProtection/>", "")
                blob = text.encode("utf-8")
            elif item.filename.startswith("xl/worksheets/sheet") and item.filename.endswith(".xml"):
                blob = _sanitize_sheet(_cache_show_formulas(blob, shared))
            dst.writestr(item.filename, blob)
    path.write_bytes(out.getvalue())


SLICE_ROOT = "Распродажа и оверсток"
SLICE_DIR = {"sale": "Распродажа", "overstock": "Оверсток"}


def slice_output(out_dir: Path, opt: Options) -> Path | None:
    """Папка режима распродажи или оверстока. Основной отчёт её не использует."""
    name = SLICE_DIR.get(opt.stock_slice)
    if not name:
        return None
    return out_dir / SLICE_ROOT / name


def _report_path(out_dir: Path, opt: Options, priced: int) -> Path:
    """Имя файла — торговая марка и группа. Папка — целевая по объёму, мелкая или без цен."""
    base = slice_output(out_dir, opt) or out_dir
    folder = base / result_folder_name(result_bucket(opt.brand, opt.group, priced))
    folder.mkdir(parents=True, exist_ok=True)
    safe = re.sub(r'[<>:"/\\|?*]+', " ", f"{opt.brand} {opt.group}")
    safe = re.sub(r"\s+", " ", safe).strip() or "Отчет"
    return folder / f"{safe}.xlsx"


def sheet_title(name: str, used: set[str]) -> str:
    base = re.sub(r"[\[\]:*?/\\']", " ", name).strip()[:31] or "Отчет"
    title, i = base, 2
    while title.lower() in used:
        suffix = f" ({i})"
        title, i = base[: 31 - len(suffix)] + suffix, i + 1
    used.add(title.lower())
    return title


# ---------------------------------------------------------------- сборка книги


@dataclass
class BuildResult:
    path: Path
    rows: int
    sheets: list[str]
    excluded: int
    priced: int
    kind: str
    brief: Path
    prep: object = None  # разбор группы — для истории показателей


def build_report(src: Sources, settings: Settings, opt: Options, out_dir: Path,
                 price_files: list[PriceFile] | None = None) -> BuildResult:
    prep = prepare(src, settings, opt, price_files)
    kept = prep.kept
    if not kept:
        raise ValueError("После фильтров не осталось ни одного товара — ослабьте фильтры.")
    sheet_rows = rows_for_workbook(prep, src, opt)

    wb = style.new_workbook()
    used: set[str] = set()
    live = Live()
    export_sheets = []
    raw_sheets: dict[str, RawTemplate] = {}
    for kind, sheet_name in (("pst", "ПСТ"), ("sbs", "СБС"), ("sales", "Продажи")):
        loc = src.located(kind)
        if not loc:
            continue
        export = wb.create_sheet(sheet_title(sheet_name, used))
        tpl = write_raw_sheet(export, loc.header.sheet.grid)
        if tpl is not None:
            raw_sheets[export.title] = tpl
        setattr(live, kind, loc)
        live.names[kind] = export.title
        export_sheets.append(export)

    refs = Refs()
    title = " · ".join(x for x in (opt.brand, opt.group) if x) or "Итоговый отчёт"
    subgroups = list(dict.fromkeys(r.subgroup for r in kept))
    infos: list[SheetInfo] = []
    if opt.subgroup_mode == "sheets" and len(subgroups) > 1:
        for sg in subgroups:
            ws = wb.create_sheet(sheet_title(sg or "Без подгруппы", used))
            part = [r for r in sheet_rows if r.subgroup == sg]
            infos.append(write_report_sheet(ws, f"{title} · {sg}", part, refs, opt, src.sales_period, src, live, settings))
    else:
        ws = wb.create_sheet(sheet_title("Товары", used))
        head = title if len(subgroups) > 1 else f"{title} · {subgroups[0]}"
        infos.append(write_report_sheet(ws, head, sheet_rows, refs, opt, src.sales_period, src, live, settings))

    crosses = wb.create_sheet(sheet_title("Кроссы", used))
    write_cross_prices(crosses, sheet_rows, opt.price_pick)

    from .dashboard import summary as dash_summary
    from .itogi import write_prices_overview, write_summary_sheet

    itogi = wb.create_sheet(sheet_title("Сводка", used))
    chart_ws = wb.create_sheet(sheet_title("Графики", used))
    write_summary_sheet(itogi, kept, infos, opt, src, title, dash_summary(kept, opt), chart_ws, getattr(prep, "signals", None))
    price_sheets = []
    if opt.subgroup_mode == "sheets" and len(subgroups) > 1:
        # отчёт разложен по подгруппам — и цены конкурентов тоже: лист «Цены …» на подгруппу
        for sg in subgroups:
            ws_p = wb.create_sheet(sheet_title(f"Цены {sg or 'без подгруппы'}", used))
            write_prices_overview(ws_p, [r for r in sheet_rows if r.subgroup == sg], infos, opt)
            price_sheets.append(ws_p)
    else:
        ws_p = wb.create_sheet(sheet_title("Цены конкурентов", used))
        write_prices_overview(ws_p, sheet_rows, infos, opt)
        price_sheets.append(ws_p)

    order = [itogi] + [wb[i.title] for i in infos] + price_sheets + [crosses, chart_ws] + export_sheets
    wb._sheets = order  # noqa: SLF001
    report_titles = {i.title for i in infos} | {itogi.title} | {w.title for w in price_sheets}
    for ws in wb.worksheets:
        if ws.title in report_titles:
            ws.sheet_properties.tabColor = None  # ярлыки без цвета, как у обычной книги
            ws.sheet_state = "visible"
        else:
            ws.sheet_state = "hidden"
    wb.active = 0
    # Палитра indexedColors от openpyxl заставляет Excel 2016 чинить styles.xml и снимать оформление.
    wb._colors = []

    priced = sum(r.any_price for r in kept)
    path = _report_path(out_dir, opt, priced)
    wb.save(path)
    from .brief import write_group_brief
    brief_dir = slice_output(out_dir, opt) or out_dir
    brief = write_group_brief(brief_dir, kept, opt, src)
    _excel2016_stylesheet(path, raw_sheets)
    return BuildResult(path, len(kept), [ws.title for ws in wb.worksheets], len(prep.dropped),
                       priced, opt.report_kind, brief, prep)


PICK_HEAD = {"min": "Мин", "median": "Медиана", "second": "2-я снизу"}


def write_cross_prices(ws, rows: list[Row], how: str = "min") -> None:
    """Выбранные кросс-артикулы и цены поставщиков. Шапка в два этажа, без объединения ячеек.

    В отчёте цена конкурента остаётся числом и на этот лист не ссылается.
    """
    ws.sheet_view.showGridLines = True
    comps = competitor_order(rows)
    suppliers: dict[str, list[str]] = {}
    for comp in comps:
        names: list[str] = []
        seen = set()
        for r in rows:
            for pick in r.picks.get(comp.key) or []:
                for offer in pick.offers:
                    if offer.supplier and offer.supplier not in seen:
                        seen.add(offer.supplier)
                        names.append(offer.supplier)
        suppliers[comp.key] = names

    # строка 1 — бренд над первым столбцом своей группы, без объединения
    # строка 2 — кросс-артикул, поставщики, минимум
    ws.cell(2, 1, "Артикул")
    ws.column_dimensions["A"].width = 18
    col = 2
    groups = []
    for comp in comps:
        brand_col = col
        ws.cell(1, brand_col, comp.name)
        ws.cell(2, col, "Кросс-артикул")
        ws.column_dimensions[get_column_letter(col)].width = 24
        col += 1
        price_cols = []
        for name in suppliers[comp.key]:
            ws.cell(2, col, name)
            dim = ws.column_dimensions[get_column_letter(col)]
            dim.width = 14
            # Поставщики одного конкурента свёрнуты: на виду кросс и минимум.
            dim.outlineLevel = 1
            dim.hidden = True
            price_cols.append(col)
            col += 1
        ws.cell(2, col, PICK_HEAD.get(how, "Мин"))
        ws.column_dimensions[get_column_letter(col)].width = 12
        groups.append((comp.key, price_cols, col))
        col += 1

    for n, r in enumerate(rows, 3):
        ws.cell(n, 1, r.article)
        for key, price_cols, min_col in groups:
            picks = r.picks.get(key) or []
            articles = []
            by_supplier: dict[str, float] = {}
            for pick in picks:
                if pick.article and pick.article not in articles:
                    articles.append(pick.article)
                for offer in pick.offers:
                    current = by_supplier.get(offer.supplier)
                    if offer.price is not None and (current is None or offer.price < current):
                        by_supplier[offer.supplier] = offer.price
            ws.cell(n, min_col - len(price_cols) - 1, ", ".join(articles))
            for offset, name in enumerate(suppliers[key]):
                price = by_supplier.get(name)
                cell = ws.cell(n, price_cols[offset], price)
                if price is not None:
                    cell.number_format = MONEY
            if price_cols and how != "min":
                # Медиана и вторая снизу считаются по поставщику, а не по складу: число, как в отчёте.
                min_cell = ws.cell(n, min_col, r.prices.get(key))
            elif price_cols:
                start, end = get_column_letter(price_cols[0]), get_column_letter(price_cols[-1])
                min_cell = ws.cell(n, min_col, f'=IF(COUNT({start}{n}:{end}{n})=0,"",MIN({start}{n}:{end}{n}))')
            else:
                min_cell = ws.cell(n, min_col, None)
            min_cell.number_format = MONEY


def write_crosses_sheet(ws, analysis: list[dict], opt: Options, start_row: int = 1) -> int:
    table = []
    for c in analysis:
        for s in c["segments"]:
            priced = s["priced"].get(s["mode"], 0)
            table.append([
                c["order"], c["name"], segment_label(s, opt.cross_split), CROSS_MODES[s["mode"]], s["items"], s["yes"],
                s["yes_articles"], s["no_only"], s["none"], priced, priced / s["items"] if s["items"] else 0, s["verdict"],
            ])
    return plain_block(
        ws,
        f"Кроссы на конкурентов · {CROSS_SPLITS[opt.cross_split]} · цены: {PRICE_SOURCES[opt.price_source]}",
        ["№", "Конкурент", "Часть отчёта", "Взяли", "Товаров", "С кроссом «Да»", "Разных кроссов «Да»",
         "Только «Нет»", "Без кроссов", "С ценой", "Покрытие ценой", "Комментарий"],
        table,
        start_row=start_row,
        widths=[5, 13, 34, 20, 10, 12, 12, 11, 11, 10, 11, 80],
    )


def write_excluded_sheet(ws, dropped: list[tuple[Row, str]], start_row: int = 1) -> int:
    reasons = Counter(reason for _, reason in dropped)
    next_row = plain_block(ws, f"Не попали в отчёт: {len(dropped)}", ["Причина", "Товаров"],
                           [[k, v] for k, v in reasons.most_common()], start_row=start_row, widths=[34, 12])
    return plain_block(ws, "Список", ["Причина", "Код товара", "Артикул", "Наименование", "Подгруппа", "Продажи, шт", "Остаток", "В пути, шт", "Потенциал, шт"],
                [[reason, r.code, r.article, r.name, r.subgroup, r.sales_qty, r.stock, r.in_transit,
                  (r.stock or 0) + (r.in_transit or 0) if r.stock is not None or r.in_transit is not None else None]
                 for r, reason in dropped],
                start_row=next_row, widths=[34, 14, 16, 40, 28, 14, 12, 12, 14])


def write_params_sheet(ws, src: Sources, settings: Settings, opt: Options, kept: int, dropped: int,
                       price_files: list[PriceFile] | None, start_row: int = 1) -> int:
    builtin = {"discounts": "из настроек программы", "gz": "из настроек программы", "positioning": "зашитое в программу"}
    files = []
    for kind in ("pst", "sales", "sbs", "transit", "acp", "crosses", "discounts", "gz", "positioning"):
        locs = src.found.get(kind)
        files.append([KIND_TITLES[kind], locs[0].label if locs else builtin.get(kind, "— нет")])
    for pf in price_files or []:
        state = "используется" if pf.enabled and pf.offers else ("выключен" if not pf.enabled else pf.problem or "не распознан")
        files.append([f"Прайс: {pf.supplier_label}", f"{pf.id} — {len(pf.offers)} цен, {state}"])
    ro = opt.recs
    params = [
        ["Сформирован", dt.datetime.now().strftime("%d.%m.%Y %H:%M")],
        ["Вид отчёта", REPORT_KINDS[opt.report_kind]],
        ["ТМ / группа", f"{opt.brand} / {opt.group}"],
        ["Основа отчёта", "ПСТ (+ СБС, Продажи, скидки, ГЗ)"],
        ["Товаров в отчёте", kept],
        ["Исключено фильтрами", dropped],
        ["Подгруппы", SUBGROUP_MODES[opt.subgroup_mode]],
        ["Цены конкурентов", f"минимальная: {PRICE_SOURCES[opt.price_source]}"],
        ["Поставщики", "только из позиционирования" if opt.monitoring_only else "все"],
        ["Наличие", "только позиции в наличии" if opt.in_stock_only else "все позиции"],
        ["Кроссы", CROSS_SPLITS[opt.cross_split]],
        ["Цель по позиционированию", TARGETS[ro.target]],
        ["Допуск позиционирования", f"±{ro.tolerance:g}%"],
        ["Рекомендованные цены", "вставлены" if ro.enabled else "не вставлялись"],
    ]
    if ro.enabled:
        params += [
            ["  не понижать больше", f"{ro.down:g}%"],
            ["  не повышать больше", ", ".join(f"{k.lower()} {v:g}%" for k, v in ro.up.items())],
            ["  минимальная ВР", f"{ro.min_vr:g}%"],
            ["  топ продаж", f"ABC из 1С: AA — топ, лимиты ×{ro.top_factor:g}, вниз не снижаем; A — середина, ×{ro.mid_factor:g}. Без ABC — первые 65% и 80% продаж"],
            ["  округление Кат. А", f"{ro.rounding:g} руб." if ro.rounding else "нет"],
        ]
    params += [[f"Источник: {title}", SOURCE_CHOICES[opt.source(k)]] for k, (title, *_) in FIELDS.items()]
    params += [[f"Фильтр: {title}", choices.get(opt.filter(k), opt.filter(k))] for k, (title, choices, _) in FILTERS.items()]
    params.append(["Фильтр: категория ABC", ", ".join(x or "пусто" for x in opt.abc) if opt.abc else "все"])
    next_row = plain_block(ws, "Параметры формирования", ["Параметр", "Значение"], params, start_row=start_row, widths=[36, 80])
    next_row = plain_block(ws, "Документы", ["Документ", "Откуда взят"], files, start_row=next_row, widths=[36, 80])
    green = "дороже" if opt.green_when == "above" else "дешевле"
    red = "дешевле" if opt.green_when == "above" else "дороже"
    legend = [
        ["Свёрнутые столбцы", "закупка, МО, марки, продажи в рублях. Кнопка «+» справа от закреплённых столбцов их разворачивает. Разновидность закреплена и в эту группу не входит"],
        ["На виду", "разновидность, продажи, остаток, товар в пути, потенциал (остаток + в пути), ABC, СБС, Кат. А, ВР, суммы, цена и отклонение каждого конкурента"],
        ["В пути, шт", "сколько штук уже заказано, итог по файлу «Товар в пути». Пусто — в файле этого артикула нет"],
        ["Потенциал, шт", "остаток + товар в пути. Пустая ячейка в одном из слагаемых считается нулём, если второе заполнено"],
        ["Синий шрифт", "исходное значение — можно править, формулы пересчитаются"],
        ["Чёрный шрифт", "формула"],
        ["Кросс курсивом (коричневый)", "взят кросс «Нет»"],
        ["Кросс серым", "кросс взят из АЦПКД"],
        ["Кат. А (жёлтая)", "МО × (1 − скидка). Скидка записана числом в формуле строки"],
        ["Закупочная цена", "цена в валюте с листа «ПСТ» × курс. Курс записан числом в формуле строки"],
        ["СБС план", "закупочная цена × ГЗ. Коэффициент записан числом в формуле строки"],
        ["СБС", "себестоимость с листа «СБС», делённая на количество, прямо в этой ячейке"],
        ["ВР", "1 − СБС / Кат. А; красным — ниже минимальной ВР"],
        ["Сумма продажи / прибыли", "Продажи, шт × Кат. А и Сумма продажи × ВР; итог в строке 2 учитывает фильтр Excel"],
        ["Цена конкурента", "минимальная цена на момент формирования, числом. Цены ниже 10 ₽ и цены в разы ниже остальных по этому артикулу не берутся"],
        ["Откл", f"1 − Кат. А / цена конкурента; зелёный — мы {green}, красный — мы {red}"],
        ["Новая цена (жёлтая)", "рекомендованная цена — можно править; Изм., новое МО, ВР после и суммы пересчитаются"],
        ["Ценовой сегмент / Автопарк", "по маркам и моделям из «Разновидности» (справочник — в настройках программы)"],
    ]
    plain_block(ws, "Как читать отчёт", ["Обозначение", "Что значит"], legend, start_row=next_row, widths=[36, 80])
    _ = WEIGHT_CLASSES
