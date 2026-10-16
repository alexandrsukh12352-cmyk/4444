"""Листы «Итоги» и «Цены» итогового отчёта.

«Итоги» — первый лист: четыре цифры, таблица по подгруппам, что сильнее меняет прибыль и что
проценить вручную. Цифры — формулы по листам отчёта: поправили Кат. А или рекомендованную цену —
итоги пересчитались. Списки товаров выбраны на момент формирования, числа в них — ссылки.

«Цены» — короткая таблица по товару: группа, подгруппа, разновидность, ABC, потенциал, продажи,
ВР, наша цена, цены конкурентов и отклонения. Числа — ссылки на лист отчёта.
"""

from __future__ import annotations

import datetime as dt

from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from . import style
from .style import argb

NUM = ">-1E+99"  # условие SUMIFS / COUNTIFS «ячейка — число»
MONEY0 = '#,##0;-#,##0;""'
MONEY = '#,##0.00'
SIGNED_RUB = '+#,##0;-#,##0;0'
PCT1 = '0.0%;-0.0%;0.0%'
PCT1S = '+0.0%;-0.0%;0.0%'
SHARE = '0%;-0%;""'
INT = '#,##0'

INK = Font(name=style.FONT, color=style.INK)
MUTED = Font(name=style.FONT, color=style.MUTED, size=9)
BOLD = Font(name=style.FONT, color=style.INK, bold=True)
HEAD_FILL = PatternFill("solid", fgColor=argb("1F3A4D"))
HEAD_FONT = Font(name=style.FONT, color=argb("FFFFFF"), bold=True, size=10)
CARD_FILL = PatternFill("solid", fgColor=argb("F3F7F8"))
TOTAL_FILL = PatternFill("solid", fgColor=argb("E8EEF1"))
LINE = Side(style="thin", color=argb("D5E0E3"))
ROW_BORDER = Border(bottom=Side(style="thin", color=argb("E3E8EC")))
UP_FONT = Font(name=style.FONT, color=argb("1D5FA8"), bold=True)
DOWN_FONT = Font(name=style.FONT, color=argb("B45309"), bold=True)
DEAR_FILL, CHEAP_FILL = PatternFill(bgColor=argb("FDE2D4")), PatternFill(bgColor=argb("DCEBFA"))


def _q(title: str) -> str:
    return "'" + title.replace("'", "''") + "'"


def _crit(value: str) -> str:
    """Текст для условия SUMIFS: без подстановочных знаков и с удвоенными кавычками."""
    text = (value or "").replace("~", "~~").replace("*", "~*").replace("?", "~?").replace('"', '""')
    return f'"{text}"'


class Across:
    """Формулы по всем листам отчёта, с условием на подгруппу или без."""

    def __init__(self, infos):
        self.infos = infos

    def parts(self, subgroup: str | None):
        for info in self.infos:
            if subgroup is None:
                yield info, None
                continue
            here = [r.subgroup for r in info.rows]
            if not here or subgroup not in here:
                continue
            if all(x == subgroup for x in here) or "subgroup" not in info.letters:
                yield info, None
            else:
                yield info, f'{info.rng("subgroup")},{_crit(subgroup)}'

    def sum(self, key: str, subgroup: str | None, *conds: tuple[str, str]) -> str:
        out = []
        for info, sg in self.parts(subgroup):
            if key not in info.letters or any(k not in info.letters for k, _ in conds):
                continue
            extra = [f"{info.rng(k)},{c}" for k, c in conds] + ([sg] if sg else [])
            out.append(f"SUMIFS({info.rng(key)},{','.join(extra)})" if extra else f"SUM({info.rng(key)})")
        return "(" + ("+".join(out) or "0") + ")"

    def count(self, key: str, subgroup: str | None, cond: str = f'"{NUM}"') -> str:
        out = []
        for info, sg in self.parts(subgroup):
            if key not in info.letters:
                continue
            out.append(f"COUNTIFS({info.rng(key)},{cond}{',' + sg if sg else ''})")
        return "(" + ("+".join(out) or "0") + ")"

    def rows(self, subgroup: str | None) -> str:
        out = []
        for info, sg in self.parts(subgroup):
            out.append(f"COUNTIFS({sg})" if sg else f"ROWS({info.rng('code')})")
        return "(" + ("+".join(out) or "0") + ")"

    def has(self, key: str) -> bool:
        return any(key in i.letters for i in self.infos)


def _cell(ws, ref: str, value, fmt: str | None = None, font: Font | None = None, fill=None, align=None):
    c = ws[ref]
    c.value = value
    if fmt:
        c.number_format = fmt
    c.font = font or INK
    if fill is not None:
        c.fill = fill
    if align is not None:
        c.alignment = align
    return c


def _header(ws, row: int, titles: list[str], first_col: int = 2, height: float = 30) -> None:
    for i, t in enumerate(titles):
        c = ws.cell(row, first_col + i, t)
        c.font, c.fill = HEAD_FONT, HEAD_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = height


def _section(ws, row: int, text: str, hint: str = "") -> None:
    _cell(ws, f"B{row}", text, font=Font(name=style.FONT, size=13, bold=True, color=style.INK))
    ws.row_dimensions[row].height = 22
    if hint:
        _cell(ws, f"B{row + 1}", hint, font=MUTED)


def _locate(infos) -> dict[str, tuple[str, int]]:
    where = {}
    for info in infos:
        for i, r in enumerate(info.rows):
            where.setdefault(r.code, (info.title, info.first + i))
    return where


def _link(where, letters_of, code: str, key: str) -> str | None:
    found = where.get(code)
    if not found:
        return None
    title, n = found
    col = letters_of[title].get(key)
    return f"{_q(title)}!${col}${n}" if col else None


def no_price_reason(r) -> str:
    picks = [p for lst in (r.picks or {}).values() for p in (lst or [])]
    if r.any_price and not r.target:
        names = ", ".join(c.name for c in r.competitors if r.prices.get(c.key))
        return f"цена есть только у {names}; цель — по первому в позиционировании"
    if any(p.article for p in picks):
        return "кросс есть, но его нет в прайсах поставщиков"
    return "нет кросса на конкурентов из позиционирования"


def _headline(title: str, summary: dict, rec: bool) -> str:
    """Заголовок листа — вывод, а не название: прибыль и где мы против главного конкурента."""
    bits = []
    if rec and summary.get("profit") is not None and summary.get("profit_rec") is not None:
        delta = (summary["profit_rec"] - summary["profit"]) / 1e6
        bits.append(f"{delta:+.1f} млн ₽ прибыли после рекомендаций".replace(".", ",").replace("-", "−"))
    elif summary.get("vr") is not None:
        bits.append(f"ВР {summary['vr'] * 100:.1f}%".replace(".", ","))
    comps = [s for s in summary.get("slices") or [] if s.get("kind") == "comp" and s.get("n")]
    if comps:
        main = max(comps, key=lambda s: s["n"])
        gap = main.get("weighted") if main.get("weighted") is not None else main.get("median")
        name = str(main["title"]).replace(" (без учёта % позиц.)", "").replace("Против ", "")
        if gap is not None and abs(gap) >= 0.005:
            side = "дешевле" if gap > 0 else "дороже"
            bits.append(f"мы {side} {name} на {abs(gap) * 100:.0f}%")
        elif gap is not None:
            bits.append(f"мы вровень с {name}")
    return f"{title}: " + ", ".join(bits) if bits else title


def _chart_data(ws, summary: dict, rec: bool, min_vr: float) -> dict:
    """Скрытый лист с числами для графиков (снимок на момент формирования). Возвращает диапазоны."""
    ws.sheet_state = "hidden"
    out = {}
    row = 1

    def block(name: str, head: list, rows: list, fmt: str = "0%;-0%;") -> tuple[int, int]:
        nonlocal row
        ws.cell(row, 1, name).font = BOLD
        for j, h in enumerate(head):
            ws.cell(row + 1, 1 + j, h)
        for i, values in enumerate(rows):
            for j, v in enumerate(values):
                c = ws.cell(row + 2 + i, 1 + j, v)
                if j:
                    c.number_format = fmt
        first, last = row + 1, row + 1 + len(rows)
        row = last + 2
        return first, last

    comps = [s for s in summary.get("slices") or [] if s.get("kind") == "comp" and s.get("n")]
    if comps:
        data = []
        for s in comps:
            name = str(s["title"]).replace(" (без учёта % позиц.)", "").replace("Против ", "")
            now = s.get("weighted") if s.get("weighted") is not None else s.get("median")
            after = s.get("weighted_rec") if s.get("weighted_rec") is not None else s.get("median_rec")
            data.append([f"{name} ({s['n']} тов.)", now, after if rec else None, s.get("target_dev")])
        out["comp"] = block("Мы против конкурентов", ["Конкурент", "Сейчас", "После", "Цель"], data, "+0%;-0%;0%")
    bridge = summary.get("bridge") if rec else None
    if bridge:
        up, down = bridge["up"]["delta"] / 1e6, bridge["down"]["delta"] / 1e6
        net = up + down
        # Водопад в Excel 2016 — столбцы с накоплением: невидимое основание и шаг.
        data = [["Повышения", 0, up, 0, 0], ["Снижения", min(up, net), 0, -down, 0], ["Итог", 0, 0, 0, net]]
        out["bridge"] = block("Откуда прибыль, млн ₽", ["Шаг", "Основание", "Рост", "Снижение", "Итог"], data, "#,##0.0;-#,##0.0;")
        out["bridge_ok"] = up >= 0 and net >= 0 and down <= 0
    pos = summary.get("position") or {}
    if pos.get("now"):
        def shares(parts):
            total = sum(p["sale"] for p in parts) or 1
            by = {p["key"]: p["sale"] / total for p in parts}
            return [by.get("no_price", 0), sum(v for k, v in by.items() if k.startswith("d")), by.get("band", 0),
                    sum(v for k, v in by.items() if k.startswith("c"))]
        data = [["Сейчас", *shares(pos["now"])]]
        if rec and pos.get("after"):
            data.append(["После", *shares(pos["after"])])
        out["pos"] = block("Позиция к цели, доля выручки", ["", "Нет цены конкурента", "Дороже цели", "В коридоре", "Дешевле цели"], data)
    abc = [x for x in summary.get("abc") or [] if x.get("items") and x.get("vr") is not None]
    if abc:
        data = [[x["name"] or "пусто", x["vr"], x["vr_rec"] if rec else None, min_vr / 100] for x in abc]
        out["abc"] = block("ВР по ABC", ["ABC", "Сейчас", "После", "Мин. ВР"], data)
    for col in "ABCDE":
        ws.column_dimensions[col].width = 22 if col == "A" else 12
    return out


GREY, BLUE, DARK = "A6A6A6", "2F6EBA", "404040"
GREEN, RED = "7FB069", "D9776A"


def _series_style(series, color: str, line: bool = False, dash: bool = False, none: bool = False) -> None:
    from openpyxl.drawing.line import LineProperties

    gp = series.graphicalProperties
    if none:
        gp.noFill = True
        gp.line = LineProperties(noFill=True)
        return
    if line:
        gp.line = LineProperties(solidFill=color, w=19050, prstDash="dash" if dash else "solid")
        return
    gp.solidFill = color
    gp.line = LineProperties(solidFill=color)


def _title(chart, text: str) -> None:
    """Заголовок графика 11 пт, не жирный в 14 пт, как ставит Excel по умолчанию."""
    from openpyxl.chart.text import RichText, Text
    from openpyxl.chart.title import Title
    from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties, RegularTextRun

    props = CharacterProperties(sz=1100, b=True, solidFill="404040")
    run = RegularTextRun(rPr=props, t=text)
    chart.title = Title(tx=Text(rich=RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=props), r=[run])])), overlay=False)


def _quiet_axes(chart, pct: bool = False, fmt: str | None = None) -> None:
    chart.y_axis.majorGridlines = None
    chart.legend.position = "b"
    chart.y_axis.delete = False
    chart.x_axis.delete = False
    if pct:
        chart.y_axis.numFmt = "0%"
    if fmt:
        chart.y_axis.numFmt = fmt


def _charts(ws, data_ws, ranges: dict, rec: bool, anchors: list[str]) -> None:
    from openpyxl.chart import BarChart, LineChart, Reference
    from openpyxl.chart.label import DataLabelList

    spots = iter(anchors)
    W, H = 12.6, 7.2
    if "comp" in ranges:
        first, last = ranges["comp"]
        ch = BarChart()
        ch.type = "bar"
        ch.grouping = "clustered"
        _title(ch, "Мы против конкурентов, плюс — мы дешевле")
        if rec:
            ch.add_data(Reference(data_ws, min_col=2, max_col=4, min_row=first, max_row=last), titles_from_data=True)
        else:
            ch.add_data(Reference(data_ws, min_col=2, max_col=2, min_row=first, max_row=last), titles_from_data=True)
            ch.add_data(Reference(data_ws, min_col=4, max_col=4, min_row=first, max_row=last), titles_from_data=True)
        ch.set_categories(Reference(data_ws, min_col=1, min_row=first + 1, max_row=last))
        colors = [GREY, BLUE, DARK] if rec else [GREY, DARK]
        for s, c in zip(ch.series, colors):
            _series_style(s, c)
        ch.gapWidth = 60
        ch.x_axis.scaling.orientation = "maxMin"
        ch.x_axis.tickLblPos = "low"  # подписи слева, а не у нуля поверх столбцов
        _quiet_axes(ch, pct=True)
        ch.width, ch.height = W, H
        ws.add_chart(ch, next(spots))
    if "bridge" in ranges and ranges.get("bridge_ok"):
        first, last = ranges["bridge"]
        ch = BarChart()
        ch.type = "col"
        ch.grouping = "stacked"
        ch.overlap = 100
        _title(ch, "Откуда прибыль, млн ₽, при тех же продажах")
        ch.add_data(Reference(data_ws, min_col=2, max_col=5, min_row=first, max_row=last), titles_from_data=True)
        ch.set_categories(Reference(data_ws, min_col=1, min_row=first + 1, max_row=last))
        _series_style(ch.series[0], "FFFFFF", none=True)
        for s, c in zip(ch.series[1:], [GREEN, RED, GREY]):
            _series_style(s, c)
            s.dLbls = DataLabelList()
            s.dLbls.showVal = True
            s.dLbls.showSerName = s.dLbls.showCatName = s.dLbls.showLegendKey = False
        ch.gapWidth = 60
        _quiet_axes(ch, fmt="#,##0.0")
        ch.legend = None
        ch.width, ch.height = W, H
        ws.add_chart(ch, next(spots))
    if "pos" in ranges:
        first, last = ranges["pos"]
        ch = BarChart()
        ch.type = "bar"
        ch.grouping = "percentStacked"
        ch.overlap = 100
        _title(ch, "Позиция к цели, доля выручки")
        ch.add_data(Reference(data_ws, min_col=2, max_col=5, min_row=first, max_row=last), titles_from_data=True)
        ch.set_categories(Reference(data_ws, min_col=1, min_row=first + 1, max_row=last))
        for s, c in zip(ch.series, ["D9D9D9", RED, GREEN, BLUE]):
            _series_style(s, c)
            s.dLbls = DataLabelList()
            s.dLbls.showVal = True
            s.dLbls.showSerName = s.dLbls.showCatName = s.dLbls.showLegendKey = False
        ch.gapWidth = 50
        ch.x_axis.scaling.orientation = "maxMin"
        _quiet_axes(ch, pct=True)
        ch.width, ch.height = W, H
        ws.add_chart(ch, next(spots))
    if "abc" in ranges:
        first, last = ranges["abc"]
        bar = BarChart()
        bar.type = "col"
        bar.grouping = "clustered"
        _title(bar, "ВР по ABC: сейчас и после")
        bar.add_data(Reference(data_ws, min_col=2, max_col=3 if rec else 2, min_row=first, max_row=last), titles_from_data=True)
        bar.set_categories(Reference(data_ws, min_col=1, min_row=first + 1, max_row=last))
        for s, c in zip(bar.series, [GREY, BLUE]):
            _series_style(s, c)
        line = LineChart()
        line.add_data(Reference(data_ws, min_col=4, max_col=4, min_row=first, max_row=last), titles_from_data=True)
        _series_style(line.series[0], RED, line=True, dash=True)
        line.series[0].marker.symbol = "none"
        bar.gapWidth = 80
        _quiet_axes(bar, pct=True)
        bar += line
        bar.width, bar.height = W, H
        ws.add_chart(bar, next(spots))


MLN = '#,##0.0,,;-#,##0.0,,;""'
HEAD_LINE = Border(bottom=Side(style="medium", color=argb("404040")))
TOTAL_LINE = Border(top=Side(style="thin", color=argb("404040")))


def _plain_header(ws, row: int, titles: list[str], first_col: int = 2) -> None:
    """Шапка как у человека: жирный текст, линия снизу, без заливки."""
    for i, t in enumerate(titles):
        c = ws.cell(row, first_col + i, t)
        c.font = Font(name=style.FONT, bold=True, color=style.INK, size=10)
        c.border = HEAD_LINE
        c.alignment = Alignment(horizontal="left" if i == 0 else "right", vertical="bottom", wrap_text=True)
    ws.row_dimensions[row].height = 30


def write_summary_sheet(ws, rows, infos, opt, src, title: str, summary: dict, data_ws=None) -> None:
    """Первый лист: вывод, четыре цифры, четыре графика Excel, короткая таблица и два списка."""
    ro = opt.recs
    a = Across(infos)
    rec = ro.enabled and a.has("rec")
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.sheet_properties.tabColor = argb("0F4C5C")
    widths = {"A": 2, "B": 34, "C": 11, "D": 11, "E": 11, "F": 11, "G": 11, "H": 11, "I": 11, "J": 11, "K": 11, "L": 11, "M": 11}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    _cell(ws, "B2", _headline(title, summary, rec), font=Font(name=style.FONT, size=15, bold=True, color=style.INK))
    ws.row_dimensions[2].height = 24
    bits = [f"{dt.datetime.now():%d.%m.%Y}"]
    if getattr(src, "sales_period", ""):
        bits.append(f"продажи {src.sales_period}")
    if getattr(src, "sales3_period", "") or getattr(src, "sales3", None):
        bits.append(f"продажи за 3 мес. {getattr(src, 'sales3_period', '')}".strip())
    from .report import PRICE_PICKS, PRICE_SOURCES

    bits.append(f"цены конкурентов: {PRICE_SOURCES.get(opt.price_source, opt.price_source)}, "
                f"{PRICE_PICKS.get(getattr(opt, 'price_pick', 'median'), '').lower()}")
    _cell(ws, "B3", " · ".join(bits), font=MUTED)

    # ---- четыре цифры: формулы по листам отчёта
    sale = a.sum("sum_sale", None)
    profit = a.sum("sum_profit", None)
    sale_vr = a.sum("sum_sale", None, ("vr", f'"{NUM}"'))
    vr_now = f'IFERROR({profit}/{sale_vr},"")'
    items = a.rows(None)
    priced_sale = a.sum("sum_sale", None, ("target", f'"{NUM}"')) if a.has("target") else "0"
    priced_items = a.count("target", None) if a.has("target") else "0"
    if rec:
        profit_r = a.sum("sum_profit_rec", None)
        sale_vr_r = a.sum("sum_sale_rec", None, ("vr_rec", f'"{NUM}"'))
        vr_rec = f'IFERROR({profit_r}/{sale_vr_r},"")'
        ups = a.count("change", None, '">0.0005"')
        downs = a.count("change", None, '"<-0.0005"')
    # Числа и подписи — в разных ячейках: кириллица и длинное тире в формате числа ломают styles.xml в Excel 2016.
    cards = [
        ("B", "C", "ВР группы", f"={vr_now}", "0.0%", "после рекомендаций:" if rec else "", f"={vr_rec}" if rec else None, "0.0%"),
        ("D", "F", "Прибыль, млн ₽", f"={profit_r}-{profit}" if rec else f"={profit}", "+#,##0.0,,;-#,##0.0,,;0" if rec else "#,##0.0,,",
         "сейчас:" if rec else "", f"={profit}" if rec else None, "#,##0.0,,"),
        ("G", "I", "С ценой конкурента, % выручки", f"=IFERROR({priced_sale}/{sale},0)", "0%", "товаров:", f"={priced_items}", "#,##0"),
        ("J", "L", "Цены вверх / вниз", f'={ups}&" / "&{downs}' if rec else "рекомендации выключены", None,
         "не трогаем:" if rec else "", f"={items}-{ups}-{downs}" if rec else None, "#,##0"),
    ]
    line = Side(style="thin", color=argb("BFBFBF"))
    for col, end, label, value, fmt, note, note_value, note_fmt in cards:
        for r in (5, 6, 7):
            for i in range(ord(col), ord(end) + 1):
                ws[f"{chr(i)}{r}"].border = Border(left=line if chr(i) == col else None, right=line if chr(i) == end else None,
                                                   top=line if r == 5 else None, bottom=line if r == 7 else None)
        _cell(ws, f"{col}5", label, font=Font(name=style.FONT, size=10, color=style.MUTED), align=Alignment(indent=1, vertical="bottom"))
        big = isinstance(value, str) and value.startswith("=")
        _cell(ws, f"{col}6", value, fmt, Font(name=style.FONT, size=18 if big else 11, bold=big, color=style.INK if big else style.MUTED),
              align=Alignment(indent=1, vertical="center", horizontal="left"))
        _cell(ws, f"{col}7", note, None, Font(name=style.FONT, size=10, color=style.MUTED), align=Alignment(indent=1, vertical="top"))
        if note_value is not None:
            nxt = chr(ord(col) + 1)
            _cell(ws, f"{nxt}7", note_value, note_fmt, Font(name=style.FONT, size=10, color=style.MUTED),
                  align=Alignment(vertical="top", horizontal="right" if col == "B" else "left"))
    ws.row_dimensions[5].height = 18
    ws.row_dimensions[6].height = 30
    ws.row_dimensions[7].height = 18

    # ---- четыре графика: числа на скрытом листе «Графики»
    if data_ws is not None:
        ranges = _chart_data(data_ws, summary, rec, ro.min_vr)
        _charts(ws, data_ws, ranges, rec, ["B9", "G9", "B25", "G25"])
        _cell(ws, "B41", "Графики — на момент формирования. Цифры сверху и таблица ниже пересчитываются, если поправить цену на листе отчёта.", font=MUTED)
        row = 43
    else:
        row = 9

    # ---- короткая таблица по подгруппам
    subgroups = list(dict.fromkeys(r.subgroup for r in rows))
    _cell(ws, f"B{row}", "По подгруппам", font=Font(name=style.FONT, size=12, bold=True, color=style.INK))
    titles = ["Подгруппа", "Товаров", "Выручка, млн ₽", "ВР"]
    if rec:
        titles += ["ВР после", "Прибыль, млн ₽", "Δ прибыли, млн ₽", "Вверх", "Вниз"]
    else:
        titles += ["Прибыль, млн ₽"]
    titles += ["С ценой конк."]
    head = row + 1
    _plain_header(ws, head, titles)
    lines = ([(sg or "без подгруппы", sg) for sg in subgroups] if len(subgroups) > 1 else []) + [("Итого", None)]
    for i, (label, sg) in enumerate(lines):
        n = head + 1 + i
        total = sg is None
        font = BOLD if total else INK
        s_sale = a.sum("sum_sale", sg)
        s_profit = a.sum("sum_profit", sg)
        s_vr = a.sum("sum_sale", sg, ("vr", f'"{NUM}"'))
        vals = [(f"={a.rows(sg)}", INT), (f"={s_sale}", MLN), (f'=IFERROR({s_profit}/{s_vr},"")', "0.0%")]
        if rec:
            p_r = a.sum("sum_profit_rec", sg)
            s_vr_r = a.sum("sum_sale_rec", sg, ("vr_rec", f'"{NUM}"'))
            vals += [(f'=IFERROR({p_r}/{s_vr_r},"")', "0.0%"), (f"={s_profit}", MLN),
                     (f"={p_r}-{s_profit}", '+#,##0.0,,;-#,##0.0,,;""'),
                     (f"={a.count('change', sg, chr(34) + '>0.0005' + chr(34))}", INT),
                     (f"={a.count('change', sg, chr(34) + '<-0.0005' + chr(34))}", INT)]
        else:
            vals.append((f"={s_profit}", MLN))
        priced = a.sum("sum_sale", sg, ("target", f'"{NUM}"')) if a.has("target") else "0"
        vals.append((f'=IFERROR({priced}/{s_sale},"")', "0%"))
        c = ws.cell(n, 2, label)
        c.font = font
        if total:
            c.border = TOTAL_LINE
        for j, (v, fmt) in enumerate(vals):
            c = ws.cell(n, 3 + j, v)
            c.number_format = fmt
            c.font = font
            c.alignment = Alignment(horizontal="right")
            if total:
                c.border = TOTAL_LINE
    last = head + len(lines)
    vr_col = "E"
    ws.conditional_formatting.add(f"{vr_col}{head + 1}:{vr_col}{last}", FormulaRule(
        formula=[f"AND(ISNUMBER({vr_col}{head + 1}),{vr_col}{head + 1}<{ro.min_vr / 100})"], font=Font(color=argb("C2410C"), bold=True)))
    if rec:
        d = "H"
        ws.conditional_formatting.add(f"{d}{head + 1}:{d}{last}", FormulaRule(
            formula=[f"AND(ISNUMBER({d}{head + 1}),{d}{head + 1}<0)"], font=Font(color=argb("B91C1C"), bold=True)))
    row = last + 2

    where = _locate(infos)
    letters_of = {i.title: i.letters for i in infos}
    changes = (summary.get("top_changes") or [])[:10] if rec else []
    if changes:
        _cell(ws, f"B{row}", "Что сильнее меняет прибыль", font=Font(name=style.FONT, size=12, bold=True, color=style.INK))
        head = row + 1
        _plain_header(ws, head, ["Разновидность", "Артикул", "ABC", "Кат. А", "Рек.", "Изм.", "Δ прибыли, тыс. ₽", "Почему"])
        ws.cell(head, 9).alignment = Alignment(horizontal="left", vertical="bottom")
        by_code = {r.code: r for r in rows}
        for i, t in enumerate(changes):
            n = head + 1 + i
            r = by_code.get(t["code"])
            link = lambda key: _link(where, letters_of, t["code"], key)  # noqa: E731
            cat_a, rec_c, chg, p0, p1 = link("cat_a"), link("rec"), link("change"), link("sum_profit"), link("sum_profit_rec")
            vals = [
                ((r.variety or r.name) if r else t["name"], None, INK, "left"),
                (t["article"], None, INK, "left"), (r.abc if r else "", None, INK, "left"),
                (f"={cat_a}" if cat_a else t["cat_a"], '#,##0', INK, "right"),
                (f"={rec_c}" if rec_c else t["rec"], '#,##0', INK, "right"),
                (f'=IFERROR({chg},"")' if chg else t["change"], '+0.0%;-0.0%;0%', INK, "right"),
                (f'=IFERROR({p1}-{p0},"")' if p0 and p1 else t["profit_delta"], '+#,##0,;-#,##0,;0', INK, "right"),
                (t["reason"], None, MUTED, "left"),
            ]
            for j, (v, fmt, font, align) in enumerate(vals):
                c = ws.cell(n, 2 + j, v)
                c.font = font
                c.alignment = Alignment(horizontal=align)
                if fmt:
                    c.number_format = fmt
        row = head + len(changes) + 2

    blind = sorted((r for r in rows if not r.target and (r.sales or 0) > 0), key=lambda r: -(r.sales or 0))[:10]
    if blind:
        blind_sale = sum(r.sales or 0 for r in rows if not r.target)
        all_sale = sum(r.sales or 0 for r in rows) or 1
        _cell(ws, f"B{row}", f"Проценить вручную — без цены конкурента {blind_sale / all_sale:.0%} продаж",
              font=Font(name=style.FONT, size=12, bold=True, color=style.INK))
        head = row + 1
        _plain_header(ws, head, ["Разновидность", "Артикул", "ABC", "Продажи, тыс. ₽", "Почему нет цены"])
        ws.cell(head, 6).alignment = Alignment(horizontal="left", vertical="bottom")
        for i, r in enumerate(blind):
            n = head + 1 + i
            vals = [(r.variety or r.name, None, INK, "left"), (r.article, None, INK, "left"), (r.abc, None, INK, "left"),
                    (r.sales, '#,##0,', INK, "right"), (no_price_reason(r), None, MUTED, "left")]
            for j, (v, fmt, font, align) in enumerate(vals):
                c = ws.cell(n, 2 + j, v)
                c.font = font
                c.alignment = Alignment(horizontal=align)
                if fmt:
                    c.number_format = fmt
    ws.freeze_panes = "A4"


def write_prices_overview(ws, rows, infos, opt) -> None:
    """Лист «Цены»: короткая таблица по товару, числа — ссылки на лист отчёта."""
    from .report import competitor_order
    from .workbook import COMP_TONES, GREEN, RED

    ro = opt.recs
    tol = ro.tolerance / 100
    where = _locate(infos)
    letters_of = {i.title: i.letters for i in infos}
    comps = competitor_order(rows)
    rec = ro.enabled and any("rec" in i.letters for i in infos)
    has_qty = any((r.sales_qty or 0) > 0 for r in rows)
    has_target = any(r.target for r in rows)
    has_potential = any("potential" in i.letters for i in infos)

    def link(r, key):
        ref = _link(where, letters_of, r.code, key)
        return f'=IF(ISNUMBER({ref}),{ref},"")' if ref else None

    # (заголовок, вид, ширина, значение(r, n, L), формат, ключ)
    cols: list[tuple] = [
        ("Артикул", "id", 14, lambda r, n, L: r.article, None, "article"),
        ("Наименование", "id", 30, lambda r, n, L: r.name, None, "name"),
        ("Группа", "info", 16, lambda r, n, L: r.group, None, "group"),
        ("Подгруппа", "info", 22, lambda r, n, L: r.subgroup, None, "subgroup"),
        ("Разновидность", "info", 36, lambda r, n, L: r.variety, None, "variety"),
        ("ABC", "calc", 6, lambda r, n, L: r.abc, None, "abc"),
        ("Потенциал, шт" if has_potential else "Остаток, шт", "calc", 10,
         lambda r, n, L: link(r, "potential" if has_potential else "stock"), '#,##0', "potential"),
        ("Продажи, руб.", "calc", 14, lambda r, n, L: link(r, "sales") or r.sales, '#,##0', "sales"),
    ]
    if has_qty:
        cols.append(("Продажи, шт", "calc", 9, lambda r, n, L: link(r, "sales_qty"), '#,##0', "sales_qty"))
    if any(getattr(r, "sales3", None) is not None for r in rows):
        qty3 = any(getattr(r, "sales3_qty", None) for r in rows)
        cols.append(("Продажи 3 мес, шт" if qty3 else "Продажи 3 мес, руб.", "calc", 10,
                     lambda r, n, L: link(r, "sales3"), '#,##0', "sales3"))
    cols += [
        ("ВР", "calc", 7, lambda r, n, L: link(r, "vr"), PCT1, "vr"),
        ("Кат. А", "calc", 12, lambda r, n, L: link(r, "cat_a"), MONEY, "cat_a"),
    ]
    for i, c in enumerate(comps):
        cols.append((f"{c.name} цена", "comp", 12, lambda r, n, L, k=c.key: link(r, f"price:{k}"), MONEY, f"price:{c.key}"))
        cols.append((f"Откл {c.name}", "comp", 8, lambda r, n, L, k=c.key:
                     f'=IF(AND(ISNUMBER({L["price:" + k]}{n}),ISNUMBER({L["cat_a"]}{n})),1-{L["cat_a"]}{n}/{L["price:" + k]}{n},"")',
                     PCT1, f"dev:{c.key}"))
    if has_target:
        cols.append(("Цель", "pos", 12, lambda r, n, L: link(r, "target"), MONEY, "target"))
        cols.append(("Откл. от цели", "pos", 9, lambda r, n, L:
                     f'=IF(AND(ISNUMBER({L["target"]}{n}),ISNUMBER({L["cat_a"]}{n})),1-{L["cat_a"]}{n}/{L["target"]}{n},"")',
                     PCT1, "gap"))
    if rec:
        cols.append(("Рек. кат. А", "rec", 12, lambda r, n, L: link(r, "rec"), MONEY, "rec"))
        cols.append(("Изм. %", "rec", 8, lambda r, n, L:
                     f'=IF(AND(ISNUMBER({L["rec"]}{n}),ISNUMBER({L["cat_a"]}{n})),{L["rec"]}{n}/{L["cat_a"]}{n}-1,"")',
                     PCT1S, "change"))

    L = {key: get_column_letter(i) for i, (*_, key) in enumerate(cols, 1)}
    ordered = sorted(rows, key=lambda r: (r.group or "", r.subgroup or "", -(r.sales or 0), r.article))
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.sheet_properties.tabColor = argb("9A5B13")
    ws.cell(2, 1, "Цены: наша, конкурентов и отклонения").font = Font(name=style.FONT, size=14, bold=True, color=style.INK)
    ws.cell(2, 6, "Отклонение = 1 − Кат. А / цена конкурента: минус — мы дороже. Числа — ссылки на лист отчёта.").font = MUTED
    ws.row_dimensions[2].height = 24
    head = 3
    for i, (t, kind, width, *_rest) in enumerate(cols, 1):
        c = ws.cell(head, i, t)
        c.font = Font(name=style.FONT, bold=True, color=style.INK, size=10)
        c.border = HEAD_LINE
        c.alignment = Alignment(wrap_text=True, vertical="bottom", horizontal="left" if i <= 6 else "right")
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[head].height = 36
    first = head + 1
    comp_index = {f"price:{c.key}": i % 2 for i, c in enumerate(comps)} | {f"dev:{c.key}": i % 2 for i, c in enumerate(comps)}
    for n, r in enumerate(ordered, first):
        for i, (_t, _kind, _w, value, fmt, key) in enumerate(cols, 1):
            v = value(r, n, L)
            c = ws.cell(n, i, v if v != "" else None)
            c.font = BOLD if key == "article" else INK
            c.border = ROW_BORDER
            if fmt:
                c.number_format = fmt
            if key in comp_index:
                c.fill = COMP_TONES[comp_index[key]]
    last = max(first, first + len(ordered) - 1)
    above, below = (GREEN, RED) if opt.green_when == "above" else (RED, GREEN)
    for key, col in L.items():
        rng = f"{col}{first}:{col}{last}"
        x = f"{col}{first}"
        if key.startswith("dev:"):
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({x}),{x}<0)"], fill=above[0], font=above[1]))
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({x}),{x}>0)"], fill=below[0], font=below[1]))
        elif key == "gap":
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({x}),{x}<-{tol})"], fill=above[0], font=above[1]))
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({x}),{x}>{tol})"], fill=below[0], font=below[1]))
        elif key == "change":
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({x}),{x}>0.0005)"], font=Font(color=argb("1D5FA8"), bold=True)))
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({x}),{x}<-0.0005)"], font=Font(color=argb("B45309"), bold=True)))
        elif key == "vr":
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({x}),{x}<{ro.min_vr / 100})"], font=Font(color=argb("C2410C"), bold=True)))
    ws.freeze_panes = f"C{first}"
    for sel in ws.sheet_view.selection:
        if sel.pane == "bottomRight":
            sel.activeCell = sel.sqref = f"C{first}"
        elif sel.pane == "topRight":
            sel.activeCell = sel.sqref = "C1"
        elif sel.pane == "bottomLeft":
            sel.activeCell = sel.sqref = f"A{first}"
    ws.auto_filter.ref = f"A{head}:{get_column_letter(len(cols))}{last}"
