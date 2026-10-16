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
MONEY0 = '#,##0;-#,##0;"—"'
MONEY = '#,##0.00'
SIGNED_RUB = '+#,##0" ₽";-#,##0" ₽";"0 ₽"'
PCT1 = '0.0%;-0.0%;0.0%'
PCT1S = '+0.0%;-0.0%;0.0%'
SHARE = '0%;-0%;"—"'
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


def write_summary_sheet(ws, rows, infos, opt, src, title: str, summary: dict) -> None:
    """Итоги группы. rows — товары отчёта, summary — dashboard.summary по ним."""
    ro = opt.recs
    tol = ro.tolerance / 100
    a = Across(infos)
    rec = ro.enabled and a.has("rec")
    pos = a.has("gap")
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.sheet_properties.tabColor = argb("0F4C5C")
    widths = {"A": 2, "B": 40, "C": 12, "D": 10, "E": 16, "F": 9, "G": 9, "H": 15, "I": 15, "J": 14,
              "K": 8, "L": 8, "M": 10, "N": 10, "O": 10, "P": 10}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    _cell(ws, "B2", f"{title} — итоги", font=Font(name=style.FONT, size=16, bold=True, color=style.ACCENT))
    ws.row_dimensions[2].height = 26
    bits = [f"сформирован {dt.datetime.now():%d.%m.%Y %H:%M}"]
    if getattr(src, "sales_period", ""):
        bits.append(f"продажи: {src.sales_period}")
    from .report import PRICE_PICKS, PRICE_SOURCES

    bits.append(f"цены конкурентов: {PRICE_SOURCES.get(opt.price_source, opt.price_source)}, "
                f"{PRICE_PICKS.get(getattr(opt, 'price_pick', 'min'), '').lower()}")
    bits.append("суммы с потенциала" if opt.money_base == "potential" else "суммы с продаж")
    _cell(ws, "B3", " · ".join(bits), font=MUTED)

    # ---- четыре цифры
    sale = a.sum("sum_sale", None)
    profit = a.sum("sum_profit", None)
    sale_vr = a.sum("sum_sale", None, ("vr", f'"{NUM}"'))
    vr_now = f'IFERROR({profit}/{sale_vr},"")'
    if rec:
        profit_r = a.sum("sum_profit_rec", None)
        sale_vr_r = a.sum("sum_sale_rec", None, ("vr_rec", f'"{NUM}"'))
        vr_rec = f'IFERROR({profit_r}/{sale_vr_r},"")'
        ups = a.count("change", None, '">0.0005"')
        downs = a.count("change", None, '"<-0.0005"')
    items = a.rows(None)
    priced_sale = a.sum("sum_sale", None, ("target", f'"{NUM}"')) if a.has("target") else "0"
    priced_items = a.count("target", None) if a.has("target") else "0"
    cards = [
        ("ВР группы", f"={vr_now}", PCT1,
         f'=IFERROR({vr_rec},"")' if rec else None, '"после рекомендаций: "0.0%' if rec else None),
        ("Прибыль", f"={profit}", MONEY0,
         f'={profit_r}-{profit}' if rec else None, '"после рекомендаций: "+#,##0" ₽";"после рекомендаций: "-#,##0" ₽";"без изменений"' if rec else None),
        ("Выручка с ценой конкурента", f'=IFERROR({priced_sale}/{sale},0)', SHARE,
         f'={priced_items}', '"товаров с ценой: "#,##0'),
        ("Двигаем цены", f'="↑ "&{ups}&"   ↓ "&{downs}' if rec else "рекомендации выключены", None,
         f"={items}-{ups}-{downs}" if rec else None, '"не трогаем: "#,##0' if rec else None),
    ]
    starts = ["B", "E", "H", "K"]
    ends = {"B": "D", "E": "G", "H": "J", "K": "M"}
    for col, (label, value, fmt, note, note_fmt) in zip(starts, cards):
        end = ends[col]
        for r in (5, 6, 7):
            ws.merge_cells(f"{col}{r}:{end}{r}")
            first = ws[f"{col}{r}"]
            first.fill = CARD_FILL
            for i in range(ord(col), ord(end) + 1):
                c = ws[f"{chr(i)}{r}"]
                c.fill = CARD_FILL
                c.border = Border(left=LINE if chr(i) == col else None, right=LINE if chr(i) == end else None,
                                  top=LINE if r == 5 else None, bottom=LINE if r == 7 else None)
        _cell(ws, f"{col}5", label, font=Font(name=style.FONT, size=10, color=style.MUTED), fill=CARD_FILL,
              align=Alignment(indent=1, vertical="bottom"))
        big = isinstance(value, str) and value.startswith("=")
        _cell(ws, f"{col}6", value, fmt, Font(name=style.FONT, size=18 if big else 11, bold=big,
                                                 color=style.ACCENT if big else style.MUTED),
              CARD_FILL, Alignment(indent=1, vertical="center", horizontal="left"))
        _cell(ws, f"{col}7", note, note_fmt, Font(name=style.FONT, size=10, color=style.MUTED), CARD_FILL,
              Alignment(indent=1, vertical="top", horizontal="left"))
    ws.row_dimensions[5].height = 18
    ws.row_dimensions[6].height = 30
    ws.row_dimensions[7].height = 18

    # ---- по подгруппам
    row = 9
    subgroups = list(dict.fromkeys(r.subgroup for r in rows))
    _section(ws, row, "По подгруппам",
             "ВР — по товарам с себестоимостью. Позиция к цели — доля выручки: дороже цели, в коридоре "
             f"±{ro.tolerance:g}%, дешевле, без цены конкурента. Всё — формулы по листам отчёта.")
    titles = ["Подгруппа", "Товаров", "С ценой конк.", "Сумма продажи", "ВР"]
    if rec:
        titles += ["ВР после", "Прибыль", "Прибыль после", "Δ прибыли", "↑", "↓"]
    else:
        titles += ["Прибыль"]
    if pos:
        titles += ["Дороже цели", "В коридоре", "Дешевле цели", "Без цены"]
    head = row + 2
    _header(ws, head, titles)
    lines = ([(sg or "без подгруппы", sg) for sg in subgroups] if len(subgroups) > 1 else []) + [("Итого", None)]
    for i, (label, sg) in enumerate(lines):
        n = head + 1 + i
        total = sg is None
        font = BOLD if total else INK
        fill = TOTAL_FILL if total else None
        vals: list[tuple[str | None, str | None]] = []
        s_sale = a.sum("sum_sale", sg)
        s_profit = a.sum("sum_profit", sg)
        s_vr = a.sum("sum_sale", sg, ("vr", f'"{NUM}"'))
        vals.append((f"={a.rows(sg)}", INT))
        vals.append((f"={a.count('target', sg)}" if a.has("target") else 0, INT))
        vals.append((f"={s_sale}", MONEY0))
        vals.append((f'=IFERROR({s_profit}/{s_vr},"")', PCT1))
        if rec:
            p_r = a.sum("sum_profit_rec", sg)
            s_vr_r = a.sum("sum_sale_rec", sg, ("vr_rec", f'"{NUM}"'))
            vals.append((f'=IFERROR({p_r}/{s_vr_r},"")', PCT1))
            vals.append((f"={s_profit}", MONEY0))
            vals.append((f"={p_r}", MONEY0))
            vals.append((f"={p_r}-{s_profit}", '+#,##0;-#,##0;0'))
            vals.append((f"={a.count('change', sg, chr(34) + '>0.0005' + chr(34))}", INT))
            vals.append((f"={a.count('change', sg, chr(34) + '<-0.0005' + chr(34))}", INT))
        else:
            vals.append((f"={s_profit}", MONEY0))
        if pos:
            dear = a.sum("sum_sale", sg, ("gap", f'"<"&{-tol}'))
            band = a.sum("sum_sale", sg, ("gap", f'">="&{-tol}'), ("gap", f'"<="&{tol}'))
            cheap = a.sum("sum_sale", sg, ("gap", f'">"&{tol}'))
            for part in (dear, band, cheap):
                vals.append((f'=IFERROR({part}/{s_sale},"")', SHARE))
            vals.append((f'=IFERROR(1-({dear}+{band}+{cheap})/{s_sale},"")', SHARE))
        _cell(ws, f"B{n}", label, font=font, fill=fill)
        ws[f"B{n}"].border = ROW_BORDER
        for j, (v, fmt) in enumerate(vals):
            c = ws.cell(n, 3 + j, v)
            c.number_format = fmt or "General"
            c.font = font
            c.border = ROW_BORDER
            if fill is not None:
                c.fill = fill
        ws.row_dimensions[n].height = 18
    last = head + len(lines)
    if rec:
        # ВР после — синим, если выше; Δ прибыли — зелёным / красным
        g = get_column_letter(7)
        ws.conditional_formatting.add(f"{g}{head + 1}:{g}{last}", FormulaRule(
            formula=[f"AND(ISNUMBER({g}{head + 1}),ISNUMBER(F{head + 1}),{g}{head + 1}>F{head + 1}+0.0005)"], font=Font(color=argb("1D5FA8"), bold=True)))
        d = get_column_letter(10)
        ws.conditional_formatting.add(f"{d}{head + 1}:{d}{last}", FormulaRule(
            formula=[f"AND(ISNUMBER({d}{head + 1}),{d}{head + 1}>0)"], font=Font(color=argb("15803D"), bold=True)))
        ws.conditional_formatting.add(f"{d}{head + 1}:{d}{last}", FormulaRule(
            formula=[f"AND(ISNUMBER({d}{head + 1}),{d}{head + 1}<0)"], font=Font(color=argb("B91C1C"), bold=True)))
    vr_col = "F"
    ws.conditional_formatting.add(f"{vr_col}{head + 1}:{vr_col}{last}", FormulaRule(
        formula=[f"AND(ISNUMBER({vr_col}{head + 1}),{vr_col}{head + 1}<{ro.min_vr / 100})"], font=Font(color=argb("C2410C"), bold=True)))
    if pos:
        first_pos = 3 + (10 if rec else 5)
        dear_col = get_column_letter(first_pos)
        ws.conditional_formatting.add(f"{dear_col}{head + 1}:{dear_col}{last}", FormulaRule(
            formula=[f"AND(ISNUMBER({dear_col}{head + 1}),{dear_col}{head + 1}>=0.2)"], fill=DEAR_FILL))
        cheap_col = get_column_letter(first_pos + 2)
        ws.conditional_formatting.add(f"{cheap_col}{head + 1}:{cheap_col}{last}", FormulaRule(
            formula=[f"AND(ISNUMBER({cheap_col}{head + 1}),{cheap_col}{head + 1}>=0.2)"], fill=CHEAP_FILL))
        none_col = get_column_letter(first_pos + 3)
        ws.conditional_formatting.add(f"{none_col}{head + 1}:{none_col}{last}", FormulaRule(
            formula=[f"AND(ISNUMBER({none_col}{head + 1}),{none_col}{head + 1}>=0.5)"], font=Font(color=argb("C2410C"), bold=True)))
    row = last + 2

    where = _locate(infos)
    letters_of = {i.title: i.letters for i in infos}

    # ---- что сильнее меняет прибыль
    changes = (summary.get("top_changes") or [])[:10] if rec else []
    if changes:
        _section(ws, row, "Что сильнее меняет прибыль",
                 "Десять товаров с самым большим изменением прибыли при тех же продажах. Цены — ссылки на лист отчёта.")
        head = row + 2
        _header(ws, head, ["Разновидность", "Артикул", "ABC", "Кат. А", "Рек. кат. А", "Изм.", "Δ прибыли", "Почему"], height=22)
        ws.merge_cells(start_row=head, start_column=9, end_row=head, end_column=16)
        by_code = {r.code: r for r in rows}
        for i, t in enumerate(changes):
            n = head + 1 + i
            r = by_code.get(t["code"])
            cat_a = _link(where, letters_of, t["code"], "cat_a")
            rec_c = _link(where, letters_of, t["code"], "rec")
            chg = _link(where, letters_of, t["code"], "change")
            p0 = _link(where, letters_of, t["code"], "sum_profit")
            p1 = _link(where, letters_of, t["code"], "sum_profit_rec")
            vals = [
                ((r.variety or r.name) if r else t["name"], None, INK), (t["article"], None, BOLD), (r.abc if r else "", None, INK),
                (f"={cat_a}" if cat_a else t["cat_a"], MONEY, INK),
                (f"={rec_c}" if rec_c else t["rec"], MONEY, INK),
                (f'=IFERROR({chg},"")' if chg else t["change"], PCT1S, UP_FONT if (t["change"] or 0) > 0 else DOWN_FONT),
                (f'=IFERROR({p1}-{p0},"")' if p0 and p1 else t["profit_delta"], '+#,##0;-#,##0;0', INK),
                (t["reason"], None, MUTED),
            ]
            for j, (v, fmt, font) in enumerate(vals):
                c = ws.cell(n, 2 + j, v)
                c.font = font
                c.border = ROW_BORDER
                if fmt:
                    c.number_format = fmt
        row = head + len(changes) + 2

    # ---- проценить вручную
    blind = sorted((r for r in rows if not r.target and (r.sales or 0) > 0), key=lambda r: -(r.sales or 0))[:15]
    if blind:
        blind_sale = sum(r.sales or 0 for r in rows if not r.target)
        all_sale = sum(r.sales or 0 for r in rows) or 1
        _section(ws, row, "Проценить вручную",
                 f"Без цены конкурента — {blind_sale / all_sale:.0%} продаж группы в рублях. Самые продаваемые из них: "
                 "по ним рекомендация видит только ВР.")
        head = row + 2
        _header(ws, head, ["Разновидность", "Артикул", "ABC", "Продажи, руб.", "Почему нет цены"], height=22)
        ws.merge_cells(start_row=head, start_column=6, end_row=head, end_column=13)
        for i, r in enumerate(blind):
            n = head + 1 + i
            vals = [(r.variety or r.name, None, INK), (r.article, None, BOLD), (r.abc, None, INK), (r.sales, MONEY0, INK),
                    (no_price_reason(r), None, MUTED)]
            for j, (v, fmt, font) in enumerate(vals):
                c = ws.cell(n, 2 + j, v)
                c.font = font
                c.border = ROW_BORDER
                if fmt:
                    c.number_format = fmt
        row = head + len(blind) + 2

    _section(ws, row, "Как читать")
    notes = [
        "Цель — цена конкурента × (1 + % позиционирования). «Откл. от цели» на листе отчёта: минус — мы дороже цели.",
        "Снижение на p при ВР m окупается, только если продажи вырастут на p / (m − p). Повышение безопасно, "
        "пока продажи падают не больше чем на p / (m + p).",
        "Топ продаж — категория AA из 1С: его не снижаем, шаг вверх меньше. A — середина, B и C — хвост.",
        "Рекомендованную цену на листе отчёта можно править — итоги пересчитаются.",
    ]
    for i, text in enumerate(notes):
        _cell(ws, f"B{row + 1 + i}", f"•  {text}", font=Font(name=style.FONT, size=10, color=style.INK))
    ws.freeze_panes = "A4"


def write_prices_overview(ws, rows, infos, opt) -> None:
    """Лист «Цены»: короткая таблица по товару, числа — ссылки на лист отчёта."""
    from .report import competitor_order
    from .workbook import COMP_TONES, HEAD_COLORS, GREEN, RED

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
        c.font = Font(name=style.FONT, bold=True, color=argb("FFFFFF"))
        c.fill = PatternFill("solid", fgColor=HEAD_COLORS[kind])
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
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
