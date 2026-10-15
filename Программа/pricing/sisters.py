"""Цена смежной марки по общему номеру ОЕМ.

Пары только эти: ADVICS и AVANTECH — диски и колодки; AVANTECH и DF — фильтры.
Файл узнаётся по столбцам «Артикул», «Торговая марка», «ОЕМ-номер».
Его ищут во «Входящие», в «Выгрузки 1С», в папках групп и в корне программы.
"""

from __future__ import annotations

from .sources import OemItem, Sources, Workspace, norm_article, norm_text

# Как зовём марку в столбце отчёта.
_DISPLAY = {"ADVICS": "ADVICS", "AVANTECH": "AVANTECH", "DF": "DOUBLE FORCE"}


def brand_key(brand: str) -> str:
    """ADVICS, AVANTECH или DF. Double Force пишут по-разному."""
    b = norm_text(brand).replace(" ", "").replace(".", "")
    if b in ("DF", "DOUBLEFORCE", "ДФ"):
        return "DF"
    return norm_text(brand)


def family_of(*texts: str) -> str:
    blob = " ".join(norm_text(t) for t in texts if t)
    if "КОЛОДК" in blob:
        return "pads"
    if "ДИСК" in blob:
        return "discs"
    if "ФИЛЬТР" in blob:
        return "filters"
    return ""


def sister_of(brand: str, group: str) -> str:
    """Марка смежного столбца. Пусто — для этой ТМ и группы столбца нет."""
    b = brand_key(brand)
    fam = family_of(group)
    if fam in ("discs", "pads") and b == "ADVICS":
        return "AVANTECH"
    if fam in ("discs", "pads") and b == "AVANTECH":
        return "ADVICS"
    if fam == "filters" and b == "AVANTECH":
        return "DF"
    if fam == "filters" and b == "DF":
        return "AVANTECH"
    return ""


def sister_title(brand: str) -> str:
    return _DISPLAY.get(brand, brand)


def _mo_rub(item, settings) -> float | None:
    """МО в рублях: столбец «МО / Руб.», иначе цена в валюте × курс."""
    if item.mo_rub is not None:
        return item.mo_rub
    if item.mo_price is None or settings is None:
        return None
    rate = settings.rate(item.mo_currency or "")
    if rate is None:
        return None
    return item.mo_price * rate


def _cat_a(item, discounts: dict, settings) -> float | None:
    disc = discounts.get((norm_text(item.subgroup), norm_text(item.order_group), norm_text(item.brand)))
    if disc is None and settings is not None:
        disc = settings.discounts.get((norm_text(item.subgroup), norm_text(item.order_group), norm_text(item.brand)))
    rub = _mo_rub(item, settings)
    if rub is not None and disc is not None:
        return round(rub * (1 - float(disc) / 100), 2)
    return rub


def price_book(ws: Workspace, settings) -> dict[tuple[str, str], float]:
    """(марка, артикул) -> Кат. А из ПСТ: МО в рублях с учётом скидки.

    Цена АЦПКД сюда не подставляется. Книга одна на все группы, пока не менялись курсы и скидки.
    """
    token = settings.calc_token()
    cached = getattr(ws, "_price_book", None)
    if cached and cached[0] == token:
        return cached[1]
    book: dict[tuple[str, str], float] = {}
    discounts = {}
    if ws.company and ws.company.discounts:
        discounts = ws.company.discounts
    items = list(ws.company.pst if ws.company else [])
    for pack in ws.packs:
        items.extend(pack.sources.pst)
    for item in items:
        price = _cat_a(item, discounts, settings)
        if price is None:
            continue
        book.setdefault((brand_key(item.brand), norm_article(item.article)), price)
    ws._price_book = (token, book)
    return book


def apply_sisters(rows, src: Sources) -> None:
    """Одна цена смежной марки: минимум Кат. А среди артикулов с общим ОЕМ.

    Столбец ставится для пары марок, даже если справочник ОЕМ ещё не прочитан.
    Тогда ячейки пустые, и в разборе написано, что файла нет.
    """
    if not rows:
        return
    sister = sister_of(rows[0].brand, rows[0].group)
    if not sister:
        return
    for r in rows:
        r.sister_brand = sister
        r.sister_price = None
    oem_rows: list[OemItem] = list(getattr(src, "oem_rows", None) or getattr(src, "oem", None) or [])
    if not oem_rows:
        return
    prices: dict = getattr(src, "sister_prices", None) or {}
    by_oem: dict[str, list[OemItem]] = {}
    own_oems: dict[tuple[str, str], list[str]] = {}
    for item in oem_rows:
        bk = brand_key(item.brand)
        ak = norm_article(item.article)
        own_oems.setdefault((bk, ak), [])
        for oem in item.oems:
            if len(oem) < 5:
                continue
            by_oem.setdefault(oem, []).append(item)
            if oem not in own_oems[(bk, ak)]:
                own_oems[(bk, ak)].append(oem)
    for r in rows:
        fam = family_of(r.group, r.name)
        oems = own_oems.get((brand_key(r.brand), norm_article(r.article)), [])
        seen: set[str] = set()
        best: float | None = None
        for oem in oems:
            for item in by_oem.get(oem, []):
                if brand_key(item.brand) != sister:
                    continue
                key = norm_article(item.article)
                if not key or key in seen:
                    continue
                item_fam = family_of(item.name)
                if fam and item_fam and fam != item_fam:
                    continue
                seen.add(key)
                price = prices.get((sister, key))
                if price is None:
                    continue
                best = price if best is None else min(best, price)
        r.sister_price = best
