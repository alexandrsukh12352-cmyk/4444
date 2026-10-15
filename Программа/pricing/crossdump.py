"""Один файл: наши кроссы и цены из прайсов, которые на них встали.

Строка — наш артикул, кросс и поставщик. Кросс без цены попадает на отдельный лист,
только если в своде стоит «Да». Сравнение артикулов — без точек и дефисов.
"""

from __future__ import annotations

from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook

from .prices import PriceFile, brand_key, supplier_key
from .report import MIN_PRICE, monitoring_map
from .sources import norm_article

CROSS_NAME = "Кроссы и цены.xlsx"
_CAP = 500_000


def _index(price_files: list[PriceFile]):
    """(бренд, артикул) -> лучшая цена каждого поставщика и города."""
    found: dict[tuple[str, str], dict[tuple, tuple]] = defaultdict(dict)
    used = []
    skipped = []
    for pf in price_files:
        if not pf.enabled or not pf.offers:
            if pf.problem or pf.doubt:
                skipped.append(pf)
            continue
        if pf.doubt:
            skipped.append(pf)
        used.append(pf)
        for offer in pf.offers:
            if not offer.price or offer.price < MIN_PRICE:
                continue
            key = (brand_key(offer.brand), offer.article_key or norm_article(offer.article))
            mark = (pf.supplier, pf.city)
            current = found[key].get(mark)
            if current is None or offer.price < current[0]:
                found[key][mark] = (offer.price, offer.qty, pf.sheet.file.name, pf.supplier)
    return found, used, skipped


def write_cross_prices(path: Path, crosses: dict, price_files: list[PriceFile], positioning) -> dict:
    """Пишет книгу. Пустой свод кроссов — ошибка вызывающего, сюда не доходит."""
    index, used, skipped = _index(price_files)
    monitored = monitoring_map(positioning) if positioning is not None else {}
    book = Workbook(write_only=True)
    prices = book.create_sheet("Цены")
    missing = book.create_sheet("Без цены")
    about = book.create_sheet("Как собрано")
    head = [
        "Наш бренд", "Наш артикул", "Бренд кросса", "Кросс-артикул", "Для сравнения",
        "Поставщик", "Город", "Цена", "Наличие", "Файл прайса", "Мониторинг",
    ]
    prices.append(head)
    missing.append(["Наш бренд", "Наш артикул", "Бренд кросса", "Кросс-артикул"])
    matched = 0
    crosses_n = 0
    yes_missing = 0
    stopped = False
    for (our_brand, our_article), groups in crosses.items():
        for rows in groups.values():
            for row in rows:
                crosses_n += 1
                hits = index.get((brand_key(row.cross_brand), row.cross_key), {})
                if not hits:
                    if row.yes:
                        yes_missing += 1
                        if yes_missing <= _CAP:
                            missing.append([our_brand, our_article, row.cross_brand, row.cross_article])
                    continue
                allowed = monitored.get(brand_key(row.cross_brand))
                for (supplier, city), (price, qty, filename, supplier_name) in hits.items():
                    if matched >= _CAP:
                        stopped = True
                        break
                    if allowed is None:
                        flag = "поставщик не мониторит бренд по позиционированию"
                    elif supplier_key(supplier_name) in allowed:
                        flag = "да"
                    else:
                        flag = "нет"
                    prices.append([
                        our_brand, our_article, row.cross_brand, row.cross_article,
                        "Да" if row.yes else "Нет", supplier, city, price, qty if qty is not None else "",
                        filename, flag,
                    ])
                    matched += 1
            if stopped:
                break
        if stopped:
            break
    about.append(["Что в файле"])
    about.append(["Кроссов в своде", crosses_n])
    about.append(["Строк с ценой", matched])
    about.append(["Кроссов «Да» без цены", yes_missing])
    about.append(["Прайсов в выгрузке", len(used)])
    about.append(["Цена ниже 10 ₽ не берётся. По поставщику и городу остаётся минимальная."])
    about.append(["Артикул сравнивается без точек, пробелов и дефисов."])
    about.append(["Цена — столбец, где не меньше 80% чисел и название не похоже на наименование."])
    about.append(["Наименование — текстовый столбец. Столбец из чисел названием не становится."])
    about.append(["Ручная настройка одного прайса применяется ко всем файлам с теми же названиями столбцов."])
    if skipped:
        about.append(["Файлы, где столбец цены стоит проверить"])
        for pf in skipped:
            about.append([pf.sheet.file.name, pf.doubt or pf.problem or pf.supplier])
    if stopped:
        about.append([f"Выгрузка обрезана на {_CAP} строках с ценой."])
    path.parent.mkdir(parents=True, exist_ok=True)
    book.save(path)
    return {
        "crosses": crosses_n,
        "matched": matched,
        "missing": yes_missing,
        "files": len(used),
        "stopped": stopped,
    }
