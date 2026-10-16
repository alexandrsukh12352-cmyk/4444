"""Номер сборки и проверки, которые видны во вкладке «Диагностика».

Если в шапке окна нет этой надписи, запущена другая папка программы.
"""

from __future__ import annotations

from pathlib import Path

BUILD = "2026-10-16.6"
BUILD_LABEL = "сборка 16 окт, 6"


def column_expectation(brand: str, group: str) -> dict:
    """Какой столбец смежной марки должен появиться у этой ТМ и группы."""
    from .sisters import sister_of, sister_title

    sister = sister_of(brand, group) if brand and group else ""
    if not sister:
        return {
            "header": "",
            "place": "",
            "text": "Для этой группы столбца смежной марки нет. Он есть у дисков и колодок ADVICS и AVANTECH и у фильтров AVANTECH и DOUBLE FORCE.",
        }
    header = sister_title(sister)
    return {
        "header": header,
        "place": "после цен конкурентов",
        "text": f"В отчёте {brand} · {group} после цен конкурентов должен быть зелёный столбец «{header}». Цена в нём — МО из ПСТ этой марки с учётом скидки.",
    }

_HERE = Path(__file__).resolve().parent


def program_marks() -> list[dict]:
    """Что умеет именно этот запущенный код, а не архив, который лежит рядом."""
    sisters = (_HERE / "sisters.py").read_text(encoding="utf-8")
    workbook = (_HERE / "workbook.py").read_text(encoding="utf-8")
    brief = (_HERE / "brief.py").read_text(encoding="utf-8")
    price_src = (_HERE / "prices.py").read_text(encoding="utf-8")
    page = (_HERE / "web" / "index.html").read_text(encoding="utf-8")
    price_at = workbook.find('f"{c.name} цена"')
    sister_at = workbook.find("sister_title(sister)")
    return [
        {
            "name": "Столбец смежной марки",
            "ok": "def sister_of" in sisters and "АЦПКД сюда не подставляется" in sisters,
            "detail": "Цена из ПСТ: МО в рублях с учётом скидки. Имя столбца — AVANTECH или DOUBLE FORCE.",
        },
        {
            "name": "Место столбца",
            "ok": price_at != -1 and sister_at > price_at,
            "detail": "Столбец стоит после цен конкурентов и перед рекомендованной ценой.",
        },
        {
            "name": "Текстовый разбор",
            "ok": 'NAME = "Разбор.txt"' in brief and "write_group_brief" in workbook,
            "detail": "Разбор.txt пишется по каждой группе из списка слева, даже если Excel по ней не собрался.",
        },
        {
            "name": "Разбор прайсов",
            "ok": "def price_date" in price_src and "def _data_from" in price_src and 'id="pa-index"' in page and 'id="dt-calendar"' in page,
            "detail": "Прайсы по папкам-датам и файл «валюта»: графики на вкладке «Прайсы», календарь и файлы по датам в «Данных». Файл без шапки читается с первой строки, столбец бренда проверяется по содержимому.",
        },
        {
            "name": "Окно загрузки",
            "ok": "PHASES" in (_HERE / "bootlog.py").read_text(encoding="utf-8") and 'id="boot-steps"' in page,
            "detail": "Загрузка показана этапами по порядку: что сделано, сколько заняло, что идёт сейчас и сколько файлов осталось. Ошибка остаётся на экране у своего этапа.",
        },
    ]
