"""Собирает пример прайса из тестовой выгрузки кроссов.

Подгруппа угадывается по префиксу артикула — это только для демонстрации,
в работе подгруппа берётся из вашего настоящего прайса.
"""

import csv
import re
import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
CROSSES = ROOT / "примеры" / "Кроссы свод (тестовая выгрузка).csv"
OUT = ROOT / "примеры" / "Прайс (пример).xlsx"

GUESS = {
    ("AVANTECH", "BR"): ("ДЕТАЛИ ТОРМОЗНОЙ СИСТЕМЫ", "Тормозные диски", "Стандартные тормозные диски"),
    ("AVANTECH", "AV"): ("ДЕТАЛИ ТОРМОЗНОЙ СИСТЕМЫ", "Тормозные колодки", "Дисковые тормозные колодки"),
    ("AVANTECH", "AF"): ("СИСТЕМЫ ФИЛЬТРАЦИИ", "Фильтры автомобильные", "Фильтры воздушные"),
    ("AVANTECH", "CF"): ("СИСТЕМЫ ФИЛЬТРАЦИИ", "Фильтры автомобильные", "Фильтры салонные"),
    ("AVANTECH", "OF"): ("СИСТЕМЫ ФИЛЬТРАЦИИ", "Фильтры автомобильные", "Фильтры масляные"),
    ("AVANTECH", "FF"): ("СИСТЕМЫ ФИЛЬТРАЦИИ", "Фильтры автомобильные", "Фильтры топливные"),
    ("ADVICS", "SN"): ("ДЕТАЛИ ТОРМОЗНОЙ СИСТЕМЫ", "Тормозные колодки", "Дисковые тормозные колодки"),
    ("DOUBLE FORCE", "DFP"): ("ДЕТАЛИ ТОРМОЗНОЙ СИСТЕМЫ", "Тормозные колодки", "Дисковые тормозные колодки"),
    ("DOUBLE FORCE", "DFA"): ("СИСТЕМЫ ФИЛЬТРАЦИИ", "Фильтры автомобильные", "Фильтры воздушные"),
    ("DOUBLE FORCE", "DFC"): ("СИСТЕМЫ ФИЛЬТРАЦИИ", "Фильтры автомобильные", "Фильтры салонные"),
    ("DOUBLE FORCE", "DFO"): ("СИСТЕМЫ ФИЛЬТРАЦИИ", "Фильтры автомобильные", "Фильтры масляные"),
}


def main(out: Path = OUT) -> Path:
    seen = {}
    with CROSSES.open(encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            if "+" in row["Артикул"]:
                continue
            prefix = re.match(r"[A-Z]+", row["Артикул поиск"])
            guess = GUESS.get((row["Бренд"], prefix.group(0) if prefix else ""))
            key = (row["Бренд"], row["Артикул поиск"])
            if guess and key not in seen:
                seen[key] = (row["Артикул"], row["Бренд"], *guess)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Прайс"
    # Порядок и набор столбцов специально отличается от свода: скрипт ищет столбцы по названию.
    ws.append(["Код товара", "Торговая марка", "Артикул", "Наименование", "Категория NEW", "Группа NEW", "Подгруппа NEW", "Цена"])
    for i, (art, brand, cat, grp, sub) in enumerate(sorted(seen.values(), key=lambda v: (v[1], v[0])), 1):
        ws.append([f"T{i:06d}", brand, art, f"{sub} {brand} {art}", cat, grp, sub, 1000 + i % 900])
    ws.append(["T999998", "AVANTECH", "AVTEST1", "Ремень ГРМ AVANTECH", "РЕМЕННОЙ И ЦЕПНОЙ ПРИВОД", "Ремни, Цепи ГРМ", "Ремни ГРМ", 1500])
    ws.append(["T999999", "SOMEBRAND", "XX-1", "Товар без позиционирования", "", "", "Свечи", 300])
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out


if __name__ == "__main__":
    print(main(Path(sys.argv[1]) if len(sys.argv) > 1 else OUT))
