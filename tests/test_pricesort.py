"""Раскладка прайсов: файлы и архивы, брошенные в «Прайсы», — по папкам «дата / поставщик»."""

import datetime as dt
import os
import sys
import time
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "Программа"))

from pricing import pricesort as P  # noqa: E402

HEAD = "Бренд;Артикул;Наименование;Цена;Наличие\n"


def put(path: Path, text: str = HEAD + "ATE;24.0122-0101;диск;1000;5\n") -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")
    return path


def setup(tmp_path):
    prices = tmp_path / "Данные" / "Прайсы"
    old = prices / "16.10.2026"
    put(old / "Armtek" / "Armtek_16.10.2026.csv")
    put(old / "Berg" / "BERG_IRK_20261016_040019.csv")
    put(old / "Rossko" / "price.csv", "Производитель;Номер;Описание;Стоимость;Склад;Город\nATE;1;x;1;1;Вл\n")
    put(old / "MXGroup" / "price.csv")
    put(old / "Форум-Авто" / "Форум-Авто_16.10.2026.csv")
    put(prices / "валюта.csv", "Дата;USD\n16.10.2026;81\n")
    return prices


def test_new_date_is_sorted_by_past_dates_names_and_archives(tmp_path):
    prices = setup(tmp_path)
    new = prices / "17.10.2026"
    put(new / "Armtek_17.10.2026.csv")
    put(new / "BERG_IRK_20261017_041120.csv")
    put(new / "Shate-M прайс.csv")  # папки ещё нет, но поставщик есть в позиционировании
    put(new / "Непонятный.csv")
    with zipfile.ZipFile(new / "Rossko_17.10.2026.zip", "w") as z:
        z.writestr("Владивосток/price.csv", HEAD + "ATE;1;x;1;1\n")
    put(prices / "Форум-Авто.csv")  # брошен прямо в «Прайсы» — дата по дню скачивания
    today = dt.date.fromtimestamp(time.time())
    result = P.sort_prices(tmp_path, ["Shate-M", "Armtek"], today=today, map_path=tmp_path / "map.json")

    assert (new / "Armtek" / "Armtek_17.10.2026.csv").is_file()
    assert (new / "Berg" / "BERG_IRK_20261017_041120.csv").is_file()
    assert (new / "Shate-M" / "Shate-M прайс.csv").is_file()
    assert (new / "Rossko" / "Rossko_17.10.2026_price.csv").is_file()  # имя архива — в имя файла
    assert not (new / "Rossko_17.10.2026.zip").exists()
    assert (tmp_path / "Данные" / "Прайсы — исходники" / "Архивы ZIP" / "17.10.2026" / "Rossko_17.10.2026.zip").is_file()
    assert (prices / today.strftime("%d.%m.%Y") / "Форум-Авто" / "Форум-Авто.csv").is_file()
    assert (prices / "валюта.csv").is_file()
    assert (new / "Непонятный.csv").is_file()
    assert [u["file"] for u in result["unknown"]] == ["Непонятный.csv"]
    hows = {m["supplier"]: m["how"] for m in result["moved"]}
    assert hows["Armtek"] == "по прошлым датам" and hows["Shate-M"] == "по имени файла" and hows["Rossko"] == "по имени архива"
    lines = " ".join(P.summary_lines(result))
    assert "Распаковано архивов: 1" in lines and "Непонятный.csv" in lines

    # второй прогон — делать нечего; руками положенный файл запоминается
    assert P.sort_prices(tmp_path, [], map_path=tmp_path / "map.json")["moved"] == []


def test_generic_name_is_recognised_by_header_and_twins_are_dropped(tmp_path):
    prices = setup(tmp_path)
    new = prices / "17.10.2026"
    # «price.csv» есть и у Rossko, и у MXGroup — по имени не решить, по шапке — Rossko
    put(new / "price.csv", "Производитель;Номер;Описание;Стоимость;Склад;Город\nATE;2;y;2;2;Вл\n")
    put(new / "Armtek" / "Armtek_17.10.2026.csv")
    put(new / "Armtek_17.10.2026 (1).csv")  # тот же файл скачан второй раз
    result = P.sort_prices(tmp_path, [], map_path=tmp_path / "map.json")
    assert (new / "Rossko" / "price.csv").is_file()
    assert {m["how"] for m in result["moved"]} == {"по шапке"}
    assert result["dup"] == 1 and not (new / "Armtek_17.10.2026 (1).csv").exists()


def test_files_with_manual_settings_and_folders_without_dates_stay(tmp_path):
    prices = setup(tmp_path)
    new = prices / "17.10.2026"
    put(new / "Armtek_17.10.2026.csv")
    P.sort_prices(tmp_path, [], keep={"Armtek_17.10.2026.csv"}, map_path=tmp_path / "map.json")
    assert (new / "Armtek_17.10.2026.csv").is_file()

    flat = tmp_path / "flat"
    put(flat / "Данные" / "Прайсы" / "Armtek" / "a.csv")
    put(flat / "Данные" / "Прайсы" / "b.csv")
    assert P.sort_prices(flat, [], map_path=tmp_path / "map2.json")["moved"] == []
    assert (flat / "Данные" / "Прайсы" / "b.csv").is_file()


def test_russian_names_in_windows_zip_and_bad_zip(tmp_path):
    prices = setup(tmp_path)
    new = prices / "17.10.2026"
    new.mkdir(parents=True, exist_ok=True)
    # имя как у архиватора русской Windows: cp866 без флага UTF-8
    name = "Прайс Омск.csv".encode("cp866")
    raw = new / "Armtek.zip"
    with zipfile.ZipFile(raw, "w") as z:
        z.writestr("A" * (len(name) - 4) + ".csv", HEAD)
    raw.write_bytes(raw.read_bytes().replace(("A" * (len(name) - 4) + ".csv").encode(), name))
    (new / "битый.zip").write_bytes(b"PK\x03\x04 broken")
    result = P.sort_prices(tmp_path, [], map_path=tmp_path / "map.json")
    assert (new / "Armtek" / "Armtek_Прайс Омск.csv").is_file()
    assert [e["file"] for e in result["errors"]] == ["битый.zip"]
    assert (new / "битый.zip").is_file()  # не распакован — остаётся на месте
    assert not any(p.name.startswith("~распаковка") for p in (tmp_path / "Данные").iterdir())


def test_name_key_ignores_dates_and_copies():
    assert P.name_key("Armtek_17.10.2026.xlsx") == P.name_key("Armtek_18.10.2026 (1).xlsx")
    assert P.name_key("BERG_IRK_20261015_040019.xlsx") == P.name_key("BERG_IRK_20261016_041120.xlsx")
    assert P.name_key("Armtek(Омск).xlsx") != P.name_key("Armtek_ЕКБ.xlsx")
    os.environ.setdefault("OPENPYXL_LXML", "False")
