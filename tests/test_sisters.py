"""Смежные марки: ADVICS↔AVANTECH по дискам и колодкам, AVANTECH↔DF по фильтрам."""
import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "Программа"))

from pricing.report import Options, prepare  # noqa: E402
from pricing.settings import Settings  # noqa: E402
from pricing.sisters import brand_key, sister_of  # noqa: E402
from pricing.sources import norm_article, scan  # noqa: E402
from pricing.workbook import HEAD_COLORS, build_report  # noqa: E402

PST = """Код товара;Артикул;Наименование;Разновидность;Группа заказа;Категория New;Группа New;Подгруппа New;Торговая марка;МО / Руб.
1;A6F-153B;Диск тормозной передний;LC 200;ГЗ;Детали тормозной системы;Тормозные диски;Стандартные тормозные диски;ADVICS;4000
2;BR0266S;Диск тормозной передний правый;LC 200;ГЗ;Детали тормозной системы;Тормозные диски;Стандартные тормозные диски;AVANTECH;1500
3;BR0267S;Диск тормозной передний левый перфорированный;LC 200;ГЗ;Детали тормозной системы;Тормозные диски;Стандартные тормозные диски;AVANTECH;900
4;AV100;Фильтр масляный;Camry;ГЗ;Система фильтрации;Фильтры;Масляные фильтры;AVANTECH;800
5;DF100;Фильтр масляный;Camry;ГЗ;Система фильтрации;Фильтры;Масляные фильтры;DOUBLE FORCE;500
"""

OEM = """Артикул;Наименование русс.;Торговая марка;ОЕМ-номер 1;ОЕМ-номер 2
A6F-153B;Диск тормозной передний;ADVICS;43512-60180;
BR0266S;Диск тормозной передний правый;AVANTECH;43512-60180 / 43512-60181;
BR0267S;Диск тормозной передний левый перфорированный;AVANTECH;43512-60180;
AV100;Фильтр масляный;AVANTECH;90915-YZZD2;
DF100;Фильтр масляный;DOUBLE FORCE;90915-YZZD2;
"""


def test_printed_1c_oem_catalog_fills_the_sister_column(tmp_path):
    """Как выгрузка 1С: шапка не в первой строке, ОЕМ во втором столбце."""
    folder = tmp_path / "in"
    folder.mkdir()
    (folder / "OEM AVANTECH+Advics Детали тормозной системы.csv").write_text(
        '"Печать справочника ""Товары""";;;;;;;\n'
        ";;;;;;;\n"
        "№ п.п.;Код товара;Наименование русс.;Артикул;Разновидность;Торговая марка;ОЕМ-номер 1;ОЕМ-номер 2\n"
        "1;100;Диск тормозной ADVICS;A6R173B;LC;ADVICS;42431-60221;\n"
        "2;200;Диск тормозной AVANTECH задний;BR0265S;LC;AVANTECH;42431-60221;\n",
        encoding="utf-8",
    )
    (folder / "пст.csv").write_text(
        "Код товара;Артикул;Наименование;Разновидность;Группа заказа;Категория New;Группа New;Подгруппа New;Торговая марка;МО / Руб.\n"
        "100;A6R173B;Диск тормозной ADVICS;LC;ГЗ;Детали тормозной системы;Тормозные диски;Стандартные тормозные диски;ADVICS;4000\n"
        "200;BR0265S;Диск тормозной AVANTECH задний;LC;ГЗ;Детали тормозной системы;Тормозные диски;Стандартные тормозные диски;AVANTECH;1500\n",
        encoding="utf-8",
    )
    src = scan([folder])
    from types import SimpleNamespace
    from pricing.sisters import price_book

    ws = SimpleNamespace(company=SimpleNamespace(pst=src.pst, acp={}, discounts={}), packs=[], acp_slices=[])
    src.sister_prices = price_book(ws, Settings.load(None))
    opt = Options(brand="ADVICS", group="Тормозные диски", sheet_filters=[])
    prep = prepare(src, Settings.load(None), opt)
    row = next(r for r in prep.kept if r.article == "A6R173B")
    assert row.sister_price == 1500
    result = build_report(src, Settings.load(None), opt, tmp_path)
    headers = [c.value for c in openpyxl.load_workbook(result.path)["Товары"][4]]
    assert "AVANTECH" in headers


def test_latin_oem_header_is_read(tmp_path):
    folder = tmp_path / "in"
    folder.mkdir()
    (folder / "OEM AVANTECH+Advics.csv").write_text(
        "Артикул;Торговая марка;OEM-номер\nAV100;AVANTECH;90915-YZZD2\nDF100;DOUBLE FORCE;90915-YZZD2\n",
        encoding="utf-8",
    )
    src = scan([folder])
    assert {i.article for i in src.oem} == {"AV100", "DF100"}
    assert src.oem_rows == src.oem


def test_sister_price_comes_from_pst_with_discount_not_from_acp():
    from types import SimpleNamespace

    from pricing.sisters import price_book

    item = SimpleNamespace(
        brand="AVANTECH", article="BR0265S", subgroup="Стандартные тормозные диски",
        order_group="ГЗ", mo_rub=1500, mo_price=None, mo_currency="",
    )
    acp = SimpleNamespace(brand="AVANTECH", article="BR0265S", cat_a_price=999)
    ws = SimpleNamespace(
        company=SimpleNamespace(pst=[], discounts={}),
        packs=[SimpleNamespace(sources=SimpleNamespace(pst=[item], acp={"1": acp}))],
        acp_slices=[SimpleNamespace(items={"1": acp})],
    )
    settings = Settings.load(None)
    settings.data["discounts"] = [["Стандартные тормозные диски", "ГЗ", "AVANTECH", 14]]
    book = price_book(ws, settings)
    assert book[("AVANTECH", norm_article("BR0265S"))] == 1290


def test_pairs_only_for_named_groups():
    assert sister_of("ADVICS", "Тормозные диски") == "AVANTECH"
    assert sister_of("AVANTECH", "Тормозные колодки") == "ADVICS"
    assert sister_of("ADVICS", "Тормозные суппорты") == ""
    assert sister_of("AVANTECH", "Фильтры") == "DF"
    assert sister_of("Double Force", "Масляные фильтры") == "AVANTECH"
    assert sister_of("ADVICS", "Фильтры") == ""
    assert brand_key("DOUBLE FORCE") == "DF"
    from pricing.sisters import sister_title
    assert sister_title("DF") == "DOUBLE FORCE"


def test_shared_oem_puts_sister_price_after_competitors(tmp_path):
    folder = tmp_path / "in"
    folder.mkdir()
    (folder / "пст.csv").write_text(PST, encoding="utf-8")
    (folder / "oem.csv").write_text(OEM, encoding="utf-8")
    src = scan([folder])
    assert {i.article for i in src.oem} == {"A6F-153B", "BR0266S", "BR0267S", "AV100", "DF100"}
    src.oem_rows = src.oem
    src.sister_prices = {
        ("AVANTECH", norm_article("BR0266S")): 1500.0,
        ("AVANTECH", norm_article("BR0267S")): 900.0,
    }
    settings = Settings.load(None)
    opt = Options(brand="ADVICS", group="Тормозные диски")
    prep = prepare(src, settings, opt)
    row = prep.kept[0]
    # Два сопоставимых артикула, 1500 и 900. В столбец идёт минимум.
    assert row.sister_brand == "AVANTECH"
    assert row.sister_price == 900
    result = build_report(src, settings, opt, tmp_path)
    wb = openpyxl.load_workbook(result.path)
    ws = wb["Товары"]
    headers = [c.value for c in ws[4]]
    price_at = headers.index("AVANTECH")
    assert headers.index("Сумма прибыли, ₽") < price_at
    assert "AVANTECH арт." not in headers
    assert ws.cell(5, price_at + 1).value == 900
    assert "E2EFDA" in (ws.cell(4, price_at + 1).fill.fgColor.rgb or "")


def test_filter_pair_and_no_column_for_other_brands(tmp_path):
    folder = tmp_path / "in"
    folder.mkdir()
    (folder / "пст.csv").write_text(PST, encoding="utf-8")
    (folder / "oem.csv").write_text(OEM, encoding="utf-8")
    src = scan([folder])
    src.oem_rows = src.oem
    src.sister_prices = {("DF", norm_article("DF100")): 500.0}
    settings = Settings.load(None)
    opt = Options(brand="AVANTECH", group="Фильтры")
    row = prepare(src, settings, opt).kept[0]
    assert row.sister_brand == "DF"
    assert row.sister_price == 500

    other = scan([folder])
    other.oem_rows = other.oem
    pads = prepare(other, settings, Options(brand="AVANTECH", group="Тормозные диски"))
    # Диск AVANTECH смотрит на ADVICS. В справочнике у ADVICS тот же ОЕМ, цены нет — артикул есть.
    assert pads.kept[0].sister_brand == "ADVICS"
    assert pads.kept[0].sister_price is None


def test_sister_column_is_next_to_cat_a_even_without_oem_file(tmp_path):
    folder = tmp_path / "in"
    folder.mkdir()
    (folder / "пст.csv").write_text(PST, encoding="utf-8")
    src = scan([folder])
    src.oem_rows = []
    src.oem = []
    settings = Settings.load(None)
    opt = Options(brand="ADVICS", group="Тормозные диски")
    prep = prepare(src, settings, opt)
    assert prep.kept[0].sister_brand == "AVANTECH"
    assert prep.kept[0].sister_price is None
    result = build_report(src, settings, opt, tmp_path)
    headers = [c.value for c in openpyxl.load_workbook(result.path)["Товары"][4]]
    assert headers.index("Сумма прибыли, ₽") < headers.index("AVANTECH")
    text = (tmp_path / "Разбор.txt").read_text(encoding="utf-8-sig")
    assert "До проценки" in text or "нет цен для анализа" in text


def test_oem_in_1c_folder_prices_the_neighbor_group(tmp_path):
    """Цена AVANTECH берётся из ПСТ соседней группы, а не из файла ОЕМ."""
    root = tmp_path / "proj"
    (root / "Выгрузки 1С").mkdir(parents=True)
    adv = root / "Данные" / "Группы" / "ADVICS — Тормозные диски"
    av = root / "Данные" / "Группы" / "AVANTECH — Тормозные диски"
    adv.mkdir(parents=True)
    av.mkdir(parents=True)
    (root / "Выгрузки 1С" / "OEM AVANTECH+Advics Детали тормозной системы.csv").write_text(
        '"Печать справочника ""Товары""";;;;;;;\n'
        ";;;;;;;\n"
        "№ п.п.;Код товара;Наименование русс.;Артикул;Разновидность;Торговая марка;ОЕМ-номер 1;ОЕМ-номер 2\n"
        "1;100;Диск тормозной ADVICS;A6R173B;LC;ADVICS;42431-60221;\n"
        "2;200;Диск тормозной AVANTECH задний;BR0265S;LC;AVANTECH;42431-60221;\n",
        encoding="utf-8",
    )
    pst = (
        "Код товара;Артикул;Наименование;Разновидность;Группа заказа;Категория New;Группа New;Подгруппа New;Торговая марка;МО / Руб.\n"
    )
    (adv / "пст.csv").write_text(
        pst + "100;A6R173B;Диск тормозной ADVICS;LC;ГЗ;Детали тормозной системы;Тормозные диски;Стандартные тормозные диски;ADVICS;4000\n",
        encoding="utf-8",
    )
    (av / "пст.csv").write_text(
        pst + "200;BR0265S;Диск тормозной AVANTECH задний;LC;ГЗ;Детали тормозной системы;Тормозные диски;Стандартные тормозные диски;AVANTECH;1500\n",
        encoding="utf-8",
    )
    from pricing.server import STATE
    from pricing.sisters import price_book

    ws = __import__("pricing.sources", fromlist=["scan_workspace"]).scan_workspace(root)
    assert len(ws.oem_rows) == 2
    src = STATE.for_group(root, "ADVICS", "Тормозные диски")
    assert len(src.oem_rows) == 2
    assert price_book(ws, Settings.load(None))[("AVANTECH", norm_article("BR0265S"))] == 1500
    opt = Options(brand="ADVICS", group="Тормозные диски", sheet_filters=["overstock", "sale"])
    result = build_report(src, Settings.load(None), opt, tmp_path)
    wb = openpyxl.load_workbook(result.path)
    wsht = wb["Товары"]
    assert wsht["A1"].value and not any("Показ" in str(c.value or "") for c in wsht[1])
    headers = [c.value for c in wsht[4]]
    col = headers.index("AVANTECH") + 1
    assert headers.index("Сумма прибыли, ₽") < headers.index("AVANTECH")
    assert wsht.cell(5, col).value == 1500
    assert (tmp_path / "Разбор.txt").is_file()
    import zipfile
    xml = zipfile.ZipFile(result.path).read("xl/worksheets/sheet1.xml").decode()
    assert "Показать,Скрыть" not in xml
