"""Кроссы через ОЕМ: кросс «Да» нашего товара-двойника с тем же номером ОЕМ."""

import sys
from pathlib import Path
from types import SimpleNamespace as NS

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "Программа"))

from pricing import oemcross, report as R  # noqa: E402
from pricing.excel import Sheet  # noqa: E402
from pricing.sources import CrossRow, OemItem, PstItem, classify_oem, fill_oem_brands, norm_article, read_oem  # noqa: E402

HEAD = ["№ п.п.", "Код товара", "Наименование русс.", "Артикул", "Разновидность", "ОЕМ-номер 1", "ОЕМ-номер 2"]


def pst(code, article, brand, group):
    return PstItem(code=code, article=article, name="", brand=brand, category="", group=group, subgroup="", order_group="",
                   variety="", discontinued=False, overstock="", abc="", assortment="", purchase_currency="", purchase_price=None,
                   mo_currency="", mo_price=None, mo_rub=None)


def test_1c_print_without_brand_column_is_read():
    grid = [["Печать справочника \"Товары\" на 17.10.2026"], [], HEAD, HEAD, HEAD,
            [1, 1001, "Фильтр салонный Avantech", "CF1103", None, "27891JY15A / 27891-ED500", None],
            [2, 1002, "Фильтр салонный DOUBLE FORCE", "DFC10112", None, "27891JY15A", None],
            [3, 1003, "Аккумулятор", "L2R62", None, "Номинальная емкость (Ач) : / EFB / Пусковой ток :  / Полярность : R", None]]
    loc = classify_oem(Sheet(Path("ОЕМ.xlsx"), "TDSheet", grid))
    assert loc is not None
    items = read_oem(loc)
    # повторы шапки — не товары; характеристики аккумулятора — не номера ОЕМ
    assert [i.article for i in items] == ["CF1103", "DFC10112"]
    assert items[0].oems == ["27891JY15A", "27891ED500"] and items[0].brand == ""
    # марка — из ПСТ по коду, а если кода в ПСТ нет — по марке из ПСТ в наименовании
    filled = fill_oem_brands(items, [pst("1001", "CF1103", "AVANTECH", "Фильтры"), pst("9", "X1", "DOUBLE FORCE", "Фильтры")])
    assert filled == 2 and [i.brand for i in items] == ["AVANTECH", "DOUBLE FORCE"]


def test_twins_need_the_same_kind_and_a_rare_oem():
    rows = [OemItem("AVANTECH", "CF1103", "Фильтр салонный", "1", ["27891JY15A"]),
            OemItem("DOUBLE FORCE", "DFC10112", "Фильтр салонный", "2", ["27891JY15A"]),
            OemItem("AVANTECH", "AF999", "Фильтр воздушный", "3", ["27891JY15A"])]
    items = [pst("1", "CF1103", "AVANTECH", "Фильтры салона"), pst("2", "DFC10112", "DOUBLE FORCE", "Фильтры салона"),
             pst("3", "AF999", "AVANTECH", "Фильтры воздушные")]
    twins = oemcross.build_twins(rows, items)
    assert [t.article for t in twins[("DOUBLE FORCE", "DFC10112")]] == ["CF1103"]  # воздушный фильтр — другой вид
    # номер ОЕМ у десятка наших товаров — болт или общий корпус: не связывает
    many = [OemItem("AVANTECH", f"X{i}", "Болт", str(10 + i), ["9017920003"]) for i in range(12)]
    assert oemcross.build_twins(many, []) == {}


def test_report_takes_the_twin_cross_only_when_asked():
    comp = NS(key="LYNXAUTO", name="LYNX", order=1)
    crosses = {("AVANTECH", "CF1103"): {"LYNXAUTO": [CrossRow("LAC-710", norm_article("LAC-710"), "LYNXAUTO", True)]}}
    src = NS(crosses=crosses, oem_rows=[OemItem("AVANTECH", "CF1103", "Фильтр салонный", "1", ["27891JY15A"]),
                                         OemItem("DOUBLE FORCE", "DFC10112", "Фильтр салонный", "2", ["27891JY15A"])],
             pst=[])
    offer = R.PriceOffer("Армтек", 500.0, 3, "прайс")
    index = NS(get=lambda comp_key, article: [offer] if article == norm_article("LAC-710") else [])

    def run(mode):
        row = NS(brand="DOUBLE FORCE", article="DFC10112", subgroup="", abc="", acp=None, competitors=[comp])
        R.attach_crosses(src, [row], R.Options(oem_cross=mode), index, {})
        return row

    shown = run("show")
    assert shown.prices["LYNXAUTO"] is None and shown.oem["LYNXAUTO"]["price"] == 500.0
    assert shown.oem["LYNXAUTO"]["twin"] == "CF1103" and not shown.oem["LYNXAUTO"]["used"]
    used = run("use")
    assert used.prices["LYNXAUTO"] == 500.0 and used.oem["LYNXAUTO"]["used"]
    assert any(p.via == "CF1103" for p in used.picks["LYNXAUTO"])
    assert "LAC-710" in R.oem_text(used, "LYNXAUTO") and "27891JY15A" in R.oem_text(used, "LYNXAUTO")
    assert run("off").oem == {}
