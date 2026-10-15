"""Снимок аналитики: рекомендованная цена меняет ВР и отклонение."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "Программа"))

from pricing.dashboard import summary  # noqa: E402
from pricing.recommend import Rec  # noqa: E402
from pricing.report import Options, Row  # noqa: E402


def row(**kw):
    base = dict(
        code="1", name="Диск", article="A1", brand="ADVICS", order_group="", category="",
        group="Тормозные диски", subgroup="Стандартные", variety="", purchase_currency="",
        purchase_price=None, purchase_rub=None, mo_currency="", mo=None, sales_qty=10, sales=None,
        stock=None, abc="A", sbs=700, sbs_plan=None, cat_a=1000, discount=None, gz=None,
        overstock="", discontinued=False, assortment="", pst=None, acp=None, target=1200,
    )
    base.update(kw)
    return Row(**base)


def test_potential_profit_uses_stock_and_transit_not_sales():
    from pricing.dashboard import money

    item = row(sales_qty=2, stock=10, in_transit=5, cat_a=100, sbs=60)
    sale, profit = money(item, 100, "potential")
    assert sale == 1500 and profit == 600
    same, _ = money(item, 100, "sales")
    assert same == 200


def test_rec_price_moves_margin_and_gap():
    rows = [
        row(code="1", article="A1", abc="A", sales_qty=10, cat_a=1000, sbs=700, target=1200,
            rec=Rec(1100, 0.1, "к цели", "up")),
        row(code="2", article="B1", abc="B", sales_qty=2, cat_a=500, sbs=400, target=500,
            rec=Rec(500, 0, "в пределах", "keep")),
    ]
    s = summary(rows, Options(rec={"enabled": True, "tolerance": 3, "min_vr": 32}))
    assert s["vr_rec"] > s["vr"]
    assert s["sale_rec"] > s["sale"]
    assert s["profit_rec"] > s["profit"]
    now = next(x for x in s["slices"] if x["kind"] == "all")
    after = next(x for x in s["slices"] if x["kind"] == "rec")
    assert abs(after["mean"]) < abs(now["mean"])
    abc_a = next(x for x in s["abc"] if x["name"] == "A")
    assert abc_a["vr_rec"] > abc_a["vr"]
    assert abc_a["mean_gap_rec"] is not None
    assert s["top_changes"][0]["article"] == "A1"


def test_later_groups_have_their_own_thresholds():
    from pricing.recommend import bundled_for

    filters = bundled_for("DOUBLE FORCE", "Фильтры автомобильные")
    assert filters["min_vr"] == 34 and filters["up"]["Премиум"] == 8
    nitto = bundled_for("NITTO", "Фильтры автомобильные")
    assert nitto["up"]["Премиум"] == 5 and nitto["down"] == 3
    sun = bundled_for("SUN", "Ремни, Цепи ГРМ")
    assert sun["tolerance"] == 5 and sun["down"] == 4
    pads = bundled_for("Double Force", "Тормозные колодки")
    assert pads["down"] == 8 and pads["min_vr"] == 34


def test_brief_states_overall_and_outside_positioning():
    from types import SimpleNamespace

    from pricing.brief import section
    from pricing.recommend import Rec

    comps = [
        SimpleNamespace(key="BREMBO", name="Brembo", order=1),
        SimpleNamespace(key="KASHIYAMA", name="Kashiyama", order=2),
    ]
    rows = [
        row(code="1", cat_a=105, target=100, sales_qty=10, sbs=70, abc="A", rec=Rec(102, 0, "", "down"),
            competitors=comps, prices={"BREMBO": 100, "KASHIYAMA": 100}, percents={"BREMBO": 0, "KASHIYAMA": 0}),
        row(code="2", cat_a=100, target=100, sales_qty=10, sbs=70, abc="A", rec=Rec(100, 0, "", "keep"),
            competitors=comps, prices={"BREMBO": 100, "KASHIYAMA": 100}, percents={"BREMBO": 0, "KASHIYAMA": 0}),
        row(code="3", cat_a=120, target=100, sales_qty=10, sbs=70, abc="C", rec=Rec(103, 0, "", "down"),
            competitors=comps, prices={"BREMBO": 100, "KASHIYAMA": 100}, percents={"BREMBO": 0, "KASHIYAMA": 0}),
    ]
    text = section(rows, Options(brand="ADVICS", group="Тормозные диски", rec={"enabled": True, "tolerance": 3}))
    assert "До проценки. Общее отклонение 8,3%, мы дороже. Brembo 8,3%, мы дороже. Kashiyama 8,3%, мы дороже." in text
    assert "Категории. A " in text and "C " in text
    assert "После проценки." in text and "ВР было" in text and "стало" in text
    assert "Вывод: в позиционировании. Поднять 0, опустить 2, не трогать 1." in text
    assert "п.п." not in text
    assert "===" not in text


def test_brief_without_a_competitor_price_has_no_fake_percent():
    from pricing.brief import section

    text = section([row(cat_a=100, target=None)], Options(brand="ADVICS", group="Тормозные диски"))
    assert text.strip() == "Вывод: нет цен для анализа."


def test_brief_is_written_next_to_the_report(tmp_path):
    from types import SimpleNamespace

    from pricing.brief import write_group_brief

    brembo = SimpleNamespace(key="BREMBO", name="Brembo", order=1)
    lynx = SimpleNamespace(key="LYNX", name="Lynx", order=1)
    path = write_group_brief(
        tmp_path,
        [row(target=100, competitors=[brembo], prices={"BREMBO": 100}, percents={"BREMBO": 0})],
        Options(brand="ADVICS", group="Тормозные диски"),
    )
    text = path.read_text(encoding="utf-8")
    assert path.name == "Разбор.txt"
    assert "Advics Тормозные диски" in text
    assert "===" not in text
    again = write_group_brief(
        tmp_path,
        [row(code="2", article="B1", target=100, competitors=[lynx], prices={"LYNX": 100}, percents={"LYNX": 0})],
        Options(brand="AVANTECH", group="Фильтры автомобильные"),
    )
    both = again.read_text(encoding="utf-8")
    assert "Advics Тормозные диски" in both and "Avantech Фильтры автомобильные" in both


def test_brief_keeps_every_group_when_one_is_rewritten(tmp_path):
    from types import SimpleNamespace

    from pricing.brief import section, write_group_brief

    comp = SimpleNamespace(key="BREMBO", name="Brembo", order=1)
    priced = dict(competitors=[comp], prices={"BREMBO": 100}, percents={"BREMBO": 0})
    assert "Вывод: нет цен для анализа." in section([], Options(brand="ADVICS", group="Тормозные диски"))
    write_group_brief(tmp_path, [row(target=100, **priced)], Options(brand="ADVICS", group="Тормозные диски"))
    write_group_brief(tmp_path, [row(target=100, **priced)], Options(brand="ADVICS", group="Тормозные колодки"))
    write_group_brief(tmp_path, [row(target=None)], Options(brand="ADVICS", group="Тормозные суппорты"))
    write_group_brief(tmp_path, [row(target=100, **priced)], Options(brand="DOUBLE FORCE", group="Пружины подвески"))
    write_group_brief(tmp_path, [row(target=None, competitors=[comp], prices={})], Options(brand="DOUBLE FORCE", group="Тормозные цилиндры и шланги"))
    write_group_brief(tmp_path, [row(target=110, cat_a=100, **priced)], Options(brand="ADVICS", group="Тормозные диски"))
    text = (tmp_path / "Разбор.txt").read_text(encoding="utf-8-sig")
    assert text.count("Advics Тормозные диски") == 1
    assert text.count("Advics Тормозные колодки") == 1
    assert "Advics Тормозные суппорты" not in text
    assert text.index("Целевые группы") < text.index("Мелкие группы") < text.index("Нет цен")
    assert text.index("Advics Тормозные колодки") < text.index("Мелкие группы")
    assert text.index("Мелкие группы") < text.index("Double Force Пружины подвески")
    assert text.index("Нет цен") < text.index("Double Force Тормозные цилиндры и шланги")
    assert text.count("Вывод:") == 4
