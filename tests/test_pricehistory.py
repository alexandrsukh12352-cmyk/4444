"""История цен: изменение в ПСТ, чьё оно и пауза после него."""

import datetime as dt
import sys
from pathlib import Path
from types import SimpleNamespace as NS

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "Программа"))

from pricing import pricehistory as H  # noqa: E402
from pricing.recommend import Rec  # noqa: E402


def row(code, mo, cat_a, rec=None, sales3=None):
    return NS(code=code, mo=mo, cat_a=cat_a, rec=rec, sales3=sales3, sales3_qty=None)


def test_change_is_ours_when_it_matches_the_recommendation(tmp_path):
    path = tmp_path / "история цен.json"
    day = dt.date(2026, 9, 1)
    first = [row("1", 1000, 800), row("2", 1000, 800), row("3", 1000, 800)]
    assert H.observe(first, day, path) == 0
    first[0].rec = Rec(840, 0.05, "к цели", "up")
    first[1].rec = Rec(840, 0.05, "к цели", "up")
    assert H.remember_recs(first, day, path) == 2
    # через две недели в ПСТ: первый — как рекомендовали, второй — иначе, третий не меняли
    later = [row("1", 1050, 840), row("2", 1100, 880), row("3", 1000, 800)]
    assert H.observe(later, day + dt.timedelta(days=14), path) == 2
    assert H.last_change("1", path)["by"] == "rec"
    assert H.last_change("2", path)["by"] == "other"
    assert H.last_change("3", path) is None
    assert "по рекомендации" in H.change_text(H.last_change("1", path))
    # пауза: через 30 дней после изменения цену не трогаем, через 100 — можно
    H.annotate(later, 90, day + dt.timedelta(days=44), path)
    assert later[0].frozen and "15.09.2026" in later[0].frozen and not later[2].frozen
    H.annotate(later, 90, day + dt.timedelta(days=120), path)
    assert not later[0].frozen and later[0].last_change
    H.annotate(later, 0, day + dt.timedelta(days=20), path)
    assert not later[0].frozen


def test_recent_sales_are_remembered_by_date(tmp_path):
    path = tmp_path / "история цен.json"
    H.observe([row("1", 1000, 800, sales3=5000)], dt.date(2026, 10, 1), path)
    H.observe([row("1", 1000, 800, sales3=7000)], dt.date(2026, 10, 1), path)
    H.observe([row("1", 1000, 800, sales3=9000)], dt.date(2026, 11, 1), path)
    series = H.load(path)["sales"]["1"]
    assert [p[0] for p in series] == ["2026-10-01", "2026-11-01"] and series[0][2] == 7000


def test_lost_competitor_prices_are_found_between_runs(tmp_path):
    path = tmp_path / "история цен.json"
    best = NS(supplier="Армтек · Владивосток")
    a = NS(code="1", prices={"BREMBO": 100.0}, best={"BREMBO": best})
    b = NS(code="2", prices={"BREMBO": 200.0}, best={"BREMBO": best})
    since, lost = H.track_priced("ADVICS|диски", [a, b], dt.date(2026, 10, 1), path)
    assert since == "" and lost == []
    # тот же день — сравнивать не с чем
    b.prices = {"BREMBO": None}
    assert H.track_priced("ADVICS|диски", [a, b], dt.date(2026, 10, 1), path)[1] == []
    # следующий день: у второго цены нет — потеря относительно 1 октября
    since, lost = H.track_priced("ADVICS|диски", [a, b], dt.date(2026, 10, 2), path)
    assert since == "2026-10-01" and [x["code"] for x in lost] == ["2"] and lost[0]["supplier"].startswith("Армтек")
    # другие настройки цен — свой снимок, без мнимых потерь
    assert H.track_priced("ADVICS|диски|median", [a, b], dt.date(2026, 10, 2), path)[1] == []


def test_cross_articles_under_another_brand_are_suggested():
    from pricing.price_check import article_hints
    from pricing.prices import Offer
    from pricing.sources import CrossRow, norm_article

    arts = [f"24.0122-0{i}" for i in range(6)]
    pf = NS(enabled=True, offers=[Offer("Continental Teves", a, 1000.0 + i, 3, i) for i, a in enumerate(arts)], supplier="Армтек")
    crosses = {("ADVICS", f"A{i}"): {"ATE": [CrossRow(a, norm_article(a), "ATE", True)]} for i, a in enumerate(arts)}
    hints = article_hints([pf], {"ATE": "ATE"}, crosses, set())
    assert hints and hints[0]["brand"] == "Continental Teves" and hints[0]["suggest"] == "ATE" and hints[0]["rows"] == 6
    # если под брендом ATE эти артикулы в прайсах тоже есть — это не потеря
    pf.offers += [Offer("ATE", a, 900.0, 1, 10 + i) for i, a in enumerate(arts)]
    assert article_hints([pf], {"ATE": "ATE"}, crosses, set()) == []


def test_stale_currency_rate_is_reported(monkeypatch):
    from pricing import server
    from pricing.settings import Settings

    s = Settings.load(None)
    s.data["currency_rates"] = {"JPY": 0.55, "USD": 81.0}
    s.data["currency_rates_date"] = (dt.date.today() - dt.timedelta(days=45)).isoformat()
    monkeypatch.setattr(server.dynamics, "read_currency",
                        lambda folder: {"rates": {"JPY": {"01.10.2026": 60.0, "15.10.2026": 61.0}, "USD": {"15.10.2026": 81.5}}})
    notes = server.rates_check(Path("."), s)
    texts = " ".join(n["text"] for n in notes)
    assert "JPY" in texts and "USD" not in texts  # иена за 100 единиц: 0,61 против 0,55; доллар в пределах 3%
    assert any(n.get("age") == 45 for n in notes)


def test_compaction_keeps_fresh_dates():
    from pricing import pricearchive as A

    shots = [NS(shot_date=d) for d in ("01.09.2026", "20.09.2026", "29.09.2026", "06.10.2026", "14.10.2026")]
    today = dt.date(2026, 10, 16)
    # старше двух недель — 01.09 и 20.09 и 29.09, но три последние даты остаются всегда
    assert A.plan(shots, today) == ["01.09.2026", "20.09.2026"]
    assert A.plan(shots, today, keep_min=1) == ["01.09.2026", "20.09.2026", "29.09.2026"]
    assert A.plan(shots[:2], today) == []
