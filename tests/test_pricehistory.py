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
