"""Реальные выгрузки из 1С и прайс поставщика (.xls) из «примеры/выгрузки_1С»."""

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "Программа"))

from pricing import prices, report as R  # noqa: E402
from pricing.excel import read_sheets  # noqa: E402
from pricing.settings import Settings  # noqa: E402
from pricing.sources import scan  # noqa: E402

REAL = ROOT / "примеры" / "выгрузки_1С"
TISS = REAL / "TissEKB.xls"

pytestmark = pytest.mark.skipif(not REAL.exists(), reason="нет папки с реальными выгрузками")


@pytest.fixture(scope="module")
def src():
    return scan([REAL])


def test_xls_reader():
    sheets = read_sheets(TISS)
    assert len(sheets) == 1
    grid = sheets[0].grid
    assert len(grid) > 1000
    assert any("Катал. номер" in str(v) for v in grid[5])


def test_1c_documents_recognised(src):
    assert src.pst and src.sbs and src.sales
    assert src.located("pst") and src.located("sales") and src.located("sbs")


def test_sales_in_pieces_and_period(src):
    assert src.sales_qty
    assert all(v >= 0 for v in src.sales_qty.values())
    assert src.sales_period


def test_tiss_price_detected(src):
    assert [s.file.name for s in src.price_sheets] == ["TissEKB.xls"]
    positioning = R.load_positioning(src)
    (pf,) = prices.load_price_files(src.price_sheets, {}, {}, R.known_suppliers(positioning))
    assert (pf.supplier, pf.city, pf.status) == ("TISS", "Екатеринбург", "auto")
    titles = pf.column_titles()
    assert titles["article"] == "Катал. номер" and titles["brand"] == "Бренд" and titles["price"] == "ОПТ"
    assert len(pf.offers) > 1000
    assert not pf.problem


def test_price_profile_applies_to_similar_file(src):
    (pf,) = prices.load_price_files(src.price_sheets, {}, {}, [])
    profile = prices.make_profile(pf)
    header, cols = prices.apply_profile(src.price_sheets[0], profile)
    assert header and cols == {k: pf.cols[k] for k in profile["columns"]}


def test_manual_header_row_is_respected(src):
    sheet = src.price_sheets[0]
    wrong = prices.load_one(sheet, {"header_row": 0}, {}, [])
    assert wrong.status == "none" or not wrong.offers
    assert wrong.problem
    manual = prices.load_one(sheet, {"header_row": 5, "cols": {"article": 3, "price": 8}, "fixed_brand": "X"}, {}, [])
    assert manual.status == "manual" and manual.offers and {o.brand for o in manual.offers} == {"X"}


def test_price_settings_endpoint(tmp_path, monkeypatch):
    import shutil

    from pricing import server, settings as settings_mod

    monkeypatch.setattr(settings_mod, "USER_FILE", tmp_path / "настройки.json")
    folder = tmp_path / "in"
    folder.mkdir()
    shutil.copy(TISS, folder)
    (f,) = server.api_scan({"folder": str(folder)})["price_files"]
    assert f["status"] == "auto" and f["offers"] > 1000
    dry = server.api_prices_save({"folder": str(folder), "id": f["id"], "header_row": 6, "cols": {}, "dry_run": True})["file"]
    assert dry["status"] == "manual" and dry["cols"]["article"] == f["cols"]["article"]
    saved = server.api_prices_save({"folder": str(folder), "id": f["id"], "supplier": "Тисс", "city": "ЕКБ",
                                    "header_row": 6, "cols": f["cols"], "remember": True})
    assert saved["message"].startswith("Формат запомнен")
    assert saved["files"][0]["supplier"] == "Тисс" and saved["files"][0]["status"] == "manual"
    off = server.api_prices_save({"folder": str(folder), "id": f["id"], "enabled": False})["files"][0]
    assert not off["enabled"] and off["supplier"] == "Тисс"
    reset = server.api_prices_save({"folder": str(folder), "id": f["id"], "reset": True})["files"][0]
    assert reset["enabled"] and reset["status"] == "saved"  # запомненный формат остался


def test_individual_settings_are_saved_apart_from_the_template(tmp_path, monkeypatch):
    from pricing import server, settings as settings_mod

    monkeypatch.setattr(settings_mod, "USER_FILE", tmp_path / "настройки.json")
    server.api_options({"options": {
        "ui": "template",
        "brand": "ADVICS",
        "group": "Тормозные диски",
        "report_kind": "full",
        "custom": {"filters": {"new": "exclude"}, "cross_split": "abc", "price_source": "prices", "rec": {"enabled": True, "target": "first"}},
    }})
    saved = settings_mod.Settings.load().options
    assert saved["ui"] == "template" and saved["report_kind"] == "full"
    assert saved["custom"]["filters"] == {"new": "exclude"}
    assert saved["custom"]["cross_split"] == "abc"
    assert saved["custom"]["rec"]["target"] == "first"


def test_missing_rates_reported(src):
    settings = Settings.load(None)
    groups = R.groups(src, "pst")
    opt = R.Options(brand=groups[0]["brand"], group=groups[0]["group"])
    rows = R.build_rows(src, settings, opt)
    missing = {m["currency"] for m in R.missing_rates(rows, settings)}
    used = {r.purchase_currency for r in rows if r.purchase_price}
    assert missing <= used
    assert all(settings.rate(c) is None for c in missing)
