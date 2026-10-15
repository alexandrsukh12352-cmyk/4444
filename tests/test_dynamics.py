"""Папка-дата, файл «валюта» и кэш сравнения цен."""
import sys
from pathlib import Path
from types import SimpleNamespace

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "Программа"))

from pricing import dynamics, prices  # noqa: E402
from pricing import report as R  # noqa: E402
from pricing.sources import CrossRow, Sources, norm_article, norm_text, scan  # noqa: E402


def _write(folder: Path, name: str, text: str) -> None:
    folder.mkdir(parents=True, exist_ok=True)
    (folder / name).write_text(text, encoding="utf-8")


def test_date_folder_keeps_the_supplier_under_the_snapshot(tmp_path):
    day = tmp_path / "Прайсы" / "2026-10-12" / "Armtek"
    _write(day, "a.csv", "Бренд;Артикул;Цена\nBrembo;A1;100\n")
    src = scan([tmp_path / "Прайсы"], [tmp_path / "Прайсы"])
    (pf,) = prices.load_price_files(src.price_sheets, {}, {}, ["Armtek"])
    assert pf.shot_date == "12.10.2026"
    assert pf.supplier == "Armtek"
    assert pf.id.startswith("12.10.2026 / Armtek /")


def test_currency_file_is_skipped_and_parsed(tmp_path):
    root = tmp_path / "Прайсы"
    (root / "валюта.csv").parent.mkdir(parents=True)
    (root / "валюта.csv").write_text(
        ";12.10.2026;13.10.2026\nUSD;89,15;88,1111\nEUR;99,7232;98,7373\n",
        encoding="utf-8",
    )
    _write(root / "12.10.2026" / "TISS", "a.csv", "Бренд;Артикул;Цена\nKashiyama;K1;1000\n")
    src = scan([root], [root])
    assert [s.file.name for s in src.price_sheets] == ["a.csv"]
    cur = dynamics.read_currency(root)
    assert cur["rates"]["USD"]["12.10.2026"] == 89.15
    assert cur["rates"]["EUR"]["13.10.2026"] == 98.7373
    again = dynamics.read_currency(root)
    assert again["cached"] is True


def test_price_change_is_grouped_and_the_second_call_is_cached(tmp_path):
    root = tmp_path / "Прайсы"
    (root / "валюта.csv").parent.mkdir(parents=True)
    (root / "валюта.csv").write_text(";12.10.2026;13.10.2026\nUSD;100;90\n", encoding="utf-8")
    text = "Бренд;Артикул;Цена\nKashiyama;K1;{price}\nValeo;V1;{price}\n"
    _write(root / "12.10.2026" / "TISS", "a.csv", text.format(price=1000))
    _write(root / "13.10.2026" / "TISS", "a.csv", text.format(price=1100))
    sheets = scan([root], [root]).price_sheets
    files = prices.load_price_files(sheets, {}, {}, ["TISS"])
    src = Sources({}, [])
    src.pst = [SimpleNamespace(
        brand="ADVICS", article="OUR1", category="ТОРМОЗА", group="Тормозные диски", subgroup="Передние",
    )]
    src.crosses = {
        (norm_text("ADVICS"), norm_article("OUR1")): {
            "KASHIYAMA": [CrossRow("K1", norm_article("K1"), "Kashiyama", True)],
        },
    }
    first = dynamics.build(src, files, R._builtin_positioning(), root, None, None)
    assert first["date_from"] == "12.10.2026" and first["date_to"] == "13.10.2026"
    assert first["cached"] is False
    row = next(r for r in first["rows"] if r["brand"] == "Kashiyama" and not r["subgroup"])
    assert abs(row["change"] - 0.1) < 1e-9
    assert all(r["brand"] != "Valeo" for r in first["rows"])
    assert any("Kashiyama" in line for line in first["conclusions"])
    usd = next(r for r in first["rates"] if r["code"] == "USD")
    assert abs(usd["change"] - (-0.1)) < 1e-9
    second = dynamics.build(src, files, R._builtin_positioning(), root, "12.10.2026", "13.10.2026")
    assert second["cached"] is True
    assert second["compared"] == 1
    dynamics._MEM_RESULT.clear()
    dynamics._MEM_INDEX.clear()
    third = dynamics.build(src, files, R._builtin_positioning(), root, "12.10.2026", "13.10.2026")
    assert third["cached"] is True
