"""Распознавание прайсов на формах настоящих поставщиков: Ivers, Форум-Авто, Armtek, MXGroup."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "Программа"))

from pricing import prices  # noqa: E402
from pricing.sources import scan  # noqa: E402

BRANDS = ["ABA", "ACDELCO", "AKEBONO", "BREMBO", "SNR", "NTN", "STELLOX", "TOTEM", "SORL", "FEBEST"]


def _sheets(tmp_path, supplier, name, text):
    base = tmp_path / "Прайсы"
    root = base / supplier
    root.mkdir(parents=True, exist_ok=True)
    (root / name).write_text(text, encoding="utf-8")
    return scan([base], [base]).price_sheets


def _load(sheets, profiles=None, files=None, pairs=None):
    keys = {prices.brand_key(b) for b in BRANDS}
    return prices.load_price_files(sheets, profiles or {}, files or {}, [], BRANDS,
                                   known_brands=keys, cross_pairs=pairs or set())


def test_ivers_header_is_recognised_with_part_number_as_article(tmp_path):
    head = "Произв.;Код;Товар;Применяемость;Кросс-номера;Номер запчасти;Ост.;Цена;Партионность\n"
    rows = "".join(f"SNR;5426{i:02d};Подшипник;Note '05- Rear;R15{i}.89;FC.4{i:04d}.S0{i % 9};{i % 7 + 1};{1500 + i};1\n"
                   for i in range(30))
    (pf,) = _load(_sheets(tmp_path, "Ivers", "IversMSK.csv", head + rows))
    assert pf.status == "auto" and pf.header.row == 0
    assert pf.column_titles() == {"article": "Номер запчасти", "price": "Цена", "brand": "Произв.",
                                  "qty": "Ост.", "name": "Товар"}
    assert len(pf.offers) == 30 and pf.offers[0].article == "FC.40000.S00"


def test_article_column_with_more_cross_hits_is_pointed_out(tmp_path):
    head = "Произв.;Код;Товар;Применяемость;Кросс-номера;Номер запчасти;Ост.;Цена\n"
    rows = "".join(f"SNR;5426{i:02d};Подшипник;x;R15{i}.89;FC.4{i:04d}.S0;2;{1500 + i}\n" for i in range(30))
    sheets = _sheets(tmp_path, "Ivers", "IversKRD.csv", head + rows)
    pairs = {("SNR", prices.norm_article(f"R15{i}.89")) for i in range(30)}
    # Пользователь указал артикулом «Номер запчасти», а с кроссами совпадает «Кросс-номера».
    (pf,) = _load(sheets, pairs=pairs)
    assert any("«Кросс-номера» совпадает с кроссами чаще" in n for n in pf.notes)


def test_forum_xlsx_model_column_is_not_taken_for_brand(tmp_path):
    head = "ФОРУМ-АВТО;;\n;тел. 789-8000;\nГРУППА;№ ПРОИЗВ.;НАИМЕНОВАНИЕ;МОДЕЛЬ;ЦЕНА, РУБ;НАЛичие\n"
    models = ["CITROEN C5, PEUGEOT 407", "FIAT Ducato 2.3JTD 02->", "TOYOTA Camry 06->", "LAND ROVER Discovery 2.7D 02->"]
    brands = ["ABA", "ACDELCO", "AKEBONO", "BREMBO", "ABA"]
    rows = "".join(f"{brands[i % 5]};YM40{i:04d};Ролик;{models[i % 4]};{150 + i};2\n" for i in range(40))
    sheets = _sheets(tmp_path, "Форум-Авто", "Форум-АвтоIRK.csv", head + rows)
    # Сохранённая форма с ошибкой: брендом указан «МОДЕЛЬ».
    bad = {"x": {"columns": {"brand": "МОДЕЛЬ", "article": "№ ПРОИЗВ.", "price": "ЦЕНА, РУБ", "qty": "НАЛичие"}}}
    (pf,) = _load(sheets, profiles=bad)
    assert pf.status == "saved"
    assert pf.column_titles()["brand"] == "ГРУППА"
    assert pf.offers[0].brand == "ABA" and any("модели машин" in n for n in pf.notes)
    # Без сохранённой формы шапка узнаётся сама: «№ ПРОИЗВ.» — это артикул.
    (auto,) = _load(sheets)
    assert auto.column_titles()["article"] == "№ ПРОИЗВ." and auto.column_titles()["brand"] == "ГРУППА"


def test_headerless_file_is_read_from_the_first_row(tmp_path):
    rows = "".join(f"STELLOX;0000{i:02d}BSX;колодки;0000{i:02d}B_SX;{i % 9 + 2};{800 + i}.50\n" for i in range(40))
    sheets = _sheets(tmp_path, "Armtek", "ArmtekMSK.csv", rows)
    # Пользователь отметил шапкой 5-ю строку — это уже строка с ценой.
    files = {prices.price_id(sheets[0]): {"header_row": 4, "cols": {"brand": 0, "article": 1, "qty": 4, "price": 5}}}
    (pf,) = _load(sheets, files=files)
    assert pf.headerless and pf.data_from == 0
    assert len(pf.offers) == 40 and pf.offers[0].article == "000000BSX" and pf.offers[0].row == 1
    assert prices.file_shape(pf).startswith("без шапки, цены с 1-й строки")
    # «Названия» из строки с ценой не становятся сохранённой формой для других файлов.
    assert prices.profiles_from_manual(sheets, files, []) == {}


def test_quantity_that_is_only_ones_is_taken_for_pack_size(tmp_path):
    stock = [6, 46, 10, 50, 1, 3, 12, 7]
    rows = "".join(f"ABA;7PK1{i:03d};Ремень;{280 + i};{stock[i % 8]};1;3180{i:04d};ДА\n" for i in range(40))
    sheets = _sheets(tmp_path, "Форум-Авто", "Форум-АвтоKRD.csv", rows)
    layout = {"header_row": 17, "cols": {"brand": 0, "article": 1, "price": 3, "qty": 5}}
    files = {prices.price_id(sheets[0]): dict(layout)}
    (manual,) = _load(sheets, files=files)
    assert manual.cols["qty"] == 5 and any("похоже на кратность" in n for n in manual.notes)
    # Та же форма, перенесённая на следующий день, берёт остаток из E.
    other = _sheets(tmp_path / "next", "Форум-Авто", "Форум-АвтоKRD.csv", rows)
    pf = prices.load_one(other[0], {}, {}, [], BRANDS, {"форумавто": layout})
    assert pf.status == "saved" and pf.cols["qty"] == 4 and pf.offers[0].qty == 6
    assert len(pf.offers) == 40


def test_english_header_is_recognised(tmp_path):
    rows = "articul;brand;fullname;stock;price\n" + "".join(
        f"F100{i:02d};16F;Защитный комплект;10;{694 + i}.98\n" for i in range(20))
    (pf,) = _load(_sheets(tmp_path, "MXGroup", "MXGroupVLD.csv", rows))
    assert pf.header.row == 0 and pf.column_titles()["article"] == "articul" and len(pf.offers) == 20


def test_numeric_brand_names_stay(tmp_path):
    rows = "Номенклатура;Бренд;Артикул;Описание;Цена, руб.;Наличие\n" + "".join(
        f"NSIN{i:010d};555;SB1{i:03d};Опора шаровая;{1500 + i};2\n" for i in range(30))
    (pf,) = _load(_sheets(tmp_path, "Rossko", "RosskoBLG.csv", rows))
    assert pf.column_titles()["brand"] == "Бренд" and not pf.notes


# ---------------------------------------------------------------- подсказки брендов


def test_brand_hints_match_whole_words_only(tmp_path):
    from pricing import price_check

    rows = "Бренд;Артикул;Цена\n" + "".join(
        f"{b};A{i};{100 + i}\n" for i, b in enumerate(
            ["HERMANN", "WUHLERMANN", "TOYOTA Camry 06->", "CITROEN C5, PEUGEOT 407", "Kashiyama Japan",
             "G-Brake Germany", "BREMBBO", "Brembo"] * 3))
    (pf,) = _load(_sheets(tmp_path, "TISS", "TissEKB.csv", rows))
    hints = price_check.brand_hints([pf], ["MANN", "TOYOTA", "Kashiyama", "G-Brake", "Brembo"])
    got = {h["brand"]: h["suggest"] for h in hints}
    assert got == {"Kashiyama Japan": "Kashiyama", "G-Brake Germany": "G-Brake", "BREMBBO": "Brembo"}
    ignore = {prices.raw_brand_key("BREMBBO")}
    again = price_check.brand_hints([pf], ["MANN", "TOYOTA", "Kashiyama", "G-Brake", "Brembo"], ignore=ignore)
    assert "BREMBBO" not in {h["brand"] for h in again}


def test_brand_aliases_api_takes_a_batch_and_an_ignore_list(tmp_path, monkeypatch):
    from pricing import server, settings as settings_mod

    monkeypatch.setattr(settings_mod, "USER_FILE", tmp_path / "настройки.json")
    settings_mod._drop_loads()
    try:
        r = server.api_brand_aliases({"add_many": [["Kashiyama Japan", "Kashiyama"], ["G-Brake Germany", "G-Brake"]],
                                      "ignore_many": ["BREMBBO"]})
        assert r["added"] == 2 and r["brand_ignore"] == ["BREMBBO"]
        assert prices.brand_key("G-Brake Germany") == prices.brand_key("G-Brake")
        r = server.api_brand_aliases({"unignore": "brembbo"})
        assert r["brand_ignore"] == [] and len(r["brand_aliases"]) == 2
    finally:
        prices.set_brand_aliases([])
        settings_mod._drop_loads()


# ---------------------------------------------------------------- город из имени файла


def test_city_from_russian_names_brackets_and_user_codes():
    cases = {
        "Прайс Армтек_Иркутск.xlsx": "Иркутск", "Прайс Армтек_ЕКБ.xlsx": "Екатеринбург",
        "Прайс Армтек_МСК_РПП.xlsx": "Москва", "FORUM_AUTO_PRICE_NVS.csv": "Новосибирск",
        "FORUM_AUTO_PRICE_VLK3.csv": "Владивосток", "Armtek(Омск).xlsx": "Омск", "BergIRK(Чита).xlsx": "Чита",
        "Shate-MKRS.csv": "Красноярск", "TissEKB.xls": "Екатеринбург", "FORUM_AUTO_PRICE.csv": "",
    }
    assert {n: prices.city_from_name(n) for n in cases} == cases
    try:
        prices.set_city_codes([["RZD", "Раздольное"]])
        assert prices.city_from_name("FORUM_AUTO_RZD.csv") == "Раздольное" and prices.city_token()
    finally:
        prices.set_city_codes([])
    assert prices.city_from_name("FORUM_AUTO_RZD.csv") == ""


# ---------------------------------------------------------------- правка одного прайса


def test_saving_one_price_rereads_only_that_supplier(tmp_path, monkeypatch):
    from test_speed import _workspace

    server, _ = _workspace(tmp_path, monkeypatch)
    other = tmp_path / "Данные" / "Прайсы" / "06.10.2026" / "Berg"
    other.mkdir(parents=True)
    (other / "BergEKB.csv").write_text("Бренд;Артикул;Цена\nBrembo;B1;1200\nAkebono;K1;600\n", encoding="utf-8")
    server.api_scan({"folder": str(tmp_path)})
    src = server.STATE.for_group(tmp_path, "", "")
    from pricing.settings import Settings

    before = {pf.supplier: pf for pf in server.STATE.price_files(src, Settings.load())}
    tiss = next(f for f in before.values() if f.supplier == "TISS")
    calls = []
    real = prices.load_one
    monkeypatch.setattr(prices, "load_one", lambda sheet, *a, **k: calls.append(sheet.file.name) or real(sheet, *a, **k))
    preview = server.api_prices_save({"folder": str(tmp_path), "id": tiss.id, "city": "Москва", "dry_run": True})
    assert preview["file"]["preview_only"] and preview["file"]["city"] == "Москва"
    calls.clear()
    server.api_prices_save({"folder": str(tmp_path), "id": tiss.id, "city": "Москва"})
    assert "BergEKB.csv" not in calls and "tiss.csv" in calls
    after = {pf.supplier: pf for pf in server.STATE.price_files(src, Settings.load())}
    assert after["Berg"] is before["Berg"] and after["TISS"].city == "Москва"


def test_city_set_by_hand_carries_to_the_same_file_on_next_dates(tmp_path):
    base = tmp_path / "Прайсы"
    body = "Бренд;Артикул;Цена\nBrembo;B1;100\nBrembo;B2;200\n"
    for date, names in (("14.10.2026", ["118482.csv", "BERG_X_20261014_040019.csv"]),
                        ("15.10.2026", ["118482.csv", "BERG_X_20261015_041120.csv", "43256.csv"])):
        folder = base / date / "Склад"
        folder.mkdir(parents=True)
        for name in names:
            (folder / name).write_text(body, encoding="utf-8")
    sheets = scan([base], [base]).price_sheets
    old = [s for s in sheets if "14.10.2026" in str(s.file)]
    files = {prices.price_id(s): {"city": "Южно-Сахалинск" if s.file.name.startswith("118482") else "Чита"} for s in old}
    loaded = prices.load_price_files(sheets, {}, files, [], BRANDS)
    got = {(pf.shot_date, pf.sheet.file.name): pf.city for pf in loaded}
    assert got[("15.10.2026", "118482.csv")] == "Южно-Сахалинск"
    assert got[("15.10.2026", "BERG_X_20261015_041120.csv")] == "Чита"
    assert got[("15.10.2026", "43256.csv")] == ""
    assert prices.stem_key("price_all2_261015_1.xlsx") == "priceall2"
