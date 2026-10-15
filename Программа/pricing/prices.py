"""Прайсы поставщиков: у каждого свой формат, поэтому столбцы ищутся по смыслу,
а если не угадались — пользователь один раз указывает их сам. Указанное сохраняется
как «формат» (по названиям столбцов) и применяется ко всем похожим файлам,
например к прайсам того же поставщика по другим городам.
"""

from __future__ import annotations

import hashlib
import pickle
import re
from dataclasses import dataclass, field
from pathlib import Path

from .bootlog import step as boot_step
from .excel import Header, Sheet, clean, num
from .filecache import CACHE_DIR
from .sources import norm_article, norm_text

FIELDS = {
    "brand": "Бренд",
    "article": "Артикул",
    "price": "Цена",
    "qty": "Наличие",
    "name": "Наименование",
}
REQUIRED = ("article", "price")
PATTERNS = {
    "brand": [r"бренд", r"производител", r"изготовител", r"^тм$", r"^tm$", r"торговая марка", r"^марка",
              r"^фирма", r"brand", r"^maker", r"^завод"],
    "article": [r"^катал", r"^артикул$", r"^артикул", r"номер детали", r"код производителя", r"каталожный",
                r"part\s*(no|number|#)", r"^article", r"^номер$", r"^арт\.?$", r"^код детали"],
    "price": [r"^опт", r"^цена$", r"^цена", r"цена", r"^price", r"стоимость", r"^руб"],
    "qty": [r"^кол-?во", r"количество", r"^остаток", r"наличие", r"^qty", r"stock", r"^склад", r"в наличии"],
    "name": [r"^наименование", r"^название", r"^описание", r"^name", r"^товар"],
}
SCAN_ROWS = 60
SAMPLE_ROWS = 300
CITY_CODES = {
    "EKB": "Екатеринбург", "SPB": "Санкт-Петербург", "MSK": "Москва", "NSK": "Новосибирск", "OMSK": "Омск",
    "KRD": "Краснодар", "KZN": "Казань", "VVO": "Владивосток", "KHV": "Хабаровск", "IRK": "Иркутск",
    "KRS": "Красноярск", "SAM": "Самара", "TMN": "Тюмень", "CHL": "Челябинск", "PRM": "Пермь", "UFA": "Уфа",
    "NN": "Нижний Новгород", "RND": "Ростов-на-Дону", "VRN": "Воронеж", "BRN": "Барнаул",
}


def brand_key(value) -> str:
    return re.sub(r"[^0-9A-ZА-Я]", "", norm_text(value))


def supplier_key(value) -> str:
    return re.sub(r"[^0-9a-zа-я]", "", clean(value).lower().replace("ё", "е"))


@dataclass
class Offer:
    brand: str
    article: str
    price: float
    qty: float | None
    row: int
    article_key: str = ""  # артикул без знаков, по нему ищем кросс


@dataclass
class PriceFile:
    sheet: Sheet
    supplier: str
    city: str
    header: Header | None
    cols: dict[str, int]
    fixed_brand: str
    status: str  # auto | saved | manual | none
    enabled: bool = True
    offers: list[Offer] = field(default_factory=list)
    problem: str = ""

    folder: str = ""  # папка внутри «Прайсы», если прайс лежит у поставщика
    name_supplier: str = ""  # как поставщик читается из имени файла
    supplier_from: str = "file"  # folder | file | manual
    mismatch: bool = False  # папка и имя файла указывают на разных известных поставщиков
    checks: dict = field(default_factory=dict)  # доля чисел в столбцах цены и наименования
    doubt: str = ""  # столбец цены похож на текст или наименование похоже на числа
    shot_date: str = ""  # папка-дата внутри «Прайсы», если прайсы лежат по дням
    why: str = ""  # почему шапка или бренд не узнались сами

    @property
    def id(self) -> str:
        return price_id(self.sheet)

    @property
    def supplier_label(self) -> str:
        return f"{self.supplier} · {self.city}" if self.city else self.supplier

    def column_titles(self) -> dict[str, str]:
        if not self.header:
            return {}
        return {k: self.header.raw[i] for k, i in self.cols.items()}


# ---------------------------------------------------------------- угадывание


_PRICE_ROOTS = {"прайсы", "prices"}
_DATE_DIR = re.compile(r"^(\d{2})\.(\d{2})\.(\d{4})$|^(\d{4})-(\d{2})-(\d{2})$|^(\d{2})-(\d{2})-(\d{4})$")


def _prices_rest(path: Path) -> tuple[str, ...] | None:
    """Части пути после папки «Прайсы», вместе с именем файла."""
    parts = Path(path).parts
    for i, part in enumerate(parts):
        if part.casefold() in _PRICE_ROOTS:
            return parts[i + 1:]
    return None


def _normalize_date(name: str) -> str:
    """«2026-10-12» и «12-10-2026» -> «12.10.2026». Уже такая запись остаётся."""
    m = _DATE_DIR.match(name)
    if not m:
        return ""
    if m.group(1):
        return f"{m.group(1)}.{m.group(2)}.{m.group(3)}"
    if m.group(4):
        return f"{m.group(6)}.{m.group(5)}.{m.group(4)}"
    return f"{m.group(7)}.{m.group(8)}.{m.group(9)}"


def _shot_sort(text: str) -> tuple[int, int, int]:
    parts = (text or "").split(".")
    if len(parts) == 3 and all(part.isdigit() for part in parts):
        day, month, year = parts
        return (int(year), int(month), int(day))
    return (0, 0, 0)


def latest_price_files(files: list[PriceFile]) -> list[PriceFile]:
    """Для отчёта — только последняя папка-дата. Старые даты нужны динамике, не каждой проценке.

    Если дат нет, остаются прайсы, которые лежат прямо в папке поставщика.
    """
    enabled = [pf for pf in files if pf.enabled]
    dated = [pf.shot_date for pf in enabled if pf.shot_date]
    if not dated:
        return enabled
    latest = max(dated, key=_shot_sort)
    return [pf for pf in enabled if pf.shot_date == latest]


def price_date(path: Path) -> str:
    """«Прайсы/12.10.2026/Armtek/файл.xlsx» -> «12.10.2026». Без папки-даты — пусто."""
    rest = _prices_rest(path)
    if not rest or len(rest) < 2:
        return ""
    return _normalize_date(rest[0])


def supplier_folder(path: Path) -> str:
    """«Прайсы/Armtek/окт.csv» и «Прайсы/12.10.2026/Armtek/окт.csv» -> «Armtek».

    Папка с датой — снимок на день, поставщик лежит внутри неё. Файл прямо в «Прайсы»
    или прямо в папке-дате поставщика из папки не получает.
    """
    rest = _prices_rest(path)
    if not rest or len(rest) < 2:
        return ""
    idx = 1 if _normalize_date(rest[0]) else 0
    if idx >= len(rest) - 1:
        return ""
    name = rest[idx]
    if name.startswith((".", "~$")):
        return ""
    return name


def price_id(sheet: Sheet) -> str:
    """Одинаковые имена файлов у разных поставщиков и дат не схлопываются в одну настройку."""
    folder = supplier_folder(sheet.file)
    date = price_date(sheet.file)
    if not date and not folder:
        return sheet.label
    return " / ".join(bit for bit in (date, folder, sheet.label) if bit)


def match_known(text: str, known: list[str]) -> str:
    """Самое длинное имя поставщика из позиционирования, которое входит в текст."""
    key = supplier_key(text)
    if not key:
        return ""
    for name in sorted(known, key=len, reverse=True):
        sk = supplier_key(name)
        if sk and sk in key:
            return name
    return ""


def city_from_name(name: str) -> str:
    """Город только по известному коду в имени: EKB, SPB, NSK. Случайное слово городом не становится."""
    letters = re.sub(r"[^A-Za-z]", " ", Path(name).stem).upper().split()
    for token in letters:
        if token in CITY_CODES:
            return CITY_CODES[token]
    return ""


def guess_supplier(name: str, known: list[str]) -> tuple[str, str]:
    """«TissEKB.xls» -> («TISS», «Екатеринбург»), «RosskoSPB1» -> («Rossko», «Санкт-Петербург»)."""
    stem = Path(name).stem
    supplier = match_known(stem, known)
    rest = stem
    if supplier:
        rest = re.sub(re.escape(supplier), "", stem, flags=re.IGNORECASE) if supplier.lower() in stem.lower() else stem
    if not supplier:
        supplier = re.split(r"[\s_\-.]+|(?<=[a-zа-я])(?=[A-ZА-Я]{2,})", stem)[0] or stem
        rest = stem[len(supplier):]
    city = city_from_name(rest)
    if not city:
        letters = re.sub(r"[^A-Za-z]", " ", rest).upper().split()
        if letters:
            city = letters[0]
    return supplier, city


def _numeric_share(sheet: Sheet, start: int, col: int, limit: int = SAMPLE_ROWS) -> float:
    rows = sheet.grid[start : start + limit]
    values = [r[col] for r in rows if col < len(r) and clean(r[col])]
    if not values:
        return 0.0
    return sum(num(v) is not None for v in values) / len(values)


def _filled_share(sheet: Sheet, start: int, col: int) -> float:
    rows = sheet.grid[start : start + SAMPLE_ROWS]
    if not rows:
        return 0.0
    return sum(1 for r in rows if col < len(r) and clean(r[col])) / len(rows)


def detect(sheet: Sheet, rows: range | None = None) -> tuple[Header | None, dict[str, int]]:
    """Ищет строку шапки, где есть артикул и цена; выбирает столбцы по названию и содержимому."""
    best: tuple[float, Header, dict[str, int]] | None = None
    for start in rows if rows is not None else range(min(SCAN_ROWS, len(sheet.grid))):
        if start < 0 or start >= len(sheet.grid):
            continue
        raw = [clean(v) for v in sheet.grid[start]]
        if sum(1 for v in raw if v) < 2:
            continue
        keys = [v.lower().replace("ё", "е") for v in raw]
        header = Header(sheet, start, 1, raw, keys)
        cols: dict[str, int] = {}
        for fld in ("article", "price", "brand", "qty", "name"):
            candidates = []
            for rank, pattern in enumerate(PATTERNS[fld]):
                rx = re.compile(pattern)
                for idx, k in enumerate(keys):
                    if k and rx.search(k) and idx not in cols.values():
                        candidates.append((rank, idx))
            if not candidates:
                continue
            data_start = start + 1
            if fld == "price":
                # Название «наименование» не берём за цену, даже если в столбце случайно много чисел.
                priced = [(i, rank) for rank, i in candidates if not _name_title(raw[i])]
                if not priced:
                    continue
                scored = [(-_numeric_share(sheet, data_start, i, 40), rank, i) for i, rank in priced]
                share, _, idx = min(scored)
                # Цена — это числа. Текстовый столбец, даже со словом «цена» в шапке, не берём.
                if -share < 0.8:
                    continue
            elif fld == "qty":
                scored = [(rank, -_filled_share(sheet, data_start, i), i) for rank, i in candidates]
                idx = min(scored)[2]
            else:
                scored = [(rank, -_filled_share(sheet, data_start, i), i) for rank, i in candidates]
                idx = min(scored)[2]
                if fld == "name" and _numeric_share(sheet, data_start, idx, 40) >= 0.5:
                    # Столбец из чисел — не наименование. Так цена не попадает в название.
                    continue
            cols[fld] = idx
        if all(f in cols for f in REQUIRED):
            # Шапка чаще всего во 2-й или 3-й строке: сверху название прайса, ниже столбцы.
            bonus = {1: 0.55, 2: 0.4, 0: 0.1}.get(start, 0)
            score = len(cols) + _filled_share(sheet, start + 1, cols["article"]) + bonus
            if best is None or score > best[0]:
                best = (score, header, cols)
    if not best:
        return None, {}
    return best[1], best[2]


_NAME_TITLE = re.compile(r"наименован|название|описание|^name$|^товар")


def _name_title(title: str) -> bool:
    return bool(_NAME_TITLE.search(clean(title).lower().replace("ё", "е")))


def apply_profile(sheet: Sheet, profile: dict) -> tuple[Header | None, dict[str, int]]:
    """Профиль — названия столбцов {поле: «название»}; ищем строку шапки, где они все есть."""
    wanted = {k: v for k, v in (profile.get("columns") or {}).items() if v}
    if not all(f in wanted for f in REQUIRED):
        return None, {}
    for start in range(min(SCAN_ROWS, len(sheet.grid))):
        raw = [clean(v) for v in sheet.grid[start]]
        norm = [norm_text(v) for v in raw]
        cols = {}
        for fld, title in wanted.items():
            t = norm_text(title)
            if t in norm:
                cols[fld] = norm.index(t)
        if all(f in cols for f in REQUIRED):
            keys = [v.lower() for v in raw]
            return Header(sheet, start, 1, raw, keys), cols
    return None, {}


def header_at(sheet: Sheet, row: int) -> Header | None:
    if row < 0 or row >= len(sheet.grid):
        return None
    raw = [clean(v) for v in sheet.grid[row]]
    return Header(sheet, row, 1, raw, [v.lower() for v in raw])


# ---------------------------------------------------------------- чтение строк


def _qty(value) -> float | None:
    n = num(value)
    if n is not None:
        return n
    text = clean(value)
    if not text:
        return 0.0
    m = re.search(r"\d+(?:[.,]\d+)?", text)
    if m:
        return float(m.group(0).replace(",", "."))
    return None


OFFER_VERSION = 1
_OFFER_MEM: dict[str, list[tuple]] = {}


def _offer_token(path: Path, header_row: int, cols: dict, fixed_brand: str) -> str:
    st = path.stat()
    raw = f"{OFFER_VERSION}|{path.resolve()}|{st.st_mtime_ns}|{st.st_size}|{header_row}|{sorted(cols.items())}|{fixed_brand}"
    return hashlib.sha1(raw.encode()).hexdigest()


def _load_offers(token: str) -> list[Offer] | None:
    packed = _OFFER_MEM.get(token)
    if packed is None:
        path = CACHE_DIR / f"off-{token}.pkl"
        if not path.is_file():
            return None
        try:
            with path.open("rb") as fh:
                blob = pickle.load(fh)
        except (OSError, pickle.PickleError, EOFError, ValueError):
            return None
        if not isinstance(blob, dict) or blob.get("version") != OFFER_VERSION or blob.get("token") != token:
            return None
        packed = blob.get("rows")
        if not isinstance(packed, list):
            return None
        _OFFER_MEM[token] = packed
    return [Offer(*row) for row in packed]


def _store_offers(token: str, offers: list[Offer]) -> None:
    packed = [(o.brand, o.article, o.price, o.qty, o.row, o.article_key) for o in offers]
    _OFFER_MEM[token] = packed
    try:
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        path = CACHE_DIR / f"off-{token}.pkl"
        tmp = path.with_suffix(".tmp")
        with tmp.open("wb") as fh:
            pickle.dump({"version": OFFER_VERSION, "token": token, "rows": packed}, fh, protocol=pickle.HIGHEST_PROTOCOL)
        tmp.replace(path)
    except (OSError, pickle.PickleError, TypeError, ValueError):
        pass


def _layout_works(sheet: Sheet, header: Header, cols: dict[str, int]) -> bool:
    if not header or not all(f in cols for f in REQUIRED):
        return False
    hits = 0
    for row in header.data[:40]:
        art = clean(row[cols["article"]]) if cols["article"] < len(row) else ""
        price = num(row[cols["price"]]) if cols["price"] < len(row) else None
        if art and price and price > 0 and norm_text(art) != norm_text(header.raw[cols["article"]]):
            hits += 1
            if hits >= 2:
                return True
    return False


def apply_supplier_layout(sheet: Sheet, layout: dict) -> tuple[Header | None, dict[str, int]]:
    """Те же номера столбцов, что уже указали у этого поставщика. Строка шапки — та же, либо соседняя."""
    base = int(layout["header_row"])
    cols = {k: int(v) for k, v in (layout.get("cols") or {}).items() if k in FIELDS}
    for delta in (0, -1, 1, -2, 2):
        row = base + delta
        if row < 0:
            continue
        header = header_at(sheet, row)
        if header and _layout_works(sheet, header, cols):
            return header, dict(cols)
    return None, {}


def explain_header(sheet: Sheet) -> str:
    """Коротко, почему шапку пришлось бы указывать руками: что лежит в первых строках."""
    bits = []
    for start in range(min(4, len(sheet.grid))):
        raw = [clean(v) for v in sheet.grid[start]]
        filled = [v for v in raw if v][:6]
        if not filled:
            bits.append(f"строка {start + 1} пустая")
        else:
            bits.append(f"строка {start + 1}: " + ", ".join(v[:28] for v in filled))
    return "Шапка не узнана сама. " + "; ".join(bits)


def read_offers(pf: PriceFile) -> None:
    pf.offers = []
    if not pf.header or not all(f in pf.cols for f in REQUIRED):
        return
    c = pf.cols
    for i, row in enumerate(pf.header.data):
        art = clean(row[c["article"]]) if c["article"] < len(row) else ""
        price = num(row[c["price"]]) if c["price"] < len(row) else None
        key = norm_article(art)
        if not key or not price or price <= 0 or norm_text(art) == norm_text(pf.header.raw[c["article"]]):
            continue
        brand = clean(row[c["brand"]]) if "brand" in c and c["brand"] < len(row) else ""
        brand = brand or pf.fixed_brand
        qty = _qty(row[c["qty"]]) if "qty" in c and c["qty"] < len(row) else None
        pf.offers.append(Offer(brand, art, price, qty, pf.header.row + pf.header.depth + i + 1, key))


def _manual(sheet: Sheet, own: dict) -> tuple[Header | None, dict[str, int]]:
    """Ручная настройка файла: строка шапки и номера столбцов; если столбцы не заданы — ищем их в этой строке."""
    row = own.get("header_row")
    if row is None:
        return None, {}
    cols = {k: int(v) for k, v in (own.get("cols") or {}).items() if k in FIELDS and v not in (None, "")}
    if not cols:
        found, cols = detect(sheet, range(row, row + 1))
        if found:
            return found, cols
    return header_at(sheet, row), cols


def guess_brand_column(sheet: Sheet, header: Header, cols: dict[str, int], hints: list[str] | None) -> int | None:
    """Если в шапке нет слова «бренд», ищем столбец, где лежат известные бренды конкурентов."""
    hint_keys = {brand_key(h) for h in hints or [] if brand_key(h)}
    if not hint_keys:
        return None
    start = header.row + header.depth
    best: tuple[float, int] | None = None
    for idx in range(len(header.raw)):
        if idx in cols.values() or _name_title(header.raw[idx]):
            continue
        if _numeric_share(sheet, start, idx, 40) > 0.35:
            continue
        hits = samples = 0
        for row in sheet.grid[start : start + 80]:
            if idx >= len(row):
                continue
            text = clean(row[idx])
            if not text:
                continue
            samples += 1
            if brand_key(text) in hint_keys:
                hits += 1
        if samples >= 3 and hits >= 2 and hits / samples >= 0.2:
            score = hits / samples
            if best is None or score > best[0]:
                best = (score, idx)
    return None if best is None else best[1]


def apply_first_profile(sheet: Sheet, profiles: dict) -> tuple[Header | None, dict[str, int]]:
    """Все сохранённые шапки проверяются за один проход по строкам, сначала 2-я и 3-я."""
    prepared = []
    for profile in profiles.values():
        wanted = {k: norm_text(v) for k, v in (profile.get("columns") or {}).items() if v and k in FIELDS}
        if all(f in wanted for f in REQUIRED):
            prepared.append(wanted)
    if not prepared:
        return None, {}
    rows = []
    for start in range(min(SCAN_ROWS, len(sheet.grid))):
        raw = [clean(v) for v in sheet.grid[start]]
        if sum(bool(v) for v in raw) < 2:
            continue
        rows.append((start, raw, [norm_text(v) for v in raw]))
    rows.sort(key=lambda item: (0 if item[0] in (1, 2) else 1, item[0]))
    for start, raw, norm in rows:
        for wanted in prepared:
            cols = {}
            for fld, title in wanted.items():
                if title in norm:
                    cols[fld] = norm.index(title)
            if all(f in cols for f in REQUIRED):
                return Header(sheet, start, 1, raw, [v.lower() for v in raw]), cols
    return None, {}


def load_one(sheet: Sheet, own: dict, profiles: dict, known_suppliers: list[str],
             competitor_brands: list[str] | None = None, supplier_layouts: dict | None = None) -> PriceFile:
    folder = supplier_folder(sheet.file)
    name_supplier, city = guess_supplier(sheet.file.name, known_suppliers)
    folder_name = match_known(folder, known_suppliers) or clean(folder) if folder else ""
    if own.get("supplier"):
        supplier = own.get("supplier")
        supplier_from = "manual"
    elif folder_name:
        supplier = folder_name
        supplier_from = "folder"
        # Город — только узнанный код (EKB, NSK). «price.xlsx» не становится городом PRICE.
        city = city if city in CITY_CODES.values() else ""
    else:
        supplier = name_supplier
        supplier_from = "file"
    city = own.get("city", city)
    file_known = match_known(sheet.file.stem, known_suppliers)
    mismatch = bool(folder_name and file_known and supplier_key(folder_name) != supplier_key(file_known))
    header, cols, status = None, {}, "none"
    manual = own.get("header_row") is not None
    if manual:
        header, cols = _manual(sheet, own)
        status = "manual" if header else "none"
    elif own.get("columns"):
        header, cols = apply_profile(sheet, own)
        status = "manual" if header else "none"
    if not header and not manual and profiles:
        boot_step("Ищу шапку из настроек", sheet.file.name)
        header, cols = apply_first_profile(sheet, profiles)
        if header:
            status = "saved"
    if not header and not manual:
        header, cols = detect(sheet)
        status = "auto" if header else "none"
    layout = (supplier_layouts or {}).get(supplier_key(folder_name or supplier))
    if not header and not manual and layout:
        boot_step("Столбцы поставщика", f"{folder_name or supplier} · {sheet.file.name}")
        header, cols = apply_supplier_layout(sheet, layout)
        if header:
            status = "saved"
    if header and "brand" not in cols and not own.get("fixed_brand"):
        found = guess_brand_column(sheet, header, cols, competitor_brands)
        if found is not None:
            cols["brand"] = found
    why = ""
    if not header:
        why = explain_header(sheet)
        if not layout and folder_name:
            why += f" Для поставщика {folder_name} столбцы ещё не сохранены: после первого «Сохранить» они применятся к следующим дням."
        elif layout:
            why += f" Столбцы как в {layout.get('example') or 'сохранённом файле'} на строке {int(layout['header_row']) + 1} здесь не дали артикул и цену."
    pf = PriceFile(
        sheet, supplier, city, header, cols, own.get("fixed_brand", ""), status,
        enabled=own.get("enabled", True), folder=folder, name_supplier=name_supplier,
        supplier_from=supplier_from, mismatch=mismatch, shot_date=price_date(sheet.file), why=why,
    )
    if header and not all(f in cols for f in REQUIRED):
        pf.problem = "укажите столбцы артикула и цены"
    elif header and "brand" not in cols and not pf.fixed_brand:
        pf.problem = "нет столбца бренда — укажите бренд для всего прайса"
    if header and all(f in cols for f in REQUIRED):
        token = _offer_token(sheet.file, header.row, cols, pf.fixed_brand)
        cached = _load_offers(token)
        if cached is not None:
            pf.offers = cached
        else:
            read_offers(pf)
            _store_offers(token, pf.offers)
    pf.checks = _column_checks(pf)
    pf.doubt = _doubt(pf.checks)
    if header and not pf.offers:
        pf.problem = pf.problem or "не нашлось ни одной строки с артикулом и ценой"
    if not header and manual:
        pf.problem = f"в строке {own['header_row'] + 1} не нашлись столбцы артикула и цены — укажите их"
    elif not header:
        pf.problem = why or "не нашлись столбцы артикула и цены — укажите их вручную"
    return pf


def can_remember(pf: PriceFile) -> bool:
    titles = pf.column_titles()
    return bool(pf.header and pf.offers and all(titles.get(f) for f in REQUIRED))


def file_settings(files: dict, sheet: Sheet) -> dict:
    """Настройки файла. Старый ключ без папки поставщика тоже подходит."""
    own = files.get(price_id(sheet))
    if own is None and price_id(sheet) != sheet.label:
        own = files.get(sheet.label)
    return own or {}


def profiles_from_manual(sheets: list[Sheet], files: dict, known_suppliers: list[str]) -> dict:
    """Ручная настройка — названия столбцов из указанной строки шапки, без повторного чтения всех цен.

    Галочка «запомнить» для этого не нужна. Без названий в шапке перенос не делается.
    """
    boot_step("Настройки прайсов", "беру столбцы, которые уже указаны")
    learned = {}
    for sheet in sheets:
        own = file_settings(files, sheet)
        if own.get("header_row") is None or not own.get("cols"):
            continue
        header = header_at(sheet, int(own["header_row"]))
        if not header:
            continue
        titles = {}
        for key, idx in own["cols"].items():
            if key in FIELDS and isinstance(idx, int) and 0 <= idx < len(header.raw) and header.raw[idx]:
                titles[key] = header.raw[idx]
        if not all(titles.get(f) for f in REQUIRED):
            continue
        boot_step("Настройки прайсов", sheet.file.name)
        signature = " | ".join(f"{k}={norm_text(titles.get(k, ''))}" for k in ("brand", "article", "price"))
        learned.setdefault(signature, {"columns": titles, "example": sheet.file.name})
    return learned


def supplier_layouts_from(sheets: list[Sheet], files: dict) -> dict[str, dict]:
    """Номера столбцов, один раз указанные у поставщика. Действуют на все его даты."""
    found: dict[str, dict] = {}
    for sheet in sheets:
        own = file_settings(files, sheet)
        if own.get("header_row") is None or not own.get("cols"):
            continue
        name = own.get("supplier") or supplier_folder(sheet.file)
        key = supplier_key(name)
        if not key:
            continue
        found[key] = {
            "header_row": int(own["header_row"]),
            "cols": {k: int(v) for k, v in own["cols"].items() if k in FIELDS and isinstance(v, int)},
            "example": sheet.file.name,
        }
    return found


def load_price_files(sheets: list[Sheet], profiles: dict, files: dict, known_suppliers: list[str],
                      competitor_brands: list[str] | None = None) -> list[PriceFile]:
    """sheets — кандидаты в прайсы; profiles — сохранённые форматы; files — настройки конкретных файлов."""
    boot_step("Настройки прайсов", "собираю столбцы поставщиков")
    merged = {**profiles_from_manual(sheets, files, known_suppliers), **(profiles or {})}
    layouts = supplier_layouts_from(sheets, files)
    loaded = []
    total = len(sheets)
    for n, sheet in enumerate(sheets, 1):
        boot_step("Разбираю прайс", f"{n} из {total} · {sheet.file.name}")
        loaded.append(load_one(
            sheet, file_settings(files, sheet), merged, known_suppliers, competitor_brands, layouts,
        ))
    return loaded


def _column_checks(pf: PriceFile) -> dict:
    if not pf.header:
        return {}
    start = pf.header.row + pf.header.depth
    out = {}
    for fld in ("price", "name", "article", "brand", "qty"):
        idx = pf.cols.get(fld)
        if idx is None:
            continue
        title = pf.header.raw[idx] if idx < len(pf.header.raw) else ""
        out[fld] = {
            "letter": pf.header.letter(idx),
            "title": title,
            "numeric": round(_numeric_share(pf.sheet, start, idx), 3),
        }
    return out


def _doubt(checks: dict) -> str:
    notes = []
    price = checks.get("price")
    name = checks.get("name")
    if price and price["numeric"] < 0.9:
        notes.append(
            f"в столбце цены {price['letter']} «{price['title']}» чисел {price['numeric']:.0%} — проверьте, что это не наименование"
        )
    if name and name["numeric"] >= 0.5:
        notes.append(
            f"столбец наименования {name['letter']} «{name['title']}» на {name['numeric']:.0%} из чисел"
        )
    return ". ".join(notes)


def make_profile(pf: PriceFile) -> dict:
    return {"columns": pf.column_titles(), "example": pf.sheet.file.name}


def profile_signature(pf: PriceFile) -> str:
    titles = pf.column_titles()
    return " | ".join(f"{k}={norm_text(titles.get(k, ''))}" for k in ("brand", "article", "price"))


def _col_letter(index: int) -> str:
    n = index + 1
    letters = ""
    while n:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _leading_blank(row: list) -> int:
    n = 0
    for value in row:
        if clean(value):
            return n
        n += 1
    return 0


def _row_cells(row: list, limit: int = 40) -> list[str]:
    last = -1
    for i, value in enumerate(row[:limit]):
        if clean(value):
            last = i
    if last < 0:
        return []
    cells = []
    for i in range(last + 1):
        text = clean(row[i]) if i < len(row) else ""
        cells.append((text[:40] + "…") if len(text) > 40 else (text or "·"))
    return cells


def file_shape(pf: PriceFile) -> str:
    """Короткая строка: где шапка, сколько пустых столбцов слева, какие поля узнались."""
    if not pf.header:
        return "шапка не найдена"
    row = pf.sheet.grid[pf.header.row] if pf.header.row < len(pf.sheet.grid) else []
    bits = [f"шапка {pf.header.row + 1}"]
    left = _leading_blank(row)
    if left:
        bits.append(f"слева {left} пустых")
    names = {"article": "артикул", "price": "цена", "brand": "бренд", "qty": "наличие", "name": "наименование"}
    for key, title in names.items():
        if key in pf.cols:
            bits.append(f"{pf.header.letter(pf.cols[key])} {title}")
    return " · ".join(bits)


def _where(pf: PriceFile) -> str:
    name = pf.sheet.file.name
    if pf.folder:
        name = f"{pf.folder}/{name}"
    if pf.shot_date:
        name = f"{pf.shot_date}/{name}"
    if pf.sheet.name:
        name = f"{name} / {pf.sheet.name}"
    return name


def layout_signature(pf: PriceFile) -> str:
    """Одинаковые прайсы (та же шапка и те же столбцы) схлопываются в одну форму."""
    if pf.header:
        row = pf.sheet.grid[pf.header.row] if pf.header.row < len(pf.sheet.grid) else []
        titles = pf.column_titles()
        parts = [f"row={pf.header.row + 1}", f"left={_leading_blank(row)}"]
        for key in ("brand", "article", "price", "qty", "name"):
            if key in pf.cols:
                parts.append(f"{key}={pf.header.letter(pf.cols[key])}:{norm_text(titles.get(key, ''))}")
        return "|".join(parts)
    for i, row in enumerate(pf.sheet.grid[:SCAN_ROWS]):
        cells = [norm_text(v) for v in row[:40] if clean(v)]
        if len(cells) >= 2:
            return "raw:" + "|".join(cells)
    return "empty"


def _column_lines(pf: PriceFile) -> list[str]:
    if not pf.header:
        return ["Столбцы: шапка не найдена"]
    names = {i: title for title, i in (
        ("бренд", pf.cols.get("brand")), ("артикул", pf.cols.get("article")), ("цена", pf.cols.get("price")),
        ("наличие", pf.cols.get("qty")), ("наименование", pf.cols.get("name")),
    ) if i is not None}
    row = pf.sheet.grid[pf.header.row] if pf.header.row < len(pf.sheet.grid) else []
    last = max(list(pf.cols.values()) + [i for i, v in enumerate(row[:40]) if clean(v)], default=-1)
    lines = ["Столбцы:"]
    for i in range(last + 1):
        title = clean(row[i]) if i < len(row) else ""
        role = names.get(i, "")
        sample = ""
        if pf.header:
            values = [clean(r[i]) for r in pf.header.data[:8] if i < len(r) and clean(r[i])]
            if values:
                sample = " | пример: " + ", ".join(values[:3])
        arrow = f" -> {role}" if role else ""
        share = ""
        if role and pf.header:
            share = f" | чисел {_numeric_share(pf.sheet, pf.header.row + pf.header.depth, i):.0%}"
        lines.append(f"  {_col_letter(i)} {title or '·'}{arrow}{share}{sample}")
    return lines


def _grid_lines(pf: PriceFile) -> list[str]:
    grid = pf.sheet.grid
    if pf.header:
        start = max(0, pf.header.row - 2)
        end = min(len(grid), pf.header.row + 4)
    else:
        start, end = 0, min(len(grid), 8)
    lines = []
    for i in range(start, end):
        cells = _row_cells(grid[i] if i < len(grid) else [])
        if not cells and not lines:
            continue
        lines.append(f"  {i + 1:>3} | " + " | ".join(cells or ["·"]))
    return lines or ["  (лист пустой)"]


LAYOUT_NAME = "Устройство прайсов.txt"


def layout_text(files: list[PriceFile], build: str = "") -> str:
    """Текст, который можно прислать целиком: формы прайсов, шапка, пустые столбцы, примеры строк."""
    groups: dict[str, list[PriceFile]] = {}
    order: list[str] = []
    for pf in files:
        sig = layout_signature(pf)
        if sig not in groups:
            order.append(sig)
            groups[sig] = []
        groups[sig].append(pf)
    recognized = sum(1 for pf in files if pf.header and pf.offers)
    lines = [
        "Устройство прайсов",
        f"Сборка {build}" if build else "Сборка",
        "Этот текст можно прислать целиком. По нему видно строку шапки, пустые столбцы слева и какой столбец за что принят.",
        "Одинаковые файлы собраны в одну форму: сетка показана на примере, остальные перечислены списком.",
        "",
        f"Файлов: {len(files)}. Распознано: {recognized}. Форм: {len(order)}.",
        "",
    ]
    if not files:
        lines.append("Прайсов нет. Положите файлы в Данные/Прайсы/Поставщик.")
        return "\n".join(lines).rstrip() + "\n"
    for n, sig in enumerate(order, 1):
        group = groups[sig]
        example = group[0]
        lines.append(f"Форма {n}. Файлов: {len(group)}. Пример: {_where(example)}")
        lines.append(file_shape(example))
        lines.extend(_column_lines(example))
        lines.append("Строки листа (· — пустая ячейка, число слева — номер строки):")
        lines.extend(_grid_lines(example))
        lines.append("Файлы:")
        for pf in group:
            note = f", {pf.problem}" if pf.problem else ""
            lines.append(f"- {_where(pf)} — {len(pf.offers)} цен, {pf.status}{note}")
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def describe(pf: PriceFile, sample: int = 6) -> dict:
    """Для окна программы: столбцы с примерами значений, распознанные поля, первые строки."""
    columns = []
    if pf.header:
        data = pf.header.data[:40]
        for i, title in enumerate(pf.header.raw):
            values = [clean(r[i]) for r in data if i < len(r) and clean(r[i])]
            if title or values:
                columns.append({"index": i, "letter": pf.header.letter(i), "title": title, "sample": values[:3]})
    return {
        "id": pf.id,
        "file": pf.sheet.file.name,
        "sheet": pf.sheet.name,
        "supplier": pf.supplier,
        "city": pf.city,
        "status": pf.status,
        "enabled": pf.enabled,
        "header_row": pf.header.row + 1 if pf.header else None,
        "cols": pf.cols,
        "fixed_brand": pf.fixed_brand,
        "columns": columns,
        "offers": len(pf.offers),
        "brands": len({brand_key(o.brand) for o in pf.offers}),
        "problem": pf.problem,
        "date": pf.shot_date,
        "folder": pf.folder,
        "supplier_from": pf.supplier_from,
        "name_supplier": pf.name_supplier,
        "mismatch": pf.mismatch,
        "fields": pf.column_titles(),
        "shape": file_shape(pf),
        "checks": pf.checks,
        "doubt": pf.doubt,
        "why": pf.why,
        "can_remember": can_remember(pf),
        "rows": len(pf.sheet.grid),
        "head_rows": [
            {"row": i + 1, "text": " · ".join(clean(v) for v in row if clean(v))[:160]}
            for i, row in enumerate(pf.sheet.grid[:SCAN_ROWS])
            if any(clean(v) for v in row)
        ][:30],
        "preview": [
            {"brand": o.brand, "article": o.article, "price": o.price, "qty": o.qty} for o in pf.offers[:sample]
        ],
    }
