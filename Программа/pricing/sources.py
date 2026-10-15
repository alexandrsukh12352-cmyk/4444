"""Распознавание и чтение выгрузок: ПСТ, Продажи, СБС, товар в пути, АЦПКД, Кроссы, Скидки, ГЗ.

Каждый документ ищется по содержимому (названиям столбцов), а не по имени файла.
Столбцы задаются списками регулярных выражений по заголовку в нижнем регистре;
многоуровневая шапка склеивается через « / » («владивосток / мо / руб.»).
"""

from __future__ import annotations

import re
from collections import defaultdict
from functools import lru_cache
from dataclasses import dataclass, field
from pathlib import Path

from .excel import SUPPORTED_EXT, Header, Sheet, cell, clean, code_key, find_header, norm_key, num
from .bootlog import step as boot_step
from .filecache import read_sheets_cached as read_sheets

# ---------------------------------------------------------------- описания столбцов

PST_COLUMNS = {
    "code": [r"^код товара$", r"^код$"],
    "name": [r"^наименование"],
    "article": [r"^артикул$"],
    "variety": [r"^разновидность"],
    "order_group": [r"^группа заказа$"],
    "category": [r"^категория new$", r"^категория$"],
    "group": [r"^группа new$", r"^группа$"],
    "subgroup": [r"^подгруппа new$", r"^подгруппа$"],
    "brand": [r"^торговая марка$", r"^тм$", r"^бренд$"],
    "discontinued": [r"^вывод из ассортимента"],
    "overstock": [r"оверсток"],
    "abc": [r"^категория abc", r"^abc"],
    "assortment": [r"^тип ассорт"],
    "purchase_currency": [r"^закупочная цена / вал"],
    "purchase_price": [r"^закупочная цена / цена$", r"^закупочная цена$"],
    "mo_currency": [r"(^|/ )мо / вал"],
    "mo_price": [r"(^|/ )мо / цена$"],
    "mo_rub": [r"(^|/ )мо / руб"],
}
PST_REQUIRED = ("code", "article", "subgroup", "brand")

SALES_COLUMNS = {
    "code": [r"^код$", r"^код товара$"],
    "article": [r"^артикул$"],
    "brand": [r"^торговая марка$", r"^бренд$"],
    "amount": [r"сумма продажи"],
    "qty": [r"количество", r"^\d*\.?\s*кол-во"],
}
SALES_REQUIRED = ("amount",)

SBS_COLUMNS = {
    "code": [r"^код$", r"^код товара$"],
    "article": [r"^артикул$"],
    "brand": [r"^торговая марка$", r"^бренд$"],
    "qty": [r"кол-во.*ост", r"^кол-во"],
    "cost": [r"себестоимость полная", r"себестоимость"],
}
SBS_REQUIRED = ("qty", "cost")

# «Товар в пути»: первая колонка — артикул, дальше по городам «Заказано, шт.» и суммы, в конце «Итого».
TRANSIT_COLUMNS = {
    "article": [r"товар\s*\.?\s*артикул", r"^артикул$"],
    "qty": [r"заказано"],
}
TRANSIT_REQUIRED = ("article", "qty")

ACP_COLUMNS = {
    "category": [r"^категория$"],
    "group": [r"^группа$"],
    "subgroup": [r"^подгруппа$"],
    "name": [r"^наименование$"],
    "article": [r"^артикул$"],
    "code": [r"^код$"],
    "abc": [r"^abc$"],
    "assortment": [r"^тип ассорт"],
    "in_matrix": [r"^включен в матрицу"],
    "brand": [r"^бренд$"],
    "promo": [r"^акция$"],
    "mo_price": [r"^цена мо$"],
    "cat_a_price": [r"^цена со скидкой кат"],
    "avg_cost": [r"^средняя себест"],
    "stock": [r"^остаток$"],
    "sold_qty": [r"^продано штук"],
    "avg_sale_price": [r"^ср\. цена продажи"],
}
ACP_REQUIRED = ("code", "mo_price", "cat_a_price")

CROSS_COLUMNS = {
    "article": [r"^артикул$"],
    "brand": [r"^бренд$", r"^торговая марка$"],
    "cross_article": [r"^артикул кросс$"],
    "cross_brand": [r"^бренд кросс$"],
    "use_for_prices": [r"использовать для сравнения цен"],
    "article_search": [r"^артикул поиск$"],
    "cross_article_search": [r"^артикул кросс поиск$"],
}
CROSS_REQUIRED = ("article", "brand", "cross_article", "cross_brand")

DISCOUNT_COLUMNS = {
    "subgroup": [r"^подгруппа"],
    "group": [r"^группа$", r"^группа new$"],
    "order_group": [r"^группа заказа$"],
    "brand": [r"^торговая марка$", r"^бренд$"],
    "division": [r"^подразделение"],
    "discount": [r"^общие скидки"],
}
DISCOUNT_REQUIRED = ("subgroup", "order_group", "brand", "discount")

POSITIONING_REQUIRED = [[r"^бренд конкурент"], [r"^подгруппа"], [r"^торговая марка$"]]

ARTICLE_RE = re.compile(r"[^0-9A-ZА-Я]")
LOOKALIKE = str.maketrans("АВЕКМНОРСТХУ", "ABEKMHOPCTXY")
YES = {"ДА", "YES", "1", "TRUE", "ИСТИНА", "+"}


def norm_article(value) -> str:
    """Артикул для поиска: без точек, дефисов, пробелов и прочего.

    «09.A417.11», «09-A417-11» и «09A41711» — один ключ. Хвост после «_»
    («G20011_ПоврежУпак») отбрасывается: это пометка, не часть номера.
    """
    text = clean(value).upper().replace("Ё", "Е").translate(LOOKALIKE)
    head, sep, _tail = text.partition("_")
    if sep and any(ch.isdigit() for ch in head):
        text = head
    return ARTICLE_RE.sub("", text)


@lru_cache(maxsize=1 << 17)
def _norm_text(value) -> str:
    return clean(value).replace("ё", "е").replace("Ё", "Е").upper()


def norm_text(value) -> str:
    """Текст для сравнения: без лишних пробелов, «ё» как «е», в верхнем регистре.

    Бренды, группы и подгруппы повторяются в каждой строке выгрузок — результат запоминается.
    """
    try:
        return _norm_text(value)
    except TypeError:  # нехешируемое значение — считаем без памяти
        return clean(value).replace("ё", "е").replace("Ё", "Е").upper()


# ---------------------------------------------------------------- модели


@dataclass
class Located:
    kind: str
    header: Header
    cols: dict[str, int]

    @property
    def label(self) -> str:
        return self.header.sheet.label

    def get(self, row: list, key: str):
        return cell(row, self.cols.get(key))

    def describe(self) -> list[tuple[str, str]]:
        h = self.header
        return [(k, f"«{h.raw[i]}» ({h.letter(i)})") for k, i in self.cols.items()]


@dataclass
class PstItem:
    code: str
    article: str
    name: str
    brand: str
    category: str
    group: str
    subgroup: str
    order_group: str
    variety: str
    discontinued: bool
    overstock: str
    abc: str
    assortment: str
    purchase_currency: str
    purchase_price: float | None
    mo_currency: str
    mo_price: float | None
    mo_rub: float | None


@dataclass
class Offer:
    supplier: str
    price: float
    qty: float | None


@dataclass
class AcpCross:
    article: str
    offers: list[Offer] = field(default_factory=list)

    @property
    def min_price(self) -> float | None:
        prices = [o.price for o in self.offers if o.price]
        return min(prices) if prices else None


@dataclass
class AcpItem:
    code: str
    article: str
    name: str
    brand: str
    category: str
    group: str
    subgroup: str
    abc: str
    assortment: str
    in_matrix: str
    promo: str
    mo_price: float | None
    cat_a_price: float | None
    avg_cost: float | None
    stock: float | None
    sold_qty: float | None
    avg_sale_price: float | None
    # бренд-конкурент (как в шапке АЦПКД) -> кроссы
    competitors: dict[str, list[AcpCross]] = field(default_factory=dict)


@dataclass
class CrossRow:
    cross_article: str
    cross_key: str
    cross_brand: str
    yes: bool


@dataclass
class Sources:
    found: dict[str, list[Located]]
    skipped: list[str]
    pst: list[PstItem] = field(default_factory=list)
    sales: dict[str, float] = field(default_factory=dict)  # код -> руб
    sales_qty: dict[str, float] = field(default_factory=dict)  # код -> шт
    sales_period: str = ""  # «2025» или «01.03.2025–31.08.2025»
    price_sheets: list[Sheet] = field(default_factory=list)  # прайсы и всё нераспознанное
    sbs: dict[str, tuple[float | None, float | None]] = field(default_factory=dict)  # код -> (остаток, СБС)
    transit: dict[str, float] = field(default_factory=dict)  # нормализованный артикул -> заказано, шт (итого)
    acp: dict[str, AcpItem] = field(default_factory=dict)  # код -> товар
    oem: list = field(default_factory=list)  # строки справочника ОЕМ: смежные марки
    crosses: dict[tuple[str, str], dict[str, list[CrossRow]]] = field(default_factory=dict)
    discounts: dict[tuple[str, str, str], float] = field(default_factory=dict)
    gz: dict[str, float] = field(default_factory=dict)  # группа заказа (верхний регистр) -> коэффициент
    warnings: list[str] = field(default_factory=list)
    dormant_prices: list[Path] = field(default_factory=list)  # прайсы прошлых дат, ещё не прочитаны
    price_order: dict[str, int] = field(default_factory=dict)  # порядок файлов при чтении папки

    def located(self, kind: str) -> Located | None:
        items = self.found.get(kind) or []
        return items[0] if items else None


# ---------------------------------------------------------------- распознавание

KINDS = [
    # порядок важен: более «узнаваемые» документы проверяются первыми
    ("acp", "АЦПКД", ACP_COLUMNS, ACP_REQUIRED),
    ("crosses", "Кроссы", CROSS_COLUMNS, CROSS_REQUIRED),
    ("discounts", "Скидки для клиентов", DISCOUNT_COLUMNS, DISCOUNT_REQUIRED),
    ("pst", "ПСТ", PST_COLUMNS, PST_REQUIRED),
    ("sbs", "СБС", SBS_COLUMNS, SBS_REQUIRED),
    ("sales", "Продажи", SALES_COLUMNS, SALES_REQUIRED),
]
KIND_TITLES = {k: t for k, t, _, _ in KINDS} | {
    "positioning": "Позиционирование", "gz": "ГЗ", "transit": "Товар в пути", "oem": "ОЕМ",
}
RESULT_MARKERS = {"Сводка", "Для заноса", "По артикулам", "Отчет", "Расч"}
DATA_DIRS = {"1c": "Выгрузки 1С", "crosses": "Кроссы", "prices": "Прайсы"}
SCAN_PRICE_ROWS = 60


def data_paths(root: Path) -> dict[str, Path]:
    """Папки, которые видит человек: общие кроссы и прайсы, и отдельная папка на группу бренда."""
    data = root / "Данные"
    return {
        "program": root / "Программа",
        "data": data,
        "groups": data / "Группы",
        "crosses": data / "Кроссы",
        "prices": data / "Прайсы",
        "output": root / "Результат",
        "legacy": root / "Входящие",
    }


def ensure_data_dirs(root: Path) -> None:
    for key in ("program", "groups", "crosses", "prices", "output"):
        data_paths(root)[key].mkdir(parents=True, exist_ok=True)
    output = data_paths(root)["output"]
    for name in ("Целевые группы", "Мелкие группы", "Нет цен"):
        (output / name).mkdir(parents=True, exist_ok=True)
    special = output / "Распродажа и оверсток"
    for kind in ("Распродажа", "Оверсток"):
        for name in ("Целевые группы", "Мелкие группы", "Нет цен"):
            (special / kind / name).mkdir(parents=True, exist_ok=True)


def locate(sheet: Sheet, columns: dict[str, list[str]], required) -> Located | None:
    header = find_header(sheet, [columns[k] for k in required])
    if not header:
        return None
    cols = {}
    for key, patterns in columns.items():
        idx = header.find(patterns)
        if idx is not None and idx not in cols.values():
            cols[key] = idx
    return Located("", header, cols)


def _transit_marker(loc: Located) -> bool:
    """Отличает «товар в пути» от прайса: есть итог или сумма в валюте поставщика."""
    blob = " | ".join(loc.header.keys)
    return "итого" in blob or "сумма в вал" in blob or "тек. курс" in blob or "тек курс" in blob


def classify_transit(sheet: Sheet) -> Located | None:
    loc = locate(sheet, TRANSIT_COLUMNS, TRANSIT_REQUIRED)
    if not loc or not _transit_marker(loc):
        return None
    loc.kind = "transit"
    return loc


def looks_like_gz(sheet: Sheet) -> bool:
    """ГЗ — две колонки без шапки (или с текстовой шапкой): «Группа заказа», коэффициент ~1.0–2.5."""
    rows = [r for r in sheet.grid if any(clean(v) for v in r)]
    if len(rows) < 3:
        return False
    good = 0
    for r in rows:
        filled = [v for v in r if clean(v)]
        coef = num(filled[1]) if len(filled) == 2 else None
        if coef is not None and num(filled[0]) is None and 0.3 <= coef <= 10:
            good += 1
    return good >= 3 and good >= 0.8 * (len(rows) - 1)


OEM_COLUMNS = {
    "code": [r"^код товара$", r"^код$"],
    "article": [r"^артикул$"],
    "name": [r"^наименование"],
    "brand": [r"^торговая марка$", r"^бренд$", r"^тм$"],
    "oem": [r"оем", r"oem"],
}
OEM_REQUIRED = ("article", "brand", "oem")


@dataclass
class OemItem:
    brand: str
    article: str
    name: str
    code: str
    oems: list[str]


def classify_oem(sheet: Sheet) -> Located | None:
    """Справочник «ОЕМ ТМ+ТМ»: артикул, марка и номера ОЕМ. Это не ПСТ и не прайс."""
    loc = locate(sheet, OEM_COLUMNS, OEM_REQUIRED)
    if not loc:
        return None
    # Файл «ОЕМ …» читаем, даже если в нём есть столбцы как у ПСТ.
    name = norm_key(getattr(loc.header.sheet.file, "name", ""))
    looks_pst = loc.header.find(PST_COLUMNS["subgroup"]) is not None and loc.header.find(PST_COLUMNS["purchase_price"])
    if looks_pst and "oem" not in name and "оем" not in name:
        return None
    loc.kind = "oem"
    loc.oem_cols = [i for i, key in enumerate(loc.header.keys) if key and ("оем" in key or "oem" in key)]
    return loc


def classify(sheet: Sheet) -> Located | None:
    if header := find_header(sheet, POSITIONING_REQUIRED):
        return Located("positioning", header, {})
    for kind, _, columns, required in KINDS:
        loc = locate(sheet, columns, required)
        if loc:
            loc.kind = kind
            return loc
    if loc := classify_transit(sheet):
        return loc
    if loc := classify_oem(sheet):
        return loc
    if looks_like_gz(sheet):
        return Located("gz", Header(sheet, 0, 0, [], []), {})
    return None


def project_dirs(root: Path) -> dict[str, Path] | None:
    """Папки проекта («Выгрузки 1С», «Кроссы», «Прайсы»), если хотя бы одна есть; иначе None."""
    dirs = {k: root / name for k, name in DATA_DIRS.items()}
    return dirs if any(d.is_dir() for d in dirs.values()) else None


def scan_paths(root: Path) -> tuple[list[Path], list[Path]]:
    """(что сканировать, папки прайсов). Без папок проекта — сама папка, как раньше.

    Папка «Входящие» рядом с новыми папками тоже читается: так не пропадают выгрузки,
    которые уже лежали по-старому.
    """
    dirs = project_dirs(root)
    if not dirs:
        return [root], []
    paths = [d for d in dirs.values() if d.is_dir()]
    legacy = root / "Входящие"
    if legacy.is_dir():
        paths.append(legacy)
    return paths, [dirs["prices"]] if dirs["prices"].is_dir() else []


@dataclass
class Pack:
    """Одна папка группы бренда: её ПСТ, СБС, продажи и АЦПКД. Кроссы и прайсы — общие."""
    name: str
    folder: Path
    sources: Sources


COMPANY_KINDS = ("pst", "sbs", "sales")


@dataclass
class AcpSlice:
    """Один файл АЦПКД. Во «Входящие» их может быть несколько, по категории."""
    label: str
    mtime: float
    items: dict
    located: object | None = None


@dataclass
class Workspace:
    packs: list[Pack]
    shared: Sources
    company: Sources = field(default_factory=lambda: Sources({}, []))
    acp_slices: list[AcpSlice] = field(default_factory=list)
    oem_rows: list = field(default_factory=list)
    oem_locs: list = field(default_factory=list)


def _has_docs(src: Sources) -> bool:
    return bool(src.pst or src.acp or src.sbs or src.sales)


def scan_workspace(root: Path) -> Workspace:
    """Кроссы и прайсы читаются один раз на всех. Каждая подпапка «Данные/Группы» — свои выгрузки.

    Пока в группах пусто, берётся «Входящие». ПСТ, СБС и Продажи из «Входящие» общие для всех групп
    и важнее таких же файлов в папке группы.
    """
    paths = data_paths(root)
    ensure_data_dirs(root)
    price_dirs = [paths["prices"]] if paths["prices"].is_dir() else []
    # Все новые файлы рабочей папки — сразу и параллельно; дальше они берутся из кэша.
    everything = collect_files([p for p in (paths["crosses"], paths["prices"], paths["groups"], paths["legacy"]) if p.is_dir()], True)
    resolved = [d.resolve() for d in price_dirs]
    later = _older_dates(everything, resolved)
    _prefetch([f for f in everything if f not in later], resolved)
    shared_paths = [p for p in (paths["crosses"], paths["prices"]) if p.is_dir()]
    shared = scan(shared_paths, price_dirs, all_dates=False) if shared_paths else Sources({}, [])
    packs = []
    if paths["groups"].is_dir():
        for folder in sorted(p for p in paths["groups"].iterdir() if p.is_dir() and not p.name.startswith(".")):
            packs.append(Pack(folder.name, folder, scan([folder])))
    if not any(_has_docs(p.sources) for p in packs):
        for name, folder in (("Входящие", paths["legacy"]), ("Выгрузки 1С", root / "Выгрузки 1С")):
            if folder.is_dir() and collect_files([folder], name == "Выгрузки 1С"):
                packs.append(Pack(name, folder, scan([folder])))
                if _has_docs(packs[-1].sources) or name == "Входящие":
                    break
    company = _company_base(paths["legacy"])
    _share_transit(shared, packs, paths)
    # Корень тоже: файл ОЕМ кладут и во «Входящие», и в «Выгрузки 1С», и рядом с программой.
    oem_rows, oem_locs = read_oem_folders([root])
    return Workspace(packs, shared, company, acp_slices_in(paths["legacy"]), oem_rows, oem_locs)


def _share_transit(shared: Sources, packs: list[Pack], paths: dict[str, Path]) -> None:
    """Один файл «товар в пути» на всю компанию виден всем группам.

    Свой файл в папке группы важнее. Если файл один (в любой группе, во «Входящие»
    или прямо в «Данные») — он подставляется тем группам, где своего нет.
    """
    seen = {p.folder.resolve() for p in packs}
    for folder in (paths["legacy"], paths["data"]):
        if folder.is_dir() and folder.resolve() not in seen and not shared.transit:
            _copy_transit(shared, scan([folder]))
    donors = [src for src in [shared, *(p.sources for p in packs)] if src.transit]
    unique: list[Sources] = []
    for src in donors:
        if not any(src.transit == other.transit for other in unique):
            unique.append(src)
    if len(unique) == 1 and not shared.transit:
        _copy_transit(shared, unique[0])


def _copy_transit(dst: Sources, src: Sources) -> None:
    if dst.transit or not src.transit:
        return
    dst.transit = src.transit
    if src.found.get("transit"):
        dst.found = dict(dst.found)
        dst.found["transit"] = src.found["transit"]


def acp_covers(items: dict, brand: str, group: str) -> bool:
    """В этом АЦПКД есть хотя бы один товар данной торговой марки и группы."""
    if not items or not (brand or group):
        return False
    b, g = norm_text(brand), norm_text(group)
    return any(norm_text(it.brand) == b and norm_text(it.group) == g for it in items.values())


def acp_slices_in(folder: Path) -> list[AcpSlice]:
    """Все файлы АЦПКД в папке, каждый целиком. Самый новый не вытесняет остальные."""
    if not folder.is_dir():
        return []
    slices = []
    for f in collect_files([folder]):
        try:
            sheets = read_sheets(f)
        except Exception:  # noqa: BLE001
            continue
        if _is_finished_report(sheets):
            continue
        mtime = f.stat().st_mtime
        for sheet in sheets:
            loc = classify(sheet)
            if not loc or loc.kind != "acp":
                continue
            slices.append(AcpSlice(loc.label, mtime, read_acp(loc), loc))
    return slices


def _company_base(folder: Path) -> Sources:
    """ПСТ, СБС и Продажи из «Входящие» — одни на весь ассортимент. Остальное отсюда не берётся."""
    if not folder.is_dir() or not collect_files([folder]):
        return Sources({}, [])
    src = scan([folder])
    found = {kind: src.found[kind] for kind in COMPANY_KINDS if src.found.get(kind)}
    out = Sources(found, [])
    if "pst" in found:
        out.pst = src.pst
    if "sbs" in found:
        out.sbs = src.sbs
    if "sales" in found:
        out.sales = src.sales
        out.sales_qty = src.sales_qty
        out.sales_period = src.sales_period
    return out


def apply_company(src: Sources, company: Sources) -> Sources:
    """Общий документ из «Входящие» заменяет такой же в папке группы. Нет общего — остаётся групповой."""
    import copy

    out = copy.copy(src)
    out.found = dict(src.found)
    out.warnings = list(src.warnings)
    for kind in COMPANY_KINDS:
        if not company.found.get(kind):
            continue
        local = src.found.get(kind)
        common = company.found[kind][0].label
        if local and local[0].label != common:
            out.warnings.append(
                f"{KIND_TITLES[kind]} в папке группы ({local[0].label}) не используется — взято из «Входящие» ({common})"
            )
        out.found[kind] = company.found[kind]
        if kind == "pst":
            out.pst = company.pst
        elif kind == "sbs":
            out.sbs = company.sbs
        else:
            out.sales = company.sales
            out.sales_qty = company.sales_qty
            out.sales_period = company.sales_period
    return out


def attach_shared(src: Sources, shared: Sources) -> Sources:
    """Копия выгрузок группы плюс общие кроссы, прайсы, ГЗ и скидки, если в папке группы их нет."""
    import copy

    out = copy.copy(src)
    out.found = dict(src.found)
    out.warnings = list(src.warnings)
    out.skipped = list(src.skipped) + [s for s in shared.skipped if s not in src.skipped]
    out.price_sheets = list(shared.price_sheets) + list(src.price_sheets)
    if shared.crosses and not out.crosses:
        out.crosses = shared.crosses
        if shared.found.get("crosses"):
            out.found["crosses"] = shared.found["crosses"]
    if shared.gz and not out.gz:
        out.gz = shared.gz
        if shared.found.get("gz"):
            out.found["gz"] = shared.found["gz"]
    if shared.discounts and not out.discounts:
        out.discounts = shared.discounts
        if shared.found.get("discounts"):
            out.found["discounts"] = shared.found["discounts"]
    if shared.transit and not out.transit:
        _copy_transit(out, shared)
    if shared.found.get("positioning") and "positioning" not in out.found:
        out.found["positioning"] = shared.found["positioning"]
    return out


def collect_files(paths: list[Path], recursive: bool = False) -> list[Path]:
    files: list[Path] = []
    for p in paths:
        if p.is_dir():
            items = p.rglob("*") if recursive else p.iterdir()
            files += sorted(x for x in items if x.is_file())
        elif p.exists():
            files.append(p)
    return [f for f in files if not f.name.startswith(("~$", ".")) and f.suffix.lower() in SUPPORTED_EXT]


def _is_finished_report(sheets: list[Sheet]) -> bool:
    """Готовый отчёт не читаем снова как выгрузку.

    Старые книги узнаём по набору листов. Новые — по листам «Отчет» и «ПСТ»
    или по листу «Данные».
    """
    names = {s.name for s in sheets}
    if (
        {"Отчет", "Расч"} <= names
        or {"Сводка", "Для заноса"} <= names
        or {"Исключено", "Параметры"} <= names
        or {"Файлы", "Сопоставление"} <= names
        or "Дашборд" in names
        or ("Отчет" in names and bool({"ПСТ", "СБС", "Продажи", "Параметры"} & names))
        or {"ПСТ", "Параметры"} <= names
        or ({"ПСТ", "СБС", "Продажи"} <= names and "Отчет" in names)
    ):
        return True
    for s in sheets:
        if s.name != "Данные" or not s.grid or not s.grid[0]:
            continue
        if clean(s.grid[0][0]) in ("Курсы", "Параметры формирования"):
            return True
    return False


def _is_currency_file(path: Path) -> bool:
    """Файл «валюта» в папке прайсов — курсы по датам, не прайс поставщика."""
    stem = path.stem.casefold().replace("ё", "е").strip()
    return stem == "валюта" or stem.startswith("валюта ") or stem.startswith("валюта_")


def _prefetch(files: list[Path], price_dirs: list[Path]) -> None:
    """Новые файлы читаются из Excel параллельно, до разбора по одному."""
    from .filecache import prefetch

    prices = {f for f in files if any(d in f.resolve().parents for d in price_dirs) and not _is_currency_file(f)}
    wanted = [f for f in files if not _is_currency_file(f)]
    prefetch(wanted, prices, lambda name, n, total: boot_step("Читаю новые файлы параллельно", f"{n} из {total} · {name}"))


def _price_sheets_of(f: Path) -> list[Sheet]:
    """Листы прайса без полной сетки в памяти: только начало (PriceSheet)."""
    from .filecache import read_price_heads

    return [s for s in read_price_heads(f) if any(clean(v) for row in s.grid[:SCAN_PRICE_ROWS] for v in row)]


def _older_dates(files: list[Path], price_dirs: list[Path]) -> set[Path]:
    """Прайсы в папках-датах старше последней. Отчёту они не нужны, только динамике."""
    from .prices import _shot_sort, price_date

    dated = {}
    for f in files:
        if any(d in f.resolve().parents for d in price_dirs) and not _is_currency_file(f):
            date = price_date(f)
            if date:
                dated[f] = date
    if len(set(dated.values())) < 2:
        return set()
    latest = max(dated.values(), key=_shot_sort)
    return {f for f, date in dated.items() if date != latest}


def scan(paths: list[Path], price_dirs: list[Path] | None = None, all_dates: bool = True) -> Sources:
    """Читает выгрузки и прайсы.

    all_dates=False — прайсы прошлых дат не читаются: отчёт берёт последнюю дату. Их пути
    лежат в src.dormant_prices, а wake_dormant дочитывает их, когда нужна динамика.
    Так читает окно программы; по умолчанию, как раньше, читается всё.
    """
    found: dict[str, list[Located]] = defaultdict(list)
    skipped: list[str] = []
    price_sheets: list[Sheet] = []
    price_dirs = [d.resolve() for d in price_dirs or []]
    recursive = bool(price_dirs) or any(p.is_dir() and p.name in DATA_DIRS.values() for p in paths)
    files = collect_files(paths, recursive)
    later = set() if all_dates else _older_dates(files, price_dirs)
    _prefetch([f for f in files if f not in later], price_dirs)
    dormant: list[Path] = []
    for f in files:
        if f in later:
            dormant.append(f)
            continue
        boot_step("Читаю файл", f.name)
        if any(d in f.resolve().parents for d in price_dirs):
            if _is_currency_file(f):
                continue
            try:
                price_sheets += _price_sheets_of(f)
            except Exception as e:  # noqa: BLE001
                skipped.append(f"{f.name}: не удалось прочитать ({e})")
            continue
        try:
            sheets = read_sheets(f)
        except Exception as e:  # noqa: BLE001
            skipped.append(f"{f.name}: не удалось прочитать ({e})")
            continue
        if _is_finished_report(sheets):
            skipped.append(f"{f.name}: это готовый отчёт — пропущен")
            continue
        located = [(s, classify(s)) for s in sheets]
        is_positioning_book = any(loc and loc.kind == "positioning" for _, loc in located)
        for s, loc in located:
            if loc:
                found[loc.kind].append(loc)
            elif is_positioning_book:
                continue  # «Поставщики по брендам», «Описание целевых групп» — часть позиционирования
            elif any(clean(v) for row in s.grid[:SCAN_PRICE_ROWS] for v in row):
                price_sheets.append(s)
    src = Sources(dict(found), skipped, price_sheets=price_sheets)
    src.dormant_prices = dormant
    src.price_order = {str(f): i for i, f in enumerate(files)}
    for kind, items in src.found.items():
        if len(items) > 1 and kind != "oem":
            items.sort(key=lambda l: l.header.sheet.file.stat().st_mtime, reverse=True)
            names = ", ".join(l.label for l in items)
            src.warnings.append(f"{KIND_TITLES[kind]}: найдено несколько ({names}) — взят {items[0].label}")
    load_all(src)
    return src


def dormant_with_settings(src: Sources, files: dict) -> list[Sheet]:
    """Листы непрочитанных прайсов прошлых дат, у которых есть своя настройка столбцов.

    По ним программа учит столбцы поставщика и форматы для следующих дат, поэтому их
    читаем и без динамики. Обычно это один-два файла, а не все прошлые даты.
    """
    from .prices import price_date, price_id, supplier_folder

    if not files:
        return []
    keys = list(files)
    out: list[Sheet] = []
    for f in getattr(src, "dormant_prices", None) or []:
        bits = [b for b in (price_date(f), supplier_folder(f), f.name) if b]
        heads = {" / ".join(bits), f.name}
        if not any(k in heads or any(k.startswith(h + " / ") for h in heads) for k in keys):
            continue
        try:
            sheets = _price_sheets_of(f)
        except Exception:  # noqa: BLE001
            continue
        out += [s for s in sheets if files.get(price_id(s)) or files.get(s.label)]
    return out


def wake_dormant(src: Sources) -> bool:
    """Дочитывает прайсы прошлых дат в src.price_sheets. Порядок листов — как при полном чтении.

    True — если что-то дочитали (тогда разобранные прайсы и группы надо собрать заново).
    """
    dormant = list(getattr(src, "dormant_prices", None) or [])
    if not dormant:
        return False
    from .filecache import prefetch

    prefetch(dormant, set(dormant), lambda name, n, total: boot_step("Читаю прайсы прошлых дат", f"{n} из {total} · {name}"))
    sheets = list(src.price_sheets)
    for f in dormant:
        boot_step("Читаю прайс прошлой даты", f.name)
        try:
            sheets += _price_sheets_of(f)
        except Exception as e:  # noqa: BLE001
            src.skipped.append(f"{f.name}: не удалось прочитать ({e})")
    order = getattr(src, "price_order", None) or {}
    big = len(order) + 1
    sheets = [s for _, s in sorted(enumerate(sheets), key=lambda item: (order.get(str(item[1].file), big), item[0]))]
    src.price_sheets = sheets
    src.dormant_prices = []
    return True


# ---------------------------------------------------------------- чтение


def _is_total(row: list, loc: Located) -> bool:
    return any(norm_text(v) in ("ИТОГО", "ВСЕГО", "TOTAL") for v in row[:4])


def read_pst(loc: Located) -> list[PstItem]:
    items = []
    for row in loc.header.data:
        code = code_key(loc.get(row, "code"))
        article = clean(loc.get(row, "article"))
        if not code or not article or _is_total(row, loc):
            continue
        items.append(
            PstItem(
                code=code,
                article=article,
                name=clean(loc.get(row, "name")),
                brand=clean(loc.get(row, "brand")),
                category=clean(loc.get(row, "category")),
                group=clean(loc.get(row, "group")),
                subgroup=clean(loc.get(row, "subgroup")),
                order_group=clean(loc.get(row, "order_group")),
                variety=clean(loc.get(row, "variety")),
                discontinued=(num(loc.get(row, "discontinued")) or 0) != 0,
                overstock=clean(loc.get(row, "overstock")),
                abc=clean(loc.get(row, "abc")).upper(),
                assortment=clean(loc.get(row, "assortment")).upper(),
                purchase_currency=clean(loc.get(row, "purchase_currency")),
                purchase_price=num(loc.get(row, "purchase_price")),
                mo_currency=clean(loc.get(row, "mo_currency")),
                mo_price=num(loc.get(row, "mo_price")),
                mo_rub=num(loc.get(row, "mo_rub")),
            )
        )
    return items


def _key(loc: Located, row: list) -> str:
    code = code_key(loc.get(row, "code"))
    if code:
        return code
    art = norm_article(loc.get(row, "article"))
    return f"art:{art}" if art else ""


def read_sales(loc: Located) -> tuple[dict[str, float], dict[str, float]]:
    """(код -> сумма, руб; код -> количество, шт). Количества может не быть в старых выгрузках."""
    amount: dict[str, float] = defaultdict(float)
    qty: dict[str, float] = defaultdict(float)
    for row in loc.header.data:
        key = _key(loc, row)
        if not key or _is_total(row, loc):
            continue
        amount[key] += num(loc.get(row, "amount")) or 0.0
        if "qty" in loc.cols:
            qty[key] += num(loc.get(row, "qty")) or 0.0
    return dict(amount), dict(qty)


def sales_period(loc: Located) -> str:
    """Период из шапки 1С над таблицей: «…: 2025 г.» или «Начало периода: 01.01.2025»."""
    text = " ".join(clean(v) for row in loc.header.sheet.grid[: loc.header.row] for v in row if clean(v))
    start = re.search(r"начало периода:\s*(\d{2}\.\d{2}\.(\d{4}))", text, re.IGNORECASE)
    end = re.search(r"конец периода:\s*(\d{2}\.\d{2}\.(\d{4}))", text, re.IGNORECASE)
    if start and end:
        if start.group(1)[:5] == "01.01" and end.group(1)[:5] == "31.12" and start.group(2) == end.group(2):
            return start.group(2)
        return f"{start.group(1)}–{end.group(1)}"
    year = re.search(r"\b(20\d{2})\s*г", text)
    return year.group(1) if year else ""


def _transit_qty_columns(loc: Located) -> list[int]:
    """Столбцы «Заказано, шт.». Если есть «Итого» — только он, иначе сумма по городам."""
    ordered = [i for i, key in enumerate(loc.header.keys) if "заказано" in key]
    total = [i for i in ordered if "итого" in loc.header.keys[i]]
    return total or ordered


def read_transit(loc: Located) -> dict[str, float]:
    """Артикул без знаков -> сколько штук в пути (итог по всем городам)."""
    cols = _transit_qty_columns(loc)
    result: dict[str, float] = defaultdict(float)
    for row in loc.header.data:
        art = norm_article(loc.get(row, "article"))
        if not art or _is_total(row, loc):
            continue
        parts = [num(cell(row, i)) for i in cols]
        if not any(p is not None for p in parts):
            continue
        result[art] += sum(p or 0.0 for p in parts)
    return dict(result)


def read_sbs(loc: Located) -> dict[str, tuple[float | None, float | None]]:
    qty: dict[str, float] = defaultdict(float)
    cost: dict[str, float] = defaultdict(float)
    for row in loc.header.data:
        key = _key(loc, row)
        if not key or _is_total(row, loc):
            continue
        qty[key] += num(loc.get(row, "qty")) or 0.0
        cost[key] += num(loc.get(row, "cost")) or 0.0
    return {k: (qty[k], cost[k] / qty[k] if qty[k] > 0 and cost[k] > 0 else None) for k in qty}


def read_acp(loc: Located) -> dict[str, AcpItem]:
    h = loc.header
    # Блоки конкурентов: «BREMBO / Артикул», «BREMBO / Склад / Цена», «BREMBO / Склад / Наличие».
    blocks: dict[str, dict] = {}
    for idx, raw in enumerate(h.raw):
        parts = raw.split(" / ")
        if len(parts) == 2 and parts[1].lower() == "артикул":
            blocks[parts[0]] = {"article": idx, "offers": {}}
        elif len(parts) == 3 and parts[0] in blocks:
            supplier = re.sub(r"\s+", " ", parts[1]).strip()
            kind = parts[2].lower()
            slot = blocks[parts[0]]["offers"].setdefault(supplier, {})
            slot["price" if kind.startswith("цен") else "qty"] = idx

    items: dict[str, AcpItem] = {}
    for row in h.data:
        code = code_key(loc.get(row, "code"))
        if not code:
            continue
        item = items.get(code)
        if item is None:
            item = AcpItem(
                code=code,
                article=clean(loc.get(row, "article")),
                name=clean(loc.get(row, "name")),
                brand=clean(loc.get(row, "brand")),
                category=clean(loc.get(row, "category")),
                group=clean(loc.get(row, "group")),
                subgroup=clean(loc.get(row, "subgroup")),
                abc=clean(loc.get(row, "abc")).upper(),
                assortment=clean(loc.get(row, "assortment")).upper(),
                in_matrix=clean(loc.get(row, "in_matrix")),
                promo=clean(loc.get(row, "promo")),
                mo_price=num(loc.get(row, "mo_price")),
                cat_a_price=num(loc.get(row, "cat_a_price")),
                avg_cost=num(loc.get(row, "avg_cost")),
                stock=num(loc.get(row, "stock")),
                sold_qty=num(loc.get(row, "sold_qty")),
                avg_sale_price=num(loc.get(row, "avg_sale_price")),
            )
            items[code] = item
        for brand, spec in blocks.items():
            article = clean(cell(row, spec["article"]))
            if not article:
                continue
            crosses = item.competitors.setdefault(brand, [])
            cross = next((c for c in crosses if norm_article(c.article) == norm_article(article)), None)
            if cross is None:
                cross = AcpCross(article)
                crosses.append(cross)
            for supplier, cols in spec["offers"].items():
                price = num(cell(row, cols.get("price")))
                if price and not any(o.supplier == supplier for o in cross.offers):
                    cross.offers.append(Offer(supplier, price, num(cell(row, cols.get("qty")))))
    return items


def read_crosses(loc: Located) -> dict[tuple[str, str], dict[str, list[CrossRow]]]:
    index: dict[tuple[str, str], dict[str, list[CrossRow]]] = defaultdict(dict)
    for row in loc.header.data:
        brand = norm_text(loc.get(row, "brand"))
        art = norm_article(loc.get(row, "article_search") or loc.get(row, "article"))
        cross_article = clean(loc.get(row, "cross_article"))
        cross_key = norm_article(loc.get(row, "cross_article_search") or cross_article)
        cross_brand = clean(loc.get(row, "cross_brand"))
        if not (brand and art and cross_key and cross_brand):
            continue
        bucket = index[(brand, art)].setdefault(norm_text(cross_brand), [])
        yes = norm_text(loc.get(row, "use_for_prices")) in YES
        existing = next((c for c in bucket if c.cross_key == cross_key), None)
        if existing:
            existing.yes = existing.yes or yes
        else:
            bucket.append(CrossRow(cross_article, cross_key, cross_brand, yes))
    return dict(index)


def read_discounts(loc: Located) -> dict[tuple[str, str, str], float]:
    result = {}
    for row in loc.header.data:
        value = num(loc.get(row, "discount"))
        if value is None:
            continue
        key = (
            norm_text(loc.get(row, "subgroup")),
            norm_text(loc.get(row, "order_group")),
            norm_text(loc.get(row, "brand")),
        )
        result.setdefault(key, value)
    return result


def read_gz(loc: Located) -> dict[str, float]:
    result = {}
    for row in loc.header.sheet.grid:
        filled = [v for v in row if clean(v)]
        if len(filled) == 2 and (coef := num(filled[1])) is not None and num(filled[0]) is None:
            result.setdefault(norm_text(filled[0]), coef)
    return result


def _oem_numbers(text) -> list[str]:
    """Номера из ячейки «ОЕМ-номер». Составные артикулы через «+» не режем: это комплект, не список ОЕМ."""
    out = []
    for part in re.split(r"[,;/|\n]+", clean(text)):
        key = norm_article(part)
        if len(key) >= 5:
            out.append(key)
    return out


def read_oem(loc: Located) -> list[OemItem]:
    cols = list(getattr(loc, "oem_cols", None) or [])
    if "oem" in loc.cols and loc.cols["oem"] not in cols:
        cols.append(loc.cols["oem"])
    items = []
    for row in loc.header.data:
        article = clean(loc.get(row, "article"))
        brand = clean(loc.get(row, "brand"))
        if not article or not brand or _is_total(row, loc):
            continue
        oems: list[str] = []
        for idx in cols:
            oems.extend(_oem_numbers(cell(row, idx)))
        if not oems:
            continue
        items.append(OemItem(
            brand=brand,
            article=article,
            name=clean(loc.get(row, "name")),
            code=code_key(loc.get(row, "code")),
            oems=list(dict.fromkeys(oems)),
        ))
    return items


# Эти папки не справочники ОЕМ. Готовый отчёт, кэш и прайсы туда не смотрим:
# иначе ради ОЕМ перечитывались бы прайсы всех дат.
_OEM_SKIP = {"результат", "кэш", "программа", ".git", "примеры", "__pycache__", "прайсы"}


def read_oem_folders(folders: list[Path]) -> tuple[list[OemItem], list[Located]]:
    """Все справочники ОЕМ под корнем программы, кроме результата, кэша и самой программы."""
    rows: list[OemItem] = []
    locs: list[Located] = []
    seen: set[str] = set()
    for folder in folders:
        if not folder or not folder.is_dir():
            continue
        key = str(folder.resolve())
        if key in seen:
            continue
        seen.add(key)
        for f in collect_files([folder], recursive=True):
            if any(part.lower() in _OEM_SKIP for part in f.parts):
                continue
            try:
                sheets = read_sheets(f)
            except Exception:  # noqa: BLE001
                continue
            name = norm_key(f.name)
            if _is_finished_report(sheets) and "oem" not in name and "оем" not in name:
                continue
            for sheet in sheets:
                loc = classify_oem(sheet)
                if not loc:
                    continue
                locs.append(loc)
                rows.extend(read_oem(loc))
    return rows, locs


def load_all(src: Sources) -> None:
    if loc := src.located("gz"):
        src.gz = read_gz(loc)
    if loc := src.located("pst"):
        src.pst = read_pst(loc)
    if loc := src.located("sales"):
        src.sales, src.sales_qty = read_sales(loc)
        src.sales_period = sales_period(loc)
    if loc := src.located("sbs"):
        src.sbs = read_sbs(loc)
    if loc := src.located("transit"):
        src.transit = read_transit(loc)
    if loc := src.located("acp"):
        src.acp = read_acp(loc)
    if loc := src.located("crosses"):
        src.crosses = read_crosses(loc)
    if loc := src.located("discounts"):
        src.discounts = read_discounts(loc)
    src.oem = []
    for loc in src.found.get("oem") or []:
        src.oem.extend(read_oem(loc))
    src.oem_rows = list(src.oem)


def discount_for(src: Sources, subgroup: str, order_group: str, brand: str) -> float | None:
    return src.discounts.get((norm_text(subgroup), norm_text(order_group), norm_text(brand)))
