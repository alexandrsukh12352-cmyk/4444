"""Итоговый отчёт ценового анализа: сборка строк, фильтры, подбор кроссов, аналитика, Excel.

Основа отчёта — ПСТ (+ СБС, Продажи, скидки, ГЗ) или АЦПКД, дополненная столбцами из ПСТ.
Для каждого поля отдельно выбирается, откуда брать значение: наши выгрузки или АЦПКД.
Цены конкурентов — минимальные из АЦПКД и/или прайсов поставщиков.
"""

from __future__ import annotations

import copy
import re
import sys
import threading
from collections import Counter, OrderedDict, defaultdict
from dataclasses import dataclass, field
from pathlib import Path

from .cars import CarInfo, CarParser
from .prices import PriceFile, brand_key, supplier_key
from .recommend import Rec, RecOptions, recommend, target_for, weight_classes
from .sisters import apply_sisters, sister_of
from .settings import Settings
from .sources import AcpItem, PstItem, Sources, norm_article, norm_text

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import build_crosses as bc  # noqa: E402

# ---------------------------------------------------------------- справочники режимов

FIELDS = {
    "mo": ("МО, руб.", "ПСТ («Руб.»)", "АЦПКД («Цена МО»)"),
    "cat_a": ("Кат. А", "МО × (1 − скидка)", "АЦПКД («Цена со скидкой кат. А»)"),
    "sbs": ("СБС", "СБС (себестоимость / кол-во)", "АЦПКД («Средняя себест-ть»)"),
    "stock": ("Остаток", "СБС («Кол-во, кон. ост.»)", "АЦПКД («Остаток»)"),
    "abc": ("Категория ABC", "ПСТ («Категория ABCD»)", "АЦПКД («ABC»)"),
    "sales_qty": ("Продажи, шт", "Продажи («Количество»)", "АЦПКД («Продано штук»)"),
    "sales": ("Продажи, руб.", "Продажи («Сумма продажи»)", "АЦПКД (шт × ср. цена)"),
    "assortment": ("Тип ассортимента", "ПСТ", "АЦПКД («Тип ассорт-та»)"),
}
SOURCE_CHOICES = {
    "own": "Наши выгрузки",
    "acp": "АЦПКД",
    "own_acp": "Наши, если нет — АЦПКД",
    "acp_own": "АЦПКД, если нет — наши",
}

CROSS_MODES = {
    "yes": "Да",
    "no": "Нет",
    "yes_no": "Да и Нет",
    "yes_fill": "Да + добор Нет",
    "off": "Не брать",
}
CROSS_MODE_HINTS = {
    "yes": "только кроссы «Использовать для сравнения цен» = Да",
    "no": "только кроссы «Нет»",
    "yes_no": "все кроссы «Да» и «Нет»",
    "yes_fill": "«Да»; если у товара нет «Да» с ценой — добираем «Нет»",
    "off": "конкурента в этой части отчёта не показываем",
}
CROSS_SPLITS = {
    "none": "Без разбивки",
    "abc": "По категориям ABC",
    "subgroup": "По подгруппам",
    "subgroup_abc": "Подгруппы и ABC",
}
PRICE_SOURCES = {"acp": "АЦПКД", "prices": "Прайсы", "both": "АЦПКД и прайсы"}
REPORT_KINDS = {"full": "Полный", "quick": "Быстрый"}
SUBGROUP_MODES = {"sheets": "Лист на подгруппу", "single": "Все на одном листе"}

FILTERS = {
    # ключ: (название, варианты, по умолчанию)
    "overstock": ("Оверсток", {"include": "Показывать", "exclude": "Скрыть", "only": "Только оверсток"}, "include"),
    "sale": ("Распродажа и вывод", {"include": "Показывать", "exclude": "Скрыть", "only": "Только они"}, "include"),
    "new": ("NEW (тип ассортимента N)", {"include": "Показывать", "exclude": "Скрыть", "only": "Только NEW"}, "include"),
    "stock": ("Остаток", {"include": "Все", "exclude": "Только без остатка", "only": "Только с остатком"}, "include"),
    "stock_transit": ("Нет остатка и нет в пути", {"include": "Показывать", "exclude": "Скрыть"}, "include"),
    "sales": ("Продажи", {"include": "Все", "exclude": "Только без продаж", "only": "Только с продажами"}, "include"),
    "mo": ("Без МО", {"include": "Показывать", "exclude": "Скрыть"}, "include"),
    "cost": ("Без себестоимости", {"include": "Показывать", "exclude": "Скрыть"}, "include"),
    "competitor_price": ("Цена конкурента", {"include": "Все", "only": "Только с ценой", "exclude": "Только без цены"}, "include"),
}
CLEAN_PRESET = {"sale": "exclude", "overstock": "exclude", "new": "exclude", "mo": "exclude", "cost": "exclude"}
ABC_VALUES = ["AA", "A", "B", "C", "D", "NEW", ""]
ABC_ORDER = {v: i for i, v in enumerate(ABC_VALUES)}


# ---------------------------------------------------------------- модель строки


@dataclass
class PriceOffer:
    supplier: str
    price: float
    qty: float | None
    source: str  # «АЦПКД» или файл прайса


@dataclass
class Pick:
    article: str
    yes: bool | None  # None — кросс взят из АЦПКД
    offers: list[PriceOffer]


@dataclass
class Row:
    code: str
    name: str
    article: str
    brand: str
    order_group: str
    category: str
    group: str
    subgroup: str
    variety: str
    purchase_currency: str
    purchase_price: float | None
    purchase_rub: float | None
    mo_currency: str
    mo: float | None
    sales_qty: float | None
    sales: float | None
    stock: float | None
    abc: str
    sbs: float | None
    sbs_plan: float | None
    cat_a: float | None
    discount: float | None
    gz: float | None
    overstock: str
    discontinued: bool
    assortment: str
    pst: PstItem | None
    acp: AcpItem | None
    in_transit: float | None = None  # заказано, шт — из «Товар в пути», итог по городам
    competitors: list = field(default_factory=list)  # bc.Competitor из позиционирования
    picks: dict[str, list[Pick]] = field(default_factory=dict)  # ключ конкурента -> кроссы
    prices: dict[str, float | None] = field(default_factory=dict)
    best: dict[str, PriceOffer] = field(default_factory=dict)
    cat_a_calc: bool = False  # Кат. А = МО × (1 − скидка), а не взята из АЦПКД
    car: CarInfo | None = None
    percents: dict[str, float] = field(default_factory=dict)  # % позиционирования по конкуренту
    target: float | None = None
    target_by: str = ""
    weight: str = "tail"
    rec: Rec | None = None
    sister_brand: str = ""
    sister_article: str = ""
    sister_price: float | None = None

    @property
    def any_price(self) -> bool:
        return any(v for v in self.prices.values())

    @property
    def gap(self) -> float | None:
        """Откл от цели как в отчёте: 1 − Кат. А / цель (> 0 — мы дешевле цели)."""
        return 1 - self.cat_a / self.target if self.cat_a and self.target else None

    @property
    def rec_gap(self) -> float | None:
        return 1 - self.rec.price / self.target if self.rec and self.rec.price and self.target else None


@dataclass
class Options:
    brand: str = ""
    group: str = ""
    base: str = "pst"
    sources: dict[str, str] = field(default_factory=dict)
    filters: dict[str, str] = field(default_factory=dict)
    abc: list[str] = field(default_factory=list)  # пусто — все
    cross_split: str = "none"
    cross_modes: dict[str, str] = field(default_factory=dict)  # «КОНКУРЕНТ|сегмент» -> режим
    price_source: str = "both"
    monitoring_only: bool = True
    in_stock_only: bool = False
    subgroup_mode: str = "sheets"
    suggest_threshold: float = 0.3
    report_kind: str = "full"  # full | quick
    rec: dict = field(default_factory=dict)
    green_when: str = "above"  # above — зелёный, когда мы дороже цели; below — когда дешевле
    money_base: str = "sales"  # sales — сумма с продаж; potential — с потенциала (остаток + в пути)
    stock_slice: str = ""  # sale / overstock — отдельный отчёт только по этой выборке
    sheet_filters: list | None = None  # None — три фильтра по умолчанию; список — что вывести над таблицей

    @classmethod
    def from_dict(cls, data: dict) -> "Options":
        known = {k: v for k, v in (data or {}).items() if k in cls.__dataclass_fields__}
        opt = cls(**known)
        if opt.cross_split not in CROSS_SPLITS:
            opt.cross_split = "none"
        if opt.price_source not in PRICE_SOURCES:
            opt.price_source = "both"
        if opt.report_kind not in REPORT_KINDS:
            opt.report_kind = "full"
        if opt.subgroup_mode not in SUBGROUP_MODES:
            opt.subgroup_mode = "sheets"
        if opt.green_when not in ("above", "below"):
            opt.green_when = "above"
        if opt.money_base not in ("sales", "potential"):
            opt.money_base = "sales"
        if opt.stock_slice not in ("sale", "overstock"):
            opt.stock_slice = ""
        opt.filters = _fold_clearance(opt.filters)
        opt = focus_stock(opt)
        if opt.sheet_filters is not None:
            opt.sheet_filters = [k for k in opt.sheet_filters if k in SHEET_FILTERS]
        return opt

    @property
    def recs(self) -> RecOptions:
        return RecOptions.from_dict(self.rec)

    def source(self, key: str) -> str:
        default = "acp_own" if self.base == "acp" else "own_acp"
        value = self.sources.get(key) or default
        return value if value in SOURCE_CHOICES else default

    def filter(self, key: str) -> str:
        if key == "discontinued":
            key = "sale"
        return self.filters.get(key) or FILTERS[key][2]


# ---------------------------------------------------------------- сегменты кроссов


def segment(r: "Row", split: str) -> tuple[str, str, str]:
    """(ключ сегмента, подгруппа, ABC) для разбивки кроссов."""
    sub, abc = r.subgroup, r.abc or ""
    if split == "abc":
        return f"abc={abc}", "", abc
    if split == "subgroup":
        return f"sub={norm_text(sub)}", sub, ""
    if split == "subgroup_abc":
        return f"sub={norm_text(sub)};abc={abc}", sub, abc
    return "*", "", ""


def mode_for(opt: Options, comp_key: str, seg_key: str, suggested: str) -> str:
    mode = opt.cross_modes.get(f"{comp_key}|{seg_key}") or opt.cross_modes.get(f"{comp_key}|*") or suggested
    return mode if mode in CROSS_MODES else suggested


# ---------------------------------------------------------------- позиционирование


def positioning_table(pos) -> list[list]:
    """Строки для редактора: ТМ, группа, подгруппа и до трёх конкурентов с процентом."""
    rows = []
    for r in pos.rows:
        comps = sorted(r.competitors, key=lambda c: (c.order, c.name))[:3]
        line = [r.brand, r.group, r.subgroup]
        for i in range(3):
            if i < len(comps):
                pct = comps[i].percents[0] if comps[i].percents else ""
                line += [comps[i].name, pct]
            else:
                line += ["", ""]
        rows.append(line)
    return rows


def positioning_from_table(rows, monitoring=None):
    """Таблица из настроек → позиционирование. Процент относится к этой подгруппе, не ко всей группе."""
    monitoring = bc.builtin_monitoring() if monitoring is None else monitoring
    parsed = []
    for raw in rows or []:
        cells = list(raw) + [""] * 9
        brand, group, subgroup = (bc.clean(cells[0]), bc.clean(cells[1]), bc.clean(cells[2]))
        if not brand and not subgroup:
            continue
        comps = []
        for order, i in enumerate((3, 5, 7), 1):
            name, pct = bc.clean(cells[i]), bc.clean(cells[i + 1])
            if name:
                comps.append(bc.Competitor(order, name, bc.norm_text(name), [str(pct)] if pct != "" else []))
        parsed.append(bc.PosRow(category="", group=group, subgroup=subgroup, brand=brand, competitors=comps))
    return bc.Positioning(parsed, monitoring, "настройки программы")


_BUILTIN_POSITIONING = None


def load_positioning(src: Sources, settings: Settings | None = None):
    if settings is not None and settings.data.get("positioning"):
        cached = getattr(settings, "_pos", None)
        if cached is None:
            settings._pos = cached = positioning_from_table(settings.data["positioning"])
        return cached
    loc = src.located("positioning")
    if loc:
        cached = getattr(loc, "_pos", None)
        if cached is None:
            table = bc.Table(loc.label, loc.header.sheet.grid)
            located = bc.locate_positioning(table, bc.load_settings(None))
            parsed = None
            if located:
                try:
                    parsed = bc.parse_positioning(located, loc.label, bc.builtin_monitoring())
                except bc.UserError as e:
                    src.warnings.append(f"{e} — взято встроенное позиционирование")
            loc._pos = cached = parsed or _builtin_positioning()
        return cached
    return _builtin_positioning()


def _builtin_positioning():
    global _BUILTIN_POSITIONING
    if _BUILTIN_POSITIONING is None:
        _BUILTIN_POSITIONING = bc.builtin_positioning(bc.load_settings(None))
    return _BUILTIN_POSITIONING


def known_suppliers(positioning) -> list[str]:
    seen: dict[str, str] = {}
    for suppliers in positioning.monitoring.values():
        for s in suppliers:
            seen.setdefault(supplier_key(s), s)
    return list(seen.values())


def monitoring_map(positioning) -> dict[str, set[str]]:
    """Бренд-конкурент (ключ) -> у каких поставщиков мониторить цену (ключи)."""
    cached = getattr(positioning, "_monitoring_map", None)
    if cached is None:
        cached = {brand_key(b): {supplier_key(s) for s in sup if s} for b, sup in positioning.monitoring.items() if sup}
        positioning._monitoring_map = cached
    return cached


def supplier_brand_book(positioning) -> dict:
    """Ключ поставщика -> каноническое имя и бренды, которые у него мониторят (лист 02)."""
    cached = getattr(positioning, "_supplier_brands", None)
    if cached is not None:
        return cached
    labels: dict[str, str] = {}
    brands: dict[str, list[str]] = {}
    for brand, suppliers in positioning.monitoring.items():
        for raw in suppliers:
            name = str(raw or "").strip()
            key = supplier_key(name)
            if not key:
                continue
            labels.setdefault(key, name)
            bucket = brands.setdefault(key, [])
            if brand not in bucket:
                bucket.append(brand)
    cached = {"labels": labels, "brands": brands}
    positioning._supplier_brands = cached
    return cached


def monitoring_suppliers(positioning) -> list[dict]:
    book = supplier_brand_book(positioning)
    return [{"supplier": book["labels"][key], "brands": book["brands"][key]} for key in book["labels"]]


# ---------------------------------------------------------------- сборка строк


def pick(mode: str, own, acp):
    """Значение поля по режиму источника; пустое (None / "") считается отсутствующим."""
    empty = lambda v: v is None or v == ""  # noqa: E731
    if mode == "own":
        return own
    if mode == "acp":
        return acp
    if mode == "acp_own":
        return own if empty(acp) else acp
    return acp if empty(own) else own


def groups(src: Sources, base: str) -> list[dict]:
    """Группы (ТМ + группа) в основе отчёта с подгруппами и количеством товаров."""
    counter: dict[tuple[str, str], Counter] = defaultdict(Counter)
    items = src.pst if base == "pst" else list(src.acp.values())
    for it in items:
        if skipped_order_group(getattr(it, "order_group", "")):
            continue
        counter[(it.brand, it.group)][it.subgroup] += 1
    result = [
        {"brand": b, "group": g, "count": sum(c.values()), "subgroups": [{"name": s, "count": n} for s, n in c.most_common()]}
        for (b, g), c in counter.items()
    ]
    return sorted(result, key=lambda x: -x["count"])


_CACHE_LOCK = threading.Lock()
_ROW_CACHE: OrderedDict = OrderedDict()
_INDEX_CACHE: OrderedDict = OrderedDict()
_FILE_OFFERS: OrderedDict = OrderedDict()


def _car_parser(settings: Settings) -> CarParser:
    parser = getattr(settings, "_cars_parser", None)
    if parser is None:
        settings._cars_parser = parser = CarParser(settings.cars)
    return parser


def build_rows(src: Sources, settings: Settings, opt: Options, positioning=None) -> list[Row]:
    """Строки группы. Повтор того же запроса копирует уже собранные, не разбирает ПСТ заново.

    Копия нужна: дальше на строку вешают кроссы и рекомендацию, а запомненный образец не трогают.
    """
    positioning = positioning or load_positioning(src, settings)
    key = (
        id(src), id(positioning), opt.brand, opt.group, opt.base,
        tuple(sorted((opt.sources or {}).items())),
        settings.calc_token(),
    )
    with _CACHE_LOCK:
        hit = _ROW_CACHE.get(key)
        if hit is not None:
            _ROW_CACHE.move_to_end(key)
    if hit is None:
        hit = _build_rows(src, settings, opt, positioning)
        with _CACHE_LOCK:
            _ROW_CACHE[key] = hit
            _ROW_CACHE.move_to_end(key)
            while len(_ROW_CACHE) > 12:
                _ROW_CACHE.popitem(last=False)
    return [copy.copy(row) for row in hit]


def _build_rows(src: Sources, settings: Settings, opt: Options, positioning) -> list[Row]:
    cars = _car_parser(settings)
    pst_by_code = {p.code: p for p in src.pst}
    gz_table = settings.gz | src.gz
    discounts = settings.discounts | src.discounts
    brand, group = norm_text(opt.brand), norm_text(opt.group)

    if opt.base == "acp":
        base = [(pst_by_code.get(a.code), a) for a in src.acp.values()]
    else:
        base = [(p, src.acp.get(p.code)) for p in src.pst]

    rows = []
    for p, a in base:
        it = p or a
        if brand and norm_text(it.brand) != brand or group and norm_text(it.group) != group:
            continue
        stock_own, sbs_own = src.sbs.get(it.code, (None, None))
        sales_acp = a.sold_qty * a.avg_sale_price if a and a.sold_qty is not None and a.avg_sale_price else None
        if a and a.sold_qty == 0:
            sales_acp = 0.0

        mo = pick(opt.source("mo"), p.mo_rub if p else None, a.mo_price if a else None)
        order_group = p.order_group if p else ""
        if skipped_order_group(order_group):
            continue
        subgroup = (p.subgroup if p else "") or (a.subgroup if a else "")
        discount = discounts.get((norm_text(subgroup), norm_text(order_group), norm_text(it.brand)))
        if discount is None and p is None and a and a.mo_price and a.cat_a_price:
            discount = round((1 - a.cat_a_price / a.mo_price) * 100, 2)
        cat_a_own = mo * (1 - discount / 100) if mo and discount is not None else None
        rate = settings.rate(p.purchase_currency) if p else None
        purchase_rub = p.purchase_price * rate if p and p.purchase_price and rate else None
        gz = gz_table.get(norm_text(order_group))
        sales_qty_own = src.sales_qty.get(it.code) if src.sales_qty else None

        row = Row(
            code=it.code,
            name=(p.name if p else "") or (a.name if a else ""),
            article=(p.article if p else "") or (a.article if a else ""),
            brand=it.brand,
            order_group=order_group,
            category=(p.category if p else "") or (a.category if a else ""),
            group=(p.group if p else "") or (a.group if a else ""),
            subgroup=subgroup,
            variety=p.variety if p else "",
            purchase_currency=p.purchase_currency if p else "",
            purchase_price=p.purchase_price if p else None,
            purchase_rub=purchase_rub,
            mo_currency=p.mo_currency if p else "",
            mo=mo,
            sales_qty=pick(opt.source("sales_qty"), sales_qty_own, a.sold_qty if a else None),
            sales=pick(opt.source("sales"), src.sales.get(it.code), sales_acp),
            stock=pick(opt.source("stock"), stock_own, a.stock if a else None),
            in_transit=src.transit.get(norm_article((p.article if p else "") or (a.article if a else ""))),
            abc=pick(opt.source("abc"), p.abc if p else "", a.abc if a else "") or "",
            sbs=pick(opt.source("sbs"), sbs_own, a.avg_cost if a else None),
            sbs_plan=purchase_rub * gz if purchase_rub and gz else None,
            cat_a=pick(opt.source("cat_a"), cat_a_own, a.cat_a_price if a else None),
            discount=discount,
            gz=gz,
            overstock=p.overstock if p else "",
            discontinued=p.discontinued if p else False,
            assortment=pick(opt.source("assortment"), p.assortment if p else "", a.assortment if a else "") or "",
            pst=p,
            acp=a,
        )
        row.cat_a_calc = cat_a_own is not None and row.cat_a == cat_a_own
        if not _subgroup_in_positioning(positioning, it.brand, it.group, subgroup):
            continue
        _, _, row.competitors, _ = positioning.resolve(row.brand, row.group, row.subgroup)
        row.percents = {c.key: _percent(c) for c in row.competitors}
        row.car = cars.parse(row.variety) if cars else None
        rows.append(row)
    return rows


def _percent(comp) -> float:
    """Процент именно этой строки позиционирования. «10 / 20» не склеиваем — берём число как написано."""
    value = comp.percents[0] if getattr(comp, "percents", None) else comp.percent
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = str(value).replace("%", "").replace("−", "-").replace(" ", "")
    match = re.match(r"-?\d+(?:[.,]\d+)?", text)
    if not match:
        return 0.0
    return float(match.group(0).replace(",", "."))


def missing_rates(rows: list[Row], settings: Settings) -> list[dict]:
    counter = Counter(
        r.purchase_currency for r in rows
        if r.purchase_price and r.purchase_currency and "?" not in r.purchase_currency and settings.rate(r.purchase_currency) is None
    )
    return [{"currency": c, "items": n} for c, n in counter.most_common()]


def missing_rates_from_pst(items, settings: Settings) -> list[dict]:
    """Валюты закупки по всем группам, для которых в настройках нет курса."""
    from .sources import norm_text

    counter: Counter = Counter()
    labels: dict[str, str] = {}
    seen: set[tuple[str, str]] = set()
    for item in items:
        cur = (getattr(item, "purchase_currency", "") or "").strip()
        if not cur or "?" in cur or not getattr(item, "purchase_price", None):
            continue
        if settings.rate(cur) is not None:
            continue
        key = norm_text(cur)
        ident = (str(getattr(item, "code", "") or ""), key)
        if ident in seen:
            continue
        seen.add(ident)
        counter[key] += 1
        labels.setdefault(key, cur)
    return [{"currency": labels[k], "items": n} for k, n in counter.most_common()]


# ---------------------------------------------------------------- фильтры


def _tri(mode: str, flag: bool) -> bool:
    """True — строка проходит фильтр «показывать / скрыть / только»."""
    return mode == "include" or (mode == "exclude" and not flag) or (mode == "only" and flag)


def _sold(r: Row) -> bool:
    return bool((r.sales_qty and r.sales_qty > 0) or (r.sales and r.sales > 0))


def is_clearance(r: Row) -> bool:
    """Распродажа и вывод: тип ассортимента R, категория ABC D или столбец «Вывод из ассортимента».

    NEW сюда не входит. Его смотрим только по типу ассортимента N, не по категории ABC.
    """
    abc = (getattr(r, "abc", "") or "").strip().upper().replace("Д", "D").replace(" ", "")
    assortment = (getattr(r, "assortment", "") or "").upper()
    return bool(getattr(r, "discontinued", False)) or assortment == "R" or abc == "D"


def skipped_order_group(name: str) -> bool:
    """Группа заказа, которую не берём ни в один отчёт."""
    return "СТОРОННИХ ПОСТАВЩИКОВ" in norm_text(name)


def _fold_clearance(filters: dict | None) -> dict:
    """Старый отдельный фильтр «Вывод» склеивается с распродажей."""
    raw = dict(filters or {})
    disc = raw.pop("discontinued", None)
    sale = raw.get("sale")
    rank = {"include": 0, "only": 1, "exclude": 2}
    chosen = sale if sale in rank else "include"
    if disc in rank and rank[disc] > rank[chosen]:
        chosen = disc
    if chosen != "include":
        raw["sale"] = chosen
    else:
        raw.pop("sale", None)
    return raw


def focus_stock(opt: Options) -> Options:
    """Отчёт распродажи или оверстока: в файл попадает только эта выборка.

    Остальные фильтры основного отчёта сохраняются. Противоположный фильтр
    не прячет строку: товар может быть и распродажей, и оверстоком сразу.
    """
    if opt.stock_slice not in ("sale", "overstock"):
        opt.stock_slice = ""
        return opt
    filters = dict(opt.filters)
    if opt.stock_slice == "sale":
        filters["sale"] = "only"
        filters.pop("overstock", None)
    else:
        filters["overstock"] = "only"
        filters.pop("sale", None)
    opt.filters = filters
    return opt


# Что можно вывести над таблицей: подпись и когда «Скрыть» прячет строку.
SHEET_FILTERS = {
    "overstock": ("Оверсток", lambda r: bool(r.overstock)),
    "sale": ("Распродажа и вывод", lambda r: is_clearance(r)),
    "competitor_price": ("Без цены", lambda r: not r.any_price),
    "new": ("NEW", lambda r: r.assortment == "N"),
    "mo": ("Без МО", lambda r: not r.mo),
    "cost": ("Без себестоимости", lambda r: not r.sbs and not r.sbs_plan),
    "stock_transit": ("Нет остатка и пути", lambda r: not ((r.stock and r.stock > 0) or (r.in_transit and r.in_transit > 0))),
    "stock": ("Без остатка", lambda r: not (r.stock and r.stock > 0)),
    "sales": ("Без продаж", lambda r: not _sold(r)),
}
DEFAULT_SHEET_FILTERS = ("overstock", "sale", "competitor_price")
# Старое имя: три фильтра, которые стоят над таблицей, пока список не меняли.
SHEET_SWITCHES = DEFAULT_SHEET_FILTERS


def sheet_filter_keys(opt: Options) -> list[str]:
    chosen = DEFAULT_SHEET_FILTERS if opt.sheet_filters is None else opt.sheet_filters
    return [k for k in chosen if k in SHEET_FILTERS]


def sheet_hide_flag(r: Row, key: str) -> bool:
    spec = SHEET_FILTERS.get(key)
    return bool(spec[1](r)) if spec else False


def sheet_switch_value(opt: Options, key: str) -> str:
    """Начальное положение списка. Совпадает с тем, что фильтр программы уже прячет."""
    mode = opt.filter(key)
    if key == "competitor_price":
        return "Скрыть" if mode == "only" else "Показать"
    if key in ("overstock", "sale", "new", "mo", "cost", "stock_transit") and mode == "exclude":
        return "Скрыть"
    return "Показать"


def exclusion_reason(r: Row, opt: Options, with_prices: bool = False, skip: set[str] | None = None) -> str:
    f = opt.filter
    sold = bool((r.sales_qty and r.sales_qty > 0) or (r.sales and r.sales > 0))
    checks = [
        ("overstock", bool(r.overstock), "оверсток", "не оверсток"),
        ("sale", is_clearance(r), "распродажа или вывод из ассортимента", "не распродажа и не вывод"),
        ("new", (r.assortment or "").upper() == "N", "NEW", "не NEW"),
        ("stock", bool(r.stock and r.stock > 0), "есть остаток", "нет остатка"),
        ("stock_transit", not (r.stock and r.stock > 0) and not (r.in_transit and r.in_transit > 0), "нет остатка и нет в пути", ""),
        ("sales", sold, "есть продажи", "нет продаж"),
        ("mo", not r.mo, "нет МО", ""),
        ("cost", not r.sbs and not r.sbs_plan, "нет себестоимости", ""),
    ]
    if with_prices:
        checks.append(("competitor_price", r.any_price, "есть цена конкурента", "нет цены конкурента"))
    skipped = skip or set()
    for key, flag, reason_yes, reason_no in checks:
        if key in skipped:
            continue
        mode = f(key)
        if not _tri(mode, flag):
            return reason_yes if mode == "exclude" else reason_no
    if opt.abc and (r.abc or "") not in opt.abc:
        return f"категория ABC «{r.abc or 'пусто'}»"
    return ""


def apply_filters(rows: list[Row], opt: Options, with_prices: bool = False) -> tuple[list[Row], list[tuple[Row, str]]]:
    kept, dropped = [], []
    for r in rows:
        reason = exclusion_reason(r, opt, with_prices)
        if reason:
            dropped.append((r, reason))
        else:
            kept.append(r)
    return kept, dropped


# ---------------------------------------------------------------- цены конкурентов


MIN_PRICE = 10.0  # 1, 2, 3 ₽ в прайсе — заглушка, не цена
ANOMALY_SHARE = 0.1  # ниже 10% уровня нормальных цен этого артикула — тоже не цена


def without_anomalies(offers: list[PriceOffer]) -> list[PriceOffer]:
    """Убирает 1–2 ₽ и цены, которые в разы ниже остальных предложений того же артикула.

    Одинокая небольшая цена остаётся: у дешёвой позиции может не быть с чем сравнивать.
    """
    kept = [o for o in offers if o.price and o.price >= MIN_PRICE]
    if len(kept) < 2:
        return kept
    ordered = sorted(kept, key=lambda o: o.price)
    upper = ordered[len(ordered) // 2:]
    anchor = upper[len(upper) // 2].price
    floor = anchor * ANOMALY_SHARE
    plausible = [o for o in ordered if o.price >= floor]
    return plausible or ordered


class PriceIndex:
    """Цены по (бренд-конкурент, артикул) из АЦПКД и/или прайсов поставщиков.

    Часть из прайсов одна на все группы: разбирается и очищается от выбросов один раз
    на набор файлов (_file_offers_clean). У группы своя только часть АЦПКД. Где артикул
    есть и в АЦПКД, и в прайсах, список собирается как раньше — сначала АЦПКД, потом
    прайсы — и очищается заново. Числа и порядок предложений те же, что при сборке
    одного общего словаря.
    """

    def __init__(self, src: Sources, opt: Options, price_files: list[PriceFile] | None = None, monitoring=None,
                 brands=None):
        self.by_source: Counter = Counter()
        self.skipped_by_monitoring = 0
        self.skipped_low = 0
        self._files: dict[tuple[str, str], list[PriceOffer]] = {}
        self._own: dict[tuple[str, str], list[PriceOffer]] = {}
        self._merged: dict | None = None
        in_stock = opt.in_stock_only
        acp: dict[tuple[str, str], list[PriceOffer]] = {}
        if opt.price_source in ("acp", "both"):
            seen: dict[tuple[str, str], set[tuple]] = defaultdict(set)
            for item in src.acp.values():
                for brand, crosses in item.competitors.items():
                    b = brand_key(brand)
                    for c in crosses:
                        key = (b, norm_article(c.article))
                        bucket = acp.setdefault(key, [])
                        known = seen[key]
                        for o in c.offers:
                            if not o.price or in_stock and o.qty is not None and o.qty <= 0:
                                continue
                            mark = (o.supplier, o.price)
                            if mark in known:
                                continue
                            known.add(mark)
                            bucket.append(PriceOffer(o.supplier, o.price, o.qty, "АЦПКД"))
                            self.by_source["АЦПКД"] += 1
        raw: dict[tuple[str, str], list[PriceOffer]] = {}
        if opt.price_source in ("prices", "both"):
            clean = _file_offers_clean(price_files, monitoring or {}, opt.monitoring_only, in_stock, brands)
            raw, self._files = clean.raw, clean.kept
            self.skipped_by_monitoring += clean.skipped
            self.skipped_low += clean.low
            if clean.total:
                self.by_source["Прайсы"] += clean.total
        for key, own in acp.items():
            from_files = raw.get(key)
            if from_files:
                # Эти цены прайсов уже очищены в общей части — пересчитываем вместе с АЦПКД.
                self.skipped_low -= len(from_files) - len(self._files.get(key, ()))
                offers = own + from_files
            else:
                offers = own
            kept = without_anomalies(offers)
            self.skipped_low += len(offers) - len(kept)
            self._own[key] = kept

    def get(self, comp_key: str, article: str) -> list[PriceOffer]:
        key = (brand_key(comp_key), norm_article(article))
        own = self._own.get(key)
        if own is not None:
            return own
        return self._files.get(key, [])

    @property
    def offers(self) -> dict[tuple[str, str], list[PriceOffer]]:
        """Все цены одним словарём, как раньше. В работе не нужен: get берёт из двух частей."""
        if self._merged is None:
            merged = dict(self._files)
            for key, kept in self._own.items():
                if kept:
                    merged[key] = kept
                else:
                    merged.pop(key, None)
            self._merged = merged
        return self._merged


@dataclass
class _CleanOffers:
    raw: dict
    kept: dict
    skipped: int
    low: int
    total: int


def _files_key(price_files) -> tuple:
    return tuple(id(pf) for pf in price_files) if price_files is not None else ()


def _cache_get(cache: OrderedDict, key: tuple, refs: tuple):
    """Попадание в кэш только если объекты те же самые, а не новые с тем же id."""
    with _CACHE_LOCK:
        hit = cache.get(key)
        if hit is None:
            return None
        held, value = hit
        if len(held) != len(refs) or any(a is not b for a, b in zip(held, refs)):
            return None
        cache.move_to_end(key)
        return value


def _cache_put(cache: OrderedDict, key: tuple, refs: tuple, value, limit: int) -> None:
    with _CACHE_LOCK:
        cache[key] = (refs, value)
        cache.move_to_end(key)
        while len(cache) > limit:
            cache.popitem(last=False)


def _file_offer_part(price_files, monitoring, monitoring_only: bool, in_stock: bool, brands=None):
    """Цены из прайсов. Одинаковы для всех групп, поэтому считаются один раз на набор файлов.

    Ключ — сами файлы, а не список: список последней даты каждый раз новый.
    brands — ключи брендов, по которым вообще ищут цену (конкуренты из позиционирования).
    Цены остальных брендов в индекс не кладутся: их всё равно никто не спрашивает.
    """
    offers, skipped, _other = _file_offer_part_ex(price_files, monitoring, monitoring_only, in_stock, brands)
    return offers, skipped


def _file_offer_part_ex(price_files, monitoring, monitoring_only: bool, in_stock: bool, brands=None):
    files = tuple(price_files or ())
    brands = frozenset(brands) if brands is not None else None
    key = ("raw", _files_key(price_files), id(monitoring), bool(monitoring_only), bool(in_stock), brands)
    refs = (*files, monitoring)
    hit = _cache_get(_FILE_OFFERS, key, refs)
    if hit is not None:
        return hit
    offers: dict[tuple[str, str], list[PriceOffer]] = {}
    skipped = other = 0
    keys: dict[str, str] = {}  # бренд как в прайсе -> ключ бренда; брендов мало, строк много
    for pf in files:
        if not pf.enabled or pf.problem and not pf.offers:
            continue
        sk = supplier_key(pf.supplier)
        label, source = pf.supplier_label, pf.sheet.file.name
        for o in pf.offers:
            b = keys.get(o.brand)
            if b is None:
                b = keys[o.brand] = brand_key(o.brand)
            allowed = monitoring.get(b)
            if monitoring_only and allowed and sk not in allowed:
                skipped += 1
                continue
            if in_stock and o.qty is not None and o.qty <= 0:
                continue
            if brands is not None and b not in brands:
                other += 1
                continue
            offers.setdefault((b, o.article_key or norm_article(o.article)), []).append(
                PriceOffer(label, o.price, o.qty, source)
            )
    stored = (offers, skipped, other)
    _cache_put(_FILE_OFFERS, key, refs, stored, 8)
    return stored


def _file_offers_clean(price_files, monitoring, monitoring_only: bool, in_stock: bool, brands=None) -> _CleanOffers:
    """Цены прайсов после отсечения заглушек и выбросов — одни на все группы."""
    files = tuple(price_files or ())
    brands = frozenset(brands) if brands is not None else None
    key = ("clean", _files_key(price_files), id(monitoring), bool(monitoring_only), bool(in_stock), brands)
    refs = (*files, monitoring)
    hit = _cache_get(_FILE_OFFERS, key, refs)
    if hit is not None:
        return hit
    raw, skipped, other = _file_offer_part_ex(price_files, monitoring, monitoring_only, in_stock, brands)
    kept: dict[tuple[str, str], list[PriceOffer]] = {}
    low = 0
    total = other  # «цен из прайсов» в окне — все строки, как раньше, хотя в индекс легли только конкуренты
    for k, offers in raw.items():
        total += len(offers)
        clean = without_anomalies(offers)
        low += len(offers) - len(clean)
        if clean:
            # Одно предложение возвращается тем же — не держим второй такой же список.
            kept[k] = offers if len(offers) == 1 and len(clean) == 1 else clean
    stored = _CleanOffers(raw, kept, skipped, low, total)
    _cache_put(_FILE_OFFERS, key, refs, stored, 8)
    return stored


def competitor_brand_keys(positioning) -> frozenset:
    """Ключи всех брендов-конкурентов: лист мониторинга и конкуренты в строках позиционирования."""
    cached = getattr(positioning, "_competitor_keys", None)
    if cached is None:
        found = {brand_key(b) for b in getattr(positioning, "monitoring", {}) or {}}
        for row in getattr(positioning, "rows", None) or []:
            for comp in getattr(row, "competitors", None) or []:
                found.add(brand_key(getattr(comp, "name", "")))
                found.add(brand_key(getattr(comp, "key", "")))
        found.discard("")
        cached = frozenset(found)
        try:
            positioning._competitor_keys = cached
        except AttributeError:
            pass
    return cached


def index_for(src: Sources, opt: Options, price_files: list[PriceFile] | None, monitoring, brands=None) -> PriceIndex:
    brands = frozenset(brands) if brands is not None else None
    key = (
        id(src), _files_key(price_files), opt.price_source,
        bool(opt.monitoring_only), bool(opt.in_stock_only), id(monitoring), brands,
    )
    refs = (src, *(price_files or ()), monitoring)
    hit = _cache_get(_INDEX_CACHE, key, refs)
    if hit is not None:
        return hit
    index = PriceIndex(src, opt, price_files, monitoring, brands)
    _cache_put(_INDEX_CACHE, key, refs, index, 24)
    return index


def acp_crosses(item: AcpItem | None, comp_key: str):
    if not item:
        return []
    b = brand_key(comp_key)
    return next((v for k, v in item.competitors.items() if brand_key(k) == b), [])


def select(mode: str, svod: list, acp: list, index, comp_key: str, has_svod: bool = True) -> list[Pick]:
    if mode == "off":
        return []
    if not has_svod:
        return [Pick(c.article.upper(), None, index.get(comp_key, c.article)) for c in acp]

    def from_svod(rows):
        return [Pick(c.cross_article, c.yes, index.get(comp_key, c.cross_key)) for c in rows]

    yes = [c for c in svod if c.yes]
    no = [c for c in svod if not c.yes]
    if mode == "yes":
        return from_svod(yes)
    if mode == "no":
        return from_svod(no)
    if mode == "yes_no":
        return from_svod(yes + no)
    if mode == "yes_fill":
        picks = from_svod(yes)
        if not any(p.offers for p in picks):
            picks += from_svod(no)
        return picks
    raise ValueError(mode)


def svod_for(src: Sources, r: Row) -> dict[str, list]:
    return src.crosses.get((norm_text(r.brand), norm_article(r.article)), {})


def svod_comp(svod: dict[str, list], comp_key: str) -> list:
    if comp_key in svod:
        return svod[comp_key]
    b = brand_key(comp_key)
    return next((v for k, v in svod.items() if brand_key(k) == b), [])


# ---------------------------------------------------------------- аналитика кроссов


def _pct(part: int, total: int) -> str:
    return f"{part / total:.0%}" if total else "0%"


RESULT_BUCKETS = (
    ("target", "Целевые группы"),
    ("small", "Мелкие группы"),
    ("noprice", "Нет цен"),
)


def result_bucket(brand: str, group: str, priced: int) -> str:
    """Куда класть файл.

    Целевые — диски и колодки ADVICS и AVANTECH, фильтры AVANTECH и DOUBLE FORCE:
    по ним судим объём и продажи, и только если есть цена конкурента.
    Мелкие — цена конкурента есть у любой другой группы. Нет цен — сравнивать не с чем.
    """
    if not priced:
        return "noprice"
    if sister_of(brand, group):
        return "target"
    return "small"


def result_folder_name(bucket: str) -> str:
    return dict(RESULT_BUCKETS)[bucket]


def group_price_gap(src: Sources, settings: Settings, index: PriceIndex, positioning, brand: str, group: str) -> tuple[int, str]:
    """Сколько товаров с ценой конкурента. Если ноль: «cross» — кроссов нет, «price» — кроссы есть, цен нет.

    Без сборки полной строки отчёта: марки авто здесь не нужны, а на всём ассортименте
    они заставляли окно крутить загрузку при каждом открытии.
    """
    del settings
    want_brand, want_group = norm_text(brand), norm_text(group)
    any_cross = False
    priced = 0
    has_svod = bool(src.crosses)
    for item in src.pst:
        if want_brand and norm_text(item.brand) != want_brand:
            continue
        if want_group and norm_text(item.group) != want_group:
            continue
        if skipped_order_group(item.order_group):
            continue
        if not _subgroup_in_positioning(positioning, item.brand, item.group, item.subgroup):
            continue
        _, _, comps, _ = positioning.resolve(item.brand, item.group, item.subgroup)
        article = norm_article(item.article)
        svod_map = src.crosses.get((norm_text(item.brand), article), {})
        acp_item = src.acp.get(item.code)
        hit = False
        for comp in comps:
            svod = svod_comp(svod_map, comp.key)
            acp = acp_crosses(acp_item, comp.key)
            if svod or acp:
                any_cross = True
            if any(pick.offers for pick in select("yes_fill", svod, acp, index, comp.key, has_svod)):
                hit = True
                break
        if hit:
            priced += 1
    if priced:
        return priced, ""
    return 0, "price" if any_cross else "cross"


def suggest(s: dict, has_svod: bool, threshold: float) -> tuple[str, str]:
    n = s["items"]
    p = s["priced"]
    if not has_svod:
        return "yes", f"Файла «Кроссы свод» нет — кроссы из АЦПКД, цена у {p['yes']} из {n}."
    if not s["yes"] and not s["no_only"]:
        return "yes_fill", "Кроссов на этот бренд нет."
    if not s["yes"]:
        return "yes_fill", f"«Да» нет совсем, берём «Нет»: цена будет у {p['yes_fill']} из {n}."
    share = s["yes"] / n if n else 0
    if share >= threshold:
        verdict = f"«Да» у {s['yes']} из {n} ({_pct(s['yes'], n)}) — кроссов хватает."
        if s["yes"] and p["yes"] == 0:
            verdict += " Цен по этим артикулам нет: прайс этого бренда не подключён или в АЦПКД пусто."
        elif p["yes_fill"] > p["yes"]:
            verdict += f" С добором «Нет» цена будет у {p['yes_fill']} вместо {p['yes']}."
        return "yes", verdict
    return "yes_fill", (
        f"«Да» только у {s['yes']} из {n} ({_pct(s['yes'], n)}, {s['yes_articles']} разных) — мало, "
        f"добираем «Нет»: цена будет у {p['yes_fill']} вместо {p['yes']}."
    )


def analyze_crosses(src: Sources, rows: list[Row], opt: Options, index: PriceIndex) -> list[dict]:
    """По каждому конкуренту и сегменту (ABC / подгруппа): кроссы «Да»/«Нет» и сколько товаров получат цену в каждом режиме."""
    has_svod = bool(src.crosses)
    comps: dict[str, dict] = {}
    for r in rows:
        svod = svod_for(src, r)
        seg_key, sub, abc = segment(r, opt.cross_split)
        for comp in r.competitors:
            c = comps.setdefault(comp.key, {"key": comp.key, "name": comp.name, "order": comp.order, "segments": {}})
            s = c["segments"].setdefault(
                seg_key,
                {"seg": seg_key, "subgroup": sub, "abc": abc, "items": 0, "yes": 0, "no_only": 0, "none": 0, "acp": 0,
                 "yes_articles": set(), "priced": Counter()},
            )
            crosses = svod_comp(svod, comp.key)
            acp = acp_crosses(r.acp, comp.key)
            yes = [x for x in crosses if x.yes]
            s["items"] += 1
            s["yes"] += bool(yes)
            s["no_only"] += bool(crosses) and not yes
            s["none"] += not crosses
            s["acp"] += bool(acp)
            s["yes_articles"].update(x.cross_key for x in yes)
            priced = _priced_modes(crosses, acp, index, comp.key, has_svod)
            for mode in priced:
                s["priced"][mode] += 1

    result = []
    for c in sorted(comps.values(), key=lambda c: (c["order"], c["name"])):
        segs = []
        for s in c["segments"].values():
            s["yes_articles"] = len(s["yes_articles"])
            s["priced"] = {m: s["priced"].get(m, 0) for m in CROSS_MODES if m != "off"}
            s["priced"]["off"] = 0
            s["suggested"], s["verdict"] = suggest(s, has_svod, opt.suggest_threshold)
            s["mode"] = mode_for(opt, c["key"], s["seg"], s["suggested"])
            s["explicit"] = f"{c['key']}|{s['seg']}" in opt.cross_modes
            segs.append(s)
        segs.sort(key=lambda s: (norm_text(s["subgroup"]), ABC_ORDER.get(s["abc"], 99), s["abc"]))
        c["segments"] = segs
        for k in ("items", "yes", "no_only", "none", "acp", "yes_articles"):
            c[k] = sum(s[k] for s in segs)
        c["priced"] = sum(s["priced"][s["mode"]] for s in segs)
        c["mode_all"] = opt.cross_modes.get(f"{c['key']}|*")
        result.append(c)
    return result


def _priced_modes(svod, acp, index: PriceIndex, comp_key: str, has_svod: bool) -> set[str]:
    """Какие режимы кросса дадут цену. Тот же ответ, что четыре вызова select, но за один проход."""
    def has_offers(rows, article_of) -> bool:
        return any(index.get(comp_key, article_of(row)) for row in rows)

    if not has_svod:
        if not has_offers(acp, lambda row: row.article):
            return set()
        return {"yes", "no", "yes_no", "yes_fill"}
    yes = [row for row in svod if row.yes]
    no = [row for row in svod if not row.yes]
    yes_hit = has_offers(yes, lambda row: row.cross_key)
    no_hit = has_offers(no, lambda row: row.cross_key)
    found = set()
    if yes_hit:
        found.update(("yes", "yes_no", "yes_fill"))
    if no_hit:
        found.add("no")
        found.add("yes_no")
        if not yes_hit:
            found.add("yes_fill")
    return found


def suggestions_map(analysis: list[dict]) -> dict[tuple[str, str], str]:
    return {(c["key"], s["seg"]): s["mode"] for c in analysis for s in c["segments"]}


def attach_crosses(src: Sources, rows: list[Row], opt: Options, index: PriceIndex, modes: dict[tuple[str, str], str]) -> None:
    has_svod = bool(src.crosses)
    for r in rows:
        svod = svod_for(src, r)
        seg_key, _, _ = segment(r, opt.cross_split)
        r.picks, r.prices, r.best = {}, {}, {}
        for comp in r.competitors:
            mode = modes.get((comp.key, seg_key)) or mode_for(opt, comp.key, seg_key, "yes_fill")
            if mode == "off":
                continue
            picks = select(mode, svod_comp(svod, comp.key), acp_crosses(r.acp, comp.key), index, comp.key, has_svod)
            r.picks[comp.key] = picks
            offers = [o for p in picks for o in p.offers]
            best = min(offers, key=lambda o: o.price) if offers else None
            r.prices[comp.key] = best.price if best else None
            if best:
                r.best[comp.key] = best


def compare_sources(rows: list[Row], src: Sources, opt: Options, settings: Settings) -> list[dict]:
    """Сколько товаров расходятся между нашими выгрузками и АЦПКД по каждому полю (для выбора источника)."""
    own = Options.from_dict({**opt.__dict__, "sources": {k: "own" for k in FIELDS}})
    acp = Options.from_dict({**opt.__dict__, "sources": {k: "acp" for k in FIELDS}})
    a_rows = {r.code: r for r in build_rows(src, settings, acp)}
    o_rows = {r.code: r for r in build_rows(src, settings, own)}
    codes = {r.code for r in rows}
    result = []
    for key, (title, own_title, acp_title) in FIELDS.items():
        both = differ = only_own = only_acp = 0
        for code in codes:
            o, a = getattr(o_rows.get(code), key, None), getattr(a_rows.get(code), key, None)
            has_o, has_a = o not in (None, ""), a not in (None, "")
            if has_o and has_a:
                both += 1
                if isinstance(o, str) or isinstance(a, str):
                    differ += str(o) != str(a)
                elif abs(o - a) > 0.01 * max(abs(o), abs(a)):
                    differ += 1
            else:
                only_own += has_o and not has_a
                only_acp += has_a and not has_o
        result.append({"key": key, "title": title, "own": own_title, "acp": acp_title,
                       "both": both, "differ": differ, "only_own": only_own, "only_acp": only_acp,
                       "source": opt.source(key)})
    return result


@dataclass
class Prepared:
    rows: list[Row]  # все товары группы
    kept: list[Row]  # после всех фильтров, с кроссами и ценами
    dropped: list[tuple[Row, str]]
    analysis: list[dict]
    index: PriceIndex
    positioning: object


def rows_for_workbook(prep: Prepared, src: Sources, opt: Options) -> list[Row]:
    """В файл попадают только строки, прошедшие фильтры программы.

    Раньше отсечённые товары возвращались в лист и прятались строкой «Показать / Скрыть».
    Excel из‑за этой строки удаляет лист «Отчет», поэтому в файл они больше не кладутся.
    """
    del src, opt
    return list(prep.kept)


def prepare(src: Sources, settings: Settings, opt: Options, price_files: list[PriceFile] | None = None) -> Prepared:
    positioning = load_positioning(src, settings)
    rows = build_rows(src, settings, opt, positioning)
    kept, dropped = apply_filters(rows, opt)
    index = index_for(src, opt, price_files, monitoring_map(positioning), competitor_brand_keys(positioning))
    analysis = analyze_crosses(src, kept, opt, index)
    attach_crosses(src, kept, opt, index, suggestions_map(analysis))
    kept, dropped_price = apply_filters(kept, opt, with_prices=True)
    apply_recommendations(kept, opt)
    kept.sort(key=lambda r: (-(r.sales_qty or 0), -(r.sales or 0), r.article))
    apply_sisters(kept, src)
    return Prepared(rows, kept, dropped + dropped_price, analysis, index, positioning)


def _subgroup_in_positioning(positioning, brand: str, group: str, subgroup: str) -> bool:
    """Подгруппа участвует, только если она есть в позиционировании этой ТМ и группы.

    Нет строки позиционирования по группе — не отсекаем: сравнивать не с чем.
    """
    b, g = norm_text(brand), norm_text(group)
    allowed = {
        norm_text(row.subgroup)
        for row in positioning.rows
        if norm_text(row.brand) == b and norm_text(row.group) == g and row.subgroup
    }
    if not allowed:
        return True
    return norm_text(subgroup) in allowed


def competitor_order(rows: list[Row]) -> list:
    """Столбец цены конкурента — если цена есть хотя бы у одного товара.

    Кросс без цены столбец не создаёт. Где стоит цена, вверху отчёта или внизу, не важно.
    """
    found: dict[str, object] = {}
    for r in rows:
        for c in r.competitors:
            if r.picks.get(c.key) is None:
                continue
            if r.prices.get(c.key) in (None, ""):
                continue
            found.setdefault(c.key, c)
    return sorted(found.values(), key=lambda c: (c.order, c.name))


def apply_recommendations(rows: list[Row], opt: Options) -> None:
    """Цель по позиционированию на каждый товар; рекомендованная Кат. А — если функция включена."""
    ro = opt.recs
    weights = weight_classes(rows, ro)
    for r in rows:
        order = [c.key for c in sorted(r.competitors, key=lambda c: c.order) if c.key in r.picks]
        names = {c.key: c.name for c in r.competitors}
        r.weight = weights.get(r.code, "tail")
        # Топ смотрит только на первого в позиционировании. Нет его цены — G-Brake и прочие не подставляем.
        mode = "first" if r.weight == "top" else ro.target
        r.target, by = target_for(r.prices, r.percents, order, mode)
        r.target_by = names.get(by, by)
        r.rec = None
        if ro.enabled:
            segment = r.car.segment if r.car else "Не определён"
            fleet = (r.car.fleet if r.car else "") or ""
            if r.overstock or is_clearance(r):
                cap = 0.0
            elif any(part in fleet.lower() for part in ("европ", "америк")):
                cap = 0.02
            else:
                cap = None
            r.rec = recommend(
                r.cat_a, r.sbs if r.sbs is not None else r.sbs_plan, r.target, segment, r.weight, ro,
                r.target_by, cap, r.abc,
            )


# ---------------------------------------------------------------- книга


def segment_label(s: dict, split: str) -> str:
    parts = []
    if split in ("subgroup", "subgroup_abc"):
        parts.append(s["subgroup"] or "без подгруппы")
    if split in ("abc", "subgroup_abc"):
        parts.append(f"ABC: {s['abc'] or 'пусто'}")
    return " · ".join(parts) or "все товары"


def build_report(src: Sources, settings: Settings, opt: Options, out_dir: Path,
                 price_files: list[PriceFile] | None = None):
    from .workbook import build_report as build

    return build(src, settings, opt, out_dir, price_files)
