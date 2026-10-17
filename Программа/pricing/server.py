"""Локальный сервер окна программы: отдаёт страницу и отвечает на запросы JSON.

Работает только на 127.0.0.1 — снаружи компьютера недоступен.
"""

from __future__ import annotations

import datetime as dt
import json
import os
import subprocess
import sys
import threading
import traceback
from collections import Counter
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from . import bootlog, check, crossdump, dynamics, filecache, insight, price_analytics, price_check, pricearchive, prices, pricesort, report
from .cars import SEGMENTS, UNKNOWN, load_default as cars_default
from .dashboard import summary as dashboard_summary
from .excel import clean, num
from .recommend import DEFAULT_UP, TARGETS, BUNDLED_GROUPS, RecOptions, bundled_for, bundled_rec
from .settings import ROOT, USER_FILE, Settings
from .buildinfo import BUILD, BUILD_LABEL, column_expectation, program_marks
from .sisters import price_book, sister_of, sister_title
from .sources import (
    DATA_DIRS, KIND_TITLES, Sources, Workspace, acp_covers, apply_company, attach_shared, collect_files,
    data_paths, dormant_with_settings, ensure_data_dirs, norm_text, scan, scan_paths, scan_workspace, wake_dormant,
)

WEB = Path(__file__).with_name("web")
EXPECTED = ["pst", "sales", "sales3", "sbs", "transit", "acp", "crosses", "discounts", "gz", "positioning", "oem"]
DOC_HINTS = {
    "pst": "основа отчёта: товары, цены МО, закупка, ABC",
    "sales": "количество и сумма продаж",
    "sales3": "продажи за последние 3 месяца, та же выгрузка 1С; в имени файла «3 мес» или период в шапке. По ним видно, продаётся ли товар сейчас",
    "sbs": "остаток и себестоимость",
    "transit": "заказано, шт — колонка «В пути» рядом с остатком. Файл группы или один на всех",
    "acp": "цены конкурентов. Файл группы или несколько файлов по категориям во «Входящие»",
    "oem": "номера ОЕМ для цены смежной марки: ADVICS–AVANTECH (диски и колодки) и AVANTECH–DF (фильтры)",
    "crosses": "«Кроссы свод» с признаком «Да»/«Нет»",
    "discounts": "если нет — берутся из настроек",
    "gz": "если нет — берутся из настроек",
    "positioning": "если нет — встроенное",
}
CONTENT_TYPES = {".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8",
                 ".css": "text/css; charset=utf-8", ".svg": "image/svg+xml"}


class State:
    def __init__(self):
        self.lock = threading.Lock()
        self.cache: dict[str, tuple[tuple, Sources]] = {}
        self.workspace_cache: tuple[tuple, Workspace] | None = None
        # (листы прайсов, настройки и позиционирование, разобранные файлы) — общие для групп.
        self.price_shared: list[tuple[tuple, tuple, list]] = []

    def sources(self, folder: Path) -> Sources:
        paths, price_dirs = scan_paths(folder)
        recursive = bool(price_dirs) or any(p.is_dir() and p.name in DATA_DIRS.values() for p in paths)
        files = collect_files(paths, recursive)
        signature = tuple((str(f), f.stat().st_mtime_ns, f.stat().st_size) for f in files)
        key = str(folder)
        with self.lock:
            cached = self.cache.get(key)
            if cached and cached[0] == signature:
                return cached[1]
        src = scan(paths, price_dirs, all_dates=False)
        src._dormant_owner = src  # type: ignore[attr-defined]
        with self.lock:
            self.cache[key] = (signature, src)
        return src

    def workspace(self, root: Path) -> Workspace:
        ensure_data_dirs(root)
        paths = data_paths(root)
        folders = [paths[k] for k in ("groups", "crosses", "prices", "legacy") if paths[k].is_dir()]
        old = root / "Выгрузки 1С"
        if old.is_dir():
            folders.append(old)
        files = collect_files(folders, True)
        signature = tuple((str(f), f.stat().st_mtime_ns, f.stat().st_size) for f in files)
        with self.lock:
            cached = self.workspace_cache
            if cached and cached[0] == signature:
                return cached[1]
        bootlog.step("Кэш", "открываю сохранённый разбор")
        stored = filecache.load_workspace(signature)
        if stored is not None:
            bootlog.step("Кэш", "сохранённый разбор подошёл")
            stored._file_signature = signature
            with self.lock:
                self.workspace_cache = (signature, stored)
            return stored
        bootlog.step("Чтение папок", "кэша нет или файлы новее")
        ws = scan_workspace(root)
        ws._file_signature = signature
        filecache.save_workspace(signature, ws)
        with self.lock:
            self.workspace_cache = (signature, ws)
        return ws

    def for_group(self, root: Path, brand: str, group: str, *, sisters: bool = True) -> Sources:
        ws = self.workspace(root)
        token = Settings.load().calc_token() if sisters else 0
        stored = getattr(ws, "_group_src", None)
        if stored is None:
            ws._group_src = stored = {}
        found = stored.get((brand, group, sisters, token))
        if found is not None:
            return found
        pack = _pick_pack(ws, brand, group)
        if pack is None:
            ready = getattr(ws, "_company_ready", None)
            if ready is None:
                ready = attach_shared(apply_company(Sources({}, []), ws.company), ws.shared)
                ws._company_ready = ready
            base = ready
        else:
            ready = getattr(pack, "_ready", None)
            if ready is None:
                ready = attach_shared(apply_company(pack.sources, ws.company), ws.shared)
                pack._ready = ready
            base = ready
        src = _with_incoming_acp(base, ws, brand, group, trust_local=bool(pack and _is_group_pack(pack)))
        src.oem_rows = list(getattr(ws, "oem_rows", None) or [])
        if any(not item.brand for item in src.oem_rows):
            from .sources import fill_oem_brands

            fill_oem_brands(src.oem_rows, getattr(src, "pst", None) or [])
        src.sister_prices = price_book(ws, Settings.load()) if sisters else {}
        src._dormant_owner = ws.shared  # type: ignore[attr-defined]
        src._group_args = (root, brand, group, sisters)  # type: ignore[attr-defined]
        locs = list(getattr(ws, "oem_locs", None) or [])
        if locs:
            src.found = dict(src.found)
            src.found["oem"] = locs
        stored[(brand, group, sisters, token)] = src
        return src

    def price_files(self, src: Sources, settings: Settings) -> list[prices.PriceFile]:
        """Прайсы с учётом сохранённых форматов; пересчитываются, только если что-то поменялось.

        Разобранные прайсы общие для всех групп с теми же листами и тем же позиционированием.
        Раньше каждая группа разбирала все прайсы заново и держала в памяти свою копию —
        на «Все группы» с большими прайсами программе не хватало памяти.
        """
        base = json.dumps([settings.price_profiles, prices.alias_token(), prices.city_token()],
                          sort_keys=True, ensure_ascii=False)
        own = json.dumps(settings.price_files, sort_keys=True, ensure_ascii=False)
        key = base + "\n" + own
        with self.lock:
            cached = getattr(src, "_price_cache", None)
            if cached and cached[0] == key:
                return cached[1]
        positioning = report.load_positioning(src)
        known = report.known_suppliers(positioning)
        brands = dynamics.competitor_brand_names(positioning)
        sheets = tuple(src.price_sheets)
        cross_brands, cross_pairs = cross_keys(src)
        shared_key = (base, tuple(known), tuple(brands), len(cross_pairs))
        files = None
        keep = None
        with self.lock:
            for held_sheets, held_key, held_own, held_files in self.price_shared:
                if held_key == shared_key and held_own == own and len(held_sheets) == len(sheets) and all(
                    a is b for a, b in zip(held_sheets, sheets)
                ):
                    files = held_files
                    break
            if files is None:
                held = next((h for h in self.price_shared if h[1] == shared_key), None)
                if held is not None:
                    keep = _reusable(held[3], json.loads(held[2]), settings.price_files, sheets)
        if files is None:
            # Прошлые даты не прочитаны, но их ручные столбцы учат следующие даты поставщика.
            owner = getattr(src, "_dormant_owner", None)
            extra = dormant_with_settings(owner, settings.price_files) if owner is not None else []
            learn = None
            if extra:
                order = getattr(owner, "price_order", None) or {}
                big = len(order) + 1
                both = list(src.price_sheets) + extra
                learn = [s for _, s in sorted(enumerate(both), key=lambda it: (order.get(str(it[1].file), big), it[0]))]
            files = prices.load_price_files(
                src.price_sheets, settings.price_profiles, settings.price_files, known, brands, learn_sheets=learn,
                known_brands=cross_brands, cross_pairs=cross_pairs, keep=keep,
            )
            with self.lock:
                self.price_shared = [(sheets, shared_key, own, files), *self.price_shared][:3]
        with self.lock:
            src._price_cache = (key, files)  # type: ignore[attr-defined]
        return files


STATE = State()


def _reusable(old_files: list, old_settings: dict, new_settings: dict, sheets: tuple) -> dict:
    """Какие уже разобранные прайсы можно не разбирать заново после правки настроек.

    Правка одного файла меняет его самого и, через «столбцы поставщика», другие файлы того же
    поставщика. Файлы, узнанные по сохранённой форме или не узнанные, тоже перепроверяются:
    их может подхватить новая форма. Остальные — те же объекты, без чтения и проверок.
    """
    changed = {k for k in set(old_settings) | set(new_settings) if old_settings.get(k) != new_settings.get(k)}
    current = {id(s) for s in sheets}
    suppliers = set()
    for k in changed:
        for side in (old_settings.get(k) or {}, new_settings.get(k) or {}):
            if side.get("supplier"):
                suppliers.add(prices.supplier_key(side["supplier"]))
        parts = [p.strip() for p in k.split(" / ")]
        for part in parts[:2]:
            if part and not prices._normalize_date(part):
                suppliers.add(prices.supplier_key(part))
                break
    keep = {}
    for pf in old_files:
        if id(pf.sheet) not in current:
            continue
        ident = prices.price_id(pf.sheet)
        if ident in changed or pf.sheet.label in changed:
            continue
        if changed and (pf.status in ("saved", "none")
                        or prices.supplier_key(pf.folder or pf.supplier) in suppliers
                        or prices.supplier_key(pf.supplier) in suppliers):
            continue
        keep[id(pf.sheet)] = pf
    return keep


_CROSS_KEYS: list = []  # [(словарь кроссов, (бренды, пары))] — держим ссылку, чтобы id не переиспользовался


def cross_keys(src: Sources) -> tuple[set, set]:
    """Ключи брендов из кроссов и пары (бренд кросса, артикул кросса) — чтобы проверять столбцы прайсов."""
    crosses = getattr(src, "crosses", None) or {}
    for held, value in _CROSS_KEYS:
        if held is crosses:
            return value
    brands: set = set()
    pairs: set = set()
    for (own_brand, _), by_brand in crosses.items():
        brands.add(prices.brand_key(own_brand))
        for rows in by_brand.values():
            for c in rows:
                bk = prices.brand_key(c.cross_brand)
                brands.add(bk)
                pairs.add((bk, c.cross_key))
    brands.discard("")
    value = (brands, pairs)
    _CROSS_KEYS[:] = [(crosses, value), *_CROSS_KEYS][:2]
    return value


def _needs_old_dates(owner: Sources, settings: Settings) -> bool:
    """Прайсы прошлых дат нужны отчёту, только если последней даты по сути нет.

    Так бывает, когда все файлы последней даты выключены или ни один не прочитался.
    Тогда, как и раньше, отчёт берёт предыдущую дату.
    """
    if not getattr(owner, "dormant_prices", None):
        return False
    dated = [s for s in owner.price_sheets if prices.price_date(s.file)]
    return not any(prices.file_settings(settings.price_files, s).get("enabled", True) for s in dated)


def _forget_groups(ws: Workspace) -> None:
    """Собранные группы держат копию списка прайсов — после дочитки собираем их заново."""
    ws._group_src = {}  # type: ignore[attr-defined]
    ws._company_ready = None  # type: ignore[attr-defined]
    for pack in ws.packs:
        pack.__dict__.pop("_ready", None)


def _wake_owner(owner: Sources) -> bool:
    """Дочитать прайсы прошлых дат. Разобранные прайсы этого источника считаются заново."""
    with STATE.lock:
        woke = wake_dormant(owner)
        if woke:
            owner.__dict__.pop("_price_cache", None)
            owner.__dict__.pop("_price_payload", None)
            ws = STATE.workspace_cache[1] if STATE.workspace_cache else None
            if ws is not None and ws.shared is owner:
                _forget_groups(ws)
    return woke


def _all_dates_sources(body: dict, settings: Settings) -> tuple[Sources, bool]:
    """Источник с прайсами всех дат: для динамики, устройства прайсов и правки старого файла."""
    src = _loaded_sources(body, settings)
    owner = getattr(src, "_dormant_owner", None)
    if owner is None or not getattr(owner, "dormant_prices", None):
        return src, False
    woke = _wake_owner(owner)
    return (_loaded_sources(body, settings) if woke else src), woke


def describe_price(pf: prices.PriceFile, positioning) -> dict:
    token = getattr(positioning, "_mon_token", None)
    if token is None:
        token = repr(sorted((str(k), list(v)) for k, v in positioning.monitoring.items()))
        try:
            positioning._mon_token = token
        except AttributeError:
            pass
    token = (token, prices.alias_token())
    held = pf.__dict__.get("_describe")
    if held is not None and held[0] == token:
        return held[1]
    d = _describe_price(pf, positioning)
    if not pf.__dict__.get("_preview"):
        pf.__dict__["_describe"] = (token, d)
    return d


def _describe_price(pf: prices.PriceFile, positioning) -> dict:
    competitors = {prices.brand_key(b): b for b in positioning.monitoring}
    keys: dict = {}  # бренд как в прайсе -> ключ; брендов в прайсе мало, строк много

    def key_of(brand):
        k = keys.get(brand)
        if k is None:
            k = keys[brand] = prices.brand_key(brand)
        return k

    found = Counter(competitors[k] for k in (key_of(o.brand) for o in pf.offers) if k in competitors)
    d = prices.describe(pf)
    ours = [o for o in pf.offers if key_of(o.brand) in competitors]
    if ours:
        d["preview"] = [{"brand": o.brand, "article": o.article, "price": o.price, "qty": o.qty} for o in ours[:6]]
    book = report.supplier_brand_book(positioning)
    key = prices.supplier_key(pf.supplier)
    expected = book["brands"].get(key, [])
    found_keys = {prices.brand_key(b) for b in found}
    d["monitored"] = key in book["labels"]
    d["competitor_brands"] = [{"brand": b, "offers": n} for b, n in found.most_common()]
    d["expected_brands"] = expected
    d["missing_brands"] = [b for b in expected if prices.brand_key(b) not in found_keys]
    return d


def price_payload(src: Sources, settings: Settings) -> list[dict]:
    files = STATE.price_files(src, settings)
    cached = getattr(src, "_price_payload", None)
    if cached and cached[0] is files:
        return cached[1]
    positioning = report.load_positioning(src, settings)
    payload = [describe_price(pf, positioning) for pf in files]
    try:
        notes, info = price_check.quality_split(files, dynamics.competitor_brand_names(positioning), cross_keys(src)[1])
    except Exception:  # noqa: BLE001 — проверка не должна ронять список прайсов
        notes, info = {}, {}
    payload = [dict(item) for item in payload]
    for item in payload:
        if notes.get(item["id"]):
            item["quality"] = notes[item["id"]]
        if info.get(item["id"]):
            item["info"] = info[item["id"]]
    src._price_payload = (files, payload)
    return payload


def brand_hints_payload(src: Sources, settings: Settings) -> list[dict]:
    """Похожие на конкурентов бренды, которые не совпали. Считается раз на набор прайсов."""
    files = STATE.price_files(src, settings)
    cached = getattr(src, "_brand_hints", None)
    ignore_token = tuple(settings.options.get("brand_ignore") or [])
    if cached and cached[0] is files and getattr(src, "_brand_hints_ignore", None) == ignore_token:
        return cached[1]
    src._brand_hints_ignore = ignore_token
    positioning = report.load_positioning(src, settings)
    ignore = {prices.raw_brand_key(b) for b in settings.options.get("brand_ignore") or []}
    try:
        hints = price_check.brand_hints(files, dynamics.competitor_brand_names(positioning), src.crosses,
                                        ignore=ignore)
    except Exception:  # noqa: BLE001
        hints = []
    src._brand_hints = (files, hints)
    return hints


def api_brand_aliases(body: dict) -> dict:
    """Синонимы брендов: «как написано у поставщика» -> бренд конкурента.

    add / add_many — считать как конкурента; remove — убрать синоним;
    ignore / ignore_many — «не тот бренд», больше не предлагать; unignore — вернуть в подсказки.
    """
    settings = Settings.load()
    rows = [list(r) for r in settings.brand_aliases if isinstance(r, (list, tuple)) and len(r) >= 2]
    pairs = list(body.get("add_many") or [])
    if body.get("add"):
        pairs.append(body["add"])
    added = 0
    for add in pairs:
        if not (isinstance(add, (list, tuple)) and len(add) >= 2 and clean(add[0]) and clean(add[1])):
            continue
        alias, canonical = clean(add[0]), clean(add[1])
        if prices.raw_brand_key(alias) == prices.raw_brand_key(canonical):
            if len(pairs) == 1:
                raise UserFacing("Это и так один бренд.")
            continue
        rows = [r for r in rows if prices.raw_brand_key(r[0]) != prices.raw_brand_key(alias)]
        rows.append([alias, canonical])
        added += 1
    remove = body.get("remove")
    if remove:
        rows = [r for r in rows if prices.raw_brand_key(r[0]) != prices.raw_brand_key(remove)]
    ignored = [clean(b) for b in settings.options.get("brand_ignore") or [] if clean(b)]
    drop = list(body.get("ignore_many") or [])
    if body.get("ignore"):
        drop.append(body["ignore"])
    for brand in drop:
        brand = clean(brand)
        if brand and prices.raw_brand_key(brand) not in {prices.raw_brand_key(b) for b in ignored}:
            ignored.append(brand)
    back = body.get("unignore")
    if back:
        ignored = [b for b in ignored if prices.raw_brand_key(b) != prices.raw_brand_key(back)]
    settings.data["brand_aliases"] = rows
    settings.options["brand_ignore"] = ignored
    settings.save()
    Settings.load()  # синонимы вступают в силу сразу
    # Синонимы меняют ключ разобранных прайсов, «не тот бренд» — только список подсказок.
    return {"brand_aliases": rows, "brand_ignore": ignored, "added": added, "reread": bool(pairs or remove)}


def resolve_folder(raw: str | None, settings: Settings) -> Path:
    text = clean(raw) if raw is not None else ""
    if raw is None:
        text = clean(settings.options.get("folder"))
    if not text:
        return ROOT
    path = Path(os.path.expandvars(os.path.expanduser(text.strip('"'))))
    return path if path.is_absolute() else (ROOT / path)


def ensure_layout(root: Path) -> None:
    """Данные по группам, общие кроссы и прайсы, папка программы, результат."""
    ensure_data_dirs(root)
    for name in ("Выгрузки 1С", "Кроссы", "Прайсы", "Настройки"):
        old = root / name
        if not old.is_dir():
            continue
        if any(p.name != ".gitkeep" for p in old.iterdir()):
            continue
        for p in old.iterdir():
            p.unlink()
        old.rmdir()


def _workspace_root(folder: Path) -> bool:
    return folder.resolve() == ROOT.resolve()


def _with_incoming_acp(src: Sources, ws: Workspace, brand: str, group: str, trust_local: bool = False) -> Sources:
    """Свой АЦПКД в папке группы важнее. Иначе берётся файл из «Входящие», где эта группа есть.

    Файлов во «Входящие» может быть несколько. Если группа попала в два, берётся более новый.
    """
    if not brand and not group:
        return src
    if trust_local and acp_covers(src.acp, brand, group):
        return src
    slices = [s for s in ws.acp_slices if acp_covers(s.items, brand, group)]
    slices.sort(key=lambda s: s.mtime, reverse=True)
    if not slices:
        return src
    import copy

    out = copy.copy(src)
    out.found = dict(src.found)
    out.warnings = list(src.warnings)
    out.acp = slices[0].items
    if slices[0].located is not None:
        out.found["acp"] = [slices[0].located]
    if len(slices) > 1:
        names = ", ".join(s.label for s in slices)
        out.warnings.append(f"АЦПКД {brand} / {group}: несколько файлов ({names}) — взят {slices[0].label}")
    return out


def _pack_covering(ws: Workspace, brand: str, group: str):
    """Папка группы, в которой лежит АЦПКД или свой ПСТ этой ТМ и группы."""
    for pack in ws.packs:
        for base in ("acp", "pst"):
            for g in report.groups(pack.sources, base):
                if (not brand or g["brand"] == brand) and (not group or g["group"] == group):
                    return pack
    return None


def _pick_pack(ws: Workspace, brand: str, group: str):
    if not ws.packs and not ws.company.pst:
        return None
    if brand or group:
        pack = _pack_covering(ws, brand, group)
        if pack:
            return pack
        if ws.company.pst:
            return None
    for pack in ws.packs:
        if pack.sources.pst or pack.sources.acp:
            return pack
    return ws.packs[0] if ws.packs else None


def _documents(src: Sources) -> list[dict]:
    docs = []
    for kind in EXPECTED:
        locs = src.found.get(kind) or []
        rows = None
        if locs and kind == "pst":
            rows = len(src.pst)
        elif locs and kind == "acp":
            rows = len(src.acp)
        elif locs and kind == "gz":
            rows = len(src.gz)
        elif locs and kind == "transit":
            rows = len(src.transit)
        elif locs:
            rows = len(locs[0].header.data)
        docs.append({"kind": kind, "title": KIND_TITLES[kind], "label": locs[0].label if locs else "",
                     "rows": rows, "hint": DOC_HINTS[kind]})
    return docs


def _positioning_pairs(src: Sources, settings: Settings) -> set[tuple[str, str]]:
    pos = report.load_positioning(src, settings)
    return {(report.norm_text(r.brand), report.norm_text(r.group)) for r in pos.rows if r.brand and r.group}


def _only_positioning(groups: list[dict], allowed: set[tuple[str, str]]) -> list[dict]:
    if not allowed:
        return groups
    return [g for g in groups if (report.norm_text(g["brand"]), report.norm_text(g["group"])) in allowed]


def _only_positioned_subgroups(groups: list[dict], src: Sources, settings: Settings) -> list[dict]:
    """В карточке группы остаются подгруппы, которые есть в позиционировании."""
    pos = report.load_positioning(src, settings)
    out = []
    for g in groups:
        allowed = {
            report.norm_text(row.subgroup)
            for row in pos.rows
            if report.norm_text(row.brand) == report.norm_text(g["brand"])
            and report.norm_text(row.group) == report.norm_text(g["group"])
            and row.subgroup
        }
        if not allowed:
            out.append(g)
            continue
        subs = [s for s in g["subgroups"] if report.norm_text(s["name"]) in allowed]
        count = sum(s["count"] for s in subs)
        if count:
            out.append({**g, "count": count, "subgroups": subs})
    return out


def _is_group_pack(pack) -> bool:
    """Папка внутри «Данные/Группы», а не запасные «Входящие»."""
    root = data_paths(ROOT)["groups"].resolve()
    try:
        pack.folder.resolve().relative_to(root)
    except ValueError:
        return False
    return True


def _pack_pairs(pack) -> list[dict]:
    """ТМ и группы, которые лежат в этой папке (свой ПСТ или АЦПКД), без повторов."""
    seen = set()
    pairs = []
    for base in ("pst", "acp"):
        for g in report.groups(pack.sources, base):
            if not g["brand"] and not g["group"]:
                continue
            key = (g["brand"], g["group"])
            if key in seen:
                continue
            seen.add(key)
            pairs.append(g)
    return pairs


def _codes_hit(codes: set[str], *maps: dict) -> bool:
    return any(code in mapping for code in codes for mapping in maps)


def _batch_positioning(ws: Workspace, settings: Settings):
    if settings.data.get("positioning"):
        return report.load_positioning(Sources({}, []), settings)
    locs = []
    for pack in ws.packs:
        locs.extend(pack.sources.found.get("positioning") or [])
    if not locs:
        return report.load_positioning(Sources({}, []), settings)
    locs.sort(key=lambda loc: loc.header.sheet.file.stat().st_mtime, reverse=True)
    return report.load_positioning(Sources({"positioning": [locs[0]]}, []), settings)


def _pst_buckets(ws: Workspace) -> dict[tuple[str, str], dict]:
    pst = ws.company.pst if ws.company.pst else [p for pack in ws.packs for p in pack.sources.pst]
    buckets: dict[tuple[str, str], dict] = {}
    for item in pst:
        if not item.brand or not item.group:
            continue
        buckets.setdefault((item.brand, item.group), {})[item.code] = item
    return buckets


def _movement_maps(ws: Workspace) -> tuple[tuple, tuple]:
    if ws.company.found.get("sbs"):
        sbs_maps = (ws.company.sbs,)
    else:
        sbs_maps = tuple(pack.sources.sbs for pack in ws.packs)
    if ws.company.found.get("sales"):
        sales_maps = (ws.company.sales, ws.company.sales_qty)
    else:
        sales_maps = tuple(pack.sources.sales for pack in ws.packs) + tuple(
            pack.sources.sales_qty for pack in ws.packs
        )
    return sbs_maps, sales_maps


def _missing_phrase(parts: list[str]) -> str:
    if not parts:
        return ""
    if len(parts) == 1:
        return f"нет {parts[0]}"
    if len(parts) == 2:
        return f"нет {parts[0]} и {parts[1]}"
    return "нет " + ", ".join(parts[:-1]) + " и " + parts[-1]


REC_FIELDS = ("tolerance", "down", "min_vr", "top_factor", "mid_factor", "rounding", "max_lift", "cooldown")


def rec_group_key(brand: str, group: str) -> str:
    """Ключ порогов: бренд и группа после той же нормализации, что и при разборе файлов."""
    return f"{norm_text(brand)}\t{norm_text(group)}"


def _clean_rec_groups(raw) -> dict:
    """Только числа порогов. Включено и «по кому цель» остаются общими."""
    out = {}
    if not isinstance(raw, dict):
        return out
    for key, value in raw.items():
        if not isinstance(key, str) or "\t" not in key or len(key) > 300 or not isinstance(value, dict):
            continue
        clean_one = {}
        for field in REC_FIELDS:
            if value.get(field) in (None, ""):
                continue
            try:
                clean_one[field] = float(value[field])
            except (TypeError, ValueError):
                continue
        up = value.get("up")
        if isinstance(up, dict):
            parsed = {}
            for seg, num_value in up.items():
                if num_value in (None, ""):
                    continue
                try:
                    parsed[str(seg)] = float(num_value)
                except (TypeError, ValueError):
                    continue
            if parsed:
                clean_one["up"] = parsed
        if clean_one:
            out[key] = clean_one
        if len(out) >= 500:
            break
    return out


def _remember_rec(settings: Settings, incoming: dict) -> None:
    """Общие пороги не затираются числами открытой группы. Запоминаются включение, цель и свои пороги."""
    if isinstance(incoming.get("custom"), dict):
        settings.options["custom"] = incoming["custom"]
    rec = incoming.get("rec")
    if isinstance(rec, dict):
        stored = dict(settings.options.get("rec") or {})
        if "enabled" in rec:
            stored["enabled"] = bool(rec["enabled"])
        if rec.get("target"):
            stored["target"] = rec["target"]
        settings.options["rec"] = stored
    if incoming.get("money_base") in ("sales", "potential"):
        settings.options["money_base"] = incoming["money_base"]
    if isinstance(incoming.get("stock_profiles"), dict):
        stored = {}
        for key in ("sale", "overstock"):
            if isinstance(incoming["stock_profiles"].get(key), dict):
                stored[key] = incoming["stock_profiles"][key]
        settings.options["stock_profiles"] = stored
    for key in ("ui", "brand", "group", "report_kind"):
        if key in incoming:
            settings.options[key] = incoming[key]
    if isinstance(incoming.get("screen"), dict):
        raw = incoming["screen"]
        screen = {"enabled": bool(raw.get("enabled", True)), "kits": bool(raw.get("kits", True))}
        try:
            screen["ratio"] = max(1.5, float(raw.get("ratio", 3)))
        except (TypeError, ValueError):
            screen["ratio"] = 3.0
        settings.options["screen"] = screen
    if isinstance(incoming.get("sheet_filters"), list):
        settings.options["sheet_filters"] = [k for k in incoming["sheet_filters"] if k in report.SHEET_FILTERS]
    scope = incoming.get("rec_scope")
    if scope in ("all", "group"):
        settings.options["rec_scope"] = scope
        if incoming.get("rec_source") in ("mine", "yours"):
            settings.options["rec_source"] = incoming["rec_source"]
        if isinstance(incoming.get("rec_group_source"), dict):
            settings.options["rec_group_source"] = {
                str(k): v for k, v in incoming["rec_group_source"].items() if v in ("mine", "yours")
            }
        if incoming.get("save_rec"):
            if isinstance(incoming.get("rec_groups"), dict):
                settings.options["rec_groups"] = _clean_rec_groups(incoming["rec_groups"])
            if isinstance(incoming.get("rec_mine"), dict):
                settings.options["rec_mine"] = _clean_rec_groups({"x\ty": incoming["rec_mine"]}).get("x\ty", {})
    elif isinstance(incoming.get("rec_groups"), dict):
        settings.options["rec_groups"] = _clean_rec_groups(incoming["rec_groups"])


def _merge_thresholds(base: dict, overlay: dict) -> dict:
    merged = {**base, **{k: v for k, v in overlay.items() if k != "up"}}
    up = {**(base.get("up") or {}), **(overlay.get("up") or {})}
    if up:
        merged["up"] = up
    return merged


def rec_for_group(settings: Settings, brand: str, group: str, incoming: dict | None = None) -> dict:
    """Пороги для сборки группы.

    Без rec_scope — как раньше: общие из настроек плюс свои числа группы.
    rec_scope=all — один набор на всех: рекомендованный или сохранённый «мой».
    rec_scope=group — у группы либо зашитый набор, либо её сохранённые числа.
    Включение и «по кому цель» всегда из запроса.
    """
    incoming = incoming or {}
    scope = incoming.get("rec_scope") or settings.options.get("rec_scope")
    if scope not in ("all", "group"):
        common = dict(settings.options.get("rec") or {})
        overlay = (settings.options.get("rec_groups") or {}).get(rec_group_key(brand, group)) or {}
        merged = _merge_thresholds(common, overlay)
    else:
        groups = incoming["rec_groups"] if isinstance(incoming.get("rec_groups"), dict) else (settings.options.get("rec_groups") or {})
        sources = incoming["rec_group_source"] if isinstance(incoming.get("rec_group_source"), dict) else (settings.options.get("rec_group_source") or {})
        mine = incoming["rec_mine"] if isinstance(incoming.get("rec_mine"), dict) else (settings.options.get("rec_mine") or {})
        bundled = bundled_for(brand, group)
        key = rec_group_key(brand, group)
        if scope == "all":
            source = incoming.get("rec_source") or settings.options.get("rec_source") or "yours"
            merged = bundled if source == "yours" else _merge_thresholds(bundled, mine)
        else:
            source = sources.get(key) or ("mine" if key in groups else "yours")
            merged = bundled if source == "yours" else _merge_thresholds(bundled, groups.get(key) or {})
    req = incoming.get("rec") if isinstance(incoming.get("rec"), dict) else {}
    if "enabled" in req:
        merged["enabled"] = bool(req["enabled"])
    if req.get("target"):
        merged["target"] = req["target"]
    return merged


def _group_rows(ws: Workspace, settings: Settings) -> list[dict]:
    """Все ТМ и группы из ПСТ, которые есть в позиционировании, с отметкой полноты документов."""
    pos = _batch_positioning(ws, settings)
    known = {(norm_text(row.brand), norm_text(row.group)) for row in pos.rows if row.brand and row.group}
    sbs_maps, sales_maps = _movement_maps(ws)
    overrides = settings.options.get("rec_groups") or {}
    found = []
    for (brand, group), by_code in _pst_buckets(ws).items():
        if (norm_text(brand), norm_text(group)) not in known:
            continue
        items = list(by_code.values())
        kept = [item for item in items
                if not report.skipped_order_group(item.order_group)
                and report._subgroup_in_positioning(pos, brand, group, item.subgroup)]
        if not kept:
            continue
        subs: Counter = Counter(item.subgroup for item in kept)
        has_sbs = _codes_hit(set(by_code), *sbs_maps)
        has_sales = _codes_hit(set(by_code), *sales_maps)
        own = any(_is_group_pack(pack) and acp_covers(pack.sources.acp, brand, group) for pack in ws.packs)
        covered = own or any(acp_covers(pack.sources.acp, brand, group) for pack in ws.packs) or any(
            acp_covers(s.items, brand, group) for s in ws.acp_slices
        )
        missing = []
        if not has_sbs:
            missing.append("СБС")
        if not has_sales:
            missing.append("Продаж")
        if not covered:
            missing.append("АЦПКД")
        key = rec_group_key(brand, group)
        found.append({
            "key": key,
            "brand": brand,
            "group": group,
            "count": len(kept),
            "subgroups": [{"name": name, "count": n} for name, n in subs.most_common()],
            "pst": True,
            "sbs": has_sbs,
            "sales": has_sales,
            "acp": covered,
            "own": own,
            "ready": not missing,
            "reason": _missing_phrase(missing),
            "rec_own": key in overrides,
        })
    return sorted(found, key=lambda g: (g["brand"], not g["ready"], g["group"]))


def complete_groups(ws: Workspace, settings: Settings) -> list[dict]:
    """ТМ и группы, где нашлись ПСТ, СБС, Продажи и АЦПКД — неважно, в какой папке.

    ПСТ и АЦПКД узнаются по бренду и группе в файле. СБС и Продажи — по коду товара.
    Достаточно одного товара группы в каждом документе. Нет строки в позиционировании — группы нет.
    """
    ready = [g for g in _group_rows(ws, settings) if g["ready"]]
    return sorted(
        ({"brand": g["brand"], "group": g["group"], "count": g["count"]} for g in ready),
        key=lambda g: (-g["count"], g["brand"], g["group"]),
    )


def _gap_token(settings: Settings) -> str:
    import hashlib
    payload = json.dumps(
        {"files": settings.price_files, "profiles": settings.price_profiles, "aliases": prices.alias_token()},
        sort_keys=True, ensure_ascii=False, default=str,
    )
    return hashlib.sha1(payload.encode()).hexdigest()


def _apply_gaps(queue: list[dict], rows: list) -> bool:
    by = {(row.get("brand"), row.get("group")): row for row in rows}
    if any((g["brand"], g["group"]) not in by for g in queue):
        return False
    for g in queue:
        hit = by[(g["brand"], g["group"])]
        g["priced_rows"] = hit.get("priced_rows") or 0
        g["price_gap"] = hit.get("price_gap") or ""
        g["bucket"] = hit.get("bucket") or report.result_bucket(g["brand"], g["group"], g["priced_rows"])
    return True


def _annotate_queue(folder: Path, settings: Settings, queue: list[dict], ws: Workspace | None) -> None:
    """Цена конкурента и причина, если её нет: кроссов нет или цен нет."""
    if not queue:
        return
    signature = getattr(ws, "_file_signature", None) if ws is not None else None
    token = _gap_token(settings)
    if signature is not None:
        cached = filecache.load_gaps(signature, token)
        if cached and _apply_gaps(queue, cached):
            return
    if ws is not None:
        src = STATE.for_group(folder, queue[0]["brand"], queue[0]["group"], sisters=False)
    else:
        src = STATE.sources(folder)
    files = _report_prices(src, settings)
    positioning = report.load_positioning(src, settings)
    acp = dict(src.acp)
    if ws is not None:
        if ws.company:
            acp.update(ws.company.acp)
        for pack in ws.packs:
            acp.update(pack.sources.acp)
        for sl in ws.acp_slices:
            acp.update(sl.items)
    host = src
    if acp is not src.acp:
        import copy
        host = copy.copy(src)
        host.acp = acp
    index = report.PriceIndex(
        host, report.Options(price_source="both"), files, report.monitoring_map(positioning),
        report.competitor_brand_keys(positioning),
    )
    stored = []
    for g in queue:
        group_src = STATE.for_group(folder, g["brand"], g["group"], sisters=False) if ws is not None else src
        priced, gap = report.group_price_gap(group_src, settings, index, positioning, g["brand"], g["group"])
        g["priced_rows"] = priced
        g["price_gap"] = gap
        g["bucket"] = report.result_bucket(g["brand"], g["group"], priced)
        stored.append({
            "brand": g["brand"], "group": g["group"], "priced_rows": priced,
            "price_gap": gap, "bucket": g["bucket"],
        })
    if signature is not None:
        filecache.save_gaps(signature, token, stored)


def group_queue(ws: Workspace, settings: Settings) -> list[dict]:
    """Очередь окна: и полные группы, и те, где документа не хватает."""
    return _group_rows(ws, settings)


def queue_from_groups(groups: list[dict], src: Sources, settings: Settings) -> list[dict]:
    """Очередь для одной папки без раскладки «Данные / Входящие»."""
    overrides = settings.options.get("rec_groups") or {}
    out = []
    has_sbs = bool(src.sbs)
    has_sales = bool(src.sales or src.sales_qty)
    for g in groups:
        has_acp = acp_covers(src.acp, g["brand"], g["group"]) if src.acp else False
        missing = []
        if not has_sbs:
            missing.append("СБС")
        if not has_sales:
            missing.append("Продаж")
        if not has_acp:
            missing.append("АЦПКД")
        key = rec_group_key(g["brand"], g["group"])
        out.append({
            **g,
            "key": key,
            "pst": True,
            "sbs": has_sbs,
            "sales": has_sales,
            "acp": has_acp,
            "own": False,
            "ready": not missing,
            "reason": _missing_phrase(missing),
            "rec_own": key in overrides,
        })
    return out


def batch_plan(ws: Workspace, settings: Settings | None = None) -> list[dict]:
    settings = settings or Settings.load()
    return [{"folder": "", "path": "", "groups": complete_groups(ws, settings)}]


def _group_list(ws: Workspace, base: str) -> list[dict]:
    if base == "pst" and ws.company.pst:
        rows = []
        legacy = data_paths(ROOT)["legacy"]
        for g in report.groups(ws.company, "pst"):
            pack = _pack_covering(ws, g["brand"], g["group"])
            rows.append({**g, "folder": pack.name if pack else "Входящие",
                         "path": str(pack.folder if pack else legacy)})
        return rows
    rows = []
    for pack in ws.packs:
        for g in report.groups(pack.sources, base):
            rows.append({**g, "folder": pack.name, "path": str(pack.folder)})
    return rows


def layout_info(folder: Path) -> list[dict]:
    if not _workspace_root(folder):
        return []
    paths = data_paths(folder)
    out = []
    for key, title, hint in (
        ("crosses", "Кроссы", "один файл на все группы"),
        ("prices", "Прайсы", "папка на каждого поставщика, внутри его прайсы"),
    ):
        path = paths[key]
        files = collect_files([path], True) if path.is_dir() else []
        out.append({"key": key, "title": title, "path": str(path), "files": len(files), "hint": hint})
    ws = STATE.workspace(folder)
    for pack in ws.packs:
        files = collect_files([pack.folder], True)
        empty = not (pack.sources.pst or pack.sources.acp)
        out.append({"key": "group", "title": pack.name, "path": str(pack.folder), "files": len(files),
                    "hint": "АЦПКД этой группы. Общие ПСТ, СБС и продажи можно положить во «Входящие»" if empty else "АЦПКД и свои выгрузки. Общие ПСТ, СБС и продажи из «Входящие» важнее",
                    "empty": empty})
    return out


def open_path(path: Path) -> None:
    if hasattr(os, "startfile"):
        os.startfile(path)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


# ---------------------------------------------------------------- обработчики API


def api_options(body: dict) -> dict:
    """Запоминает режим, общие фильтры и пороги рекомендаций по группам."""
    settings = Settings.load()
    _remember_rec(settings, body.get("options") or {})
    settings.save()
    return {"ok": True}


def api_init(_body: dict) -> dict:
    ensure_layout(ROOT)
    s = Settings.load()
    rec = RecOptions.from_dict(s.options.get("rec"))
    return {
        "options": s.options,
        "folder": str(resolve_folder(None, s)),
        "fields": {k: {"title": t, "own": o, "acp": a} for k, (t, o, a) in report.FIELDS.items()},
        "sources": report.SOURCE_CHOICES,
        "filters": {k: {"title": t, "choices": c, "default": d} for k, (t, c, d) in report.FILTERS.items()},
        "sheet_filter_choices": [{"key": k, "title": title} for k, (title, _) in report.SHEET_FILTERS.items()],
        "default_sheet_filters": list(report.DEFAULT_SHEET_FILTERS),
        "clean_preset": report.CLEAN_PRESET,
        "cross_modes": report.CROSS_MODES,
        "cross_mode_hints": report.CROSS_MODE_HINTS,
        "cross_splits": report.CROSS_SPLITS,
        "price_sources": report.PRICE_SOURCES,
        "price_picks": report.PRICE_PICKS,
        "price_fields": prices.FIELDS,
        "abc_values": report.ABC_VALUES,
        "report_kinds": report.REPORT_KINDS,
        "subgroup_modes": report.SUBGROUP_MODES,
        "targets": TARGETS,
        "segments": [*SEGMENTS, UNKNOWN],
        "default_up": DEFAULT_UP,
        "rec": dataclasses_rec(rec),
        "bundled_rec": bundled_rec(),
        "bundled_groups": BUNDLED_GROUPS,
        "can_pick_folder": _has_tk(),
        "build": BUILD,
        "build_label": BUILD_LABEL,
    }


def dataclasses_rec(rec: RecOptions) -> dict:
    return {
        "enabled": rec.enabled, "target": rec.target, "tolerance": rec.tolerance, "down": rec.down,
        "up": rec.up, "min_vr": rec.min_vr, "max_lift": rec.max_lift, "cooldown": rec.cooldown,
        "top_factor": rec.top_factor, "mid_factor": rec.mid_factor, "rounding": rec.rounding,
        "priced_first": rec.priced_first,
    }


def _workspace_missing_rates(ws, settings: Settings) -> list[dict]:
    """Курсы, которых не хватает по ПСТ всех групп, а не только открытой."""
    items = []
    if ws.company:
        items.extend(ws.company.pst)
    for pack in ws.packs:
        items.extend(pack.sources.pst)
    return report.missing_rates_from_pst(items, settings)


def api_progress(_body: dict) -> dict:
    """Текущий шаг загрузки. Отвечает сразу, не ждёт чтения Excel."""
    return bootlog.snapshot()


RATE_GAP = 0.03  # курс в настройках расходится с файлом «валюта» больше чем на 3% — сказать
RATE_AGE_DAYS = 30


def rates_check(folder: Path, settings: Settings) -> list[dict]:
    """Курсы в настройках против файла «валюта» в прайсах и давность последней правки курсов.

    По курсу считаются закупка в рублях и СБС план. Старый курс — ВР «по плану» неверная.
    """
    notes: list[dict] = []
    ours = {clean(k).upper(): v for k, v in (settings.data.get("currency_rates") or {}).items() if v not in ("", None)}
    try:
        cur = dynamics.read_currency(folder)
    except Exception:  # noqa: BLE001 — проверка не должна мешать
        cur = {"rates": {}}
    for code, by_date in (cur.get("rates") or {}).items():
        mine = ours.get(code.upper())
        if not mine or not by_date:
            continue
        date = max(by_date, key=prices._shot_sort)
        value = by_date[date]
        ratio = value / mine
        if 50 < ratio < 200:  # в файле за 100 единиц (иена у ЦБ)
            value, ratio = value / 100, ratio / 100
        elif 0.005 < ratio < 0.02:
            value, ratio = value * 100, ratio * 100
        if abs(ratio - 1) > RATE_GAP:
            notes.append({"currency": code, "ours": mine, "file": round(value, 4), "date": date, "gap": ratio - 1,
                          "text": f"Курс {code} в настройках {mine:g}, в файле «валюта» на {date} — {value:.4g} "
                                  f"({(ratio - 1) * 100:+.0f}%). Закупка в рублях и СБС план считаются по старому курсу."})
    saved = settings.data.get("currency_rates_date")
    if saved:
        try:
            age = (dt.date.today() - dt.date.fromisoformat(saved)).days
        except ValueError:
            age = 0
        if age > RATE_AGE_DAYS and ours:
            notes.append({"currency": "", "age": age,
                          "text": f"Курсы валют в настройках не меняли {age} дн. (с {dt.date.fromisoformat(saved):%d.%m.%Y}). Проверьте, что они актуальны."})
    return notes


def api_scan(body: dict) -> dict:
    bootlog.begin("scan")
    try:
        result = _api_scan_body(body)
    except Exception as exc:
        bootlog.fail(str(exc))
        raise
    bootlog.finish()
    return result


def _api_scan_body(body: dict) -> dict:
    filecache.reset_stats()
    settings = Settings.load()
    folder = resolve_folder(body.get("folder"), settings)
    bootlog.step("Папка", str(folder))
    if not folder.exists():
        raise UserFacing(f"Папка не найдена: {folder}")
    sorted_prices = sort_prices(folder, settings) if _workspace_root(folder) else None
    if _workspace_root(folder):
        ensure_layout(folder)
        ws = STATE.workspace(folder)
        src = STATE.for_group(folder, settings.options.get("brand") or "", settings.options.get("group") or "")
        paths = data_paths(folder)
        files = collect_files([p for p in (paths["groups"], paths["crosses"], paths["prices"], paths["legacy"]) if p.is_dir()], True)
        groups = {"pst": _group_list(ws, "pst"), "acp": _group_list(ws, "acp")}
    else:
        src = STATE.sources(folder)
        paths, price_dirs = scan_paths(folder)
        recursive = bool(price_dirs) or any(p.is_dir() and p.name in ("Выгрузки 1С", "Кроссы", "Прайсы") for p in paths)
        files = collect_files(paths, recursive)
        groups = {"pst": report.groups(src, "pst"), "acp": report.groups(src, "acp")}
    allowed = _positioning_pairs(src, settings)
    groups = {key: _only_positioned_subgroups(_only_positioning(rows, allowed), src, settings) for key, rows in groups.items()}
    groups = {
        key: [{**g, "key": rec_group_key(g.get("brand") or "", g.get("group") or "")} for g in rows]
        for key, rows in groups.items()
    }
    if _workspace_root(folder):
        queue = group_queue(ws, settings)
        batch = batch_plan(ws, settings)
    else:
        queue = queue_from_groups(groups.get("pst") or [], src, settings)
        batch = []
    bootlog.step("Кроссы", "проверяю, у каких групп есть цена конкурента")
    _annotate_queue(folder, settings, queue, ws if _workspace_root(folder) else None)
    bootlog.step("Прайсы", "собираю распознанные файлы")
    payload = price_payload(src, settings)
    return {
        "folder": str(folder),
        "files": len(files),
        "dirs": layout_info(folder),
        "documents": _documents(src),
        "skipped": src.skipped,
        "warnings": src.warnings,
        "groups": groups,
        "queue": queue,
        "price_files": payload,
        "workspace": _workspace_root(folder),
        "batch": batch,
        "missing_rates": _workspace_missing_rates(ws, settings) if _workspace_root(folder) else report.missing_rates_from_pst(
            getattr(src, "pst", []) or [], settings),
        "reread": filecache.misses,
        "monitoring_suppliers": report.monitoring_suppliers(report.load_positioning(src, settings)),
        "brand_hints": brand_hints_payload(src, settings),
        "rates_check": rates_check(folder, settings),
        "price_archive": pricearchive.preview(folder) if _workspace_root(folder) else None,
        "price_sorted": sorted_prices,
        "brand_aliases": settings.brand_aliases,
        "brand_ignore": settings.options.get("brand_ignore") or [],
        "price_waiting": _waiting_prices(src),
    }


def sort_prices(folder: Path, settings: Settings) -> dict | None:
    """Новые файлы и архивы в «Прайсах» — по папкам «дата / поставщик». Ошибка раскладки не мешает запуску."""
    try:
        from types import SimpleNamespace

        known = report.known_suppliers(report.load_positioning(SimpleNamespace(located=lambda kind: None), settings))
    except Exception:  # noqa: BLE001
        known = []
    keep = set()
    for key in settings.price_files:
        parts = [p.strip() for p in str(key).split(" / ")]
        if len(parts) >= 2 and prices._normalize_date(parts[0]):
            keep.add(parts[1])  # «дата / файл / лист» — у файла в папке-дате есть ручные настройки
    try:
        bootlog.step("Раскладываю прайсы", "новые файлы и архивы — по папкам поставщиков")
        result = pricesort.sort_prices(folder, known, keep=keep)
    except Exception as exc:  # noqa: BLE001
        return {"moved": [], "unpacked": [], "unknown": [], "dup": 0, "errors": [{"file": "", "why": str(exc)}], "lines": []}
    result["lines"] = pricesort.summary_lines(result)
    return result


def _waiting_prices(src: Sources) -> list[dict]:
    """Прайсы прошлых дат, которые ещё не разобраны: для календаря в «Данных»."""
    owner = getattr(src, "_dormant_owner", None) or src
    out = []
    for path in getattr(owner, "dormant_prices", None) or []:
        out.append({"date": prices.price_date(path), "supplier": prices.supplier_folder(path) or path.stem,
                    "file": path.name})
    return out


def _options(body: dict) -> report.Options:
    opt = report.Options.from_dict(body.get("options") or {})
    opt.history_write = True  # окно программы ведёт историю цен; тесты и разбор файлов — нет
    return opt


def _sister_status(src, rows) -> dict:
    key = next((r.sister_brand for r in rows if r.sister_brand), "")
    return {
        "brand": sister_title(key) if key else "",
        "oem_rows": len(getattr(src, "oem_rows", None) or getattr(src, "oem", None) or []),
        "priced": sum(r.sister_price is not None for r in rows),
        "rows": sum(bool(r.sister_brand) for r in rows),
    }


def _report_prices(src: Sources, settings: Settings) -> list:
    """Отчёт и проценка смотрят только последнюю дату. Все даты остаются для динамики.

    Прошлые даты при старте не читаются. Если последней даты по сути нет — все её файлы
    выключены или не прочитались, — прошлые даты дочитываются и отчёт берёт предыдущую.
    """
    owner = getattr(src, "_dormant_owner", None)
    if owner is not None and _needs_old_dates(owner, settings) and _wake_owner(owner):
        args = getattr(src, "_group_args", None)
        if args:
            src = STATE.for_group(args[0], args[1], args[2], sisters=args[3])
    return prices.latest_price_files(STATE.price_files(src, settings))


def api_analyze(body: dict) -> dict:
    settings = Settings.load()
    folder = resolve_folder(body.get("folder"), settings)
    opt = _options(body)
    src = STATE.for_group(folder, opt.brand, opt.group) if _workspace_root(folder) else STATE.sources(folder)
    bootlog.step("Отчёт", f"{opt.brand} · {opt.group}".strip(" ·"))
    files = _report_prices(src, settings)
    prep = report.prepare(src, settings, opt, files)
    rows, final = prep.rows, prep.kept
    reasons = Counter(reason for _, reason in prep.dropped)
    abc = Counter(r.abc or "" for r in rows)
    subgroups = Counter(r.subgroup for r in final)
    comps = {}
    for r in final:
        for c in r.competitors:
            if c.key in r.picks:
                comps.setdefault(c.key, {"key": c.key, "name": c.name, "order": c.order})
    preview = [
        {
            "code": r.code, "article": r.article, "name": r.name, "subgroup": r.subgroup,
            "sales_qty": r.sales_qty, "sales": r.sales, "stock": r.stock, "in_transit": r.in_transit,
            "abc": r.abc, "sbs": r.sbs, "cat_a": r.cat_a,
            "fleet": r.car.fleet if r.car else "", "segment": r.car.segment if r.car else "",
            "makes": r.car.label if r.car else "",
            "target": r.target, "gap": r.gap, "weight": r.weight,
            "rec": r.rec.price if r.rec else None, "rec_change": r.rec.change if r.rec else None,
            "rec_kind": r.rec.kind if r.rec else "", "reason": r.rec.reason if r.rec else "",
            "crosses": {k: [{"a": p.article, "yes": p.yes} for p in v] for k, v in r.picks.items()},
            "prices": r.prices,
            "best": {k: o.supplier for k, o in r.best.items()},
        }
        for r in final[:15]
    ]
    by_source = getattr(prep.index, "by_source", {})
    any_index = _any_index(src, opt, files, prep.positioning)
    funnel = insight.funnel(src, final, prep.analysis, prep.index, any_index, files, prep.positioning)
    seen = set()
    positioning_rows = []
    for r in final:
        for c in r.competitors:
            key = (r.subgroup, c.key)
            if key in seen:
                continue
            seen.add(key)
            positioning_rows.append({"subgroup": r.subgroup, "name": c.name, "percent": r.percents.get(c.key, 0)})
    return {
        "total": len(rows),
        "kept": len(final),
        "priced": sum(r.any_price for r in final),
        "reasons": reasons.most_common(),
        "abc": dict(abc),
        "subgroups": [{"name": k, "count": v} for k, v in subgroups.items()],
        "crosses": prep.analysis,
        "cross_split": opt.cross_split,
        "compare": report.compare_sources(rows, src, opt, settings),
        "competitors": sorted(comps.values(), key=lambda c: (c["order"], c["name"])),
        "preview": preview,
        "no_positioning": sum(1 for r in rows if not r.competitors),
        "positioning": prep.positioning.source,
        "positioning_rows": positioning_rows,
        "has_svod": bool(src.crosses),
        "has_transit": bool(src.transit),
        "has_acp": bool(src.acp),
        "missing_rates": report.missing_rates(rows, settings),
        "sister": _sister_status(src, final),
        "offers": dict(by_source),
        "skipped_by_monitoring": prep.index.skipped_by_monitoring,
        "skipped_low": prep.index.skipped_low,
        "price_date": files[0].shot_date if files and files[0].shot_date else "",
        "sales_period": src.sales_period,
        "dashboard": dashboard_summary(final, opt),
        "funnel": funnel,
        "screened": insight.excluded_counts(final),
        "screen": insight.screen_options(settings),
        "history": insight.history_for(opt.brand, opt.group),
        "losses": insight.loss_reasons(prep, files),
        "signals": list(prep.signals) + currency_signals(folder, final),
        "oem": insight.oem_summary(final, src, opt),
        "sales_trend": _group_trend(src, final),
    }


def _group_trend(src, rows):
    try:
        from .salestrend import group_trend

        return group_trend(src, rows)
    except Exception:  # noqa: BLE001
        return None


RATE_MOVE = 0.03  # курс сдвинулся на 3% и больше — сигнал


def currency_signals(folder: Path, rows) -> list[dict]:
    """Курс закупочной валюты этой группы заметно сдвинулся за последние две недели (по файлу «валюта»)."""
    try:
        rates = (dynamics.read_currency(folder) or {}).get("rates") or {}
    except Exception:  # noqa: BLE001
        return []
    used: dict[str, int] = {}
    for r in rows:
        cur = str(getattr(r, "purchase_currency", "") or "").upper()
        if cur and cur not in ("RUB", "РУБ", "RUR"):
            used[cur] = used.get(cur, 0) + 1
    out = []
    for code, series in rates.items():
        items = sum(n for cur, n in used.items() if code.upper() in cur)
        if not items:
            continue
        points = []
        for day, value in (series or {}).items():
            try:
                points.append((dt.datetime.strptime(day, "%d.%m.%Y").date(), float(value)))
            except (TypeError, ValueError):
                continue
        points.sort()
        if len(points) < 2:
            continue
        last_day, last = points[-1]
        older = [p for p in points if p[0] <= last_day - dt.timedelta(days=14)] or \
            [p for p in points if p[0] <= last_day - dt.timedelta(days=7)]
        if not older or not older[-1][1]:
            continue
        base_day, base = older[-1]
        change = last / base - 1
        if abs(change) < RATE_MOVE:
            continue
        pct = f"{change * 100:+.1f}".replace(".", ",")
        side = "дорожает" if change > 0 else "дешевеет"
        out.append({"comp": code, "side": "up" if change > 0 else "down", "n": items, "of": len(rows), "median": change,
                    "since": base_day.strftime("%d.%m"), "kind": "rate",
                    "text": f"Курс {code} {pct}% с {base_day:%d.%m} — закупка {side} у {items} товаров",
                    "hint": "СБС план и ВР по ним сдвинутся со следующей закупкой"})
    return out


def _any_index(src: Sources, opt: report.Options, files, positioning):
    """Цены всех прайсов и АЦПКД — без отбора по листу мониторинга и наличию. Для воронки и дыр."""
    loose = report.Options(price_source=opt.price_source, monitoring_only=False, in_stock_only=False)
    return report.index_for(src, loose, files, report.monitoring_map(positioning), report.competitor_brand_keys(positioning))


def _remember_history(opt: report.Options, result) -> None:
    prep = getattr(result, "prep", None)
    if prep is None:
        return
    if opt.recs.enabled:
        try:
            from .pricehistory import remember_recs

            remember_recs(prep.kept)
        except Exception:  # noqa: BLE001 — история не должна мешать сборке отчёта
            pass
    try:
        insight.record_history(opt.brand, opt.group, prep.kept, dashboard_summary(prep.kept, opt), opt.stock_slice or "")
    except Exception:  # noqa: BLE001 — история не должна мешать сборке отчёта
        pass


def api_crosses_gaps(body: dict) -> dict:
    """Книга для базы кроссов по открытой группе: нет «Да», «Да» без цены, отсеянное."""
    settings = Settings.load()
    folder = resolve_folder(body.get("folder"), settings)
    opt = _options(body)
    src = STATE.for_group(folder, opt.brand, opt.group) if _workspace_root(folder) else STATE.sources(folder)
    files = _report_prices(src, settings)
    prep = report.prepare(src, settings, opt, files)
    if not prep.kept:
        raise UserFacing("После фильтров не осталось товаров — выгружать нечего.")
    any_index = _any_index(src, opt, files, prep.positioning)
    out = _report_dir(settings, folder)
    title = " ".join(x for x in (opt.brand, opt.group) if x) or "группа"
    path = out / f"{insight.GAPS_NAME} — {title}.xlsx"
    try:
        info = insight.gaps_workbook(path, src, prep.kept, prep.analysis, prep.index, any_index, prep.positioning, files)
    except PermissionError:
        raise UserFacing("Не удалось сохранить файл — возможно, он открыт в Excel. Закройте его и повторите.") from None
    from .workbook import _excel2016_stylesheet

    _excel2016_stylesheet(path)
    if body.get("open", True):
        _try_open(path)
    return {**info, "path": str(path), "name": path.name, "folder": str(path.parent)}


def _slice_payloads(raw: dict) -> list[dict]:
    """Пустой список из одного набора — обычный отчёт. both — два прохода."""
    value = raw.get("stock_slice") or ""
    names = ["sale", "overstock"] if value == "both" else [value] if value in ("sale", "overstock") else [""]
    out = []
    for name in names:
        item = dict(raw)
        if name:
            item["stock_slice"] = name
        else:
            item.pop("stock_slice", None)
        out.append(item)
    return out


def _api_build_stock(body: dict, raws: list[dict]) -> dict:
    """Один или два отчёта текущей группы: только распродажа, только оверсток или оба."""
    settings = Settings.load()
    folder = resolve_folder(body.get("folder"), settings)
    out_dir = data_paths(ROOT)["output"] if _workspace_root(folder) else resolve_folder(settings.options.get("output") or "Результат", settings)
    labels = {"sale": "Распродажа", "overstock": "Оверсток"}
    made, skipped = [], []
    for raw in raws:
        opt = report.Options.from_dict(raw)
        src = STATE.for_group(folder, opt.brand, opt.group) if _workspace_root(folder) else STATE.sources(folder)
        label = labels.get(opt.stock_slice, "")
        try:
            result = report.build_report(src, settings, opt, out_dir, _report_prices(src, settings))
        except PermissionError:
            skipped.append({"brand": opt.brand, "group": opt.group, "slice": label,
                            "reason": "файл открыт в Excel — закройте его и повторите"})
            continue
        except ValueError as e:
            skipped.append({"brand": opt.brand, "group": opt.group, "slice": label, "reason": str(e)})
            continue
        _remember_history(opt, result)
        made.append({
            "path": str(result.path), "name": result.path.name, "folder": str(result.path.parent),
            "brand": opt.brand, "group": opt.group, "slice": label, "rows": result.rows,
            "priced": result.priced, "brief": str(result.brief),
        })
    root = out_dir / "Распродажа и оверсток"
    if made and body.get("open", True):
        _try_open(root if len(raws) > 1 else Path(made[0]["folder"]))
    return {"made": made, "skipped": skipped, "folder": str(root), "stock": True}


def api_build(body: dict) -> dict:
    raws = _slice_payloads(body.get("options") or {})
    if len(raws) > 1 or raws[0].get("stock_slice"):
        return _api_build_stock(body, raws)
    settings = Settings.load()
    folder = resolve_folder(body.get("folder"), settings)
    opt = _options(body)
    src = STATE.for_group(folder, opt.brand, opt.group) if _workspace_root(folder) else STATE.sources(folder)
    out_dir = data_paths(ROOT)["output"] if _workspace_root(folder) else resolve_folder(settings.options.get("output") or "Результат", settings)
    try:
        result = report.build_report(src, settings, opt, out_dir, _report_prices(src, settings))
    except PermissionError:
        raise UserFacing("Не удалось сохранить файл — возможно, он открыт в Excel. Закройте его и повторите.") from None
    except ValueError as e:
        raise UserFacing(str(e)) from None
    _remember_history(opt, result)
    incoming = body.get("options") or {}
    settings.options.pop("price_agg", None)
    settings.options.update(
        folder="" if folder == ROOT else str(folder),
        base=opt.base, sources=opt.sources, filters=opt.filters,
        subgroup_mode=opt.subgroup_mode, abc=opt.abc, cross_modes=opt.cross_modes, cross_split=opt.cross_split,
        price_source=opt.price_source, monitoring_only=opt.monitoring_only, in_stock_only=opt.in_stock_only,
        brand=opt.brand, group=opt.group, report_kind=opt.report_kind,
        sheet_filters=report.sheet_filter_keys(opt),
        green_when=opt.green_when, reason_cols=opt.reason_cols, oem_cross=opt.oem_cross, demand=opt.demand, ui=incoming.get("ui") or settings.options.get("ui") or "template",
    )
    _remember_rec(settings, incoming)
    settings.save()
    if body.get("open", settings.options.get("open_result", True)):
        _try_open(result.path)
    return {"path": str(result.path), "name": result.path.name, "folder": str(result.path.parent),
            "rows": result.rows, "sheets": result.sheets, "excluded": result.excluded, "priced": result.priced,
            "kind": result.kind, "brief": str(result.brief), "brief_name": result.brief.name}


def _ensure_brief(out_dir, src, settings, opt) -> None:
    """Блок в «Разбор.txt», даже если Excel по этой группе не собрался."""
    from .brief import write_group_brief

    try:
        prep = report.prepare(src, settings, opt, _report_prices(src, settings))
        rows = prep.kept
    except Exception:
        rows = []
    write_group_brief(out_dir, rows, opt, src)


def api_build_all(body: dict) -> dict:
    """Excel — где есть ПСТ, СБС, Продажи и АЦПКД. Разбор.txt — по каждой группе из списка.

    Фильтры и кроссы — общие. Пороги рекомендаций — свои у группы, если их задавали, иначе общие.
    Включение рекомендаций и «по кому цель» одни на всю сборку. Прайс может быть пустым.
    Ошибка одной группы не обрывает разбор остальных.
    """
    settings = Settings.load()
    folder = resolve_folder(body.get("folder"), settings)
    if not _workspace_root(folder):
        raise UserFacing("Сразу все группы собираются из папки программы: «Данные» и «Входящие».")
    incoming = body.get("options") or {}
    _remember_rec(settings, incoming)
    settings.save()
    ws = STATE.workspace(folder)
    queue = group_queue(ws, settings)
    if not queue:
        raise UserFacing("В позиционировании нет групп из выгрузок.")
    complete = {(g["brand"], g["group"]) for g in complete_groups(ws, settings)}
    out_dir = data_paths(ROOT)["output"]
    made, skipped = [], []
    slice_raws = _slice_payloads(incoming)
    stock = any(item.get("stock_slice") for item in slice_raws)
    labels = {"sale": "Распродажа", "overstock": "Оверсток"}
    for base_raw in slice_raws:
        for g in queue:
            raw = dict(base_raw)
            raw.pop("custom", None)
            raw.pop("rec_groups", None)
            raw["brand"] = g["brand"]
            raw["group"] = g["group"]
            raw["rec"] = rec_for_group(settings, g["brand"], g["group"], incoming)
            opt = report.Options.from_dict(raw)
            src = STATE.for_group(folder, g["brand"], g["group"])
            slice_label = labels.get(opt.stock_slice, "")
            if (g["brand"], g["group"]) not in complete:
                if not stock:
                    _ensure_brief(out_dir, src, settings, opt)
                skipped.append({"folder": slice_label, "brand": g["brand"], "group": g["group"],
                                "reason": g.get("reason") or "нет полного набора документов"})
                continue
            try:
                result = report.build_report(src, settings, opt, out_dir, _report_prices(src, settings))
            except PermissionError:
                if not stock:
                    _ensure_brief(out_dir, src, settings, opt)
                skipped.append({"folder": slice_label, "brand": g["brand"], "group": g["group"],
                                "reason": "файл открыт в Excel — закройте его и повторите"})
                continue
            except Exception as e:  # noqa: BLE001 — одна группа не должна обрывать разбор остальных
                if not stock:
                    _ensure_brief(out_dir, src, settings, opt)
                skipped.append({"folder": slice_label, "brand": g["brand"], "group": g["group"], "reason": str(e)})
                continue
            _remember_history(opt, result)
            note = "" if result.priced else "цен конкурентов нет, файл всё равно создан"
            made.append({
                "path": str(result.path), "name": result.path.name, "folder_name": slice_label,
                "brand": g["brand"], "group": g["group"], "rows": result.rows, "priced": result.priced,
                "excluded": result.excluded, "kind": result.kind, "note": note,
            })
    open_dir = out_dir / "Распродажа и оверсток" if stock else out_dir
    if made and body.get("open"):
        _try_open(open_dir)
    brief = (open_dir / "Разбор.txt") if stock else (out_dir / "Разбор.txt")
    if stock and made:
        briefs = [open_dir / "Распродажа" / "Разбор.txt", open_dir / "Оверсток" / "Разбор.txt"]
        brief = next((path for path in briefs if path.is_file()), brief)
    missing = _workspace_missing_rates(ws, settings)
    return {"folder": str(open_dir), "made": made, "skipped": skipped,
            "brief": str(brief) if brief.is_file() else "", "brief_name": brief.name if brief.is_file() else "",
            "missing_rates": missing}


def _loaded_sources(body: dict, settings: Settings) -> Sources:
    folder = resolve_folder(body.get("folder"), settings)
    opt = body.get("options") or {}
    if _workspace_root(folder):
        return STATE.for_group(folder, opt.get("brand") or settings.options.get("brand") or "",
                               opt.get("group") or settings.options.get("group") or "")
    return STATE.sources(folder)


def api_prices_save(body: dict) -> dict:
    """Настройка одного прайса: поставщик, город, вкл/выкл, строка шапки и столбцы.

    dry_run — только показать, что получится (по первым строкам, без чтения всего файла);
    remember — запомнить формат для похожих файлов.
    """
    if body.get("dry_run"):
        return _prices_save(body)
    bootlog.begin("prices")
    try:
        bootlog.step("Сохраняю столбцы", str(body.get("id") or ""))
        result = _prices_save(body)
    except Exception as exc:
        bootlog.fail(str(exc))
        raise
    bootlog.finish()
    return result


def _prices_save(body: dict) -> dict:
    settings = Settings.load()
    src = _loaded_sources(body, settings)
    ident = body.get("id")
    sheet = next((s for s in src.price_sheets if prices.price_id(s) == ident or s.label == ident), None)
    if not sheet:
        # Файл прошлой даты: такие читаются, только когда открывают динамику.
        src, _ = _all_dates_sources(body, settings)
        sheet = next((s for s in src.price_sheets if prices.price_id(s) == ident or s.label == ident), None)
    if not sheet:
        raise UserFacing("Прайс не найден — пересканируйте папку.")
    positioning = report.load_positioning(src)
    key = prices.price_id(sheet)
    if body.get("reset"):
        settings.price_files.pop(key, None)
        if key != sheet.label:
            settings.price_files.pop(sheet.label, None)
        settings.save()
        bootlog.step("Обновляю список прайсов", "перечитываю только этот файл и файлы того же поставщика")
        return {"files": price_payload(src, Settings.load())}

    own = dict(prices.file_settings(settings.price_files, sheet))
    for field in ("supplier", "city", "fixed_brand"):
        if field in body:
            own[field] = clean(body[field])
    if "enabled" in body:
        own["enabled"] = bool(body["enabled"])
    if "header_row" in body:
        row = body["header_row"]
        if row in (None, ""):
            own.pop("header_row", None)
            own.pop("cols", None)
        else:
            own["header_row"] = int(row) - 1
            own["cols"] = {k: int(v) for k, v in (body.get("cols") or {}).items() if v not in (None, "")}
        own.pop("columns", None)
    brands = dynamics.competitor_brand_names(positioning)
    cross_brands, cross_pairs = cross_keys(src)
    pf = prices.load_one(sheet, own, settings.price_profiles, report.known_suppliers(positioning), brands,
                         known_brands=cross_brands, cross_pairs=cross_pairs, preview=True)
    if body.get("dry_run"):
        return {"file": describe_price(pf, positioning)}

    if "header_row" not in own:
        own.pop("cols", None)
    if key != sheet.label:
        settings.price_files.pop(sheet.label, None)
    settings.price_files[key] = own
    remembered = ""
    if body.get("remember"):
        if prices.can_remember(pf):
            settings.price_profiles[prices.profile_signature(pf)] = prices.make_profile(pf)
            remembered = "Формат запомнен — он применится к похожим прайсам."
        else:
            remembered = "Формат не запомнен: у столбцов артикула и цены нет названий в строке шапки."
    elif prices.can_remember(pf) and own.get("cols"):
        remembered = "Столбцы этого поставщика применятся к его следующим файлам, в том числе в других датах."
    settings.save()
    bootlog.step("Обновляю список прайсов", "перечитываю только этот файл и файлы того же поставщика")
    return {"files": price_payload(src, Settings.load()), "message": remembered}


def api_prices_layout(body: dict) -> dict:
    """Пишет «Устройство прайсов.txt»: шапка, пустые столбцы слева, какая колонка за что принята.

    Файл кладётся в папку результата. Его можно прислать целиком — по нему видна форма прайсов.
    """
    settings = Settings.load()
    folder = resolve_folder(body.get("folder"), settings)
    src, _ = _all_dates_sources(body, settings)
    files = STATE.price_files(src, settings)
    text = prices.layout_text(files, BUILD)
    out = _report_dir(settings, folder)
    out.mkdir(parents=True, exist_ok=True)
    path = out / prices.LAYOUT_NAME
    path.write_text(text, encoding="utf-8-sig")
    return {
        "path": str(path.resolve()),
        "name": path.name,
        "folder": str(out.resolve()),
        "files": len(files),
        "forms": text.count("\nФорма "),
    }


def api_prices_crosses(body: dict) -> dict:
    """Один файл: наши кроссы и цены поставщиков, которые на них встали."""
    settings = Settings.load()
    folder = resolve_folder(body.get("folder"), settings)
    src = _loaded_sources(body, settings)
    if not src.crosses:
        raise UserFacing("Нет файла кроссов. Положите «Кроссы свод» в Данные/Кроссы.")
    files = _report_prices(src, settings)
    positioning = report.load_positioning(src, settings)
    out = _report_dir(settings, folder)
    out.mkdir(parents=True, exist_ok=True)
    path = out / crossdump.CROSS_NAME
    info = crossdump.write_cross_prices(path, src.crosses, files, positioning)
    return {**info, "path": str(path.resolve()), "name": path.name, "folder": str(out.resolve())}


def api_prices_dynamics(body: dict) -> dict:
    """Как изменились цены конкурентов между двумя папками-датами и как за это время изменился курс."""
    settings = Settings.load()
    folder = resolve_folder(body.get("folder"), settings)
    src, woke = _all_dates_sources(body, settings)
    files = STATE.price_files(src, settings)
    positioning = report.load_positioning(src, settings)
    result = dynamics.build(src, files, positioning, folder, body.get("date_from"), body.get("date_to"))
    if woke:
        # Прошлые даты прочитались только сейчас — список на вкладке «Прайсы» дополняется ими.
        result["price_files"] = price_payload(src, settings)
    return result


_CATALOGS: list = []  # [((кроссы, ПСТ, позиционирование), каталог)]


def _all_pst(folder: Path, src: Sources) -> list:
    """ПСТ всех групп: категория, группа и подгруппа нашего артикула для аналитики по прайсам."""
    if not _workspace_root(folder):
        return list(getattr(src, "pst", None) or [])
    ws = STATE.workspace(folder)
    seen = set()
    out = []
    for items in [ws.company.pst or []] + [p.sources.pst or [] for p in ws.packs]:
        for item in items:
            key = (norm_text(item.brand), item.article)
            if key not in seen:
                seen.add(key)
                out.append(item)
    return out


def _price_scope(body: dict, settings: Settings):
    """Всё для аналитики и выгрузки цен: прайсы всех дат, бренды позиционирования, пары кроссов, каталог."""
    folder = resolve_folder(body.get("folder"), settings)
    bootlog.step("Читаю прайсы прошлых дат", "смотрю, что уже прочитано")
    src, woke = _all_dates_sources(body, settings)
    bootlog.step("Обновляю список прайсов", "уже разобранные файлы берутся из кэша")
    files = STATE.price_files(src, settings)
    # даты, сжатые в архив, — к динамике; дата, которая есть файлами, из архива не берётся
    files = files + pricearchive.archived_files(folder, {pf.shot_date for pf in files if pf.shot_date})
    positioning = report.load_positioning(src, settings)
    names = {prices.brand_key(n): n for n in dynamics.competitor_brand_names(positioning)}
    for key, label in dynamics._brand_names(files).items():
        if key in names and label:
            names[key] = label
    names.pop("", None)
    _brands, pairs = cross_keys(src)
    pst = _all_pst(folder, src)
    catalog = None
    held_key = (src.crosses, len(pst), id(positioning))
    for key, value in _CATALOGS:
        if key[0] is held_key[0] and key[1:] == held_key[1:]:
            catalog = value
    if catalog is None:
        from types import SimpleNamespace

        catalog = dynamics._catalog(SimpleNamespace(pst=pst, crosses=src.crosses), positioning)
        _CATALOGS[:] = [(held_key, catalog), *_CATALOGS][:2]
    panel = price_analytics.build_panel(files, set(names), pairs, names)
    return folder, src, woke, files, positioning, pairs, pst, catalog, panel


def _with_boot(plan: str, fn, body: dict) -> dict:
    bootlog.begin(plan)
    try:
        result = fn(body)
    except Exception as exc:
        bootlog.fail(str(exc))
        raise
    bootlog.finish()
    return result


def api_prices_analytics(body: dict) -> dict:
    return _with_boot("analytics", _prices_analytics, body)


def api_prices_export(body: dict) -> dict:
    return _with_boot("export", _prices_export, body)


def _prices_analytics(body: dict) -> dict:
    """Прайсы: что и где поменялось между двумя датами, индекс брендов по всем датам, курс."""
    settings = Settings.load()
    folder, src, woke, files, _pos, pairs, _pst, catalog, panel = _price_scope(body, settings)
    bootlog.step("Считаю изменения", f"{len(panel['cells'])} позиций · {len(panel['dates'])} дат")
    currency = dynamics.read_currency(folder)
    result = price_analytics.analyze(panel, catalog, currency, body.get("date_from"), body.get("date_to"))
    result["crossed"] = len(pairs)
    result["has_crosses"] = bool(src.crosses)
    result["currency_path"] = currency.get("path") or ""
    if woke:
        result["price_files"] = price_payload(src, settings)
    return result


def _prices_export(body: dict) -> dict:
    """Выгрузка цен для макроса: позиции из кроссов, все даты, файл и строка источника."""
    settings = Settings.load()
    folder, src, _woke, files, positioning, pairs, pst, _catalog, panel = _price_scope(body, settings)
    if not src.crosses:
        raise UserFacing("Нет файла кроссов. Положите «Кроссы свод» в Данные/Кроссы — выгрузка берёт только позиции из кроссов.")
    if not files:
        raise UserFacing("Прайсов нет.")
    currency = dynamics.read_currency(folder)
    out = _report_dir(settings, folder)
    out.mkdir(parents=True, exist_ok=True)
    path = out / price_analytics.EXPORT_NAME
    bootlog.step("Пишу Excel", path.name)
    info = price_analytics.write_export(path, files, panel, src.crosses, pst, positioning, currency, pairs)
    return {**info, "path": str(path.resolve()), "folder": str(out.resolve())}


def api_prices_dynamics_export(body: dict) -> dict:
    """Подробная книга динамики: выводы, бренды, срезы и каждый артикул."""
    settings = Settings.load()
    folder = resolve_folder(body.get("folder"), settings)
    src, _ = _all_dates_sources(body, settings)
    files = STATE.price_files(src, settings)
    if not files:
        raise UserFacing("Прайсов нет.")
    positioning = report.load_positioning(src, settings)
    out = _report_dir(settings, folder)
    out.mkdir(parents=True, exist_ok=True)
    path = out / dynamics.EXPORT_NAME
    info = dynamics.write_dynamics_export(
        path, src, files, positioning, folder, body.get("date_from"), body.get("date_to"),
    )
    return {**info, "path": str(path.resolve()), "folder": str(out.resolve())}


def api_prices_compact(body: dict) -> dict:
    """Сжать прайсы старше двух недель. dry — только посчитать, что сожмётся."""
    settings = Settings.load()
    folder = resolve_folder(body.get("folder"), settings)
    if body.get("dry", True):
        return pricearchive.preview(folder)
    return _with_boot("compact", _prices_compact, body)


def _prices_compact(body: dict) -> dict:
    settings = Settings.load()
    folder = resolve_folder(body.get("folder"), settings)
    bootlog.step("Читаю прайсы прошлых дат", "для архива нужны цены из старых файлов")
    src, _ = _all_dates_sources(body, settings)
    files = STATE.price_files(src, settings)
    positioning = report.load_positioning(src, settings)
    keep = {prices.brand_key(n) for n in dynamics.competitor_brand_names(positioning)} | set(cross_keys(src)[0])
    keep.discard("")
    bootlog.step("Сжимаю", "пишу архив и переношу исходники")
    result = pricearchive.compact(folder, files, keep)
    with STATE.lock:
        STATE.workspace_cache = None
        STATE.cache.clear()
        STATE.price_shared.clear()
    return result


def api_prices_forget(body: dict) -> dict:
    settings = Settings.load()
    settings.data["price_profiles"] = {}
    settings.save()
    src = _loaded_sources(body, settings)
    return {"files": price_payload(src, settings)}


def api_check(body: dict) -> dict:
    settings = Settings.load()
    folder = resolve_folder(body.get("folder"), settings)
    src = _loaded_sources(body, settings)
    if not src.pst:
        raise UserFacing("В папке нет ПСТ — проверять нечего.")
    out_dir = resolve_folder(settings.options.get("output") or "Результат", settings)
    out = out_dir / f"Проверка_папки_{dt.datetime.now():%Y-%m-%d_%H-%M-%S}.xlsx"
    check.build_report(src, out)
    _try_open(out)
    return {"path": str(out), "name": out.name, "folder": str(out.parent)}


def api_open(body: dict) -> dict:
    path = Path(body.get("path") or "")
    if not path.exists():
        raise UserFacing(f"Не найдено: {path}")
    open_path(path)
    return {"ok": True}


def _try_open(path: Path) -> None:
    try:
        open_path(path)
    except OSError:
        pass


def _has_tk() -> bool:
    try:
        import tkinter  # noqa: F401
    except ImportError:
        return False
    return os.name == "nt" or bool(os.environ.get("DISPLAY"))


def api_pick_folder(body: dict) -> dict:
    import tkinter
    from tkinter import filedialog

    root = tkinter.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    folder = filedialog.askdirectory(initialdir=body.get("folder") or str(ROOT), title="Папка с выгрузками")
    root.destroy()
    return {"folder": folder or ""}


def api_settings_get(_body: dict) -> dict:
    s = Settings.load()
    defaults = Settings.load(None)
    return {
        "currency_rates": s.data.get("currency_rates") or {},
        "gz": s.data.get("gz") or [],
        "discounts": s.data.get("discounts") or [],
        "suggest_threshold": s.options.get("suggest_threshold", 0.3),
        "open_result": s.options.get("open_result", True),
        "output": s.options.get("output", "Результат"),
        "customized": USER_FILE.exists(),
        "cars": s.cars,
        "cars_custom": bool(s.data.get("cars")),
        "fleets": sorted({str(r[3]) for r in cars_default() if len(r) > 3 and r[3]}),
        "segments": [*SEGMENTS, UNKNOWN],
        "positioning": s.data.get("positioning") or report.positioning_table(report.load_positioning(Sources({}, []))),
        "positioning_custom": bool(s.data.get("positioning")),
        "defaults_count": {"gz": len(defaults.data.get("gz") or []), "discounts": len(defaults.data.get("discounts") or [])},
        "city_codes": s.data.get("city_codes") or [],
    }


def api_settings_save(body: dict) -> dict:
    s = Settings.load()
    if "currency_rates" in body:
        rates = {}
        for k, v in body["currency_rates"].items():
            if clean(k):
                value = num(v)
                rates[clean(k).upper()] = value if value is not None else ""
        if rates != (s.data.get("currency_rates") or {}):
            s.data["currency_rates_date"] = dt.date.today().isoformat()
        s.data["currency_rates"] = rates
    if "gz" in body:
        s.data["gz"] = [[clean(n), float(c)] for n, c in body["gz"] if clean(n) and c not in ("", None)]
    if "discounts" in body:
        s.data["discounts"] = [
            [clean(x) for x in row[:5]] + [float(row[5])] for row in body["discounts"] if len(row) >= 6 and row[5] not in ("", None)
        ]
    if "cars" in body:
        rows = []
        for row in body["cars"]:
            cells = [clean(x) for x in (list(row) + [""] * 6)[:6]]
            if cells[0]:
                rows.append(cells)
        s.data["cars"] = rows
    if "positioning" in body:
        rows = []
        for row in body["positioning"]:
            cells = [clean(x) for x in (list(row) + [""] * 9)[:9]]
            if any(cells[:3]):
                rows.append(cells)
        s.data["positioning"] = rows
    if "city_codes" in body:
        s.data["city_codes"] = [[clean(a), clean(b)] for a, b, *_ in body["city_codes"] if clean(a) and clean(b)]
    for key in ("suggest_threshold", "open_result", "output"):
        if key in body:
            s.options[key] = body[key]
    s.save()
    Settings.load()  # коды городов вступают в силу сразу
    return api_settings_get({})


def api_settings_import(body: dict) -> dict:
    """Переносит ГЗ и скидки из файлов в папке в сохранённые настройки."""
    s = Settings.load()
    src = _loaded_sources(body, s)
    imported = []
    if src.gz and (loc := src.located("gz")):
        names = {}
        for row in loc.header.sheet.grid:
            filled = [v for v in row if clean(v)]
            if len(filled) == 2:
                names.setdefault(clean(filled[0]).upper(), clean(filled[0]))
        s.data["gz"] = [[names.get(k, k), v] for k, v in src.gz.items()]
        imported.append(f"ГЗ: {len(src.gz)}")
    if src.discounts and (loc := src.located("discounts")):
        rows, seen = [], set()
        for row in loc.header.data:
            value = num(loc.get(row, "discount"))
            if value is None:
                continue
            values = [clean(loc.get(row, k)) for k in ("category", "group", "subgroup", "order_group", "brand")]
            key = tuple(v.upper() for v in values[2:])
            if key not in seen:
                seen.add(key)
                rows.append(values + [value])
        s.data["discounts"] = rows
        imported.append(f"скидки: {len(rows)}")
    if not imported:
        raise UserFacing("В папке нет файлов ГЗ или «Скидка для клиентов».")
    s.save()
    return {**api_settings_get({}), "imported": imported}


def api_settings_reset(_body: dict) -> dict:
    s = Settings.load()
    defaults = Settings.load(None)
    for key in ("currency_rates", "gz", "discounts"):
        s.data[key] = defaults.data.get(key)
    s.data.pop("cars", None)
    s.data.pop("positioning", None)
    s.save()
    return api_settings_get({})


def _report_dir(settings: Settings, folder: Path | None = None) -> Path:
    """Куда ляжет следующий отчёт. Та же папка, что у «Сформировать отчёт»."""
    folder = folder if folder is not None else resolve_folder(None, settings)
    if _workspace_root(folder):
        return data_paths(ROOT)["output"]
    return resolve_folder(settings.options.get("output") or "Результат", settings)


def _folder_files(folder: Path) -> list[dict]:
    if not folder.is_dir():
        return []
    rows = []
    for path in folder.iterdir():
        if path.name.startswith(".") or not path.is_file():
            continue
        st = path.stat()
        rows.append({
            "name": path.name,
            "size": st.st_size,
            "mtime": dt.datetime.fromtimestamp(st.st_mtime).strftime("%d.%m.%Y %H:%M"),
        })
    rows.sort(key=lambda row: row["mtime"], reverse=True)
    return rows


def api_diagnose(body: dict) -> dict:
    """Что запущено и куда пишется результат. Без сборки отчёта."""
    settings = Settings.load()
    folder = resolve_folder(body.get("folder"), settings)
    out = _report_dir(settings, folder)
    brief = out / "Разбор.txt"
    brand = clean(body.get("brand") or settings.options.get("brand") or "")
    group = clean(body.get("group") or settings.options.get("group") or "")
    expect = column_expectation(brand, group)
    oem_rows = 0
    sister_prices = 0
    try:
        if _workspace_root(folder):
            src = STATE.for_group(folder, brand, group)
        else:
            src = STATE.sources(folder)
        oem_rows = len(getattr(src, "oem_rows", None) or getattr(src, "oem", None) or [])
        sister = sister_of(brand, group) if brand and group else ""
        book = getattr(src, "sister_prices", None) or {}
        sister_prices = sum(1 for key in book if key[0] == sister) if sister else 0
    except Exception as exc:  # noqa: BLE001
        expect = {**expect, "text": f"{expect['text']} Папку прочитать не удалось: {exc}"}
    return {
        "build": BUILD,
        "build_label": BUILD_LABEL,
        "root": str(ROOT.resolve()),
        "output": str(out.resolve()),
        "brief_path": str(brief.resolve()),
        "brief_exists": brief.is_file(),
        "brief_size": brief.stat().st_size if brief.is_file() else 0,
        "brief_mtime": dt.datetime.fromtimestamp(brief.stat().st_mtime).strftime("%d.%m.%Y %H:%M") if brief.is_file() else "",
        "files": _folder_files(out),
        "marks": program_marks(),
        "brand": brand,
        "group": group,
        "oem_rows": oem_rows,
        "sister_prices": sister_prices,
        "column": expect,
    }


def api_diagnose_write(body: dict) -> dict:
    """Кладёт «Диагностика.txt» в папку результата, чтобы её было видно в проводнике."""
    info = api_diagnose(body)
    out = Path(info["output"])
    out.mkdir(parents=True, exist_ok=True)
    path = out / "Диагностика.txt"
    lines = [
        f"Сборка {info['build']} ({info['build_label']})",
        f"Папка программы: {info['root']}",
        f"Папка результата: {info['output']}",
        "",
        "Разбор.txt появляется здесь же, когда нажимаете «Сформировать отчёт» или «Все группы».",
        f"Сейчас Разбор.txt: {'есть, ' + info['brief_mtime'] if info['brief_exists'] else 'ещё нет'}.",
        "",
        info["column"]["text"],
        f"Строк ОЕМ прочитано: {info['oem_rows']}. Цен смежной марки в ПСТ: {info['sister_prices']}.",
        "",
        "Проверки кода:",
    ]
    for mark in info["marks"]:
        lines.append(f"- {'да' if mark['ok'] else 'нет'}: {mark['name']}. {mark['detail']}")
    if info["files"]:
        lines += ["", "Файлы в папке результата:"]
        lines += [f"- {row['name']} ({row['mtime']})" for row in info["files"][:30]]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8-sig")
    return {**info, "probe": str(path.resolve()), "probe_name": path.name}


ROUTES = {
    "/api/init": api_init,
    "/api/options": api_options,
    "/api/scan": api_scan,
    "/api/progress": api_progress,
    "/api/analyze": api_analyze,
    "/api/build": api_build,
    "/api/build-all": api_build_all,
    "/api/check": api_check,
    "/api/diagnose": api_diagnose,
    "/api/diagnose/write": api_diagnose_write,
    "/api/prices/save": api_prices_save,
    "/api/prices/layout": api_prices_layout,
    "/api/prices/crosses": api_prices_crosses,
    "/api/crosses-gaps": api_crosses_gaps,
    "/api/brand-aliases": api_brand_aliases,
    "/api/prices/analytics": api_prices_analytics,
    "/api/prices/export": api_prices_export,
    "/api/prices/dynamics": api_prices_dynamics,
    "/api/prices/dynamics-export": api_prices_dynamics_export,
    "/api/prices/forget": api_prices_forget,
    "/api/prices/compact": api_prices_compact,
    "/api/open": api_open,
    "/api/pick-folder": api_pick_folder,
    "/api/settings": api_settings_get,
    "/api/settings/save": api_settings_save,
    "/api/settings/import": api_settings_import,
    "/api/settings/reset": api_settings_reset,
}


class UserFacing(Exception):
    """Ошибка, текст которой показываем пользователю как есть."""


class Handler(BaseHTTPRequestHandler):
    server_version = "PriceAnalysis/1.0"

    def log_message(self, fmt, *args):  # тише в консоли
        if os.environ.get("PRICE_DEBUG"):
            super().log_message(fmt, *args)

    def _send(self, status: int, body: bytes, ctype: str) -> None:
        self.send_response(status)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _json(self, status: int, data) -> None:
        self._send(status, json.dumps(data, ensure_ascii=False, default=str).encode("utf-8"), "application/json; charset=utf-8")

    def do_GET(self):  # noqa: N802
        path = self.path.split("?", 1)[0]
        if path in ROUTES:
            return self._dispatch(path, {})
        name = "index.html" if path in ("/", "") else path.lstrip("/")
        file = (WEB / name).resolve()
        if WEB.resolve() not in file.parents or not file.is_file():
            return self._send(HTTPStatus.NOT_FOUND, b"not found", "text/plain")
        self._send(HTTPStatus.OK, file.read_bytes(), CONTENT_TYPES.get(file.suffix, "application/octet-stream"))

    def do_POST(self):  # noqa: N802
        path = self.path.split("?", 1)[0]
        if path not in ROUTES:
            return self._json(HTTPStatus.NOT_FOUND, {"error": "нет такого действия"})
        length = int(self.headers.get("Content-Length") or 0)
        try:
            body = json.loads(self.rfile.read(length) or b"{}")
        except ValueError:
            return self._json(HTTPStatus.BAD_REQUEST, {"error": "неверный запрос"})
        self._dispatch(path, body)

    def _dispatch(self, path: str, body: dict) -> None:
        try:
            self._json(HTTPStatus.OK, ROUTES[path](body))
        except UserFacing as e:
            self._json(HTTPStatus.UNPROCESSABLE_ENTITY, {"error": str(e)})
        except Exception as e:  # noqa: BLE001
            traceback.print_exc()
            self._json(HTTPStatus.INTERNAL_SERVER_ERROR, {"error": f"Ошибка: {e}"})


def serve(host: str, port: int) -> ThreadingHTTPServer:
    for candidate in range(port, port + 20):
        try:
            return ThreadingHTTPServer((host, candidate), Handler)
        except OSError:
            continue
    raise OSError(f"Не удалось занять порт {port}–{port + 19}")
