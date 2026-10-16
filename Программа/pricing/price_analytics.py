"""Аналитика цен конкурентов по датам выгрузки и выгрузка цен для макроса.

Берутся только позиции из наших кроссов (бренд и артикул кросса) и только бренды
из позиционирования. Одна «ячейка» — поставщик + город + бренд + артикул: так видно,
где именно поменялась цена. Если в один день у поставщика и города два файла,
берётся меньшая цена.

Сравнение дат честное: изменение считается только по позициям, которые есть в обе даты
у того же поставщика и города. Пропавшие и новые позиции считаются отдельно.
Индекс бренда — цепной: от даты к дате медиана изменения по общим позициям, перемножается.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import math
import pickle
import random
import statistics
from collections import Counter, defaultdict
from pathlib import Path

from . import dynamics
from .excel import clean
from .filecache import CACHE_DIR
from .prices import PriceFile, brand_key, supplier_key
from .report import MIN_PRICE, supplier_brand_book
from .sources import norm_article, norm_text

PANEL_VERSION = 1
EXPORT_NAME = "Цены конкурентов по датам.xlsx"
SAME = 0.0005  # изменение меньше 0,05% — цена та же (копейки округления)
_ROWS_MEM: dict[str, dict] = {}
_PANELS: list = []  # [(ключ, панель)] — две последние
_TOKENS: list = []  # [(множество пар, отпечаток)]

BUCKETS = [
    (-math.inf, -0.10, "< −10%"), (-0.10, -0.05, "−10…−5%"), (-0.05, -0.02, "−5…−2%"),
    (-0.02, -SAME, "−2…0%"), (-SAME, SAME, "без изменений"), (SAME, 0.02, "0…+2%"),
    (0.02, 0.05, "+2…5%"), (0.05, 0.10, "+5…10%"), (0.10, math.inf, "> +10%"),
]


# ---------------------------------------------------------------- позиции из кроссов в одном файле


def scope_token(allowed: set, pairs: set) -> str:
    """Отпечаток набора: бренды и пары кроссов. Хэш строк в Python разный от запуска к запуску — берём sha1."""
    for held, token in _TOKENS:
        if held is pairs:
            return hashlib.sha1((token + "|" + ",".join(sorted(allowed))).encode()).hexdigest()
    h = hashlib.sha1()
    for bk, ak in sorted(pairs):
        h.update(f"{bk}\t{ak}\n".encode())
    token = h.hexdigest()
    _TOKENS[:] = [(pairs, token), *_TOKENS][:2]
    return hashlib.sha1((token + "|" + ",".join(sorted(allowed))).encode()).hexdigest()


def _rows_slot(token: str) -> Path:
    return CACHE_DIR / f"pan-{token}.pkl"


def crossed_rows(pf: PriceFile, allowed: set, pairs: set, scope: str) -> dict:
    """(бренд, артикул) -> (цена, наличие, строка файла, бренд как в прайсе, артикул как в прайсе).

    Только позиции из кроссов и бренды из позиционирования; из повторов — меньшая цена.
    """
    if not pf.enabled or not pf.offers:
        return {}
    token = hashlib.sha1(f"{PANEL_VERSION}|{dynamics._offer_token(pf)}|{scope}".encode()).hexdigest()
    hit = _ROWS_MEM.get(token)
    if hit is not None:
        return hit
    path = _rows_slot(token)
    if path.is_file():
        try:
            with path.open("rb") as fh:
                blob = pickle.load(fh)
            if isinstance(blob, dict) and blob.get("token") == token:
                _ROWS_MEM[token] = blob["rows"]
                return blob["rows"]
        except (OSError, pickle.PickleError, EOFError, ValueError, KeyError):
            pass
    keys: dict[str, str] = {}
    out: dict = {}
    for o in pf.offers:
        if not o.price or o.price < MIN_PRICE:
            continue
        bk = keys.get(o.brand)
        if bk is None:
            bk = keys[o.brand] = brand_key(o.brand)
        if bk not in allowed:
            continue
        key = (bk, o.article_key or norm_article(o.article))
        if key not in pairs:
            continue
        prev = out.get(key)
        if prev is None or o.price < prev[0]:
            out[key] = (o.price, o.qty, o.row, clean(o.brand), clean(o.article))
    _ROWS_MEM[token] = out
    try:
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix(".tmp")
        with tmp.open("wb") as fh:
            pickle.dump({"token": token, "rows": out}, fh, protocol=pickle.HIGHEST_PROTOCOL)
        tmp.replace(path)
    except (OSError, pickle.PickleError):
        pass
    return out


# ---------------------------------------------------------------- панель: ячейка -> цена по датам


def file_name(pf: PriceFile) -> str:
    """Имя файла прайса; у даты из архива — как файл назывался до сжатия."""
    return getattr(pf, "source_name", "") or pf.sheet.file.name


def _city(pf: PriceFile) -> str:
    return clean(pf.city) or "—"


def file_dates(files: list[PriceFile]) -> list[str]:
    dates = sorted({pf.shot_date or "без даты" for pf in files if pf.enabled and pf.offers}, key=dynamics._date_key)
    dated = [d for d in dates if d != "без даты"]
    return dated or dates


def build_panel(files: list[PriceFile], allowed: set, pairs: set, names: dict[str, str]) -> dict:
    scope = scope_token(allowed, pairs)
    dates = file_dates(files)
    key = (scope, tuple(sorted(dynamics._offer_token(pf) for pf in files if pf.enabled)))
    for held, panel in _PANELS:
        if held == key:
            return panel
    cells: dict[tuple, dict[str, float]] = defaultdict(dict)
    suppliers: dict[str, str] = {}
    shots = []
    for pf in files:
        date = pf.shot_date or "без даты"
        if date not in dates or not pf.enabled:
            continue
        sup = supplier_key(pf.supplier) or "—"
        suppliers.setdefault(sup, pf.supplier or "—")
        city = _city(pf)
        rows = crossed_rows(pf, allowed, pairs, scope)
        shots.append({"date": date, "supplier": sup, "city": city, "file": file_name(pf), "rows": len(rows),
                      "offers": len(pf.offers)})
        for (bk, ak), (price, *_rest) in rows.items():
            bucket = cells[(sup, city, bk, ak)]
            old = bucket.get(date)
            if old is None or price < old:
                bucket[date] = price
    panel = {"dates": dates, "cells": dict(cells), "suppliers": suppliers, "brands": dict(names),
             "shots": shots, "scope": scope}
    _PANELS[:] = [(key, panel), *_PANELS][:2]
    return panel


# ---------------------------------------------------------------- расчёты


def _median(values):
    return statistics.median(values) if values else None


def _stat(changes: list[float]) -> dict:
    up = sum(1 for c in changes if c > SAME)
    down = sum(1 for c in changes if c < -SAME)
    return {
        "n": len(changes), "up": up, "down": down, "same": len(changes) - up - down,
        "median": _median(changes), "mean": (sum(changes) / len(changes)) if changes else None,
        "moved": (up + down) / len(changes) if changes else 0.0,
    }


def _bucket(change: float) -> int:
    for i, (lo, hi, _label) in enumerate(BUCKETS):
        if lo <= change < hi:
            return i
    return len(BUCKETS) - 1


def _date_obj(text: str) -> dt.date | None:
    try:
        d, m, y = (int(x) for x in text.split("."))
        return dt.date(y, m, d)
    except (ValueError, AttributeError):
        return None


def presets(dates: list[str], date_to: str) -> dict[str, str]:
    """С чем сравнить последнюю дату: предыдущая выгрузка, неделя и месяц назад, первая дата."""
    before = [d for d in dates if dynamics._date_key(d) < dynamics._date_key(date_to)]
    out = {}
    if not before:
        return out
    out["prev"] = before[-1]
    end = _date_obj(date_to)
    if end:
        for name, days in (("week", 7), ("month", 30)):
            limit = end - dt.timedelta(days=days)
            fit = [d for d in before if (_date_obj(d) or end) <= limit]
            if fit:
                out[name] = fit[-1]
    out["first"] = before[0]
    return out


def _rates_on(currency: dict, dates: list[str]) -> dict[str, list]:
    """Курс на каждую дату прайсов. Нет курса в этот день — берётся последний известный до него."""
    out = {}
    for code, series in (currency.get("rates") or {}).items():
        known = sorted(((dynamics._date_key(d), v) for d, v in series.items()), key=lambda x: x[0])
        values = []
        for d in dates:
            k = dynamics._date_key(d)
            last = None
            for kk, v in known:
                if kk <= k:
                    last = v
                else:
                    break
            values.append(last)
        if any(v is not None for v in values):
            out[code] = values
    return out


def _pearson(xs: list[float], ys: list[float]) -> float | None:
    if len(xs) < 4:
        return None
    mx, my = sum(xs) / len(xs), sum(ys) / len(ys)
    vx = sum((x - mx) ** 2 for x in xs)
    vy = sum((y - my) ** 2 for y in ys)
    if vx <= 0 or vy <= 0:
        return None
    return sum((x - mx) * (y - my) for x, y in zip(xs, ys)) / math.sqrt(vx * vy)


def analyze(panel: dict, catalog: dict, currency: dict, date_from: str | None, date_to: str | None) -> dict:
    dates = panel["dates"]
    names = panel["brands"]
    suppliers = panel["suppliers"]
    if not dates:
        return {"dates": [], "empty": True}
    if date_to not in dates:
        date_to = dates[-1]
    pre = presets(dates, date_to)
    if date_from not in dates or dynamics._date_key(date_from or "") >= dynamics._date_key(date_to):
        date_from = pre.get("prev", "")
    cells = panel["cells"]
    label = lambda bk: names.get(bk, bk)  # noqa: E731

    # ------------------------------------------------ сравнение двух дат
    by_brand: dict[str, list] = defaultdict(list)
    by_place: dict[tuple, list] = defaultdict(list)
    by_where: dict[tuple, list] = defaultdict(list)
    by_feed: dict[tuple, list] = defaultdict(list)
    hist: dict[str, list[int]] = defaultdict(lambda: [0] * len(BUCKETS))
    moves: dict[str, list] = defaultdict(list)
    gone: Counter = Counter()
    new: Counter = Counter()
    both_places = set()
    if date_from:
        places_from = {(s, c) for (s, c, _b, _a), v in cells.items() if date_from in v}
        places_to = {(s, c) for (s, c, _b, _a), v in cells.items() if date_to in v}
        both_places = places_from & places_to
    for (sup, city, bk, ak), series in cells.items():
        if not date_from:
            break
        old, cur = series.get(date_from), series.get(date_to)
        if (sup, city) in both_places:
            if old and not cur:
                gone[bk] += 1
            elif cur and not old:
                new[bk] += 1
        if not old or not cur:
            continue
        change = cur / old - 1
        by_brand[bk].append(change)
        by_where[(sup, city, bk)].append(change)
        by_where[(sup, city, "")].append(change)
        by_feed[(bk, sup, city)].append(change)
        b = _bucket(change)
        hist[bk][b] += 1
        hist[""][b] += 1
        places = catalog.get((bk, ak)) or [("", "", "")]
        seen = set()
        for category, group, subgroup in places:
            cat = category or "Без категории"
            levels = [(cat, "", "")]
            if group:
                levels.append((cat, group, ""))
                if subgroup:
                    levels.append((cat, group, subgroup))
            for level in levels:
                if level not in seen:
                    seen.add(level)
                    by_place[(bk, *level)].append(change)
        if abs(change) > SAME:
            first = places[0]
            moves[bk].append({
                "brand": label(bk), "article": ak, "supplier": suppliers.get(sup, sup), "city": city,
                "old": old, "new": cur, "change": change,
                "category": first[0], "group": first[1], "subgroup": first[2],
            })
    all_changes = [c for v in by_brand.values() for c in v]
    brands = []
    for bk, changes in by_brand.items():
        st = _stat(changes)
        st.update({"key": bk, "brand": label(bk), "gone": gone.get(bk, 0), "new": new.get(bk, 0)})
        brands.append(st)
    brands.sort(key=lambda r: (-(r["moved"] * math.sqrt(r["n"])), r["brand"]))
    places = []
    for (bk, cat, group, sub), changes in by_place.items():
        st = _stat(changes)
        st.update({"key": bk, "brand": label(bk), "category": cat, "group": group, "subgroup": sub})
        places.append(st)
    places.sort(key=lambda r: (r["brand"], r["category"], r["group"], r["subgroup"]))
    where = []
    for (sup, city, bk), changes in by_where.items():
        st = _stat(changes)
        st.update({"supplier": suppliers.get(sup, sup), "supplier_key": sup, "city": city, "key": bk})
        where.append(st)
    feed = []
    for (bk, sup, city), changes in by_feed.items():
        st = _stat(changes)
        if st["n"] < 3 or st["moved"] < 0.1 or st["median"] is None:
            continue
        st.update({"key": bk, "brand": label(bk), "supplier": suppliers.get(sup, sup), "city": city})
        feed.append(st)
    feed.sort(key=lambda r: -(abs(r["mean"] or 0) * math.sqrt(r["n"]) + r["moved"] * 0.01))
    articles = {}
    for bk, rows in moves.items():
        rows.sort(key=lambda r: -abs(r["change"]))
        articles[bk] = rows[:200]

    # ------------------------------------------------ индекс по всем датам и курс
    steps: list[dict[str, float]] = []
    steps_n: list[dict[str, int]] = []
    for prev, cur in zip(dates, dates[1:]):
        ratios: dict[str, list] = defaultdict(list)
        for (_s, _c, bk, _a), series in cells.items():
            a, b = series.get(prev), series.get(cur)
            if a and b:
                ratios[bk].append(b / a)
                ratios[""].append(b / a)
        steps.append({bk: statistics.median(v) for bk, v in ratios.items()})
        steps_n.append({bk: len(v) for bk, v in ratios.items()})
    index: dict[str, list] = {}
    for bk in [""] + sorted({k for s in steps for k in s if k}, key=label):
        level, line = 100.0, [100.0]
        for st in steps:
            r = st.get(bk)
            if r is not None:
                level *= r
            line.append(round(level, 3))  # нет общих позиций на шаге — уровень тот же
        index[bk] = line
    rates = _rates_on(currency, dates)
    rate_index = {}
    for code, values in rates.items():
        base = next((v for v in values if v), None)
        rate_index[code] = [round(v / base * 100, 3) if v and base else None for v in values]
    corr = []
    for bk in sorted({k for s in steps for k in s if k}, key=label):
        xs = [math.log(s[bk]) if s.get(bk) else None for s in steps]
        for code, values in rates.items():
            ys = [math.log(values[i + 1] / values[i]) if values[i] and values[i + 1] else None for i in range(len(values) - 1)]
            best = None
            for lag in range(0, 5):
                pairs = [(xs[t], ys[t - lag]) for t in range(len(xs)) if t - lag >= 0 and xs[t] is not None and ys[t - lag] is not None]
                r = _pearson([p[0] for p in pairs], [p[1] for p in pairs])
                if r is not None and (best is None or abs(r) > abs(best[1])):
                    best = (lag, r, len(pairs))
            if best:
                corr.append({"key": bk, "brand": label(bk), "code": code, "lag": best[0], "r": round(best[1], 3), "n": best[2]})
    rate_change = {}
    if date_from:
        i, j = dates.index(date_from), dates.index(date_to)
        for code, values in rates.items():
            a, b = values[i], values[j]
            rate_change[code] = (b / a - 1) if a and b else None
    return {
        "empty": False,
        "dates": dates,
        "date_from": date_from,
        "date_to": date_to,
        "presets": pre,
        "summary": {**_stat(all_changes), "brands": len(by_brand),
                    "places": len({(s, c) for (s, c, b) in by_where if b == ""}),
                    "gone": sum(gone.values()), "new": sum(new.values())},
        "brands": brands,
        "places": places,
        "where": where,
        "feed": feed[:80],
        "hist": {"labels": [b[2] for b in BUCKETS], "values": {k: v for k, v in hist.items()}},
        "articles": articles,
        "index": {"dates": dates, "lines": {k: v for k, v in index.items()}, "n": steps_n,
                  "labels": {k: label(k) for k in index if k}},
        "rates": {"codes": sorted(rates), "values": rates, "index": rate_index, "change": rate_change},
        "corr": corr,
        "corr_note": "" if len(dates) >= 6 else f"Связь с курсом считается от 6 дат выгрузки, сейчас {len(dates)}.",
        "shots": _shots(panel),
    }


def _shots(panel: dict) -> list[dict]:
    by = defaultdict(lambda: {"files": 0, "rows": 0})
    for s in panel["shots"]:
        cell = by[(s["date"], s["supplier"])]
        cell["files"] += 1
        cell["rows"] += s["rows"]
    return [{"date": d, "supplier": panel["suppliers"].get(sup, sup), **v} for (d, sup), v in by.items()]


# ---------------------------------------------------------------- выгрузка для макроса


def _reverse_crosses(crosses: dict) -> dict[tuple, list]:
    rev: dict[tuple, list] = defaultdict(list)
    for (own_brand, own_key), by_brand in (crosses or {}).items():
        for rows in by_brand.values():
            for c in rows:
                rev[(brand_key(c.cross_brand), c.cross_key)].append((own_brand, own_key, c.yes))
    return rev


def write_export(path: Path, files: list[PriceFile], panel: dict, crosses: dict, pst: list, positioning,
                 currency: dict, pairs: set | None = None, sample: int = 50) -> dict:
    """Плоская книга: одна строка — цена позиции из кроссов у поставщика и города на дату.

    Если артикул конкурента стоит в кроссах у нескольких наших товаров, строка повторяется
    для каждого, а в «Первая строка цены» единица только у первой — так суммы не задваиваются.
    """
    from .xlsx_write import write_book

    dates = panel["dates"]
    scope = panel["scope"]
    names = panel["brands"]
    rev = _reverse_crosses(crosses)
    items = {}
    for item in pst or []:
        items.setdefault((norm_text(item.brand), norm_article(item.article)), item)
    book = supplier_brand_book(positioning)
    monitored = {sup: {brand_key(b) for b in brands} for sup, brands in book["brands"].items()}
    allowed = set(names)
    pairs = pairs if pairs is not None else set(rev)
    head = ["Дата", "Поставщик", "Город", "Файл", "Строка в файле", "Бренд в прайсе", "Бренд", "Артикул в прайсе",
            "Артикул (поиск)", "Цена", "Наличие", "Прошлая дата", "Цена на прошлую дату", "Изменение",
            "Наш ТМ", "Наш артикул", "Код товара", "Наименование", "Категория", "Группа", "Подгруппа",
            "Кросс «Да»", "Поставщик мониторит бренд", "Первая строка цены"]
    formats = {0: "date", 4: "int", 9: "money", 10: "int", 11: "date", 12: "money", 13: "pct", 23: "int"}
    widths = [11, 14, 14, 26, 8, 14, 12, 16, 16, 11, 9, 11, 11, 9, 10, 16, 10, 36, 22, 20, 24, 8, 10, 8]
    cells = panel["cells"]
    sample_pool: list = []
    meta = []

    def rows():
        order = sorted((pf for pf in files if pf.enabled and (pf.shot_date or "без даты") in dates),
                       key=lambda pf: (dynamics._date_key(pf.shot_date or "без даты"), pf.supplier, _city(pf), file_name(pf)))
        for pf in order:
            date = pf.shot_date or "без даты"
            day = _date_obj(date) or date
            sup = supplier_key(pf.supplier) or "—"
            city = _city(pf)
            found = crossed_rows(pf, allowed, pairs, scope) if pairs else {}
            meta.append([day, pf.supplier, city, file_name(pf), pf.sheet.name, pf.status, len(pf.offers), len(found),
                         "; ".join(pf.notes) or pf.problem or ""])
            di = dates.index(date)
            for (bk, ak), (price, qty, row, brand_raw, art_raw) in sorted(found.items(), key=lambda kv: kv[1][2]):
                series = cells.get((sup, city, bk, ak)) or {}
                prev_date = next((d for d in reversed(dates[:di]) if series.get(d)), None)
                prev = series.get(prev_date) if prev_date else None
                change = (price / prev - 1) if prev else None
                sample_pool.append((pf, row, brand_raw, art_raw, price))
                watch = "да" if bk in monitored.get(sup, ()) else "нет"
                links = rev.get((bk, ak)) or [("", "", None)]
                for n, (own_brand, own_key, yes) in enumerate(links):
                    item = items.get((own_brand, own_key))
                    yield [
                        day, pf.supplier, city, file_name(pf), row, brand_raw, names.get(bk, bk), art_raw, ak,
                        price, qty, (_date_obj(prev_date) or prev_date) if prev_date else None, prev, change,
                        (item.brand if item else own_brand), (item.article if item else own_key),
                        (clean(item.code) if item else ""), (clean(item.name) if item else ""),
                        (clean(item.category) if item else ""), (clean(item.group) if item else ""),
                        (clean(item.subgroup) if item else ""),
                        "" if yes is None else ("да" if yes else "нет"), watch, 1 if n == 0 else 0,
                    ]

    # Строки собираются заранее: по ходу заполняются список файлов и выборка для проверки.
    data_rows = list(rows())
    chunks = [data_rows[i:i + 1_000_000] for i in range(0, len(data_rows), 1_000_000)] or [[]]
    out_sheets = []
    for n, chunk in enumerate(chunks):
        out_sheets.append({"name": "Цены" if n == 0 else f"Цены {n + 1}", "header": head, "rows": chunk,
                           "formats": formats, "widths": widths})
    out_sheets.append({
        "name": "Даты и файлы",
        "header": ["Дата", "Поставщик", "Город", "Файл", "Лист", "Как распознан", "Цен в файле", "Позиций из кроссов", "Замечания"],
        "rows": meta, "formats": {0: "date", 6: "int", 7: "int"}, "widths": [11, 16, 14, 30, 14, 14, 11, 11, 60],
    })
    rate_rows = []
    for code, series in sorted((currency.get("rates") or {}).items()):
        for d, v in sorted(series.items(), key=lambda kv: dynamics._date_key(kv[0])):
            rate_rows.append([_date_obj(d) or d, code, v])
    out_sheets.append({"name": "Курсы", "header": ["Дата", "Валюта", "Курс"], "rows": rate_rows,
                       "formats": {0: "date"}, "widths": [11, 8, 10]})
    check = []
    rnd = random.Random(7)
    picked = rnd.sample(sample_pool, min(sample, len(sample_pool)))
    by_file: dict[int, list] = defaultdict(list)
    for item in picked:
        by_file[id(item[0])].append(item)
    for group in by_file.values():
        pf = group[0][0]
        try:
            grid = pf.sheet.full_grid()
        except Exception:  # noqa: BLE001 — проверка не должна ронять выгрузку
            grid = []
        for _pf, row, brand_raw, art_raw, price in sorted(group, key=lambda x: x[1]):
            line = grid[row - 1] if 0 < row <= len(grid) else []
            text = " | ".join(clean(v) for v in line if clean(v))
            check.append([pf.shot_date, pf.supplier, _city(pf), file_name(pf), row, brand_raw, art_raw, price, text])
    out_sheets.append({
        "name": "Проверка",
        "header": ["Дата", "Поставщик", "Город", "Файл", "Строка", "Бренд", "Артикул", "Цена", "Строка файла целиком"],
        "rows": check, "formats": {4: "int", 7: "money"}, "widths": [11, 16, 14, 30, 8, 14, 16, 11, 120],
    })
    counts = write_book(path, out_sheets)
    return {"rows": len(data_rows), "files": len(meta), "dates": len(dates), "checked": len(check),
            "sheets": len(out_sheets), "name": path.name, "counts": counts}
