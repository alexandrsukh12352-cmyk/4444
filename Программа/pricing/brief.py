"""Текстовый разбор в папку «Результат».

Пишется, когда софт собирает отчёт: один блок на торговую марку и группу.
Повторная сборка той же группы заменяет её блок, остальные остаются.
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path

from .dashboard import summary
from .report import Options, Row

NAME = "Разбор.txt"
_SECTION_TITLES = ("Целевые группы", "Мелкие группы", "Нет цен")
_SECTION_TAG = {"целевые": "Целевые группы", "мелкие": "Мелкие группы", "нетцен": "Нет цен"}


def _heading(brand: str, group: str) -> str:
    """«Advics Тормозные колодки», без рамок и без точки посередине."""
    def word(text: str) -> str:
        return text[:1].upper() + text[1:].lower() if text else text

    brand_text = " ".join(word(part) for part in (brand or "").split())
    group_text = (group or "").strip()
    if group_text and not any(ch.islower() for ch in group_text):
        bits = []
        for part in group_text.split():
            core = part.strip(".,")
            if core.isalpha() and core.isupper() and len(core) <= 4:
                bits.append(part)
            else:
                bits.append(word(part))
        group_text = " ".join(bits)
    return " ".join(part for part in (brand_text, group_text) if part)


def _watch_all(brand: str) -> bool:
    """NITTO и SUN не сводим к одному конкуренту."""
    key = (brand or "").upper().replace(" ", "").replace(".", "")
    return key in {"NITTO", "SUN"}


def _num_pct(fraction: float) -> str:
    points = abs(fraction) * 100
    rounded = round(points, 1)
    if abs(rounded - round(rounded)) < 0.05:
        return f"{int(round(rounded))}%"
    return f"{rounded:.1f}".replace(".", ",") + "%"


def _weighted(pairs: list[tuple[float, float]]) -> float | None:
    """(отклонение, продажи шт.). Нет продаж — простое среднее, чтобы группа не осталась без числа."""
    if not pairs:
        return None
    weight = sum(qty for _, qty in pairs)
    if weight > 0:
        return sum(gap * qty for gap, qty in pairs) / weight
    return sum(gap for gap, _ in pairs) / len(pairs)


def _article_gap(row: Row, price: float | None, brand: str) -> float | None:
    """Наша цена / цель − 1. Плюс — мы дороже цели позиционирования."""
    if not price:
        return None
    if _watch_all(brand):
        gaps = []
        for comp in sorted(row.competitors or [], key=lambda c: (c.order, c.name)):
            theirs = (row.prices or {}).get(comp.key)
            if not theirs:
                continue
            percent = (row.percents or {}).get(comp.key) or 0
            target = theirs * (1 + percent / 100)
            if target:
                gaps.append(price / target - 1)
        return sum(gaps) / len(gaps) if gaps else None
    if not row.target:
        return None
    return price / row.target - 1


def _pairs(rows: list[Row], brand: str, price_of) -> list[tuple[float, float]]:
    found = []
    for row in rows:
        gap = _article_gap(row, price_of(row), brand)
        if gap is None:
            continue
        found.append((gap, row.sales_qty or 0))
    return found


def _phrase(mean: float | None, *, direction: bool) -> str:
    if mean is None:
        return "нет"
    number = _num_pct(mean)
    if not direction or abs(mean) * 100 < 0.05:
        return number
    word = "мы дороже" if mean > 0 else "мы дешевле"
    return f"{number}, {word}"


def _verdict(mean: float | None, tolerance: float, plural: bool) -> str:
    if mean is None:
        return "Вывод: нет цен для анализа."
    if abs(mean) * 100 <= tolerance + 1e-9:
        return "Вывод: в позиционировании."
    who = "конкурентов" if plural else "конкурента"
    if mean > 0:
        return f"Вывод: мы дороже {who}."
    return f"Вывод: мы дешевле {who}."


def _gap_vs(row: Row, price: float | None, comp) -> float | None:
    """Наша цена / цель этого конкурента − 1. Плюс — мы дороже его цели."""
    if not price:
        return None
    theirs = (row.prices or {}).get(comp.key)
    if not theirs:
        return None
    percent = (row.percents or {}).get(comp.key) or 0
    target = theirs * (1 + percent / 100)
    if not target:
        return None
    return price / target - 1


def _ranked_competitors(rows: list[Row]) -> list:
    found = {}
    for row in rows:
        for comp in row.competitors or []:
            found.setdefault(comp.key, comp)
    return sorted(found.values(), key=lambda comp: (comp.order, comp.name))


def _named_gaps(rows: list[Row], price_of) -> list[str]:
    """Первый и второй конкурент по порядку позиционирования, если по ним есть цена."""
    bits = []
    for comp in _ranked_competitors(rows)[:2]:
        pairs = []
        for row in rows:
            gap = _gap_vs(row, price_of(row), comp)
            if gap is None:
                continue
            pairs.append((gap, row.sales_qty or 0))
        mean = _weighted(pairs)
        if mean is None:
            continue
        bits.append(f"{comp.name} {_phrase(mean, direction=True)}")
    return bits


def _category_line(rows: list[Row], brand: str, price_of) -> str:
    from .report import ABC_ORDER

    buckets: dict[str, list[tuple[float, float]]] = {}
    for row in rows:
        gap = _article_gap(row, price_of(row), brand)
        if gap is None:
            continue
        buckets.setdefault(row.abc or "", []).append((gap, row.sales_qty or 0))
    if not buckets:
        return ""
    bits = []
    for abc in sorted(buckets, key=lambda name: (ABC_ORDER.get(name, 99), name)):
        mean = _weighted(buckets[abc])
        if mean is None:
            continue
        bits.append(f"{abc or 'без ABC'} {_phrase(mean, direction=True)}")
    if not bits:
        return ""
    return "Категории. " + ". ".join(bits) + "."


def _action(rows: list[Row]) -> str:
    up = down = keep = 0
    for row in rows:
        kind = row.rec.kind if row.rec else ""
        if kind in ("up", "vr"):
            up += 1
        elif kind == "down":
            down += 1
        elif kind == "keep":
            keep += 1
    if not (up or down or keep):
        return ""
    return f" Поднять {up}, опустить {down}, не трогать {keep}."


def section(rows: list[Row], opt: Options, oem_rows: int | None = None) -> str:
    """До и после проценки: общее, первый и второй конкурент, категории ABC, что сделать."""
    del oem_rows
    if not rows:
        return "Вывод: нет цен для анализа.\n"
    brand = opt.brand or (rows[0].brand if rows else "")
    tolerance = opt.recs.tolerance
    before_price = lambda row: row.cat_a
    before = _pairs(rows, brand, before_price)
    overall = _weighted(before)
    if overall is None:
        return "Вывод: нет цен для анализа.\n"
    pieces = [f"До проценки. Общее отклонение {_phrase(overall, direction=True)}."]
    pieces.extend(bit + "." for bit in _named_gaps(rows, before_price))
    lines = [" ".join(pieces)]
    categories = _category_line(rows, brand, before_price)
    if categories:
        lines.append(categories)
    verdict_from = overall
    snap = summary(rows, opt)
    if opt.recs.enabled:
        after_price = lambda row: row.rec.price if row.rec and row.rec.price else None
        after_all = _weighted(_pairs(rows, brand, after_price))
        if after_all is not None:
            line = f"После проценки. Общее отклонение {_phrase(after_all, direction=True)}."
            if snap["vr"] is not None and snap["vr_rec"] is not None:
                line += f" ВР было {_num_pct(snap['vr'])}, стало {_num_pct(snap['vr_rec'])}."
            elif snap["vr"] is not None:
                line += f" ВР было {_num_pct(snap['vr'])}."
            lines.append(line)
            verdict_from = after_all
        elif snap["vr"] is not None:
            lines.append(f"ВР было {_num_pct(snap['vr'])}.")
    elif snap["vr"] is not None:
        lines[0] += f" ВР было {_num_pct(snap['vr'])}."
    verdict = _verdict(verdict_from, tolerance, _watch_all(brand))
    if opt.recs.enabled:
        verdict = verdict.rstrip() + _action(rows)
        if not verdict.endswith("."):
            verdict += "."
    lines.append(verdict)
    return "\n".join(lines) + "\n"


def _read_text(path: Path) -> str:
    data = path.read_bytes()
    for enc in ("utf-8-sig", "utf-8", "cp1251"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def _is_heading(line: str) -> bool:
    text = line.strip()
    if not text or text in _SECTION_TITLES:
        return False
    return not text.startswith((
        "До проценки", "После проценки", "Вывод:", "Ценовой", "Сформирован",
        "Отклонение", "В выборку", "ВР ", "Категории", "Раздел:",
    ))


def _split(text: str) -> tuple[str, dict[str, tuple[str, str]]]:
    """Блоки групп и раздел, под которым они стояли в файле."""
    head, blocks = [], {}
    current = None
    opened = "мелкие"
    active = "мелкие"
    buf: list[str] = []
    titles = {title: tag for tag, title in _SECTION_TAG.items()}

    def flush() -> None:
        nonlocal current, buf
        if current:
            blocks[current] = (opened, "\n".join(buf).strip() + "\n")
        current, buf = None, []

    for line in text.splitlines():
        if line.startswith("=== ") and line.endswith(" ==="):
            flush()
            raw = line[4:-4].strip()
            if " · " in raw:
                brand, group = raw.split(" · ", 1)
                current = _heading(brand, group)
            else:
                current = raw
            opened = active
            buf = []
            continue
        if line.strip() in titles:
            active = titles[line.strip()]
            continue
        if _is_heading(line):
            flush()
            current = line.strip()
            opened = active
            continue
        if current is None:
            if line.strip():
                head.append(line)
        else:
            buf.append(line)
    flush()
    return "\n".join(head).strip(), blocks


def _has_positioned_competitor(rows: list[Row]) -> bool:
    return any(row.competitors for row in rows)


def _section_tag(opt: Options, body: str) -> str:
    if "До проценки" not in body:
        return "нетцен"
    from .sisters import sister_of
    if sister_of(opt.brand, opt.group):
        return "целевые"
    return "мелкие"


def _visible(body: str) -> str:
    lines = [line for line in body.splitlines() if not line.startswith("Раздел:")]
    return "\n".join(lines).strip()


def write_group_brief(out_dir: Path, rows: list[Row], opt: Options, src=None) -> Path:
    """Кладёт или обновляет «Разбор.txt» в папке результата.

    Целевые группы, затем мелкие, в конце группы без цен.
    Группа без конкурента в позиционировании в файл не попадает.
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / NAME
    key = _heading(opt.brand, opt.group) or "Группа"
    include = not rows or _has_positioned_competitor(rows)
    try:
        body = section(rows, opt, None if src is None else len(getattr(src, "oem_rows", None) or [])) if include else ""
    except Exception as exc:  # noqa: BLE001 — файл в «Результат» должен появиться, даже если блок группы сломался
        body = f"Вывод: разбор не собрался ({exc}).\n"
        include = True
    if include and body:
        body = body if body.endswith("\n") else body + "\n"
    old = _read_text(path) if path.is_file() else ""
    _, blocks = _split(old)
    if include and body:
        blocks[key] = (_section_tag(opt, body), body)
    else:
        blocks.pop(key, None)
    stamp = dt.datetime.now().strftime("%d.%m.%Y %H:%M")
    grouped: dict[str, list[str]] = {tag: [] for tag in _SECTION_TAG}
    for name in sorted(blocks):
        tag, stored = blocks[name]
        grouped.setdefault(tag, []).append(f"{name}\n{_visible(stored)}\n")
    parts = [
        "Ценовой разбор ассортимента\n"
        f"Сформирован {stamp}.\n"
        "Отклонение от цели позиционирования, с весом продаж в штуках.\n"
    ]
    for tag, title in _SECTION_TAG.items():
        if not grouped[tag]:
            continue
        parts.append(title + "\n")
        parts.extend(grouped[tag])
    path.write_text("\n".join(parts), encoding="utf-8-sig")
    return path
