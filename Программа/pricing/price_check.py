"""Проверки прайсов: похожие на конкурентов бренды, которые не совпали, и «файл поехал» по сравнению
с прошлой датой того же поставщика.

Сводка по каждому прочитанному файлу (сколько цен, медиана цены по брендам-конкурентам)
копится в «Программа/кэш/price_stats.json». Поэтому сравнение с прошлой датой не требует
читать прайсы прошлых дат заново.
"""

from __future__ import annotations

import json
import re
import statistics
from collections import Counter, defaultdict
from difflib import SequenceMatcher

from . import prices
from .prices import PriceFile, brand_key, raw_brand_key, supplier_key

STATS_NAME = "price_stats.json"
STATS_VERSION = 1
MIN_BRAND_ROWS = 20  # сравниваем медиану бренда, только если цен у него хотя бы столько


# ---------------------------------------------------------------- похожие бренды


_WORDS = re.compile(r"[0-9A-ZА-ЯЁ]+")


def _words(text: str) -> list[str]:
    return _WORDS.findall(str(text).upper().replace("Ё", "Е"))


def _similar(name: str, key: str, targets: dict[str, str]) -> tuple[str, float] | None:
    """Похож ли бренд на конкурента. Совпадение — целым словом, а не куском слова.

    «Kashiyama Japan» → KASHIYAMA (слово). «G-Brake» → GBRAKE (слова слитно).
    «HERMANN» не похож на MANN: это кусок слова. Опечатки («BREMBBO») — по сходству букв.
    """
    words = _words(name)
    best = None
    for target in targets:
        score = 0.0
        if target in words and len(words) > 1:
            score = 0.95
        elif len(target) >= 4 and len(words) > 1 and any(
            "".join(words[i:j]) == target for i in range(len(words)) for j in range(i + 2, len(words) + 1)
        ):
            score = 0.95
        elif len(key) >= 5 and len(target) >= 5 and key[0] == target[0]:
            ratio = SequenceMatcher(None, key, target).ratio()
            if ratio >= 0.86:
                score = ratio
        if score and (best is None or score > best[1]):
            best = (target, score)
    return best


def brand_hints(files: list[PriceFile], competitor_names: list[str], crosses: dict | None = None,
                limit: int = 200, ignore: set | None = None) -> list[dict]:
    """Бренды из прайсов и кроссов, которые похожи на конкурента, но с ним не совпали.

    «Kashiyama Japan», «LYNX», «G-Brake Germany» — кандидаты в синонимы. Сам синоним человек
    подтверждает в окне: программа ничего не склеивает без спроса. Отклонённые («не тот бренд»)
    больше не предлагаются. Строки, не похожие на бренд («TOYOTA Camry 06->», «CITROEN C5,
    PEUGEOT 407»), — это не бренды, а модели машин из чужого столбца; их не предлагаем.
    """
    ignore = ignore or set()
    targets = {}
    for name in competitor_names:
        key = brand_key(name)
        if key:
            targets.setdefault(key, name)
    if not targets:
        return []
    seen: Counter = Counter()
    where: dict[str, set] = defaultdict(set)
    names: dict[str, str] = {}
    for pf in files:
        if not pf.enabled:
            continue
        local = pf.__dict__.get("_brand_counts")
        if local is None:
            local = pf.__dict__["_brand_counts"] = Counter(o.brand for o in pf.offers)
        for brand, n in local.items():
            raw = raw_brand_key(brand)
            if not raw:
                continue
            seen[raw] += n
            where[raw].add(pf.supplier)
            names.setdefault(raw, str(brand).strip())
    for groups in (crosses or {}).values():
        for rows in groups.values():
            for row in rows:
                raw = raw_brand_key(row.cross_brand)
                if raw:
                    seen[raw] += 1
                    where[raw].add("кроссы")
                    names.setdefault(raw, str(row.cross_brand).strip())
    out = []
    for raw, n in seen.most_common():
        if brand_key(names[raw]) in targets or raw in ignore:
            continue  # уже совпадает, в том числе через синоним, или человек сказал «не тот бренд»
        if not prices._brand_like(names[raw]):
            continue
        hit = _similar(names[raw], raw, targets)
        if not hit:
            continue
        out.append({
            "brand": names[raw], "rows": n, "where": sorted(where[raw]),
            "suggest": targets[hit[0]], "score": round(hit[1], 2),
        })
        if len(out) >= limit:
            break
    return out


# ---------------------------------------------------------------- «файл поехал»


def _stats_path():
    return prices.CACHE_DIR / STATS_NAME


def _load_stats() -> dict:
    try:
        data = json.loads(_stats_path().read_text(encoding="utf-8"))
        if data.get("version") == STATS_VERSION and isinstance(data.get("groups"), dict):
            return data
    except (OSError, ValueError, AttributeError):
        pass
    return {"version": STATS_VERSION, "groups": {}}


def _group_key(pf: PriceFile) -> str:
    return f"{supplier_key(pf.supplier)}|{(pf.city or '').lower()}"


def _summary(files: list[PriceFile], competitor_keys: set[str]) -> dict:
    offers = 0
    by_brand: dict[str, list[float]] = defaultdict(list)
    for pf in files:
        offers += len(pf.offers)
        for o in pf.offers:
            key = brand_key(o.brand)
            if key in competitor_keys and o.price:
                by_brand[key].append(o.price)
    return {
        "offers": offers,
        "brands": {k: [len(v), statistics.median(v)] for k, v in by_brand.items()},
    }


_SUMS: dict = {}  # (группа, дата, бренды) -> (файлы, сводка): та же группа файлов не пересчитывается


def _cached_summary(gkey: str, date: str, part: list[PriceFile], competitor_keys: set[str]) -> dict:
    token = (gkey, date, tuple(sorted(competitor_keys)))
    held = _SUMS.get(token)
    if held is not None and len(held[0]) == len(part) and all(a is b for a, b in zip(held[0], part)):
        return held[1]
    summary = _summary(part, competitor_keys)
    _SUMS[token] = (list(part), summary)
    return summary


INFO_PREFIXES = ("цен меньше", "пропал бренд")


def quality_split(files: list[PriceFile], competitor_names: list[str]) -> tuple[dict, dict]:
    """(замечания, справки). «Цен меньше» и «пропал бренд» — справка: так бывает, когда в этот день
    у поставщика и города меньше файлов. «Медиана цены ушла» — замечание: похоже, столбец не тот."""
    warn: dict[str, list[str]] = {}
    info: dict[str, list[str]] = {}
    for ident, notes in quality_notes(files, competitor_names).items():
        for text in notes:
            (info if text.startswith(INFO_PREFIXES) else warn).setdefault(ident, []).append(text)
    return warn, info


def quality_notes(files: list[PriceFile], competitor_names: list[str]) -> dict[str, list[str]]:
    """id файла -> замечания по сравнению с прошлой датой того же поставщика и города.

    Сильно меньше цен, пропал бренд-конкурент, медиана цены бренда сдвинулась больше чем
    на треть — частые признаки, что поставщик поменял формат и столбцы встали не те.
    """
    competitor_keys = {brand_key(n) for n in competitor_names if brand_key(n)}
    stats = _load_stats()
    groups = stats["groups"]
    current: dict[tuple[str, str], list[PriceFile]] = defaultdict(list)
    for pf in files:
        if pf.enabled and pf.shot_date and pf.offers:
            current[(_group_key(pf), pf.shot_date)].append(pf)
    changed = False
    for (gkey, date), part in current.items():
        summary = _cached_summary(gkey, date, part, competitor_keys)
        if groups.setdefault(gkey, {}).get(date) != summary:
            groups[gkey][date] = summary
            changed = True
    if changed:
        try:
            prices.CACHE_DIR.mkdir(parents=True, exist_ok=True)
            tmp = _stats_path().with_suffix(".tmp")
            tmp.write_text(json.dumps(stats, ensure_ascii=False), encoding="utf-8")
            tmp.replace(_stats_path())
        except OSError:
            pass
    notes: dict[str, list[str]] = {}
    for (gkey, date), part in current.items():
        history = groups.get(gkey, {})
        earlier = [d for d in history if prices._shot_sort(d) < prices._shot_sort(date)]
        if not earlier:
            continue
        prev_date = max(earlier, key=prices._shot_sort)
        prev, now = history[prev_date], history[date]
        found = []
        if prev["offers"] >= 100 and now["offers"] < prev["offers"] * 0.6:
            found.append(f"цен меньше, чем {prev_date}: {now['offers']:,} против {prev['offers']:,}".replace(",", " "))
        for key, (n_prev, med_prev) in prev["brands"].items():
            if n_prev < MIN_BRAND_ROWS:
                continue
            n_now, med_now = now["brands"].get(key, [0, None])
            if not n_now:
                found.append(f"пропал бренд {key} (был {n_prev} цен на {prev_date})")
            elif n_now >= MIN_BRAND_ROWS and med_prev and abs(med_now / med_prev - 1) > 0.35:
                found.append(f"медиана цены {key} {med_now:,.0f} против {med_prev:,.0f} на {prev_date} — проверьте столбец цены".replace(",", " "))
        if found:
            for pf in part:
                notes[pf.id] = found
    return notes
