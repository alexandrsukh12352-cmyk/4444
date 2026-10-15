"""Аналитика по ценам конкурентов: отсев несопоставимого, воронка покрытия, надёжность цены,
дыры в кроссах и история показателей группы.

Ничего здесь не меняет расчёт Кат. А, ВР и рекомендаций — только решает, какие пары
«наш артикул — цена конкурента» вообще сравнимы, и объясняет, почему у товара нет цены.
"""

from __future__ import annotations

import datetime as dt
import json
import re
import statistics
from collections import Counter
from pathlib import Path

from .prices import brand_key, supplier_key
from .sources import norm_article, norm_text

# ---------------------------------------------------------------- отсев несопоставимого

SCREEN_DEFAULTS = {"enabled": True, "kits": True, "ratio": 3.0}


def screen_options(settings) -> dict:
    raw = (getattr(settings, "options", None) or {}).get("screen") or {}
    out = dict(SCREEN_DEFAULTS)
    if isinstance(raw, dict):
        if "enabled" in raw:
            out["enabled"] = bool(raw["enabled"])
        if "kits" in raw:
            out["kits"] = bool(raw["kits"])
        try:
            ratio = float(raw.get("ratio", out["ratio"]))
            if ratio >= 1.5:
                out["ratio"] = ratio
        except (TypeError, ValueError):
            pass
    return out


def is_kit(row) -> bool:
    """Комплект «диски + колодки» и подобное: артикул через «+» или наименование «Комплект…»."""
    article = str(getattr(row, "article", "") or "")
    return "+" in article or norm_text(getattr(row, "name", "")).startswith("КОМПЛЕКТ")


def screen_pairs(rows, settings) -> Counter:
    """Убирает цену конкурента там, где сравнивать нечего. Что убрано и почему — в row.excluded.

    Комплект против одной детали и цена, отличающаяся от нашей в разы, — это ошибка кросса
    или комплектности, а не позиционирование. В расчёт отклонений и выводов они не идут.
    """
    cfg = screen_options(settings)
    reasons: Counter = Counter()
    ratio = cfg["ratio"]
    for r in rows:
        r.excluded = {}
        if not cfg["enabled"]:
            continue
        kit = cfg["kits"] and is_kit(r)
        for key, price in list(r.prices.items()):
            if not price:
                continue
            reason = kind = ""
            if kit:
                reason = kind = "комплект против одной детали"
            elif r.cat_a:
                q = r.cat_a / price
                if q > ratio:
                    kind = "мы дороже в разы"
                    reason = f"мы дороже в {q:.1f} раза — проверить кросс и комплектность"
                elif q * ratio < 1:
                    kind = "мы дешевле в разы"
                    reason = f"мы дешевле в {1 / q:.1f} раза — проверить кросс и комплектность"
            if reason:
                best = r.best.get(key)
                r.excluded[key] = {"reason": reason, "kind": kind, "price": price, "supplier": best.supplier if best else ""}
                r.prices[key] = None
                r.best.pop(key, None)
                reasons[kind] += 1
    return reasons


# ---------------------------------------------------------------- воронка и надёжность


def _items(n: int) -> str:
    """1 товар, 2 товара, 5 товаров."""
    if n % 10 == 1 and n % 100 != 11:
        return f"{n} товар"
    if 2 <= n % 10 <= 4 and not 12 <= n % 100 <= 14:
        return f"{n} товара"
    return f"{n} товаров"


def _of_items(n: int) -> str:
    """у 1 товара, у 5 товаров."""
    return f"{n} товара" if n % 10 == 1 and n % 100 != 11 else f"{n} товаров"


def _offers(row, comp_key: str) -> list:
    return [o for p in (row.picks.get(comp_key) or []) for o in p.offers]


def _yes_keys(src, row, comp_key: str) -> list[str]:
    from .report import acp_crosses, svod_comp, svod_for

    if src.crosses:
        return [c.cross_key for c in svod_comp(svod_for(src, row), comp_key) if c.yes]
    return [norm_article(c.article) for c in acp_crosses(row.acp, comp_key)]


def _acp_has(row, comp_key: str, keys: set[str]) -> bool:
    from .report import acp_crosses

    return any(norm_article(c.article) in keys and any(o.price for o in c.offers) for c in acp_crosses(row.acp, comp_key))


def funnel(src, rows, analysis: list[dict], index, any_index, price_files, positioning) -> list[dict]:
    """По каждому конкуренту: товаров → есть кросс «Да» → цена есть в каких-то прайсах →
    цена у поставщиков мониторинга → осталась после отсева. И что сделать, чтобы цен стало больше.

    any_index — цены всех прайсов без отбора по листу мониторинга и наличию.
    """
    from .report import monitoring_map, supplier_brand_book

    loaded = {supplier_key(pf.supplier) for pf in price_files or [] if pf.enabled and pf.offers}
    labels = supplier_brand_book(positioning)["labels"]
    monitoring = monitoring_map(positioning)
    out = []
    for comp in analysis:
        key = comp["key"]
        bk = brand_key(key)
        part = [r for r in rows if any(c.key == key for c in r.competitors)]
        if not part:
            continue
        sold_total = sum(r.sales_qty or 0 for r in part)
        yes = in_prices = monitored = final = excluded = by_no = 0
        sold_final = 0.0
        outside: Counter = Counter()
        for r in part:
            keys = _yes_keys(src, r, key)
            if keys:
                yes += 1
                ks = set(keys)
                anywhere = [o for k in ks for o in any_index.get(key, k)]
                if anywhere or _acp_has(r, key, ks):
                    in_prices += 1
                    if any(index.get(key, k) for k in ks):
                        monitored += 1
                    else:
                        for o in anywhere:
                            outside[o.supplier.split(" · ")[0]] += 1
            if key in getattr(r, "excluded", {}):
                excluded += 1
            price = r.prices.get(key)
            if price:
                final += 1
                sold_final += r.sales_qty or 0
                best = r.best.get(key)
                pick = next((p for p in r.picks.get(key) or [] if best in p.offers), None)
                if pick is not None and pick.yes is False:
                    by_no += 1
        wanted = monitoring.get(bk) or set()
        missing = sorted(labels.get(s, s) for s in wanted if s not in loaded)
        actions = []
        n = len(part)
        if not yes:
            actions.append(f"Нет кроссов «Да» на {comp['name']} — {_items(n)} не с чем сравнить. Нужны кроссы в базе.")
        else:
            no_price = yes - monitored
            if no_price and missing:
                actions.append(f"Нет прайса {', '.join(missing)} — у {_of_items(no_price)} кросс «Да» есть, цены нет.")
            if yes - in_prices > 0 and not missing:
                actions.append(f"У {_of_items(yes - in_prices)} кросс «Да» есть, но этих артикулов нет ни в одном прайсе.")
            if in_prices - monitored > 0:
                names = ", ".join(s for s, _ in outside.most_common(3))
                actions.append(f"У {_of_items(in_prices - monitored)} цена есть только у поставщиков не из листа 02 ({names}).")
            if n - yes > 0:
                actions.append(f"Без кросса «Да» ещё {_items(n - yes)}.")
        if excluded:
            actions.append(f"Отсеяно как несопоставимое: {excluded}.")
        out.append({
            "key": key, "name": comp["name"], "order": comp.get("order"),
            "items": n, "yes": yes, "in_prices": in_prices, "monitored": monitored,
            "excluded": excluded, "final": final, "by_no": by_no,
            "sold_total": sold_total, "sold_final": sold_final,
            "sold_share": sold_final / sold_total if sold_total else None,
            "missing_suppliers": missing, "actions": actions,
            "reliability": reliability(part, key),
        })
    return out


_CITY_PREFIX = re.compile(r"^[A-ZА-Я]{2,4}\.")
_TERM = re.compile(r"\s*\((?:срок|center|rzd|msk|spb)[^)]*\)", re.IGNORECASE)


def supplier_base(label: str) -> str:
    """«KRD.Форум-Авто(center) (срок дней - 2 )» и «TISS · Екатеринбург» -> «Форум-Авто», «TISS»."""
    text = str(label or "").split(" · ")[0]
    text = _TERM.sub("", text)
    text = _CITY_PREFIX.sub("", text.strip())
    text = re.sub(r"\(.*?\)", "", text)
    text = text.split(".")[0] if "." in text and not text.split(".")[0].isupper() else text
    return text.strip() or str(label or "")


def reliability(rows, comp_key: str) -> dict:
    """Насколько можно доверять минимальной цене: сколько поставщиков, есть ли наличие, кто даёт минимум
    и насколько дороже ближайший другой поставщик (большой разрыв — повод проверить минимум)."""
    priced = [r for r in rows if r.prices.get(comp_key)]
    single = no_stock = 0
    gaps = []
    leaders: Counter = Counter()
    for r in priced:
        best_by: dict[str, float] = {}
        for o in _offers(r, comp_key):
            base = supplier_key(supplier_base(o.supplier))
            if o.price and (base not in best_by or o.price < best_by[base]):
                best_by[base] = o.price
        best = r.best.get(comp_key)
        if best is not None:
            leaders[supplier_base(best.supplier)] += 1
            if best.qty is not None and best.qty <= 0:
                no_stock += 1
        ordered = sorted(best_by.values())
        if len(ordered) <= 1:
            single += 1
        elif ordered[0]:
            gaps.append(ordered[1] / ordered[0] - 1)
    return {
        "priced": len(priced),
        "single": single,
        "no_stock": no_stock,
        "second_gap": statistics.median(gaps) if gaps else None,
        "leaders": [{"supplier": s, "rows": n} for s, n in leaders.most_common(3)],
    }


def excluded_counts(rows) -> dict:
    out: Counter = Counter()
    for r in rows:
        for info in (getattr(r, "excluded", None) or {}).values():
            out[info.get("kind") or info["reason"]] += 1
    return dict(out)


# ---------------------------------------------------------------- дыры в кроссах

GAPS_NAME = "Дыры в кроссах"


def gaps_workbook(path: Path, src, rows, comps: list, index, any_index, positioning, price_files) -> dict:
    """Книга для работы с базой кроссов: где «Да» нет, где «Да» без цены, что отсеяно."""
    from . import style
    from .report import monitoring_map, supplier_brand_book, svod_comp, svod_for

    loaded = {supplier_key(pf.supplier) for pf in price_files or [] if pf.enabled and pf.offers}
    labels = supplier_brand_book(positioning)["labels"]
    monitoring = monitoring_map(positioning)
    no_yes, yes_no_price, odd = [], [], []
    for r in rows:
        svod = svod_for(src, r) if src.crosses else {}
        for comp in r.competitors:
            key = comp.key
            crosses = svod_comp(svod, key) if svod else []
            yes = [c for c in crosses if c.yes]
            base = [r.code, r.article, r.name, r.subgroup, r.abc or "", r.sales_qty or 0, comp.name]
            if not yes:
                no = [c for c in crosses if not c.yes]
                priced = [(o.price, c.cross_article) for c in no for o in index.get(key, c.cross_key)]
                best = min(priced) if priced else None
                no_yes.append(base + [
                    ", ".join(c.cross_article for c in no[:5]),
                    best[1] if best else "", best[0] if best else None,
                ])
                continue
            if r.prices.get(key) or key in (r.excluded or {}):
                continue
            wanted = monitoring.get(brand_key(key)) or set()
            missing = sorted(labels.get(s, s) for s in wanted if s not in loaded)
            for c in yes:
                anywhere = any_index.get(key, c.cross_key)
                if anywhere:
                    where = "цена есть только вне листа 02: " + ", ".join(sorted({o.supplier for o in anywhere})[:3])
                elif missing:
                    where = "нет прайса " + ", ".join(missing)
                else:
                    where = "артикула нет в прайсах"
                yes_no_price.append(base + [c.cross_article, where])
        for key, info in (r.excluded or {}).items():
            name = next((c.name for c in r.competitors if c.key == key), key)
            cross = ""
            best_pick = next((p.article for p in r.picks.get(key) or [] if p.offers), "")
            cross = best_pick or cross
            odd.append([r.code, r.article, r.name, r.subgroup, r.cat_a, name, cross, info["price"], info["supplier"], info["reason"]])
    no_yes.sort(key=lambda x: -(x[5] or 0))
    yes_no_price.sort(key=lambda x: -(x[5] or 0))
    wb = style.new_workbook()
    ws = wb.active
    ws.title = "Нет кросса Да"
    style.write_table(
        ws, "Товары без кросса «Да» — по продажам",
        ["Код", "Артикул", "Наименование", "Подгруппа", "ABC", "Продажи, шт", "Конкурент", "Кроссы «Нет»",
         "«Нет» с ценой", "Цена по «Нет»"],
        no_yes, widths=[10, 16, 40, 26, 6, 11, 14, 30, 16, 12],
    )
    ws2 = wb.create_sheet("Да без цены")
    style.write_table(
        ws2, "Кросс «Да» есть, цены нет — почему",
        ["Код", "Артикул", "Наименование", "Подгруппа", "ABC", "Продажи, шт", "Конкурент", "Кросс-артикул", "Причина"],
        yes_no_price, widths=[10, 16, 40, 26, 6, 11, 14, 16, 46],
    )
    ws3 = wb.create_sheet("Несопоставимо")
    style.write_table(
        ws3, "Цены, которые не взяты в сравнение",
        ["Код", "Артикул", "Наименование", "Подгруппа", "Кат. А", "Конкурент", "Кросс-артикул", "Цена конкурента",
         "Поставщик", "Причина"],
        odd, widths=[10, 16, 40, 26, 11, 14, 16, 12, 26, 46],
    )
    wb._colors = []
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return {"no_yes": len(no_yes), "yes_no_price": len(yes_no_price), "excluded": len(odd)}


# ---------------------------------------------------------------- история показателей

HISTORY_NAME = "история.json"


def _history_path() -> Path:
    from .settings import PROGRAM

    return PROGRAM / HISTORY_NAME


def _read_history(path: Path) -> list:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except (OSError, ValueError):
        return []


def record_history(brand: str, group: str, rows, summary: dict, slice_name: str = "", path: Path | None = None) -> None:
    """Запись по группе за сегодня: покрытие ценами, доля в коридоре, отклонение, ВР. Одна запись в день."""
    path = path or _history_path()
    main = next((s for s in summary.get("slices") or [] if s.get("kind") == "all"), {}) or {}
    sold = sum(r.sales_qty or 0 for r in rows)
    sold_priced = sum(r.sales_qty or 0 for r in rows if r.target is not None)
    record = {
        "date": dt.date.today().isoformat(),
        "brand": brand, "group": group, "slice": slice_name,
        "items": summary.get("items"), "priced": summary.get("priced"), "with_target": summary.get("with_target"),
        "in_band": summary.get("in_band"),
        "sold_share": sold_priced / sold if sold else None,
        "median": main.get("median"), "weighted": main.get("weighted"),
        "vr": summary.get("vr"),
    }
    items = [h for h in _read_history(path) if not (
        h.get("date") == record["date"] and h.get("brand") == brand and h.get("group") == group
        and h.get("slice", "") == slice_name)]
    items.append(record)
    items = items[-3000:]
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix(".tmp")
        tmp.write_text(json.dumps(items, ensure_ascii=False, indent=0), encoding="utf-8")
        tmp.replace(path)
    except OSError:
        pass


def history_for(brand: str, group: str, limit: int = 12, path: Path | None = None) -> list:
    path = path or _history_path()
    b, g = norm_text(brand), norm_text(group)
    found = [h for h in _read_history(path)
             if norm_text(h.get("brand", "")) == b and norm_text(h.get("group", "")) == g and not h.get("slice")]
    found.sort(key=lambda h: h.get("date", ""))
    return found[-limit:]
