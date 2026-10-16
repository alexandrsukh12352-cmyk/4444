"""Дашборд: главное о группе на одном листе — в полном и в быстром отчёте.

Карточки, корзины отклонений и ВР по ABC / подгруппам / сегментам — живые формулы по всем
листам отчёта: поправили Кат. А или рекомендованную цену — дашборд пересчитался.
Статистика отклонений (среднее, медиана, очистка от крайних и выбросов) и выводы — снимок
на момент формирования, считаются здесь же, в Python (их же показывает окно программы).
"""

from __future__ import annotations

import datetime as dt
import statistics
from dataclasses import dataclass

from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from . import style
from .cars import SEGMENTS, UNKNOWN
from .report import ABC_ORDER, PRICE_SOURCES, REPORT_KINDS, SUBGROUP_MODES, Options, Row, competitor_order
from .recommend import TARGETS, sales_weight

NUM = ">-1E+99"  # условие SUMIFS / COUNTIFS «ячейка — число»
CARD_FILL = PatternFill("solid", fgColor="F3F7F8")
CARD_LINE = Side(style="thin", color="D5E0E3")
HEAD_FILL = PatternFill("solid", fgColor=style.ACCENT)
SOFT_FILL = PatternFill("solid", fgColor="F7F9FA")
MONEY0 = '#,##0;-#,##0;"—"'
PCT1 = '0.0%;-0.0%;0.0%'
PCT1S = '+0.0%;-0.0%;0.0%'
PP = '+0.0" п.п.";-0.0" п.п.";0.0" п.п."'
SHARE = '0%;-0%;"—"'
GREENS = ["15803D", "22C55E", "4ADE80"]  # от сильного отклонения к слабому
REDS = ["B91C1C", "DC2626", "F87171"]


# ---------------------------------------------------------------- расчёт в Python (снимок)


def money_qty(r: Row, basis: str) -> float | None:
    """Штуки, на которые умножается цена. Потенциал — остаток плюс товар в пути."""
    if basis == "potential":
        if r.stock is None and r.in_transit is None:
            return None
        return (r.stock or 0) + (r.in_transit or 0)
    return r.sales_qty


def money(r: Row, price: float | None, basis: str = "sales") -> tuple[float | None, float | None]:
    """Сумма и прибыль при цене price. basis=potential считает на остатке и товаре в пути."""
    qty = money_qty(r, basis)
    if price is None or (qty is None and not r.cat_a):
        return None, None
    if qty is not None:
        sale = qty * price
    elif basis != "potential" and r.sales is not None and r.cat_a:
        sale = r.sales * price / r.cat_a
    else:
        return None, None
    profit = sale * (1 - r.sbs / price) if r.sbs is not None and price else None
    return sale, profit


def margin(rows: list[Row], rec: bool = False, basis: str = "sales") -> tuple[float, float, float | None]:
    """(сумма, сумма прибыли, ВР) — ВР по товарам, у которых есть СБС."""
    sale_all = profit_all = sale_vr = 0.0
    for r in rows:
        price = (r.rec.price if r.rec else None) if rec else r.cat_a
        sale, profit = money(r, price, basis)
        sale_all += sale or 0
        if profit is not None:
            profit_all += profit
            sale_vr += sale or 0
    return sale_all, profit_all, (profit_all / sale_vr if sale_vr else None)


def buckets(tol: float) -> list[tuple[str, float | None, float | None]]:
    """Корзины отклонения от цели: (подпись, от, до). Откл = 1 − наша / цель: минус — мы дороже."""
    edges = [tol, *[x for x in (0.1, 0.2) if x > tol]]

    def name(i: int) -> str:
        lo = f"{edges[i] * 100:g}"
        return f"{lo}–{edges[i + 1] * 100:g}%" if i + 1 < len(edges) else f"больше {lo}%"

    dearer = [(f"Дороже {name(i)}", -edges[i + 1] if i + 1 < len(edges) else None, -edges[i]) for i in range(len(edges))]
    cheaper = [(f"Дешевле {name(i)}", edges[i], edges[i + 1] if i + 1 < len(edges) else None) for i in range(len(edges))]
    return [*reversed(dearer), (f"В пределах ±{tol * 100:g}%", -tol, tol), *cheaper]


def in_bucket(v: float, lo: float | None, hi: float | None, tol: float) -> bool:
    if lo == -tol and hi == tol:
        return -tol <= v <= tol
    if hi is not None and hi <= -tol:  # «дороже»: [lo, hi)
        return (lo is None or v >= lo) and v < hi
    return v > lo and (hi is None or v <= hi)  # «дешевле»: (lo, hi]


def describe(values: list[float], weights: list[float] | None = None) -> dict:
    """Среднее с выбросами и без, медиана, очистка 5% / 10% крайних, выбросы по 1,5 × IQR."""
    n = len(values)
    if not n:
        return {"n": 0}
    s = sorted(values)

    def trimmed(p: float) -> float:
        k = int(n * p)
        part = s[k: n - k] or s
        return sum(part) / len(part)

    q1, q3 = (statistics.quantiles(s, n=4, method="inclusive")[i] for i in (0, 2)) if n > 1 else (s[0], s[0])
    lo, hi = q1 - 1.5 * (q3 - q1), q3 + 1.5 * (q3 - q1)
    clean = [v for v in s if lo <= v <= hi]
    weighted = None
    if weights and sum(weights) > 0:
        weighted = sum(v * w for v, w in zip(values, weights)) / sum(weights)
    return {
        "n": n, "mean": sum(s) / n, "weighted": weighted, "median": statistics.median(s),
        "trim5": trimmed(0.05), "trim10": trimmed(0.10), "iqr": sum(clean) / len(clean) if clean else None,
        "outliers": n - len(clean), "min": s[0], "max": s[-1],
    }


@dataclass
class Slice:
    title: str
    kind: str  # all | abc | comp | rec
    stats: dict


def deviation_slices(rows: list[Row], rec_enabled: bool) -> list[Slice]:
    weight = sales_weight(rows)

    def gaps(rs, attr="gap"):
        pairs = [(getattr(r, attr), weight(r)) for r in rs if getattr(r, attr) is not None]
        return describe([g for g, _ in pairs], [w for _, w in pairs])

    result = [Slice("Все товары с целью", "all", gaps(rows))]
    if rec_enabled:
        result.append(Slice("Все — после рекомендаций", "rec", gaps(rows, "rec_gap")))
    for abc in sorted({r.abc for r in rows}, key=lambda a: (ABC_ORDER.get(a, 99), a)):
        result.append(Slice(f"ABC: {abc or 'пусто'}", "abc", gaps([r for r in rows if r.abc == abc])))
    for c in competitor_order(rows):
        now, after = [], []
        for r in rows:
            price = r.prices.get(c.key)
            if not (r.cat_a and price):
                continue
            now.append((1 - r.cat_a / price, weight(r)))
            if rec_enabled and r.rec and r.rec.price:
                after.append((1 - r.rec.price / price, weight(r)))
        stats = describe([v for v, _ in now], [w for _, w in now])
        pcts = [r.percents.get(c.key) for r in rows if r.prices.get(c.key) and r.percents.get(c.key) is not None]
        if pcts:
            stats["target_dev"] = -statistics.median(pcts) / 100  # где цель в тех же единицах: «−5%» → мы дешевле на 5%
        if after:
            rec_stats = describe([v for v, _ in after], [w for _, w in after])
            stats["mean_rec"] = rec_stats.get("mean")
            stats["median_rec"] = rec_stats.get("median")
            stats["weighted_rec"] = rec_stats.get("weighted")
        result.append(Slice(f"Против {c.name} (без учёта % позиц.)", "comp", stats))
    return [s for s in result if s.stats["n"]]


def summary(rows: list[Row], opt: Options) -> dict:
    """Всё, что показывает дашборд, — числами (для окна программы, выводов и проверки формул)."""
    ro = opt.recs
    tol = ro.tolerance / 100
    basis = opt.money_base if opt.money_base in ("sales", "potential") else "sales"
    sale, profit, vr = margin(rows, basis=basis)
    sale_r, profit_r, vr_r = margin(rows, rec=True, basis=basis) if ro.enabled else (None, None, None)
    with_target = [r for r in rows if r.gap is not None]
    bucket_rows = []
    for label, lo, hi in buckets(tol):
        now = sum(in_bucket(r.gap, lo, hi, tol) for r in with_target)
        after = sum(in_bucket(r.rec_gap, lo, hi, tol) for r in rows if r.rec_gap is not None) if ro.enabled else None
        bucket_rows.append({"label": label, "now": now, "rec": after})

    def group_table(key, values):
        out = []
        for v in values:
            part = [r for r in rows if key(r) == v]
            if not part:
                continue
            s, p, m = margin(part, basis=basis)
            _, _, m_r = margin(part, rec=True, basis=basis) if ro.enabled else (None, None, None)
            gaps = [r.gap for r in part if r.gap is not None]
            gaps_rec = [r.rec_gap for r in part if r.rec_gap is not None] if ro.enabled else []
            out.append({
                "name": v, "items": len(part), "sales_qty": sum(r.sales_qty or 0 for r in part), "sale": s, "profit": p,
                "vr": m, "vr_rec": m_r, "priced": len(gaps), "mean_gap": sum(gaps) / len(gaps) if gaps else None,
                "mean_gap_rec": sum(gaps_rec) / len(gaps_rec) if gaps_rec else None,
                "in_band": sum(abs(g) <= tol + 1e-12 for g in gaps),
                "in_band_rec": sum(abs(g) <= tol + 1e-12 for g in gaps_rec) if ro.enabled else None,
            })
        return out

    abc_values = sorted({r.abc for r in rows}, key=lambda a: (ABC_ORDER.get(a, 99), a))
    seg_values = [s for s in [*reversed(SEGMENTS), UNKNOWN] if any(r.car and r.car.segment == s for r in rows)]
    kinds = {"up": 0, "down": 0, "keep": 0, "vr": 0, "none": 0}
    for r in rows:
        if r.rec:
            kinds[r.rec.kind] = kinds.get(r.rec.kind, 0) + 1
    impacts = []
    if ro.enabled:
        for r in rows:
            if r.rec and r.rec.price and r.cat_a and abs(r.rec.price - r.cat_a) > 0.005:
                _, p0 = money(r, r.cat_a, basis)
                _, p1 = money(r, r.rec.price, basis)
                impacts.append({"code": r.code, "article": r.article, "name": r.name, "cat_a": r.cat_a, "rec": r.rec.price,
                                "change": r.rec.change, "sales_qty": r.sales_qty, "weight": r.weight,
                                "profit_delta": (p1 or 0) - (p0 or 0), "reason": r.rec.reason})
        impacts.sort(key=lambda x: -abs(x["profit_delta"]))
    low_vr = sum(1 for r in rows if r.sbs is not None and r.cat_a and 1 - r.sbs / r.cat_a < ro.min_vr / 100)
    low_vr_rec = (sum(1 for r in rows if r.sbs is not None and r.rec and r.rec.price and 1 - r.sbs / r.rec.price < ro.min_vr / 100)
                  if ro.enabled else None)
    result = {
        "items": len(rows), "priced": sum(r.any_price for r in rows), "with_target": len(with_target),
        "in_band": sum(abs(r.gap) <= tol + 1e-12 for r in with_target),
        "in_band_rec": sum(abs(r.rec_gap) <= tol + 1e-12 for r in rows if r.rec_gap is not None) if ro.enabled else None,
        "sale": sale, "profit": profit, "vr": vr, "sale_rec": sale_r, "profit_rec": profit_r, "vr_rec": vr_r,
        "low_vr": low_vr, "low_vr_rec": low_vr_rec, "min_vr": ro.min_vr, "tolerance": ro.tolerance,
        "buckets": bucket_rows, "abc": group_table(lambda r: r.abc, abc_values),
        "subgroups": group_table(lambda r: r.subgroup, list(dict.fromkeys(r.subgroup for r in rows))),
        "segments": group_table(lambda r: r.car.segment if r.car else UNKNOWN, seg_values),
        "slices": [{"title": s.title, "kind": s.kind, **s.stats} for s in deviation_slices(rows, ro.enabled)],
        "rec_enabled": ro.enabled, "money_base": basis, "kinds": kinds, "top_changes": impacts[:12],
    }
    result["change_hist"] = _change_hist(rows) if ro.enabled else None
    result["limits"] = _limits(rows) if ro.enabled else None
    result["position"] = _position(rows, tol, ro.enabled, basis)
    result["bridge"] = _bridge(rows, basis) if ro.enabled else None
    result["weight_by"] = "qty" if any((r.sales_qty or 0) > 0 for r in rows) else "rub"
    result["conclusions"] = conclusions(result)
    return result


POSITION = [
    ("no_price", "нет цены конкурента"), ("d20", "дороже цели >20%"), ("d10", "дороже 10–20%"), ("d", "дороже до 10%"),
    ("band", "в коридоре"), ("c", "дешевле до 10%"), ("c10", "дешевле 10–20%"), ("c20", "дешевле цели >20%"),
]


def _pos_key(gap: float | None, tol: float) -> str:
    if gap is None:
        return "no_price"
    if abs(gap) <= tol + 1e-12:
        return "band"
    if gap < 0:
        return "d20" if gap < -0.2 else "d10" if gap < -0.1 else "d"
    return "c20" if gap > 0.2 else "c10" if gap > 0.1 else "c"


def _position(rows: list[Row], tol: float, rec: bool, basis: str) -> dict:
    """Где наша цена относительно цели — доля выручки и товаров, сейчас и после рекомендаций."""
    def mix(attr: str, price_of) -> list[dict]:
        acc = {k: [0, 0.0] for k, _ in POSITION}
        for r in rows:
            key = _pos_key(getattr(r, attr), tol)
            sale, _ = money(r, price_of(r), basis)
            acc[key][0] += 1
            acc[key][1] += sale or 0
        return [{"key": k, "label": t, "items": acc[k][0], "sale": acc[k][1]} for k, t in POSITION]

    out = {"now": mix("gap", lambda r: r.cat_a)}
    if rec:
        out["after"] = mix("rec_gap", lambda r: r.rec.price if r.rec and r.rec.price else r.cat_a)
    return out


def _bridge(rows: list[Row], basis: str) -> dict:
    """Откуда меняется прибыль при тех же продажах и сколько продаж нужно, чтобы шаг окупился.

    Снижение окупается, если продажи этих товаров вырастут на (прибыль до / прибыль после − 1).
    Повышение безопасно, пока продажи упадут не больше чем на (1 − прибыль до / прибыль после).
    """
    parts = {"up": [0, 0.0, 0.0, 0.0, []], "down": [0, 0.0, 0.0, 0.0, []]}
    profit_now = 0.0
    for r in rows:
        _, p0 = money(r, r.cat_a, basis)
        profit_now += p0 or 0
        if not (r.rec and r.rec.price and r.cat_a) or abs(r.rec.price - r.cat_a) < 0.005:
            continue
        _, p1 = money(r, r.rec.price, basis)
        sale0, _ = money(r, r.cat_a, basis)
        side = "up" if r.rec.price > r.cat_a else "down"
        x = parts[side]
        x[0] += 1
        x[1] += p0 or 0
        x[2] += p1 or 0
        x[3] += sale0 or 0
        x[4].append(r.rec.change)
    out = {"profit": profit_now}
    for side, (n, p0, p1, sale0, changes) in parts.items():
        item = {"items": n, "profit_before": p0, "profit_after": p1, "delta": p1 - p0, "sale": sale0,
                "median_change": statistics.median(changes) if changes else None}
        if side == "down":
            item["need_lift"] = (p0 / p1 - 1) if p1 > 0 else None
        else:
            item["may_lose"] = (1 - p0 / p1) if p1 > 0 else None
        out[side] = item
    out["after"] = profit_now + out["up"]["delta"] + out["down"]["delta"]
    return out


CHANGE_BUCKETS = [
    (-1e9, -0.10, "< −10%"), (-0.10, -0.05, "−10…−5%"), (-0.05, -0.02, "−5…−2%"), (-0.02, -0.0005, "−2…0%"),
    (-0.0005, 0.0005, "0"), (0.0005, 0.02, "0…+2%"), (0.02, 0.05, "+2…5%"), (0.05, 0.10, "+5…10%"), (0.10, 1e9, "> +10%"),
]


def _change_hist(rows: list[Row]) -> dict:
    """Насколько меняются цены по рекомендации: товары и продажи в штуках в каждом диапазоне."""
    items = [0] * len(CHANGE_BUCKETS)
    sales = [0.0] * len(CHANGE_BUCKETS)
    for r in rows:
        if not r.rec or r.rec.change is None or not r.cat_a:
            continue
        for i, (lo, hi, _label) in enumerate(CHANGE_BUCKETS):
            if lo <= r.rec.change < hi:
                items[i] += 1
                sales[i] += r.sales_qty or 0
                break
    return {"labels": [b[2] for b in CHANGE_BUCKETS], "items": items, "sales": sales}


def _limits(rows: list[Row]) -> list[dict]:
    """Почему рекомендация такая: по тексту причины у каждого товара."""
    found = {"band": 0, "free": 0, "capped": 0, "payback": 0, "vr": 0, "top": 0, "no_price": 0}
    for r in rows:
        if not r.rec:
            continue
        reason = r.rec.reason or ""
        if reason.startswith("нет цены конкурента"):
            found["no_price"] += 1
        elif reason.startswith("в пределах позиционирования"):
            found["band"] += 1
        elif reason.startswith("топ продаж"):
            found["top"] += 1
        elif r.rec.kind == "vr" or "не ниже ВР" in reason:
            found["vr"] += 1
        elif "окупаемостью" in reason:
            found["payback"] += 1
        elif "ограничено" in reason:
            found["capped"] += 1
        else:
            found["free"] += 1
    titles = {
        "free": "дошли до цели", "capped": "упёрлись в лимит шага", "band": "уже в коридоре — не трогаем",
        "payback": "снижение не окупится — режем меньше",
        "vr": "подняли до минимальной ВР", "top": "топ продаж категории A — не снижаем", "no_price": "нет цены конкурента",
    }
    return [{"key": k, "title": titles[k], "items": v} for k, v in found.items() if v]


def _p(v: float | None, signed: bool = False) -> str:
    if v is None:
        return "—"
    text = f"{v * 100:+.1f}%" if signed else f"{v * 100:.1f}%"
    return text.replace(".", ",").replace("-", "−")


def _pp(points: float) -> str:
    text = f"{abs(points):.1f}".replace(".", ",")
    sign = "+" if points > 0 else ("−" if points < 0 else "")
    return f"{sign}{text} п.п."


def conclusions(s: dict) -> list[str]:
    out = []
    n = s["items"]
    if not n:
        return out
    out.append(f"Цена конкурента есть у {s['priced']} из {n} товаров ({s['priced'] / n:.0%}).")
    if s["with_target"]:
        t = s["with_target"]
        dearer = sum(b["now"] for b in s["buckets"] if b["label"].startswith("Дороже"))
        cheaper = sum(b["now"] for b in s["buckets"] if b["label"].startswith("Дешевле"))
        out.append(f"В пределах ±{s['tolerance']:g}% от цели — {s['in_band']} из {t} ({s['in_band'] / t:.0%}); "
                   f"дороже цели — {dearer}, дешевле — {cheaper}.")
        main = next((x for x in s["slices"] if x["kind"] == "all"), None)
        if main:
            side = "дороже" if main["median"] < 0 else "дешевле"
            out.append(f"Медианное отклонение от цели {_p(main['median'], True)} — в целом мы {side} цели. "
                       f"Без выбросов среднее {_p(main['iqr'], True)} (выбросов {main['outliers']}); "
                       f"сырое среднее {_p(main['mean'], True)} тянут единичные кроссы с ошибочной ценой.")
        worst = [x for x in s["abc"] if x["mean_gap"] is not None and x["priced"] >= 3 and abs(x["mean_gap"]) < 2]
        if worst:
            w = min(worst, key=lambda x: x["mean_gap"])
            if w["mean_gap"] < 0:
                out.append(f"Сильнее всего дороже цели категория {w['name'] or 'без ABC'}: в среднем {_p(w['mean_gap'], True)}.")
    if s["vr"] is not None:
        line = f"ВР группы сейчас {_p(s['vr'])}"
        if s["low_vr"]:
            line += f"; ниже {s['min_vr']:g}% — {s['low_vr']} товаров"
        out.append(line + ".")
    if s["rec_enabled"] and s["vr_rec"] is not None and s["vr"] is not None:
        k = s["kinds"]
        delta = (s["vr_rec"] - s["vr"]) * 100
        sale_ch = s["sale_rec"] / s["sale"] - 1 if s["sale"] else None
        profit_delta = s["profit_rec"] - s["profit"]
        profit_text = f"{abs(profit_delta):,.0f}".replace(",", " ") + " руб."
        if profit_delta < 0:
            profit_text = "−" + profit_text
        elif profit_delta > 0:
            profit_text = "+" + profit_text
        out.append(f"Рекомендации: поднять {k['up'] + k['vr']}, понизить {k['down']}, не трогать {k['keep']}. "
                   f"ВР {_p(s['vr'])} → {_p(s['vr_rec'])} ({_pp(delta)}), "
                   f"выручка при тех же продажах {_p(sale_ch, True)}, прибыль {profit_text}")
        if s["in_band_rec"] is not None and s["with_target"]:
            out.append(f"В позиционировании после рекомендаций — {s['in_band_rec']} из {s['with_target']} "
                       f"(было {s['in_band']}): остальное не дотянуть в пределах лимитов изменения цены.")
    return out


# ---------------------------------------------------------------- лист


def write_dashboard(ws, prep, opt: Options, infos: list, title: str, src) -> dict:
    rows = prep.kept
    ro = opt.recs
    tol = ro.tolerance / 100
    s = summary(rows, opt)
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.column_dimensions["A"].width = 2
    for col in "BCDEFGHIJKLMNOP":
        ws.column_dimensions[col].width = 13
    ws.column_dimensions["B"].width = 30

    def across(key: str, template: str, *extra: str) -> str:
        """Сумма template по всем листам отчёта; template — с {r} и {k} ({k} — диапазоны extra)."""
        parts = []
        for info in infos:
            if key not in info.letters or any(e not in info.letters for e in extra):
                continue
            text = template.replace("{r}", info.rng(key))
            for i, e in enumerate(extra):
                text = text.replace("{k" + str(i) + "}", info.rng(e))
            parts.append(text)
        return "+".join(parts) if parts else "0"

    def cell(ref: str, value, fmt: str | None = None, font: Font | None = None, fill=None, align=None):
        c = ws[ref]
        c.value = value
        if fmt:
            c.number_format = fmt
        c.font = font or Font(name=style.FONT, color=style.INK)
        if fill:
            c.fill = fill
        if align:
            c.alignment = align
        return c

    ws["B2"] = f"Дашборд · {title}"
    ws["B2"].font = Font(name=style.FONT, size=18, bold=True, color=style.ACCENT)
    ws.row_dimensions[2].height = 30
    period = f" · продажи: {src.sales_period}" if getattr(src, "sales_period", "") else ""
    ws["B3"] = (f"{REPORT_KINDS[opt.report_kind]} отчёт · {dt.datetime.now():%d.%m.%Y %H:%M} · "
                f"{SUBGROUP_MODES[opt.subgroup_mode].lower()} · цены: {PRICE_SOURCES[opt.price_source]}{period} · "
                f"цель: {TARGETS[ro.target].split(' (')[0].lower()}")
    ws["B3"].font = Font(name=style.FONT, color=style.MUTED)

    # ---- карточки
    items = f"=({across('code', 'COUNTA({r})')})"
    sale = f"({across('sum_sale', 'SUM({r})')})"
    profit = f"({across('sum_profit', 'SUM({r})')})"
    sale_vr = f"({across('sum_sale', f'SUMIFS({{r}},{{k0}},\"{NUM}\")', 'vr')})"
    priced = f"({across('target', 'COUNT({r})')})"
    in_band = f"({across('gap', f'COUNTIFS({{r}},\">=\"&-{tol},{{r}},\"<=\"&{tol})')})"
    vr_now = f"IFERROR({profit}/{sale_vr},\"\")"
    rec = ro.enabled and any("rec" in i.letters for i in infos)
    if rec:
        sale_r = f"({across('sum_sale_rec', 'SUM({r})')})"
        profit_r = f"({across('sum_profit_rec', 'SUM({r})')})"
        sale_vr_r = f"({across('sum_sale_rec', f'SUMIFS({{r}},{{k0}},\"{NUM}\")', 'vr_rec')})"
        vr_rec = f"IFERROR({profit_r}/{sale_vr_r},\"\")"
    low_vr = f"({across('vr', f'COUNTIFS({{r}},\"<\"&{ro.min_vr / 100})')})"

    cards = [
        ("Товаров в отчёте", items, '#,##0', f"=IFERROR({priced}/{items[1:]},0)", '"с ценой конкурента: "0%'),
        ("ВР сейчас", f"={vr_now}", PCT1, f"={low_vr}", f'"ниже {ro.min_vr:g}%: "0" шт."'),
        ("ВР после рекомендаций", f"={vr_rec}" if rec else "не вставлялись", PCT1,
         f'=IFERROR(({vr_rec}-{vr_now})*100,"")' if rec else None, PP),
        ("Сумма продажи", f"={sale}", MONEY0, f"=IFERROR({sale_r}/{sale}-1,\"\")" if rec else None,
         '"после рек.: "+0.0%;"после рек.: "-0.0%;"после рек.: 0%"'),
        ("Сумма прибыли", f"={profit}", MONEY0, f"={profit_r}-{profit}" if rec else None,
         '"после рек.: "+#,##0" ₽";"после рек.: "-#,##0" ₽";"после рек.: 0 ₽"'),
        (f"В позиционировании ±{ro.tolerance:g}%", f"={in_band}", '#,##0', f"=IFERROR({in_band}/{priced},0)",
         '0%" товаров с целью"'),
    ]
    cols = ["B", "D", "F", "H", "J", "L"]
    pairs = {"B": "C", "D": "E", "F": "G", "H": "I", "J": "K", "L": "M"}
    ws.column_dimensions["B"].width = 16
    for col, (label, value, fmt, note, note_fmt) in zip(cols, cards):
        right = pairs[col]
        for r in (5, 6, 7):
            ws.merge_cells(f"{col}{r}:{right}{r}")
            for c in (col, right):
                ws[f"{c}{r}"].fill = CARD_FILL
                ws[f"{c}{r}"].border = Border(left=CARD_LINE if c == col else None, right=CARD_LINE if c == right else None,
                                              top=CARD_LINE if r == 5 else None, bottom=CARD_LINE if r == 7 else None)
        cell(f"{col}5", label, font=Font(name=style.FONT, size=10, color=style.MUTED), fill=CARD_FILL,
             align=Alignment(indent=1, vertical="bottom"))
        big = value if isinstance(value, str) and value.startswith("=") else value
        cell(f"{col}6", big, fmt, Font(name=style.FONT, size=20 if big != "не вставлялись" else 12, bold=True,
                                         color=style.ACCENT if big != "не вставлялись" else style.MUTED),
             CARD_FILL, Alignment(indent=1, vertical="center"))
        cell(f"{col}7", note, note_fmt, Font(name=style.FONT, size=10, color=style.MUTED), CARD_FILL,
             Alignment(indent=1, vertical="top"))
    ws.row_dimensions[5].height = 20
    ws.row_dimensions[6].height = 34
    ws.row_dimensions[7].height = 20
    ws.merge_cells("N5:P7")
    cell("N5", "ВР — по товарам, у которых есть СБС. Всё в карточках и таблицах ниже — формулы по листам отчёта.",
         font=Font(name=style.FONT, italic=True, size=9, color=style.MUTED),
         align=Alignment(wrap_text=True, vertical="center"))

    def section(row: int, text: str, hint: str = "") -> None:
        cell(f"B{row}", text, font=Font(name=style.FONT, size=13, bold=True, color=style.INK))
        if hint:
            cell(f"F{row}", hint, font=Font(name=style.FONT, italic=True, size=9, color=style.MUTED))
        ws.row_dimensions[row].height = 22

    def header(row: int, titles: list[str], start: str = "B") -> None:
        base = ord(start)
        for i, t in enumerate(titles):
            c = ws[f"{chr(base + i)}{row}"]
            c.value = t
            c.font = Font(name=style.FONT, bold=True, color="FFFFFF", size=10)
            c.fill = HEAD_FILL
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[row].height = 30

    # ---- корзины отклонений
    row = 10
    section(row, "Отклонение Кат. А от цели позиционирования",
            "Откл = 1 − Кат. А / цель: минус — мы дороже цели. Сейчас и после рекомендованных цен.")
    header(row + 1, ["Корзина", "Товаров сейчас", "Доля", *(["После рек.", "Доля"] if rec else [])])
    first_b = row + 2
    gap_bucket = []
    bucket_list = buckets(tol)
    dearer_colors, cheaper_colors = (GREENS, REDS) if opt.green_when == "above" else (REDS, GREENS)
    for i, (label, lo, hi) in enumerate(bucket_list):
        r = first_b + i
        band = lo == -tol and hi == tol
        if band:
            color = "E2E8F0"
        else:
            dearer_side = hi is not None and hi <= -tol
            shades = dearer_colors if dearer_side else cheaper_colors
            k = i if dearer_side else len(bucket_list) - 1 - i
            color = shades[min(k, len(shades) - 1)]
        cell(f"B{r}", label, font=Font(name=style.FONT, bold=True, color=style.INK if band else "FFFFFF"),
             fill=PatternFill("solid", fgColor=color))

        def crit(key):
            if lo == -tol and hi == tol:
                return across(key, f'COUNTIFS({{r}},">="&{-tol},{{r}},"<="&{tol})')
            if hi is not None and hi <= -tol:
                cond = (f'{{r}},">="&{lo},' if lo is not None else "") + f'{{r}},"<"&{hi}'
            else:
                cond = f'{{r}},">"&{lo}' + (f',{{r}},"<="&{hi}' if hi is not None else "")
            return across(key, f"COUNTIFS({cond})")

        cell(f"C{r}", f"={crit('gap')}", '#,##0')
        cell(f"D{r}", f"=IFERROR(C{r}/SUM($C${first_b}:$C${first_b + len(bucket_list) - 1}),0)", SHARE)
        if rec:
            cell(f"E{r}", f"={crit('gap_rec')}", '#,##0')
            cell(f"F{r}", f"=IFERROR(E{r}/SUM($E${first_b}:$E${first_b + len(bucket_list) - 1}),0)", SHARE)
        gap_bucket.append(r)
    last_b = gap_bucket[-1]
    chart = BarChart()
    chart.type = "bar"
    chart.style = 10
    chart.title = "Товаров по отклонению от цели"
    chart.y_axis.majorGridlines = None
    chart.x_axis.scaling.orientation = "maxMin"
    data = Reference(ws, min_col=3, min_row=row + 1, max_row=last_b)
    chart.add_data(data, titles_from_data=True)
    if rec:
        chart.add_data(Reference(ws, min_col=5, min_row=row + 1, max_row=last_b), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=2, min_row=first_b, max_row=last_b))
    for ser, color in zip(chart.series, ["0F4C5C", "60A5FA"]):
        ser.graphicalProperties.solidFill = color
        ser.graphicalProperties.line.solidFill = color
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showVal = True
    chart.height, chart.width = 7.5, 15
    chart.legend.position = "b"
    ws.add_chart(chart, f"H{row + 1}")
    row = max(last_b, row + 15) + 2

    # ---- ВР по срезам
    def margin_table(row: int, name: str, key: str, values: list[str], label=lambda v: v or "пусто") -> int:
        if not values:
            return row
        section(row, name)
        titles = ["Срез", "Товаров", "Продажи, шт", "Сумма продажи", "ВР сейчас", "С целью", f"В ±{ro.tolerance:g}%",
                  "Ср. откл. от цели"]
        if rec:
            titles += ["ВР после рек.", "Δ ВР", "Изм. выручки"]
        header(row + 1, titles)
        for i, v in enumerate(values):
            r = row + 2 + i
            crit = f'"{v}"' if v else '""'
            cell(f"B{r}", label(v), font=Font(name=style.FONT, bold=True, color=style.INK), fill=SOFT_FILL if i % 2 else None)
            cell(f"C{r}", f"={across(key, f'COUNTIFS({{r}},{crit})')}", '#,##0')
            cell(f"D{r}", f"={across(key, f'SUMIFS({{k0}},{{r}},{crit})', 'sales_qty')}", '#,##0')
            cell(f"E{r}", f"={across(key, f'SUMIFS({{k0}},{{r}},{crit})', 'sum_sale')}", MONEY0)
            p = across(key, f"SUMIFS({{k0}},{{r}},{crit})", "sum_profit")
            sv = across(key, f'SUMIFS({{k0}},{{r}},{crit},{{k1}},"{NUM}")', "sum_sale", "vr")
            cell(f"F{r}", f'=IFERROR(({p})/({sv}),"")', PCT1, Font(name=style.FONT, bold=True, color=style.INK))
            cell(f"G{r}", f"={across(key, f'COUNTIFS({{r}},{crit},{{k0}},\"{NUM}\")', 'gap')}", '#,##0')
            cell(f"H{r}", f'={across(key, f"COUNTIFS({{r}},{crit},{{k0}},\">=\"&{-tol},{{k0}},\"<=\"&{tol})", "gap")}', '#,##0')
            gs = across(key, f"SUMIFS({{k0}},{{r}},{crit})", "gap")
            cell(f"I{r}", f'=IFERROR(({gs})/G{r},"")', PCT1S)
            if rec:
                pr = across(key, f"SUMIFS({{k0}},{{r}},{crit})", "sum_profit_rec")
                svr = across(key, f'SUMIFS({{k0}},{{r}},{crit},{{k1}},"{NUM}")', "sum_sale_rec", "vr_rec")
                sr = across(key, f"SUMIFS({{k0}},{{r}},{crit})", "sum_sale_rec")
                cell(f"J{r}", f'=IFERROR(({pr})/({svr}),"")', PCT1, Font(name=style.FONT, bold=True, color="1D5FA8"))
                cell(f"K{r}", f'=IFERROR((J{r}-F{r})*100,"")', PP)
                cell(f"L{r}", f'=IFERROR(({sr})/E{r}-1,"")', PCT1S)
            if i % 2:
                for c in "CDEFGHIJKL":
                    ws[f"{c}{r}"].fill = SOFT_FILL
        last = row + 1 + len(values)
        rng = f"F{row + 2}:F{last}"
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER(F{row + 2}),F{row + 2}<{ro.min_vr / 100})"],
                                                       font=Font(color="C2410C", bold=True)))
        dearer_fill, cheaper_fill = ("DCFCE7", "FEE2E2") if opt.green_when == "above" else ("FEE2E2", "DCFCE7")
        ws.conditional_formatting.add(f"I{row + 2}:I{last}", FormulaRule(formula=[f"AND(ISNUMBER(I{row + 2}),I{row + 2}<-{tol})"],
                                                                         fill=PatternFill(bgColor=dearer_fill)))
        ws.conditional_formatting.add(f"I{row + 2}:I{last}", FormulaRule(formula=[f"AND(ISNUMBER(I{row + 2}),I{row + 2}>{tol})"],
                                                                         fill=PatternFill(bgColor=cheaper_fill)))
        if rec:
            ws.conditional_formatting.add(f"K{row + 2}:K{last}", FormulaRule(formula=[f"AND(ISNUMBER(K{row + 2}),K{row + 2}>0.05)"],
                                                                             font=Font(color="15803D", bold=True)))
            ws.conditional_formatting.add(f"K{row + 2}:K{last}", FormulaRule(formula=[f"AND(ISNUMBER(K{row + 2}),K{row + 2}<-0.05)"],
                                                                             font=Font(color="B91C1C", bold=True)))
        return last

    abc_values = [x["name"] for x in s["abc"]]
    abc_head = row
    row = margin_table(row, "ВР и позиционирование по категориям ABC", "abc", abc_values) + 2
    if abc_values:
        chart = BarChart()
        chart.type = "col"
        chart.title = "ВР по ABC"
        chart.y_axis.numFmt = "0%"
        chart.y_axis.majorGridlines = None
        chart.add_data(Reference(ws, min_col=6, min_row=abc_head + 1, max_row=abc_head + 1 + len(abc_values)), titles_from_data=True)
        if rec:
            chart.add_data(Reference(ws, min_col=10, min_row=abc_head + 1, max_row=abc_head + 1 + len(abc_values)), titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=2, min_row=abc_head + 2, max_row=abc_head + 1 + len(abc_values)))
        for ser, color in zip(chart.series, ["0F4C5C", "60A5FA"]):
            ser.graphicalProperties.solidFill = color
            ser.graphicalProperties.line.solidFill = color
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showVal = True
        chart.height, chart.width = 7, 13
        chart.legend.position = "b"
        ws.add_chart(chart, f"{'N' if rec else 'K'}{abc_head + 1}")
        row = max(row, abc_head + 16)
    subgroups = [x["name"] for x in s["subgroups"]]
    if len(subgroups) > 1:
        row = margin_table(row, "По подгруппам", "subgroup", subgroups, lambda v: v or "без подгруппы") + 2
    segs = [x["name"] for x in s["segments"]]
    if segs and any("segment" in i.letters for i in infos):
        row = margin_table(row, "По ценовому сегменту автопарка", "segment", segs) + 2

    # ---- статистика отклонений (снимок)
    section(row, "Статистика отклонений от цели", "на момент формирования; минус — мы дороже")
    titles = ["Срез", "Товаров", "Среднее", "Взвеш. по продажам", "Медиана", "Без 5% крайних", "Без 10% крайних",
              "Без выбросов (1,5×IQR)", "Выбросов", "Мин.", "Макс."]
    header(row + 1, titles)
    r = row + 2
    for i, sl in enumerate(s["slices"]):
        vals = [sl["title"], sl["n"], sl["mean"], sl["weighted"], sl["median"], sl["trim5"], sl["trim10"], sl["iqr"],
                sl["outliers"], sl["min"], sl["max"]]
        for j, v in enumerate(vals):
            c = ws.cell(r, 2 + j, v)
            c.font = Font(name=style.FONT, color=style.INK, bold=j == 0 or j == 4,
                          italic=sl["kind"] in ("comp", "rec") and j == 0)
            c.number_format = '#,##0' if j in (1, 8) else PCT1S
            if sl["kind"] == "rec":
                c.fill = PatternFill("solid", fgColor="EAF2FD")
            elif i % 2:
                c.fill = SOFT_FILL
            if j in (2, 3, 4, 5, 6, 7) and isinstance(v, float):
                if abs(v) > tol:
                    good = (v < 0) == (opt.green_when == "above")
                    c.font = Font(name=style.FONT, color="15803D" if good else "B91C1C", bold=j == 4)
        r += 1
    ws.column_dimensions["B"].width = 34
    row = r + 1
    cell(f"B{row - 1}", "Медиана и «без выбросов» — честнее среднего: единичные кроссы с ошибочной ценой его сильно сдвигают.",
         font=Font(name=style.FONT, italic=True, size=9, color=style.MUTED))
    row += 1

    # ---- рекомендации
    if rec:
        section(row, "Рекомендованные цены", f"лимиты: −{ro.down:g}% / +{', '.join(f'{v:g}' for v in ro.up.values())}%, "
                                               f"ВР не ниже {ro.min_vr:g}%, топ: первый конкурент, A вниз не снижаем, прочий топ до −2%")
        header(row + 1, ["Действие", "Товаров"])
        ch = lambda tmpl: across("change", tmpl)  # noqa: E731
        lines = [
            ("Поднять цену", f"={ch('COUNTIFS({r},\">0.0005\")')}", "15803D"),
            ("Понизить цену", f"={ch('COUNTIFS({r},\"<-0.0005\")')}", "B45309"),
            ("Оставить", f"={ch('COUNTIFS({r},\">=-0.0005\",{r},\"<=0.0005\")')}", style.INK),
            ("из них подняты до мин. ВР", s["kinds"].get("vr", 0), "C2410C"),
        ]
        for i, (label, value, color) in enumerate(lines):
            rr = row + 2 + i
            cell(f"B{rr}", label, font=Font(name=style.FONT, color=color, bold=True))
            cell(f"C{rr}", value, '#,##0', Font(name=style.FONT, color=color, bold=True))
        row += 2 + len(lines) + 1
        if s["top_changes"]:
            cell(f"B{row}", "Крупнейшие изменения по влиянию на прибыль (снимок)",
                 font=Font(name=style.FONT, bold=True, color=style.INK))
            header(row + 1, ["Наименование", "Артикул", "Кат. А", "Рек. кат. А", "Изм.", "Продажи, шт", "Δ прибыли", "Вес", "Основание"])
            for i, t in enumerate(s["top_changes"]):
                rr = row + 2 + i
                vals = [t["name"], t["article"], t["cat_a"], t["rec"], t["change"], t["sales_qty"], t["profit_delta"],
                        {"top": "топ", "mid": "середина"}.get(t["weight"], "хвост"), t["reason"]]
                fmts = [None, None, MONEY0, MONEY0, PCT1S, '#,##0', '+#,##0;-#,##0;0', None, None]
                for j, (v, f) in enumerate(zip(vals, fmts)):
                    c = cell(f"{chr(ord('B') + j)}{rr}", v, f, Font(name=style.FONT, color=style.INK, size=10))
                    if j == 4:
                        c.font = Font(name=style.FONT, size=10, bold=True, color="1D5FA8" if (v or 0) > 0 else "B45309")
                    if j == 8:
                        c.font = Font(name=style.FONT, size=9, color=style.MUTED)
            row += 2 + len(s["top_changes"]) + 1

    # ---- выводы
    section(row, "Выводы", "на момент формирования")
    for i, text in enumerate(s["conclusions"]):
        rr = row + 1 + i
        ws.merge_cells(f"B{rr}:P{rr}")
        cell(f"B{rr}", f"•  {text}", font=Font(name=style.FONT, color=style.INK, size=11),
             align=Alignment(wrap_text=True, vertical="top"))
        ws.row_dimensions[rr].height = 30 if len(text) > 120 else 18
    ws.freeze_panes = "A4"
    return s
