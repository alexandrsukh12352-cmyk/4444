"""Рекомендованная Кат. А: подгонка к позиционированию с ограничениями.

Цель по позиционированию = цена конкурента × (1 + % позиционирования). Если наша Кат. А
в пределах допуска от цели — не трогаем. Иначе двигаем к цели, но не больше лимитов:
вниз — не больше N %, вверх — по ценовому сегменту автопарка (премиум можно сильнее,
бюджетный — почти нельзя). Для топовых по продажам артикулов лимит вверх меньше: они сильнее
всего влияют на ВР группы. Вниз топ почти не режем: категория A и AA — никак,
остальной топ — не больше 2%. Цель топа — только первый конкурент в позиционировании.
Нет его цены — на второго и третьего не ориентируемся. ВР после изменения — не ниже минимальной.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .cars import UNKNOWN
from .sources import norm_text

TARGETS = {
    "first": "Только первый в списке позиционирования. Нет цены — на остальных не ориентируемся",
    "main": "По списку: сначала первый, если цены нет — следующий, со своим процентом",
    "avg": "Среднее по всем, у кого есть цена",
    "min": "Самая низкая цель среди всех, у кого есть цена",
}
WEIGHT_CLASSES = {"top": "Топ продаж", "mid": "Середина", "tail": "Хвост"}
DEFAULT_UP = {"Премиум": 10.0, "Средний": 5.0, "Бюджетный": 2.0, UNKNOWN: 5.0}


@dataclass
class RecOptions:
    enabled: bool = False
    target: str = "main"
    tolerance: float = 3.0  # ± % — «в пределах позиционирования»
    down: float = 10.0  # не понижать больше, %
    up: dict[str, float] = field(default_factory=lambda: dict(DEFAULT_UP))  # не повышать больше, % — по сегменту
    min_vr: float = 32.0
    top_share: float = 50.0  # топ — артикулы, дающие первые N % продаж в шт
    mid_share: float = 80.0
    top_factor: float = 0.5  # во сколько раз ужимаются лимиты для топа
    mid_factor: float = 0.75
    rounding: float = 1.0  # шаг округления Кат. А, руб.
    max_lift: float = 20.0  # снижение окупается, только если продажи вырастут не больше чем на N %; 0 — не проверять
    priced_first: bool = True

    @classmethod
    def from_dict(cls, data: dict | None) -> "RecOptions":
        data = dict(data or {})
        known = {k: v for k, v in data.items() if k in cls.__dataclass_fields__}
        for key in ("tolerance", "down", "min_vr", "top_share", "mid_share", "top_factor", "mid_factor", "rounding", "max_lift"):
            if known.get(key) not in (None, ""):
                known[key] = float(known[key])
        if "enabled" in known:
            known["enabled"] = bool(known["enabled"])
        if "priced_first" in known:
            known["priced_first"] = bool(known["priced_first"])
        opt = cls(**known)
        opt.up = {**DEFAULT_UP, **{k: float(v) for k, v in (opt.up or {}).items() if v not in ("", None)}}
        if opt.target not in TARGETS:
            opt.target = "main"
        return opt


def bundled_rec() -> dict:
    """Зашитые пороги. Это стартовый набор, не разбор конкретных отчётов.

    После разбора файлов сюда лягут числа по группам в BUNDLED_GROUPS.
    """
    base = RecOptions()
    return {
        "tolerance": base.tolerance, "down": base.down, "min_vr": base.min_vr,
        "top_share": base.top_share, "mid_share": base.mid_share,
        "top_factor": base.top_factor, "mid_factor": base.mid_factor,
        "rounding": base.rounding, "max_lift": base.max_lift, "up": dict(base.up),
    }


# Ключ — «БРЕНД\\tГРУППА» после norm_text. Числа по уже разобранным отчётам.
# Включение и «по кому цель» переключатель в окне не затирает, если его прислали.
BUNDLED_GROUPS: dict[str, dict] = {
    # Диски: против Brembo есть запас вверх, но объём без Brembo уже дороже G-brake. Вверх умеренно.
    "ADVICS\tТОРМОЗНЫЕ ДИСКИ": {"min_vr": 34, "down": 6, "tolerance": 3, "top_factor": 0.5,
                               "up": {"Премиум": 6, "Средний": 4, "Бюджетный": 2, "Не определён": 4}},
    # Колодки: сильно дороже Akebono, ВР группы тонкая. Вниз с хвоста, вверх почти нет.
    "ADVICS\tТОРМОЗНЫЕ КОЛОДКИ": {"min_vr": 34, "down": 8, "tolerance": 3, "top_factor": 0.4,
                                 "up": {"Премиум": 2, "Средний": 2, "Бюджетный": 2, "Не определён": 2}},
    "ADVICS\tТОРМОЗНЫЕ СУППОРТЫ": {"min_vr": 32, "down": 4, "tolerance": 4, "up": {"Премиум": 4, "Средний": 3, "Бюджетный": 2, "Не определён": 3}},
    "ADVICS\tТОРМОЗНЫЕ ЦИЛИНДРЫ И ШЛАНГИ": {"min_vr": 34, "down": 4, "tolerance": 3, "up": {"Премиум": 4, "Средний": 3, "Бюджетный": 2, "Не определён": 3}},
    # Диски дешевле цели LYNX. Подтянуть можно, топ не разгонять.
    "AVANTECH\tТОРМОЗНЫЕ ДИСКИ": {"min_vr": 34, "down": 4, "tolerance": 3, "top_factor": 0.5,
                                 "up": {"Премиум": 8, "Средний": 5, "Бюджетный": 2, "Не определён": 5}},
    # Колодки почти в цели у Kashiyama.
    "AVANTECH\tТОРМОЗНЫЕ КОЛОДКИ": {"min_vr": 34, "down": 3, "tolerance": 4, "up": {"Премиум": 3, "Средний": 3, "Бюджетный": 2, "Не определён": 3}},
    # Фильтры дешевле LYNX, но много оверстока и распродажи — их не повышает отдельное правило.
    "AVANTECH\tФИЛЬТРЫ АВТОМОБИЛЬНЫЕ": {"min_vr": 34, "down": 6, "tolerance": 3, "top_factor": 0.5,
                                       "up": {"Премиум": 8, "Средний": 6, "Бюджетный": 3, "Не определён": 5}},
    "AVANTECH\tДЕТАЛИ ПОДВЕСКИ И РУЛЕВОГО УПРАВЛЕНИЯ": {"min_vr": 32, "down": 8, "up": {"Премиум": 2, "Средний": 2, "Бюджетный": 2, "Не определён": 2}},
    "AVANTECH\tВТУЛКИ И САЙЛЕНТБЛОКИ": {"min_vr": 32, "down": 8, "up": {"Премиум": 2, "Средний": 2, "Бюджетный": 2, "Не определён": 2}},
    "AVANTECH\tРЕМНИ, ЦЕПИ ГРМ": {"min_vr": 32, "down": 4, "up": {"Премиум": 3, "Средний": 2, "Бюджетный": 2, "Не определён": 2}},
    "AVANTECH\tПРИВОД И КОМПЛЕКТУЮЩИЕ": {"min_vr": 34, "down": 4, "up": {"Премиум": 5, "Средний": 4, "Бюджетный": 2, "Не определён": 4}},
    "AVANTECH\tОПОРЫ ДВИГАТЕЛЯ И КПП": {"min_vr": 32, "down": 5, "up": {"Премиум": 2, "Средний": 2, "Бюджетный": 2, "Не определён": 2}},
    # Фильтры дешевле MILES примерно на 23 п.п. JD в файле не было, у обоих процент 0.
    # Шаг вверх как у фильтров AVANTECH, не закрываем разрыв за один раз.
    "DOUBLE FORCE\tФИЛЬТРЫ АВТОМОБИЛЬНЫЕ": {"min_vr": 34, "down": 4, "tolerance": 3, "top_factor": 0.5,
                                           "up": {"Премиум": 8, "Средний": 6, "Бюджетный": 3, "Не определён": 5}},
    # Колодки примерно на 48% дороже JD при цели «вровень». ВР 44%, режем хвост, не всю группу.
    "DOUBLE FORCE\tТОРМОЗНЫЕ КОЛОДКИ": {"min_vr": 34, "down": 8, "tolerance": 3, "top_factor": 0.4,
                                       "up": {"Премиум": 2, "Средний": 2, "Бюджетный": 2, "Не определён": 2}},
    # Стойки около MILES, на 6 п.п. дешевле цели. Ходы короткие.
    "DOUBLE FORCE\tСТОЙКИ И АМОРТИЗАТОРЫ": {"min_vr": 34, "down": 3, "tolerance": 4, "top_factor": 0.5,
                                           "up": {"Премиум": 4, "Средний": 3, "Бюджетный": 2, "Не определён": 3}},
    # Шрусы, ступицы и катушки дороже JD. Вниз умеренно: кросс есть не у всех.
    "DOUBLE FORCE\tПРИВОД И КОМПЛЕКТУЮЩИЕ": {"min_vr": 34, "down": 6, "tolerance": 3,
                                            "up": {"Премиум": 2, "Средний": 2, "Бюджетный": 2, "Не определён": 2}},
    "DOUBLE FORCE\tСТУПИЦЫ И КОМПЛЕКТУЮЩИЕ": {"min_vr": 34, "down": 6, "tolerance": 3,
                                             "up": {"Премиум": 2, "Средний": 2, "Бюджетный": 2, "Не определён": 2}},
    "DOUBLE FORCE\tПРОВОДА И КАТУШКИ ЗАЖИГАНИЯ": {"min_vr": 34, "down": 6, "tolerance": 3, "top_factor": 0.5,
                                                 "up": {"Премиум": 2, "Средний": 2, "Бюджетный": 2, "Не определён": 2}},
    # Пружины без продаж, шланги на копейки. Рамку не раздуваем.
    "DOUBLE FORCE\tПРУЖИНЫ ПОДВЕСКИ": {"min_vr": 32, "down": 4, "tolerance": 4,
                                      "up": {"Премиум": 2, "Средний": 2, "Бюджетный": 2, "Не определён": 2}},
    "DOUBLE FORCE\tТОРМОЗНЫЕ ЦИЛИНДРЫ И ШЛАНГИ": {"min_vr": 32, "down": 4, "tolerance": 4,
                                                 "up": {"Премиум": 3, "Средний": 2, "Бюджетный": 2, "Не определён": 2}},
    # NITTO: масляные у VIC в цели, воздушные дешевле примерно на 17 п.п. MANN в файле нет,
    # поэтому вверх короче, чем у DF: шаг к VIC, не закрытие разрыва.
    "NITTO\tФИЛЬТРЫ АВТОМОБИЛЬНЫЕ": {"min_vr": 34, "down": 3, "tolerance": 3, "top_factor": 0.5,
                                    "up": {"Премиум": 5, "Средний": 4, "Бюджетный": 2, "Не определён": 4}},
    # SUN: цена конкурента есть у горстки артикулов, на ней мы дороже цели «−15%».
    # Вниз короткий шаг, вверх почти нет, допуск шире, чтобы шум не двигал группу.
    "SUN\tРЕМНИ, ЦЕПИ ГРМ": {"min_vr": 34, "down": 4, "tolerance": 5, "top_factor": 0.5,
                           "up": {"Премиум": 2, "Средний": 2, "Бюджетный": 2, "Не определён": 2}},
}


def bundled_group_key(brand: str, group: str) -> str:
    return f"{norm_text(brand)}\t{norm_text(group)}"


def bundled_for(brand: str, group: str) -> dict:
    base = bundled_rec()
    extra = BUNDLED_GROUPS.get(bundled_group_key(brand, group)) or {}
    up = {**base["up"], **(extra.get("up") or {})}
    return {**base, **{k: v for k, v in extra.items() if k != "up"}, "up": up}


@dataclass
class Rec:
    price: float | None
    change: float | None  # доля: 0.05 = +5 %
    reason: str
    kind: str  # keep | up | down | vr | none


def sales_weight(rows) -> "callable":
    """Вес товара: продажи в штуках, а если штук в выгрузке нет ни у кого — продажи в рублях.

    В «Продажах» бывает только «Сумма продажи, руб» — тогда топ и середина считаются по рублям,
    иначе защита топа от снижения не срабатывает ни у одного товара.
    """
    if any((r.sales_qty or 0) > 0 for r in rows):
        return lambda r: r.sales_qty or 0
    return lambda r: r.sales or 0


def weight_classes(rows, opt: RecOptions) -> dict[str, str]:
    """Код -> top / mid / tail по накопленной доле продаж в шт (или в рублях, если штук нет)."""
    w = sales_weight(rows)
    sold = sorted(((w(r), r.code) for r in rows if w(r) > 0), reverse=True)
    total = sum(q for q, _ in sold)
    result = {r.code: "tail" for r in rows}
    acc = 0.0
    for qty, code in sold:
        share_before = acc / total * 100 if total else 100
        acc += qty
        if share_before < opt.top_share:
            result[code] = "top"
        elif share_before < opt.mid_share:
            result[code] = "mid"
    return result


def target_for(prices: dict[str, float | None], percents: dict[str, float], order: list[str], mode: str) -> tuple[float | None, str]:
    """(цель, по кому). Порядок order — как в позиционировании, первый элемент главный.

    first — только он. Нет его цены — цели нет, следующие не подставляются.
    main — он, а если цены нет, следующий по списку.
    """
    def goal(key: str) -> float:
        return prices[key] * (1 + percents.get(key, 0) / 100)

    if mode == "first":
        if not order or not prices.get(order[0]):
            return None, ""
        return goal(order[0]), order[0]
    targets = [(k, goal(k)) for k in order if prices.get(k)]
    if not targets:
        return None, ""
    if mode == "avg":
        return sum(t for _, t in targets) / len(targets), "среднее"
    if mode == "min":
        k, t = min(targets, key=lambda x: x[1])
        return t, k
    return targets[0][1], targets[0][0]


def _round(value: float, step: float) -> float:
    """Округление «половина вверх»: 102,5 ₽ при шаге 1 ₽ становится 103."""
    from decimal import Decimal, ROUND_HALF_UP

    quantum = Decimal("0.01") if not step or step <= 0 else Decimal(str(step))
    amount = Decimal(format(float(value), ".10f"))
    return float((amount / quantum).quantize(Decimal("1"), rounding=ROUND_HALF_UP) * quantum)


# Топ продаж вниз почти не двигаем: это объём группы. 0.02 = 2%.
TOP_DOWN_CAP = 0.02


def abc_key(abc: str) -> str:
    """AA и A — одна полка. Кириллическую «А» приводим к латинской."""
    return (abc or "").strip().upper().replace("А", "A").replace(" ", "")


def is_core_abc(abc: str) -> bool:
    return abc_key(abc) in {"A", "AA"}


def recommend(cur: float | None, sbs: float | None, target: float | None, segment: str, weight: str,
              opt: RecOptions, who: str = "", raise_cap: float | None = None, abc: str = "") -> Rec:
    if not cur:
        return Rec(None, None, "нет Кат. А", "none")
    factor = {"top": opt.top_factor, "mid": opt.mid_factor}.get(weight, 1.0)
    up = opt.up.get(segment, opt.up.get(UNKNOWN, 5.0)) * factor / 100
    down = opt.down * factor / 100
    if weight == "top":
        down = 0.0 if is_core_abc(abc) else min(down, TOP_DOWN_CAP)
    # raise_cap — доля. 0: оверсток и распродажа, цену не поднимаем. 0.02: Европа и Америка.
    if raise_cap is not None:
        up = min(up, max(raise_cap, 0.0))
    floor = sbs / (1 - opt.min_vr / 100) if sbs and opt.min_vr < 100 else None
    notes = []
    if weight == "top":
        if is_core_abc(abc):
            notes.append("топ продаж, категория A — цену не снижаем")
        else:
            notes.append("топ продаж — вниз не больше 2%")
        if factor != 1:
            notes.append(f"вверх ×{opt.top_factor:g}")
    elif weight == "mid" and factor != 1:
        notes.append(f"середина продаж — лимиты ×{opt.mid_factor:g}")

    if target:
        gap = 1 - cur / target  # как «Откл» в отчёте: > 0 — мы дешевле цели
        if abs(gap) * 100 <= opt.tolerance:
            price, kind = cur, "keep"
            reason = f"в пределах позиционирования ({who})" if who else "в пределах позиционирования"
        else:
            change = target / cur - 1
            if change < 0 and down <= 0:
                price, kind = cur, "keep"
                reason = "топ продаж, категория A — не снижаем"
                if who:
                    reason += f" ({who})"
            else:
                limited = max(-down, min(up, change))
                payback = ""
                if limited < 0 and opt.max_lift > 0 and sbs and cur:
                    # Снижение на p при марже m окупается, если продажи вырастут на p / (m − p).
                    m = 1 - sbs / cur
                    lift = opt.max_lift / 100
                    cap = lift * m / (1 + lift) if m > 0 else 0.0
                    if -limited > cap:
                        need = -limited / (m + limited) if m + limited > 0 else None
                        limited = -cap
                        payback = (f"снижение ограничено окупаемостью: для −{abs(change) * 100:.0f}% "
                                   f"нужно {'+' + format(need * 100, '.0f') + '% продаж' if need else 'продавать в разы больше'}")
                if abs(limited) * 100 < max(opt.tolerance, 0.5) and payback:
                    price, kind, reason = cur, "keep", payback
                else:
                    price = cur * (1 + limited)
                    kind = "up" if limited > 0 else "down"
                    reason = f"к цели {who}".strip()
                    if payback:
                        reason += f"; {payback}"
                    elif limited != change:
                        reason += f", ограничено {'+' if limited > 0 else '−'}{abs(limited) * 100:.1f}%"
                        if limited > 0:
                            reason += f" ({segment.lower()} сегмент)"
    else:
        price, kind, reason = cur, "keep", "нет цены конкурента"

    if floor and price < floor:
        if price < cur:
            price = min(cur, floor)
            kind = "keep" if price >= cur else "down"
            reason += f"; не ниже ВР {opt.min_vr:g}%"
        else:
            raised = min(floor, cur * (1 + up))
            if raised > price:
                price = raised
                kind = "vr"
                reason = f"ВР ниже {opt.min_vr:g}% — поднимаем" + (", упёрлись в лимит" if raised < floor else "")
    price = _round(price, opt.rounding)
    if abs(price - cur) < max(opt.rounding, 0.01) / 2:
        price = round(cur, 2)
        if kind in ("up", "down", "vr"):
            kind = "keep"
    if notes and kind != "keep":
        reason += "; " + "; ".join(notes)
    return Rec(price, price / cur - 1, reason, kind)
