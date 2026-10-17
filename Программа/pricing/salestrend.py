"""Динамика продаж по выгрузке «Продажи за 3 месяца» с разбивкой по дням, неделям или месяцам.

Период выгрузки раскладывается по дням (месяц — поровну на его дни), дальше считаются окна:
последние 4 недели против 4 недель до них. Так же — продажи до и после изменения цены.
"""

from __future__ import annotations

import datetime as dt
import re

WINDOW = 28
MIN_BASE = 3  # меньше трёх штук в базовом окне — динамика ничего не значит


def period_end(period: str) -> dt.date | None:
    """«01.08.2026–16.10.2026» -> 16.10.2026."""
    found = re.findall(r"(\d{2})\.(\d{2})\.(\d{4})", period or "")
    if not found:
        return None
    d, m, y = found[-1]
    try:
        return dt.date(int(y), int(m), int(d))
    except ValueError:
        return None


def _amount(series, start: dt.date, end: dt.date, use_qty: bool = True) -> float:
    """Сколько продано с start по end включительно; период частично в окне — пропорционально дням."""
    total = 0.0
    for a, b, qty, rub in series:
        p0, p1 = dt.date.fromisoformat(a), dt.date.fromisoformat(b)
        lo, hi = max(p0, start), min(p1, end)
        if lo > hi:
            continue
        share = ((hi - lo).days + 1) / ((p1 - p0).days + 1)
        total += (qty if use_qty else rub) * share
    return total


def _use_qty(series) -> bool:
    return any(q for _a, _b, q, _r in series)


def trend(series, end: dt.date | None, window: int = WINDOW) -> tuple[float, float, float | None]:
    """(последние 4 недели, 4 недели до них, изменение долей). Изменение — None, если база меньше трёх."""
    if not series or end is None:
        return 0.0, 0.0, None
    q = _use_qty(series)
    recent = _amount(series, end - dt.timedelta(days=window - 1), end, q)
    before = _amount(series, end - dt.timedelta(days=2 * window - 1), end - dt.timedelta(days=window), q)
    if before < (MIN_BASE if q else 1):
        return recent, before, None
    return recent, before, recent / before - 1


def around(series, day: dt.date, end: dt.date | None, window: int = WINDOW) -> tuple[float, float, int] | None:
    """Продажи в день до и после даты изменения цены: (до, после, дней после). None — данных мало."""
    if not series or end is None or day > end:
        return None
    after_days = min(window, (end - day).days + 1)
    if after_days < 14:
        return None
    first = min(dt.date.fromisoformat(a) for a, *_ in series)
    q = _use_qty(series)
    before_days = min(window, (day - first).days)
    if before_days < 14:
        return None
    before = _amount(series, day - dt.timedelta(days=before_days), day - dt.timedelta(days=1), q) / before_days
    after = _amount(series, day, day + dt.timedelta(days=after_days - 1), q) / after_days
    return before, after, after_days


def effect_text(series, change: dict | None, end: dt.date | None) -> str:
    """«продажи после изменения +12% в день (28 дн.)» — для столбца «Цену меняли»."""
    if not change:
        return ""
    try:
        day = dt.date.fromisoformat(change.get("date", ""))
    except ValueError:
        return ""
    res = around(series, day, end)
    if not res:
        return ""
    before, after, days = res
    if before <= 0:
        return f"до изменения продаж не было, после — {after * 7:.1f} в неделю".replace(".", ",")
    pct = f"{(after / before - 1) * 100:+.0f}"
    return f"продажи после изменения {pct}% в день ({days} дн.)"


def group_trend(src, rows) -> dict | None:
    """Продажи группы за 4 недели к 4 неделям до них (по всем товарам отчёта)."""
    days = getattr(src, "sales3_days", None) or {}
    end = period_end(getattr(src, "sales3_period", ""))
    if not days or end is None:
        return None
    recent = before = 0.0
    qty = any(_use_qty(days.get(r.code) or []) for r in rows[:200])
    for r in rows:
        series = days.get(r.code)
        if not series:
            continue
        recent += _amount(series, end - dt.timedelta(days=WINDOW - 1), end, qty)
        before += _amount(series, end - dt.timedelta(days=2 * WINDOW - 1), end - dt.timedelta(days=WINDOW), qty)
    if before <= 0:
        return None
    up = sum(1 for r in rows if (getattr(r, "sales_trend", None) or 0) >= 0.2)
    down = sum(1 for r in rows if getattr(r, "sales_trend", None) is not None and r.sales_trend <= -0.2)
    return {"recent": recent, "before": before, "change": recent / before - 1, "unit": "шт" if qty else "₽",
            "end": end.strftime("%d.%m"), "up": up, "down": down}
