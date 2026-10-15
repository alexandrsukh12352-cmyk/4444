"""Проверка папки: что нашлось, что с чем сопоставилось, где расходятся ПСТ и АЦПКД.

Запуск: python Программа/pricing/check.py [папка или файлы] — по умолчанию папка «Входящие».
"""

from __future__ import annotations

import datetime as dt
import os
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path

APP = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(APP))

from build_crosses import builtin_positioning, load_settings  # noqa: E402
from pricing import style  # noqa: E402
from pricing.sources import KIND_TITLES, Sources, discount_for, norm_article, norm_text, scan  # noqa: E402

ROOT = APP.parent
DEFAULT_INPUT = ROOT / "Входящие"
DEFAULT_OUTPUT = ROOT / "Результат"
EXPECTED = ["pst", "sales", "sbs", "transit", "acp", "crosses", "discounts"]
TOLERANCE = 0.01


@dataclass
class Row:
    """Сведения по одному товару ПСТ для проверки."""

    item: object
    sales: float | None
    stock: float | None
    sbs: float | None
    discount: float | None
    cat_a: float | None
    acp: object | None


def collect(src: Sources) -> list[Row]:
    rows = []
    for p in src.pst:
        stock, sbs = src.sbs.get(p.code, (None, None))
        discount = discount_for(src, p.subgroup, p.order_group, p.brand)
        cat_a = p.mo_rub * (1 - discount / 100) if p.mo_rub and discount is not None else None
        rows.append(Row(p, src.sales.get(p.code), stock, sbs, discount, cat_a, src.acp.get(p.code)))
    return rows


def differs(a, b) -> bool:
    if a is None or b is None:
        return False
    if a == b == 0:
        return False
    return abs(a - b) > TOLERANCE * max(abs(a), abs(b))


def build_report(src: Sources, out: Path) -> dict:
    positioning = builtin_positioning(load_settings(None))
    rows = collect(src)
    wb = style.new_workbook()
    stats = {}

    # --- Файлы
    ws = wb.active
    ws.title = "Файлы"
    table = []
    for kind in EXPECTED + [k for k in src.found if k not in EXPECTED]:
        locs = src.found.get(kind) or []
        if not locs:
            table.append([KIND_TITLES[kind], "не найден", None, "", "нет"])
            continue
        loc = locs[0]
        cols = "; ".join(f"{k}: {v}" for k, v in loc.describe()) if loc.cols else ""
        table.append([KIND_TITLES[kind], loc.label, len(loc.header.data), cols, "да"])
    for s in src.skipped:
        table.append(["—", s, None, "", "пропущен"])
    style.write_table(
        ws,
        "Что нашлось в папке",
        ["Документ", "Где найден", "Строк", "Распознанные столбцы", "Найден"],
        table,
        widths=[22, 45, 10, 120, 10],
        status_col=5,
    )

    # --- Сопоставление по подгруппам
    ws = wb.create_sheet("Сопоставление")
    by_sub: dict[tuple[str, str], Counter] = defaultdict(Counter)
    for r in rows:
        key = (r.item.brand, r.item.subgroup)
        c = by_sub[key]
        c["all"] += 1
        c["sales"] += r.sales is not None
        c["sbs"] += r.sbs is not None
        c["acp"] += r.acp is not None
        c["discount"] += r.discount is not None
        c["mo"] += bool(r.item.mo_rub)
        c["pos"] += bool(positioning.resolve(r.item.brand, r.item.group, r.item.subgroup)[2])
    table = [
        [b, s, c["all"], c["mo"], c["discount"], c["sales"], c["sbs"], c["acp"], "да" if c["pos"] else "нет"]
        for (b, s), c in sorted(by_sub.items(), key=lambda kv: -kv[1]["all"])
    ]
    style.write_table(
        ws,
        "Товары ПСТ: что к ним нашлось (по коду товара)",
        ["ТМ", "Подгруппа", "Товаров в ПСТ", "Есть МО", "Есть скидка", "Есть продажи", "Есть СБС", "Есть в АЦПКД", "Позиционирование"],
        table,
        widths=[16, 34, 14, 12, 12, 13, 12, 13, 17],
        status_col=9,
    )

    # --- Кроссы по конкурентам
    ws = wb.create_sheet("Кроссы")
    cover: dict[tuple, Counter] = defaultdict(Counter)
    for r in rows:
        p = r.item
        _, _, comps, _ = positioning.resolve(p.brand, p.group, p.subgroup)
        svod = src.crosses.get((norm_text(p.brand), norm_article(p.article)), {})
        for comp in comps:
            c = cover[(p.brand, p.subgroup, comp.order, comp.name)]
            c["all"] += 1
            crosses = svod.get(comp.key, [])
            yes = [x for x in crosses if x.yes]
            c["yes"] += bool(yes)
            c["no_only"] += bool(crosses) and not yes
            c["none"] += not crosses
            acp_crosses = r.acp.competitors.get(comp.key, []) if r.acp else []
            if not acp_crosses and r.acp:
                acp_crosses = next((v for k, v in r.acp.competitors.items() if norm_text(k) == comp.key), [])
            priced = {norm_article(x.article) for x in acp_crosses if x.min_price}
            c["acp_cross"] += bool(acp_crosses)
            c["acp_price"] += bool(priced)
            c["yes_priced"] += any(x.cross_key in priced for x in yes)
            c["no_priced"] += not any(x.cross_key in priced for x in yes) and any(
                x.cross_key in priced for x in crosses if not x.yes
            )
    table = []
    for (b, s, n, comp), c in sorted(cover.items()):
        total = c["all"] or 1
        table.append(
            [b, s, n, comp, c["all"], c["yes"], c["no_only"], c["none"], c["acp_cross"], c["acp_price"],
             c["yes_priced"], c["no_priced"], c["acp_price"] / total]
        )
    style.write_table(
        ws,
        "Кроссы на конкурентов из позиционирования",
        ["ТМ", "Подгруппа", "№", "Конкурент", "Товаров", "Кросс «Да» (свод)", "Только «Нет»", "Нет кроссов",
         "Кросс в АЦПКД", "Цена в АЦПКД", "Цена у кросса «Да»", "Цена только у «Нет»", "Покрытие ценой"],
        table,
        widths=[14, 30, 5, 13, 10, 13, 12, 11, 12, 12, 13, 14, 12],
        percent_cols=[13],
    )

    # --- ПСТ против АЦПКД
    ws = wb.create_sheet("ПСТ и АЦПКД")
    diffs = Counter()
    detail = []
    for r in rows:
        a = r.acp
        if not a:
            continue
        diffs["both"] += 1
        checks = [
            ("МО", r.item.mo_rub, a.mo_price),
            ("Кат. А", r.cat_a, a.cat_a_price),
            ("СБС", r.sbs, a.avg_cost),
            ("Остаток", r.stock, a.stock),
            ("ABC", r.item.abc or "—", a.abc or "—"),
        ]
        for name, ours, theirs in checks:
            bad = ours != theirs if name == "ABC" else differs(ours, theirs)
            if bad:
                diffs[name] += 1
                detail.append([r.item.code, r.item.article, r.item.subgroup, name, ours, theirs])
    summary = [[name, diffs["both"], diffs[name]] for name in ("МО", "Кат. А", "СБС", "Остаток", "ABC")]
    style.write_table(
        ws,
        f"Расхождения больше {TOLERANCE:.0%}: наши выгрузки (ПСТ, СБС, скидки) против АЦПКД",
        ["Показатель", "Товаров в обоих", "Расходится"],
        summary,
        widths=[16, 16, 14, 30, 16, 16],
    )
    ws.append([])
    start = ws.max_row + 1
    style.write_table(
        ws,
        "Где именно",
        ["Код товара", "Артикул", "Подгруппа", "Показатель", "Наши данные", "АЦПКД"],
        detail,
        start_row=start,
        number_cols=[5, 6],
    )

    stats.update(items=len(rows), diffs=dict(diffs), files={k: [l.label for l in v] for k, v in src.found.items()})
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return stats


def main(argv=None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass
    args = [Path(a) for a in (argv if argv is not None else sys.argv[1:]) if a != "--no-open"]
    no_open = "--no-open" in (argv if argv is not None else sys.argv[1:])
    src = scan(args or [DEFAULT_INPUT])
    print()
    for kind in EXPECTED:
        locs = src.found.get(kind)
        print(f"  {KIND_TITLES[kind]:<22} {locs[0].label if locs else '— не найден'}")
    for s in src.skipped:
        print(f"  ! {s}")
    for w in src.warnings:
        print(f"  ! {w}")
    if not src.pst:
        print("\nПСТ не найден — проверять нечего.\n")
        return 1
    out = DEFAULT_OUTPUT / f"Проверка_папки_{dt.datetime.now():%Y-%m-%d_%H-%M-%S}.xlsx"
    stats = build_report(src, out)
    print(f"\nТоваров в ПСТ: {stats['items']}. Отчёт о проверке: {out}\n")
    if not no_open and hasattr(os, "startfile"):
        os.startfile(out)  # type: ignore[attr-defined]
    return 0


if __name__ == "__main__":
    sys.exit(main())
