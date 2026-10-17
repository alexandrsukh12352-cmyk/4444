"""Кроссы через ОЕМ: цена конкурента для товара, у которого своего кросса нет.

У наших марок много товаров-двойников: диск ADVICS, диск AVANTECH и диск DOUBLE FORCE на одну
машину имеют один номер ОЕМ. Кроссы в «Кроссы свод» заведены не у всех двойников одинаково:
у AVANTECH CF1103 есть кросс «Да» на LYNX, а у DOUBLE FORCE DFC10112 — нет. Раз оба наших товара —
одна и та же деталь по ОЕМ, подтверждённый кросс двойника годится и этому товару.

Это кандидат, а не кросс: он помечается «по ОЕМ» с артикулом двойника. Режимы:
  off  — не искать;
  show — искать и показывать, сколько товаров получили бы цену (цель не меняется);
  use  — брать цену кандидата, если своей цены у конкурента нет.

Защита от ложных двойников:
  * двойник — того же вида: та же группа в ПСТ, а если её нет — те же первые два слова наименования;
  * номер ОЕМ, общий больше чем у MAX_SHARE наших товаров, не связывает (болт, гайка, общий корпус);
  * берутся только кроссы «Да» двойника.
"""

from __future__ import annotations

from dataclasses import dataclass

from .excel import clean
from .sources import norm_article, norm_text

MAX_SHARE = 8
MODES = {"off": "Не искать", "show": "Показывать", "use": "Брать в цену"}


@dataclass
class Twin:
    brand: str
    article: str
    via: str  # общий номер ОЕМ


def _kind(name: str, group: str) -> str:
    if group:
        return "g:" + norm_text(group)
    words = norm_text(name).split()
    return "n:" + " ".join(words[:2])


def build_twins(oem_rows, pst) -> dict[tuple[str, str], list[Twin]]:
    """(марка, артикул) нашего товара -> двойники по ОЕМ того же вида."""
    group_of = {p.code: p.group for p in pst or [] if p.code}
    group_by_key = {(norm_text(p.brand), norm_article(p.article)): p.group for p in pst or []}
    items = []
    for it in oem_rows or []:
        if not it.brand or not it.article:
            continue
        key = (norm_text(it.brand), norm_article(it.article))
        group = group_of.get(it.code) or group_by_key.get(key) or ""
        items.append((key, _kind(it.name, group), it.oems, clean(it.brand), clean(it.article)))
    by_oem: dict[str, list] = {}
    for item in items:
        for oem in item[2]:
            by_oem.setdefault(oem, []).append(item)
    out: dict[tuple[str, str], list[Twin]] = {}
    for key, kind, oems, _b, _a in items:
        seen = {key}
        twins = out.setdefault(key, [])
        for oem in oems:
            group = by_oem.get(oem) or []
            if len({g[0] for g in group}) > MAX_SHARE:
                continue
            for other_key, other_kind, _o, brand, article in group:
                if other_key in seen or other_kind != kind:
                    continue
                seen.add(other_key)
                twins.append(Twin(brand, article, oem))
    return {k: v for k, v in out.items() if v}


def twins_for(src) -> dict:
    """Двойники считаются один раз на набор строк ОЕМ и ПСТ."""
    oem_rows = getattr(src, "oem_rows", None) or getattr(src, "oem", None) or []
    pst = getattr(src, "pst", None) or []
    token = (id(oem_rows), len(oem_rows), id(pst), len(pst))
    held = getattr(src, "_oem_twins", None)
    if held and held[0] == token:
        return held[1]
    twins = build_twins(oem_rows, pst)
    try:
        src._oem_twins = (token, twins)
    except AttributeError:
        pass
    return twins


def candidates(src, brand: str, article: str, comp_key: str, svod_comp) -> list[tuple]:
    """Кроссы «Да» двойников на этого конкурента: [(CrossRow, двойник)]."""
    twins = twins_for(src).get((norm_text(brand), norm_article(article))) or []
    out, seen = [], set()
    for t in twins:
        rows = svod_comp(src.crosses.get((norm_text(t.brand), norm_article(t.article)), {}), comp_key)
        for row in rows:
            if row.yes and row.cross_key not in seen:
                seen.add(row.cross_key)
                out.append((row, t))
    return out
