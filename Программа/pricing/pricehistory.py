"""История цен: что и когда менялось в ПСТ и совпало ли это с нашей рекомендацией.

Хранить старые ПСТ не нужно. Каждый раз, когда программа открывает группу, она сравнивает МО товара
с прошлым разом. Изменилось — пишет дату, было, стало и чьё это изменение: «по рекомендации»,
если новая Кат. А совпала с рекомендацией, которую программа выдала до этого (±1%), иначе
«вне программы». После изменения цену не трогаем N дней (по умолчанию 90), чтобы увидеть, как
ответили продажи. Продажи за 3 месяца запоминаются по датам — для сверки «что изменили → что стало».

Файл «история цен.json» лежит рядом с программой и в архив сборки не входит.
"""

from __future__ import annotations

import datetime as dt
import json
from pathlib import Path

NAME = "история цен.json"
CHANGE = 0.005  # меньше 0,5% — округление, а не изменение цены
SAME_AS_REC = 0.01  # новая Кат. А в пределах 1% от рекомендации — это наша рекомендация
KEEP_CHANGES = 6
KEEP_SALES = 12

_memo: dict = {"path": None, "mtime": None, "data": None}


def history_path() -> Path:
    from .settings import PROGRAM

    return PROGRAM / NAME


def _empty() -> dict:
    return {"version": 1, "prices": {}, "changes": {}, "recs": {}, "sales": {}}


def load(path: Path | None = None) -> dict:
    path = path or history_path()
    try:
        mtime = path.stat().st_mtime
    except OSError:
        return _empty()
    if _memo["path"] == path and _memo["mtime"] == mtime and _memo["data"] is not None:
        return _memo["data"]
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            data = _empty()
    except (OSError, ValueError):
        data = _empty()
    for key, value in _empty().items():
        data.setdefault(key, value)
    _memo.update(path=path, mtime=mtime, data=data)
    return data


def save(data: dict, path: Path | None = None) -> None:
    path = path or history_path()
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix(".tmp")
        tmp.write_text(json.dumps(data, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
        tmp.replace(path)
        _memo.update(path=path, mtime=path.stat().st_mtime, data=data)
    except OSError:
        pass


def _day(value: str) -> dt.date | None:
    try:
        return dt.date.fromisoformat(value)
    except (TypeError, ValueError):
        return None


def _ru(day: dt.date) -> str:
    return day.strftime("%d.%m.%Y")


def observe(rows, today: dt.date | None = None, path: Path | None = None) -> int:
    """Сравнивает МО (или Кат. А, если МО нет) с прошлым разом. Возвращает, сколько цен изменилось."""
    today = today or dt.date.today()
    data = load(path)
    prices, changes, recs = data["prices"], data["changes"], data["recs"]
    found = dirty = 0
    for r in rows:
        base = getattr(r, "mo", None) or getattr(r, "cat_a", None)
        if not base:
            continue
        code = str(r.code)
        prev = prices.get(code)
        if prev and prev.get("base") and abs(base / prev["base"] - 1) > CHANGE:
            rec = recs.get(code)
            cat_a = getattr(r, "cat_a", None)
            ours = bool(rec and rec.get("rec") and cat_a and (_day(rec.get("date")) or today) <= today
                        and abs(cat_a / rec["rec"] - 1) <= SAME_AS_REC)
            item = {"date": today.isoformat(), "old": prev.get("cat_a"), "new": cat_a,
                    "old_base": prev["base"], "new_base": base, "by": "rec" if ours else "other"}
            changes[code] = (changes.get(code) or [])[-(KEEP_CHANGES - 1):] + [item]
            found += 1
        if not prev or prev.get("base") != base or prev.get("cat_a") != getattr(r, "cat_a", None):
            prices[code] = {"base": base, "cat_a": getattr(r, "cat_a", None), "seen": today.isoformat()}
            dirty += 1
        sales3 = getattr(r, "sales3", None)
        if sales3 is not None:
            qty = getattr(r, "sales3_qty", None)
            point = [today.isoformat(), qty if qty is not None else None, sales3]
            series = data["sales"].get(code) or []
            if not series or series[-1][0] != point[0]:
                data["sales"][code] = (series + [point])[-KEEP_SALES:]
                dirty += 1
            elif series[-1] != point:
                series[-1] = point
                dirty += 1
    if dirty or found:
        save(data, path)
    return found


def remember_recs(rows, today: dt.date | None = None, path: Path | None = None) -> int:
    """Рекомендации из собранного отчёта: потом по ним видно, чьё было изменение цены."""
    today = today or dt.date.today()
    data = load(path)
    n = 0
    for r in rows:
        rec = getattr(r, "rec", None)
        if not rec or rec.kind not in ("up", "down", "vr") or not rec.price:
            continue
        data["recs"][str(r.code)] = {"date": today.isoformat(), "cur": r.cat_a, "rec": rec.price}
        n += 1
    if n:
        save(data, path)
    return n


def last_change(code: str, path: Path | None = None) -> dict | None:
    items = load(path)["changes"].get(str(code)) or []
    return items[-1] if items else None


def change_text(change: dict | None) -> str:
    """«16.09.2026: +5,0%, по рекомендации» — для отчёта."""
    if not change:
        return ""
    day = _day(change.get("date"))
    old, new = change.get("old"), change.get("new")
    if not (old and new):
        old, new = change.get("old_base"), change.get("new_base")
    pct = f"{(new / old - 1) * 100:+.1f}%".replace(".", ",") if old and new else ""
    who = "по рекомендации" if change.get("by") == "rec" else "вне программы"
    return ", ".join(x for x in (f"{_ru(day)}: {pct}" if day else pct, who) if x)


def annotate(rows, cooldown_days: float = 90, today: dt.date | None = None, path: Path | None = None) -> int:
    """Ставит r.last_change и, если цену меняли недавно, r.frozen — рекомендацию по ней не даём."""
    today = today or dt.date.today()
    changes = load(path)["changes"]
    frozen = 0
    for r in rows:
        items = changes.get(str(r.code)) or []
        change = items[-1] if items else None
        r.last_change = change
        r.frozen = ""
        if not change or not cooldown_days:
            continue
        day = _day(change.get("date"))
        if day is None:
            continue
        until = day + dt.timedelta(days=int(cooldown_days))
        if today < until:
            r.frozen = (f"цену меняли {change_text(change)} — не трогаем до {_ru(until)}, "
                        f"ждём, как ответят продажи")
            frozen += 1
    return frozen
