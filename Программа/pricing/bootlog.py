"""Стадия загрузки. Окно опрашивает её, пока идёт чтение папки.

Отдельный замок: опрос не ждёт, пока освободится разбор Excel.

Шаги программы (step) складываются в крупные этапы (PHASES): окно загрузки показывает их
списком по порядку — что сделано, сколько заняло, что идёт сейчас и сколько осталось файлов.
Этапы, которые не понадобились (всё взято из кэша), помечаются «не понадобилось».
"""

from __future__ import annotations

import re
import threading
import time

PHASES = {
    "folder": "Папка и сохранённый разбор",
    "files": "Чтение файлов Excel",
    "prices": "Распознавание прайсов",
    "groups": "Группы, кроссы и цены",
    "analytics": "Расчёт изменений цен",
    "export": "Запись книги Excel",
    "compact": "Архив старых дат",
}
PLANS = {
    "scan": ["folder", "files", "prices", "groups"],
    "prices": ["files", "prices"],
    "analytics": ["files", "prices", "analytics"],
    "export": ["files", "prices", "export"],
    "compact": ["files", "prices", "compact"],
}
_STAGE_PHASE = {
    "Папка": "folder", "Кэш": "folder", "Чтение папок": "files", "Читаю файл": "files",
    "Читаю новые файлы параллельно": "files", "Читаю прайсы прошлых дат": "files", "Читаю прайс прошлой даты": "files",
    "Настройки прайсов": "prices", "Ищу шапку из настроек": "prices", "Столбцы поставщика": "prices",
    "Разбираю прайс": "prices", "Сохраняю столбцы": "prices", "Обновляю список прайсов": "prices",
    "Прайсы": "groups", "Кроссы": "groups", "Отчёт": "groups",
    "Считаю изменения": "analytics", "Пишу Excel": "export", "Сжимаю": "compact",
}
_PROGRESS = re.compile(r"^(\d+) из (\d+)")

_lock = threading.Lock()
_state = {
    "stage": "Жду окно",
    "detail": "",
    "started": 0.0,
    "stage_at": 0.0,
    "error": "",
    "done": True,
    "log": [],
    "plan": [],
    "phases": {},
    "current": "",
}


def begin(plan: str = "scan") -> None:
    now = time.time()
    with _lock:
        _state.update(stage="Старт", detail="", started=now, stage_at=now, error="", done=False, log=[],
                      plan=list(PLANS.get(plan, PLANS["scan"])), phases={}, current="")


def _enter(phase: str, now: float) -> dict:
    phases = _state["phases"]
    current = _state["current"]
    if current and current != phase and current in phases:
        phases[current]["end"] = now
    item = phases.get(phase)
    if item is None:
        item = phases[phase] = {"start": now, "end": None, "steps": 0, "detail": "", "progress": None}
        if phase not in _state["plan"]:
            _state["plan"].append(phase)
    else:
        item["end"] = None
    _state["current"] = phase
    return item


def step(stage: str, detail: str = "") -> None:
    now = time.time()
    with _lock:
        last = _state["log"][-1] if _state["log"] else None
        if not last or last["stage"] != stage or last["detail"] != detail:
            _state["log"].append({"stage": stage, "detail": detail})
            _state["log"] = _state["log"][-8:]
        _state["stage"] = stage
        _state["detail"] = detail
        _state["stage_at"] = now
        _state["error"] = ""
        _state["done"] = False
        phase = _STAGE_PHASE.get(stage)
        if phase:
            item = _enter(phase, now)
            item["steps"] += 1
            item["detail"] = f"{stage}: {detail}" if detail else stage
            m = _PROGRESS.match(detail or "")
            item["progress"] = [int(m.group(1)), int(m.group(2))] if m else item["progress"]


def fail(message: str) -> None:
    with _lock:
        _state["error"] = message
        _state["done"] = True
        _state["detail"] = message


def finish() -> None:
    now = time.time()
    with _lock:
        current = _state["current"]
        if current and current in _state["phases"]:
            _state["phases"][current]["end"] = now
        _state["done"] = True
        _state["stage"] = "Готово"
        _state["detail"] = ""
        _state["error"] = ""


def _phase_rows(now: float) -> list[dict]:
    rows = []
    plan = _state["plan"]
    phases = _state["phases"]
    reached = max((plan.index(k) for k in phases if k in plan), default=-1)
    for i, key in enumerate(plan):
        item = phases.get(key)
        if item is None:
            state = "skip" if (i < reached or _state["done"]) else "wait"
            rows.append({"key": key, "title": PHASES.get(key, key), "state": state})
            continue
        end = item["end"]
        if _state["error"] and key == _state["current"]:
            state = "error"
        elif key == _state["current"] and not _state["done"]:
            state = "active"
        else:
            state = "done"
        rows.append({
            "key": key, "title": PHASES.get(key, key), "state": state,
            "elapsed": round((end or now) - item["start"], 1), "detail": item["detail"],
            "progress": item["progress"], "steps": item["steps"],
        })
    return rows


def snapshot() -> dict:
    now = time.time()
    with _lock:
        started = _state["started"] or now
        stage_at = _state["stage_at"] or started
        return {
            "stage": _state["stage"],
            "detail": _state["detail"],
            "error": _state["error"],
            "done": _state["done"],
            "started": _state["started"],
            "elapsed": round(now - started, 1),
            "stage_elapsed": round(now - stage_at, 1),
            "log": list(_state["log"]),
            "phases": _phase_rows(now),
        }
