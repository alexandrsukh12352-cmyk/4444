"""Стадия загрузки. Окно опрашивает её, пока идёт чтение папки.

Отдельный замок: опрос не ждёт, пока освободится разбор Excel.
"""

from __future__ import annotations

import threading
import time

_lock = threading.Lock()
_state = {
    "stage": "Жду окно",
    "detail": "",
    "started": 0.0,
    "stage_at": 0.0,
    "error": "",
    "done": True,
    "log": [],
}


def begin() -> None:
    now = time.time()
    with _lock:
        _state.update(stage="Старт", detail="", started=now, stage_at=now, error="", done=False, log=[])


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


def fail(message: str) -> None:
    with _lock:
        _state["error"] = message
        _state["done"] = True
        _state["detail"] = message


def finish() -> None:
    with _lock:
        _state["done"] = True
        _state["stage"] = "Готово"
        _state["detail"] = ""
        _state["error"] = ""


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
        }
