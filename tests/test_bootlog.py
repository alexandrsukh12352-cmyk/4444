"""Окно загрузки видит шаг и текст ошибки."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "Программа"))

from pricing import bootlog, server  # noqa: E402


def test_loading_window_keeps_the_stage_and_the_error():
    bootlog.begin()
    bootlog.step("Читаю файл", "TissEKB.xls")
    snap = bootlog.snapshot()
    assert snap["done"] is False
    assert snap["stage"] == "Читаю файл" and snap["detail"] == "TissEKB.xls"
    assert snap["log"][-1]["detail"] == "TissEKB.xls"
    bootlog.fail("не удалось прочитать лист")
    failed = server.api_progress({})
    assert failed["done"] is True
    assert failed["error"] == "не удалось прочитать лист"
