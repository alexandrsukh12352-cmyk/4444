"""Вкладка «Диагностика»: номер сборки и ожидаемый столбец смежной марки."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "Программа"))

from pricing.buildinfo import BUILD, column_expectation, program_marks  # noqa: E402


def test_this_build_writes_the_sister_column_and_the_brief():
    marks = {m["name"]: m["ok"] for m in program_marks()}
    assert marks["Столбец смежной марки"]
    assert marks["Место столбца"]
    assert marks["Текстовый разбор"]
    assert BUILD


def test_filters_expect_double_force_after_competitors():
    got = column_expectation("AVANTECH", "Фильтры автомобильные")
    assert got["header"] == "DOUBLE FORCE"
    assert got["place"] == "после цен конкурентов"
    assert column_expectation("ADVICS", "Тормозные диски")["header"] == "AVANTECH"
    assert column_expectation("ADVICS", "Тормозные суппорты")["header"] == ""
