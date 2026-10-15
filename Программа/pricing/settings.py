"""Настройки программы: курсы валют, ГЗ, скидки, марки авто, выбранные режимы.

Зашитые значения лежат в defaults.json и cars.json рядом с модулем. Всё, что пользователь
поменял в окне «Настройки», сохраняется в «Программа/настройки.json» и перекрывает
зашитое. Файл можно удалить — тогда вернутся значения по умолчанию.
"""

from __future__ import annotations

import json
import shutil
from pathlib import Path

from .cars import load_default as cars_default
from .sources import DATA_DIRS, norm_text  # noqa: F401

DEFAULTS_FILE = Path(__file__).with_name("defaults.json")
PROGRAM = Path(__file__).resolve().parent.parent
ROOT = PROGRAM.parent
USER_FILE = PROGRAM / "настройки.json"
OLD_USER_FILES = (ROOT / "Настройки" / "настройки.json", ROOT / "настройки.json")
OUTPUT_DIR = "Результат"

DEFAULT_OPTIONS = {
    "folder": "",
    "output": "Результат",
    "base": "pst",
    "sources": {},
    "filters": {},
    "price_source": "both",
    "monitoring_only": True,
    "in_stock_only": False,
    "cross_split": "none",
    "subgroup_mode": "sheets",
    "suggest_threshold": 0.3,
    "open_result": True,
}


def _read(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


_LOADS: dict[tuple, "Settings"] = {}
_TOKEN = iter(range(1, 10**12))


def _drop_loads() -> None:
    _LOADS.clear()


class Settings:
    def __init__(self, data: dict):
        self.data = data

    @classmethod
    def load(cls, user_file: Path | None | object = ...) -> "Settings":
        """user_file=None — только зашитые значения; по умолчанию — с «настройки.json»."""
        if user_file is ...:
            user_file = USER_FILE
            if not user_file.exists():
                for old in OLD_USER_FILES:
                    if old.exists():
                        user_file.parent.mkdir(parents=True, exist_ok=True)
                        shutil.move(str(old), user_file)
                        break
        key = None
        if user_file:
            try:
                st = user_file.stat()
                key = (str(user_file.resolve()), st.st_mtime_ns, st.st_size)
            except OSError:
                key = (str(user_file), 0, 0)
            hit = _LOADS.get(key)
            if hit is not None:
                return hit
        data = _read(DEFAULTS_FILE)
        data["options"] = dict(DEFAULT_OPTIONS)
        if user_file:
            user = _read(user_file)
            for item in ("gz", "discounts", "price_profiles", "price_files", "cars", "positioning", "brand_aliases", "city_codes"):
                if item in user:
                    data[item] = user[item]
            if "currency_rates" in user:
                data["currency_rates"] = {**{k: "" for k in data.get("currency_rates") or {}}, **user["currency_rates"]}
            data["options"].update(user.get("options") or {})
        loaded = cls(data)
        if user_file:
            from .prices import set_brand_aliases

            set_brand_aliases(data.get("brand_aliases"))
            from .prices import set_city_codes

            set_city_codes(data.get("city_codes"))
        if key is not None:
            _LOADS[key] = loaded
            if len(_LOADS) > 8:
                _LOADS.pop(next(iter(_LOADS)))
        return loaded

    def save(self, user_file: Path | None = None) -> None:
        user_file = user_file or USER_FILE
        user_file.parent.mkdir(parents=True, exist_ok=True)
        user_file.write_text(json.dumps(self.data, ensure_ascii=False, indent=1), encoding="utf-8")
        self._drop_derived()
        _drop_loads()

    def _drop_derived(self) -> None:
        for name in ("_rates", "_gz", "_discounts", "_cars_rows", "_cars_parser", "_pos", "_calc_gen"):
            self.__dict__.pop(name, None)

    def calc_token(self) -> int:
        """Меняется, когда настройки перечитывают с диска. Кэш строк на него опирается."""
        gen = self.__dict__.get("_calc_gen")
        if gen is None:
            gen = next(_TOKEN)
            self._calc_gen = gen
        return gen

    # ------------------------------------------------------------ доступ

    @property
    def options(self) -> dict:
        return self.data["options"]

    @property
    def rates(self) -> dict[str, float]:
        cached = self.__dict__.get("_rates")
        if cached is None:
            cached = {norm_text(k): float(v) for k, v in (self.data.get("currency_rates") or {}).items() if v not in (None, "")}
            self._rates = cached
        return cached

    def rate(self, currency: str) -> float | None:
        """Курс валюты из столбца ПСТ «Закупочная цена / Вал.»; рубли — 1."""
        cur = norm_text(currency).strip("-– ")
        if not cur:
            return None
        if cur.rstrip(".") in ("РУБ", "RUB", "RUR"):
            return 1.0
        return self.rates.get(cur)

    @property
    def price_profiles(self) -> dict:
        return self.data.setdefault("price_profiles", {})

    @property
    def price_files(self) -> dict:
        return self.data.setdefault("price_files", {})

    @property
    def brand_aliases(self) -> list:
        """[[как бренд написан у поставщика или в кроссах, бренд конкурента], ...]."""
        return self.data.setdefault("brand_aliases", [])

    @property
    def cars(self) -> list[list]:
        cached = self.__dict__.get("_cars_rows")
        if cached is None:
            rows = self.data.get("cars")
            cached = rows if rows else cars_default()
            self._cars_rows = cached
        return cached

    @property
    def gz(self) -> dict[str, float]:
        cached = self.__dict__.get("_gz")
        if cached is None:
            cached = {norm_text(name): float(coef) for name, coef in self.data.get("gz") or [] if coef not in (None, "")}
            self._gz = cached
        return cached

    @property
    def discounts(self) -> dict[tuple[str, str, str], float]:
        cached = self.__dict__.get("_discounts")
        if cached is not None:
            return cached
        result = {}
        for row in self.data.get("discounts") or []:
            *_, subgroup, order_group, brand, value = row
            if value in (None, ""):
                continue
            result.setdefault((norm_text(subgroup), norm_text(order_group), norm_text(brand)), float(value))
        self._discounts = result
        return result
