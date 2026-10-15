"""Марки и модели авто из столбца «Разновидность» → автопарк и ценовой сегмент.

Интернет не нужен: разбор идёт по справочнику марок и моделей (встроенный — cars.json,
правится во вкладке «Настройки»). В справочнике у марки есть автопарк (японский, европейский…)
и сегмент; у модели можно задать свой сегмент (Land Cruiser — премиум, Solaris — бюджетный)
и пометку «праворульная» (модели только для японского рынка).
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path

CARS_FILE = Path(__file__).with_name("cars.json")
SEGMENTS = ["Бюджетный", "Средний", "Премиум"]
SEGMENT_RANK = {s: i for i, s in enumerate(SEGMENTS)}
UNKNOWN = "Не определён"


@dataclass
class Entry:
    make: str
    model: str  # пусто — запись про марку целиком
    aliases: list[str]
    fleet: str
    segment: str
    rhd: bool = False


@dataclass
class CarInfo:
    makes: list[str]  # «Toyota», «Lexus»
    models: list[str]  # «Toyota Corolla»
    fleet: str  # «Японский», «Японский (праворульный)», «Смешанный: …»
    segment: str  # Премиум / Средний / Бюджетный / Не определён

    @property
    def label(self) -> str:
        if self.models:
            return ", ".join(self.models)
        return ", ".join(self.makes)


_DEFAULT_CARS: list[list] | None = None


def load_default() -> list[list]:
    global _DEFAULT_CARS
    if _DEFAULT_CARS is None:
        _DEFAULT_CARS = json.loads(CARS_FILE.read_text(encoding="utf-8"))
    return _DEFAULT_CARS


def entries_from(rows: list[list]) -> list[Entry]:
    """Строки справочника: [марка, модель, как пишется (через запятую), автопарк, сегмент, праворульная]."""
    result = []
    for row in rows:
        make, model, aliases, fleet, segment, *rest = (list(row) + [""] * 6)[:6]
        make = str(make or "").strip()
        if not make:
            continue
        model = str(model or "").strip()
        words = [a.strip() for a in str(aliases or "").split(",") if a.strip()]
        words.append(f"{make} {model}".strip() if model else make)
        if model and len(model) >= 3 and not model.isdigit():
            words.append(model)
        result.append(Entry(make, model, sorted({_norm(w) for w in words if _norm(w)}, key=len, reverse=True),
                            str(fleet or "").strip(), str(segment or "").strip() or "Средний",
                            str(rest[0]).strip().lower() in ("1", "да", "true", "yes", "+")))
    return result


def _norm(text: str) -> str:
    text = str(text).upper().replace("Ё", "Е")
    text = re.sub(r"[^0-9A-ZА-Я]+", " ", text)
    return f" {text.strip()} " if text.strip() else ""


class CarParser:
    def __init__(self, rows: list[list] | None = None):
        self.entries = entries_from(rows if rows is not None else load_default())
        self.models = sorted(
            (e for e in self.entries if e.model),
            key=lambda e: -max((len(a) for a in e.aliases), default=0),
        )
        self.makes = {e.make.upper(): e for e in self.entries if not e.model}
        self.cache: dict[str, CarInfo] = {}

    def parse(self, variety: str) -> CarInfo:
        if variety in self.cache:
            return self.cache[variety]
        text = _norm(variety)
        hits: list[tuple[int, Entry]] = []
        taken: list[tuple[int, int]] = []

        def find(entry: Entry) -> int | None:
            for alias in entry.aliases:
                start = text.find(alias)
                while start >= 0:
                    span = (start, start + len(alias))
                    if not any(a < span[1] and span[0] < b for a, b in taken):
                        taken.append(span)
                        return start
                    start = text.find(alias, start + 1)
            return None

        for entry in self.models:
            pos = find(entry)
            if pos is not None:
                hits.append((pos, entry))
        for entry in self.makes.values():
            pos = find(entry)
            if pos is not None:
                hits.append((pos, entry))
        hits.sort(key=lambda h: h[0])

        makes: list[str] = []
        models: list[str] = []
        segments: list[str] = []
        fleets: list[str] = []
        rhd = False
        for _, e in hits:
            make_entry = self.makes.get(e.make.upper())
            if e.make not in makes:
                makes.append(e.make)
            if e.model:
                models.append(e.model if e.model.upper().startswith(e.make.upper()) else f"{e.make} {e.model}")
            segments.append(e.segment if e.model else (make_entry.segment if make_entry else e.segment))
            fleet = e.fleet or (make_entry.fleet if make_entry else "")
            if fleet and fleet not in fleets:
                fleets.append(fleet)
            rhd = rhd or e.rhd
        # марки, у которых нашлись модели, в списке моделей уже есть
        bare = [m for m in makes if not any(x.upper().startswith(m.upper()) for x in models)]
        models = models + bare if models else []
        if not fleets:
            fleet = UNKNOWN
        elif len(fleets) == 1:
            fleet = fleets[0]
        else:
            fleet = "Смешанный: " + ", ".join(fleets)
        if rhd and fleets == ["Японский"]:
            fleet = "Японский (праворульный)"
        segment = max(segments, key=lambda s: SEGMENT_RANK.get(s, -1)) if segments else UNKNOWN
        info = CarInfo(makes, models, fleet, segment)
        self.cache[variety] = info
        return info
