"""Кэш разобранных файлов.

Файл читается из Excel заново, только если изменились дата или размер.
Рабочая область на диск пишется без сеток ячеек: сетки и так лежат в кэше
каждого файла. Целиком с сетками прайсов она слишком большая и при открытии
окна не успевает ни сохраниться, ни прочитаться.
"""

from __future__ import annotations

import hashlib
import json
import os
import pickle
import threading
from pathlib import Path

from .excel import PriceSheet, read_sheets

CACHE_VERSION = 3
# Общий файл рабочей области версии 3 хранил сетки всех прайсов и подвешивал запуск.
# Версия 5: прайсы прошлых дат в рабочую область не читаются (src.dormant_prices).
# Версия 6: от прайса в рабочей области только начало листа (PriceSheet), без полной сетки.
WORKSPACE_VERSION = 6
# Сколько строк прайса держится в памяти: шапка ищется в первых 60, столбцы проверяются по 300.
HEAD_ROWS = 400
CACHE_DIR = Path(__file__).resolve().parent.parent / "кэш"
_lock = threading.Lock()
_memory: dict[tuple, list] = {}
misses = 0


def reset_stats() -> None:
    global misses
    misses = 0


def _slot(path: Path, mtime_ns: int, size: int, cache_dir: Path | None = None) -> Path:
    raw = f"{CACHE_VERSION}|{path.resolve()}|{mtime_ns}|{size}".encode()
    return (cache_dir or CACHE_DIR) / f"{hashlib.sha1(raw).hexdigest()}.pkl"


# ---------------------------------------------------------------- параллельное чтение новых файлов

PARALLEL_MIN_BYTES = 8 * 1024 * 1024  # меньше — процессы дороже самого чтения
PARALLEL_MAX = 6
# Включает main.py. Процесс-помощник заново запускает главный файл программы, поэтому
# параллельное чтение разрешено только там, где главный файл защищён `if __name__ == "__main__"`.
PARALLEL = False


def _cached_on_disk(path: Path, want_head: bool) -> bool:
    st = path.stat()
    if not _slot(path, st.st_mtime_ns, st.st_size).is_file():
        return False
    return not want_head or _head_slot(path, st.st_mtime_ns, st.st_size).is_file()


def _prefetch_one(args: tuple) -> bool:
    """В отдельном процессе: прочитать файл и положить сетку (и начало прайса) в кэш на диске."""
    raw_path, cache_dir, want_head = args
    path, cache = Path(raw_path), Path(cache_dir)
    st = path.stat()
    full_slot = _slot(path, st.st_mtime_ns, st.st_size, cache)
    sheets = None
    if full_slot.is_file():
        try:
            with full_slot.open("rb") as fh:
                sheets = pickle.load(fh)
        except (OSError, pickle.PickleError, EOFError, AttributeError, ValueError):
            sheets = None
    if sheets is None:
        sheets = read_sheets(path)
        _store(full_slot, sheets, cache)
    if want_head:
        raw = f"head|{CACHE_VERSION}|{HEAD_ROWS}|{path.resolve()}|{st.st_mtime_ns}|{st.st_size}".encode()
        head_slot = cache / f"head-{hashlib.sha1(raw).hexdigest()}.pkl"
        _store(head_slot, [(s.name, len(s.grid), s.grid[:HEAD_ROWS]) for s in sheets], cache)
    return True


def prefetch(paths: list[Path], price_paths: set[Path] | None = None, on_step=None) -> int:
    """Файлы, которых нет в кэше, читаются из Excel параллельно. Возвращает, сколько прочитано.

    Новые прайсы приходят каждую неделю, и первый запуск после этого — самый долгий.
    Несколько файлов сразу читают несколько ядер процессора. Если процессы не запустились,
    ничего страшного: файлы прочитаются по одному, как раньше.
    """
    global misses
    if not PARALLEL:
        return 0
    price_paths = price_paths or set()
    todo = []
    for p in paths:
        try:
            if not _cached_on_disk(p, p in price_paths):
                todo.append(p)
        except OSError:
            continue
    if len(todo) < 2 or sum(p.stat().st_size for p in todo) < PARALLEL_MIN_BYTES:
        return 0
    cpus = os.cpu_count() or 1
    if cpus < 2:
        return 0
    workers = min(len(todo), max(2, cpus - 1), PARALLEL_MAX)
    # Крупные файлы первыми: так процессы заканчивают примерно вместе.
    todo.sort(key=lambda p: p.stat().st_size, reverse=True)
    done = 0
    try:
        import multiprocessing
        from concurrent.futures import ProcessPoolExecutor, as_completed

        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        # spawn — как в Windows: процесс стартует с чистого листа и сам находит модули программы.
        with ProcessPoolExecutor(max_workers=workers, mp_context=multiprocessing.get_context("spawn")) as pool:
            jobs = {pool.submit(_prefetch_one, (str(p), str(CACHE_DIR), p in price_paths)): p for p in todo}
            for job in as_completed(jobs):
                try:
                    job.result()
                    done += 1
                    if on_step:
                        on_step(jobs[job].name, done, len(todo))
                except Exception:  # noqa: BLE001 — этот файл прочитается обычным порядком
                    continue
    except Exception:  # noqa: BLE001 — процессы не поднялись: читаем по одному
        return done
    with _lock:
        misses += done
    return done


def read_sheets_cached(path: Path) -> list:
    """Сетка листов. Повторно из Excel — только если файл новее или другого размера."""
    global misses
    st = path.stat()
    key = (str(path.resolve()), st.st_mtime_ns, st.st_size)
    with _lock:
        cached = _memory.get(key)
    if cached is not None:
        return cached
    slot = _slot(path, st.st_mtime_ns, st.st_size)
    if slot.is_file():
        try:
            with slot.open("rb") as fh:
                sheets = pickle.load(fh)
            if isinstance(sheets, list):
                with _lock:
                    _memory[key] = sheets
                return sheets
        except (OSError, pickle.PickleError, EOFError, AttributeError, ValueError):
            pass
    with _lock:
        misses += 1
    sheets = read_sheets(path)
    with _lock:
        _memory[key] = sheets
    try:
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        tmp = slot.with_suffix(".tmp")
        with tmp.open("wb") as fh:
            pickle.dump(sheets, fh, protocol=pickle.HIGHEST_PROTOCOL)
        tmp.replace(slot)
    except OSError:
        pass
    return sheets


def read_full_uncached(path: Path) -> list:
    """Полная сетка из кэша файла (или из Excel), не оставляя её в памяти программы."""
    global misses
    st = path.stat()
    key = (str(path.resolve()), st.st_mtime_ns, st.st_size)
    with _lock:
        cached = _memory.get(key)
    if cached is not None:
        return cached
    slot = _slot(path, st.st_mtime_ns, st.st_size)
    if slot.is_file():
        try:
            with slot.open("rb") as fh:
                sheets = pickle.load(fh)
            if isinstance(sheets, list):
                return sheets
        except (OSError, pickle.PickleError, EOFError, AttributeError, ValueError):
            pass
    with _lock:
        misses += 1
    sheets = read_sheets(path)
    _store(slot, sheets)
    return sheets


def _store(slot: Path, data, cache_dir: Path | None = None) -> None:
    try:
        (cache_dir or CACHE_DIR).mkdir(parents=True, exist_ok=True)
        tmp = slot.with_suffix(".tmp")
        with tmp.open("wb") as fh:
            pickle.dump(data, fh, protocol=pickle.HIGHEST_PROTOCOL)
        tmp.replace(slot)
    except OSError:
        pass


_heads: dict[tuple, list] = {}


def _head_slot(path: Path, mtime_ns: int, size: int) -> Path:
    raw = f"head|{CACHE_VERSION}|{HEAD_ROWS}|{path.resolve()}|{mtime_ns}|{size}".encode()
    return CACHE_DIR / f"head-{hashlib.sha1(raw).hexdigest()}.pkl"


def read_price_heads(path: Path) -> list[PriceSheet]:
    """Листы прайса: первые HEAD_ROWS строк и размер листа. Полная сетка остаётся в кэше на диске."""
    st = path.stat()
    key = (str(path.resolve()), st.st_mtime_ns, st.st_size)
    with _lock:
        hit = _heads.get(key)
    if hit is None:
        slot = _head_slot(path, st.st_mtime_ns, st.st_size)
        if slot.is_file():
            try:
                with slot.open("rb") as fh:
                    hit = pickle.load(fh)
                if not isinstance(hit, list):
                    hit = None
            except (OSError, pickle.PickleError, EOFError, AttributeError, ValueError):
                hit = None
        if hit is None:
            full = read_full_uncached(path)
            hit = [(s.name, len(s.grid), s.grid[:HEAD_ROWS]) for s in full]
            _store(slot, hit)
        with _lock:
            _heads[key] = hit
    return [PriceSheet(path, name, list(head), total=total, index=i) for i, (name, total, head) in enumerate(hit)]


def _signature_rows(signature: tuple) -> list:
    return [[path, mtime, size] for path, mtime, size in signature]


def _iter_sheets(workspace):
    seen: set[int] = set()

    def take(sheet) -> None:
        if sheet is None or id(sheet) in seen:
            return
        seen.add(id(sheet))
        found.append(sheet)

    def walk(src) -> None:
        if not src:
            return
        for locs in (getattr(src, "found", None) or {}).values():
            for loc in locs:
                header = getattr(loc, "header", None)
                take(getattr(header, "sheet", None) if header else None)
        for sheet in getattr(src, "price_sheets", None) or []:
            take(sheet)

    found: list = []
    walk(getattr(workspace, "shared", None))
    walk(getattr(workspace, "company", None))
    for pack in getattr(workspace, "packs", None) or []:
        walk(getattr(pack, "sources", None))
    for sl in getattr(workspace, "acp_slices", None) or []:
        loc = getattr(sl, "located", None)
        header = getattr(loc, "header", None) if loc else None
        take(getattr(header, "sheet", None) if header else None)
    for loc in getattr(workspace, "oem_locs", None) or []:
        header = getattr(loc, "header", None)
        take(getattr(header, "sheet", None) if header else None)
    return found


def _strip_grids(workspace) -> list:
    saved = []
    for sheet in _iter_sheets(workspace):
        if isinstance(sheet, PriceSheet):
            continue  # начало прайса маленькое — хранится в рабочей области как есть
        saved.append((sheet, sheet.grid))
        sheet.grid = None
    return saved


def _restore_grids(saved) -> None:
    for sheet, grid in saved:
        sheet.grid = grid


def _rehydrate(workspace) -> bool:
    """Сетки, которые не стали класть в общий кэш, читаются из кэша своего файла."""
    pending: dict[Path, list] = {}
    for sheet in _iter_sheets(workspace):
        if sheet.grid is not None:
            continue
        pending.setdefault(Path(sheet.file), []).append(sheet)
    for path, sheets in pending.items():
        try:
            loaded = {item.name: item for item in read_sheets_cached(path)}
        except (OSError, pickle.PickleError, EOFError, AttributeError, ValueError):
            return False
        for sheet in sheets:
            src = loaded.get(sheet.name)
            if src is None:
                return False
            sheet.grid = src.grid
    return True


def load_workspace(signature: tuple):
    """Разобранная рабочая область, если набор файлов и их даты те же.

    Старый общий кэш со сетками не открываем: он и подвешивал запуск.
    """
    meta_path = CACHE_DIR / "workspace.json"
    blob_path = CACHE_DIR / "workspace.pkl"
    if not meta_path.is_file() or not blob_path.is_file():
        return None
    try:
        meta = json.loads(meta_path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    if meta.get("version") != WORKSPACE_VERSION or meta.get("signature") != _signature_rows(signature):
        return None
    try:
        with blob_path.open("rb") as fh:
            workspace = pickle.load(fh)
    except (OSError, pickle.PickleError, EOFError, AttributeError, ValueError):
        return None
    if not _rehydrate(workspace):
        return None
    return workspace


def save_workspace(signature: tuple, workspace) -> None:
    saved = _strip_grids(workspace)
    try:
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        tmp = CACHE_DIR / "workspace.tmp"
        with tmp.open("wb") as fh:
            pickle.dump(workspace, fh, protocol=pickle.HIGHEST_PROTOCOL)
        tmp.replace(CACHE_DIR / "workspace.pkl")
        (CACHE_DIR / "workspace.json").write_text(
            json.dumps({"version": WORKSPACE_VERSION, "signature": _signature_rows(signature)}, ensure_ascii=False),
            encoding="utf-8",
        )
    except (OSError, pickle.PickleError, TypeError, ValueError):
        pass
    finally:
        _restore_grids(saved)


def load_gaps(signature: tuple, token: str) -> list | None:
    """Уже посчитанные «кроссов нет / цен нет» для того же набора файлов и прайсов."""
    path = CACHE_DIR / "gaps.json"
    if not path.is_file():
        return None
    try:
        meta = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    if meta.get("version") != WORKSPACE_VERSION or meta.get("token") != token:
        return None
    if meta.get("signature") != _signature_rows(signature):
        return None
    rows = meta.get("rows")
    return rows if isinstance(rows, list) else None


def save_gaps(signature: tuple, token: str, rows: list) -> None:
    try:
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        (CACHE_DIR / "gaps.json").write_text(
            json.dumps(
                {"version": WORKSPACE_VERSION, "token": token, "signature": _signature_rows(signature), "rows": rows},
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
    except (OSError, TypeError, ValueError):
        pass
