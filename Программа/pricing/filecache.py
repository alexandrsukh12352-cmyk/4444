"""Кэш разобранных файлов.

Файл читается из Excel заново, только если изменились дата или размер.
Рабочая область на диск пишется без сеток ячеек: сетки и так лежат в кэше
каждого файла. Целиком с сетками прайсов она слишком большая и при открытии
окна не успевает ни сохраниться, ни прочитаться.
"""

from __future__ import annotations

import hashlib
import json
import pickle
import threading
from pathlib import Path

from .excel import read_sheets

CACHE_VERSION = 3
# Общий файл рабочей области версии 3 хранил сетки всех прайсов и подвешивал запуск.
WORKSPACE_VERSION = 4
CACHE_DIR = Path(__file__).resolve().parent.parent / "кэш"
_lock = threading.Lock()
_memory: dict[tuple, list] = {}
misses = 0


def reset_stats() -> None:
    global misses
    misses = 0


def _slot(path: Path, mtime_ns: int, size: int) -> Path:
    raw = f"{CACHE_VERSION}|{path.resolve()}|{mtime_ns}|{size}".encode()
    return CACHE_DIR / f"{hashlib.sha1(raw).hexdigest()}.pkl"


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
