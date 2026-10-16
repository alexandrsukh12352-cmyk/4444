"""Архив прайсов: старые даты — компактно, без исходных файлов в «Прайсах».

Каждый день в «Прайсах» появляется папка-дата с файлами поставщиков. Отчёту нужна только
последняя дата, старые — динамике цен. Чтобы папка и запуск не росли месяц за месяцем, даты
старше двух недель сжимаются: от каждого файла остаются строки брендов конкурентов (бренд,
артикул, цена, наличие, строка файла) — это в десятки раз меньше исходников. Архив лежит в
«Данные/Архив цен», по файлу на дату. Сами папки-даты переносятся в «Данные/Прайсы — исходники»:
программа их больше не читает, удалить их можно руками, когда захочется.
"""

from __future__ import annotations

import datetime as dt
import gzip
import pickle
import shutil
from pathlib import Path

from .excel import Sheet
from .prices import Offer, PriceFile, _shot_sort, brand_key

VERSION = 1
ARCHIVE_DIR = "Архив цен"
RAW_DIR = "Прайсы — исходники"
KEEP_DAYS = 14  # даты свежее двух недель остаются файлами
KEEP_MIN = 3  # и не меньше трёх последних дат, даже если выгрузок давно не было

_memo: dict = {}


def data_dir(folder: Path) -> Path:
    data = folder / "Данные"
    return data if data.is_dir() else folder


def archive_dir(folder: Path) -> Path:
    return data_dir(folder) / ARCHIVE_DIR


def _day(text: str) -> dt.date | None:
    y, m, d = _shot_sort(text)
    try:
        return dt.date(y, m, d) if y else None
    except ValueError:
        return None


def _slot(folder: Path, date: str) -> Path:
    day = _day(date)
    return archive_dir(folder) / f"{day.isoformat() if day else date}.pkl.gz"


def plan(files: list[PriceFile], today: dt.date | None = None, keep_days: int | None = None, keep_min: int | None = None) -> list[str]:
    """Какие даты сжимать: старше keep_days, кроме keep_min последних."""
    today = today or dt.date.today()
    keep_days = KEEP_DAYS if keep_days is None else keep_days
    keep_min = KEEP_MIN if keep_min is None else keep_min
    dates = sorted({pf.shot_date for pf in files if pf.shot_date and _day(pf.shot_date)}, key=_shot_sort)
    keep = set(dates[-keep_min:]) if keep_min else set()
    return [d for d in dates if d not in keep and _day(d) < today - dt.timedelta(days=keep_days)]


def _date_folders(files: list[PriceFile], date: str) -> set[Path]:
    """Папки-даты внутри «Прайсов», где лежат файлы этой даты."""
    out = set()
    for pf in files:
        if pf.shot_date != date:
            continue
        path = Path(pf.sheet.file)
        for parent in path.parents:
            if parent.parent.name.casefold() in {"прайсы", "prices"}:
                out.add(parent)
                break
    return out


def _size(paths) -> int:
    total = 0
    for p in paths:
        if p.is_file():
            total += p.stat().st_size
        elif p.is_dir():
            total += sum(f.stat().st_size for f in p.rglob("*") if f.is_file())
    return total


def preview(folder: Path, today: dt.date | None = None) -> dict:
    """Что сожмётся, без чтения прайсов: по папкам-датам в «Прайсах»."""
    from .dynamics import prices_dir
    from .prices import _normalize_date

    root = prices_dir(folder)
    shots = []
    if root is not None:
        for d in root.iterdir():
            date = _normalize_date(d.name) if d.is_dir() else ""
            if date:
                shots.append(type("Shot", (), {"shot_date": date, "path": d})())
    dates = plan(shots, today)
    dirs = [s.path for s in shots if s.shot_date in dates]
    files = sum(1 for d in dirs for f in d.rglob("*") if f.is_file())
    return {"dates": dates, "files": files, "bytes": _size(dirs), "folders": len(dirs), "archived": archived_dates(folder),
            "keep_days": KEEP_DAYS}


def compact(folder: Path, files: list[PriceFile], keep_brands: set[str], today: dt.date | None = None) -> dict:
    """Сжимает старые даты в архив и переносит их папки в «Прайсы — исходники».

    keep_brands — ключи брендов, строки которых сохраняются (конкуренты из позиционирования и кроссов).
    Папка-дата переносится, только если архив этой даты записан и прочитан обратно без потерь.
    """
    dates = plan(files, today)
    target = archive_dir(folder)
    target.mkdir(parents=True, exist_ok=True)
    raw_root = data_dir(folder) / RAW_DIR
    done, moved, kept_rows, before = [], 0, 0, 0
    for date in dates:
        part = [pf for pf in files if pf.shot_date == date and pf.enabled and pf.offers]
        records = []
        for pf in part:
            keys: dict[str, str] = {}
            rows = []
            for o in pf.offers:
                bk = keys.get(o.brand)
                if bk is None:
                    bk = keys[o.brand] = brand_key(o.brand)
                if bk in keep_brands:
                    rows.append((o.brand, o.article, o.price, o.qty, o.row, o.article_key))
            records.append({"supplier": pf.supplier, "city": pf.city, "name": Path(pf.sheet.file).name,
                            "sheet": pf.sheet.name, "status": pf.status, "offers": len(pf.offers), "rows": rows})
        blob = {"version": VERSION, "date": date, "files": records}
        slot = _slot(folder, date)
        tmp = slot.with_suffix(".tmp")
        with gzip.open(tmp, "wb") as fh:
            pickle.dump(blob, fh, protocol=pickle.HIGHEST_PROTOCOL)
        tmp.replace(slot)
        back = _read(slot)
        if not back or sum(len(r["rows"]) for r in back["files"]) != sum(len(r["rows"]) for r in records):
            continue  # архив не сошёлся — исходники не трогаем
        kept_rows += sum(len(r["rows"]) for r in records)
        dirs = _date_folders(files, date)
        before += _size(dirs)
        for src in dirs:
            dst = raw_root / src.name
            if dst.exists():
                dst = raw_root / f"{src.name} ({dt.datetime.now():%H%M%S})"
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(src), str(dst))
            moved += 1
        done.append(date)
    _memo.clear()
    return {"dates": done, "folders": moved, "rows": kept_rows, "bytes_before": before,
            "bytes_archive": _size([_slot(folder, d) for d in done])}


def _read(path: Path) -> dict | None:
    try:
        st = path.stat()
    except OSError:
        return None
    key = (str(path), st.st_mtime_ns, st.st_size)
    if key in _memo:
        return _memo[key]
    try:
        with gzip.open(path, "rb") as fh:
            blob = pickle.load(fh)
    except (OSError, pickle.PickleError, EOFError, ValueError):
        return None
    if not isinstance(blob, dict) or blob.get("version") != VERSION:
        return None
    _memo[key] = blob
    return blob


def archived_dates(folder: Path) -> list[str]:
    out = []
    for path in sorted(archive_dir(folder).glob("*.pkl.gz")):
        blob = _read(path)
        if blob:
            out.append(blob["date"])
    return out


def archived_files(folder: Path, skip_dates: set[str] | None = None) -> list[PriceFile]:
    """Архивные даты как прайсы для динамики: поставщик, город, дата и строки конкурентов.

    Дата, которая есть и файлами в «Прайсах», из архива не берётся — файлы важнее.
    """
    skip = skip_dates or set()
    out: list[PriceFile] = []
    for path in sorted(archive_dir(folder).glob("*.pkl.gz")):
        blob = _read(path)
        if not blob or blob["date"] in skip:
            continue
        for i, rec in enumerate(blob["files"]):
            pf = PriceFile(
                sheet=Sheet(path, rec.get("sheet") or "", []), supplier=rec["supplier"], city=rec.get("city") or "",
                header=None, cols={"архив": i}, fixed_brand="", status="архив",
                offers=[Offer(*row) for row in rec["rows"]], shot_date=blob["date"],
            )
            pf.source_name = rec.get("name") or path.name  # для выгрузки: как назывался файл
            out.append(pf)
    return out
