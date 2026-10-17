"""Раскладка прайсов: скачанные файлы и архивы — по папкам «дата / поставщик».

Прайсы скачиваются как есть — Excel, CSV или ZIP — и бросаются в папку с датой
(«Прайсы/17.10.2026») или прямо в «Прайсы». Раскладка:

* распаковывает ZIP (и RAR/7z, если в Windows есть tar), вложенные архивы тоже;
* узнаёт поставщика и кладёт файл в «Прайсы/<дата>/<поставщик>/»;
* файл, брошенный прямо в «Прайсы», получает дату скачивания;
* сам архив после распаковки уезжает в «Прайсы — исходники/Архивы ZIP/<дата>».

Поставщик узнаётся так, по порядку:
1. по прошлым датам: файл с таким же именем (без даты в имени) уже лежал в папке поставщика;
2. по имени файла или архива: в нём есть название папки поставщика или поставщика из позиционирования;
3. по шапке: такая же шапка таблицы раньше была только у одного поставщика.
Не узнан — файл остаётся в папке-дате как есть (программа и так его прочитает, поставщик
возьмётся из имени), а в окне программы будет подсказка разложить его один раз руками.
Дальше такой файл узнаётся сам.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import json
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

from .excel import SUPPORTED_EXT, clean, norm_key
from .prices import _normalize_date, _shot_sort, match_known, stem_key, supplier_key
from .sources import _is_currency_file

MAP_NAME = "раскладка прайсов.json"
RAW_DIR = "Прайсы — исходники"
ZIP_DIR = "Архивы ZIP"
REPLACED_DIR = "Заменённые"
ARCHIVE_EXT = {".zip", ".rar", ".7z"}
SKIP_EXT = {".crdownload", ".part", ".partial", ".download", ".tmp", ".opdownload"}
SKIP_NAMES = {".gitkeep", "desktop.ini", "thumbs.db"}
MAX_DEPTH = 3
MIN_KEY = 3  # «MX» или «АП» в имени файла — не повод решать, чей это прайс

_DATE_IN_NAME = re.compile(r"(?<!\d)(\d{1,2})[._\-](\d{1,2})[._\-](20\d{2}|\d{2})(?!\d)|(?<!\d)(20\d{2})[._\-](\d{2})[._\-](\d{2})(?!\d)")
_COPY_MARK = re.compile(r"\s*\(\d{1,2}\)$|\s*-\s*(копия|copy)(\s*\(\d+\))?$", re.IGNORECASE)


def _program_dir() -> Path:
    from .settings import PROGRAM

    return PROGRAM


def name_key(name: str) -> str:
    """Имя файла без даты выгрузки и без «(1)» от браузера: по нему файл узнаётся в следующие дни."""
    stem = Path(name).stem
    stem = _COPY_MARK.sub("", stem)
    stem = _DATE_IN_NAME.sub(" ", stem)
    return stem_key(stem + ".x")


def _wanted(path: Path) -> bool:
    low = path.name.casefold()
    if low in SKIP_NAMES or low.startswith(("~$", ".")) or path.suffix.lower() in SKIP_EXT:
        return False
    return path.suffix.lower() in SUPPORTED_EXT or path.suffix.lower() in ARCHIVE_EXT


def _ru(day: dt.date) -> str:
    return day.strftime("%d.%m.%Y")


# ---------------------------------------------------------------- что уже известно


class Book:
    """Кто чей: имена файлов и шапки из уже разложенных папок плюс запомненное раньше."""

    def __init__(self, root: Path, raw: Path | None, known: list[str], map_path: Path):
        self.root = root
        self.map_path = map_path
        self.folders: dict[str, str] = {}  # ключ поставщика -> как называется его папка
        self.names: dict[str, set[str]] = {}  # name_key -> папки
        self.live_files: list[tuple[Path, str]] = []
        for base in (root, raw):
            if base is None or not base.is_dir():
                continue
            for date_dir in base.iterdir():
                if not (date_dir.is_dir() and _normalize_date(date_dir.name)):
                    continue
                for sup in date_dir.iterdir():
                    if not sup.is_dir() or sup.name.startswith((".", "~")):
                        continue
                    self.folders.setdefault(supplier_key(sup.name), sup.name)
                    for f in sup.rglob("*"):
                        if f.is_file() and f.suffix.lower() in SUPPORTED_EXT and not f.name.startswith(("~$", ".")):
                            key = name_key(f.name)
                            if key:
                                self.names.setdefault(key, set()).add(sup.name)
                            if base == root:
                                self.live_files.append((f, sup.name))
        self.saved = self._load()
        for key, sup in (self.saved.get("names") or {}).items():
            self.folders.setdefault(supplier_key(sup), sup)
        self.known = [k for k in known if len(supplier_key(k)) >= MIN_KEY]
        self._heads: dict[str, set[str]] | None = None

    def _load(self) -> dict:
        try:
            data = json.loads(self.map_path.read_text(encoding="utf-8"))
            return data if isinstance(data, dict) else {}
        except (OSError, ValueError):
            return {}

    def save(self) -> None:
        data = dict(self.saved)
        names = dict(data.get("names") or {})
        for key, sups in self.names.items():
            if len(sups) == 1:
                names[key] = next(iter(sups))
        data["names"] = names
        try:
            self.map_path.parent.mkdir(parents=True, exist_ok=True)
            tmp = self.map_path.with_suffix(".tmp")
            tmp.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
            tmp.replace(self.map_path)
        except OSError:
            pass

    def folder_for(self, supplier: str) -> str:
        """Папка уже есть — пишем в неё, как она названа («Armtek», а не «ARMTEK»)."""
        return self.folders.get(supplier_key(supplier)) or clean(supplier)

    def by_name(self, name: str) -> str:
        key = name_key(name)
        if not key:
            return ""
        live = self.names.get(key) or set()
        if len(live) == 1:
            return next(iter(live))
        if len(live) > 1:
            return ""  # одно и то же имя у разных поставщиков («price.xlsx») — по имени не решить
        return (self.saved.get("names") or {}).get(key, "")

    def by_title(self, name: str) -> str:
        stem = Path(name).stem
        names = sorted(set(self.folders.values()) | set(self.known), key=lambda s: len(supplier_key(s)), reverse=True)
        names = [n for n in names if len(supplier_key(n)) >= MIN_KEY]
        hit = match_known(stem, names)
        return self.folder_for(hit) if hit else ""

    def by_head(self, path: Path) -> str:
        fp = head_print(path)
        if not fp:
            return ""
        if self._heads is None:
            self._heads = {}
            dates = sorted({f.parent.parent.name for f, _ in self.live_files}, key=lambda d: _shot_sort(_normalize_date(d)))[-3:]
            for f, sup in self.live_files:
                if f.parent.parent.name not in dates:
                    continue  # шапки берём с последних трёх дат: старые файлы читать незачем
                other = head_print(f)
                if other:
                    self._heads.setdefault(other, set()).add(sup)
            for other, sup in (self.saved.get("heads") or {}).items():
                self._heads.setdefault(other, set()).add(sup)
        sups = self._heads.get(fp) or set()
        return next(iter(sups)) if len(sups) == 1 else ""

    def learn(self, name: str, supplier: str, path: Path | None = None) -> None:
        key = name_key(name)
        if key:
            self.names.setdefault(key, set()).add(supplier)
        if path is not None:
            fp = head_print(path)
            if fp:
                heads = dict(self.saved.get("heads") or {})
                heads[fp] = supplier
                self.saved["heads"] = heads


def head_print(path: Path) -> str:
    """Отпечаток шапки: строка с названиями столбцов в начале первого листа с данными."""
    if path.suffix.lower() not in SUPPORTED_EXT:
        return ""
    try:
        from .filecache import read_price_heads

        sheets = read_price_heads(path)
    except Exception:  # noqa: BLE001 — битый или не таблица: просто не узнаём по шапке
        return ""
    for sheet in sheets:
        best: list[str] = []
        for row in (sheet.grid or [])[:40]:
            words = [norm_key(c) for c in row if clean(c) and not re.fullmatch(r"[\d\s.,\-]+", clean(c)) and len(clean(c)) <= 40]
            if len(words) > len(best):
                best = words
        if len(best) >= 3:
            return hashlib.sha1("|".join(best).encode("utf-8")).hexdigest()[:16]
    return ""


# ---------------------------------------------------------------- распаковка


def _zip_name(info: zipfile.ZipInfo) -> str:
    name = info.filename
    if not info.flag_bits & 0x800:
        try:
            name = name.encode("cp437").decode("cp866")  # архивы из русской Windows
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
    return name.replace("\\", "/").rsplit("/", 1)[-1]


def _unpack(archive: Path, into: Path) -> list[Path]:
    """Файлы архива — плоско в папку into. Ошибка — исключение с понятным текстом."""
    into.mkdir(parents=True, exist_ok=True)
    out: list[Path] = []
    if archive.suffix.lower() == ".zip":
        try:
            with zipfile.ZipFile(archive) as zf:
                for info in zf.infolist():
                    if info.is_dir() or "__MACOSX" in info.filename:
                        continue
                    name = _zip_name(info)
                    if not name or name.startswith(("~$", ".")):
                        continue
                    target = _free(into / name)
                    with zf.open(info) as src, target.open("wb") as dst:
                        shutil.copyfileobj(src, dst)
                    out.append(target)
        except zipfile.BadZipFile as exc:
            raise ValueError("архив повреждён или скачан не до конца") from exc
        except RuntimeError as exc:  # архив с паролем
            raise ValueError("архив с паролем") from exc
        return out
    tar = shutil.which("tar")
    if not tar:
        raise ValueError(f"{archive.suffix.upper()[1:]} распаковать нечем — распакуйте руками или скачайте ZIP")
    stage = Path(tempfile.mkdtemp(dir=into))
    run = subprocess.run([tar, "-xf", str(archive), "-C", str(stage)], capture_output=True, timeout=600)
    if run.returncode != 0:
        shutil.rmtree(stage, ignore_errors=True)
        raise ValueError(f"{archive.suffix.upper()[1:]} не распаковался — распакуйте руками или скачайте ZIP")
    for f in sorted(stage.rglob("*")):
        if f.is_file() and not f.name.startswith(("~$", ".")):
            target = _free(into / f.name)
            shutil.move(str(f), str(target))
            out.append(target)
    shutil.rmtree(stage, ignore_errors=True)
    return out


def _free(path: Path) -> Path:
    if not path.exists():
        return path
    for i in range(2, 1000):
        candidate = path.with_name(f"{path.stem} ({i}){path.suffix}")
        if not candidate.exists():
            return candidate
    return path


def _twin(folder: Path, name: str, f: Path) -> Path | None:
    """Файл с тем же содержимым и тем же именем (без даты и «(1)») уже лежит в папке."""
    if not folder.is_dir():
        return None
    key = name_key(name)
    for other in folder.iterdir():
        if other.is_file() and other.resolve() != f.resolve() and name_key(other.name) == key and _same(other, f):
            return other
    return None


def _same(a: Path, b: Path) -> bool:
    try:
        if a.stat().st_size != b.stat().st_size:
            return False
        with a.open("rb") as fa, b.open("rb") as fb:
            while True:
                x, y = fa.read(1 << 20), fb.read(1 << 20)
                if x != y:
                    return False
                if not x:
                    return True
    except OSError:
        return False


# ---------------------------------------------------------------- раскладка


def sort_prices(folder: Path, known: list[str] | None = None, today: dt.date | None = None,
                keep: set[str] | None = None, map_path: Path | None = None) -> dict:
    """Раскладывает новые файлы в «Прайсах». Возвращает, что сделано, — для окна и для батника.

    keep — имена файлов с ручными настройками: такие из папки-даты не двигаем, иначе настройки потеряются.
    """
    from .dynamics import prices_dir

    root = prices_dir(folder)
    result = {"moved": [], "unpacked": [], "unknown": [], "errors": [], "dup": 0}
    if root is None:
        return result
    today = today or dt.date.today()
    keep = keep or set()
    dates = [d for d in root.iterdir() if d.is_dir() and _normalize_date(d.name)]
    loose_root = [f for f in root.iterdir() if f.is_file() and _wanted(f) and not _is_currency_file(f)]
    loose_dated = []
    for d in dates:
        for f in d.iterdir():
            if f.is_file() and _wanted(f) and not _is_currency_file(f) and f.name not in keep:
                loose_dated.append((f, d))
            elif f.is_dir():
                loose_dated += [(a, d) for a in f.rglob("*") if a.is_file() and a.suffix.lower() in ARCHIVE_EXT]
    if not dates and not any(f.suffix.lower() in ARCHIVE_EXT for f in loose_root):
        # без папок-дат программа читает прайсы прямо из папок поставщиков — так и оставляем
        loose_root = []
    if not loose_root and not loose_dated:
        return result

    data = root.parent  # «Данные»: исходники и временная распаковка — рядом с «Прайсами», не внутри
    raw = data / RAW_DIR
    book = Book(root, raw, known or [], map_path or (_program_dir() / MAP_NAME))
    work = Path(tempfile.mkdtemp(prefix="~распаковка ", dir=data))
    try:
        jobs: list[tuple[Path, Path, list[str], str]] = []  # файл, папка-дата, имена архивов, поставщик из папки
        for f in loose_root:
            day = dt.date.fromtimestamp(f.stat().st_mtime)
            day = min(day, today)
            date_dir = root / _ru(day)
            jobs.append((f, date_dir, [], ""))
        for f, d in loose_dated:
            sup = f.parent.name if f.parent != d else ""
            jobs.append((f, d, [], sup))

        while jobs:
            f, date_dir, hints, fixed = jobs.pop(0)
            if f.suffix.lower() in ARCHIVE_EXT:
                if len(hints) >= MAX_DEPTH:
                    result["errors"].append({"file": f.name, "why": "слишком много вложенных архивов"})
                    continue
                try:
                    inner = _unpack(f, Path(tempfile.mkdtemp(dir=work)))
                except (ValueError, OSError, subprocess.SubprocessError) as exc:
                    result["errors"].append({"file": f.name, "why": str(exc)})
                    continue
                result["unpacked"].append({"file": f.name, "files": len(inner), "date": _normalize_date(date_dir.name)})
                for g in inner:
                    if _wanted(g):
                        jobs.append((g, date_dir, hints + [f.name], fixed))
                if not str(f).startswith(str(work)):
                    dest = _free(raw / ZIP_DIR / date_dir.name / f.name)
                    try:
                        dest.parent.mkdir(parents=True, exist_ok=True)
                        shutil.move(str(f), str(dest))
                    except OSError:
                        result["errors"].append({"file": f.name, "why": "распакован, но архив не убрать — открыт или нет прав"})
                continue

            supplier, how = fixed, "папка" if fixed else ""
            name = f.name
            if not supplier:
                supplier, how = book.by_name(name), "по прошлым датам"
            if not supplier:
                supplier, how = book.by_title(name), "по имени файла"
            for h in hints:
                if supplier:
                    break
                supplier, how = book.by_name(h) or book.by_title(h), "по имени архива"
                if supplier and name_key(h) and name_key(h) not in name_key(name):
                    # «Armtek_Омск.zip» -> «price.xlsx»: имя архива — в имя файла, там город и поставщик
                    name = f"{Path(h).stem}_{name}"
            if not supplier:
                supplier, how = book.by_head(f), "по шапке"
            date = _normalize_date(date_dir.name)
            if supplier:
                target_dir = date_dir / book.folder_for(supplier)
            elif f.parent == date_dir:
                result["unknown"].append({"file": f.name, "date": date})
                continue
            else:
                target_dir = date_dir
            target = target_dir / name
            twin = _twin(target_dir, name, f)
            if twin is not None:
                # тот же прайс скачан второй раз («Armtek (1).xlsx») — второй экземпляр не нужен
                try:
                    f.unlink()
                    result["dup"] += 1
                except OSError:
                    result["errors"].append({"file": f.name, "why": "повтор уже лежащего файла, но удалить его не вышло"})
                continue
            try:
                target_dir.mkdir(parents=True, exist_ok=True)
                if target.exists() and target.resolve() != f.resolve():
                    if _same(target, f):
                        f.unlink()
                        result["dup"] += 1
                        continue
                    old = _free(raw / REPLACED_DIR / date_dir.name / target_dir.name / target.name)
                    old.parent.mkdir(parents=True, exist_ok=True)
                    shutil.move(str(target), str(old))
                shutil.move(str(f), str(target))
            except OSError:
                result["errors"].append({"file": f.name, "why": "файл открыт в Excel или нет прав — закройте и обновите"})
                continue
            if supplier:
                if how not in ("папка",):
                    book.learn(name, target_dir.name, target)
                result["moved"].append({"file": name, "supplier": target_dir.name, "date": date, "how": how})
            else:
                result["unknown"].append({"file": name, "date": date})
        book.save()
    finally:
        shutil.rmtree(work, ignore_errors=True)
    # папка из архива, ставшая пустой
    for d in dates:
        for sub in d.iterdir():
            if sub.is_dir() and not any(sub.rglob("*")):
                try:
                    sub.rmdir()
                except OSError:
                    pass
    return result


def summary_lines(result: dict) -> list[str]:
    """Отчёт для батника и подсказки в окне."""
    lines = []
    by: dict[str, int] = {}
    for m in result["moved"]:
        by[m["supplier"]] = by.get(m["supplier"], 0) + 1
    if result["unpacked"]:
        n = sum(u["files"] for u in result["unpacked"])
        lines.append(f"Распаковано архивов: {len(result['unpacked'])}, файлов в них: {n}")
    if by:
        lines.append(f"Разложено файлов: {len(result['moved'])} — " + ", ".join(f"{k} {v}" for k, v in sorted(by.items())))
    if result["dup"]:
        lines.append(f"Повторы, которые уже лежали в папке, убраны: {result['dup']}")
    for u in result["unknown"]:
        lines.append(f"Не узнал поставщика: {u['file']} ({u['date']}) — положите его в папку поставщика руками один раз, дальше узнаю сам")
    for e in result["errors"]:
        lines.append(f"Не получилось: {e['file']} — {e['why']}")
    return lines
