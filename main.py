#!/usr/bin/env python3
"""Ценовой анализ: окно программы в браузере.

Запуск: «Запуск.bat» (Windows) или `python main.py`.
Ключи: --port N (по умолчанию 47615), --no-browser (не открывать окно сам),
--sort-prices (только разложить прайсы по папкам, как «Разложить прайсы.bat»).
"""

from __future__ import annotations

import os

# До импорта openpyxl: с lxml Excel видит битый порядок элементов и удаляет лист.
os.environ["OPENPYXL_LXML"] = "False"

import argparse
import shutil
import subprocess
import sys
import threading
import webbrowser
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "Программа"))

DEFAULT_PORT = 47615


def open_window(url: str) -> None:
    """Открывает страницу отдельным окном Edge/Chrome (без вкладок), иначе — обычным браузером."""
    if os.name == "nt":
        candidates = [
            shutil.which("msedge"),
            os.path.expandvars(r"%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"),
            os.path.expandvars(r"%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"),
            os.path.expandvars(r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"),
            os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe"),
        ]
        for exe in candidates:
            if exe and Path(exe).exists():
                try:
                    subprocess.Popen([exe, f"--app={url}", "--window-size=1360,900"])
                    return
                except OSError:
                    continue
    webbrowser.open(url)


def sort_only() -> int:
    """«Разложить прайсы.bat»: распаковать архивы и разложить файлы по папкам «дата / поставщик»."""
    from pricing.server import ensure_layout, sort_prices
    from pricing.settings import Settings

    ensure_layout(ROOT)
    print()
    print("  Раскладываю прайсы в «Данные/Прайсы»...")
    result = sort_prices(ROOT, Settings.load())
    lines = (result or {}).get("lines") or []
    print()
    for line in lines or ["Новых файлов нет — всё уже разложено."]:
        print("  " + line)
    print()
    return 0


def main(argv=None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass
    parser = argparse.ArgumentParser(description="Ценовой анализ")
    parser.add_argument("--port", type=int, default=DEFAULT_PORT)
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--no-browser", action="store_true")
    parser.add_argument("--sort-prices", action="store_true", help="только разложить прайсы по папкам и выйти")
    args = parser.parse_args(argv)

    try:
        import openpyxl  # noqa: F401
    except ImportError:
        print("Не установлен модуль openpyxl. Выполните: python -m pip install openpyxl")
        return 1

    if args.sort_prices:
        return sort_only()

    from pricing import filecache
    from pricing.server import ensure_layout, serve

    filecache.PARALLEL = True  # новые файлы читаются в нескольких процессах сразу

    ensure_layout(ROOT)
    (ROOT / "Входящие").mkdir(exist_ok=True)
    server = serve(args.host, args.port)
    url = f"http://{args.host}:{server.server_address[1]}/"
    print()
    print("  Ценовой анализ запущен.")
    print(f"  Окно программы: {url}")
    print("  Кроссы — в «Данные/Кроссы», прайсы — в «Данные/Прайсы».")
    print("  На каждую группу бренда — своя папка в «Данные/Группы».")
    print("  Пока там пусто, читается «Входящие». Отчёты — в «Результат».")
    print("  Чтобы закрыть программу — закройте это окно.")
    print()
    if not args.no_browser:
        threading.Timer(0.4, open_window, args=(url,)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
