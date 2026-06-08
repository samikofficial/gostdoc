"""Точка входа для сборки автономного .exe (PyInstaller).

PyInstaller требует файл-скрипт, а не entry-point из pyproject. Этот модуль просто
вызывает CLI. Сборка (см. packaging/build.md):
    pyinstaller --onefile --collect-data docx --name gostdoc packaging/gostdoc_entry.py
"""

import sys

from gostdoc.cli import main

if __name__ == "__main__":
    sys.exit(main())
