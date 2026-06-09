"""Точка входа для сборки GUI-.exe (PyInstaller, windowed).

Сборка (см. packaging/build.md):
    pyinstaller --onefile --windowed --collect-data docx --name gostdoc-gui \
        packaging/gostdoc_gui_entry.py
"""

import sys

from gostdoc.gui import main

if __name__ == "__main__":
    sys.exit(main())
