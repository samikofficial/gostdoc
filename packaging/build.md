# Сборка автономного .exe (Windows)

Чтобы инструментом можно было пользоваться без установки Python, собирается
один самодостаточный `gostdoc.exe`.

## Как собрать

```
python -m venv .venv
.venv\Scripts\activate
pip install -e .
pip install pyinstaller
pyinstaller --onefile --collect-data docx --name gostdoc ^
    --distpath dist --workpath build/pyi --specpath build ^
    packaging/gostdoc_entry.py
```

Результат: `dist/gostdoc.exe` (~12 МБ, не требует Python).

`--collect-data docx` обязателен: иначе в сборку не попадёт шаблон `default.docx`
из python-docx и `Document()` упадёт.

### GUI-версия (окно с интерфейсом)

```
pyinstaller --onefile --windowed --collect-data docx --name gostdoc-gui ^
    --distpath dist --workpath build/pyi_gui --specpath build ^
    packaging/gostdoc_gui_entry.py
```

Результат: `dist/gostdoc-gui.exe` (~15 МБ, включает Tcl/Tk). `--windowed` — без консольного
окна. Двойной клик открывает окно с выбором файла и настройками; можно также перетащить
`.docx` на `.exe` — путь подставится в окно.

## Использование .exe

```
gostdoc.exe ДИПЛОМ.docx
gostdoc.exe ДИПЛОМ.docx --no-detect-structure
gostdoc.exe ДИПЛОМ.docx --check
```

Студент может перетащить `.docx` на `gostdoc.exe` или вызвать из консоли.
Результат пишется рядом: `ДИПЛОМ.gost.docx`.

При запуске двойным кликом/перетаскиванием (своя новая консоль) программа покажет
результат и подождёт нажатия Enter, чтобы окно не закрылось мгновенно. В обычном
терминале и в пайпах (CI) паузы нет — определяется через `GetConsoleProcessList`.

> `dist/` и `build/` в `.gitignore` — артефакты сборки в репозиторий не коммитятся.
