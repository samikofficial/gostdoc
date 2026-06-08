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

## Использование .exe

```
gostdoc.exe ДИПЛОМ.docx
gostdoc.exe ДИПЛОМ.docx --no-detect-structure
gostdoc.exe ДИПЛОМ.docx --check
```

Студент может перетащить `.docx` на `gostdoc.exe` или вызвать из консоли.
Результат пишется рядом: `ДИПЛОМ.gost.docx`.

> `dist/` и `build/` в `.gitignore` — артефакты сборки в репозиторий не коммитятся.
