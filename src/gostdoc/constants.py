"""ГОСТ 7.32-2017: значения оформления. Единственный источник правды.

Менять параметры ГОСТ можно ТОЛЬКО здесь. Нигде в коде эти числа не хардкодить
повторно. Маппинг на python-docx указан рядом с каждым значением (см. ТЗ, разд. 7.1).
"""

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Cm, Mm, Pt

# --- Поля страницы (7.32-2017, п.6.1) ---
MARGIN_LEFT = Mm(30)
MARGIN_RIGHT = Mm(15)
MARGIN_TOP = Mm(20)
MARGIN_BOTTOM = Mm(20)

# --- Шрифт ---
# Единая гарнитура (применяется во все слоты rFonts: ascii/hAnsi/cs/eastAsia).
FONT_NAME = "Times New Roman"
# Кегль тела: 14 пт (стандарт задаёт «не менее 12»; 14 — практика 7.32 для ВКР/курсовых).
FONT_SIZE_BODY = Pt(14)
# Цвет текста — авто (чёрный). Записывается как w:color val="auto" на уровне XML.
COLOR_AUTO = "auto"

# --- Интервалы и отступы тела (body) ---
# Полуторный: присвоение line_spacing=1.5 даёт line=360/lineRule=auto
# (python-docx читает это правило как ONE_POINT_FIVE) — это и есть полуторный по ГОСТ.
LINE_SPACING_BODY = 1.5
# Подписи и абзацы таблиц — одинарный интервал.
LINE_SPACING_RULE_SINGLE = WD_LINE_SPACING.SINGLE
SPACE_BEFORE = Pt(0)
SPACE_AFTER = Pt(0)
# Абзацный отступ первой строки тела (7.32-2017, п.6.1).
FIRST_LINE_INDENT = Cm(1.25)
# Без отступа первой строки (подписи, ячейки таблиц, списки).
NO_FIRST_LINE_INDENT = Cm(0)
ALIGN_BODY = WD_ALIGN_PARAGRAPH.JUSTIFY

# --- Имена стилей и их локализованные синонимы (подводный камень №8) ---
# Доступ к стилям бывает по русским именам; держим карту синонимов.
STYLE_NAMES_NORMAL = ("Normal", "Обычный")
STYLE_NAMES_BODY_TEXT = ("Body Text", "Основной текст")
# Префиксы имён стилей заголовков (рус./англ.).
STYLE_PREFIXES_HEADING = ("Heading", "Заголовок")
STYLE_NAMES_CAPTION = ("Caption", "Подпись")
# Служебные стили (оглавление) — содержимое и порядок не трогаем.
STYLE_PREFIXES_TOC = ("TOC", "Оглавление")
