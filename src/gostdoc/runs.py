"""Нормализация run'а: гарнитура, кегль, цвет — с сохранением смысловых акцентов.

Главные риски (ТЗ, разд. 8):
- №1 прямое форматирование перебивает стиль → правим именно run, а не только стиль.
- №2 нельзя терять b/i/u/vertAlign → их вообще не трогаем.
- №16 пересоздание run теряет рисунки/объекты → меняем только rPr существующего run'а.
"""

from __future__ import annotations

from docx.text.run import Run

from . import constants as c
from . import xml_utils


def normalize_run(run: Run, font_name: str = c.FONT_NAME, font_size=c.FONT_SIZE_BODY) -> None:
    """Привести гарнитуру/кегль/цвет run'а к ГОСТ, сохранив акценты и вложенные объекты."""
    rpr = run._r.get_or_add_rPr()
    xml_utils.set_run_fonts(rpr, font_name)
    xml_utils.set_run_size(rpr, font_size)
    xml_utils.set_run_color_auto(rpr)


def normalize_heading_run(run: Run, font_name: str = c.FONT_NAME, font_size=c.FONT_SIZE_BODY) -> None:
    """Run заголовка: гарнитура/кегль/цвет + принудительно полужирный, без подчёркивания.

    Полужирность и снятие подчёркивания для заголовков предписаны ГОСТ 7.32-2017
    (п.6.2.3); прямое форматирование студента (b=False/u=True) перебивается явно.
    """
    normalize_run(run, font_name, font_size)
    run.bold = True
    run.underline = False
