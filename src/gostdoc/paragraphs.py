"""Интервалы, отступы и выравнивание абзаца — по его категории (ТЗ, разд. 7.2).

Риски (ТЗ, разд. 8):
- №9 правило межстрочного интервала: явно ставим MULTIPLE + 1.5 (не полагаемся на EXACTLY).
- №10 абзацный отступ только у body; спискам/подписям/таблицам — не навязывать.
- №11 ячейки таблиц: одинарный интервал, без отступа первой строки, выравнивание не трогаем.
"""

from __future__ import annotations

from docx.enum.text import WD_LINE_SPACING
from docx.text.paragraph import Paragraph

from . import constants as c
from .classify import (
    CATEGORY_BODY,
    CATEGORY_CAPTION,
    CATEGORY_LIST_ITEM,
    CATEGORY_TABLE_CELL,
)


def _zero_vertical_spacing(pf) -> None:
    pf.space_before = c.SPACE_BEFORE
    pf.space_after = c.SPACE_AFTER


def apply_paragraph_format(paragraph: Paragraph, category: str) -> None:
    """Применить интервалы/отступ/выравнивание согласно категории абзаца.

    heading / empty / other / unknown — paragraph_format не трогаем
    (заголовкам гарнитуру/полужирность наводят runs.py и styles.py).
    """
    pf = paragraph.paragraph_format

    if category == CATEGORY_BODY:
        # Присвоение float принудительно ставит lineRule=auto, перебивая EXACTLY (подв. камень №9).
        pf.line_spacing = c.LINE_SPACING_BODY
        _zero_vertical_spacing(pf)
        pf.first_line_indent = c.FIRST_LINE_INDENT
        pf.alignment = c.ALIGN_BODY

    elif category == CATEGORY_LIST_ITEM:
        pf.line_spacing = c.LINE_SPACING_BODY
        _zero_vertical_spacing(pf)
        # first_line_indent НЕ трогаем — отступ списка задаётся нумерацией.

    elif category == CATEGORY_TABLE_CELL:
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
        _zero_vertical_spacing(pf)
        pf.first_line_indent = c.NO_FIRST_LINE_INDENT
        # Выравнивание ячеек не навязываем.

    elif category == CATEGORY_CAPTION:
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
        _zero_vertical_spacing(pf)
        pf.first_line_indent = c.NO_FIRST_LINE_INDENT
        # Выравнивание подписи не навязываем агрессивно.
