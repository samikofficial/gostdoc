"""Поля всех секций и сквозная нумерация страниц (ТЗ, разд. 7.5).

Риски (ТЗ, разд. 8):
- №3 несколько секций → обходим document.sections целиком, а не [0].
- №4 поля PAGE нет в API → вставляем через xml_utils.
- №5 титул без номера, но в сквозной нумерации → different first page + пустой
  первый колонтитул, поле PAGE в основном; старт нумерации не сбрасываем.
- №14 идемпотентность → add_page_number_field не дублирует поле.
- №18 зеркальные поля/переплёт → снимаем mirrorMargins и обнуляем gutter.
"""

from __future__ import annotations

from docx.document import Document as _Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

from . import constants as c
from . import xml_utils


def _disable_mirror_margins(document: _Document) -> None:
    settings_el = document.settings.element
    mirror = settings_el.find(qn("w:mirrorMargins"))
    if mirror is not None:
        settings_el.remove(mirror)


def normalize_margins(document: _Document) -> None:
    """Выставить ГОСТ-поля каждой секции; снять зеркальные поля и переплёт."""
    _disable_mirror_margins(document)
    for section in document.sections:
        section.left_margin = c.MARGIN_LEFT
        section.right_margin = c.MARGIN_RIGHT
        section.top_margin = c.MARGIN_TOP
        section.bottom_margin = c.MARGIN_BOTTOM
        section.gutter = 0


def _set_centered_page_field(footer) -> None:
    para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    xml_utils.add_page_number_field(para)


def setup_page_numbering(document: _Document) -> None:
    """Сквозная нумерация внизу по центру; на первой странице номер не печатается."""
    sections = document.sections
    if not sections:
        return
    first = sections[0]
    first.different_first_page_header_footer = True
    footer = first.footer
    footer.is_linked_to_previous = False  # создаём собственный колонтитул
    _set_centered_page_field(footer)
    # Остальные секции наследуют основной колонтитул — нумерация остаётся сквозной.
    for section in sections[1:]:
        section.footer.is_linked_to_previous = True
