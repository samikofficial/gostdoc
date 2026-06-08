"""Категоризация абзаца по его СТИЛЮ и контексту (ТЗ, разд. 7.2).

Категория определяет, какие правила оформления применит paragraphs.py.
Классифицируем по стилю и структуре, а не «на глаз» (распознавание неразмеченной
структуры — Фаза 2, не здесь).

Вероятные баги этого модуля и защита от них:
- numPr у списков лежит на уровне СТИЛЯ, а не абзаца → резолвим цепочку стилей.
- абзац с одним рисунком имеет пустой text → пустым его считать нельзя.
- русские имена стилей (Заголовок 1, Обычный) → сравнение по префиксам/синонимам.
- абзац в ячейке таблицы не отличить по самому абзацу → in_table передаём явно.
"""

from __future__ import annotations

import re

from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from . import constants as c

CATEGORY_BODY = "body"
CATEGORY_HEADING = "heading"
CATEGORY_CAPTION = "caption"
CATEGORY_TABLE_CELL = "table-cell"
CATEGORY_LIST_ITEM = "list-item"
CATEGORY_OTHER = "other"  # оглавление и прочее служебное — не трогаем содержимое
CATEGORY_EMPTY = "empty"

_EMBEDDED_TAGS = ("w:drawing", "w:object", "w:pict")

# Строка оглавления: номер страницы в конце И лидер до него — точечный (.. или …,
# бывает разбит пробелами/коротким) ИЛИ табуляция (w:tab → "\t" в тексте). Связка
# «лидер/таб + хвостовой номер» высокоточна — обычная проза/заголовок так не выглядят.
_TOC_LEADER_RUN = re.compile(r"[.…]{2,}")
_ENDS_WITH_PAGE_NO = re.compile(r"\d\s*$")


def _is_toc_line(text: str) -> bool:
    if not _ENDS_WITH_PAGE_NO.search(text):
        return False
    return bool(_TOC_LEADER_RUN.search(text)) or "\t" in text


def _is_empty(paragraph: Paragraph) -> bool:
    """Пустой абзац: нет текста И нет встроенных объектов (рисунок/OLE/pict)."""
    if paragraph.text.strip():
        return False
    for tag in _EMBEDDED_TAGS:
        if paragraph._p.findall(".//" + qn(tag)):
            return False
    return True


def _style_name(paragraph: Paragraph) -> str:
    style = paragraph.style
    name = getattr(style, "name", None)
    return name or ""


def _element_has_numpr(element) -> bool:
    ppr = element.find(qn("w:pPr"))
    return ppr is not None and ppr.find(qn("w:numPr")) is not None


def _has_numbering(paragraph: Paragraph) -> bool:
    """numPr на самом абзаце ИЛИ в цепочке его стилей (basedOn)."""
    if _element_has_numpr(paragraph._p):
        return True
    style = paragraph.style
    seen: set[int] = set()
    while style is not None and id(style) not in seen:
        seen.add(id(style))
        if _element_has_numpr(style._element):
            return True
        style = style.base_style
    return False


def _starts_with_any(name: str, prefixes: tuple[str, ...]) -> bool:
    return any(name.startswith(p) for p in prefixes)


def classify_paragraph(paragraph: Paragraph, in_table: bool) -> str:
    """Вернуть категорию абзаца. Порядок проверок задаёт приоритет правил."""
    if _is_empty(paragraph):
        return CATEGORY_EMPTY

    name = _style_name(paragraph)
    if name and _starts_with_any(name, c.STYLE_PREFIXES_TOC):
        return CATEGORY_OTHER
    # Неразмеченная строка оглавления (лидер + номер страницы) — защищаем от body-правил.
    if _is_toc_line(paragraph.text):
        return CATEGORY_OTHER
    if name and _starts_with_any(name, c.STYLE_PREFIXES_HEADING):
        return CATEGORY_HEADING
    if name in c.STYLE_NAMES_CAPTION:
        return CATEGORY_CAPTION
    if _has_numbering(paragraph):
        return CATEGORY_LIST_ITEM
    if in_table:
        return CATEGORY_TABLE_CELL
    # Normal/Body Text или неизвестный безопасный стиль (вне таблицы, без numPr,
    # непустой) — относим к телу. Консервативно: служебное уже отсеяно выше.
    return CATEGORY_BODY
