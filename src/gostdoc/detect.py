"""Фаза 2: распознавание НЕразмеченной структуры (ТЗ, разд. 12).

Принципы (ТЗ): только высокоуверенные сигналы; авторазметка обратимая и логируемая;
развиваем от реальных провалов на настоящих ВКР, а не спекулятивно.

Распознаём (опционально, по флагу):
- структурные элементы по ключевым словам (ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ, СПИСОК... ) →
  оформляем прописными (w:caps, БЕЗ изменения текста), по центру, с новой страницы;
- главы/разделы «ГЛАВА N …», «РАЗДЕЛ N …» → Heading 1;
- многоуровневые нумерованные заголовки «N.N …», «N.N.N …» → Heading 2/3 по глубине.

Сознательно НЕ трогаем (реальные ложные срабатывания из ВКР):
- строки оглавления (лидеры + номер страницы) — classify уже относит их к служебным;
- одноуровневые «1. …» — это поля-задания и пункты списков, а не заголовки;
- длинные абзацы (заголовок — короткий).
"""

from __future__ import annotations

import re

from docx.document import Document as _Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from . import constants as c
from .classify import (
    CATEGORY_EMPTY,
    CATEGORY_HEADING,
    CATEGORY_OTHER,
    _has_numbering,
    classify_paragraph,
)

# Заголовок — короткий (это название, не предложение).
_MAX_HEADING_LEN = 100
_MAX_STRUCT_LEN = 70

_LEADER = re.compile(r"[.…]{3,}")
_TRAILING_PAGE = re.compile(r"\s+\d{1,4}\s*$")
# Многоуровневый номер: минимум два уровня (N.N…), отсекает одноуровневые поля/списки.
_NUMBERED_MULTILEVEL = re.compile(r"^\s*(\d+(?:\.\d+)+)\.?\s+\S")
_CHAPTER = re.compile(r"^\s*(ГЛАВА|РАЗДЕЛ)\s+\d", re.IGNORECASE)

# Подпись: ключевое слово + номер + разделитель (— . :) ИЛИ конец строки.
# Разделитель отсекает ссылки в тексте («Рисунок 1 показывает …»).
_FIG_CAPTION = re.compile(r"^\s*(Рисунок|Рис\.)\s*\d+(?:\.\d+)*\s*([—–\-.:]|$)")
_TBL_CAPTION = re.compile(r"^\s*(Таблица|Табл\.)\s*\d+(?:\.\d+)*\s*([—–\-.:]|$)")
_MAX_CAPTION_LEN = 200


def _normalize_struct_text(text: str) -> str:
    cleaned = _LEADER.sub(" ", text)
    cleaned = _TRAILING_PAGE.sub("", cleaned)
    return cleaned.strip().rstrip(".:").strip().upper()


def _is_struct_word(text: str) -> bool:
    norm = _normalize_struct_text(text)
    return norm in c.STRUCT_WORDS or any(norm.startswith(p) for p in c.STRUCT_PREFIXES)


def _ensure_style(document: _Document, name: str):
    try:
        return document.styles[name]
    except KeyError:
        style = document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = document.styles["Normal"]
        return style


def _apply_display_caps(paragraph: Paragraph) -> None:
    """Отображать прописными через w:caps — текст НЕ меняем (инвариант сохранён)."""
    for run in paragraph.runs:
        rpr = run._r.get_or_add_rPr()
        caps = rpr.find(qn("w:caps"))
        if caps is None:
            caps = OxmlElement("w:caps")
            rpr.append(caps)
        caps.set(qn("w:val"), "1")


def _set_style(document: _Document, paragraph: Paragraph, name: str) -> None:
    _ensure_style(document, name)
    paragraph.style = document.styles[name]


def _mark_structural(document: _Document, paragraph: Paragraph) -> None:
    _set_style(document, paragraph, "Heading 1")
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.page_break_before = True
    _apply_display_caps(paragraph)


def _numbered_level(text: str) -> int:
    token = text.split()[0].rstrip(".")
    return min(token.count(".") + 1, 3)


def _mark_caption(document: _Document, paragraph: Paragraph, centered: bool) -> None:
    _set_style(document, paragraph, "Caption")
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER if centered else WD_ALIGN_PARAGRAPH.LEFT


def detect_and_mark(document: _Document) -> list[str]:
    """Распознать и разметить структуру стилями. Вернуть лог изменений (обратимых)."""
    log: list[str] = []
    for paragraph in document.paragraphs:
        category = classify_paragraph(paragraph, in_table=False)
        if category in (CATEGORY_HEADING, CATEGORY_EMPTY, CATEGORY_OTHER):
            continue  # уже размечено или служебное (в т.ч. TOC-строки)
        text = paragraph.text.strip()
        if not text:
            continue

        if _is_struct_word(text) and len(text) <= _MAX_STRUCT_LEN:
            _mark_structural(document, paragraph)
            log.append(f"структурный элемент: {text[:45]!r} → по центру, прописными, с новой страницы")
            continue

        if _CHAPTER.match(text) and len(text) <= _MAX_HEADING_LEN:
            _set_style(document, paragraph, "Heading 1")
            log.append(f"заголовок главы: {text[:45]!r} → Heading 1")
            continue

        if (
            _NUMBERED_MULTILEVEL.match(text)
            and len(text) <= _MAX_HEADING_LEN
            and not _has_numbering(paragraph)
        ):
            level = _numbered_level(text)
            _set_style(document, paragraph, f"Heading {level}")
            log.append(f"нумерованный заголовок ур.{level}: {text[:45]!r} → Heading {level}")
            continue

        if len(text) <= _MAX_CAPTION_LEN and _FIG_CAPTION.match(text):
            _mark_caption(document, paragraph, centered=True)
            log.append(f"подпись рисунка: {text[:45]!r} → Caption (по центру)")
            continue

        if len(text) <= _MAX_CAPTION_LEN and _TBL_CAPTION.match(text):
            _mark_caption(document, paragraph, centered=False)
            log.append(f"наименование таблицы: {text[:45]!r} → Caption (слева)")

    return log
