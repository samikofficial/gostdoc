"""Приведение стилей документа к ГОСТ: docDefaults, Normal, Heading*, Caption.

Стиль задаёт базу; прямое форматирование run'ов добивается отдельно (runs.py),
потому что прямое форматирование перебивает стиль (ТЗ, разд. 8.1).

Риски (ТЗ, разд. 8):
- №7 истинные дефолты лежат в docDefaults → правим и его, и Normal.
- №8 локализованные имена стилей (Заголовок 1, Подпись) → ищем по префиксам/синонимам.
"""

from __future__ import annotations

from docx.document import Document as _Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from . import constants as c
from . import xml_utils
from .profile import GOST, Profile


def _doc_defaults_rpr(document: _Document):
    """Вернуть (создав при отсутствии) элемент rPr из docDefaults/rPrDefault."""
    styles_el = document.styles.element
    dd = styles_el.find(qn("w:docDefaults"))
    if dd is None:
        dd = OxmlElement("w:docDefaults")
        styles_el.insert(0, dd)
    rprd = dd.find(qn("w:rPrDefault"))
    if rprd is None:
        rprd = OxmlElement("w:rPrDefault")
        dd.insert(0, rprd)
    rpr = rprd.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        rprd.append(rpr)
    return rpr


def _apply_font_to_rpr(rpr) -> None:
    xml_utils.set_run_fonts(rpr, c.FONT_NAME)
    xml_utils.set_run_size(rpr, c.FONT_SIZE_BODY)
    xml_utils.set_run_color_auto(rpr)


def _normalize_doc_defaults(document: _Document) -> None:
    _apply_font_to_rpr(_doc_defaults_rpr(document))


def _normalize_normal(document: _Document) -> None:
    normal = document.styles["Normal"]
    _apply_font_to_rpr(normal.element.get_or_add_rPr())
    pf = normal.paragraph_format
    pf.line_spacing = c.LINE_SPACING_BODY
    pf.space_before = c.SPACE_BEFORE
    pf.space_after = c.SPACE_AFTER
    pf.first_line_indent = c.FIRST_LINE_INDENT
    pf.alignment = c.ALIGN_BODY


def _is_heading_name(name: str | None) -> bool:
    return name is not None and any(name.startswith(p) for p in c.STYLE_PREFIXES_HEADING)


def _normalize_headings(document: _Document, profile: Profile) -> None:
    for style in document.styles:
        if style.type == WD_STYLE_TYPE.PARAGRAPH and _is_heading_name(style.name):
            _apply_font_to_rpr(style.element.get_or_add_rPr())
            style.font.bold = profile.bold_headings
            style.font.underline = False


def _normalize_captions(document: _Document) -> None:
    for style in document.styles:
        if style.name in c.STYLE_NAMES_CAPTION:
            _apply_font_to_rpr(style.element.get_or_add_rPr())


def normalize_styles(document: _Document, profile: Profile = GOST) -> None:
    """Привести docDefaults и стили Normal/Heading*/Caption к ГОСТ."""
    _normalize_doc_defaults(document)
    _normalize_normal(document)
    _normalize_headings(document, profile)
    _normalize_captions(document)
