"""Тесты Фазы 2: распознавание неразмеченной структуры (ТЗ, разд. 12).

Фикстура unmarked_structure кодирует реальные паттерны и ложные срабатывания,
найденные на 5 настоящих ВКР.
"""

from __future__ import annotations

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

from gostdoc.classify import CATEGORY_HEADING, classify_paragraph
from gostdoc.detect import detect_and_mark
from gostdoc.formatter import extract_body_text, format_document


def _by_text(doc, needle):
    for p in doc.paragraphs:
        if p.text.strip().startswith(needle):
            return p
    raise AssertionError(f"не найден абзац: {needle!r}")


def _exact(doc, text):
    for p in doc.paragraphs:
        if p.text.strip() == text:
            return p
    raise AssertionError(f"не найден абзац с точным текстом: {text!r}")


def test_structural_elements_detected(load_fixture):
    doc = load_fixture("unmarked_structure.docx")
    detect_and_mark(doc)
    for word in ("ВВЕДЕНИЕ", "Заключение", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"):
        p = _exact(doc, word)
        assert classify_paragraph(p, in_table=False) == CATEGORY_HEADING
        assert p.alignment == WD_ALIGN_PARAGRAPH.CENTER
        assert p.paragraph_format.page_break_before is True
        # Прописные через w:caps, текст не изменён.
        assert any(r._r.find(qn("w:rPr")).find(qn("w:caps")) is not None for r in p.runs)


def test_numbered_multilevel_headings_detected(load_fixture):
    doc = load_fixture("unmarked_structure.docx")
    detect_and_mark(doc)
    assert _by_text(doc, "1.1 История возникновения").style.name == "Heading 2"
    assert _by_text(doc, "1.1.1 Подробности").style.name == "Heading 3"


def test_false_positives_not_marked(load_fixture):
    doc = load_fixture("unmarked_structure.docx")
    detect_and_mark(doc)
    # Поле-задание и пункт списка (одноуровневые «1.») — НЕ заголовки.
    assert _by_text(doc, "1. Тема:").style.name == "Normal"
    assert _by_text(doc, "1. Кубики Кооса").style.name == "Normal"
    # Обычный абзац тела — не тронут.
    assert _by_text(doc, "Текст введения").style.name == "Normal"


def test_toc_lines_protected_not_marked(load_fixture):
    doc = load_fixture("unmarked_structure.docx")
    detect_and_mark(doc)
    # Строки оглавления (лидеры + номер страницы) не должны стать заголовками.
    assert _by_text(doc, "ВВЕДЕНИЕ ......").style.name == "Normal"
    assert _by_text(doc, "1.1 История вопроса .....").style.name == "Normal"


def test_captions_detected(load_fixture):
    doc = load_fixture("unmarked_structure.docx")
    detect_and_mark(doc)
    fig = _by_text(doc, "Рисунок 1 — Схема")
    fig_abbr = _by_text(doc, "Рис. 1.1.")
    tbl = _by_text(doc, "Таблица 2 —")
    assert fig.style.name == "Caption"
    assert fig.alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert fig_abbr.style.name == "Caption"
    assert tbl.style.name == "Caption"
    assert tbl.alignment == WD_ALIGN_PARAGRAPH.LEFT


def test_figure_reference_in_text_not_caption(load_fixture):
    doc = load_fixture("unmarked_structure.docx")
    detect_and_mark(doc)
    # «Рисунок 1 показывает …» — ссылка в тексте, не подпись.
    assert _by_text(doc, "Рисунок 1 показывает").style.name == "Normal"


def test_detect_preserves_text_invariant(fixtures_dir, tmp_path):
    src = fixtures_dir / "unmarked_structure.docx"
    before = extract_body_text(Document(str(src)))
    out = tmp_path / "out.docx"
    format_document(str(src), str(out), detect_structure=True)
    after = extract_body_text(Document(str(out)))
    assert before == after  # caps — это отображение, текст не меняется


def test_detect_log_returned(fixtures_dir, tmp_path):
    out = tmp_path / "out.docx"
    messages = format_document(str(fixtures_dir / "unmarked_structure.docx"), str(out), detect_structure=True)
    assert any("Фаза 2" in m for m in messages)


def test_default_does_not_detect(fixtures_dir, tmp_path):
    # Без флага Фаза 2 не работает: нумерованный заголовок остаётся Normal.
    out = tmp_path / "out.docx"
    format_document(str(fixtures_dir / "unmarked_structure.docx"), str(out))
    doc = Document(str(out))
    assert _by_text(doc, "1.1 История возникновения").style.name == "Normal"
