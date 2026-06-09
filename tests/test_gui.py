"""Тесты логики GUI (format_file) — без запуска tkinter (UI не нужен для проверки логики)."""

from __future__ import annotations

import shutil
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Mm

from gostdoc.gui import PAGE_NUMBER_LABELS, default_output, format_file
from gostdoc.profile import PageNumberPosition
from gostdoc.xml_utils import lengths_equal, paragraph_has_page_field


def test_format_file_applies_options(fixtures_dir, tmp_path):
    out = tmp_path / "o.docx"
    result, _ = format_file(
        str(fixtures_dir / "unmarked_structure.docx"),
        str(out),
        detect_structure=True,
        margins="20,20,20,20",
        page_number="bottom-right",
        page_number_size_pt=12,
        bold_headings=False,
    )
    assert result == str(out)
    doc = Document(str(out))
    assert lengths_equal(doc.sections[0].left_margin, Mm(20))
    footer = doc.sections[0].footer
    page_paras = [p for p in footer.paragraphs if paragraph_has_page_field(p)]
    assert page_paras and page_paras[0].alignment == WD_ALIGN_PARAGRAPH.RIGHT


def test_format_file_default_output(fixtures_dir, tmp_path):
    src = tmp_path / "вход.docx"
    shutil.copy(fixtures_dir / "calibri_default.docx", src)
    result, _ = format_file(str(src))
    assert result.endswith("вход.gost.docx")
    assert Path(result).exists()
    assert src.exists()  # исходник не перезаписан


def test_default_output_helper():
    assert default_output("C:/x/Диплом.docx").endswith("Диплом.gost.docx")


def test_page_number_labels_are_valid():
    for value in PAGE_NUMBER_LABELS.values():
        PageNumberPosition(value)  # не бросает — значит метка валидна
