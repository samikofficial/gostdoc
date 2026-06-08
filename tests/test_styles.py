"""Тесты стилей: docDefaults, Normal, Heading, Caption (ТЗ, разд. 8.7, 8.8 + DoD)."""

from __future__ import annotations

from docx.oxml.ns import qn
from docx.shared import Pt

from gostdoc import constants as c
from gostdoc.runs import normalize_heading_run
from gostdoc.styles import normalize_styles


def _docdefaults_rfonts_ascii(document):
    sel = document.styles.element
    rpr = sel.find(qn("w:docDefaults") + "/" + qn("w:rPrDefault") + "/" + qn("w:rPr"))
    assert rpr is not None
    rf = rpr.find(qn("w:rFonts"))
    return None if rf is None else rf.get(qn("w:ascii"))


def test_doc_defaults_font(load_fixture):
    doc = load_fixture("calibri_default.docx")
    normalize_styles(doc)
    assert _docdefaults_rfonts_ascii(doc) == c.FONT_NAME


def test_normal_style_font(load_fixture):
    doc = load_fixture("calibri_default.docx")
    normalize_styles(doc)
    normal = doc.styles["Normal"]
    assert normal.font.name == c.FONT_NAME
    assert normal.font.size == Pt(14)


def test_heading_styles_bold_no_underline(load_fixture):
    doc = load_fixture("styled_headings.docx")
    normalize_styles(doc)
    for name in ("Heading 1", "Heading 2"):
        st = doc.styles[name]
        assert st.font.bold is True
        assert st.font.underline is not True
        assert st.font.name == c.FONT_NAME


def test_normalize_heading_run_forces_bold_removes_underline(load_fixture):
    doc = load_fixture("styled_headings.docx")
    heading = doc.paragraphs[0]  # «Введение», в фикстуре подчёркнут и НЕ полужирный
    for run in heading.runs:
        normalize_heading_run(run)
    for run in heading.runs:
        assert run.bold is True
        assert run.underline is not True
        assert run.font.name == c.FONT_NAME
