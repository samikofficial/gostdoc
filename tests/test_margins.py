"""Тесты полей страницы для ВСЕХ секций (ТЗ, DoD + подв. камень №3)."""

from __future__ import annotations

from docx.oxml.ns import qn

from gostdoc import constants as c
from gostdoc.sections import normalize_margins
from gostdoc.xml_utils import lengths_equal


def test_all_sections_get_gost_margins(load_fixture):
    doc = load_fixture("multi_section.docx")
    assert len(doc.sections) == 2  # фикстура с двумя секциями и разными полями
    normalize_margins(doc)
    for s in doc.sections:
        assert lengths_equal(s.left_margin, c.MARGIN_LEFT)
        assert lengths_equal(s.right_margin, c.MARGIN_RIGHT)
        assert lengths_equal(s.top_margin, c.MARGIN_TOP)
        assert lengths_equal(s.bottom_margin, c.MARGIN_BOTTOM)
        assert lengths_equal(s.gutter, 0)


def test_mirror_margins_normalized(load_fixture):
    doc = load_fixture("mirror_margins.docx")
    # Фикстура реально с зеркальными полями (подв. камень №18).
    assert doc.settings.element.find(qn("w:mirrorMargins")) is not None
    normalize_margins(doc)
    assert doc.settings.element.find(qn("w:mirrorMargins")) is None
    for s in doc.sections:
        assert lengths_equal(s.left_margin, c.MARGIN_LEFT)
        assert lengths_equal(s.right_margin, c.MARGIN_RIGHT)
