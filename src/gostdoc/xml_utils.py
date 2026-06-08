"""Низкоуровневые lxml-хелперы для мест, где python-docx не хватает (ТЗ, разд. 7.3).

Здесь: установка всех слотов rFonts, размера, цвета auto на уровне rPr; поле PAGE
в колонтитуле. Все функции меняют существующие элементы, ничего не пересоздавая.
"""

from __future__ import annotations

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Word хранит поля и отступы в твипах (1/1440 дюйма), а python-docx отдаёт EMU.
# Из-за квантования Cm/Mm не возвращаются побайтно — сравнивать длины нужно в твипах.
EMU_PER_TWIP = 635


def emu_to_twips(value) -> int:
    """EMU → твипы (как хранит Word). Снимает ошибку квантования при сравнении длин."""
    return round(int(value) / EMU_PER_TWIP)


def lengths_equal(a, b) -> bool:
    """Равны ли две длины с точностью до твипа (None трактуется как 0)."""
    return emu_to_twips(a or 0) == emu_to_twips(b or 0)


# --- run properties (rPr) ---


def set_run_fonts(rpr, name: str) -> None:
    """Единая гарнитура во все слоты rFonts (подводный камень №6).

    ascii/hAnsi/cs ставим всегда; eastAsia — только если он уже задан.
    """
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn("w:ascii"), name)
    rfonts.set(qn("w:hAnsi"), name)
    rfonts.set(qn("w:cs"), name)
    if rfonts.get(qn("w:eastAsia")) is not None:
        rfonts.set(qn("w:eastAsia"), name)


def set_run_size(rpr, length) -> None:
    """Кегль run'а в w:sz и w:szCs. length — Length (Pt); значение — половинки пункта.

    szCs нужен для run'ов с флагом complex-script, иначе их размером управляет он, а не sz
    (по аналогии с подв. камнем №6 для гарнитуры). У CT_RPr нет get_or_add для szCs —
    создаём вручную сразу после sz, сохраняя порядок дочерних элементов по схеме.
    """
    half_points = str(int(round(length.pt * 2)))
    sz = rpr.get_or_add_sz()
    sz.set(qn("w:val"), half_points)
    szcs = rpr.find(qn("w:szCs"))
    if szcs is None:
        szcs = OxmlElement("w:szCs")
        sz.addnext(szcs)
    szcs.set(qn("w:val"), half_points)


def set_run_color_auto(rpr) -> None:
    """Цвет текста — auto. Снимаем theme-атрибуты, иначе они перебьют val."""
    color = rpr.get_or_add_color()
    color.set(qn("w:val"), "auto")
    for attr in ("w:themeColor", "w:themeShade", "w:themeTint"):
        key = qn(attr)
        if color.get(key) is not None:
            del color.attrib[key]


# --- нумерация страниц (поле PAGE) ---


def paragraph_has_page_field(paragraph) -> bool:
    """Есть ли в абзаце уже поле PAGE (для идемпотентности, подводный камень №14)."""
    for instr in paragraph._p.findall(".//" + qn("w:instrText")):
        if "PAGE" in (instr.text or "").upper():
            return True
    # Поле может быть оформлено как simple field.
    for fld in paragraph._p.findall(".//" + qn("w:fldSimple")):
        if "PAGE" in (fld.get(qn("w:instr")) or "").upper():
            return True
    return False


def add_page_number_field(paragraph) -> None:
    """Вставить поле { PAGE } в абзац (python-docx не умеет это штатно, подв. камень №4).

    Идемпотентно: если поле PAGE уже есть — ничего не делает.
    """
    if paragraph_has_page_field(paragraph):
        return
    run = paragraph.add_run()
    r = run._r

    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")

    r.append(begin)
    r.append(instr)
    r.append(end)
