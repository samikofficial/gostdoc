"""Конвейер: открыть → привести оформление к ГОСТ → сохранить (ТЗ, разд. 7.5).

Главный инвариант: меняем только оформление. Текст тела (абзацы + ячейки таблиц)
до и после обработки совпадает побайтно; единственный добавленный текст — поле PAGE
в колонтитуле (в тело не входит).

Риски (ТЗ, разд. 8):
- №12,13 не-.docx/.doc/повреждён/запаролен → понятная ошибка, без трассировки.
- полнота обхода → рекурсивно обходим тело, включая вложенные таблицы.
- №16 рисунки в run'ах → run'ы не пересоздаём (см. runs.py).
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.document import Document as _Document
from docx.table import Table

from . import classify, constants as c, paragraphs, runs, sections, styles
from .xml_utils import lengths_equal, paragraph_has_page_field

_OLE2_MAGIC = b"\xd0\xcf\x11\xe0"
_ZIP_MAGIC = b"PK"


class GostDocError(Exception):
    """Понятная пользователю ошибка (плохой формат, повреждённый/запароленный файл)."""


def _validate_and_open(input_path: str) -> _Document:
    path = Path(input_path)
    if not path.exists():
        raise GostDocError(f"Файл не найден: {input_path}")

    ext = path.suffix.lower()
    if ext == ".doc":
        raise GostDocError("Формат .doc не поддерживается, пересохраните файл в .docx.")
    if ext != ".docx":
        raise GostDocError("Это не файл .docx.")

    signature = path.read_bytes()[:4]
    if signature == _OLE2_MAGIC:
        raise GostDocError(
            "Файл .docx в двоичном формате (повреждён или защищён паролем). "
            "Снимите защиту и пересохраните."
        )
    if signature[:2] != _ZIP_MAGIC:
        raise GostDocError("Файл повреждён или не является документом .docx.")

    try:
        return Document(str(path))
    except Exception as exc:  # noqa: BLE001 — переводим в понятную ошибку, не глотаем
        raise GostDocError(
            f"Не удалось открыть документ (повреждён или защищён паролем): {input_path}"
        ) from exc


def _normalize_runs_in_paragraph(paragraph, category: str) -> None:
    is_heading = category == classify.CATEGORY_HEADING
    for run in paragraph.runs:
        if is_heading:
            runs.normalize_heading_run(run)
        else:
            runs.normalize_run(run)


def _process_paragraph(paragraph, in_table: bool) -> None:
    category = classify.classify_paragraph(paragraph, in_table)
    paragraphs.apply_paragraph_format(paragraph, category)
    _normalize_runs_in_paragraph(paragraph, category)


def _process_table(table: Table) -> None:
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                _process_paragraph(paragraph, in_table=True)
            for nested in cell.tables:  # вложенные таблицы
                _process_table(nested)


def _normalize_body(document: _Document) -> None:
    for paragraph in document.paragraphs:
        _process_paragraph(paragraph, in_table=False)
    for table in document.tables:
        _process_table(table)


def _normalize_header_footer_fonts(document: _Document) -> None:
    """Гарнитуру/кегль колонтитулов (включая поле PAGE) — к ГОСТ; формат не трогаем."""
    for section in document.sections:
        for part in (
            section.header,
            section.footer,
            section.first_page_header,
            section.first_page_footer,
        ):
            for paragraph in part.paragraphs:
                for run in paragraph.runs:
                    runs.normalize_run(run)


def extract_body_text(document: _Document) -> str:
    """Чистый текст тела: абзацы документа + абзацы ячеек таблиц (без колонтитулов).

    Общий хелпер для теста инварианта и проверок (подводный камень №15).
    """
    parts: list[str] = []

    def collect_table(table: Table) -> None:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    parts.append(paragraph.text)
                for nested in cell.tables:
                    collect_table(nested)

    for paragraph in document.paragraphs:
        parts.append(paragraph.text)
    for table in document.tables:
        collect_table(table)
    return "\n".join(parts)


def format_document(input_path: str, output_path: str) -> str:
    """Привести оформление input_path к ГОСТ 7.32-2017 и сохранить в output_path."""
    document = _validate_and_open(input_path)
    styles.normalize_styles(document)
    sections.normalize_margins(document)
    sections.setup_page_numbering(document)
    _normalize_body(document)
    _normalize_header_footer_fonts(document)
    document.save(output_path)
    return output_path


def check_compliance(document: _Document) -> list[str]:
    """Вернуть список несоответствий ГОСТ-константам (пусто = соответствует).

    Проверяем верифицируемые ключевые параметры (поля, шрифт Normal, нумерацию).
    """
    problems: list[str] = []
    for i, section in enumerate(document.sections, start=1):
        if not lengths_equal(section.left_margin, c.MARGIN_LEFT):
            problems.append(f"Секция {i}: левое поле не равно 30 мм.")
        if not lengths_equal(section.right_margin, c.MARGIN_RIGHT):
            problems.append(f"Секция {i}: правое поле не равно 15 мм.")
        if not lengths_equal(section.top_margin, c.MARGIN_TOP):
            problems.append(f"Секция {i}: верхнее поле не равно 20 мм.")
        if not lengths_equal(section.bottom_margin, c.MARGIN_BOTTOM):
            problems.append(f"Секция {i}: нижнее поле не равно 20 мм.")

    normal = document.styles["Normal"]
    if normal.font.name != c.FONT_NAME:
        problems.append("Стиль «Обычный»: шрифт не Times New Roman.")
    if normal.font.size != c.FONT_SIZE_BODY:
        problems.append("Стиль «Обычный»: кегль не равен 14 пт.")

    if document.sections:
        first = document.sections[0]
        if not first.different_first_page_header_footer:
            problems.append("Не включён особый колонтитул первой страницы (номер на титуле).")
        if not any(paragraph_has_page_field(p) for p in first.footer.paragraphs):
            problems.append("В колонтитуле нет поля номера страницы.")
    return problems


def check_file(input_path: str) -> list[str]:
    """Открыть документ (с валидацией формата) и проверить соответствие ГОСТ."""
    return check_compliance(_validate_and_open(input_path))
