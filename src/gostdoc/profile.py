"""Профиль оформления: переопределяемые параметры, расходящиеся между методичками.

Дефолт — базовый ГОСТ 7.32-2017 (значения из constants). Расходятся только поля,
позиция/кегль номера страницы и полужирность заголовков (см. анализ 10 методичек).
Всё остальное (шрифт, кегль тела, интервал, отступ) у методичек совпадает с ГОСТ,
поэтому в профиль не выносится (YAGNI).
"""

from __future__ import annotations

from dataclasses import dataclass, replace
from enum import Enum

from docx.shared import Length, Mm, Pt

from . import constants as c


class PageNumberPosition(str, Enum):
    BOTTOM_CENTER = "bottom-center"
    BOTTOM_RIGHT = "bottom-right"
    TOP_RIGHT = "top-right"
    NONE = "none"


@dataclass(frozen=True)
class Profile:
    """Неизменяемый набор параметров оформления. Менять — через replace()/build_profile."""

    margin_left: Length = c.MARGIN_LEFT
    margin_right: Length = c.MARGIN_RIGHT
    margin_top: Length = c.MARGIN_TOP
    margin_bottom: Length = c.MARGIN_BOTTOM
    page_number: PageNumberPosition = PageNumberPosition.BOTTOM_CENTER
    page_number_size: Length = c.FONT_SIZE_BODY
    bold_headings: bool = True


# Базовый профиль — чистый ГОСТ 7.32-2017.
GOST = Profile()


def parse_margins(text: str) -> tuple[Length, Length, Length, Length]:
    """«30,10,20,20» (мм, порядок Л,П,В,Н) → кортеж длин EMU. Бросает ValueError."""
    parts = [p.strip() for p in text.split(",")]
    if len(parts) != 4:
        raise ValueError("Поля задаются как Л,П,В,Н в мм, например: 30,10,20,20")
    try:
        left, right, top, bottom = (int(p) for p in parts)
    except ValueError as exc:
        raise ValueError("Поля должны быть целыми числами (мм), например: 30,10,20,20") from exc
    return (Mm(left), Mm(right), Mm(top), Mm(bottom))


def build_profile(
    margins: str | None = None,
    page_number: str | None = None,
    page_number_size_pt: int | None = None,
    bold_headings: bool | None = None,
) -> Profile:
    """Собрать профиль поверх ГОСТ из CLI-переопределений (None — оставить дефолт ГОСТ)."""
    changes: dict = {}
    if margins is not None:
        left, right, top, bottom = parse_margins(margins)
        changes.update(
            margin_left=left, margin_right=right, margin_top=top, margin_bottom=bottom
        )
    if page_number is not None:
        changes["page_number"] = PageNumberPosition(page_number)
    if page_number_size_pt is not None:
        changes["page_number_size"] = Pt(page_number_size_pt)
    if bold_headings is not None:
        changes["bold_headings"] = bold_headings
    return replace(GOST, **changes)
