"""CLI gostdoc (ТЗ, разд. 7.4).

    gostdoc INPUT.docx [-o OUTPUT.docx] [--check]

Коды выхода: 0 — успех/соответствует; 1 — есть несоответствия (--check);
2 — ошибка формата/открытия (понятное сообщение, без трассировки).
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .formatter import GostDocError, check_file, format_document
from .profile import PageNumberPosition, build_profile

EXIT_OK = 0
EXIT_NONCOMPLIANT = 1
EXIT_ERROR = 2


def _setup_console() -> None:
    """Сделать вывод в консоль безопасным для Unicode (рус. текст, стрелки и т.п.).

    На Windows консоль по умолчанию cp1251 и падает на символах вроде «→».
    Переключаем кодовую страницу на UTF-8 и поток на utf-8 с заменой непечатаемых.
    """
    if sys.platform == "win32":
        try:
            import ctypes

            ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        except Exception:  # noqa: BLE001 — косметика вывода, не критично
            pass
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:
            continue
        try:
            reconfigure(encoding="utf-8", errors="replace")
        except Exception:  # noqa: BLE001
            pass


def _default_output(input_path: str) -> str:
    """INPUT.docx → INPUT.gost.docx (исходник не перезаписываем)."""
    p = Path(input_path)
    return str(p.parent / (p.stem + ".gost.docx"))


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="gostdoc",
        description="Приведение оформления .docx к ГОСТ 7.32-2017 без изменения текста.",
    )
    parser.add_argument("input", help="входной файл .docx")
    parser.add_argument("-o", "--output", help="куда сохранить результат (по умолчанию INPUT.gost.docx)")
    parser.add_argument(
        "--check",
        action="store_true",
        help="не записывать файл, только проверить соответствие ГОСТ",
    )
    parser.add_argument(
        "--detect-structure",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="распознавать неразмеченные заголовки/структурные элементы (по умолчанию вкл; "
        "--no-detect-structure отключает, оставляя только чистое v0-форматирование)",
    )
    # Переопределения под методичку (дефолт — базовый ГОСТ 7.32-2017).
    overrides = parser.add_argument_group("переопределения под методичку")
    overrides.add_argument(
        "--margins",
        metavar="Л,П,В,Н",
        help="поля в мм через запятую (лево,право,верх,низ), например: 30,10,20,20",
    )
    overrides.add_argument(
        "--page-number",
        choices=[p.value for p in PageNumberPosition],
        help="позиция номера страницы (по умолчанию bottom-center; none — без нумерации)",
    )
    overrides.add_argument(
        "--page-number-size", type=int, metavar="ПТ", help="кегль номера страницы в пунктах"
    )
    overrides.add_argument(
        "--no-bold-headings",
        dest="bold_headings",
        action="store_false",
        default=None,
        help="не делать заголовки полужирными (некоторые методички запрещают)",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    _setup_console()
    args = _build_parser().parse_args(argv)
    try:
        profile = build_profile(
            margins=args.margins,
            page_number=args.page_number,
            page_number_size_pt=args.page_number_size,
            bold_headings=args.bold_headings,
        )
    except ValueError as exc:
        print(f"Ошибка: {exc}", file=sys.stderr)
        return EXIT_ERROR

    try:
        if args.check:
            problems = check_file(args.input, profile)
            if problems:
                print("Несоответствия профилю ГОСТ 7.32-2017:")
                for problem in problems:
                    print(f"  - {problem}")
                return EXIT_NONCOMPLIANT
            print("Документ соответствует профилю ГОСТ 7.32-2017.")
            return EXIT_OK

        output = args.output or _default_output(args.input)
        messages = format_document(
            args.input, output, detect_structure=args.detect_structure, profile=profile
        )
        print(f"Готово: {output}")
        for message in messages:
            print(message)
        return EXIT_OK
    except GostDocError as exc:
        print(f"Ошибка: {exc}", file=sys.stderr)
        return EXIT_ERROR


if __name__ == "__main__":
    sys.exit(main())
