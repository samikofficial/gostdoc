"""Тесты CLI: понятные ошибки и коды выхода (ТЗ, разд. 10 + подв. камни №12,13)."""

from __future__ import annotations

import shutil

from gostdoc.cli import EXIT_ERROR, EXIT_NONCOMPLIANT, EXIT_OK, main
from gostdoc.formatter import format_document


def test_legacy_doc_extension_error(fixtures_dir, capsys):
    rc = main([str(fixtures_dir / "legacy.doc")])
    assert rc == EXIT_ERROR
    err = capsys.readouterr().err
    assert ".doc" in err
    assert "Traceback" not in err


def test_not_docx_extension(tmp_path, capsys):
    f = tmp_path / "notes.txt"
    f.write_text("просто текст", encoding="utf-8")
    rc = main([str(f)])
    assert rc == EXIT_ERROR
    assert "не файл .docx" in capsys.readouterr().err


def test_broken_docx(fixtures_dir, capsys):
    rc = main([str(fixtures_dir / "broken.docx")])
    assert rc == EXIT_ERROR
    err = capsys.readouterr().err
    assert "Ошибка" in err
    assert "Traceback" not in err


def test_missing_file(tmp_path, capsys):
    rc = main([str(tmp_path / "нет.docx")])
    assert rc == EXIT_ERROR
    assert "не найден" in capsys.readouterr().err


def test_check_dirty_returns_nonzero(fixtures_dir):
    assert main([str(fixtures_dir / "calibri_default.docx"), "--check"]) == EXIT_NONCOMPLIANT


def test_check_formatted_returns_zero(fixtures_dir, tmp_path):
    out = tmp_path / "ok.docx"
    format_document(str(fixtures_dir / "calibri_default.docx"), str(out))
    assert main([str(out), "--check"]) == EXIT_OK


def test_default_output_path(fixtures_dir, tmp_path, capsys):
    src = tmp_path / "report.docx"
    shutil.copy(fixtures_dir / "calibri_default.docx", src)
    rc = main([str(src)])
    assert rc == EXIT_OK
    assert (tmp_path / "report.gost.docx").exists()
    assert src.exists()  # исходник не перезаписан


def test_explicit_output_path(fixtures_dir, tmp_path):
    out = tmp_path / "custom.docx"
    rc = main([str(fixtures_dir / "calibri_default.docx"), "-o", str(out)])
    assert rc == EXIT_OK
    assert out.exists()
