"""Тесты неблокирующих предупреждений: исправления/комментарии (ТЗ, подв. камень №19)."""

from __future__ import annotations

from gostdoc.cli import EXIT_OK, main
from gostdoc.formatter import format_document


def test_revisions_produce_warning(fixtures_dir, tmp_path):
    out = tmp_path / "rev.docx"
    warnings = format_document(str(fixtures_dir / "with_revisions.docx"), str(out))
    assert any("исправления" in w for w in warnings)


def test_clean_document_has_no_warnings(fixtures_dir, tmp_path):
    out = tmp_path / "clean.docx"
    warnings = format_document(str(fixtures_dir / "calibri_default.docx"), str(out))
    assert warnings == []


def test_cli_prints_warning_and_succeeds(fixtures_dir, tmp_path, capsys):
    out = tmp_path / "rev.docx"
    rc = main([str(fixtures_dir / "with_revisions.docx"), "-o", str(out)])
    assert rc == EXIT_OK
    assert "Предупреждение" in capsys.readouterr().out
