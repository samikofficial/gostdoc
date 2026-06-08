"""Общие фикстуры pytest: гарантируем наличие тестовых .docx и даём к ним доступ."""

from __future__ import annotations

from pathlib import Path

import pytest
from docx import Document

import make_fixtures

FIXTURES_DIR = make_fixtures.FIXTURES_DIR


@pytest.fixture(scope="session", autouse=True)
def _ensure_fixtures() -> None:
    """Сгенерировать недостающие фикстуры перед прогоном тестов."""
    make_fixtures.generate_all(force=False)


@pytest.fixture
def fixtures_dir() -> Path:
    return FIXTURES_DIR


@pytest.fixture
def load_fixture():
    """Открыть фикстуру по имени файла и вернуть Document."""

    def _load(name: str) -> Document:
        return Document(str(FIXTURES_DIR / name))

    return _load
