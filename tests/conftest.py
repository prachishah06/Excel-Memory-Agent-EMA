"""Shared pytest fixtures for EMA test workbooks."""

from __future__ import annotations

import importlib
from pathlib import Path
from typing import Generator

import pytest

import ema.config as config
from fixtures.make_fixtures import (
    make_formula,
    make_offset_headers,
    make_plain,
    make_table,
)


@pytest.fixture
def ema_home(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> Generator[Path, None, None]:
    """Point `EMA_HOME` at a temporary test directory for the duration of a test."""

    ema_home_path = tmp_path / "ema-home"
    monkeypatch.setenv("EMA_HOME", str(ema_home_path))
    importlib.reload(config)

    try:
        yield ema_home_path
    finally:
        importlib.reload(config)


@pytest.fixture
def plain_wb(tmp_path: Path) -> Path:
    """Create a simple tabular workbook fixture."""

    return make_plain(tmp_path / "plain.xlsx")


@pytest.fixture
def table_wb(tmp_path: Path) -> Path:
    """Create a workbook fixture whose data lives in an Excel Table."""

    return make_table(tmp_path / "table.xlsx")


@pytest.fixture
def formula_wb(tmp_path: Path) -> Path:
    """Create a workbook fixture that includes a formula column."""

    return make_formula(tmp_path / "formula.xlsx")


@pytest.fixture
def offset_wb(tmp_path: Path) -> Path:
    """Create a workbook fixture with headers below metadata rows."""

    return make_offset_headers(tmp_path / "offset.xlsx")
