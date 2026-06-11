"""Application configuration constants for Excel Memory Agent.

This module intentionally contains constants only. Values are read once at
import time and can be overridden through environment variables, which keeps
the rest of the codebase simple and easy to test.

Runtime state should live outside cloud-synced folders such as OneDrive to
reduce the risk of sync locks and conflict copies.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Final

EMA_HOME: Final[Path] = Path(os.getenv("EMA_HOME", Path.home() / ".ema"))
"""Root directory for EMA runtime state."""

REGISTRY_PATH: Final[Path] = EMA_HOME / "registry.json"
"""Location of the JSON workbook registry."""

BACKUP_DIR: Final[Path] = EMA_HOME / "backups"
"""Directory where workbook backups are stored."""

LOG_PATH: Final[Path] = EMA_HOME / "ema.log"
"""Path to the EMA log file."""

LOCK_SUFFIX: Final[str] = "~$"
"""Excel lock-file prefix marker used to detect open workbooks."""

MAX_BACKUPS_PER_WORKBOOK: Final[int] = 20
"""Maximum number of retained backups per workbook."""

SCHEMA_VERSION: Final[int] = 1
"""Persisted workbook schema version."""

LLM_ENABLED: Final[bool] = os.getenv("EMA_LLM_ENABLED", "false").lower() == "true"
"""Whether the optional local-LLM flow is enabled."""

LLM_MODEL: Final[str] = os.getenv("EMA_LLM_MODEL", "qwen3:4b")
"""Default local model name used by the LLM integration."""

OLLAMA_HOST: Final[str] = os.getenv("EMA_OLLAMA_HOST", "http://localhost:11434")
"""Base URL for the local Ollama server."""

REQUIRE_CONFIRM_FOR_LLM: Final[bool] = True
"""Require preview/confirmation before any LLM-sourced write."""

EXPOSE_TEXT_TOOL: Final[bool] = False
"""Whether to expose the optional MCP text-append tool."""

__all__ = [
    "BACKUP_DIR",
    "EMA_HOME",
    "EXPOSE_TEXT_TOOL",
    "LLM_ENABLED",
    "LLM_MODEL",
    "LOCK_SUFFIX",
    "LOG_PATH",
    "MAX_BACKUPS_PER_WORKBOOK",
    "OLLAMA_HOST",
    "REGISTRY_PATH",
    "REQUIRE_CONFIRM_FOR_LLM",
    "SCHEMA_VERSION",
]
