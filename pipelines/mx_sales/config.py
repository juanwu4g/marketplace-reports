"""Paths and pipeline-wide settings.

Everything the pipeline touches on disk is resolved from here, so relocating the
project (for example moving the code up next to ``downloads/``) is a one-line change.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
# The code lives in <repo>/pipelines; the data does not. The monthly run always
# sets MX_SALES_* (platforms/mercadolibre/transform.data_env, from core/layout).
# Run by hand without them, the pipeline falls back to the old data folder
# <repo>/downloads — never to its own code folder.
DEFAULT_DATA_ROOT = PROJECT_ROOT.parent / "downloads"


@dataclass(frozen=True)
class Settings:
    """Resolved locations for the raw drops and the processed outputs."""

    raw_dir: Path
    processed_dir: Path
    warehouse_path: Path

    @classmethod
    def from_env(cls) -> "Settings":
        root = Path(os.environ.get("MX_SALES_ROOT", DEFAULT_DATA_ROOT)).resolve()
        return cls(
            raw_dir=Path(os.environ.get("MX_SALES_RAW_DIR", root)),
            processed_dir=Path(os.environ.get("MX_SALES_PROCESSED_DIR", root / "data" / "processed")),
            warehouse_path=Path(os.environ.get("MX_SALES_WAREHOUSE", root / "data" / "mx_sales.duckdb")),
        )


SETTINGS = Settings.from_env()

# Directories under raw_dir that hold code or outputs rather than downloads.
EXCLUDED_DIRS = {"mx_sales", "data", "tests", ".git", ".venv", "__pycache__", ".pytest_cache"}
