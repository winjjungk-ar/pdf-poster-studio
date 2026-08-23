import sqlite3
from pathlib import Path
import sys

from app.config import DB_PATH
from app.config import DATA_DIR


EMBEDDED_SCHEMA_SQL = r"""
CREATE TABLE IF NOT EXISTS app_settings (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS recent_projects (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_path TEXT NOT NULL,
    opened_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS project_settings (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_name TEXT NOT NULL,
    pdf_path TEXT,
    output_mode TEXT DEFAULT 'no_cut',
    split_layout TEXT DEFAULT 'left_right',
    scale_percent REAL DEFAULT 96.0,
    overlap_mm REAL DEFAULT 5.0,
    safe_margin_mm REAL DEFAULT 5.0,
    add_cut_guides INTEGER DEFAULT 1,
    add_tile_numbers INTEGER DEFAULT 1,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS patch_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    version TEXT NOT NULL,
    title TEXT NOT NULL,
    applied_at TEXT NOT NULL,
    status TEXT NOT NULL,
    notes TEXT
);

CREATE TABLE IF NOT EXISTS audit_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    event_time TEXT NOT NULL,
    event_type TEXT NOT NULL,
    message TEXT NOT NULL
);
"""



def _resource_path(relative_path: str) -> Path:
    """Resolve resource path both in source and PyInstaller build."""
    base_path = getattr(sys, "_MEIPASS", None)

    if base_path:
        return Path(base_path) / relative_path

    # Source tree fallback
    return Path(__file__).resolve().parents[2] / relative_path




def initialize_database() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    schema = None
    tried_paths = []

    candidate_paths = [
        _resource_path("app/db/schema.sql"),
        Path(__file__).resolve().parent / "schema.sql",
    ]

    for schema_path in candidate_paths:
        tried_paths.append(str(schema_path))

        if schema_path.exists():
            schema = schema_path.read_text(encoding="utf-8")
            break

    # Fallback for PyInstaller builds where schema.sql is not available as a file.
    if schema is None:
        schema = EMBEDDED_SCHEMA_SQL

    if not schema or not schema.strip():
        raise FileNotFoundError(
            "schema.sql not found and embedded schema is empty. Tried: "
            + " | ".join(tried_paths)
        )

    with sqlite3.connect(DB_PATH) as conn:
        conn.executescript(schema)
        conn.commit()

