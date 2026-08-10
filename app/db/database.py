import sqlite3
from pathlib import Path

from app.config import DB_PATH
from app.config import DATA_DIR


def initialize_database() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    schema_path = Path(__file__).resolve().parent / "schema.sql"

    with sqlite3.connect(DB_PATH) as conn:
        schema = schema_path.read_text(encoding="utf-8")
        conn.executescript(schema)
        conn.commit()