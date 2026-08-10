from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
PROJECTS_DIR = BASE_DIR / "projects"
EXPORTS_DIR = BASE_DIR / "exports"
BACKUPS_DIR = BASE_DIR / "backups"
PATCHES_DIR = BASE_DIR / "patches"
LOGS_DIR = BASE_DIR / "logs"
DOCS_DIR = BASE_DIR / "docs"

DB_PATH = DATA_DIR / "app.db"
SETTINGS_DIR = DATA_DIR / "settings"
RECENT_PROJECTS_PATH = SETTINGS_DIR / "recent_projects.json"