from pathlib import Path
import sys


def _get_base_dir() -> Path:
    """Return app base directory for source and PyInstaller builds.

    In source mode:
        project root

    In PyInstaller one-folder mode:
        folder containing the EXE

    This keeps data, exports, and projects writable beside the EXE.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent

    return Path(__file__).resolve().parent.parent


BASE_DIR = _get_base_dir()
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
