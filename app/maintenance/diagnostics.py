import sys
from pathlib import Path


def run_basic_diagnostics() -> dict:
    return {
        "python_version": sys.version,
        "current_working_directory": str(Path.cwd()),
        "status": "ok",
    }