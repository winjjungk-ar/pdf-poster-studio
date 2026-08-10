import json
from pathlib import Path
from typing import Any


PROJECT_FILE_EXTENSION = ".a4a3proj"


class ProjectManager:
    def save_project(self, path: str | Path, data: dict[str, Any]) -> None:
        project_path = Path(path)
        if project_path.suffix.lower() != PROJECT_FILE_EXTENSION:
            project_path = project_path.with_suffix(PROJECT_FILE_EXTENSION)

        project_path.parent.mkdir(parents=True, exist_ok=True)

        content = json.dumps(
            data,
            ensure_ascii=False,
            indent=2,
        )

        project_path.write_text(content, encoding="utf-8")

    def load_project(self, path: str | Path) -> dict[str, Any]:
        project_path = Path(path)

        if not project_path.exists():
            raise FileNotFoundError(f"Project file not found: {project_path}")

        content = project_path.read_text(encoding="utf-8")
        data = json.loads(content)

        if not isinstance(data, dict):
            raise RuntimeError("Invalid project file format.")

        return data