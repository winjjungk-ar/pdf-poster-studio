import json
from pathlib import Path


class RecentProjects:
    def __init__(self, storage_path: Path, limit: int = 5) -> None:
        self.storage_path = storage_path
        self.limit = limit

    def load(self) -> list[str]:
        if not self.storage_path.exists():
            return []

        try:
            data = json.loads(self.storage_path.read_text(encoding="utf-8"))
        except Exception:
            return []

        if not isinstance(data, list):
            return []

        result = []
        for item in data:
            if isinstance(item, str) and item not in result:
                result.append(item)

        return result[: self.limit]

    def save(self, paths: list[str]) -> None:
        self.storage_path.parent.mkdir(parents=True, exist_ok=True)
        clean_paths = []

        for path in paths:
            if path and path not in clean_paths:
                clean_paths.append(path)

        self.storage_path.write_text(
            json.dumps(clean_paths[: self.limit], ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def add(self, path: str) -> list[str]:
        paths = self.load()

        if path in paths:
            paths.remove(path)

        paths.insert(0, path)
        paths = paths[: self.limit]
        self.save(paths)

        return paths