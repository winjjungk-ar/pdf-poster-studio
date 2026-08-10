from pathlib import Path


def normalize_path(path: str) -> Path:
    return Path(path).expanduser().resolve()


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)