import json
import shutil
from pathlib import Path


SETTINGS_PATH = Path(__file__).parent / "ocr_settings.json"


DEFAULT_WINDOWS_PATHS = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
]


def find_tesseract_executable() -> str:
    from_path = shutil.which("tesseract")
    if from_path:
        return from_path

    for path in DEFAULT_WINDOWS_PATHS:
        if Path(path).exists():
            return path

    return ""


def load_ocr_settings() -> dict:
    data = {
        "tesseract_cmd": "",
        "tessdata_dir": str((Path(__file__).parent / "tessdata").resolve()),
        "language": "eng+kor",
        "dpi": 200,
    }

    try:
        if SETTINGS_PATH.exists():
            saved = json.loads(SETTINGS_PATH.read_text(encoding="utf-8-sig"))
            data.update(saved)
    except Exception:
        pass

    if not data.get("tesseract_cmd"):
        data["tesseract_cmd"] = find_tesseract_executable()

    return data


def save_ocr_settings(settings: dict) -> None:
    SETTINGS_PATH.write_text(
        json.dumps(settings, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
