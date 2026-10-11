import json
import shutil
import sys
from pathlib import Path


SETTINGS_PATH = Path(__file__).parent / "ocr_settings.json"


DEFAULT_WINDOWS_PATHS = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
]


def app_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def bundled_tesseract_path() -> str:
    base = app_base_dir()

    candidates = [
        base / "tesseract" / "tesseract.exe",
        base / "Tesseract-OCR" / "tesseract.exe",
        base / "_internal" / "Tesseract-OCR" / "tesseract.exe",
        base / "app" / "tesseract" / "tesseract.exe",
        base / "_internal" / "app" / "tesseract" / "tesseract.exe",
    ]

    for path in candidates:
        if path.exists():
            return str(path)

    return ""


def bundled_tessdata_dir() -> str:
    base = app_base_dir()

    candidates = [
        base / "tesseract" / "tessdata",
        base / "Tesseract-OCR" / "tessdata",
        base / "_internal" / "Tesseract-OCR" / "tessdata",
        base / "app" / "tessdata",
        base / "_internal" / "app" / "tessdata",
        Path(__file__).resolve().parent / "tessdata",
    ]

    for path in candidates:
        if path.exists():
            return str(path.resolve())

    return str((Path(__file__).resolve().parent / "tessdata").resolve())


def find_tesseract_executable() -> str:
    bundled = bundled_tesseract_path()
    if bundled:
        return bundled

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
        "tessdata_dir": bundled_tessdata_dir(),
        "language": "eng+jpn+kor",
        "dpi": 300,
    }

    try:
        if SETTINGS_PATH.exists():
            saved = json.loads(SETTINGS_PATH.read_text(encoding="utf-8-sig"))
            data.update(saved)
    except Exception:
        pass

    # 배포본 내부 tesseract를 우선 사용
    bundled = bundled_tesseract_path()
    if bundled:
        data["tesseract_cmd"] = bundled
        data["tessdata_dir"] = bundled_tessdata_dir()
    elif not data.get("tesseract_cmd"):
        data["tesseract_cmd"] = find_tesseract_executable()

    if not data.get("tessdata_dir"):
        data["tessdata_dir"] = bundled_tessdata_dir()

    return data


def save_ocr_settings(settings: dict) -> None:
    SETTINGS_PATH.write_text(
        json.dumps(settings, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
