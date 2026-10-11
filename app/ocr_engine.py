import hashlib
import json
import os
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import List

import fitz
import pytesseract
from PIL import Image

from app.ocr_settings import load_ocr_settings


CACHE_DIR = Path(__file__).parent / "cache" / "ocr"


@dataclass
class OcrTextBlock:
    page_index: int
    text: str
    confidence: float
    x: int
    y: int
    width: int
    height: int
    dpi: int


def configure_tesseract() -> dict:
    settings = load_ocr_settings()

    tesseract_cmd = settings.get("tesseract_cmd", "")
    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

    tessdata_dir = settings.get("tessdata_dir", "")
    if tessdata_dir:
        tessdata_dir = str(Path(tessdata_dir).resolve())
        settings["tessdata_dir"] = tessdata_dir
        os.environ["TESSDATA_PREFIX"] = tessdata_dir

    return settings


def _file_fingerprint(pdf_path: str) -> str:
    path = Path(pdf_path)
    stat = path.stat()
    raw = f"{path.resolve()}|{stat.st_size}|{int(stat.st_mtime)}"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:24]


def _cache_path(pdf_path: str, page_index: int, language: str, dpi: int) -> Path:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)

    fingerprint = _file_fingerprint(pdf_path)
    safe_lang = language.replace("+", "_").replace("/", "_").replace("\\", "_")

    return CACHE_DIR / f"{fingerprint}_p{page_index + 1}_{safe_lang}_{dpi}.json"


def _load_cached_blocks(pdf_path: str, page_index: int, language: str, dpi: int) -> List[OcrTextBlock] | None:
    path = _cache_path(pdf_path, page_index, language, dpi)

    if not path.exists():
        return None

    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
        return [OcrTextBlock(**item) for item in data.get("blocks", [])]
    except Exception:
        return None


def _save_cached_blocks(pdf_path: str, page_index: int, language: str, dpi: int, blocks: List[OcrTextBlock]) -> None:
    path = _cache_path(pdf_path, page_index, language, dpi)

    try:
        payload = {
            "pdf_path": str(Path(pdf_path).resolve()),
            "page_index": page_index,
            "language": language,
            "dpi": dpi,
            "blocks": [asdict(block) for block in blocks],
        }
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def clear_ocr_cache() -> None:
    if not CACHE_DIR.exists():
        return

    for path in CACHE_DIR.glob("*.json"):
        try:
            path.unlink()
        except Exception:
            pass


def render_page_to_image(page: fitz.Page, dpi: int = 200) -> Image.Image:
    scale = dpi / 72
    matrix = fitz.Matrix(scale, scale)
    pixmap = page.get_pixmap(matrix=matrix, alpha=False)

    return Image.frombytes(
        "RGB",
        (pixmap.width, pixmap.height),
        pixmap.samples,
    )


def run_ocr_on_page(
    page: fitz.Page,
    page_index: int,
    language: str = "eng+jpn+kor",
    dpi: int = 200,
    pdf_path: str | None = None,
) -> List[OcrTextBlock]:
    settings = configure_tesseract()
    language = settings.get("language", language)
    dpi = int(settings.get("dpi", dpi))

    if pdf_path:
        cached = _load_cached_blocks(pdf_path, page_index, language, dpi)
        if cached is not None:
            return cached

    image = render_page_to_image(page, dpi=dpi)

    tessdata_dir = settings.get("tessdata_dir", "")
    config = ""
    if tessdata_dir and Path(tessdata_dir).exists():
        config = f"--tessdata-dir {tessdata_dir}"

    data = pytesseract.image_to_data(
        image,
        lang=language,
        config=config,
        output_type=pytesseract.Output.DICT,
    )

    results: List[OcrTextBlock] = []

    for i, raw_text in enumerate(data.get("text", [])):
        text = raw_text.strip()
        if not text:
            continue

        try:
            confidence = float(data["conf"][i])
        except Exception:
            confidence = -1.0

        if confidence < 35:
            continue

        results.append(
            OcrTextBlock(
                page_index=page_index,
                text=text,
                confidence=confidence,
                x=int(data["left"][i]),
                y=int(data["top"][i]),
                width=int(data["width"][i]),
                height=int(data["height"][i]),
                dpi=dpi,
            )
        )

    if pdf_path:
        _save_cached_blocks(pdf_path, page_index, language, dpi, results)

    return results


def image_rect_to_pdf_rect(x: int, y: int, width: int, height: int, dpi: int):
    scale = 72 / dpi
    return fitz.Rect(
        x * scale,
        y * scale,
        (x + width) * scale,
        (y + height) * scale,
    )
