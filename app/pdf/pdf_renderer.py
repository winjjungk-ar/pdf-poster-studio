from pathlib import Path

import fitz
from PyQt6.QtGui import QImage, QPixmap


class PdfRenderer:
    def __init__(self) -> None:
        self.document: fitz.Document | None = None
        self.pdf_path: Path | None = None

    def open(self, pdf_path: str) -> int:
        self.close()
        self.pdf_path = Path(pdf_path)
        self.document = fitz.open(str(self.pdf_path))

        if self.page_count <= 0:
            raise RuntimeError("PDF has no pages.")

        return self.page_count

    @property
    def page_count(self) -> int:
        if self.document is None:
            return 0
        return len(self.document)

    def render_page(self, page_index: int, zoom: float = 1.5) -> QPixmap:
        if self.document is None:
            raise RuntimeError("PDF document is not opened.")

        if page_index < 0 or page_index >= self.page_count:
            raise IndexError("Page index is out of range.")

        page = self.document.load_page(page_index)
        matrix = fitz.Matrix(zoom, zoom)

        pix = page.get_pixmap(
            matrix=matrix,
            alpha=False,
            colorspace=fitz.csRGB,
        )

        image = QImage(
            pix.samples,
            pix.width,
            pix.height,
            pix.stride,
            QImage.Format.Format_RGB888,
        )

        if image.isNull():
            raise RuntimeError("Rendered QImage is null.")

        pixmap = QPixmap.fromImage(image.copy())

        if pixmap.isNull():
            raise RuntimeError("Rendered QPixmap is null.")

        return pixmap

    def get_page_pixmap(self, page_index: int, zoom: float = 1.5) -> QPixmap:
        return self.render_page(page_index, zoom=zoom)
    def close(self) -> None:
        if self.document is not None:
            self.document.close()

        self.document = None
        self.pdf_path = None