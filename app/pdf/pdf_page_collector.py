from dataclasses import dataclass
from pathlib import Path

import fitz


@dataclass
class PdfCollectItem:
    pdf_path: str
    page_ranges: str


class PdfPageCollector:
    def collect(
        self,
        items: list[PdfCollectItem],
        output_pdf_path: str,
    ) -> None:
        if not items:
            raise RuntimeError("No PDF files were added.")

        output_path = Path(output_pdf_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        out_doc = fitz.open()

        try:
            for item in items:
                src_path = Path(item.pdf_path)

                if not src_path.exists():
                    raise FileNotFoundError(f"PDF file not found: {src_path}")

                src_doc = fitz.open(str(src_path))

                try:
                    page_indices = self.parse_page_ranges(
                        page_ranges=item.page_ranges,
                        page_count=len(src_doc),
                    )

                    for page_index in page_indices:
                        out_doc.insert_pdf(
                            src_doc,
                            from_page=page_index,
                            to_page=page_index,
                        )

                finally:
                    src_doc.close()

            if len(out_doc) == 0:
                raise RuntimeError("No pages were collected.")

            out_doc.save(str(output_path), garbage=4, deflate=True)

        finally:
            out_doc.close()

    def parse_page_ranges(
        self,
        page_ranges: str,
        page_count: int,
    ) -> list[int]:
        text = (page_ranges or "").strip().lower()

        if text in ("", "all", "*"):
            return list(range(page_count))

        result: list[int] = []
        parts = [part.strip() for part in text.split(",") if part.strip()]

        for part in parts:
            if "-" in part:
                start_text, end_text = part.split("-", 1)
                start = int(start_text.strip())
                end = int(end_text.strip())

                if start <= 0 or end <= 0:
                    raise ValueError("Page numbers must be greater than zero.")

                if start > end:
                    raise ValueError(f"Invalid page range: {part}")

                for page_no in range(start, end + 1):
                    self._append_page_number(result, page_no, page_count)
            else:
                page_no = int(part)

                if page_no <= 0:
                    raise ValueError("Page numbers must be greater than zero.")

                self._append_page_number(result, page_no, page_count)

        if not result:
            raise ValueError("No valid page numbers were specified.")

        return result

    def _append_page_number(
        self,
        result: list[int],
        page_no: int,
        page_count: int,
    ) -> None:
        if page_no > page_count:
            raise ValueError(
                f"Page number {page_no} is out of range. Total pages: {page_count}"
            )

        page_index = page_no - 1

        if page_index not in result:
            result.append(page_index)