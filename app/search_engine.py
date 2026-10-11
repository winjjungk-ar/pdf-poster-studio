import time
from dataclasses import dataclass
from typing import List, Optional

import fitz

from app.ocr_engine import image_rect_to_pdf_rect, run_ocr_on_page


@dataclass
class SearchResult:
    page_index: int
    text: str
    source: str
    rect: Optional[fitz.Rect] = None
    confidence: Optional[float] = None


def search_pdf_text(pdf_path: str, query: str) -> List[SearchResult]:
    doc = fitz.open(pdf_path)
    results: List[SearchResult] = []

    try:
        for page_index, page in enumerate(doc):
            rects = page.search_for(query)
            for rect in rects:
                results.append(
                    SearchResult(
                        page_index=page_index,
                        text=query,
                        source="text",
                        rect=rect,
                    )
                )
    finally:
        doc.close()

    return results


def search_pdf_ocr(
    pdf_path: str,
    query: str,
    language: str = "eng+jpn+kor",
    dpi: int = 200,
    max_pages: Optional[int] = None,
) -> List[SearchResult]:
    doc = fitz.open(pdf_path)
    results: List[SearchResult] = []
    query_lower = query.lower()

    try:
        page_count = len(doc)
        if max_pages is not None:
            page_count = min(page_count, max_pages)

        for page_index in range(page_count):
            page = doc[page_index]
            blocks = run_ocr_on_page(
                page,
                page_index=page_index,
                language=language,
                dpi=dpi,
                pdf_path=pdf_path,
            )

            for block in blocks:
                if query_lower in block.text.lower():
                    rect = image_rect_to_pdf_rect(
                        block.x,
                        block.y,
                        block.width,
                        block.height,
                        block.dpi,
                    )
                    results.append(
                        SearchResult(
                            page_index=page_index,
                            text=block.text,
                            source="ocr",
                            rect=rect,
                            confidence=block.confidence,
                        )
                    )
    finally:
        doc.close()

    return results


def search_pdf_with_optional_ocr(
    pdf_path: str,
    query: str,
    use_ocr_if_needed: bool = True,
    language: str = "eng+jpn+kor",
    dpi: int = 200,
) -> List[SearchResult]:
    results = search_pdf_text(pdf_path, query)

    if results or not use_ocr_if_needed:
        return results

    return search_pdf_ocr(
        pdf_path=pdf_path,
        query=query,
        language=language,
        dpi=dpi,
    )

def search_pdf_ocr_from_page(
    pdf_path: str,
    query: str,
    start_page_index: int = 0,
    language: str = "eng+jpn+kor",
    dpi: int = 200,
    max_pages: Optional[int] = None,
) -> List[SearchResult]:
    doc = fitz.open(pdf_path)
    query_lower = query.lower()

    try:
        total_pages = len(doc)
        if total_pages <= 0:
            return []

        if max_pages is None:
            max_pages = total_pages

        start_page_index = max(0, min(start_page_index, total_pages - 1))

        # 현재 페이지부터 끝까지, 그 다음 처음부터 현재 페이지 전까지 순환 검색
        page_order = list(range(start_page_index, total_pages)) + list(range(0, start_page_index))

        checked = 0
        for page_index in page_order:
            if checked >= max_pages:
                break

            page = doc[page_index]
            blocks = run_ocr_on_page(
                page,
                page_index=page_index,
                language=language,
                dpi=dpi,
                pdf_path=pdf_path,
            )

            page_results: List[SearchResult] = []

            for block in blocks:
                if query_lower in block.text.lower():
                    rect = image_rect_to_pdf_rect(
                        block.x,
                        block.y,
                        block.width,
                        block.height,
                        block.dpi,
                    )
                    page_results.append(
                        SearchResult(
                            page_index=page_index,
                            text=block.text,
                            source="ocr",
                            rect=rect,
                            confidence=block.confidence,
                        )
                    )

            if page_results:
                return page_results

            checked += 1

    finally:
        doc.close()

    return []

def search_pdf_ocr_from_page_limited(
    pdf_path: str,
    query: str,
    start_page_index: int = 0,
    language: str = "eng+jpn+kor",
    dpi: int = 200,
    max_pages: int = 3,
    max_seconds: float = 8.0,
) -> List[SearchResult]:
    started_at = time.monotonic()

    doc = fitz.open(pdf_path)
    query_lower = query.lower()

    try:
        total_pages = len(doc)
        if total_pages <= 0:
            return []

        start_page_index = max(0, min(start_page_index, total_pages - 1))
        page_order = list(range(start_page_index, total_pages)) + list(range(0, start_page_index))

        checked = 0

        for page_index in page_order:
            if checked >= max_pages:
                break

            if time.monotonic() - started_at >= max_seconds:
                break

            page = doc[page_index]
            blocks = run_ocr_on_page(
                page,
                page_index=page_index,
                language=language,
                dpi=dpi,
                pdf_path=pdf_path,
            )

            page_results: List[SearchResult] = []

            for block in blocks:
                if query_lower in block.text.lower():
                    rect = image_rect_to_pdf_rect(
                        block.x,
                        block.y,
                        block.width,
                        block.height,
                        block.dpi,
                    )
                    page_results.append(
                        SearchResult(
                            page_index=page_index,
                            text=block.text,
                            source="ocr",
                            rect=rect,
                            confidence=block.confidence,
                        )
                    )

            if page_results:
                return page_results

            checked += 1

    finally:
        doc.close()

    return []

def search_pdf_ocr_from_page_limited_with_progress(
    pdf_path: str,
    query: str,
    start_page_index: int = 0,
    language: str = "eng+jpn+kor",
    dpi: int = 200,
    max_pages: int = 3,
    max_seconds: float = 8.0,
    progress_callback=None,
    cancel_callback=None,
) -> List[SearchResult]:
    started_at = time.monotonic()

    doc = fitz.open(pdf_path)
    query_lower = query.lower()

    try:
        total_pages = len(doc)
        if total_pages <= 0:
            return []

        start_page_index = max(0, min(start_page_index, total_pages - 1))
        page_order = list(range(start_page_index, total_pages)) + list(range(0, start_page_index))

        checked = 0

        for page_index in page_order:
            if checked >= max_pages:
                break

            if time.monotonic() - started_at >= max_seconds:
                break

            if cancel_callback and cancel_callback():
                break

            if progress_callback:
                progress_callback(page_index, checked, max_pages, total_pages)

            page = doc[page_index]
            blocks = run_ocr_on_page(
                page,
                page_index=page_index,
                language=language,
                dpi=dpi,
                pdf_path=pdf_path,
            )

            if cancel_callback and cancel_callback():
                break

            page_results: List[SearchResult] = []

            for block in blocks:
                if query_lower in block.text.lower():
                    rect = image_rect_to_pdf_rect(
                        block.x,
                        block.y,
                        block.width,
                        block.height,
                        block.dpi,
                    )
                    page_results.append(
                        SearchResult(
                            page_index=page_index,
                            text=block.text,
                            source="ocr",
                            rect=rect,
                            confidence=block.confidence,
                        )
                    )

            if page_results:
                return page_results

            checked += 1

    finally:
        doc.close()

    return []

def search_pdf_ocr_all_with_progress(
    pdf_path: str,
    query: str,
    language: str = "eng+jpn+kor",
    dpi: int = 200,
    progress_callback=None,
    cancel_callback=None,
) -> List[SearchResult]:
    doc = fitz.open(pdf_path)
    results: List[SearchResult] = []
    query_lower = query.lower()

    try:
        total_pages = len(doc)

        for page_index in range(total_pages):
            if cancel_callback and cancel_callback():
                break

            if progress_callback:
                progress_callback(page_index, total_pages)

            page = doc[page_index]
            blocks = run_ocr_on_page(
                page,
                page_index=page_index,
                language=language,
                dpi=dpi,
                pdf_path=pdf_path,
            )

            if cancel_callback and cancel_callback():
                break

            for block in blocks:
                if query_lower in block.text.lower():
                    rect = image_rect_to_pdf_rect(
                        block.x,
                        block.y,
                        block.width,
                        block.height,
                        block.dpi,
                    )

                    results.append(
                        SearchResult(
                            page_index=page_index,
                            text=block.text,
                            source="ocr",
                            rect=rect,
                            confidence=block.confidence,
                        )
                    )

    finally:
        doc.close()

    return results
