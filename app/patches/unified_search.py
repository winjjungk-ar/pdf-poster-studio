from __future__ import annotations

from types import MethodType

from PyQt6.QtWidgets import QMessageBox, QToolBar, QPushButton, QToolButton
from PyQt6.QtGui import QAction

from app.localization import t


def _is_pdf_text_search_label(text: str) -> bool:
    text = (text or "").strip().lower()
    compact = text.replace(" ", "").replace("&", "")
    return (
        "pdf텍스트찾기" in compact
        or "pdftextsearch" in compact
        or "pdftext" in compact
        or "텍스트찾기" in compact
    )


def _is_ocr_search_label(text: str) -> bool:
    text = (text or "").strip().lower()
    compact = text.replace(" ", "").replace("&", "")
    return (
        "ocr글자찾기" in compact
        or "ocr텍스트찾기" in compact
        or "quickocr" in compact
        or "ocrsearch" in compact
        or "ocr" == compact
    )


def _safe_set_search_status(window, text: str) -> None:
    try:
        if hasattr(window, "set_search_status"):
            window.set_search_status(text)
        elif hasattr(window, "search_status_label"):
            window.search_status_label.setText(text)
    except Exception:
        pass


def _unified_search_from_toolbar(self) -> None:
    """
    Unified search:
    1. Try embedded PDF text search.
    2. If no result, ask user whether to continue with OCR.
    """
    if not hasattr(self, "search_line_edit"):
        try:
            self.find_text_dialog()
        except Exception:
            pass
        return

    query = self.search_line_edit.text().strip()

    if not query:
        self.search_line_edit.setFocus()
        return

    if not getattr(self, "current_pdf_path", None):
        QMessageBox.information(self, t("find_text"), t("please_open_pdf_first"))
        return

    self.last_find_query = query
    self.search_mode = "pdf_text"
    self.search_results = []
    self.search_result_index = -1
    self.last_ocr_search_page = -1

    # 1) PDF text-layer search first.
    try:
        import app.main_window as main_window_module

        search_pdf_text = getattr(main_window_module, "search_pdf_text", None)

        if search_pdf_text is not None:
            _safe_set_search_status(self, "PDF 텍스트에서 찾는 중...")
            self.search_results = search_pdf_text(self.current_pdf_path, query) or []
            self.search_result_index = -1

            if hasattr(self, "update_search_results_panel"):
                self.update_search_results_panel()

            if self.search_results:
                self.search_mode = "pdf_text"
                _safe_set_search_status(self, f"PDF 텍스트 검색 결과: {len(self.search_results)}건")
                self.find_next_text()
                return

    except Exception as exc:
        try:
            self._log(f"Unified search PDF text phase failed: {exc}")
        except Exception:
            pass

    # 2) Ask OCR fallback.
    answer = QMessageBox.question(
        self,
        "OCR 검색 확인",
        (
            f"PDF 텍스트에서 '{query}' 검색 결과가 없습니다.\n\n"
            "스캔 문서이거나 이미지 안의 글자일 수 있습니다.\n"
            "OCR로 검색하시겠습니까?"
        ),
        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        QMessageBox.StandardButton.Yes,
    )

    if answer != QMessageBox.StandardButton.Yes:
        _safe_set_search_status(self, t("search_status_no_result"))
        return

    # 3) OCR quick search via existing find_next_text logic.
    self.search_mode = "quick_ocr"
    self.search_results = []
    self.search_result_index = -1
    self.last_ocr_search_page = -1

    try:
        _safe_set_search_status(self, "OCR로 글자 찾는 중...")
        self.find_next_text()
    except Exception as exc:
        _safe_set_search_status(self, t("search_status_no_result"))
        QMessageBox.information(
            self,
            "Find Text",
            f"An error occurred during OCR search.\n\n{exc}",
        )
        try:
            self._log(f"Unified OCR fallback failed: {exc}")
        except Exception:
            pass


def _unified_pdf_text_search_from_toolbar(self) -> None:
    return _unified_search_from_toolbar(self)


def _unified_quick_ocr_search_from_toolbar(self) -> None:
    """
    Keep compatibility: old OCR action now runs unified search too.
    """
    return _unified_search_from_toolbar(self)


def _cleanup_search_ui(window) -> None:
    """
    Best-effort UI cleanup:
    - Rename PDF text search action/button to 'Find Text'
    - Hide OCR-only action/button from toolbars
    """
    # QAction cleanup
    try:
        actions = window.findChildren(QAction)
    except Exception:
        actions = []

    seen_search_action = False

    for action in actions:
        try:
            text = action.text() or ""
            tooltip = action.toolTip() or ""
            status_tip = action.statusTip() or ""

            combined = " ".join([text, tooltip, status_tip])

            if _is_ocr_search_label(combined):
                # Hide OCR-only action. If it is in menu, this also hides it.
                action.setVisible(False)
                continue

            if _is_pdf_text_search_label(combined):
                if not seen_search_action:
                    action.setText("Find Text")
                    action.setToolTip("Search PDF text first. If no result is found, OCR search will be suggested.")
                    action.setStatusTip("Find Text")
                    seen_search_action = True
                else:
                    action.setVisible(False)
        except Exception:
            pass

    # QPushButton / QToolButton cleanup
    button_classes = [QPushButton, QToolButton]

    for cls in button_classes:
        try:
            buttons = window.findChildren(cls)
        except Exception:
            buttons = []

        seen_button = False

        for button in buttons:
            try:
                text = button.text() or ""
                tooltip = button.toolTip() or ""
                combined = " ".join([text, tooltip])

                if _is_ocr_search_label(combined):
                    button.hide()
                    continue

                if _is_pdf_text_search_label(combined):
                    if not seen_button:
                        button.setText("Find Text")
                        button.setToolTip("Search PDF text first. If no result is found, OCR search will be suggested.")
                        seen_button = True
                    else:
                        button.hide()
            except Exception:
                pass

    # Toolbar duplicate cleanup after visibility changes.
    try:
        for toolbar in window.findChildren(QToolBar):
            for action in list(toolbar.actions()):
                text = action.text() or ""
                tooltip = action.toolTip() or ""
                combined = " ".join([text, tooltip])

                if _is_ocr_search_label(combined):
                    toolbar.removeAction(action)
    except Exception:
        pass


def install_unified_search(window) -> None:
    """
    Monkey-patch MainWindow search entry points without editing large main_window.py.
    """
    try:
        window.search_from_toolbar = MethodType(_unified_search_from_toolbar, window)
        window.pdf_text_search_from_toolbar = MethodType(_unified_pdf_text_search_from_toolbar, window)
        window.quick_ocr_search_from_toolbar = MethodType(_unified_quick_ocr_search_from_toolbar, window)
        window._unified_search_installed = True

        _cleanup_search_ui(window)

        try:
            window._log("Unified search installed: PDF text first, OCR fallback.")
        except Exception:
            pass

    except Exception as exc:
        try:
            window._log(f"Unified search install failed: {exc}")
        except Exception:
            pass
