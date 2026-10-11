from __future__ import annotations

from pathlib import Path
import traceback

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QFileDialog,
    QMessageBox,
    QInputDialog,
    QProgressDialog,
    QApplication,
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QFormLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QComboBox,
    QCheckBox,
    QDialogButtonBox,
    QGroupBox,
    QStackedWidget,
    QToolBar,
)
from PyQt6.QtGui import QAction

from app.localization import t
try:
    from app.converters.ppt_quality_checker import create_ppt_quality_report
except Exception:
    create_ppt_quality_report = None

from app.converters.office_exporter import (
    OfficeExporter,
    OfficeExportOptions,
    parse_page_range,
)


def _get_current_pdf_path(window) -> str | None:
    candidates = [
        "current_pdf_path",
        "pdf_path",
        "source_pdf_path",
        "current_file_path",
        "file_path",
    ]

    for name in candidates:
        value = getattr(window, name, None)
        if value:
            path = str(value)
            if path.lower().endswith(".pdf") and Path(path).exists():
                return path

    try:
        workspace_items = getattr(window, "workspace_items", None)
        current_workspace_index = getattr(window, "current_workspace_index", None)

        if (
            isinstance(workspace_items, list)
            and current_workspace_index is not None
            and 0 <= int(current_workspace_index) < len(workspace_items)
        ):
            item = workspace_items[int(current_workspace_index)]
            if isinstance(item, dict):
                for key in ("pdf_path", "path", "file_path"):
                    value = item.get(key)
                    if value and str(value).lower().endswith(".pdf") and Path(str(value)).exists():
                        return str(value)
    except Exception:
        pass

    return None


def _choose_source_pdf(window) -> str | None:
    current = _get_current_pdf_path(window)
    if current:
        return current

    path, _ = QFileDialog.getOpenFileName(
        window,
        "PDF Select",
        "",
        "PDF Files (*.pdf)",
    )
    return path or None


def _get_page_count(source_pdf_path: str) -> int:
    try:
        import fitz

        doc = fitz.open(source_pdf_path)
        count = len(doc)
        doc.close()
        return count
    except Exception:
        return 0


class PowerPointConvertDialog(QDialog):
    def __init__(self, parent, source_pdf_path: str) -> None:
        super().__init__(parent)

        self.setWindowTitle("PowerPoint Conversion Settings")
        self.setMinimumWidth(520)

        self.source_pdf_path = source_pdf_path
        self.page_count = _get_page_count(source_pdf_path)

        root = QVBoxLayout(self)

        source_group = QGroupBox("Target PDF")
        source_layout = QVBoxLayout(source_group)
        source_label = QLabel(str(source_pdf_path))
        source_label.setWordWrap(True)
        source_layout.addWidget(source_label)
        root.addWidget(source_group)

        form_group = QGroupBox("Conversion Settings")
        form = QFormLayout(form_group)

        self.page_range_edit = QLineEdit()
        self.page_range_edit.setPlaceholderText("Example: 1-3,5,8-10 / blank for all")
        form.addRow(
            f"Page Range{f' (total {self.page_count} pages)' if self.page_count else ''}:",
            self.page_range_edit,
        )

        self.mode_combo = QComboBox()
        self.mode_combo.addItems(
            [
                "Auto Mode - Best Selection per Page",
                "Image Slides Preserving Original Layout",
                "Editable Text / Images",
            ]
        )
        form.addRow("PowerPoint Conversion Mode:", self.mode_combo)

        self.dpi_combo = QComboBox()
        self.dpi_combo.addItems(
            [
                "150 - Fast",
                "200 - Recommended",
                "300 - High Quality",
            ]
        )
        self.dpi_combo.setCurrentIndex(1)
        form.addRow("Quality:", self.dpi_combo)

        root.addWidget(form_group)

        option_group = QGroupBox("Options")
        option_layout = QVBoxLayout(option_group)

        self.open_after_check = QCheckBox("Open PowerPoint file after conversion")
        self.open_after_check.setChecked(True)

        self.quality_report_check = QCheckBox("Generate Quality Comparison Report")
        self.quality_report_check.setChecked(False)

        option_layout.addWidget(self.open_after_check)
        option_layout.addWidget(self.quality_report_check)

        if create_ppt_quality_report is None:
            self.quality_report_check.setEnabled(False)
            self.quality_report_check.setText("Generate Quality Comparison Report (Unavailable)")

        root.addWidget(option_group)

        notice = QLabel(
            "Guide:\n"
            "- Auto Mode converts normal pages as editable content and complex drawing pages as original images.\n"
            "- Editable Text / Images avoids text overlap, but some complex shapes may be missing.\n"
            "- Image Slides Preserving Original Layout preserves the original appearance best, but text editing is not available."
        )
        notice.setWordWrap(True)
        notice.setStyleSheet("color: #555;")
        root.addWidget(notice)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("Convert")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Cancel")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

    def get_options(self) -> tuple[OfficeExportOptions, bool, bool] | None:
        try:
            page_range = (
                parse_page_range(self.page_range_edit.text(), self.page_count)
                if self.page_count
                else None
            )
        except Exception as exc:
            QMessageBox.warning(
                self,
                "Page Range Error",
                f"Invalid page range format.\n\n{exc}",
            )
            return None

        dpi_text = self.dpi_combo.currentText()
        dpi_value = int(dpi_text.split()[0])

        mode_text = self.mode_combo.currentText()

        if "Auto Mode" in mode_text:
            powerpoint_mode = "auto"
            include_editable_text_layer = True
        elif "Editable" in mode_text:
            powerpoint_mode = "text_only"
            include_editable_text_layer = True
        else:
            powerpoint_mode = "image"
            include_editable_text_layer = False

        options = OfficeExportOptions(
            dpi=dpi_value,
            page_range=page_range,
            fit_mode="fit",
            include_page_breaks=True,
            include_editable_text_layer=include_editable_text_layer,
            powerpoint_mode=powerpoint_mode,
            max_drawing_objects=1200 if powerpoint_mode in ("text_only", "auto") else 0,
            faded_background_transparency=78.0,
            text_box_fill_transparency=18.0,
        )

        return (
            options,
            self.open_after_check.isChecked(),
            self.quality_report_check.isChecked(),
        )


def _ask_options(
    window,
    source_pdf_path: str,
    for_powerpoint: bool = False,
) -> OfficeExportOptions | None:
    """
    Legacy prompt for Word/Excel and fallback.
    PowerPoint uses PowerPointConvertDialog.
    """
    page_count = _get_page_count(source_pdf_path)
    dpi_value = 200

    if for_powerpoint:
        dialog = PowerPointConvertDialog(window, source_pdf_path)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return None

        result = dialog.get_options()
        if result is None:
            return None

        options, _open_after, _quality_report = result
        options._open_after_export = _open_after
        options._create_quality_report = _quality_report
        return options

    range_text, ok = QInputDialog.getText(
        window,
        "Page Range",
        "Enter page range to convert\n"
        "Example: 1-3,5,8-10\n"
        "blank for all pages"
        + (f"\nTotal pages: {page_count}" if page_count else ""),
    )

    if not ok:
        return None

    try:
        page_range = parse_page_range(range_text, page_count) if page_count else None
    except Exception as exc:
        QMessageBox.warning(
            window,
            "Page Range Error",
            f"Invalid page range format.\n\n{exc}",
        )
        return None

    return OfficeExportOptions(
        dpi=dpi_value,
        page_range=page_range,
        fit_mode="fit",
        include_page_breaks=True,
        include_editable_text_layer=False,
        powerpoint_mode="image",
    )


def _build_auto_mode_summary(options) -> str:
    mode = getattr(options, "powerpoint_mode", "image")

    if mode == "auto":
        report = getattr(options, "auto_mode_report", []) or []
        image_count = sum(1 for item in report if item.get("mode") == "image")
        editable_count = sum(1 for item in report if item.get("mode") == "text_only")

        detail_lines = []
        for item in report[:20]:
            mode_label = "Original Image" if item.get("mode") == "image" else "Editable"
            reason = item.get("reason", "")
            page = item.get("page", "more")
            detail_lines.append(f"- page {page}: {mode_label} ({reason})")

        if len(report) > 20:
            detail_lines.append(f"- ... more {len(report) - 20} pages")

        note = (
            "Converted using Auto Mode.\n"
            f"- Image slides: {image_count} pages\n"
            f"- Editable Text / Images: {editable_count} pages"
        )

        if detail_lines:
            note += "\n\nPer-page processing:\n" + "\n".join(detail_lines)

        return note

    if mode == "text_only":
        return "Converted using Editable Text / Images mode."

    return "Converted using Image Slides Preserving Original Layout mode."


def export_current_pdf_to_powerpoint(window, source_pdf_path: str | None = None) -> None:
    source_pdf = source_pdf_path or _choose_source_pdf(window)

    if not source_pdf:
        return

    options = _ask_options(window, source_pdf, for_powerpoint=True)

    if options is None:
        return

    default_name = str(Path(source_pdf).with_suffix(".pptx"))

    output_path, _ = QFileDialog.getSaveFileName(
        window,
        "Save as PowerPoint",
        default_name,
        "PowerPoint Files (*.pptx)",
    )

    if not output_path:
        return

    if not output_path.lower().endswith(".pptx"):
        output_path += ".pptx"

    progress = None

    try:
        progress = QProgressDialog(
            "Preparing PowerPoint conversion...",
            "Cancel",
            0,
            100,
            window,
        )
        progress.setWindowTitle("Converting to PowerPoint")
        progress.setMinimumDuration(0)
        progress.setAutoClose(False)
        progress.setAutoReset(False)
        progress.setValue(3)
        progress.setLabelText("Starting PowerPoint conversion...")
        progress.show()
        QApplication.processEvents()

        def on_progress(done, total, message):
            if total <= 0:
                value = 0
            else:
                value = int((done / total) * 100)

            progress.setValue(max(0, min(100, value)))
            progress.setLabelText(message)
            QApplication.processEvents()

            return not progress.wasCanceled()

        OfficeExporter().export_pdf_to_powerpoint(
            source_pdf,
            output_path,
            options,
            progress_callback=on_progress,
        )

        progress.setValue(100)
        progress.close()

        note = _build_auto_mode_summary(options)

        QMessageBox.information(
            window,
            "Conversion Complete",
            f"PowerPoint conversion completed.\n\n{output_path}\n\n{note}",
        )

        if create_ppt_quality_report is not None and getattr(options, "_create_quality_report", False):
            try:
                report_dir = str(Path(output_path).with_suffix("")) + "_quality_report"
                report = create_ppt_quality_report(
                    source_pdf_path=source_pdf,
                    pptx_path=output_path,
                    output_dir=report_dir,
                    page_indexes=options.page_range,
                    dpi=120,
                )

                QMessageBox.information(
                    window,
                    "Quality Report Generated",
                    f"Average Score: {report.average_score:.2f}\n\n{report.report_html}",
                )

                if hasattr(window, "_open_file_with_default_app"):
                    window._open_file_with_default_app(report.report_html)

            except Exception as report_exc:
                QMessageBox.warning(
                    window,
                    "Quality Report Failed",
                    f"PowerPoint conversion completed, but quality report generation failed.\n\n{report_exc}",
                )

        if getattr(options, "_open_after_export", True):
            try:
                if hasattr(window, "_open_file_with_default_app"):
                    window._open_file_with_default_app(output_path)
            except Exception:
                pass

    except Exception as exc:
        try:
            if progress is not None:
                progress.close()
        except Exception:
            pass

        QMessageBox.critical(
            window,
            "PowerPoint Conversion Failed",
            f"{exc}\n\n{traceback.format_exc()}",
        )


def export_current_pdf_to_word(window, source_pdf_path: str | None = None) -> None:
    source_pdf = source_pdf_path or _choose_source_pdf(window)

    if not source_pdf:
        return

    options = _ask_options(window, source_pdf, for_powerpoint=False)

    if options is None:
        return

    word_mode, ok = QInputDialog.getItem(
        window,
        "Word Conversion Mode",
        "Conversion Mode:",
        [
            "Auto Mode - Best Selection per Page",
            "Text Only",
            "Page Image + Text",
            "Image Document Preserving Original Layout",
            "Editable Layout Document - Requires Word",
        ],
        0,
        False,
    )

    if not ok:
        return

    if "Auto Mode" in word_mode:
        options.word_mode = "auto_layout"
        options.max_drawing_objects = 1200
    elif "Editable Layout" in word_mode:
        options.word_mode = "layout_editable_com"
        options.max_drawing_objects = 1200
    elif "Original Layout" in word_mode:
        options.word_mode = "layout_image"
    elif "Image" in word_mode:
        options.word_mode = "image_text"
    else:
        options.word_mode = "text"

    default_name = str(Path(source_pdf).with_suffix(".docx"))

    output_path, _ = QFileDialog.getSaveFileName(
        window,
        "Save as Word",
        default_name,
        "Word Files (*.docx)",
    )

    if not output_path:
        return

    if not output_path.lower().endswith(".docx"):
        output_path += ".docx"

    try:
        OfficeExporter().export_pdf_to_word(source_pdf, output_path, options)

        note = ""
        if getattr(options, "word_mode", "") == "auto_layout":
            report = getattr(options, "word_auto_report", []) or []
            image_count = sum(1 for item in report if item.get("mode") == "image")
            editable_count = sum(1 for item in report if item.get("mode") == "editable")

            lines = []
            for item in report[:20]:
                mode_label = "Original Image" if item.get("mode") == "image" else "Editable"
                lines.append(f"- page {item.get('page')}: {mode_label} ({item.get('reason', '')})")

            if len(report) > 20:
                lines.append(f"- ... more {len(report) - 20} pages")

            note = (
                "\n\nAuto Mode processing result:\n"
                f"- Original image pages: {image_count} pages\n"
                f"- Editable pages: {editable_count} pages"
            )

            if lines:
                note += "\n\nPer-page processing:\n" + "\n".join(lines)

        QMessageBox.information(
            window,
            "Conversion Complete",
            f"Word conversion completed.\n\n{output_path}{note}",
        )

        try:
            if hasattr(window, "_open_file_with_default_app"):
                window._open_file_with_default_app(output_path)
        except Exception:
            pass

    except Exception as exc:
        QMessageBox.critical(window, "Word Conversion Failed", f"{exc}\n\n{traceback.format_exc()}")


def export_current_pdf_to_excel(window, source_pdf_path: str | None = None) -> None:
    source_pdf = source_pdf_path or _choose_source_pdf(window)

    if not source_pdf:
        return

    options = _ask_options(window, source_pdf, for_powerpoint=False)

    if options is None:
        return

    excel_mode, ok = QInputDialog.getItem(
        window,
        "Excel Conversion Mode",
        "Conversion Mode:",
        [
            "Extract Text by Lines",
            "Estimate Table Layout by Coordinates",
        ],
        1,
        False,
    )

    if not ok:
        return

    options.excel_mode = "table" if "Table" in excel_mode else "lines"

    default_name = str(Path(source_pdf).with_suffix(".xlsx"))

    output_path, _ = QFileDialog.getSaveFileName(
        window,
        "Save as Excel",
        default_name,
        "Excel Files (*.xlsx)",
    )

    if not output_path:
        return

    if not output_path.lower().endswith(".xlsx"):
        output_path += ".xlsx"

    try:
        OfficeExporter().export_pdf_to_excel(source_pdf, output_path, options)
        QMessageBox.information(window, "Conversion Complete", f"Excel conversion completed.\n\n{output_path}")

        try:
            if hasattr(window, "_open_file_with_default_app"):
                window._open_file_with_default_app(output_path)
        except Exception:
            pass

    except Exception as exc:
        QMessageBox.critical(window, "Excel Conversion Failed", f"{exc}\n\n{traceback.format_exc()}")


def show_office_convert_dialog(window) -> None:
    """
    Main popup-style launcher for Office conversion.
    """
    source_pdf = _choose_source_pdf(window)
    if not source_pdf:
        return

    dialog = QDialog(window)
    dialog.setWindowTitle(t("office_convert"))
    dialog.setMinimumWidth(420)

    layout = QVBoxLayout(dialog)

    title = QLabel("Convert PDF to Office File")
    title.setStyleSheet("font-size: 18px; font-weight: bold;")
    layout.addWidget(title)

    pdf_label = QLabel(f"Target PDF:\n{source_pdf}")
    pdf_label.setWordWrap(True)
    layout.addWidget(pdf_label)

    desc = QLabel(
        "Select a conversion format.\n\n"
        "- PowerPoint: slide conversion / Auto Mode supported\n"
        "- Word: Extract PDF text to a Word document\n"
        "- Excel: Extract PDF text by rows"
    )
    desc.setWordWrap(True)
    layout.addWidget(desc)

    button_row = QHBoxLayout()

    ppt_button = QPushButton("PowerPoint")
    word_button = QPushButton("Word")
    excel_button = QPushButton("Excel")

    button_row.addWidget(ppt_button)
    button_row.addWidget(word_button)
    button_row.addWidget(excel_button)

    layout.addLayout(button_row)

    close_button = QPushButton("Close")
    layout.addWidget(close_button)

    def run_ppt():
        dialog.accept()
        export_current_pdf_to_powerpoint(window, source_pdf)

    def run_word():
        dialog.accept()
        export_current_pdf_to_word(window, source_pdf)

    def run_excel():
        dialog.accept()
        export_current_pdf_to_excel(window, source_pdf)

    ppt_button.clicked.connect(run_ppt)
    word_button.clicked.connect(run_word)
    excel_button.clicked.connect(run_excel)
    close_button.clicked.connect(dialog.reject)

    dialog.exec()




# --- toolbar cleanup helpers v1 ---
def _is_fit_width_text(text: str) -> bool:
    text = (text or "").strip().lower()
    compact = text.replace(" ", "").replace("&", "")
    return (
        "너비맞춤" in compact
        or "fitwidth" in compact
        or "fit-width" in compact
        or "폭맞춤" in compact
    )


def cleanup_duplicate_toolbar_actions(window) -> None:
    """
    Cleanup toolbar-only duplicates without touching menu actions.

    - Keep first Fit Width action visible in toolbar.
    - Hide duplicated Fit Width actions from toolbars.
    - Remove Office Convert toolbar created by earlier patch.
    """
    try:
        toolbars = window.findChildren(QToolBar)
    except Exception:
        return

    # 1) Remove Office Convert toolbar if it was created by earlier patch.
    for toolbar in list(toolbars):
        try:
            object_name = toolbar.objectName() or ""
            title = toolbar.windowTitle() or ""

            if object_name == "office_convert_toolbar" or title in ("Office ??", "Office Convert", t("office_convert")):
                toolbar.hide()
                toolbar.setParent(None)
        except Exception:
            pass

    # Refresh toolbar list after removal.
    try:
        toolbars = window.findChildren(QToolBar)
    except Exception:
        toolbars = []

    # 2) Remove duplicate Fit Width actions from toolbars.
    seen_fit_width = False

    for toolbar in toolbars:
        try:
            for action in list(toolbar.actions()):
                text = action.text() or ""
                tooltip = action.toolTip() or ""
                status_tip = action.statusTip() or ""

                if (
                    _is_fit_width_text(text)
                    or _is_fit_width_text(tooltip)
                    or _is_fit_width_text(status_tip)
                ):
                    if not seen_fit_width:
                        seen_fit_width = True
                    else:
                        toolbar.removeAction(action)
        except Exception:
            continue



def install_office_actions(window) -> None:
    menu_bar = window.menuBar()
    office_menu = menu_bar.addMenu(t("office_convert"))

    main_action = QAction(t("office_convert") + "...", window)
    main_action.triggered.connect(lambda: show_office_convert_dialog(window))
    office_menu.addAction(main_action)

    office_menu.addSeparator()

    ppt_action = QAction("PDF → PowerPoint...", window)
    ppt_action.triggered.connect(lambda: export_current_pdf_to_powerpoint(window))
    office_menu.addAction(ppt_action)

    word_action = QAction("PDF → Word...", window)
    word_action.triggered.connect(lambda: export_current_pdf_to_word(window))
    office_menu.addAction(word_action)

    excel_action = QAction("PDF → Excel...", window)
    excel_action.triggered.connect(lambda: export_current_pdf_to_excel(window))
    office_menu.addAction(excel_action)

    office_menu.addSeparator()

    info_action = QAction(t("conversion_mode_guide"), window)

    def show_info():
        QMessageBox.information(
            window,
            t("office_convert_guide"),
            t("office_convert_guide_text"),
        )

    info_action.triggered.connect(show_info)
    office_menu.addAction(info_action)

    cleanup_duplicate_toolbar_actions(window)

    window._office_export_actions = {
        "menu": office_menu,
        "main": main_action,
        "ppt": ppt_action,
        "word": word_action,
        "excel": excel_action,
        "info": info_action,
    }



# --- unified office convert dialog v2 ---
class UnifiedOfficeConvertDialog(QDialog):
    def __init__(self, parent, source_pdf_path: str) -> None:
        super().__init__(parent)

        self.setWindowTitle(t("office_convert"))
        self.setMinimumWidth(560)

        self.source_pdf_path = source_pdf_path
        self.page_count = _get_page_count(source_pdf_path)

        root = QVBoxLayout(self)

        title = QLabel("Convert PDF to Office File")
        title.setStyleSheet("font-size: 18px; font-weight: bold;")
        root.addWidget(title)

        pdf_label = QLabel(f"Target PDF:\n{source_pdf_path}")
        pdf_label.setWordWrap(True)
        pdf_label.setStyleSheet("color: #444;")
        root.addWidget(pdf_label)

        common_group = QGroupBox("Common Settings")
        common_form = QFormLayout(common_group)

        self.format_combo = QComboBox()
        self.format_combo.addItems(["PowerPoint", "Word", "Excel"])
        common_form.addRow("Conversion Format:", self.format_combo)

        self.page_range_edit = QLineEdit()
        self.page_range_edit.setPlaceholderText("Example: 1-3,5,8-10 / blank for all")
        common_form.addRow(
            f"Page Range{f' (total {self.page_count} pages)' if self.page_count else ''}:",
            self.page_range_edit,
        )

        self.open_after_check = QCheckBox("Open file after conversion")
        self.open_after_check.setChecked(True)
        common_form.addRow("", self.open_after_check)

        root.addWidget(common_group)

        self.stack = QStackedWidget()
        root.addWidget(self.stack)

        # PowerPoint settings
        ppt_group = QGroupBox("PowerPoint Settings")
        ppt_form = QFormLayout(ppt_group)

        self.ppt_mode_combo = QComboBox()
        self.ppt_mode_combo.addItems(
            [
                "Auto Mode - Best Selection per Page",
                "Image Slides Preserving Original Layout",
                "Editable Text / Images",
            ]
        )
        ppt_form.addRow("Conversion Mode:", self.ppt_mode_combo)

        self.ppt_dpi_combo = QComboBox()
        self.ppt_dpi_combo.addItems(["150 - Fast", "200 - Recommended", "300 - High Quality"])
        self.ppt_dpi_combo.setCurrentIndex(1)
        ppt_form.addRow("Quality:", self.ppt_dpi_combo)

        self.ppt_quality_report_check = QCheckBox("Generate Quality Comparison Report")
        self.ppt_quality_report_check.setChecked(False)

        if create_ppt_quality_report is None:
            self.ppt_quality_report_check.setEnabled(False)
            self.ppt_quality_report_check.setText("Generate Quality Comparison Report (Unavailable)")

        ppt_form.addRow("", self.ppt_quality_report_check)

        ppt_notice = QLabel(
            "Auto Mode converts normal pages as editable content and complex drawing pages as original images."
        )
        ppt_notice.setWordWrap(True)
        ppt_notice.setStyleSheet("color: #666;")
        ppt_form.addRow("", ppt_notice)

        self.stack.addWidget(ppt_group)

        # Word settings
        word_group = QGroupBox("Word Settings")
        word_form = QFormLayout(word_group)

        self.word_mode_combo = QComboBox()
        self.word_mode_combo.addItems(
            [
                "Auto Mode - Best Selection per Page",
                "Text Only",
                "Page Image + Text",
                "Image Document Preserving Original Layout",
                "Editable Layout Document - Requires Word",
            ]
        )
        word_form.addRow("Conversion Mode:", self.word_mode_combo)

        word_notice = QLabel(
            "If original layout is important, use 'Image Document Preserving Original Layout'.\n"
            "If text editing is needed, use 'Editable Layout Document'.",
        )
        word_notice.setWordWrap(True)
        word_notice.setStyleSheet("color: #666;")
        word_form.addRow("", word_notice)

        self.stack.addWidget(word_group)

        # Excel settings
        excel_group = QGroupBox("Excel Settings")
        excel_form = QFormLayout(excel_group)

        self.excel_mode_combo = QComboBox()
        self.excel_mode_combo.addItems(
            [
                "Estimate Table Layout by Coordinates",
                "Extract Text by Lines",
            ]
        )
        excel_form.addRow("Conversion Mode:", self.excel_mode_combo)

        excel_notice = QLabel(
            "For tables/lists, use 'Estimate Table Layout by Coordinates'.\n"
            "Complex design documents may not convert perfectly.",
        )
        excel_notice.setWordWrap(True)
        excel_notice.setStyleSheet("color: #666;")
        excel_form.addRow("", excel_notice)

        self.stack.addWidget(excel_group)

        self.format_combo.currentIndexChanged.connect(self.stack.setCurrentIndex)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("Convert")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Cancel")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

    def _page_range(self):
        if not self.page_count:
            return None

        return parse_page_range(self.page_range_edit.text(), self.page_count)

    def selected_format(self) -> str:
        return self.format_combo.currentText()

    def open_after_export(self) -> bool:
        return self.open_after_check.isChecked()

    def build_options(self) -> OfficeExportOptions | None:
        try:
            page_range = self._page_range()
        except Exception as exc:
            QMessageBox.warning(
                self,
                "Page Range Error",
                f"Invalid page range format.\n\n{exc}",
            )
            return None

        fmt = self.selected_format()

        if fmt == "PowerPoint":
            dpi = int(self.ppt_dpi_combo.currentText().split()[0])
            mode_text = self.ppt_mode_combo.currentText()

            if "Auto Mode" in mode_text:
                powerpoint_mode = "auto"
                include_editable = True
            elif "Editable" in mode_text:
                powerpoint_mode = "text_only"
                include_editable = True
            else:
                powerpoint_mode = "image"
                include_editable = False

            options = OfficeExportOptions(
                dpi=dpi,
                page_range=page_range,
                fit_mode="fit",
                include_page_breaks=True,
                include_editable_text_layer=include_editable,
                powerpoint_mode=powerpoint_mode,
                max_drawing_objects=1200 if powerpoint_mode in ("text_only", "auto") else 0,
            )
            options._create_quality_report = self.ppt_quality_report_check.isChecked()
            options._open_after_export = self.open_after_export()
            return options

        if fmt == "Word":
            mode_text = self.word_mode_combo.currentText()

            if "Auto Mode" in mode_text:
                word_mode = "auto_layout"
                max_drawings = 1200
            elif "Editable Layout" in mode_text:
                word_mode = "layout_editable_com"
                max_drawings = 1200
            elif "Original Layout" in mode_text:
                word_mode = "layout_image"
                max_drawings = 0
            elif "Image" in mode_text:
                word_mode = "image_text"
                max_drawings = 0
            else:
                word_mode = "text"
                max_drawings = 0

            options = OfficeExportOptions(
                dpi=200,
                page_range=page_range,
                fit_mode="fit",
                include_page_breaks=True,
                powerpoint_mode="image",
                max_drawing_objects=max_drawings,
            )
            options.word_mode = word_mode
            options._open_after_export = self.open_after_export()
            return options

        # Excel
        mode_text = self.excel_mode_combo.currentText()
        excel_mode = "table" if "Table" in mode_text else "lines"

        options = OfficeExportOptions(
            dpi=200,
            page_range=page_range,
            fit_mode="fit",
            include_page_breaks=True,
            powerpoint_mode="image",
        )
        options.excel_mode = excel_mode
        options._open_after_export = self.open_after_export()
        return options


def _run_office_conversion_with_options(
    window,
    source_pdf: str,
    fmt: str,
    options: OfficeExportOptions,
) -> None:
    suffix_map = {
        "PowerPoint": ".pptx",
        "Word": ".docx",
        "Excel": ".xlsx",
    }

    filter_map = {
        "PowerPoint": "PowerPoint Files (*.pptx)",
        "Word": "Word Files (*.docx)",
        "Excel": "Excel Files (*.xlsx)",
    }

    suffix = suffix_map.get(fmt, ".pptx")
    default_name = str(Path(source_pdf).with_suffix(suffix))

    output_path, _ = QFileDialog.getSaveFileName(
        window,
        f"Save as {fmt}",
        default_name,
        filter_map.get(fmt, "All Files (*.*)"),
    )

    if not output_path:
        return

    if not output_path.lower().endswith(suffix):
        output_path += suffix

    progress = None

    try:
        if fmt == "PowerPoint":
            progress = QProgressDialog("Preparing PowerPoint conversion...", "Cancel", 0, 100, window)
            progress.setWindowTitle("Converting to PowerPoint")
            progress.setMinimumDuration(0)
            progress.setAutoClose(False)
            progress.setAutoReset(False)
            progress.setValue(3)
            progress.show()
            QApplication.processEvents()

            def on_progress(done, total, message):
                value = int((done / total) * 100) if total > 0 else 0
                progress.setValue(max(0, min(100, value)))
                progress.setLabelText(message)
                QApplication.processEvents()
                return not progress.wasCanceled()

            OfficeExporter().export_pdf_to_powerpoint(
                source_pdf,
                output_path,
                options,
                progress_callback=on_progress,
            )

            progress.setValue(100)
            progress.close()

            note = _build_auto_mode_summary(options)

            QMessageBox.information(
                window,
                "Conversion Complete",
                f"PowerPoint conversion completed.\n\n{output_path}\n\n{note}",
            )

            if create_ppt_quality_report is not None and getattr(options, "_create_quality_report", False):
                try:
                    report_dir = str(Path(output_path).with_suffix("")) + "_quality_report"
                    report = create_ppt_quality_report(
                        source_pdf_path=source_pdf,
                        pptx_path=output_path,
                        output_dir=report_dir,
                        page_indexes=options.page_range,
                        dpi=120,
                    )
                    QMessageBox.information(
                        window,
                        "Quality Report Generated",
                        f"Average Score: {report.average_score:.2f}\n\n{report.report_html}",
                    )
                    if hasattr(window, "_open_file_with_default_app"):
                        window._open_file_with_default_app(report.report_html)
                except Exception as report_exc:
                    QMessageBox.warning(
                        window,
                        "Quality Report Failed",
                        f"PowerPoint conversion completed, but quality report generation failed.\n\n{report_exc}",
                    )

        elif fmt == "Word":
            OfficeExporter().export_pdf_to_word(source_pdf, output_path, options)

            note = ""
            if getattr(options, "word_mode", "") == "auto_layout":
                report = getattr(options, "word_auto_report", []) or []
                image_count = sum(1 for item in report if item.get("mode") == "image")
                editable_count = sum(1 for item in report if item.get("mode") == "editable")

                lines = []
                for item in report[:20]:
                    mode_label = "Original Image" if item.get("mode") == "image" else "Editable"
                    lines.append(f"- page {item.get('page')}: {mode_label} ({item.get('reason', '')})")

                note = (
                    "\n\nAuto Mode processing result:\n"
                    f"- Original image pages: {image_count} pages\n"
                    f"- Editable pages: {editable_count} pages"
                )

                if lines:
                    note += "\n\nPer-page processing:\n" + "\n".join(lines)

            QMessageBox.information(
                window,
                "Conversion Complete",
                f"Word conversion completed.\n\n{output_path}{note}",
            )

        else:
            OfficeExporter().export_pdf_to_excel(source_pdf, output_path, options)
            QMessageBox.information(
                window,
                "Conversion Complete",
                f"Excel conversion completed.\n\n{output_path}",
            )

        if getattr(options, "_open_after_export", True):
            try:
                if hasattr(window, "_open_file_with_default_app"):
                    window._open_file_with_default_app(output_path)
            except Exception:
                pass

    except Exception as exc:
        try:
            if progress is not None:
                progress.close()
        except Exception:
            pass

        QMessageBox.critical(
            window,
            f"{fmt} Conversion Failed",
            f"{exc}\n\n{traceback.format_exc()}",
        )


def show_office_convert_dialog(window) -> None:
    source_pdf = _choose_source_pdf(window)
    if not source_pdf:
        return

    dialog = UnifiedOfficeConvertDialog(window, source_pdf)

    if dialog.exec() != QDialog.DialogCode.Accepted:
        return

    options = dialog.build_options()
    if options is None:
        return

    _run_office_conversion_with_options(
        window,
        source_pdf,
        dialog.selected_format(),
        options,
    )


