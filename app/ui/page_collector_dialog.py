from pathlib import Path

import fitz
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QFileDialog,
    QTableWidget,
    QTableWidgetItem,
    QMessageBox,
    QLabel,
    QLineEdit,
)

from app.config import EXPORTS_DIR
from app.pdf.pdf_page_collector import PdfCollectItem, PdfPageCollector


class PageCollectorDialog(QDialog):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)

        self.setWindowTitle("Collect Pages from PDFs")
        self.resize(900, 520)

        self.collector = PdfPageCollector()
        self._create_ui()

    def _create_ui(self) -> None:
        layout = QVBoxLayout(self)

        info = QLabel(
            "Add PDF files and specify pages to collect.\n"
            "Page range examples: all, 1,3,5, 2-4, 1,3,6-9"
        )

        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["PDF File", "Pages", "Page Count"])
        self.table.setColumnWidth(0, 520)
        self.table.setColumnWidth(1, 160)
        self.table.setColumnWidth(2, 100)

        button_layout = QHBoxLayout()

        add_button = QPushButton("Add PDF Files")
        add_button.clicked.connect(self.add_pdf_files)

        move_up_button = QPushButton("Move Up")
        move_up_button.clicked.connect(self.move_selected_row_up)

        move_down_button = QPushButton("Move Down")
        move_down_button.clicked.connect(self.move_selected_row_down)

        validate_button = QPushButton("Validate Page Ranges")
        validate_button.clicked.connect(self.validate_page_ranges)

        remove_button = QPushButton("Remove Selected")
        remove_button.clicked.connect(self.remove_selected)

        clear_button = QPushButton("Clear List")
        clear_button.clicked.connect(self.clear_list)

        button_layout.addWidget(add_button)
        button_layout.addWidget(remove_button)
        button_layout.addWidget(move_up_button)
        button_layout.addWidget(move_down_button)
        button_layout.addWidget(validate_button)
        button_layout.addWidget(clear_button)

        output_layout = QHBoxLayout()

        self.output_path_edit = QLineEdit()
        self.output_path_edit.setPlaceholderText("Output PDF path")

        browse_output_button = QPushButton("Browse Output")
        browse_output_button.clicked.connect(self.browse_output)

        output_layout.addWidget(QLabel("Output"))
        output_layout.addWidget(self.output_path_edit)
        output_layout.addWidget(browse_output_button)

        action_layout = QHBoxLayout()

        create_button = QPushButton("Create Collected PDF")
        create_button.clicked.connect(self.create_collected_pdf)
        create_button.setStyleSheet("font-weight: bold;")

        close_button = QPushButton("Close")
        close_button.clicked.connect(self.close)

        action_layout.addStretch()
        action_layout.addWidget(create_button)
        action_layout.addWidget(close_button)

        layout.addWidget(info)
        layout.addLayout(button_layout)
        layout.addWidget(self.table)
        layout.addLayout(output_layout)
        layout.addLayout(action_layout)

    def add_pdf_files(self) -> None:
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Add PDF Files",
            "",
            "PDF Files (*.pdf)"
        )

        if not file_paths:
            return

        for file_path in file_paths:
            self._add_pdf_row(file_path)

        if not self.output_path_edit.text().strip() and file_paths:
            first_name = Path(file_paths[0]).stem
            default_path = EXPORTS_DIR / f"{first_name}_collected.pdf"
            self.output_path_edit.setText(str(default_path))

    def _add_pdf_row(self, file_path: str) -> None:
        try:
            doc = fitz.open(file_path)
            page_count = len(doc)
            doc.close()
        except Exception as exc:
            QMessageBox.warning(
                self,
                "PDF Error",
                f"Failed to read PDF:\n{file_path}\n\n{exc}"
            )
            return

        row = self.table.rowCount()
        self.table.insertRow(row)

        file_item = QTableWidgetItem(file_path)
        pages_item = QTableWidgetItem("all")
        count_item = QTableWidgetItem(str(page_count))
        count_item.setFlags(count_item.flags() & ~Qt.ItemFlag.ItemIsEditable)

        self.table.setItem(row, 0, file_item)
        self.table.setItem(row, 1, pages_item)
        self.table.setItem(row, 2, count_item)

    def remove_selected(self) -> None:
        selected_rows = sorted(
            {index.row() for index in self.table.selectedIndexes()},
            reverse=True,
        )

        for row in selected_rows:
            self.table.removeRow(row)

    def clear_list(self) -> None:
        self.table.setRowCount(0)

    def browse_output(self) -> None:
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Output PDF",
            str(EXPORTS_DIR / "collected.pdf"),
            "PDF Files (*.pdf)"
        )

        if file_path:
            self.output_path_edit.setText(file_path)


    def _swap_table_rows(self, row_a: int, row_b: int) -> None:
        if row_a < 0 or row_b < 0:
            return

        if row_a >= self.table.rowCount() or row_b >= self.table.rowCount():
            return

        column_count = self.table.columnCount()
        values_a = []
        values_b = []

        for col in range(column_count):
            item_a = self.table.item(row_a, col)
            item_b = self.table.item(row_b, col)
            values_a.append("" if item_a is None else item_a.text())
            values_b.append("" if item_b is None else item_b.text())

        for col in range(column_count):
            self.table.setItem(row_a, col, QTableWidgetItem(values_b[col]))
            self.table.setItem(row_b, col, QTableWidgetItem(values_a[col]))

        self.table.selectRow(row_b)

    def move_selected_row_up(self) -> None:
        row = self.table.currentRow()

        if row <= 0:
            return

        self._swap_table_rows(row, row - 1)

    def move_selected_row_down(self) -> None:
        row = self.table.currentRow()

        if row < 0 or row >= self.table.rowCount() - 1:
            return

        self._swap_table_rows(row, row + 1)

    def validate_page_ranges(self) -> None:
        if self.table.rowCount() <= 0:
            QMessageBox.information(self, "Validate Page Ranges", "No PDF files to validate.")
            return

        errors = []

        for row in range(self.table.rowCount()):
            pdf_item = self.table.item(row, 0)
            pages_item = self.table.item(row, 1)
            count_item = self.table.item(row, 2)

            if pdf_item is None:
                errors.append(f"Row {row + 1}: PDF path is empty.")
                continue

            pdf_path = pdf_item.text().strip()
            page_ranges = "all"

            if pages_item is not None:
                page_ranges = pages_item.text().strip() or "all"

            try:
                if count_item is not None and count_item.text().strip().isdigit():
                    page_count = int(count_item.text().strip())
                else:
                    doc = fitz.open(pdf_path)
                    page_count = doc.page_count
                    doc.close()

                self.collector.parse_page_ranges(page_ranges, page_count)

            except Exception as exc:
                errors.append(f"Row {row + 1}: {exc}")

        if errors:
            QMessageBox.warning(
                self,
                "Validate Page Ranges",
                "Some page ranges are invalid:\n\n" + "\n".join(errors)
            )
            return

        QMessageBox.information(
            self,
            "Validate Page Ranges",
            "All page ranges are valid."
        )


    def create_collected_pdf(self) -> None:
        output_path = self.output_path_edit.text().strip()

        if not output_path:
            QMessageBox.warning(self, "Output Required", "Please select output PDF path.")
            return

        if self.table.rowCount() <= 0:
            QMessageBox.warning(self, "No PDFs", "Please add PDF files first.")
            return

        items = []

        for row in range(self.table.rowCount()):
            pdf_item = self.table.item(row, 0)
            pages_item = self.table.item(row, 1)

            if pdf_item is None:
                continue

            pdf_path = pdf_item.text().strip()
            page_ranges = "all"

            if pages_item is not None:
                page_ranges = pages_item.text().strip() or "all"

            items.append(
                PdfCollectItem(
                    pdf_path=pdf_path,
                    page_ranges=page_ranges,
                )
            )

        try:
            self.collector.collect(
                items=items,
                output_pdf_path=output_path,
            )

            result = QMessageBox.question(
                self,
                "Complete",
                f"Collected PDF was created:\n\n{output_path}\n\nOpen this PDF in the main window?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )

            if result == QMessageBox.StandardButton.Yes and self.parent() is not None:
                parent = self.parent()

                if hasattr(parent, "_open_pdf_path"):
                    parent._open_pdf_path(output_path, clear_selections=True)
                    parent.is_modified = True

                self.close()

        except Exception as exc:
            QMessageBox.critical(
                self,
                "Collect Error",
                f"Failed to create collected PDF.\n\n{exc}"
            )