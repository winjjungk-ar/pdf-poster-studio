from PyQt6.QtCore import Qt, QRect, QPoint, pyqtSignal
from PyQt6.QtGui import QPainter, QPen, QColor, QPixmap
from PyQt6.QtWidgets import QLabel


class PdfPreviewWidget(QLabel):
    selection_changed = pyqtSignal(QRect)
    zoom_requested = pyqtSignal(int)

    def __init__(self) -> None:
        super().__init__()

        self._pixmap: QPixmap | None = None
        self._selection_rect = QRect()
        self._drag_start = QPoint()
        self._dragging = False

        self.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        self.setMouseTracking(True)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        # Keep the preview area clean before a PDF is loaded.
        # Selection rectangles are drawn only when a preview pixmap exists.
        self.setStyleSheet(
            "QLabel { border: none; background: transparent; font-size: 18px; }"
        )

    def set_preview_pixmap(self, pixmap: QPixmap) -> None:
        self._pixmap = pixmap
        self.setPixmap(pixmap)

        # Important:
        # Keep widget geometry exactly equal to the displayed pixmap.
        # This prevents wrong selection coordinates after heavy zoom-out.
        self.setMinimumSize(pixmap.size())
        self.setMaximumSize(pixmap.size())
        self.resize(pixmap.size())

        self._selection_rect = self._clamp_rect_to_pixmap(self._selection_rect)
        self.update()

    def clear_preview(self) -> None:
        self.preview_pixmap = None

        if hasattr(self, "selection_rect"):
            self.selection_rect = QRect()

        if hasattr(self, "_selection_rect"):
            self._selection_rect = QRect()

        self.setMinimumSize(1, 1)
        self.resize(1, 1)
        self.update()


    def set_selection_rect(self, rect: QRect) -> None:
        self._selection_rect = self._clamp_rect_to_pixmap(rect.normalized())
        self.update()

    def get_selection_rect(self) -> QRect:
        return self._selection_rect.normalized()

    def clear_selection(self) -> None:
        self._selection_rect = QRect()
        self.selection_changed.emit(self._selection_rect)
        self.update()

    def auto_select_content(self, threshold: int = 245, padding: int = 6) -> QRect:
        if self._pixmap is None:
            return QRect()

        image = self._pixmap.toImage()
        width = image.width()
        height = image.height()

        if width <= 0 or height <= 0:
            return QRect()

        min_x = width
        min_y = height
        max_x = -1
        max_y = -1

        # Sampling step keeps it fast on large previews.
        step = max(1, min(width, height) // 900)

        for y in range(0, height, step):
            for x in range(0, width, step):
                color = image.pixelColor(x, y)

                # Detect non-white / non-background pixels.
                if (
                    color.red() < threshold
                    or color.green() < threshold
                    or color.blue() < threshold
                ):
                    min_x = min(min_x, x)
                    min_y = min(min_y, y)
                    max_x = max(max_x, x)
                    max_y = max(max_y, y)

        if max_x < min_x or max_y < min_y:
            return QRect()

        min_x = max(0, min_x - padding)
        min_y = max(0, min_y - padding)
        max_x = min(width - 1, max_x + padding)
        max_y = min(height - 1, max_y + padding)

        rect = QRect(
            min_x,
            min_y,
            max_x - min_x + 1,
            max_y - min_y + 1,
        ).normalized()

        self._selection_rect = self._clamp_rect_to_pixmap(rect)
        self.selection_changed.emit(self._selection_rect)
        self.update()

        return self._selection_rect

    def mousePressEvent(self, event) -> None:
        if event.button() != Qt.MouseButton.LeftButton:
            return

        if self._pixmap is None:
            return

        pos = self._clamp_point_to_pixmap(event.position().toPoint())

        self._dragging = True
        self._drag_start = pos
        self._selection_rect = QRect(self._drag_start, self._drag_start)
        self.update()

    def mouseMoveEvent(self, event) -> None:
        if not self._dragging:
            return

        current_pos = self._clamp_point_to_pixmap(event.position().toPoint())
        self._selection_rect = QRect(self._drag_start, current_pos).normalized()
        self.update()

    def mouseReleaseEvent(self, event) -> None:
        if event.button() != Qt.MouseButton.LeftButton:
            return

        if not self._dragging:
            return

        self._dragging = False
        current_pos = self._clamp_point_to_pixmap(event.position().toPoint())
        self._selection_rect = QRect(self._drag_start, current_pos).normalized()
        self._selection_rect = self._clamp_rect_to_pixmap(self._selection_rect)

        self.selection_changed.emit(self._selection_rect)
        self.update()

    def wheelEvent(self, event) -> None:
        if self._pixmap is None:
            return

        delta = event.angleDelta().y()

        if delta > 0:
            self.zoom_requested.emit(1)
        elif delta < 0:
            self.zoom_requested.emit(-1)

        event.accept()

    def paintEvent(self, event) -> None:
        # Do not draw placeholder rectangle before a PDF is loaded.
        pixmap = getattr(self, 'preview_pixmap', None)
        if pixmap is None:
            pixmap = getattr(self, '_pixmap', None)
        if pixmap is None:
            pixmap = getattr(self, 'pixmap', None)
        if pixmap is None or (hasattr(pixmap, 'isNull') and pixmap.isNull()):
            return

        super().paintEvent(event)

        if self._selection_rect.isNull() or self._selection_rect.width() <= 0 or self._selection_rect.height() <= 0:
            return

        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        fill_color = QColor(255, 0, 0, 35)
        border_pen = QPen(QColor(220, 0, 0), 2, Qt.PenStyle.SolidLine)

        painter.fillRect(self._selection_rect, fill_color)
        painter.setPen(border_pen)
        painter.drawRect(self._selection_rect)

        painter.end()

    def _clamp_point_to_pixmap(self, point: QPoint) -> QPoint:
        if self._pixmap is None:
            return point

        x = max(0, min(point.x(), self._pixmap.width() - 1))
        y = max(0, min(point.y(), self._pixmap.height() - 1))
        return QPoint(x, y)

    def _clamp_rect_to_pixmap(self, rect: QRect) -> QRect:
        if self._pixmap is None:
            return rect.normalized()

        bounds = QRect(0, 0, self._pixmap.width(), self._pixmap.height())
        return rect.normalized().intersected(bounds)