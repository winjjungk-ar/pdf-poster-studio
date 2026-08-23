
from pathlib import Path
from PyQt6.QtGui import QPixmap, QPainter, QColor, QPen, QBrush, QFont, QIcon
from PyQt6.QtWidgets import QApplication
import sys

app = QApplication(sys.argv)

out = Path("app_icon.ico")

pixmap = QPixmap(256, 256)
pixmap.fill(QColor("#00000000"))

p = QPainter(pixmap)
p.setRenderHint(QPainter.RenderHint.Antialiasing)

# Background
p.setBrush(QBrush(QColor("#DDF3F5")))
p.setPen(QPen(QColor("#8CCFD6"), 8))
p.drawEllipse(12, 12, 232, 232)

# Paper
p.setBrush(QBrush(QColor("#FFFFFF")))
p.setPen(QPen(QColor("#5F7F95"), 7))
p.drawRoundedRect(70, 38, 116, 164, 16, 16)

# Poster split lines
p.setPen(QPen(QColor("#8EA7FF"), 7))
p.drawLine(128, 60, 128, 180)
p.drawLine(88, 120, 168, 120)

# Selection rectangle
p.setBrush(QBrush(QColor(255, 188, 120, 110)))
p.setPen(QPen(QColor("#FF9F68"), 7))
p.drawRoundedRect(88, 82, 80, 70, 12, 12)

# PDF text
p.setFont(QFont("Segoe UI", 28, QFont.Weight.Bold))
p.setPen(QColor("#395B64"))
p.drawText(78, 218, "PDF")

p.end()

icon = QIcon(pixmap)
icon.addPixmap(pixmap)
icon_engine_sizes = [16, 24, 32, 48, 64, 128, 256]
icon.save(str(out))

print(f"[OK] Icon created: {out.resolve()}")
