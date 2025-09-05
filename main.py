#!/usr/bin/env python3

import sys
import math
from PyQt5.QtGui import QBrush, QPen, QPolygonF, QColor, QPainter, QFont
from PyQt5.QtCore import Qt, QPointF, QRectF
from PyQt5.QtWidgets import QApplication, QGraphicsView, QGraphicsScene, QGraphicsPolygonItem, QMainWindow, QGraphicsTextItem


def hex_corners(center_x, center_y, size):
    """Return list of 6 QPointF for a pointy-top hex centered at (center_x, center_y)."""
    # pointy-top orientation: first corner at angle = 30 deg
    corners = []
    for i in range(6):
        angle_deg = 30 + 60 * i
        angle_rad = math.radians(angle_deg)
        x = center_x + size * math.cos(angle_rad)
        y = center_y + size * math.sin(angle_rad)
        corners.append(QPointF(x, y))
    return corners


def column_to_excel_name(col_index):
    """Convert column index (0-based) to Excel-style column name (A, B, C, ..., Z, AA, AB, ...)"""
    result = ""
    while col_index >= 0:
        result = chr(65 + (col_index % 26)) + result
        col_index = col_index // 26 - 1
        if col_index < 0:
            break
    return result


class HexItem(QGraphicsPolygonItem):
    def __init__(self, polygon, row, col):
        super().__init__(polygon)
        self.row = row
        self.col = col
        self.setPen(QPen(QColor("#cbd5e1"), 0.8))
        self.setBrush(QBrush(QColor("#ffffff")))

    def hoverEnterEvent(self, event):
        self.setBrush(QBrush(QColor("#f1f5f9")))
        super().hoverEnterEvent(event)

    def hoverLeaveEvent(self, event):
        self.setBrush(QBrush(QColor("#ffffff")))
        super().hoverLeaveEvent(event)


class HexGridScene(QGraphicsScene):
    def __init__(self, rows, cols, size, spacing=0.0, parent=None):
        super().__init__(parent)
        self.rows = rows
        self.cols = cols
        self.size = size                # distance center -> corner
        self.spacing = spacing          # extra gap (set 0 for perfect tiling)
        self.build_grid()

    def build_grid(self):
        self.clear()
        s = self.size
        # pointy-top metrics:
        hex_w = math.sqrt(3) * s      # width corner-to-corner horizontally
        hex_h = 2 * s                  # height corner-to-corner vertically
        horiz = hex_w                  # horizontal distance between columns' centers
        vert = 3.0/4.0 * hex_h         # vertical distance between rows' centers

        left_margin = s + 40  # Increased for row labels
        top_margin = s + 30   # Increased for column labels

        # Create column headers (A, B, C, D, ...)
        for c in range(self.cols):
            cx = left_margin + c * horiz + (0 % 2) * (horiz / 2.0)  # Use row 0 for positioning
            cy = top_margin - 25  # Above the grid
            
            col_name = column_to_excel_name(c)
            text_item = QGraphicsTextItem(col_name)
            text_item.setDefaultTextColor(QColor("#374151"))
            text_item.setFont(QFont("Arial", 10, QFont.Bold))
            
            # Center the text
            text_rect = text_item.boundingRect()
            text_item.setPos(cx - text_rect.width()/2, cy - text_rect.height()/2)
            self.addItem(text_item)

        # Create row headers (1, 2, 3, 4, ...)
        for r in range(self.rows):
            cx = left_margin - 30  # To the left of the grid
            cy = top_margin + r * vert
            
            row_name = str(r + 1)  # 1-based row numbering
            text_item = QGraphicsTextItem(row_name)
            text_item.setDefaultTextColor(QColor("#374151"))
            text_item.setFont(QFont("Arial", 10, QFont.Bold))
            
            # Center the text
            text_rect = text_item.boundingRect()
            text_item.setPos(cx - text_rect.width()/2, cy - text_rect.height()/2)
            self.addItem(text_item)

        # Create hex grid
        for r in range(self.rows):
            for c in range(self.cols):
                cx = left_margin + c * horiz + (r % 2) * (horiz / 2.0)
                cy = top_margin + r * vert
                pts = hex_corners(cx, cy, s - self.spacing)
                poly = QPolygonF(pts)
                item = HexItem(poly, r, c)
                item.setAcceptHoverEvents(True)
                self.addItem(item)

        # set scene rect to tightly enclose grid including labels
        width = left_margin + (self.cols - 1) * horiz + hex_w + 40
        height = top_margin + (self.rows - 1) * vert + hex_h + 40
        self.setSceneRect(QRectF(0, 0, width, height))


class HexView(QGraphicsView):
    def __init__(self, scene: HexGridScene, parent=None):
        super().__init__(scene, parent)
        self.setRenderHints(self.renderHints() | QPainter.Antialiasing | QPainter.SmoothPixmapTransform)
        self.setViewportUpdateMode(QGraphicsView.BoundingRectViewportUpdate)
        self.setDragMode(QGraphicsView.NoDrag)
        self._last_mouse_pos = None
        self._pan_active = False
        self._zoom = 1.0

        # Nice background color
        self.setBackgroundBrush(QBrush(QColor("#f8fafc")))

    # Mouse pan
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and event.modifiers() & Qt.ControlModifier:
            self._pan_active = True
            self._last_mouse_pos = event.pos()
            self.setCursor(Qt.ClosedHandCursor)
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._pan_active and self._last_mouse_pos is not None:
            delta = event.pos() - self._last_mouse_pos
            self._last_mouse_pos = event.pos()
            self.horizontalScrollBar().setValue(self.horizontalScrollBar().value() - delta.x())
            self.verticalScrollBar().setValue(self.verticalScrollBar().value() - delta.y())
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton and self._pan_active:
            self._pan_active = False
            self.setCursor(Qt.ArrowCursor)
            event.accept()
            return
        super().mouseReleaseEvent(event)

    # Wheel to zoom
    def wheelEvent(self, event):
        angle = event.angleDelta().y()
        factor = 1.001 ** angle
        self._zoom *= factor
        self.scale(factor, factor)


class MainWindow(QMainWindow):
    def __init__(self, rows=20, cols=12, hex_size=30):
        super().__init__()
        self.setWindowTitle("Hexcel")
        scene = HexGridScene(rows=rows, cols=cols, size=hex_size)
        view = HexView(scene)
        self.setCentralWidget(view)
        self.resize(1000, 700)


def main():
    app = QApplication(sys.argv)
    # create and show
    win = MainWindow(rows=16, cols=22, hex_size=28)
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
