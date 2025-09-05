#!/usr/bin/env python3

import sys
import math
from PyQt5.QtGui import QBrush, QPen, QPolygonF, QColor, QPainter, QFont
from PyQt5.QtCore import Qt, QPointF, QRectF
from PyQt5.QtWidgets import (QApplication, QGraphicsView, QGraphicsScene, 
                           QGraphicsPolygonItem, QMainWindow, QGraphicsTextItem,
                           QLineEdit, QGraphicsProxyWidget)


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
    if col_index < 0:
        return ""
    result = ""
    while col_index >= 0:
        result = chr(65 + (col_index % 26)) + result
        col_index = col_index // 26 - 1
        if col_index < 0:
            break
    return result


class HexItem(QGraphicsPolygonItem):
    def __init__(self, polygon, row, col, scene):
        super().__init__(polygon)
        self.row = row
        self.col = col
        self.scene_ref = scene
        self.value = ""
        self.selected = False
        self.setPen(QPen(QColor("#A4A7A6"), 0.8))
        self.setBrush(QBrush(QColor("#ffffff")))
        
        # Create text item for displaying value
        self.text_item = QGraphicsTextItem()
        self.text_item.setDefaultTextColor(QColor("#1f2937"))
        self.text_item.setFont(QFont("Arial", 8))
        self.text_item.setParentItem(self)
        self.update_text_position()

    def update_text_position(self):
        """Center the text in the hexagon"""
        if self.text_item:
            text_rect = self.text_item.boundingRect()
            hex_rect = self.boundingRect()
            x = hex_rect.center().x() - text_rect.width() / 2
            y = hex_rect.center().y() - text_rect.height() / 2
            self.text_item.setPos(x, y)

    def set_value(self, value):
        """Set the value and update display"""
        self.value = str(value)
        self.text_item.setPlainText(self.value)
        self.update_text_position()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            # Clear previous selection
            self.scene_ref.clear_selection()
            # Select this hex with highlighted border
            self.selected = True
            self.setPen(QPen(QColor("#0F773D"), 3))  # Thick green border
            self.setBrush(QBrush(QColor("#ffffff")))  # Keep white fill
            
            # Bring selected hex to front, but text even higher
            self.setZValue(1)
            self.text_item.setZValue(2)  # Text stays on top
            
            print(f"Selected cell [{self.row}][{self.col}]: '{self.value}'")
        super().mousePressEvent(event)

    def mouseDoubleClickEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.start_editing()
        super().mouseDoubleClickEvent(event)

    def start_editing(self):
        """Start editing the cell value"""
        self.scene_ref.start_cell_editing(self)

    def hoverEnterEvent(self, event):
        if not self.selected:
            self.setBrush(QBrush(QColor("#f1f5f9")))
        super().hoverEnterEvent(event)

    def hoverLeaveEvent(self, event):
        if not self.selected:
            self.setBrush(QBrush(QColor("#ffffff")))
            self.setPen(QPen(QColor("#A4A7A6"), 0.8))  # Reset border on hover leave
        super().hoverLeaveEvent(event)


class InfiniteHexGridScene(QGraphicsScene):
    def __init__(self, size, spacing=0.0, parent=None):
        super().__init__(parent)
        self.hex_size = size
        self.spacing = spacing
        self.hex_items = {}  # Store hex items by (row, col)
        self.current_editor = None
        self.visible_items = set()  # Track currently visible items
        
        # Calculate hex metrics
        s = self.hex_size
        self.hex_w = math.sqrt(3) * s
        self.hex_h = 2 * s
        self.horiz = self.hex_w
        self.vert = 3.0/4.0 * self.hex_h
        
        self.left_margin = s + 40
        self.top_margin = s + 30
        
        # Set a very large scene rect to allow infinite scrolling
        large_size = 1000000
        self.setSceneRect(QRectF(-large_size, -large_size, 2*large_size, 2*large_size))

    def get_hex_position(self, row, col):
        """Calculate the center position of a hex at given row, col"""
        cx = self.left_margin + col * self.horiz + (row % 2) * (self.horiz / 2.0)
        cy = self.top_margin + row * self.vert
        return cx, cy

    def get_visible_range(self, view_rect):
        """Calculate which hex cells should be visible in the given view rectangle"""
        # Add some padding to ensure smooth scrolling
        padding = max(self.hex_w, self.hex_h) * 2
        expanded_rect = view_rect.adjusted(-padding, -padding, padding, padding)
        
        # Calculate approximate row/col ranges
        min_row = max(0, int((expanded_rect.top() - self.top_margin) / self.vert) - 1)
        max_row = int((expanded_rect.bottom() - self.top_margin) / self.vert) + 2
        
        min_col = max(0, int((expanded_rect.left() - self.left_margin) / self.horiz) - 2)
        max_col = int((expanded_rect.right() - self.left_margin) / self.horiz) + 3
        
        return min_row, max_row, min_col, max_col

    def create_hex_item(self, row, col):
        """Create a hex item at the given row, col"""
        if (row, col) in self.hex_items:
            return self.hex_items[(row, col)]
            
        cx, cy = self.get_hex_position(row, col)
        pts = hex_corners(cx, cy, self.hex_size - self.spacing)
        poly = QPolygonF(pts)
        item = HexItem(poly, row, col, self)
        item.setAcceptHoverEvents(True)
        
        self.hex_items[(row, col)] = item
        self.addItem(item)
        return item

    def create_header_items(self, view_rect):
        """Create column and row headers for visible area"""
        min_row, max_row, min_col, max_col = self.get_visible_range(view_rect)
        
        # Create column headers
        for c in range(min_col, max_col + 1):
            header_key = ('col_header', c)
            if header_key not in self.hex_items:
                cx, _ = self.get_hex_position(0, c)
                cy = self.top_margin - 40
                
                col_name = column_to_excel_name(c)
                text_item = QGraphicsTextItem(col_name)
                text_item.setDefaultTextColor(QColor("#374151"))
                text_item.setFont(QFont("Arial", 10, QFont.Bold))
                
                text_rect = text_item.boundingRect()
                text_item.setPos(cx - text_rect.width()/2, cy - text_rect.height()/2)
                text_item.setZValue(5)  # Keep headers on top
                
                self.hex_items[header_key] = text_item
                self.addItem(text_item)
        
        # Create row headers
        for r in range(min_row, max_row + 1):
            header_key = ('row_header', r)
            if header_key not in self.hex_items:
                cx = self.left_margin - 40
                _, cy = self.get_hex_position(r, 0)
                
                row_name = str(r + 1)
                text_item = QGraphicsTextItem(row_name)
                text_item.setDefaultTextColor(QColor("#374151"))
                text_item.setFont(QFont("Arial", 10, QFont.Bold))
                
                text_rect = text_item.boundingRect()
                text_item.setPos(cx - text_rect.width()/2, cy - text_rect.height()/2)
                text_item.setZValue(5)  # Keep headers on top
                
                self.hex_items[header_key] = text_item
                self.addItem(text_item)

    def update_visible_items(self, view_rect):
        """Update which items are visible and create/remove as needed"""
        min_row, max_row, min_col, max_col = self.get_visible_range(view_rect)
        
        # Create headers
        self.create_header_items(view_rect)
        
        # Track new visible items
        new_visible = set()
        
        # Create hex items for visible area
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                if r >= 0 and c >= 0:  # Only positive coordinates
                    item = self.create_hex_item(r, c)
                    new_visible.add((r, c))
        
        # Remove items that are too far away (optional optimization)
        items_to_remove = []
        for key, item in self.hex_items.items():
            if isinstance(key, tuple) and len(key) == 2:
                if isinstance(key[0], int) and isinstance(key[1], int):  # Regular hex item
                    r, c = key
                    if (r < min_row - 10 or r > max_row + 10 or 
                        c < min_col - 10 or c > max_col + 10):
                        items_to_remove.append(key)
        
        # Remove distant items to save memory
        for key in items_to_remove:
            item = self.hex_items[key]
            self.removeItem(item)
            del self.hex_items[key]
        
        self.visible_items = new_visible

    def clear_selection(self):
        for item in self.hex_items.values():
            if hasattr(item, 'selected') and item.selected:
                item.selected = False
                item.setPen(QPen(QColor("#A4A7A6"), 0.8))
                item.setBrush(QBrush(QColor("#ffffff")))
                item.setZValue(0)
                if hasattr(item, 'text_item'):
                    item.text_item.setZValue(0)

    def start_cell_editing(self, hex_item):
        """Start editing a cell"""
        if self.current_editor:
            self.finish_editing()
    
        # Create line edit
        line_edit = QLineEdit()
        line_edit.setText(hex_item.value)
        line_edit.setAlignment(Qt.AlignCenter)
        line_edit.setStyleSheet("""
            QLineEdit {
                border: none;
                background: transparent;
                font-size: 8pt;
                color: #1f2937;
            }
        """)
    
        # Create proxy widget to embed QLineEdit in scene
        proxy = QGraphicsProxyWidget()
        proxy.setWidget(line_edit)
        
        # Position the editor over the hex
        hex_rect = hex_item.boundingRect()
        editor_width = 60
        editor_height = 20
        x = hex_rect.center().x() - editor_width / 2
        y = hex_rect.center().y() - editor_height / 2
        proxy.setPos(x, y)
        proxy.resize(editor_width, editor_height)
        proxy.setZValue(10)
        
        self.addItem(proxy)
        self.current_editor = {
            'proxy': proxy,
            'line_edit': line_edit,
            'hex_item': hex_item
        }
    
        line_edit.returnPressed.connect(self.finish_editing)
        line_edit.editingFinished.connect(self.finish_editing)
        line_edit.setFocus()
        line_edit.selectAll()

    def finish_editing(self):
        """Finish editing and update the cell value"""
        if not self.current_editor:
            return

        line_edit = self.current_editor['line_edit']
        hex_item = self.current_editor['hex_item']
        proxy = self.current_editor['proxy']

        new_value = line_edit.text()
        hex_item.set_value(new_value)

        self.removeItem(proxy)
        proxy.deleteLater()
        self.current_editor = None
        print(f"Updated cell [{hex_item.row}][{hex_item.col}] = '{new_value}'")

    def keyPressEvent(self, event):
        """Handle key presses for navigation and editing"""
        if event.key() == Qt.Key_Escape and self.current_editor:
            self.removeItem(self.current_editor['proxy'])
            self.current_editor['proxy'].deleteLater()
            self.current_editor = None
        super().keyPressEvent(event)


class InfiniteHexView(QGraphicsView):
    def __init__(self, scene: InfiniteHexGridScene, parent=None):
        super().__init__(scene, parent)
        self.setRenderHints(self.renderHints() | QPainter.Antialiasing | QPainter.SmoothPixmapTransform)
        self.setViewportUpdateMode(QGraphicsView.BoundingRectViewportUpdate)
        self.setDragMode(QGraphicsView.NoDrag)
        self._last_mouse_pos = None
        self._pan_active = False
        self._zoom = 1.0

        # Nice background color
        self.setBackgroundBrush(QBrush(QColor("#f8fafc")))
        
        # Connect scroll events to update visible items
        self.horizontalScrollBar().valueChanged.connect(self.update_scene_items)
        self.verticalScrollBar().valueChanged.connect(self.update_scene_items)

    def update_scene_items(self):
        """Update visible items when scrolling"""
        visible_rect = self.mapToScene(self.viewport().rect()).boundingRect()
        self.scene().update_visible_items(visible_rect)

    def resizeEvent(self, event):
        """Update items when view is resized"""
        super().resizeEvent(event)
        self.update_scene_items()

    def showEvent(self, event):
        """Update items when view is first shown"""
        super().showEvent(event)
        self.update_scene_items()

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
        # Update items after zoom
        self.update_scene_items()


class MainWindow(QMainWindow):
    def __init__(self, hex_size=28):
        super().__init__()
        self.setWindowTitle("Hexcel")
        
        # Start in fullscreen
        self.showMaximized()

        self.scene = InfiniteHexGridScene(size=hex_size)
        self.view = InfiniteHexView(self.scene)
        self.setCentralWidget(self.view)

    def keyPressEvent(self, event):
        """Forward key events to the scene"""
        if event.key() == Qt.Key_F11:
            # Toggle fullscreen with F11
            if self.isFullScreen():
                self.showMaximized()
            else:
                self.showFullScreen()
        else:
            self.scene.keyPressEvent(event)
        super().keyPressEvent(event)


def main():
    app = QApplication(sys.argv)
    win = MainWindow(hex_size=28)
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()

