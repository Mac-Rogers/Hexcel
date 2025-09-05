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
        self.is_fill_selected = False  # For drag-fill selection
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

    def update_appearance(self):
        """Update visual appearance based on selection state"""
        if self.selected:
            self.setPen(QPen(QColor("#0F773D"), 3))  # Thick green border
            self.setBrush(QBrush(QColor("#ffffff")))  # Keep white fill
            self.setZValue(1)
            self.text_item.setZValue(2)
        elif self.is_fill_selected:
            self.setPen(QPen(QColor("#0F773D"), 2))  # Blue border for fill selection
            self.setBrush(QBrush(QColor("#ffffff")))  # Light blue fill
            self.setZValue(1)
            self.text_item.setZValue(2)
        else:
            self.setPen(QPen(QColor("#A4A7A6"), 0.8))
            self.setBrush(QBrush(QColor("#ffffff")))
            self.setZValue(0)
            self.text_item.setZValue(0)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            # Clear previous selection
            self.scene_ref.clear_selection()
            # Select this hex
            self.selected = True
            self.update_appearance()
            
            # Start drag operation
            self.scene_ref.start_drag_operation(self, event.scenePos())
            
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
        if not self.selected and not self.is_fill_selected:
            self.setBrush(QBrush(QColor("#f1f5f9")))
        super().hoverEnterEvent(event)

    def hoverLeaveEvent(self, event):
        if not self.selected and not self.is_fill_selected:
            self.setBrush(QBrush(QColor("#ffffff")))
            self.setPen(QPen(QColor("#A4A7A6"), 0.8))
        super().hoverLeaveEvent(event)



class InfiniteHexGridScene(QGraphicsScene):
    def __init__(self, size, spacing=0.0, parent=None):
        super().__init__(parent)
        self.hex_size = size
        self.spacing = spacing
        self.hex_items = {}
        self.current_editor = None
        self.visible_items = set()
        
        # Drag operation state
        self.drag_active = False
        self.drag_start_item = None
        self.drag_start_pos = None
        self.fill_selection = set()
        
        # Calculate hex metrics
        s = self.hex_size
        self.hex_w = math.sqrt(3) * s
        self.hex_h = 2 * s
        self.horiz = self.hex_w
        self.vert = 3.0/4.0 * self.hex_h
        
        self.left_margin = s + 40
        self.top_margin = s + 30
        
        # Set scene rect starting from origin (0,0)
        large_size = 100000
        self.setSceneRect(QRectF(0, 0, large_size, large_size))

    def get_hex_position(self, row, col):
        """Calculate the center position of a hex at given row, col"""
        cx = self.left_margin + col * self.horiz + (row % 2) * (self.horiz / 2.0)
        cy = self.top_margin + row * self.vert
        return cx, cy

    def get_hex_at_position(self, scene_pos):
        """Find which hex cell is at the given scene position"""
        # Approximate calculation - find closest hex
        x, y = scene_pos.x(), scene_pos.y()
        
        # Rough estimate of row and column
        approx_row = max(0, int((y - self.top_margin) / self.vert))
        approx_col = max(0, int((x - self.left_margin) / self.horiz))
        
        # Check nearby cells to find the exact one
        min_distance = float('inf')
        closest_cell = None
        
        for r in range(max(0, approx_row - 2), approx_row + 3):
            for c in range(max(0, approx_col - 2), approx_col + 3):
                if (r, c) in self.hex_items:
                    hex_item = self.hex_items[(r, c)]
                    if isinstance(hex_item, HexItem):
                        hex_center = hex_item.boundingRect().center()
                        hex_pos = hex_item.scenePos() + hex_center
                        distance = ((hex_pos.x() - x) ** 2 + (hex_pos.y() - y) ** 2) ** 0.5
                        if distance < min_distance and distance < self.hex_size:
                            min_distance = distance
                            closest_cell = (r, c)
        
        return closest_cell

    def get_hex_neighbors(self, row, col):
        """Get the 6 neighboring hex coordinates"""
        neighbors = []
        if row % 2 == 0:  # Even row
            offsets = [(-1, -1), (-1, 0), (0, -1), (0, 1), (1, -1), (1, 0)]
        else:  # Odd row
            offsets = [(-1, 0), (-1, 1), (0, -1), (0, 1), (1, 0), (1, 1)]
        
        for dr, dc in offsets:
            new_row, new_col = row + dr, col + dc
            if new_row >= 0 and new_col >= 0:
                neighbors.append((new_row, new_col))
        return neighbors

    def get_line_cells(self, start_row, start_col, end_row, end_col):
        """Get all cells in a line from start to end (hexagonal path)"""
        if start_row == end_row and start_col == end_col:
            return [(start_row, start_col)]
        
        cells = [(start_row, start_col)]
        
        # Determine direction
        row_diff = end_row - start_row
        col_diff = end_col - start_col
        
        # Simple linear interpolation for hex grid
        current_row, current_col = start_row, start_col
        
        while (current_row, current_col) != (end_row, end_col):
            # Move towards target
            if abs(row_diff) > abs(col_diff):
                # Primarily vertical movement
                if row_diff > 0:
                    current_row += 1
                else:
                    current_row -= 1
            else:
                # Primarily horizontal movement
                if col_diff > 0:
                    current_col += 1
                else:
                    current_col -= 1
            
            # Ensure we don't go negative
            current_row = max(0, current_row)
            current_col = max(0, current_col)
            
            cells.append((current_row, current_col))
            
            # Update differences
            row_diff = end_row - current_row
            col_diff = end_col - current_col
            
            # Safety check to prevent infinite loops
            if len(cells) > 100:
                break
        
        return cells

    def start_drag_operation(self, hex_item, start_pos):
        """Start a drag operation from the given hex item"""
        self.drag_active = True
        self.drag_start_item = hex_item
        self.drag_start_pos = start_pos
        self.fill_selection = set()

    def update_drag_selection(self, current_pos):
        """Update the drag selection based on current mouse position"""
        if not self.drag_active or not self.drag_start_item:
            return
        
        # Find hex at current position
        current_cell = self.get_hex_at_position(current_pos)
        if not current_cell:
            return
        
        current_row, current_col = current_cell
        start_row, start_col = self.drag_start_item.row, self.drag_start_item.col
        
        # Clear previous fill selection
        for row, col in self.fill_selection:
            if (row, col) in self.hex_items:
                hex_item = self.hex_items[(row, col)]
                if isinstance(hex_item, HexItem):
                    hex_item.is_fill_selected = False
                    hex_item.update_appearance()
        
        # Get cells in line from start to current
        line_cells = self.get_line_cells(start_row, start_col, current_row, current_col)
        
        # Update fill selection
        self.fill_selection = set(line_cells)
        for row, col in self.fill_selection:
            # Create hex if it doesn't exist
            if (row, col) not in self.hex_items:
                self.create_hex_item(row, col)
            
            hex_item = self.hex_items[(row, col)]
            if isinstance(hex_item, HexItem):
                if (row, col) != (start_row, start_col):  # Don't change the original selection
                    hex_item.is_fill_selected = True
                hex_item.update_appearance()

    def finish_drag_operation(self):
        """Finish the drag operation and fill cells with the original value"""
        if not self.drag_active or not self.drag_start_item:
            return
        
        original_value = self.drag_start_item.value
        
        # Fill all selected cells with the original value
        for row, col in self.fill_selection:
            if (row, col) in self.hex_items:
                hex_item = self.hex_items[(row, col)]
                if isinstance(hex_item, HexItem):
                    hex_item.set_value(original_value)
                    hex_item.is_fill_selected = False
                    hex_item.update_appearance()
        
        # Reset drag state
        self.drag_active = False
        self.drag_start_item = None
        self.drag_start_pos = None
        self.fill_selection = set()
        
        print(f"Filled {len(self.fill_selection)} cells with value: '{original_value}'")

    def clear_selection(self):
        """Clear all selections"""
        for item in self.hex_items.values():
            if hasattr(item, 'selected') and hasattr(item, 'is_fill_selected'):
                item.selected = False
                item.is_fill_selected = False
                item.update_appearance()

    def mouseMoveEvent(self, event):
        """Handle mouse move for drag operations"""
        if self.drag_active:
            self.update_drag_selection(event.scenePos())
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        """Handle mouse release to finish drag operations"""
        if event.button() == Qt.LeftButton and self.drag_active:
            self.finish_drag_operation()
        super().mouseReleaseEvent(event)

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
        """Update items when view is first shown and center on origin"""
        super().showEvent(event)
        # Start at the top-left (origin)
        self.ensureVisible(0, 0, 1, 1, 0, 0)  # Show origin without margins
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
