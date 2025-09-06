#!/usr/bin/env python3

import sys
import math
import os  # ribbon path utils
import glob
import re   # <<< ADDED: simple formula parsing

from PyQt5.QtGui import QBrush, QPen, QPolygonF, QColor, QPainter, QFont, QPixmap
from PyQt5.QtCore import Qt, QPointF, QRectF, pyqtSignal
from PyQt5.QtWidgets import (QApplication, QGraphicsView, QGraphicsScene,
                             QGraphicsPolygonItem, QMainWindow, QGraphicsTextItem,
                             QLineEdit, QGraphicsProxyWidget,
                             QWidget, QHBoxLayout, QVBoxLayout, QLabel, QAction)

# ---------- ribbon helpers ----------
def _base_dir():
    if '__file__' in globals():
        return os.path.dirname(os.path.abspath(__file__))
    return os.getcwd()

def find_ribbon_image():
    base = _base_dir()
    candidates = [
        os.path.join(base, "ribbon.png"),
        os.path.join(base, "ribbon.jpg"),
        os.path.join(base, "ribbon.jpeg"),
        os.path.join(base, "ribbon.bmp"),
    ]
    candidates += glob.glob(os.path.join(base, "Screenshot*.*"))
    for p in candidates:
        ext = os.path.splitext(p)[1].lower()
        if os.path.isfile(p) and ext in {".png", ".jpg", ".jpeg", ".bmp", ".gif"}:
            return p
    return None

class RibbonBanner(QWidget):
    def __init__(self, image_path: str, height: int = 274, parent=None):
        super().__init__(parent)
        self._img_path = image_path
        self._pix = QPixmap(self._img_path) if self._img_path and os.path.exists(self._img_path) else QPixmap()
        self._label = QLabel()
        self._label.setAlignment(Qt.AlignCenter)
        self._label.setStyleSheet("QLabel { background: #E5E7EB; }")
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)
        lay.addWidget(self._label)
        self.setFixedHeight(height)
        self._rescale()

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self._rescale()

    def _rescale(self):
        if not self._pix.isNull():
            scaled = self._pix.scaled(
                self.width(), self.height(),
                Qt.KeepAspectRatioByExpanding,
                Qt.SmoothTransformation
            )
            self._label.setPixmap(scaled)
            self._label.setText("")
        else:
            self._label.setPixmap(QPixmap())
            self._label.setText(f" Ribbon image not found at: {self._img_path or '(none)'} ")
            self._label.setAlignment(Qt.AlignCenter)
            self._label.setStyleSheet("QLabel { background:#E5E7EB; color:#374151; padding:6px; }")

# ---------- grid + formula utils ----------
def hex_corners(center_x, center_y, size):
    corners = []
    for i in range(6):
        angle_deg = 30 + 60 * i
        angle_rad = math.radians(angle_deg)
        x = center_x + size * math.cos(angle_rad)
        y = center_y + size * math.sin(angle_rad)
        corners.append(QPointF(x, y))
    return corners

def column_to_excel_name(col_index):
    if col_index < 0:
        return ""
    result = ""
    while col_index >= 0:
        result = chr(65 + (col_index % 26)) + result
        col_index = col_index // 26 - 1
        if col_index < 0:
            break
    return result

# >>> ADDED: parse "AB12" -> (row, col)
def excel_col_to_index(name: str) -> int:
    name = name.strip().upper()
    total = 0
    for ch in name:
        if not ('A' <= ch <= 'Z'):
            return -1
        total = total * 26 + (ord(ch) - 64)
    return total - 1

_cell_re = re.compile(r'^([A-Z]+)(\d+)$')
def parse_cell_ref(ref: str):
    m = _cell_re.match(ref.strip().upper())
    if not m:
        return None
    col_s, row_s = m.groups()
    col = excel_col_to_index(col_s)
    row = int(row_s) - 1
    return (row, col) if (row >= 0 and col >= 0) else None

class HexItem(QGraphicsPolygonItem):
    def __init__(self, polygon, row, col, scene):
        super().__init__(polygon)
        self.row = row
        self.col = col
        self.scene_ref = scene
        self.value = ""               # raw string the user typed (formula or literal)
        self.selected = False
        self.is_fill_selected = False
        self.setPen(QPen(QColor("#A4A7A6"), 0.8))
        self.setBrush(QBrush(QColor("#ffffff")))
        self.text_item = QGraphicsTextItem()
        self.text_item.setDefaultTextColor(QColor("#1f2937"))
        self.text_item.setFont(QFont("Arial", 8))
        self.text_item.setParentItem(self)
        self.update_text_position()

    def update_text_position(self):
        if self.text_item:
            text_rect = self.text_item.boundingRect()
            hex_rect = self.boundingRect()
            x = hex_rect.center().x() - text_rect.width() / 2
            y = hex_rect.center().y() - text_rect.height() / 2
            self.text_item.setPos(x, y)

    def _display_text_for_value(self, value_str: str) -> str:
        v = "" if value_str is None else str(value_str)
        if v.startswith("="):
            try:
                return self.scene_ref.evaluate_formula_str(v, caller=self)
            except Exception:
                # show the raw formula if something went wrong
                return v
        return v

    def set_value(self, value):
        """Set raw value (string) and update what you SEE (formula result or literal)."""
        self.value = "" if value is None else str(value)
        display = self._display_text_for_value(self.value)
        self.text_item.setPlainText(display)
        self.update_text_position()

    def update_appearance(self):
        if self.selected:
            self.setPen(QPen(QColor("#0F773D"), 3))
            self.setBrush(QBrush(QColor("#ffffff")))
            self.setZValue(1)
            self.text_item.setZValue(2)
        elif self.is_fill_selected:
            self.setPen(QPen(QColor("#0F773D"), 2))
            self.setBrush(QBrush(QColor("#ffffff")))
            self.setZValue(1)
            self.text_item.setZValue(2)
        else:
            self.setPen(QPen(QColor("#A4A7A6"), 0.8))
            self.setBrush(QBrush(QColor("#ffffff")))
            self.setZValue(0)
            self.text_item.setZValue(0)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.scene_ref.clear_selection()
            self.selected = True
            self.update_appearance()
            self.scene_ref.cellSelected.emit(self)  # sync Name/fx
            # Start potential drag (or edit if released without drag)
            self.scene_ref.start_drag_operation(self, event.scenePos())
            print(f"Selected cell [{self.row}][{self.col}]: '{self.value}'")
        super().mousePressEvent(event)

    def mouseDoubleClickEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.start_editing()
        super().mouseDoubleClickEvent(event)

    def start_editing(self):
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
    cellSelected = pyqtSignal(object)

    def __init__(self, size, spacing=0.0, parent=None):
        super().__init__(parent)
        self.hex_size = size
        self.spacing = spacing
        self.hex_items = {}
        self.current_editor = None
        self.visible_items = set()

        # drag-fill state
        self.drag_active = False
        self.drag_start_item = None
        self.drag_start_pos = None
        self.fill_selection = set()

        s = self.hex_size
        self.hex_w = math.sqrt(3) * s
        self.hex_h = 2 * s
        self.horiz = self.hex_w
        self.vert = 3.0 / 4.0 * self.hex_h

        self.left_margin = s + 40
        self.top_margin = s + 30

        large_size = 100000
        self.setSceneRect(QRectF(0, 0, large_size, large_size))

    # ---------- formula engine helpers ----------
    def _get_cell_item(self, row: int, col: int):
        it = self.hex_items.get((row, col))
        return it if isinstance(it, HexItem) else None

    def _is_numeric_cell(self, row: int, col: int) -> bool:
        it = self._get_cell_item(row, col)
        if not it: return False
        try:
            float(it.value)
            return True
        except (ValueError, TypeError):
            return False

    def _split_args(self, s: str):
        args, buf, depth = [], [], 0
        for ch in s:
            if ch == '(':
                depth += 1; buf.append(ch)
            elif ch == ')':
                depth = max(0, depth - 1); buf.append(ch)
            elif ch == ',' and depth == 0:
                token = ''.join(buf).strip()
                if token: args.append(token)
                buf = []
            else:
                buf.append(ch)
        token = ''.join(buf).strip()
        if token: args.append(token)
        return args

    def _expand_range(self, a_ref: str, b_ref: str):
        a = parse_cell_ref(a_ref); b = parse_cell_ref(b_ref)
        if not a or not b:
            return []
        (r1, c1), (r2, c2) = a, b
        if r1 > r2: r1, r2 = r2, r1
        if c1 > c2: c1, c2 = c2, c1
        out = []
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                out.append((r, c))
        return out

    def _values_for_math(self, token: str):
        t = token.strip()
        if not t:
            return []
        # Range A1:B3
        if ':' in t:
            left, right = t.split(':', 1)
            cells = self._expand_range(left, right)
            vals = []
            for (r, c) in cells:
                it = self._get_cell_item(r, c)
                if it:
                    try:
                        vals.append(float(it.value))
                    except (ValueError, TypeError):
                        pass
            return vals
        # Single cell A1
        rc = parse_cell_ref(t)
        if rc:
            r, c = rc
            it = self._get_cell_item(r, c)
            if it:
                try:
                    return [float(it.value)]
                except (ValueError, TypeError):
                    return []
            return []
        # Nested FUNC(...)
        if re.match(r'^[A-Za-z_][A-Za-z0-9_]*\s*\(.*\)$', t):
            val_str = self.evaluate_formula_str('=' + t)
            try:
                return [float(val_str)]
            except (ValueError, TypeError):
                return []
        # Plain number
        try:
            return [float(t)]
        except ValueError:
            return []

    def _count_numerics(self, token: str) -> int:
        t = token.strip()
        if not t:
            return 0
        if ':' in t:
            left, right = t.split(':', 1)
            count = 0
            for (r, c) in self._expand_range(left, right):
                if self._is_numeric_cell(r, c):
                    count += 1
            return count
        rc = parse_cell_ref(t)
        if rc:
            r, c = rc
            return 1 if self._is_numeric_cell(r, c) else 0
        if re.match(r'^[A-Za-z_][A-Za-z0-9_]*\s*\(.*\)$', t):
            v = self.evaluate_formula_str('=' + t)
            try:
                float(v); return 1
            except (ValueError, TypeError):
                return 0
        try:
            float(t); return 1
        except ValueError:
            return 0

    def evaluate_formula_str(self, text: str, caller: "HexItem" = None) -> str:
        """
        Supported:
          =SUM(...), =AVERAGE/AVG(...), =MIN(...), =MAX(...),
          =COUNT(...), =PRODUCT(...), =ABS(x)
        Args can be numbers, A1 refs, A1:B3 ranges, or nested FUNC(...).
        """
        s = text.strip()
        if not s.startswith('='):
            return s
        body = s[1:].strip()
        m = re.match(r'^([A-Za-z_][A-Za-z0-9_]*)\s*\((.*)\)$', body)
        if not m:
            return text  # not recognized, show raw
        func_name, inner = m.group(1).upper(), m.group(2)
        args = self._split_args(inner)

        if func_name in ('SUM', 'ADD'):
            vals = []; [vals.extend(self._values_for_math(t)) for t in args]
            total = sum(vals) if vals else 0.0
            return str(int(total)) if total.is_integer() else str(total)

        if func_name in ('AVERAGE', 'AVG', 'MEAN'):
            vals = []; [vals.extend(self._values_for_math(t)) for t in args]
            if not vals: return "0"
            avg = sum(vals)/len(vals)
            return str(int(avg)) if avg.is_integer() else str(avg)

        if func_name == 'MIN':
            vals = []; [vals.extend(self._values_for_math(t)) for t in args]
            if not vals: return "0"
            mn = min(vals)
            return str(int(mn)) if mn.is_integer() else str(mn)

        if func_name == 'MAX':
            vals = []; [vals.extend(self._values_for_math(t)) for t in args]
            if not vals: return "0"
            mx = max(vals)
            return str(int(mx)) if mx.is_integer() else str(mx)

        if func_name == 'COUNT':
            cnt = 0
            for t in args:
                cnt += self._count_numerics(t)
            return str(cnt)

        if func_name == 'PRODUCT':
            vals = []; [vals.extend(self._values_for_math(t)) for t in args]
            if not vals: return "0"
            prod = 1.0
            for v in vals: prod *= v
            return str(int(prod)) if prod.is_integer() else str(prod)

        if func_name == 'ABS':
            vals = []; [vals.extend(self._values_for_math(t)) for t in args]
            x = abs(vals[0]) if vals else 0.0
            return str(int(x)) if x.is_integer() else str(x)

        return text  # unknown -> show raw

    def recalc_all(self):
        """Re-evaluate display text for all formula cells currently created."""
        for key, item in self.hex_items.items():
            if isinstance(key, tuple) and isinstance(item, HexItem):
                if isinstance(key[0], int) and isinstance(key[1], int):
                    if item.value.startswith('='):
                        try:
                            disp = self.evaluate_formula_str(item.value, caller=item)
                            item.text_item.setPlainText(disp)
                            item.update_text_position()
                        except Exception:
                            # leave as-is on error
                            pass

    # ---------- grid building / visibility ----------
    def get_hex_position(self, row, col):
        cx = self.left_margin + col * self.horiz + (row % 2) * (self.horiz / 2.0)
        cy = self.top_margin + row * self.vert
        return cx, cy

    def get_hex_at_position(self, scene_pos):
        x, y = scene_pos.x(), scene_pos.y()
        approx_row = max(0, int((y - self.top_margin) / self.vert))
        approx_col = max(0, int((x - self.left_margin) / self.horiz))
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

    def start_drag_operation(self, hex_item, start_pos):
        self.drag_active = True
        self.drag_start_item = hex_item
        self.drag_start_pos = start_pos
        self.fill_selection = set([(hex_item.row, hex_item.col)])

    def update_drag_selection(self, current_pos):
        if not self.drag_active or not self.drag_start_item:
            return
        current_cell = self.get_hex_at_position(current_pos)
        if not current_cell:
            return
        current_row, current_col = current_cell
        start_row, start_col = self.drag_start_item.row, self.drag_start_item.col

        for row, col in self.fill_selection:
            if (row, col) in self.hex_items:
                hi = self.hex_items[(row, col)]
                if isinstance(hi, HexItem):
                    hi.is_fill_selected = False
                    hi.update_appearance()

        path = self.get_line_cells(start_row, start_col, current_row, current_col)
        self.fill_selection = set(path)

        for row, col in self.fill_selection:
            if (row, col) != (start_row, start_col):
                if (row, col) not in self.hex_items:
                    self.create_hex_item(row, col)
                hi = self.hex_items[(row, col)]
                if isinstance(hi, HexItem):
                    hi.is_fill_selected = True
                    hi.update_appearance()

    def finish_drag_operation(self):
        if not self.drag_active or not self.drag_start_item:
            return
        original_value = self.drag_start_item.value
        for row, col in self.fill_selection:
            if (row, col) in self.hex_items and (row, col) != (self.drag_start_item.row, self.drag_start_item.col):
                hi = self.hex_items[(row, col)]
                if isinstance(hi, HexItem):
                    hi.set_value(original_value)
                    hi.is_fill_selected = False
                    hi.update_appearance()
        self.drag_active = False
        self.drag_start_item = None
        self.drag_start_pos = None
        self.fill_selection = set()
        # After fills, recalc formulas that might depend on these cells
        self.recalc_all()

    def get_hex_neighbors(self, row, col):
        neighbors = []
        if row % 2 == 0:
            offsets = [(-1, -1), (-1, 0), (0, -1), (0, 1), (1, -1), (1, 0)]
        else:
            offsets = [(-1, 0), (-1, 1), (0, -1), (0, 1), (1, 0), (1, 1)]
        for dr, dc in offsets:
            new_row, new_col = row + dr, col + dc
            if new_row >= 0 and new_col >= 0:
                neighbors.append((new_row, new_col))
        return neighbors

    def get_line_cells(self, start_row, start_col, end_row, end_col):
        if start_row == end_row and start_col == end_col:
            return [(start_row, start_col)]
        cells = [(start_row, start_col)]
        row_diff = end_row - start_row
        col_diff = end_col - start_col
        current_row, current_col = start_row, start_col
        while (current_row, current_col) != (end_row, end_col):
            if abs(row_diff) > abs(col_diff):
                current_row += 1 if row_diff > 0 else -1
            else:
                current_col += 1 if col_diff > 0 else -1
            current_row = max(0, current_row)
            current_col = max(0, current_col)
            cells.append((current_row, current_col))
            row_diff = end_row - current_row
            col_diff = end_col - current_col
            if len(cells) > 100:
                break
        return cells

    def cancel_editing(self):
        if not self.current_editor:
            return
        proxy = self.current_editor['proxy']
        self.removeItem(proxy)
        proxy.deleteLater()
        self.current_editor = None

    def clear_selection(self):
        for item in self.hex_items.values():
            if hasattr(item, 'selected') and hasattr(item, 'is_fill_selected'):
                item.selected = False
                item.is_fill_selected = False
                item.update_appearance()

    def mouseMoveEvent(self, event):
        if self.drag_active:
            self.update_drag_selection(event.scenePos())
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton and self.drag_active:
            if len(self.fill_selection) <= 1:
                self.drag_active = False
                self.start_cell_editing(self.drag_start_item)
            else:
                self.finish_drag_operation()
        super().mouseReleaseEvent(event)

    def get_visible_range(self, view_rect):
        padding = max(self.hex_w, self.hex_h) * 2
        expanded_rect = view_rect.adjusted(-padding, -padding, padding, padding)
        min_row = max(0, int((expanded_rect.top() - self.top_margin) / self.vert) - 1)
        max_row = int((expanded_rect.bottom() - self.top_margin) / self.vert) + 2
        min_col = max(0, int((expanded_rect.left() - self.left_margin) / self.horiz) - 2)
        max_col = int((expanded_rect.right() - self.left_margin) / self.horiz) + 3
        return min_row, max_row, min_col, max_col

    def create_hex_item(self, row, col):
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
        min_row, max_row, min_col, max_col = self.get_visible_range(view_rect)
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
                text_item.setZValue(5)
                self.hex_items[header_key] = text_item
                self.addItem(text_item)
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
                text_item.setZValue(5)
                self.hex_items[header_key] = text_item
                self.addItem(text_item)

    def update_visible_items(self, view_rect):
        min_row, max_row, min_col, max_col = self.get_visible_range(view_rect)
        self.create_header_items(view_rect)
        new_visible = set()
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                if r >= 0 and c >= 0:
                    item = self.create_hex_item(r, c)
                    new_visible.add((r, c))
        items_to_remove = []
        for key, item in self.hex_items.items():
            if isinstance(key, tuple) and len(key) == 2:
                if isinstance(key[0], int) and isinstance(key[1], int):
                    r, c = key
                    if (r < min_row - 10 or r > max_row + 10 or
                        c < min_col - 10 or c > max_col + 10):
                        items_to_remove.append(key)
        for key in items_to_remove:
            item = self.hex_items[key]
            self.removeItem(item)
            del self.hex_items[key]
        self.visible_items = new_visible

    def start_cell_editing(self, hex_item):
        if self.current_editor:
            self.finish_editing()
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
        proxy = QGraphicsProxyWidget()
        proxy.setWidget(line_edit)
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
        self.cellSelected.emit(hex_item)
        print(f"Updated cell [{hex_item.row}][{hex_item.col}] = '{new_value}'")
        # After any edit, recalc formulas everywhere (simple but effective)
        self.recalc_all()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape and self.current_editor:
            self.cancel_editing()
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
        self.setBackgroundBrush(QBrush(QColor("#f8fafc")))
        self.horizontalScrollBar().valueChanged.connect(self.update_scene_items)
        self.verticalScrollBar().valueChanged.connect(self.update_scene_items)

    def update_scene_items(self):
        visible_rect = self.mapToScene(self.viewport().rect()).boundingRect()
        self.scene().update_visible_items(visible_rect)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.update_scene_items()

    def showEvent(self, event):
        super().showEvent(event)
        self.ensureVisible(0, 0, 1, 1, 0, 0)
        self.update_scene_items()

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

    def wheelEvent(self, event):
        angle = event.angleDelta().y()
        factor = 1.001 ** angle
        self._zoom *= factor
        self.scale(factor, factor)
        self.update_scene_items()

class MainWindow(QMainWindow):
    def __init__(self, hex_size=28):
        super().__init__()
        self.setWindowTitle("Hexcel")
        self.showFullScreen()

        self.scene = InfiniteHexGridScene(size=hex_size)
        self.view = InfiniteHexView(self.scene)

        # Name + Formula bar
        self.name_label = QLabel("A1")
        self.name_label.setStyleSheet("QLabel { padding: 2px 6px; color:#111827; }")

        self.formula_edit = QLineEdit()
        self.formula_edit.setPlaceholderText("fx")
        self.formula_edit.setStyleSheet("""
            QLineEdit {
                background: white;
                border: 1px solid #D1D5DB;
                border-radius: 4px;
                padding: 4px 8px;
                font-size: 10pt;
                color: #111827;
            }
        """)
        self.formula_edit.returnPressed.connect(self._apply_formula_edit)

        formula_row = QWidget()
        h = QHBoxLayout(formula_row); h.setContentsMargins(8, 6, 8, 6); h.setSpacing(8)
        name_box_label = QLabel("Name"); name_box_label.setStyleSheet("QLabel { color:#6B7280; }")
        fx_label = QLabel("fx"); fx_label.setStyleSheet("QLabel { color:#6B7280; }")
        h.addWidget(name_box_label)
        h.addWidget(self.name_label, 0)
        h.addSpacing(12)
        h.addWidget(fx_label)
        h.addWidget(self.formula_edit, 1)

        # Ribbon (screenshot) at top
        img_path = find_ribbon_image()
        print("Ribbon path:", img_path if img_path else "(not found in project folder)")
        self.ribbon = RibbonBanner(img_path or "", height=274)

        central = QWidget()
        v = QVBoxLayout(central); v.setContentsMargins(0,0,0,0); v.setSpacing(0)
        v.addWidget(self.ribbon)
        v.addWidget(formula_row)
        v.addWidget(self.view)
        self.setCentralWidget(central)

        # File menu (Cancel Edit + Exit) and small Help item
        # mb = self.menuBar()
        # file_menu = mb.addMenu("&File")
        # cancel_act = QAction("Cancel Edit", self); cancel_act.setShortcut("Esc")
        # cancel_act.triggered.connect(self._cancel_edit)
        # exit_act = QAction("Exit", self); exit_act.triggered.connect(self.close)
        # file_menu.addAction(cancel_act); file_menu.addSeparator(); file_menu.addAction(exit_act)

        # help_menu = mb.addMenu("&Help")
        # help_text = QAction("Formula Help", self)
        # help_text.triggered.connect(lambda: self.statusBar().showMessage(
        #     "Functions: SUM, AVERAGE/AVG, MIN, MAX, COUNT, PRODUCT, ABS. Use A1 or A1:B3; nesting allowed.",
        #     8000))
        # help_menu.addAction(help_text)

        self.statusBar().showMessage("Ready — Ctrl+Click to pan • Wheel to zoom")
        self.scene.cellSelected.connect(self._on_cell_selected)
        self._current_item = None

    def _cancel_edit(self):
        self.scene.cancel_editing()

    def _on_cell_selected(self, hex_item):
        if hex_item is None:
            return
        self._current_item = hex_item
        addr = f"{column_to_excel_name(hex_item.col)}{hex_item.row+1}"
        self.name_label.setText(addr)
        self.formula_edit.setText(hex_item.value)

    def _apply_formula_edit(self):
        if self._current_item is None:
            return
        self._current_item.set_value(self.formula_edit.text())
        self.scene.recalc_all()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_F11:
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
