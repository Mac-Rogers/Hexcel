import tkinter as tk
from tkinter.simpledialog import askstring
import math

CELL_RADIUS = 32
HEX_ROWS = 8      # Number of rows in grid
HEX_COLS = 10     # Number of cols in grid

class HexCell:
    def __init__(self, polygon, label, value=""):
        self.polygon = polygon
        self.label = label
        self.value = value

class HexcelApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Hexcel: Hexagonal Excel')
        self.geometry('950x640')
        self.canvas = tk.Canvas(self, width=920, height=620, bg="white")
        self.canvas.pack()
        self.grid_cells = {}      # (q, r) : HexCell
        self.selected_cell = None
        self._draw_grid()
        self.canvas.bind("<Button-1>", self.on_click)
        self.canvas.bind("<Double-Button-1>", self.on_dbl_click)

    def hex_center(self, q, r):
        # Pointy-top hex math for neat honeycomb tiling
        x = CELL_RADIUS * 3/2 * q + 80
        y = CELL_RADIUS * math.sqrt(3) * (r + 0.5 * (q % 2)) + 80
        return x, y

    def draw_one_hex(self, cx, cy):
        points = []
        for i in range(6):
            angle_rad = math.pi / 180 * (60 * i - 30)
            points.append(cx + CELL_RADIUS * math.cos(angle_rad))
            points.append(cy + CELL_RADIUS * math.sin(angle_rad))
        return points

    def _draw_grid(self):
        # Draw column headers
        for q in range(HEX_COLS):
            cx, cy = self.hex_center(q, -0.9)
            self.canvas.create_text(cx, cy, text=chr(q+65), font=('Arial', 13, 'bold'), fill="#222")
        # Draw hex cells
        for q in range(HEX_COLS):
            for r in range(HEX_ROWS):
                cx, cy = self.hex_center(q, r)
                pts = self.draw_one_hex(cx, cy)
                poly = self.canvas.create_polygon(pts, outline="#bbb", fill="#f8faff", width=2)
                label = self.canvas.create_text(cx, cy, text="", font=('Arial', 12, 'bold'), fill="#222")
                self.grid_cells[(q, r)] = HexCell(polygon=poly, label=label)
    
    def find_cell(self, x, y):
        # Geometric hit-test for cell under mouse
        for (q, r), cell in self.grid_cells.items():
            cx, cy = self.hex_center(q, r)
            dx, dy = abs(x-cx), abs(y-cy)
            # Hexes fit in a circle for click testing
            if dx < CELL_RADIUS * 0.9 and dy < CELL_RADIUS * 0.9:
                return (q, r)
        return None

    def on_click(self, event):
        cell_pos = self.find_cell(event.x, event.y)
        # Deselect all
        for c in self.grid_cells.values():
            self.canvas.itemconfig(c.polygon, outline="#bbb", fill="#f8faff", width=2)
        if cell_pos:
            c = self.grid_cells[cell_pos]
            self.canvas.itemconfig(c.polygon, outline="#0366d6", fill="#ffe19d", width=3)
            self.selected_cell = cell_pos

    def on_dbl_click(self, event):
        cell_pos = self.find_cell(event.x, event.y)
        if cell_pos:
            c = self.grid_cells[cell_pos]
            value = askstring("Edit Cell", "Enter value:", initialvalue=c.value)
            if value is not None:
                c.value = value
                self.canvas.itemconfig(c.label, text=value)
                self.canvas.itemconfig(c.polygon, fill="#ffe19d")
                self.selected_cell = cell_pos

if __name__ == "__main__":
    HexcelApp().mainloop()
