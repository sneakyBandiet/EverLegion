import json
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.worksheet.filters import AutoFilter

# Color mapping based on your images
COLOR_MAP = {
    "S": "FF0000",    # Red
    "S+": "C00000",   # Dark Red
    "A": "FFFF00",    # Yellow
    "B": "00B0F0",    # Blue
    "C": "00FF00",    # Green
    "D": "92D050",    # Light Green
    "F": "0070C0"     # Dark Blue
}

# Columns to color
COLOR_COLUMNS = ["PvP", "Story/Tower", "Faction Tower", "Bosses"]

# Load JSON data
with open("herotier_full.json", "r") as f:
    data = json.load(f)

wb = Workbook()
ws = wb.active
ws.title = "Hero Tier"

# Write header
headers = list(data[0].keys())

import json
import sys
from openpyxl import Workbook
from openpyxl.styles import PatternFill

def export_json_to_excel(json_path, excel_path):
    with open(json_path, 'r') as f:
        data = json.load(f)

    wb = Workbook()
    ws = wb.active
    ws.title = "Hero Tier List"

    headers = list(data[0].keys())
    ws.append(headers)

    color_map = {
        'S+': '8B0000', # Dark Red
        'S': 'FF0000',  # Red
        'A': 'FFA500',  # Orange
        'B': 'FFFF00',  # Yellow
        'C': '90EE90',  # Light Green
        'D': '00FF00',  # Green
        'F': '00BFFF',  # Blue
        '-': 'FFFFFF',  # White
    }

    for row in data:
        ws.append([row.get(h, '') for h in headers])

    for r in range(2, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            val = str(cell.value)
            if val in color_map:
                cell.fill = PatternFill(start_color=color_map[val], end_color=color_map[val], fill_type="solid")

        # SAVE (overwrite safely)
    tmp_path = excel_path + ".tmp"
    wb.save(tmp_path)              # write to a temp file first
    wb.replace(tmp_path, excel_path)  # atomically replace (overwrites if exists)

if __name__ == "__main__":
    export_json_to_excel("herotier_full.json", "herotier_full.xlsx")
