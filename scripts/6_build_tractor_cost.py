"""Build the standalone tractor cost-per-hour workbook -> Tractor Cost per Hour.xlsx."""
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

import tractor_model as tm

OUT = os.path.join(tm.ROOT, "Tractor Cost per Hour.xlsx")

wb = Workbook()
ws = wb.active
ws.title = "Tractor Cost"
ref = tm.write_data_sheets(wb)

ws.sheet_view.showGridLines = False
ws["B2"] = "Tractor Cost per Hour"
ws["B2"].font = Font(bold=True, size=20, color=tm.NAVY)
ws["B3"] = ("Pick a state to pull its regional diesel price, then describe the tractor. "
            "Amber cells are inputs.")
ws["B3"].font = Font(italic=True, size=10, color="595959")
ws.merge_cells("B3:F3")

c = ws["B5"]
c.value, c.font = "  1 · LOCATION", Font(bold=True, size=12, color="FFFFFF")
c.fill = PatternFill("solid", fgColor=tm.NAVY)
c.alignment = tm.LEFT
b = ws["B6"]
b.value, b.font = "State:", Font(bold=True, size=11, color="FFFFFF")
b.fill, b.alignment, b.border = PatternFill("solid", fgColor=tm.BLUE), tm.RIGHT, tm.BORDER
s = ws["C6"]
s.value = "Iowa"
s.fill = PatternFill("solid", fgColor=tm.AMBER)
s.alignment, s.border, s.font = tm.LEFT, tm.BORDER, Font(bold=True, size=11)

wb.defined_names.add(DefinedName("FuelStates", attr_text=ref["sf_key"]))
dv = DataValidation(type="list", formula1="FuelStates", allow_blank=True)
ws.add_data_validation(dv)
dv.add(ws["C6"])

tm.write_calculator(ws, 8, "$C$6", ref, title="  2 · TRACTOR & COST PER HOUR")

ws.column_dimensions["A"].width = 2
ws.column_dimensions["B"].width = 38
ws.column_dimensions["C"].width = 22
ws.column_dimensions["D"].width = 40
ws.freeze_panes = "A7"

wb.save(OUT)
print("saved:", OUT)
