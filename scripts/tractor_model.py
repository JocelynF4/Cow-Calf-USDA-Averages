"""Tractor cost-per-hour model, shared by the standalone workbook and the combined tool.

The method follows the American Society of Agricultural and Biological Engineers
machinery-cost standard (ASABE D497 / EP496), the same method used by Iowa State
University's "Estimating Farm Machinery Costs" (Ag Decision Maker A3-29):

  fuel (gal/hr)          = 0.044 x max PTO horsepower            (average diesel use)
  lubrication            = 15% of fuel cost
  repairs & maintenance  = list price x RF1 x (hours/1000)^RF2   (accumulated; the cost
                           for this year is the change across this year's hours)
  remaining value        = list price x (C1 - C2 x sqrt(age) - C3 x sqrt(hours/yr))^2
  depreciation           = this year's drop in remaining value
  interest               = average value this year x interest rate
  taxes/insurance/housing= average value this year x that percentage

Every per-hour figure = the year's cost / hours used per year.
"""
import csv
import os

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

FUEL_GAL_PER_HP_HR = 0.044
LUBE_SHARE_OF_FUEL = 0.15

# Repair factors and estimated life by drive type (ASABE D497.7, table 3).
DRIVE_TYPES = [
    # label, RF1, RF2, estimated life (hours)
    ("2WD", 0.007, 2.0, 12000),
    ("4WD / track", 0.003, 2.0, 16000),
]

# Remaining-value coefficients by tractor size (ASABE D497.7, table 6).
SIZE_CLASSES = [
    # lower horsepower bound, label, C1, C2, C3
    (0, "Small (under 80 hp)", 0.981, 0.093, 0.0058),
    (80, "Medium (80-150 hp)", 0.942, 0.100, 0.0008),
    (150, "Large (over 150 hp)", 0.976, 0.119, 0.0019),
]

BASES = [("52-week average", 3), ("Latest week", 2), ("2025 average", 4)]

DEFAULTS = {
    "hp": 100, "drive": "2WD", "list_price": 120000, "hours_on": 3000,
    "hours_yr": 400, "age": 5, "interest": 0.06, "tih": 0.01,
}

NAVY, BLUE, LIGHT, AMBER, GREY, GREEN = "1F3A5F", "2F5C8F", "DCE6F1", "FFF2CC", "F2F2F2", "E2EFDA"
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR = Font(bold=True, color="FFFFFF", size=10)
CEN = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")
NOTE = Font(italic=True, size=9, color="808080")
MONEY_HR = '"$"#,##0.00" /hr"'


def load_diesel():
    with open(os.path.join(ROOT, "data", "diesel_by_region.csv")) as fh:
        regions = list(csv.DictReader(fh))
    with open(os.path.join(ROOT, "data", "state_diesel_region.csv")) as fh:
        states = list(csv.DictReader(fh))
    return regions, states


def _hdr(cell):
    cell.font, cell.alignment, cell.border = HDR, CEN, BORDER
    cell.fill = PatternFill("solid", fgColor=BLUE)


def write_data_sheets(wb, states_to_include=None):
    """Add Diesel_by_Region, State_Fuel_Region and Tractor_Factors; return lookup ranges."""
    regions, states = load_diesel()
    if states_to_include is not None:
        keep = set(states_to_include)
        states = [s for s in states if s["state"] in keep]

    dr = wb.create_sheet("Diesel_by_Region")
    heads = ["Price Series", "Latest Week ($/gal)", "52-Week Avg ($/gal)", "2025 Avg ($/gal)",
             "Latest Week Date", "52-Week Window Starts"]
    for j, h in enumerate(heads, start=1):
        dr.cell(1, j, h)
        _hdr(dr.cell(1, j))
    for i, r in enumerate(regions, start=2):
        vals = [r["region"], float(r["latest_price"]), float(r["avg_52wk"]),
                float(r["avg_2025"]), r["latest_week"], r["avg_52wk_from"]]
        for j, v in enumerate(vals, start=1):
            c = dr.cell(i, j, v)
            if 2 <= j <= 4:
                c.number_format = '"$"0.000'
    dr_last = 1 + len(regions)
    for col, w in zip("ABCDEF", (28, 16, 16, 16, 16, 18)):
        dr.column_dimensions[col].width = w

    sf = wb.create_sheet("State_Fuel_Region")
    for j, h in enumerate(["State", "PADD", "Fuel Region", "Price Series Used"], start=1):
        sf.cell(1, j, h)
        _hdr(sf.cell(1, j))
    for i, r in enumerate(sorted(states, key=lambda s: s["state"]), start=2):
        sf.cell(i, 1, r["state"])
        sf.cell(i, 2, r["padd"])
        sf.cell(i, 3, r["padd_name"])
        sf.cell(i, 4, r["price_series"] or "none published")
    sf_last = 1 + len(states)
    for col, w in zip("ABCD", (26, 10, 22, 28)):
        sf.column_dimensions[col].width = w

    tf = wb.create_sheet("Tractor_Factors")
    for j, h in enumerate(["Drive Type", "Repair Factor 1", "Repair Factor 2",
                           "Estimated Life (hrs)"], start=1):
        tf.cell(1, j, h)
        _hdr(tf.cell(1, j))
    for i, row in enumerate(DRIVE_TYPES, start=2):
        for j, v in enumerate(row, start=1):
            tf.cell(i, j, v)
    drive_last = 1 + len(DRIVE_TYPES)
    s0 = drive_last + 2
    for j, h in enumerate(["Min Horsepower", "Size Class", "Value Coeff 1",
                           "Value Coeff 2", "Value Coeff 3"], start=1):
        tf.cell(s0, j, h)
        _hdr(tf.cell(s0, j))
    for i, row in enumerate(SIZE_CLASSES, start=s0 + 1):
        for j, v in enumerate(row, start=1):
            tf.cell(i, j, v)
    size_first, size_last = s0 + 1, s0 + len(SIZE_CLASSES)
    k = size_last + 2
    tf.cell(k, 1, "Fuel use (gal per hp-hour)")
    tf.cell(k, 2, FUEL_GAL_PER_HP_HR)
    tf.cell(k + 1, 1, "Lubrication (share of fuel cost)")
    tf.cell(k + 1, 2, LUBE_SHARE_OF_FUEL)
    tf.cell(k + 1, 2).number_format = "0%"
    tf.cell(k + 3, 1, "Source: ASABE D497.7 (repair factors table 3, remaining-value "
                      "coefficients table 6); fuel and lubrication per ASABE EP496 / "
                      "Iowa State Ag Decision Maker A3-29.").font = NOTE
    tf.column_dimensions["A"].width = 30
    tf.column_dimensions["B"].width = 20
    for col in "CDE":
        tf.column_dimensions[col].width = 16

    wb.defined_names.add(DefinedName("DriveTypes", attr_text=f"Tractor_Factors!$A$2:$A${drive_last}"))
    wb.defined_names.add(DefinedName("DieselBasis", attr_text=f"Tractor_Factors!$G$2:$G${1 + len(BASES)}"))
    tf.cell(1, 7, "Diesel Price Basis")
    tf.cell(1, 8, "Column")
    _hdr(tf.cell(1, 7))
    _hdr(tf.cell(1, 8))
    for i, (lab, col) in enumerate(BASES, start=2):
        tf.cell(i, 7, lab)
        tf.cell(i, 8, col)
    tf.column_dimensions["G"].width = 20

    return {
        "dr": f"Diesel_by_Region!$A$2:$F${dr_last}",
        "dr_key": f"Diesel_by_Region!$A$2:$A${dr_last}",
        "sf_key": f"State_Fuel_Region!$A$2:$A${sf_last}",
        "sf_padd": f"State_Fuel_Region!$B$2:$B${sf_last}",
        "sf_name": f"State_Fuel_Region!$C$2:$C${sf_last}",
        "sf_series": f"State_Fuel_Region!$D$2:$D${sf_last}",
        "drive_key": f"Tractor_Factors!$A$2:$A${drive_last}",
        "rf1": f"Tractor_Factors!$B$2:$B${drive_last}",
        "rf2": f"Tractor_Factors!$C$2:$C${drive_last}",
        "size_min": f"Tractor_Factors!$A${size_first}:$A${size_last}",
        "size_lab": f"Tractor_Factors!$B${size_first}:$B${size_last}",
        "c1": f"Tractor_Factors!$C${size_first}:$C${size_last}",
        "c2": f"Tractor_Factors!$D${size_first}:$D${size_last}",
        "c3": f"Tractor_Factors!$E${size_first}:$E${size_last}",
        "fuel_k": f"Tractor_Factors!$B${k}",
        "lube_k": f"Tractor_Factors!$B${k + 1}",
        "basis_key": f"Tractor_Factors!$G$2:$G${1 + len(BASES)}",
        "basis_col": f"Tractor_Factors!$H$2:$H${1 + len(BASES)}",
        "states": [s["state"] for s in sorted(states, key=lambda s: s["state"])],
    }


def write_calculator(ws, top, state_cell, ref, title="  TRACTOR COST PER HOUR"):
    """Write the input + result block into ws starting at row `top` (columns B:D).

    `state_cell` is an absolute reference to the cell holding the full state name.
    Returns the last row used.
    """
    def section(row, text):
        c = ws[f"B{row}"]
        c.value, c.font = text, Font(bold=True, size=12, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = LEFT

    def field(row, text, value=None, fmt=None, note=None):
        c = ws[f"B{row}"]
        c.value, c.font = text, Font(bold=True, size=11, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=BLUE)
        c.alignment, c.border = RIGHT, BORDER
        v = ws[f"C{row}"]
        v.value = value
        v.fill = PatternFill("solid", fgColor=AMBER)
        v.alignment, v.border, v.font = LEFT, BORDER, Font(bold=True, size=11)
        if fmt:
            v.number_format = fmt
        if note:
            ws[f"D{row}"].value, ws[f"D{row}"].font = note, NOTE
        return f"$C${row}"

    def derived(row, text, formula, fmt=None, note=None):
        c = ws[f"B{row}"]
        c.value, c.font, c.alignment = text, Font(italic=True, size=10, color="595959"), RIGHT
        v = ws[f"C{row}"]
        v.value, v.font, v.alignment = formula, Font(italic=True, size=10, color="595959"), CEN
        if fmt:
            v.number_format = fmt
        if note:
            ws[f"D{row}"].value, ws[f"D{row}"].font = note, NOTE
        return f"$C${row}"

    def out(row, text, formula, fmt=MONEY_HR, strong=False, fill=LIGHT):
        c = ws[f"B{row}"]
        c.value, c.alignment, c.border = text, LEFT, BORDER
        c.font = Font(bold=strong, size=12 if strong else 11, color=NAVY if strong else "000000")
        c.fill = PatternFill("solid", fgColor=GREEN if strong else GREY)
        v = ws[f"C{row}"]
        v.value = formula
        v.font = Font(bold=True, size=14 if strong else 11, color=NAVY)
        v.fill = PatternFill("solid", fgColor=GREEN if strong else fill)
        v.alignment, v.border, v.number_format = CEN, BORDER, fmt
        return f"$C${row}"

    d = DEFAULTS
    r = top
    section(r, title)
    r += 1
    st = state_cell
    series = derived(r, "Fuel region:",
                     f'=IF({st}="","",IFERROR(INDEX({ref["sf_padd"]},MATCH({st},{ref["sf_key"]},0))'
                     f'&" · "&INDEX({ref["sf_name"]},MATCH({st},{ref["sf_key"]},0)),"n/a"))')
    r += 1
    ser = derived(r, "Diesel price series:",
                  f'=IF({st}="","",IFERROR(INDEX({ref["sf_series"]},MATCH({st},{ref["sf_key"]},0)),"n/a"))')
    r += 1
    basis = field(r, "Diesel price basis:", BASES[0][0])
    dv_basis = DataValidation(type="list", formula1="DieselBasis", allow_blank=False)
    ws.add_data_validation(dv_basis)
    dv_basis.add(ws[basis.replace("$", "")])
    r += 1
    regional = derived(
        r, "Regional diesel price:",
        f'=IFERROR(INDEX({ref["dr"]},MATCH({ser},{ref["dr_key"]},0),'
        f'INDEX({ref["basis_col"]},MATCH({basis},{ref["basis_key"]},0))),"n/a")',
        '"$"0.000" /gal"', "on-highway retail price, includes road taxes")
    r += 1
    own = field(r, "Your own diesel price (optional):", None, '"$"0.000" /gal"',
                "leave blank to use the regional price")
    r += 1
    price = derived(r, "Diesel price used:", f'=IF(ISNUMBER({own}),{own},{regional})',
                    '"$"0.000" /gal"')
    r += 2

    hp = field(r, "Max PTO horsepower:", d["hp"], "0")
    r += 1
    drive = field(r, "Drive type:", d["drive"])
    dv_drive = DataValidation(type="list", formula1="DriveTypes", allow_blank=False)
    ws.add_data_validation(dv_drive)
    dv_drive.add(ws[drive.replace("$", "")])
    r += 1
    lp = field(r, "List price, new equivalent ($):", d["list_price"], '"$"#,##0',
               "today's list price of a comparable new tractor")
    r += 1
    hon = field(r, "Hours already on tractor:", d["hours_on"], "#,##0")
    r += 1
    hyr = field(r, "Hours used per year:", d["hours_yr"], "#,##0")
    r += 1
    age = field(r, "Tractor age (years):", d["age"], "0")
    r += 1
    rate = field(r, "Interest rate:", d["interest"], "0.0%")
    r += 1
    tih = field(r, "Taxes, insurance & housing:", d["tih"], "0.0%", "% of tractor value per year")
    r += 1
    size = derived(r, "Size class:",
                   f'=IFERROR(INDEX({ref["size_lab"]},MATCH({hp},{ref["size_min"]},1)),"")')
    r += 2

    def size_k(rng):
        return f'INDEX({ref[rng]},MATCH({hp},{ref["size_min"]},1))'

    def drive_k(rng):
        return f'INDEX({ref[rng]},MATCH({drive},{ref["drive_key"]},0))'

    def value_at(age_expr):
        return (f'{lp}*MAX(0,{size_k("c1")}-{size_k("c2")}*SQRT({age_expr})'
                f'-{size_k("c3")}*SQRT({hyr}))^2')

    v0 = derived(r, "Value at start of this year:", f'=IFERROR({value_at(age)},"")', '"$"#,##0')
    r += 1
    v1 = derived(r, "Value at end of this year:", f'=IFERROR({value_at(age + "+1")},"")', '"$"#,##0')
    r += 1
    gal = derived(r, "Fuel use:", f'=IFERROR({ref["fuel_k"]}*{hp},"")', '0.0" gal/hr"')
    r += 2

    ws[f"B{r}"].value, ws[f"B{r}"].font = "Operating costs (per hour)", Font(bold=True, size=11, color=NAVY)
    r += 1
    fuel = out(r, "⛽  Fuel", f'=IFERROR({gal}*{price},"n/a")')
    r += 1
    lube = out(r, "🛢️  Lubrication", f'=IFERROR({ref["lube_k"]}*{fuel},"n/a")')
    r += 1
    repair = out(r, "🔧  Repairs & maintenance",
                 f'=IFERROR({lp}*{drive_k("rf1")}*((({hon}+{hyr})/1000)^{drive_k("rf2")}'
                 f'-({hon}/1000)^{drive_k("rf2")})/{hyr},"n/a")')
    r += 1
    opsub = out(r, "Operating subtotal", f'=IFERROR({fuel}+{lube}+{repair},"n/a")')
    ws[f"B{r}"].font = Font(bold=True, size=11)
    r += 2

    ws[f"B{r}"].value, ws[f"B{r}"].font = "Ownership costs (per hour)", Font(bold=True, size=11, color=NAVY)
    r += 1
    dep = out(r, "📉  Depreciation", f'=IFERROR(({v0}-{v1})/{hyr},"n/a")')
    r += 1
    intr = out(r, "🏦  Interest on investment", f'=IFERROR(({v0}+{v1})/2*{rate}/{hyr},"n/a")')
    r += 1
    tihc = out(r, "🏠  Taxes, insurance & housing", f'=IFERROR(({v0}+{v1})/2*{tih}/{hyr},"n/a")')
    r += 1
    owsub = out(r, "Ownership subtotal", f'=IFERROR({dep}+{intr}+{tihc},"n/a")')
    ws[f"B{r}"].font = Font(bold=True, size=11)
    r += 2

    out(r, "🔩  Wear & tear (repairs + depreciation)", f'=IFERROR({repair}+{dep},"n/a")',
        strong=True)
    r += 1
    total = out(r, "💰  Total cost per hour", f'=IFERROR({opsub}+{owsub},"n/a")', strong=True)
    r += 1
    out(r, "📅  Total cost per year", f'=IFERROR({total}*{hyr},"n/a")', '"$"#,##0" /yr"')
    r += 2

    notes = [
        "Method: ASABE machinery-cost standard (D497 / EP496), as used in Iowa State Ag Decision Maker A3-29.",
        "Fuel = 0.044 gal per hp-hour × horsepower × diesel price; lubrication = 15% of fuel.",
        "Repairs = this year's share of list price × repair factor 1 × (hours/1000)^repair factor 2.",
        "Depreciation = this year's drop in the standard's remaining-value curve for the size class.",
        "Diesel is on-highway retail and includes road taxes; off-road farm diesel is usually cheaper — "
        "enter your own price to use it. Territories have no regional series.",
        "Labor is not included.",
    ]
    for line in notes:
        ws[f"B{r}"].value, ws[f"B{r}"].font = line, NOTE
        ws.merge_cells(f"B{r}:F{r}")
        r += 1
    return r - 1
