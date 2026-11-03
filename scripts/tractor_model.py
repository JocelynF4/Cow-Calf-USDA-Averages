"""Two-wheel-drive (2WD) tractor variable (operating) cost-per-hour model, shared by the standalone workbook
and the combined tool.

Only variable costs are measured — costs that rise with each hour the tractor runs.
Ownership costs (depreciation, interest, taxes, insurance, housing) are left out.

  fuel (gal/hr)          = 0.044 x max PTO horsepower      (average diesel use;
                           Iowa State Ag Decision Maker A3-29)
  fuel cost              = gal/hr x regional diesel price
  lubrication            = entered separately in $/hr (oil, filters, grease) — not
                           tied to the diesel price
  repairs & maintenance  = list price x RF1 x (hours/1000)^RF2   (ASABE D497; accumulated,
                           so this year's cost is the change across this year's hours)

Repairs per hour = this year's repair cost / hours used per year.
"""
import csv
import os

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

FUEL_GAL_PER_HP_HR = 0.044

# The model covers two-wheel-drive (2WD) tractors only.
# Repair factors and estimated life for 2WD tractors (ASABE D497.7, table 3).
DRIVE_TYPES = [
    # label, RF1, RF2, estimated life (hours)
    ("2WD", 0.007, 2.0, 12000),
]

# Diesel price basis -> column in Diesel_by_Region.
BASES = [("5-year average", 3), ("52-week average", 4), ("Latest week", 2), ("2025 average", 5)]

DEFAULTS = {
    "hp": 100, "list_price": 120000, "hours_on": 3000,
    "hours_yr": 400, "lube": 3.00,
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
    heads = ["Price Series", "Latest Week ($/gal)", "5-Year Avg ($/gal)", "52-Week Avg ($/gal)",
             "2025 Avg ($/gal)", "Latest Week Date", "5-Year Window Starts",
             "52-Week Window Starts"]
    for j, h in enumerate(heads, start=1):
        dr.cell(1, j, h)
        _hdr(dr.cell(1, j))
    for i, r in enumerate(regions, start=2):
        vals = [r["region"], float(r["latest_price"]), float(r["avg_5yr"]), float(r["avg_52wk"]),
                float(r["avg_2025"]), r["latest_week"], r["avg_5yr_from"], r["avg_52wk_from"]]
        for j, v in enumerate(vals, start=1):
            c = dr.cell(i, j, v)
            if 2 <= j <= 5:
                c.number_format = '"$"0.000'
    dr_last = 1 + len(regions)
    for col, w in zip("ABCDEFGH", (28, 16, 16, 16, 16, 16, 18, 18)):
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
    k = drive_last + 2
    tf.cell(k, 1, "Fuel use (gal per hp-hour)")
    tf.cell(k, 2, FUEL_GAL_PER_HP_HR)
    tf.cell(k + 2, 1, "Source: 2WD tractor repair factors from ASABE D497.7 (table 3); fuel use from "
                      "Iowa State Ag Decision Maker A3-29. Estimated life is for reference "
                      "only.").font = NOTE
    tf.column_dimensions["A"].width = 30
    tf.column_dimensions["B"].width = 20
    for col in "CD":
        tf.column_dimensions[col].width = 18

    tf.cell(1, 7, "Diesel Price Basis")
    tf.cell(1, 8, "Column")
    _hdr(tf.cell(1, 7))
    _hdr(tf.cell(1, 8))
    for i, (lab, col) in enumerate(BASES, start=2):
        tf.cell(i, 7, lab)
        tf.cell(i, 8, col)
    tf.column_dimensions["G"].width = 20
    wb.defined_names.add(DefinedName("DieselBasis", attr_text=f"Tractor_Factors!$G$2:$G${1 + len(BASES)}"))

    return {
        "dr": f"Diesel_by_Region!$A$2:$H${dr_last}",
        "dr_key": f"Diesel_by_Region!$A$2:$A${dr_last}",
        "sf_key": f"State_Fuel_Region!$A$2:$A${sf_last}",
        "sf_padd": f"State_Fuel_Region!$B$2:$B${sf_last}",
        "sf_name": f"State_Fuel_Region!$C$2:$C${sf_last}",
        "sf_series": f"State_Fuel_Region!$D$2:$D${sf_last}",
        "rf1": "Tractor_Factors!$B$2",
        "rf2": "Tractor_Factors!$C$2",
        "fuel_k": f"Tractor_Factors!$B${k}",
        "basis_key": f"Tractor_Factors!$G$2:$G${1 + len(BASES)}",
        "basis_col": f"Tractor_Factors!$H$2:$H${1 + len(BASES)}",
        "states": [s["state"] for s in sorted(states, key=lambda s: s["state"])],
    }


def write_calculator(ws, top, state_cell, ref, title="  TRACTOR VARIABLE COST PER HOUR"):
    """Write the input + result block into ws starting at row `top` (columns B:D).

    `state_cell` is an absolute reference to the cell holding the full state name.
    Returns the last row used.
    """
    def section(row, text):
        c = ws[f"B{row}"]
        c.value, c.font = text, Font(bold=True, size=12, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = LEFT

    def heading(row, text):
        ws[f"B{row}"].value, ws[f"B{row}"].font = text, Font(bold=True, size=11, color=NAVY)

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

    def out(row, text, formula, fmt=MONEY_HR, strong=False):
        c = ws[f"B{row}"]
        c.value, c.alignment, c.border = text, LEFT, BORDER
        c.font = Font(bold=strong, size=12 if strong else 11, color=NAVY if strong else "000000")
        c.fill = PatternFill("solid", fgColor=GREEN if strong else GREY)
        v = ws[f"C{row}"]
        v.value = formula
        v.font = Font(bold=True, size=14 if strong else 11, color=NAVY)
        v.fill = PatternFill("solid", fgColor=GREEN if strong else LIGHT)
        v.alignment, v.border, v.number_format = CEN, BORDER, fmt
        return f"$C${row}"

    def dropdown(cell_ref, name):
        dv = DataValidation(type="list", formula1=name, allow_blank=False)
        ws.add_data_validation(dv)
        dv.add(ws[cell_ref.replace("$", "")])

    d = DEFAULTS
    st = state_cell
    r = top
    section(r, title)
    r += 1

    heading(r, "Fuel")
    r += 1
    derived(r, "Fuel region:",
            f'=IF({st}="","",IFERROR(INDEX({ref["sf_padd"]},MATCH({st},{ref["sf_key"]},0))'
            f'&" · "&INDEX({ref["sf_name"]},MATCH({st},{ref["sf_key"]},0)),"n/a"))')
    r += 1
    ser = derived(r, "Diesel price series:",
                  f'=IF({st}="","",IFERROR(INDEX({ref["sf_series"]},MATCH({st},{ref["sf_key"]},0)),"n/a"))')
    r += 1
    basis = field(r, "Diesel price basis:", BASES[0][0])
    dropdown(basis, "DieselBasis")
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
    r += 1
    hp = field(r, "Max PTO horsepower:", d["hp"], "0")
    r += 1
    gal = derived(r, "Fuel use:", f'=IFERROR({ref["fuel_k"]}*{hp},"")', '0.0" gal/hr"')
    r += 2

    heading(r, "Lubrication")
    r += 1
    lube_in = field(r, "Lubrication cost ($/hr):", d["lube"], '"$"#,##0.00',
                    "oil, filters & grease per running hour — enter your own")
    r += 2

    heading(r, "Repairs & maintenance")
    r += 1
    derived(r, "Drive type:", "2WD", note="model covers 2WD tractors only")
    r += 1
    lp = field(r, "List price, new equivalent ($):", d["list_price"], '"$"#,##0',
               "today's list price of a comparable new tractor")
    r += 1
    hon = field(r, "Hours already on tractor:", d["hours_on"], "#,##0")
    r += 1
    hyr = field(r, "Hours used per year:", d["hours_yr"], "#,##0")
    r += 2

    heading(r, "Variable costs (per hour)")
    r += 1
    fuel = out(r, "⛽  Fuel", f'=IFERROR({gal}*{price},"n/a")')
    r += 1
    lube = out(r, "🛢️  Lubrication", f'=IF(ISNUMBER({lube_in}),{lube_in},"n/a")')
    r += 1
    repair = out(r, "🔧  Repairs & maintenance (wear & tear)",
                 f'=IFERROR({lp}*{ref["rf1"]}*((({hon}+{hyr})/1000)^{ref["rf2"]}'
                 f'-({hon}/1000)^{ref["rf2"]})/{hyr},"n/a")')
    r += 1
    total = out(r, "💰  Total variable cost per hour",
                f'=IFERROR({fuel}+{lube}+{repair},"n/a")', strong=True)
    r += 1
    out(r, "📅  Total variable cost per year", f'=IFERROR({total}*{hyr},"n/a")', '"$"#,##0" /yr"')
    r += 2

    notes = [
        "2WD tractors only. Variable costs only — depreciation, interest, taxes, insurance, housing and labor are not included.",
        "Fuel = 0.044 gal per hp-hour × horsepower × diesel price (Iowa State Ag Decision Maker A3-29).",
        "Lubrication is entered separately and is not tied to the diesel price.",
        "Repairs = this year's share of list price × repair factor 1 × (hours/1000)^repair factor 2 (ASABE D497).",
        "Diesel is on-highway retail and includes road taxes; off-road farm diesel is usually cheaper — "
        "enter your own price to use it. Territories have no regional series.",
    ]
    for line in notes:
        ws[f"B{r}"].value, ws[f"B{r}"].font = line, NOTE
        ws.merge_cells(f"B{r}:F{r}")
        r += 1
    return r - 1
