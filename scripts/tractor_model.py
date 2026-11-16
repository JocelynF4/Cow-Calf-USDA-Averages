"""Two-wheel-drive (2WD) tractor variable (operating) cost-per-hour model, shared by the standalone workbook
and the combined tool.

Only variable costs are measured — costs that rise with each hour the tractor runs.
Depreciation is added as its own line. Lubrication, interest, taxes, insurance,
housing and labor are left out.

  fuel (gal/hr)          = 0.044 x max PTO horsepower      (average diesel use;
                           Iowa State Ag Decision Maker A3-29)
  fuel cost              = gal/hr x the state's AAA 2026 average diesel price
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

# Remaining-value coefficients by tractor size (ASABE D497.7, table 6), used for
# depreciation. Entered from memory of the standard; not yet checked against it.
SIZE_CLASSES = [
    # lower horsepower bound, label, C1, C2, C3
    (0, "Small (under 80 hp)", 0.981, 0.093, 0.0058),
    (80, "Medium (80-150 hp)", 0.942, 0.100, 0.0008),
    (150, "Large (over 150 hp)", 0.976, 0.119, 0.0019),
]

DEFAULTS = {
    "hp": 100, "list_price": 120000, "hours_on": 0,
    "hours_yr": 400, "age": 0,
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
    """State -> AAA 2026 average diesel price ($/gal), from data/diesel_by_state.csv."""
    with open(os.path.join(ROOT, "data", "diesel_by_state.csv")) as fh:
        return [(r["state"], float(r["diesel_2026_avg"])) for r in csv.DictReader(fh)]


def _hdr(cell):
    cell.font, cell.alignment, cell.border = HDR, CEN, BORDER
    cell.fill = PatternFill("solid", fgColor=BLUE)


def write_data_sheets(wb, states_to_include=None):
    """Add Diesel_by_State and Tractor_Factors; return lookup ranges."""
    states = sorted(load_diesel())
    if states_to_include is not None:
        keep = set(states_to_include)
        states = [s for s in states if s[0] in keep]

    ds = wb.create_sheet("Diesel_by_State")
    for j, h in enumerate(["State", "Diesel 2026 Avg ($/gal)"], start=1):
        ds.cell(1, j, h)
        _hdr(ds.cell(1, j))
    for i, (st, price) in enumerate(states, start=2):
        ds.cell(i, 1, st)
        ds.cell(i, 2, price).number_format = '"$"0.000'
    ds_last = 1 + len(states)
    ds.cell(ds_last + 2, 1, "Source: AAA state diesel price averages for 2026 "
                            "(gasprices.aaa.com/state-gas-price-averages).").font = NOTE
    ds.column_dimensions["A"].width = 24
    ds.column_dimensions["B"].width = 24

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
    tf.cell(k + 2, 1, "Source: 2WD tractor repair factors (table 3) and remaining-value coefficients "
                      "(table 6) from ASABE D497.7; fuel use from Iowa State Ag Decision Maker "
                      "A3-29. Entered from memory, not yet checked against the published "
                      "tables. Estimated life is for reference only.").font = NOTE
    tf.column_dimensions["A"].width = 30
    tf.column_dimensions["B"].width = 20
    for col in "CDE":
        tf.column_dimensions[col].width = 18


    return {
        "ds_key": f"Diesel_by_State!$A$2:$A${ds_last}",
        "ds_price": f"Diesel_by_State!$B$2:$B${ds_last}",
        "rf1": "Tractor_Factors!$B$2",
        "rf2": "Tractor_Factors!$C$2",
        "fuel_k": f"Tractor_Factors!$B${k}",
        "size_min": f"Tractor_Factors!$A${size_first}:$A${size_last}",
        "size_lab": f"Tractor_Factors!$B${size_first}:$B${size_last}",
        "c1": f"Tractor_Factors!$C${size_first}:$C${size_last}",
        "c2": f"Tractor_Factors!$D${size_first}:$D${size_last}",
        "c3": f"Tractor_Factors!$E${size_first}:$E${size_last}",
        "states": [s[0] for s in states],
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
    state_price = derived(
        r, "State diesel price (AAA 2026 avg):",
        f'=IF({st}="","",IFERROR(INDEX({ref["ds_price"]},MATCH({st},{ref["ds_key"]},0)),"n/a"))',
        '"$"0.000" /gal"', "on-road retail price, includes road taxes")
    r += 1
    own = field(r, "Your own diesel price (optional):", None, '"$"0.000" /gal"',
                "leave blank to use the state price")
    r += 1
    price = derived(r, "Diesel price used:", f'=IF(ISNUMBER({own}),{own},{state_price})',
                    '"$"0.000" /gal"')
    r += 1
    hp = field(r, "Max PTO horsepower:", d["hp"], "0")
    r += 1
    gal = derived(r, "Fuel use:", f'=IFERROR({ref["fuel_k"]}*{hp},"")', '0.0" gal/hr"')
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

    heading(r, "Depreciation")
    r += 1
    age = field(r, "Tractor age (years):", d["age"], "0")
    r += 1
    derived(r, "Size class:",
            f'=IFERROR(INDEX({ref["size_lab"]},MATCH({hp},{ref["size_min"]},1)),"")')
    r += 1

    def size_k(rng):
        return f'INDEX({ref[rng]},MATCH({hp},{ref["size_min"]},1))'

    def value_at(age_expr):
        return (f'{lp}*MAX(0,{size_k("c1")}-{size_k("c2")}*SQRT({age_expr})'
                f'-{size_k("c3")}*SQRT({hyr}))^2')

    v0 = derived(r, "Value at start of this year:", f'=IFERROR({value_at(age)},"")', '"$"#,##0')
    r += 1
    v1 = derived(r, "Value at end of this year:", f'=IFERROR({value_at(age + "+1")},"")',
                 '"$"#,##0')
    r += 1
    deprec_yr = derived(r, "Depreciation this year:", f'=IFERROR({v0}-{v1},"")', '"$"#,##0',
                        "falls mostly with age, not hours of use")
    r += 2

    heading(r, "Variable costs (per hour)")
    r += 1
    fuel = out(r, "⛽  Fuel", f'=IFERROR({gal}*{price},"n/a")')
    r += 1
    repair = out(r, "🔧  Repairs & maintenance (wear & tear)",
                 f'=IFERROR({lp}*{ref["rf1"]}*((({hon}+{hyr})/1000)^{ref["rf2"]}'
                 f'-({hon}/1000)^{ref["rf2"]})/{hyr},"n/a")')
    r += 1
    var = out(r, "Variable cost subtotal", f'=IFERROR({fuel}+{repair},"n/a")')
    ws[f"B{r}"].font = Font(bold=True, size=11)
    r += 2

    heading(r, "Depreciation (per hour)")
    r += 1
    dep = out(r, "📉  Depreciation", f'=IFERROR({deprec_yr}/{hyr},"n/a")')
    r += 2

    out(r, "🔩  Wear & tear (repairs + depreciation)", f'=IFERROR({repair}+{dep},"n/a")',
        strong=True)
    r += 1
    total = out(r, "💰  Total cost per hour (variable + depreciation)",
                f'=IFERROR({var}+{dep},"n/a")', strong=True)
    r += 1
    out(r, "📅  Total cost per year", f'=IFERROR({total}*{hyr},"n/a")', '"$"#,##0" /yr"')
    r += 2

    notes = [
        "2WD tractors only. Includes fuel, repairs and depreciation; lubrication, interest, taxes, insurance, housing and labor are not included.",
        "Depreciation = this year's drop in the remaining-value curve for the size class (ASABE D497), ÷ hours used per year.",
        "Fuel = 0.044 gal per hp-hour × horsepower × diesel price (Iowa State Ag Decision Maker A3-29).",
        "Repairs = this year's share of list price × repair factor 1 × (hours/1000)^repair factor 2 (ASABE D497).",
        "Diesel = AAA 2026 state average (on-road retail, includes road taxes); off-road farm diesel is "
        "usually cheaper — enter your own price to use it.",
    ]
    for line in notes:
        ws[f"B{r}"].value, ws[f"B{r}"].font = line, NOTE
        ws.merge_cells(f"B{r}:F{r}")
        r += 1
    return r - 1
