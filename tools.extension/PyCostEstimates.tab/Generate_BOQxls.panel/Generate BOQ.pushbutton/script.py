#! python3
import os
import string
import xlsxwriter
import clr

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import MessageBox
from pyrevit import revit, DB

# --- Save path ---
desktop = os.path.expanduser("~/Desktop")
xlsx_path = os.path.join(desktop, "BOQ_Export_From_Model.xlsx")

# --- Parameters ---
PARAM_COST = "Cost"
PARAM_TOTAL = "Test_1234"

CATEGORY_MAP = {
    "Walls": DB.BuiltInCategory.OST_Walls,
    "Doors": DB.BuiltInCategory.OST_Doors,
    "Windows": DB.BuiltInCategory.OST_Windows,
    "Structural Foundations": DB.BuiltInCategory.OST_StructuralFoundation,
    "Structural Framing": DB.BuiltInCategory.OST_StructuralFraming,
    "Structural Columns": DB.BuiltInCategory.OST_StructuralColumns,
    "Roofs": DB.BuiltInCategory.OST_Roofs,
    "Ceilings": DB.BuiltInCategory.OST_Ceilings,  # ✅ Added
    "Electrical": [
        DB.BuiltInCategory.OST_Conduit,
        DB.BuiltInCategory.OST_LightingFixtures,
        DB.BuiltInCategory.OST_LightingDevices,
        DB.BuiltInCategory.OST_ElectricalFixtures
    ]
}

CATEGORY_DESCRIPTIONS = {
    "Walls": "Blockwork in hollow concrete blocks load bearing walls, plastered both sides and painted, including all mortar and reinforcement ties.",
    "Doors": "Flush doors with hardwood frame, inclusive of ironmongery, architraves, painting, and necessary fixing accessories.",
    "Windows": "Aluminium sliding windows including glazing, window stays, handles, mosquito gauze and fixing to concrete or blockwork reveals.",
    "Structural Foundations": "Mass concrete, RC footing, bedding and hardcore compacted in layers, damp-proof membrane and blinding, including formwork and reinforcement.",
    "Structural Framing": "Mild steel beams and trusses complete with welding, surface preparation, primer coating, and installation.",
    "Structural Columns": "Steel and concrete columns including starter bars, ties, shuttering, and specified concrete class with admixtures.",
    "Roofs": "0.5mm IBR/IT4 Pre-painted roof sheeting fixed to purlins with appropriate screws, complete with ridge capping, barge boards, insulation and accessories.",
    "Ceilings": "Particle board ceilings (to BS EN 312, P3 grade where humid) and PVC tongue-and-groove ceiling panels (fire-retardant to BS EN 13501-1) shall be fixed to galvanized steel or timber framing with corrosion-resistant fasteners, complete with all trims, joint covers, sealants and allowances for movement, delivering a smooth, level, washable finish free of visible defects.",
    "Electrical": "Steel conduits to BS 4568 -1 are available in Class 3 (pre-galvanized, medium to heavy duty) and Class 4 (hot-dip galvanized heavy duty), with diameters typically from 20 to 50 mm, wall thickness of 1.3–1.6 mm, metric threading, and supplied in 3.75 m lengths. Armoured cables to SANS 1507 standard, suitable for indoor and outdoor installation, laid in ducts or clipped directly. IP-rated junction boxes and fittings shall be provided where necessary."
}

# --- Unit conversions ---
FT3_TO_M3 = 0.0283168
FT2_TO_M2 = 0.092903
FT_TO_M = 0.3048
CONCRETE_NAME = "Concrete - Cast-in-Place Concrete"
STEEL_NAME = "Metal - Steel 43-275"

# --- Create workbook ---
wb = xlsxwriter.Workbook(xlsx_path)
sheet = wb.add_worksheet("BOQ Export")
sheet.freeze_panes(1, 0)

# --- Formats ---
font = 'Century Gothic'
def col_fmt(bold=False, italic=False, underline=False, wrap=False, num_fmt=None):
    fmt = {'valign': 'top', 'font_name': font, 'border': 1}
    if bold: fmt['bold'] = True
    if italic: fmt['italic'] = True
    if underline: fmt['underline'] = True
    if wrap: fmt['text_wrap'] = True
    if num_fmt: fmt['num_format'] = num_fmt
    return wb.add_format(fmt)

fmt_header = col_fmt(bold=True)
fmt_bold = col_fmt(bold=True)
fmt_italic_wrap = col_fmt(italic=True, underline=True, wrap=True)
fmt_section = col_fmt(bold=True)
fmt_money = col_fmt(num_fmt='#,##0.00')
fmt_normal = col_fmt()

# --- Headers ---
headers = ["ITEM", "DESCRIPTION", "UNIT", "QTY", "RATE (ZMW)", "AMOUNT (ZMW)"]
for col, header in enumerate(headers):
    sheet.write(0, col, header, fmt_header)

# --- Column widths ---
sheet.set_column(1, 1, 30)
sheet.set_column(4, 4, 12)
sheet.set_column(5, 5, 16)

row = 1
skipped = 0
category_totals = []

for category_name, bic in CATEGORY_MAP.items():
    elements = []
    if isinstance(bic, list):
        for sub_cat in bic:
            elements += DB.FilteredElementCollector(revit.doc).OfCategory(sub_cat).WhereElementIsNotElementType().ToElements()
    else:
        elements = DB.FilteredElementCollector(revit.doc).OfCategory(bic).WhereElementIsNotElementType().ToElements()

    grouped = {}
    item_letter = iter(string.ascii_uppercase)

    for el in elements:
        try:
            el_type = revit.doc.GetElement(el.GetTypeId())
            name = el_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            cost_param = el_type.LookupParameter(PARAM_COST)
            total_param = el.LookupParameter(PARAM_TOTAL)
            if not (cost_param and total_param):
                skipped += 1
                continue

            rate = cost_param.AsDouble()
            total = total_param.AsDouble()
            qty = 1.0
            unit = "No."

            if category_name in ["Doors", "Windows"]:
                qty = 1
            elif category_name in ["Walls", "Roofs", "Ceilings"]:
                area_param = el.LookupParameter("Area")
                if area_param and area_param.HasValue:
                    qty = area_param.AsDouble() * FT2_TO_M2
                    unit = "m²"
            elif category_name == "Structural Foundations":
                vol_param = el.LookupParameter("Volume")
                if vol_param and vol_param.HasValue:
                    qty = vol_param.AsDouble() * FT3_TO_M3
                    unit = "m³"
            elif category_name == "Structural Framing":
                len_param = el.get_Parameter(DB.BuiltInParameter.CURVE_ELEM_LENGTH)
                if len_param and len_param.HasValue:
                    qty = len_param.AsDouble() * FT_TO_M
                    unit = "m"
            elif category_name == "Structural Columns":
                mat_param = el.LookupParameter("Structural Material")
                mat_elem = revit.doc.GetElement(mat_param.AsElementId()) if mat_param else None
                mat_name = mat_elem.Name if mat_elem else ""
                if mat_name == CONCRETE_NAME:
                    vol_param = el.LookupParameter("Volume")
                    if vol_param and vol_param.HasValue:
                        qty = vol_param.AsDouble() * FT3_TO_M3
                        unit = "m³"
                elif mat_name == STEEL_NAME:
                    len_param = el.get_Parameter(DB.BuiltInParameter.CURVE_ELEM_LENGTH)
                    if len_param and len_param.HasValue:
                        qty = len_param.AsDouble() * FT_TO_M
                        unit = "m"
            elif category_name == "Electrical":
                len_param = el.get_Parameter(DB.BuiltInParameter.CURVE_ELEM_LENGTH)
                if len_param and len_param.HasValue:
                    qty = len_param.AsDouble() * FT_TO_M
                    unit = "m"

            if name not in grouped:
                grouped[name] = {"qty": 0.0, "rate": rate, "total": 0.0, "unit": unit}
            grouped[name]["qty"] += qty
            grouped[name]["total"] += total

        except:
            skipped += 1
            continue

    if grouped:
        sheet.write(row, 1, category_name.upper(), fmt_section)
        row += 1
        if category_name in CATEGORY_DESCRIPTIONS:
            sheet.write(row, 1, CATEGORY_DESCRIPTIONS[category_name], fmt_italic_wrap)
            row += 1

        subtotal = 0.0
        for name, data in grouped.items():
            letter = next(item_letter)
            sheet.write(row, 0, letter, fmt_normal)
            sheet.write(row, 1, name, fmt_normal)
            sheet.write(row, 2, data["unit"], fmt_normal)
            sheet.write(row, 3, round(data["qty"], 2), fmt_normal)
            sheet.write(row, 4, round(data["rate"], 2), fmt_money)
            sheet.write(row, 5, round(data["total"], 2), fmt_money)
            subtotal += data["total"]
            row += 1

        sheet.write(row, 1, category_name.upper() + " TO COLLECTION", fmt_bold)
        sheet.write(row, 5, round(subtotal, 2), fmt_money)
        category_totals.append((category_name.upper(), round(subtotal, 2)))
        row += 2

# --- Summary Collection ---
sheet.write(row, 0, "COLLECTION", fmt_bold)
row += 1
for name, total in category_totals:
    sheet.write(row, 0, name, fmt_normal)
    sheet.write(row, 5, total, fmt_money)
    row += 1

# --- Grand Total ---
grand_total = sum([t[1] for t in category_totals])
sheet.write(row, 0, "GRAND TOTAL", fmt_bold)
sheet.write(row, 5, grand_total, fmt_money)

# --- Close workbook ---
wb.close()
MessageBox.Show("BOQ export complete!\nSaved to Desktop:\n{}\nSkipped: {}".format(xlsx_path, skipped), "✅ XLSX Export")
