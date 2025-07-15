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
    "Block Work in Walls": DB.BuiltInCategory.OST_Walls,
    "Doors": DB.BuiltInCategory.OST_Doors,
    "Windows": DB.BuiltInCategory.OST_Windows,
    "Structural Foundations": DB.BuiltInCategory.OST_StructuralFoundation,
    "Structural Framing": DB.BuiltInCategory.OST_StructuralFraming,
    "Structural Columns": DB.BuiltInCategory.OST_StructuralColumns,
    "Structural Rebar": DB.BuiltInCategory.OST_Rebar,
    "Roofs": DB.BuiltInCategory.OST_Roofs,
    "Ceilings": DB.BuiltInCategory.OST_Ceilings,
    "Wall and Floor Finishes": DB.BuiltInCategory.OST_GenericModel,
    "Plumbing": [
        DB.BuiltInCategory.OST_PlumbingFixtures,
        DB.BuiltInCategory.OST_PipeCurves,
        DB.BuiltInCategory.OST_PipeFitting
    ],
    "Electrical": [
        DB.BuiltInCategory.OST_Conduit,
        DB.BuiltInCategory.OST_LightingFixtures,
        DB.BuiltInCategory.OST_LightingDevices,
        DB.BuiltInCategory.OST_ElectricalFixtures,
        DB.BuiltInCategory.OST_ElectricalEquipment
    ]
}

CATEGORY_DESCRIPTIONS = {
    "Block Work in Walls": "Blockwork in hollow concrete blocks load bearing walls, plastered both sides and painted, including all mortar and reinforcement ties.",
    "Doors": "Flush doors with hardwood frame, inclusive of ironmongery, architraves, painting, and necessary fixing accessories.",
    "Windows": "Aluminium sliding windows including glazing, window stays, handles, mosquito gauze and fixing to concrete or blockwork reveals.",
    "Structural Foundations": "Mass concrete, RC footing, bedding and hardcore compacted in layers, damp‑proof membrane and blinding, including formwork and reinforcement.",
    "Structural Framing": "Mild steel beams and trusses complete with welding, surface preparation, primer coating, and installation...",
    "Structural Columns": "Steel and concrete columns including starter bars, ties, shuttering, and specified concrete class with admixtures.",
    "Structural Rebar": (
        "High‑yield deformed steel bars to BS 4449:2005 Grade B500B, supplied in standard lengths and bent to shape as scheduled, "
        "including all cutting, bending, fixing, tying with 16‑gauge annealed wire, and providing necessary spacers, chairs, laps, and hooks. "
        "Bars shall be clean, free from loose rust, oil, or paint, and shall be fixed as per BS 8666. "
        "All reinforcement shall be placed accurately to the specified cover, securely supported during concreting, and handled to prevent displacement."
    ),
    "Roofs": "0.5mm IBR/IT4 pre‑painted roof sheeting fixed to purlins with appropriate screws, complete with ridge capping, barge boards, insulation and accessories.",
    "Ceilings": "Particle board ceilings (to BS EN 312...) and PVC tongue‑and‑groove ceiling panels...",
    "Wall and Floor Finishes": "British Standards for wall and floor finishes—primarily BS 5385, BS 8203, and BS 5325—establish best practices...",
    "Plumbing": (
        "Sanitary appliances including WC pans, flush tanks, wash basins, sinks and urinals, complete with all fixings, traps and connections; "
        "and cold/hot water pipework inclusive of fittings, jointing and testing."
    ),
    "Electrical": "Steel conduits to BS 4568‑1, armoured cables to SANS 1507, IP‑rated junction boxes and fittings..."
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

# --- Excel Formats ---
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
fmt_italic = col_fmt(italic=True, wrap=True)
fmt_description = col_fmt(italic=True, underline=True, wrap=True)
fmt_section = col_fmt(bold=True)
fmt_money = col_fmt(num_fmt='#,##0.00')
fmt_normal = col_fmt()

# --- Column Titles & Widths ---
headers = ["ITEM", "DESCRIPTION", "UNIT", "QTY", "RATE (ZMW)", "AMOUNT (ZMW)"]
for col, header in enumerate(headers):
    sheet.write(0, col, header, fmt_header)
sheet.set_column(1, 1, 45)
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

            # Quantity logic
            if category_name in ["Doors", "Windows"]:
                qty = 1

            elif category_name in ["Wall and Floor Finishes", "Roofs", "Ceilings"]:
                area = el.LookupParameter("Area")
                if area and area.HasValue:
                    qty = area.AsDouble() * FT2_TO_M2
                    unit = "m²"

            elif category_name == "Structural Foundations":
                vol = el.LookupParameter("Volume")
                if vol and vol.HasValue:
                    qty = vol.AsDouble() * FT3_TO_M3
                    unit = "m³"

            elif category_name == "Structural Framing":
                length = el.get_Parameter(DB.BuiltInParameter.CURVE_ELEM_LENGTH)
                if length and length.HasValue:
                    qty = length.AsDouble() * FT_TO_M
                    unit = "m"

            elif category_name == "Structural Columns":
                mat = el.LookupParameter("Structural Material")
                mat_elem = revit.doc.GetElement(mat.AsElementId()) if mat else None
                mname = mat_elem.Name if mat_elem else ""
                if mname == CONCRETE_NAME:
                    vol = el.LookupParameter("Volume")
                    if vol and vol.HasValue:
                        qty = vol.AsDouble() * FT3_TO_M3
                        unit = "m³"
                elif mname == STEEL_NAME:
                    length = el.get_Parameter(DB.BuiltInParameter.CURVE_ELEM_LENGTH)
                    if length and length.HasValue:
                        qty = length.AsDouble() * FT_TO_M
                        unit = "m"

            elif category_name == "Structural Rebar":
                bl = el.LookupParameter("Total Bar Length")
                if bl and bl.HasValue:
                    qty = bl.AsDouble() * FT_TO_M
                    unit = "m"

            elif category_name == "Plumbing":
                cat_id = el.Category.Id.IntegerValue
                if cat_id in (
                    int(DB.BuiltInCategory.OST_PlumbingFixtures),
                    int(DB.BuiltInCategory.OST_PipeFitting)
                ):
                    qty = 1
                    unit = "No."
                elif cat_id == int(DB.BuiltInCategory.OST_PipeCurves):
                    length = el.get_Parameter(DB.BuiltInParameter.CURVE_ELEM_LENGTH)
                    if length and length.HasValue:
                        qty = length.AsDouble() * FT_TO_M
                        unit = "m"

            elif category_name == "Electrical":
                length = el.get_Parameter(DB.BuiltInParameter.CURVE_ELEM_LENGTH)
                if length and length.HasValue:
                    qty = length.AsDouble() * FT_TO_M
                    unit = "m"

            comment = ""
            cp = el_type.LookupParameter("Type Comments")
            if cp and cp.HasValue:
                comment = cp.AsString()

            grouped.setdefault(name, {"qty": 0.0, "rate": rate, "total": 0.0, "unit": unit, "comment": comment})
            grouped[name]["qty"] += qty
            grouped[name]["total"] += total

        except:
            skipped += 1

    if grouped:
        sheet.write(row, 1, category_name.upper(), fmt_section)
        row += 1
        if category_name in CATEGORY_DESCRIPTIONS:
            sheet.write(row, 1, CATEGORY_DESCRIPTIONS[category_name], fmt_description)
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
            row += 1
            if data["comment"]:
                sheet.write(row, 1, data["comment"], fmt_italic)
                row += 1
            subtotal += data["total"]

        sheet.write(row, 1, category_name.upper()+" TO COLLECTION", fmt_bold)
        sheet.write(row, 5, round(subtotal, 2), fmt_money)
        category_totals.append((category_name.upper(), round(subtotal, 2)))
        row += 2

# --- Summary ---
sheet.write(row, 0, "COLLECTION", fmt_bold); row += 1
for name, total in category_totals:
    sheet.write(row, 0, name, fmt_normal)
    sheet.write(row, 5, total, fmt_money)
    row += 1

sheet.write(row, 0, "GRAND TOTAL", fmt_bold)
sheet.write(row, 5, sum(t[1] for t in category_totals), fmt_money)

wb.close()
MessageBox.Show(
    "BOQ export complete!\nSaved to Desktop:\n{}\nSkipped: {}".format(xlsx_path, skipped),
    "✅ XLSX Export"
)
