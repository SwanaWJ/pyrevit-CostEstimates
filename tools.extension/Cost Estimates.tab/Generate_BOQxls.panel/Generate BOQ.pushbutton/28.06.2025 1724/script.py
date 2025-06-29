# -*- coding: utf-8 -*-
import os
import csv
import string
import codecs
from pyrevit import revit, DB, forms

# --- Save to Desktop ---
desktop = os.path.expanduser("~/Desktop")
csv_path = os.path.join(desktop, "BOQ_Export_From_Model.csv")

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
    "Roofs": DB.BuiltInCategory.OST_Roofs
}

CATEGORY_DESCRIPTIONS = {
    "Walls": "Blockwork in hollow concrete blocks load bearing (crushing strength not less than 3.5N/mm2) in cement mortar (1:4) as described.",
    "Doors": "Flush doors with hardwood frame and internal plywood finish as specified.",
    "Windows": "Aluminium sliding windows with clear glass and burglar bars as specified.",
    "Structural Foundations": "Mass concrete, RC footing and ground beams cast in situ as described.",
    "Structural Framing": "Mild steel beams and trusses to detail.",
    "Structural Columns": "Steel and concrete columns as detailed on structural drawings.",
    "Roofs": "0.5mm IBR/IT4 Pre-painted roof sheeting and fixing in accordance with manufacturer's instructions (measured net - no allowance made for laps)."
}

# --- Constants ---
FT3_TO_M3 = 0.0283168
FT2_TO_M2 = 0.092903
FT_TO_M = 0.3048
CONCRETE_NAME = "Concrete - Cast-in-Place Concrete"
STEEL_NAME = "Metal - Steel 43-275"

# --- Start CSV Export ---
skipped = 0
category_totals = []

with codecs.open(csv_path, 'w', encoding='utf-8') as f:
    writer = csv.writer(f)
    writer.writerow(["ITEM", "DESCRIPTION", "UNIT", "QTY", "RATE (ZAR)", "AMOUNT (ZAR)"])

    for category_name, bic in CATEGORY_MAP.items():
        item_letter = iter(string.ascii_uppercase)
        elements = DB.FilteredElementCollector(revit.doc).OfCategory(bic).WhereElementIsNotElementType().ToElements()
        grouped = {}

        for el in elements:
            try:
                el_type = revit.doc.GetElement(el.GetTypeId())
                type_param = el_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                if not type_param:
                    skipped += 1
                    continue

                name = type_param.AsString()
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
                    unit = "No."
                elif category_name in ["Walls", "Roofs"]:
                    area_param = el.LookupParameter("Area")
                    if area_param and area_param.HasValue:
                        qty = area_param.AsDouble() * FT2_TO_M2
                        unit = "m²"
                elif category_name in ["Structural Foundations", "Floors"]:
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

                if name not in grouped:
                    grouped[name] = {"qty": 0.0, "rate": rate, "total": 0.0, "unit": unit}
                grouped[name]["qty"] += qty
                grouped[name]["total"] += total

            except:
                skipped += 1
                continue

        if grouped:
            writer.writerow([])  # Blank line
            writer.writerow(["", category_name.upper(), "", "", "", ""])
            if category_name in CATEGORY_DESCRIPTIONS:
                writer.writerow(["", CATEGORY_DESCRIPTIONS[category_name], "", "", "", ""])
            subtotal = 0.0

            for name, data in grouped.items():
                letter = next(item_letter)
                writer.writerow([
                    letter,
                    name,
                    data["unit"],
                    round(data["qty"], 2),
                    round(data["rate"], 2),
                    round(data["total"], 2)
                ])
                subtotal += data["total"]

            writer.writerow(["", category_name.upper() + " TO COLLECTION", "", "", "", round(subtotal, 2)])
            category_totals.append((category_name.upper(), round(subtotal, 2)))

    # Final collection summary
    writer.writerow([])
    writer.writerow(["COLLECTION"])
    for name, total in category_totals:
        writer.writerow([name, "", "", "", "", total])

# --- Notify ---
forms.alert(
    "BOQ export complete:\n\nSaved to Desktop:\n{}\nSkipped: {}".format(csv_path, skipped),
    title="BOQ CSV Export"
)
