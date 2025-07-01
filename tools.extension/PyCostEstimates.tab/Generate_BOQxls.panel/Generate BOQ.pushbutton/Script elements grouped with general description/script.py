# -*- coding: utf-8 -*-
import os
import csv
import string
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
}

BLOCKWORK_DESC = "Blockwork in hollow concrete blocks load bearing (crushing strength not less than 3.5N/mm2) in cement mortar (1:4) as described."

# --- Start CSV Export ---
skipped = 0
item_letter = iter(string.ascii_uppercase)

with open(csv_path, 'w') as f:
    writer = csv.writer(f)
    writer.writerow(["ITEM", "DESCRIPTION", "UNIT", "QTY", "RATE ZMW", "Test_1234 ZMW"])

    for category_name, bic in CATEGORY_MAP.items():
        elements = DB.FilteredElementCollector(revit.doc).OfCategory(bic).WhereElementIsNotElementType().ToElements()
        grouped = {}

        for el in elements:
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
            unit = "ea"

            # Fallback order: Area → Volume → Length
            if el.LookupParameter("Area"):
                param = el.LookupParameter("Area")
                qty = param.AsDouble()
                unit = "m²"
            elif el.LookupParameter("Volume"):
                param = el.LookupParameter("Volume")
                qty = param.AsDouble()
                unit = "m³"
            elif el.get_Parameter(DB.BuiltInParameter.CURVE_ELEM_LENGTH):
                param = el.get_Parameter(DB.BuiltInParameter.CURVE_ELEM_LENGTH)
                qty = param.AsDouble()
                unit = "m"

            if name not in grouped:
                grouped[name] = {"qty": 0.0, "rate": rate, "total": 0.0, "unit": unit}
            grouped[name]["qty"] += qty
            grouped[name]["total"] += total

        if grouped:
            writer.writerow([])  # spacing
            writer.writerow([category_name.upper()])

            # Insert general description under Walla only
            if category_name == "Walls":
                writer.writerow(["", BLOCKWORK_DESC])  # this line becomes bold+underline in .xlsx later

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

# --- Notify ---
forms.alert("BOQ export complete:\n\nItems written to Desktop:\n{}\nSkipped: {}".format(csv_path, skipped), title="BOQ CSV Export")
