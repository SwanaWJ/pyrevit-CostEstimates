# -*- coding: utf-8 -*-
import os
import csv
from pyrevit import revit, DB, script

# Setup
doc = revit.doc
output = script.get_output()
output_path = os.path.join(os.path.dirname(__file__), "BOQ_Export_From_Model.csv")

# Configuration
TARGET_PARAM = "Test_1234"
COST_PARAM = "Cost"

# CSV Headers
headers = ["Item", "Description", "Unit", "Qty", "Rate (ZMW)", "Amount (ZMW)"]
rows = []

# Target Categories
categories = {
    "Walls": DB.BuiltInCategory.OST_Walls,
    "Floors": DB.BuiltInCategory.OST_Floors,
    "Structural Foundations": DB.BuiltInCategory.OST_StructuralFoundation,
    "Structural Framing": DB.BuiltInCategory.OST_StructuralFraming,
    "Windows": DB.BuiltInCategory.OST_Windows,
    "Doors": DB.BuiltInCategory.OST_Doors
}

item_index = 0
skipped_count = 0

for cat_name, cat_enum in categories.items():
    elements = DB.FilteredElementCollector(doc).OfCategory(cat_enum).WhereElementIsNotElementType().ToElements()
    for el in elements:
        try:
            # Get Type
            symbol = el.Symbol if hasattr(el, "Symbol") else doc.GetElement(el.GetTypeId())
            desc_param = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
            description = desc_param.AsString() if desc_param else "Unknown"

            # Quantity (prioritized)
            qty = None
            unit = ""

            if el.LookupParameter("Count") and el.LookupParameter("Count").AsInteger() > 0:
                qty = el.LookupParameter("Count").AsInteger()
                unit = "item"
            elif el.LookupParameter("Area") and el.LookupParameter("Area").AsDouble() > 0:
                qty = el.LookupParameter("Area").AsDouble() * 0.092903  # ft² to m²
                unit = "m²"
            elif el.LookupParameter("Volume") and el.LookupParameter("Volume").AsDouble() > 0:
                qty = el.LookupParameter("Volume").AsDouble() * 0.0283168  # ft³ to m³
                unit = "m³"
            elif el.LookupParameter("Length") and el.LookupParameter("Length").AsDouble() > 0:
                qty = el.LookupParameter("Length").AsDouble() * 0.3048  # ft to m
                unit = "m"
            else:
                qty = 1
                unit = "item"

            # Rate and Total
            cost_param = symbol.LookupParameter(COST_PARAM)
            total_param = el.LookupParameter(TARGET_PARAM)

            rate = cost_param.AsDouble() if cost_param else 0
            total = total_param.AsDouble() if total_param else 0

            # Item Code A, B, C, ...
            item_code = chr(65 + item_index)
            item_index += 1

            rows.append([item_code, description, unit, round(qty, 2), round(rate, 2), round(total, 2)])

        except Exception as e:
            skipped_count += 1

# Write to CSV
with open(output_path, "w") as f:
    writer = csv.writer(f)
    writer.writerow(headers)
    writer.writerows(rows)

# Output
script.get_output().print_md(
    "**BOQ export complete:**\n\n{} items written to:\n{}\n\nSkipped: {}".format(
        len(rows), output_path, skipped_count
    )
)
