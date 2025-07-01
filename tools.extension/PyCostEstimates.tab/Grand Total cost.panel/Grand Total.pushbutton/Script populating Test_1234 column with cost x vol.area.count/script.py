# -*- coding: utf-8 -*-
from pyrevit import revit, DB
from pyrevit import script

output = script.get_output()
doc = revit.doc

# Constants
PARAM_COST = "Cost"
PARAM_TARGET = "Test_1234"
FT3_TO_M3 = 0.0283168
FT2_TO_M2 = 0.092903

# Supported categories
categories = [
    DB.BuiltInCategory.OST_Doors,
    DB.BuiltInCategory.OST_Walls,
    DB.BuiltInCategory.OST_Windows,
    DB.BuiltInCategory.OST_Floors,
    DB.BuiltInCategory.OST_StructuralFraming,
    DB.BuiltInCategory.OST_StructuralFoundation
]

# Collect all relevant elements
elements = []
for cat in categories:
    elements += DB.FilteredElementCollector(doc)\
                 .OfCategory(cat)\
                 .WhereElementIsNotElementType()\
                 .ToElements()

# Begin transaction
t = DB.Transaction(doc, "Set Test_1234 = Cost × Quantity (metric)")
t.Start()

updated = 0
skipped = []

for elem in elements:
    try:
        # Get type element and parameters
        type_elem = doc.GetElement(elem.GetTypeId())
        cost_param = type_elem.LookupParameter(PARAM_COST)
        target_param = elem.LookupParameter(PARAM_TARGET)

        if not cost_param or not target_param or target_param.IsReadOnly:
            raise Exception("Missing or read-only parameter")

        cost_val = cost_param.AsDouble()
        factor = 1.0  # fallback for Count

        # Use Volume (convert ft³ to m³)
        vol_param = elem.LookupParameter("Volume")
        if vol_param and vol_param.HasValue:
            factor = vol_param.AsDouble() * FT3_TO_M3

        # Else use Area (convert ft² to m²)
        elif elem.LookupParameter("Area") and elem.LookupParameter("Area").HasValue:
            factor = elem.LookupParameter("Area").AsDouble() * FT2_TO_M2

        result = cost_val * factor
        target_param.Set(result)
        updated += 1

    except Exception as e:
        skipped.append((elem.Id, str(e)))

t.Commit()

# Output
output.print_md("Updated {} element(s) with '{}' = Cost × Quantity (in metric units).".format(updated, PARAM_TARGET))
if skipped:
    output.print_md("Skipped {} element(s) due to issues:".format(len(skipped)))
    for item in skipped:
        output.print_md("- Element ID: {} | Reason: {}".format(item[0], item[1]))
