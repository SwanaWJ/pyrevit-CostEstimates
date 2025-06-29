# -*- coding: utf-8 -*-
from pyrevit import revit, DB, forms

# Settings
SHARED_PARAM_NAME = "Test_1234"
COST_PARAM_NAME = "Cost"
AREA_PARAM_NAME = "Area"
VOLUME_PARAM_NAME = "Volume"

# Collect all elements of relevant categories
categories = [
    DB.BuiltInCategory.OST_Walls,
    DB.BuiltInCategory.OST_Floors,
    DB.BuiltInCategory.OST_Ceilings,
    DB.BuiltInCategory.OST_Roofs,
    DB.BuiltInCategory.OST_Doors,
    DB.BuiltInCategory.OST_Windows
]

all_elements = []
for bic in categories:
    collector = DB.FilteredElementCollector(revit.doc).OfCategory(bic).WhereElementIsNotElementType()
    all_elements.extend(collector)

updated = []
skipped = []

t = DB.Transaction(revit.doc, "Populate Test_1234 values")
t.Start()

for el in all_elements:
    try:
        cost = el.LookupParameter(COST_PARAM_NAME)
        area = el.LookupParameter(AREA_PARAM_NAME)
        volume = el.LookupParameter(VOLUME_PARAM_NAME)
        target = el.LookupParameter(SHARED_PARAM_NAME)

        if not (cost and target and (area or volume)):
            skipped.append("{} (missing parameter)".format(el.Id))
            continue

        cost_val = cost.AsDouble() if cost.HasValue else 0
        area_val = area.AsDouble() if (area and area.HasValue) else 0
        volume_val = volume.AsDouble() if (volume and volume.HasValue) else 0

        amount = cost_val * (area_val if area_val > 0 else volume_val)
        target.Set(amount)
        updated.append(str(el.Id))

    except Exception as e:
        skipped.append("{} (error: {})".format(el.Id, str(e)))

t.Commit()

# Show results
msg = "✅ Populated Test_1234 for {} elements.".format(len(updated))
if skipped:
    msg += "\n\n⚠️ Skipped {} elements:\n".format(len(skipped))
    msg += "\n".join(skipped[:10])  # Show only first 10

forms.alert(msg, title="Test_1234 Update Complete", warn_icon=False)
