# --- Imports ---
import os
import csv
from pyrevit import revit, DB, forms

# --- Load Material Costs from CSV ---
script_dir = os.path.dirname(__file__)
csv_path = os.path.join(script_dir, "material_unit_costs.csv")

material_prices = {}
with open(csv_path, 'r') as file:
    reader = csv.DictReader(file)
    for row in reader:
        name = row["Item"].strip()
        try:
            cost = float(row["UnitCost"])
            material_prices[name] = cost
        except:
            continue  # Skip invalid values

# --- Define Composite Wall Recipes ---
wall_recipes = {
    "Wall-190mm-Block": {
        "Block_190mm": 12,
        "Cement_50kg": 0.08,
        "RiverSand_m3": 0.05
    },
    # Add more walls here if needed
}

# --- Start Transaction ---
updated = []
skipped = []

with revit.Transaction("Set Composite Wall Costs"):
    walltypes = DB.FilteredElementCollector(revit.doc).OfClass(DB.WallType)
    for wt in walltypes:
        # Safe way to get wall type name
        name_param = wt.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if not name_param:
            skipped.append("Unnamed WallType")
            continue
        wall_name = name_param.AsString().strip()

        if wall_name in wall_recipes:
            total_cost = 0
            missing_price = False
            for material, qty in wall_recipes[wall_name].items():
                unit_cost = material_prices.get(material)
                if unit_cost is not None:
                    total_cost += qty * unit_cost
                else:
                    skipped.append(wall_name + " (missing price for {})".format(material))
                    missing_price = True
                    break

            if not missing_price:
                cost_param = wt.LookupParameter("Cost")
                if cost_param:
                    try:
                        if cost_param.StorageType == DB.StorageType.Double:
                            cost_param.Set(total_cost)
                            updated.append((wall_name, total_cost))
                        else:
                            skipped.append(wall_name + " (Cost not numeric)")
                    except Exception as e:
                        skipped.append(wall_name + " (Set error: {})".format(str(e)))
                else:
                    skipped.append(wall_name + " (No Cost parameter)")

# --- Result Summary ---
if updated:
    summary = ""
    for item in updated:
        name, cost = item
        summary += "{}: {} ZMW\n".format(name, cost)
else:
    summary = "No matching wall types found."

if skipped:
    summary += "\n\nSkipped:\n" + "\n".join(skipped)

forms.alert(summary, title="Composite Wall Cost Update")
