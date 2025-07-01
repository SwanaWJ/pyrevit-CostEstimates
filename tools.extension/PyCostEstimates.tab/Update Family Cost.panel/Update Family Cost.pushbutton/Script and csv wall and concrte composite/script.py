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
            continue  # Skip bad data

# --- Define Wall Recipes ---
wall_recipes = {
    "Wall-190mm-Block": {
        "Block_190mm": 12,
        "Cement_50kg": 0.08,
        "RiverSand_m3": 0.05
    },
    # Add more walls if needed
}

# --- Define Concrete Mix Recipes ---
concrete_recipes = {
    "C25 Concrete": {
        "Cement_50kg_m3": 12,
        "Fine Aggregates_m3": 0.08,
        "Corse Aggregates_m3": 0.05
    },
    # Add more concrete grades if needed
}

# --- Start Transaction ---
updated = []
skipped = []

with revit.Transaction("Set Composite Costs"):

    # ---- WALL TYPES ----
    walltypes = DB.FilteredElementCollector(revit.doc).OfClass(DB.WallType)
    for wt in walltypes:
        name_param = wt.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if not name_param:
            skipped.append("Unnamed WallType")
            continue
        wall_name = name_param.AsString().strip()

        if wall_name in wall_recipes:
            total_cost = 0
            recipe = wall_recipes[wall_name]
            for mat, qty in recipe.items():
                price = material_prices.get(mat)
                if price is None:
                    skipped.append(wall_name + " (missing price for {})".format(mat))
                    break
                total_cost += qty * price
            else:
                cost_param = wt.LookupParameter("Cost")
                if cost_param and cost_param.StorageType == DB.StorageType.Double:
                    cost_param.Set(total_cost)
                    updated.append((wall_name, total_cost))
                else:
                    skipped.append(wall_name + " (no Cost param)")

    # ---- FLOOR TYPES (for Concrete) ----
    floortypes = DB.FilteredElementCollector(revit.doc).OfClass(DB.FloorType)
    for ft in floortypes:
        name_param = ft.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if not name_param:
            skipped.append("Unnamed FloorType")
            continue
        floor_name = name_param.AsString().strip()

        if floor_name in concrete_recipes:
            total_cost = 0
            recipe = concrete_recipes[floor_name]
            for mat, qty in recipe.items():
                price = material_prices.get(mat)
                if price is None:
                    skipped.append(floor_name + " (missing price for {})".format(mat))
                    break
                total_cost += qty * price
            else:
                cost_param = ft.LookupParameter("Cost")
                if cost_param and cost_param.StorageType == DB.StorageType.Double:
                    cost_param.Set(total_cost)
                    updated.append((floor_name, total_cost))
                else:
                    skipped.append(floor_name + " (no Cost param)")

# --- Result Summary ---
if updated:
    summary = ""
    for item in updated:
        name, cost = item
        summary += "{}: {} ZMW\n".format(name, cost)
else:
    summary = "No matching types found."

if skipped:
    summary += "\n\nSkipped:\n" + "\n".join(skipped)

forms.alert(summary, title="Composite Material Cost Update")
