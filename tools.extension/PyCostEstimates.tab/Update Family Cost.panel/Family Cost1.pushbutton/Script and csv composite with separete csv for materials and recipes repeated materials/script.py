import os
import csv
from pyrevit import revit, DB, forms

# --- Paths ---
script_dir = os.path.dirname(__file__)
csv_costs_path = os.path.join(script_dir, "material_unit_costs.csv")
csv_recipes_path = os.path.join(script_dir, "recipes.csv")

# --- Load Material Unit Costs ---
material_prices = {}
with open(csv_costs_path, 'r') as file:
    reader = csv.DictReader(file)
    for row in reader:
        try:
            material_prices[row["Item"].strip()] = float(row["UnitCost"])
        except:
            continue

# --- Load Recipes ---
recipes = {}
with open(csv_recipes_path, 'r') as file:
    reader = csv.DictReader(file)
    for row in reader:
        type_name = row["Type"].strip()
        component = row["Component"].strip()
        try:
            qty = float(row["Quantity"])
        except:
            continue
        if type_name not in recipes:
            recipes[type_name] = {}
        recipes[type_name][component] = qty

# --- Process Revit Types ---
updated = []
skipped = []

with revit.Transaction("Set Composite Costs from CSV"):

    # WALL TYPES
    for wt in DB.FilteredElementCollector(revit.doc).OfClass(DB.WallType):
        name_param = wt.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if not name_param:
            continue
        wall_name = name_param.AsString().strip()

        if wall_name in recipes:
            total_cost = 0
            missing = False
            for mat, qty in recipes[wall_name].items():
                if mat in material_prices:
                    total_cost += qty * material_prices[mat]
                else:
                    skipped.append("{} (missing price: {})".format(wall_name, mat))
                    missing = True
                    break
            if not missing:
                cost_param = wt.LookupParameter("Cost")
                if cost_param and cost_param.StorageType == DB.StorageType.Double:
                    cost_param.Set(total_cost)
                    updated.append((wall_name, total_cost))

    # FLOOR TYPES (e.g. concrete mixes)
    for ft in DB.FilteredElementCollector(revit.doc).OfClass(DB.FloorType):
        name_param = ft.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if not name_param:
            continue
        floor_name = name_param.AsString().strip()

        if floor_name in recipes:
            total_cost = 0
            missing = False
            for mat, qty in recipes[floor_name].items():
                if mat in material_prices:
                    total_cost += qty * material_prices[mat]
                else:
                    skipped.append("{} (missing price: {})".format(floor_name, mat))
                    missing = True
                    break
            if not missing:
                cost_param = ft.LookupParameter("Cost")
                if cost_param and cost_param.StorageType == DB.StorageType.Double:
                    cost_param.Set(total_cost)
                    updated.append((floor_name, total_cost))

# --- Show Summary ---
if updated:
    summary = ""
    for name, cost in updated:
        summary += "{}: {} ZMW\n".format(name, cost)
else:
    summary = "No matching types found."

if skipped:
    summary += "\n\nSkipped:\n" + "\n".join(skipped)

forms.alert(summary, title="Composite Cost Calculation")
