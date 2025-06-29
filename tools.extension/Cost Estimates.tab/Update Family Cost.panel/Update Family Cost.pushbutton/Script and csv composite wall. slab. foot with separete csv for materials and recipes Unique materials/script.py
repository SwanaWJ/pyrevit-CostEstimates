# -*- coding: utf-8 -*-
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
            name = row["Item"].strip()
            unit_cost = float(row["UnitCost"])
            material_prices[name] = unit_cost
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

# --- Process Types ---
updated = []
skipped = []
missing_materials = set()

with revit.Transaction("Set Composite Costs from CSV"):

    # WALL TYPES
    for wt in DB.FilteredElementCollector(revit.doc).OfClass(DB.WallType):
        name = wt.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if not name:
            continue
        type_name = name.AsString().strip()
        if type_name in recipes:
            total_cost = 0
            valid = True
            for mat, qty in recipes[type_name].items():
                if mat in material_prices:
                    total_cost += qty * material_prices[mat]
                else:
                    missing_materials.add(mat)
                    skipped.append("{0} (missing price for {1})".format(type_name, mat))
                    valid = False
                    break
            if valid:
                cost_param = wt.LookupParameter("Cost")
                if cost_param and cost_param.StorageType == DB.StorageType.Double:
                    cost_param.Set(total_cost)
                    updated.append((type_name, total_cost))

    # FLOOR TYPES
    for ft in DB.FilteredElementCollector(revit.doc).OfClass(DB.FloorType):
        name = ft.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if not name:
            continue
        type_name = name.AsString().strip()
        if type_name in recipes:
            total_cost = 0
            valid = True
            for mat, qty in recipes[type_name].items():
                if mat in material_prices:
                    total_cost += qty * material_prices[mat]
                else:
                    missing_materials.add(mat)
                    skipped.append("{0} (missing price for {1})".format(type_name, mat))
                    valid = False
                    break
            if valid:
                cost_param = ft.LookupParameter("Cost")
                if cost_param and cost_param.StorageType == DB.StorageType.Double:
                    cost_param.Set(total_cost)
                    updated.append((type_name, total_cost))

    # STRUCTURAL FRAMING (Generic beams)
    for sym in DB.FilteredElementCollector(revit.doc).OfClass(DB.FamilySymbol):
        if sym.Category and sym.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_StructuralFraming):
            name = sym.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
            if not name:
                continue
            type_name = name.AsString().strip()
            if type_name in recipes:
                total_cost = 0
                valid = True
                for mat, qty in recipes[type_name].items():
                    if mat in material_prices:
                        total_cost += qty * material_prices[mat]
                    else:
                        missing_materials.add(mat)
                        skipped.append("{0} (missing price for {1})".format(type_name, mat))
                        valid = False
                        break
                if valid:
                    cost_param = sym.LookupParameter("Cost")
                    if cost_param and cost_param.StorageType == DB.StorageType.Double:
                        cost_param.Set(total_cost)
                        updated.append((type_name, total_cost))

    # STRUCTURAL FOUNDATIONS (WallFoundationType — like RC beams)
    for wf in DB.FilteredElementCollector(revit.doc).OfClass(DB.WallFoundationType):
        name = wf.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if not name:
            continue
        type_name = name.AsString().strip()
        if type_name in recipes:
            total_cost = 0
            valid = True
            for mat, qty in recipes[type_name].items():
                if mat in material_prices:
                    total_cost += qty * material_prices[mat]
                else:
                    missing_materials.add(mat)
                    skipped.append("{0} (missing price for {1})".format(type_name, mat))
                    valid = False
                    break
            if valid:
                cost_param = wf.LookupParameter("Cost")
                if cost_param and cost_param.StorageType == DB.StorageType.Double:
                    cost_param.Set(total_cost)
                    updated.append((type_name, total_cost))

# --- Summary ---
summary = ""

if updated:
    summary += "Updated Types:\n"
    for name, cost in updated:
        summary += "{0}: {1} ZMW\n".format(name, cost)

if skipped:
    summary += "\nSkipped Types:\n" + "\n".join(skipped)

if missing_materials:
    summary += "\n\nMissing materials in material_unit_costs.csv:\n"
    summary += "\n".join(sorted(missing_materials))

if not updated and not skipped:
    summary = "No matching types found."

forms.alert(summary, title="Composite Cost Calculation")
