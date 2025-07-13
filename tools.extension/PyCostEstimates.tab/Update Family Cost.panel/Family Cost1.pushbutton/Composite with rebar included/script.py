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

# --- Initialize Trackers ---
updated = []
skipped = []
missing_materials = set()

with revit.Transaction("Set Composite Costs from CSV"):

    # Utility function to apply cost if matched
    def apply_cost_to_elements(collected, enum_value, name_param=True):
        for elem in collected:
            if not elem.Category or elem.Category.Id.IntegerValue != int(enum_value):
                continue
            if name_param:
                param = elem.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
            else:
                param = elem.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
            if not param:
                continue
            type_name = param.AsString().strip()
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
                    cost_param = elem.LookupParameter("Cost")
                    if cost_param and cost_param.StorageType == DB.StorageType.Double and not cost_param.IsReadOnly:
                        cost_param.Set(total_cost)
                        updated.append((type_name, total_cost))
                    else:
                        skipped.append("{0} (no editable 'Cost' parameter)".format(type_name))

    # Apply cost to major categories
    apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfClass(DB.WallType), DB.BuiltInCategory.OST_Walls)
    apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfClass(DB.FloorType), DB.BuiltInCategory.OST_Floors)
    apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfClass(DB.WallFoundationType), DB.BuiltInCategory.OST_StructuralFoundation)
    apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfClass(DB.FamilySymbol), DB.BuiltInCategory.OST_StructuralFraming)
    apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfClass(DB.FamilySymbol), DB.BuiltInCategory.OST_GenericModel)
    apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfClass(DB.RoofType), DB.BuiltInCategory.OST_Roofs)
    apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfClass(DB.CeilingType), DB.BuiltInCategory.OST_Ceilings)
    apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfClass(DB.FamilySymbol), DB.BuiltInCategory.OST_StructuralColumns)
    apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfClass(DB.FamilySymbol), DB.BuiltInCategory.OST_Doors)
    apply_cost_to_elements(DB.FilteredElementCollector(revit.doc).OfClass(DB.FamilySymbol), DB.BuiltInCategory.OST_Windows)

    # ✅ Added: Structural Rebar support
    apply_cost_to_elements(
        DB.FilteredElementCollector(revit.doc).OfClass(DB.Structure.RebarBarType),
        DB.BuiltInCategory.OST_Rebar,
        name_param=False
    )

# --- Summary ---
summary = ""

if updated:
    summary += "✅ Updated Types:\n"
    for name, cost in updated:
        summary += "- {}: {:.2f} ZMW\n".format(name, cost)

if skipped:
    summary += "\n⚠️ Skipped Types:\n" + "\n".join(skipped)

if missing_materials:
    summary += "\n\n❗Missing materials in material_unit_costs.csv:\n"
    summary += "\n".join(sorted(missing_materials))

if not updated and not skipped:
    summary = "No matching types found."

forms.alert(summary, title="Composite Cost Calculation")
