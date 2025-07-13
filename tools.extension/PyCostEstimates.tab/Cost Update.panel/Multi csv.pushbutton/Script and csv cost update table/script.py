# --- Imports ---
import os
import csv
from pyrevit import revit, DB, forms

# --- CSV file path (must be in same folder as this script) ---
script_dir = os.path.dirname(__file__)
csv_file_path = os.path.join(script_dir, "MPI_Lusaka_Max_Prices.csv")

# --- Read CSV into dictionary ---
cost_dict = {}
with open(csv_file_path, mode='r') as file:
    reader = csv.DictReader(file)
    for row in reader:
        name = row["Product Description"].strip()
        try:
            cost = float(row["Max"].replace(",", ""))
            cost_dict[name] = cost
        except:
            continue  # Skip non-numeric rows

# --- Start transaction to update family type costs ---
updated = []
skipped = []

with revit.Transaction("Update Costs from CSV"):
    collector = DB.FilteredElementCollector(revit.doc).OfClass(DB.FamilySymbol)
    for symbol in collector:
        try:
            name_param = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
            if not name_param:
                continue
            type_name = name_param.AsString().strip()
            if type_name in cost_dict:
                cost_param = symbol.LookupParameter("Cost")
                if cost_param and cost_param.StorageType == DB.StorageType.Double:
                    cost_param.Set(cost_dict[type_name])
                    updated.append(type_name)
                else:
                    skipped.append(type_name + " (no cost param)")
        except Exception as e:
            skipped.append("Error: " + str(e))

# --- Summary popup ---
forms.alert(
    "{} type(s) updated.\n{} type(s) skipped.".format(len(updated), len(skipped)),
    title="Cost Update Complete"
)
