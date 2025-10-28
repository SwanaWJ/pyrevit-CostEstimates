# -*- coding: utf-8 -*-
from __future__ import print_function
__title__  = "Material Schedule (DEBUG)"
__author__ = "Wachama J. Swana"
__doc__    = "Material schedule with debug logs & summary; saves to Desktop."

import os, csv, datetime
from collections import defaultdict

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter
from Autodesk.Revit.UI import TaskDialog
from pyrevit import revit, script

import System
DESKTOP = System.Environment.GetFolderPath(System.Environment.SpecialFolder.DesktopDirectory)
STAMP   = datetime.datetime.now().strftime("%Y%m%d_%H%M")
OUT_XLSX= os.path.join(DESKTOP, "Material_Schedule_{}.xlsx".format(STAMP))
DBG_BASES = os.path.join(DESKTOP, "WJS_material_bases_DEBUG.csv")
DBG_MATCH = os.path.join(DESKTOP, "WJS_recipe_matches_DEBUG.csv")

SCRIPT_DIR  = os.path.dirname(__file__)
RECIPES_CSV = os.path.join(SCRIPT_DIR, "recipes.csv")
COST_DIR    = os.path.join(SCRIPT_DIR, "material_costs")

from helpers import load_cost_folder, load_recipes, norm, price_lookup

doc = revit.doc

def alert(msg):
    TaskDialog.Show("Material Schedule", msg)

# --- sanity checks
if not os.path.isfile(RECIPES_CSV):
    alert("Missing recipes.csv at:\n{}".format(RECIPES_CSV)); script.exit()
if not os.path.isdir(COST_DIR):
    alert("Missing material_costs folder at:\n{}".format(COST_DIR)); script.exit()

cost_map = load_cost_folder(COST_DIR)
recipes  = load_recipes(RECIPES_CSV)

if not cost_map:
    alert("No rows loaded from cost CSVs in:\n{}\n\nCheck headers (Item/Material/Description, UoM/Unit, Rate/Price/Cost).".format(COST_DIR)); script.exit()
if not recipes:
    alert("No recipe rows loaded from:\n{}\n\nCheck headers (Category, FamilyOrTypePattern, BaseUnit, Constituent, Unit, QtyPerBase, [Waste%]).".format(RECIPES_CSV)); script.exit()

# ---- Category rules (align recipe BaseUnit to these)
CAT_RULES = [
    ("Block Work in Walls", BuiltInCategory.OST_Walls,               "m²", BuiltInParameter.HOST_AREA_COMPUTED),
    ("Concrete Works",      BuiltInCategory.OST_StructuralFoundation,"m³", BuiltInParameter.HOST_VOLUME_COMPUTED),
    ("Concrete Works",      BuiltInCategory.OST_Floors,              "m³", BuiltInParameter.HOST_VOLUME_COMPUTED),
    ("Concrete Works",      BuiltInCategory.OST_StructuralFraming,   "m³", BuiltInParameter.HOST_VOLUME_COMPUTED),
    ("Concrete Works",      BuiltInCategory.OST_StructuralColumns,   "m³", BuiltInParameter.HOST_VOLUME_COMPUTED),
]
CAT_BASEUNIT = {c:u for (c,_,u,_) in CAT_RULES}

def get_item_display_name(el):
    try:
        sym = getattr(el, "Symbol", None)
        if sym:
            fam = getattr(sym, "Family", None)
            return u"{} : {}".format(fam.Name if fam else "", sym.Name or "").strip()
    except Exception:
        pass
    try:
        return el.Name or ""
    except Exception:
        return ""

# ---- 1) Collect model bases
bases = defaultdict(lambda: defaultdict(float))  # {cat: {name: qty}}
total_elements = 0
for catname, bic, base_unit, bip in CAT_RULES:
    try:
        for el in FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType():
            total_elements += 1
            p = el.get_Parameter(bip)
            if not p: continue
            val = p.AsDouble()
            if base_unit == "m²":
                qty = val * 0.09290304   # ft² -> m²
            elif base_unit == "m³":
                qty = val * 0.028316846592  # ft³ -> m³
            else:
                qty = val
            if qty <= 1e-9: continue
            bases[catname][get_item_display_name(el)] += qty
    except Exception:
        continue

# Write bases debug
try:
    with open(DBG_BASES, "w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["Category","Item Name","Base Unit","Total Base Qty"])
        for cat in sorted(bases.keys()):
            for name, qty in sorted(bases[cat].items()):
                w.writerow([cat, name, CAT_BASEUNIT.get(cat,""), qty])
except Exception:
    pass

# ---- 2) Expand with recipes
materials_by_cat = {}  # output aggregation
match_rows = []        # debug rows (category, item, regex, base_qty, constituent, unit, perbase, waste, qty)

for catname, name_qty in bases.items():
    rules = recipes.get(catname, [])
    if not rules:
        continue
    materials_by_cat[catname] = {}
    base_unit = CAT_BASEUNIT.get(catname, "")
    for item_name, base_total in name_qty.items():
        for r in rules:
            if r.get("base_unit") != base_unit:
                continue
            try:
                if not r["regex"].search(item_name or ""):
                    continue
            except Exception:
                continue
            perb  = float(r.get("per_base",0.0) or 0.0)
            waste = float(r.get("waste",0.0) or 0.0)
            qty   = perb * base_total
            if waste > 0: qty *= (1.0 + waste/100.0)
            mat_name = r["material"].strip()
            unit     = r["unit"].strip()
            rate, src, unit_from_price = price_lookup(cost_map, mat_name)
            if unit_from_price: unit = unit_from_price
            key = (norm(mat_name), unit)
            cur = materials_by_cat[catname].get(key, {"name": mat_name, "unit": unit, "qty": 0.0, "rate": 0.0, "src": src})
            cur["qty"] += qty
            if cur["rate"] == 0.0 and rate:
                cur["rate"] = rate; cur["src"] = src
            materials_by_cat[catname][key] = cur
            match_rows.append([catname, item_name, r["regex"].pattern, base_total, mat_name, unit, perb, waste, qty])

# Write matches debug
try:
    with open(DBG_MATCH, "w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["Category","Matched Item Name","Recipe Pattern","Item Base Qty",
                    "Constituent","Unit","QtyPerBase","Waste%","Constituent Qty"])
        for row in match_rows:
            w.writerow(row)
except Exception:
    pass

# ---- 3) Write Excel (or CSV) to Desktop
wrote_any = False
total_lines = 0
try:
    import xlsxwriter
    wb = xlsxwriter.Workbook(OUT_XLSX)
    ws = wb.add_worksheet("MATERIAL SCHEDULE")
    ws.set_tab_color("#70AD47")
    fmt_title = wb.add_format({'bold': True, 'font_size': 12, 'align':'center'})
    fmt_head  = wb.add_format({'bold': True, 'bg_color':'#DDDDDD', 'border':1, 'align':'center', 'valign':'vcenter'})
    fmt_txt   = wb.add_format({'border':1})
    fmt_num   = wb.add_format({'border':1, 'num_format':'#,##0.00'})
    fmt_cat   = wb.add_format({'bold': True, 'bg_color':'#E2F0D9', 'border':1})
    fmt_sub   = wb.add_format({'bold': True, 'border':1, 'num_format':'#,##0.00'})

    ws.merge_range(0,0,0,6, "MATERIAL SCHEDULE (Constituents)", fmt_title)
    ws.write_row(2, 0, ["No.","Description of Material","Unit","Quantity","Rate","Amount","Price Source (CSV)"], fmt_head)
    ws.set_column(0,0,6); ws.set_column(1,1,50); ws.set_column(2,2,10); ws.set_column(3,5,14); ws.set_column(6,6,24)

    r = 3; i = 1
    for cat in sorted(materials_by_cat.keys()):
        bucket = materials_by_cat[cat]
        if not bucket:
            continue
        ws.merge_range(r,0,r,6,cat,fmt_cat); r += 1
        cat_amount_cells = []
        for (_, unit), d in sorted(bucket.items(), key=lambda kv: kv[1]['name'].lower()):
            ws.write_number(r,0,i,fmt_txt)
            ws.write_string(r,1,d['name'],fmt_txt)
            ws.write_string(r,2,unit,fmt_txt)
            ws.write_number(r,3,float(d['qty']),fmt_num)
            ws.write_number(r,4,float(d.get('rate',0.0)),fmt_num)
            ws.write_formula(r,5,"=D{0}*E{0}".format(r+1),fmt_num)
            ws.write_string(r,6,d.get('src',''),fmt_txt)
            cat_amount_cells.append("F{}".format(r+1))
            r += 1; i += 1; total_lines += 1
        if cat_amount_cells:
            ws.write(r,4,"SUBTOTAL",fmt_cat)
            ws.write_formula(r,5,"=SUM({})".format(",".join(cat_amount_cells)),fmt_sub)
            r += 2
    if i > 1:
        ws.write(r,4,"GRAND TOTAL",fmt_head)
        ws.write_formula(r,5,"=SUM(F4:F{})".format(r),fmt_sub)
    wb.close()
    wrote_any = True
except Exception:
    # CSV fallback
    out_csv = OUT_XLSX.replace(".xlsx",".csv")
    with open(out_csv, "w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["MATERIAL SCHEDULE (Constituents)"]); w.writerow([])
        w.writerow(["No.","Description","Unit","Quantity","Rate","Amount","Price Source (CSV)"])
        idx=1
        for cat in sorted(materials_by_cat.keys()):
            w.writerow([]); w.writerow([cat])
            for (_,unit), d in sorted(materials_by_cat[cat].items(), key=lambda kv: kv[1]['name'].lower()):
                qty=float(d['qty']); rate=float(d.get('rate',0.0))
                w.writerow([idx, d['name'], unit, qty, rate, qty*rate, d.get('src','')])
                idx += 1; total_lines += 1
    OUT_XLSX = out_csv
    wrote_any = True

# ---- 4) Final summary
base_items = sum(len(b) for b in bases.values())
msg = [
    "Scan summary:",
    "- Elements scanned: {}".format(total_elements),
    "- Base items found (unique names across categories): {}".format(base_items),
    "- Recipe matches (rows): {}".format(len(match_rows)),
    "- Output lines written: {}".format(total_lines),
    "",
    "Files saved to Desktop:",
    "- {}".format(OUT_XLSX),
    "- {}".format(DBG_BASES),
    "- {}".format(DBG_MATCH),
]
if total_lines == 0:
    msg.append("")
    msg.append("No schedule lines were produced.")
    msg.append("Most likely causes:")
    msg.append("• Recipe patterns didn’t match your item names.")
    msg.append("• Recipe BaseUnit didn’t match the category’s base unit.")
    msg.append("Open the two DEBUG CSVs to see the base names & which recipes matched.")
alert("\n".join(msg))
