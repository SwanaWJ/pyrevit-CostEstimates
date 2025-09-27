# -*- coding: utf-8 -*-
import os
import string
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
import clr

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import MessageBox
from pyrevit import revit, DB

# --- Save path ---
desktop = os.path.expanduser("~/Desktop")
xlsx_path = os.path.join(desktop, "BOQ_Export_From_Model.xlsx")

# --- Parameters ---
PARAM_COST  = "Cost"        # rate on type / material
PARAM_TOTAL = "Test_1234"   # kept for backward compatibility (not used to write Amount)

# --- Ordered Categories (must match CATEGORY_MAP keys exactly) ---
CATEGORY_ORDER = [
    "Structural Foundations",
    "Block Work in Walls",
    "Structural Columns",
    "Structural Framing",
    "Structural Rebar",
    "Roofs",
    "Windows",
    "Doors",
    "Electrical",
    "Plumbing",
    "Painting",                  # <-- virtual: painted wall faces (all together)
    "Wall and Floor Finishes",
    "Furniture",                 # <-- includes Furniture + Furniture Systems
]

# Sentinel for virtual categories
VIRTUAL_PAINT = object()

# --- Categories map (keys match above 1:1) ---
CATEGORY_MAP = {
    "Structural Foundations": DB.BuiltInCategory.OST_StructuralFoundation,
    "Block Work in Walls":    DB.BuiltInCategory.OST_Walls,
    "Structural Columns":     DB.BuiltInCategory.OST_StructuralColumns,
    "Structural Framing":     DB.BuiltInCategory.OST_StructuralFraming,
    "Structural Rebar":       DB.BuiltInCategory.OST_Rebar,
    "Roofs":                  DB.BuiltInCategory.OST_Roofs,
    "Windows":                DB.BuiltInCategory.OST_Windows,
    "Doors":                  DB.BuiltInCategory.OST_Doors,
    "Electrical": [
        DB.BuiltInCategory.OST_Conduit,
        DB.BuiltInCategory.OST_LightingFixtures,
        DB.BuiltInCategory.OST_LightingDevices,
        DB.BuiltInCategory.OST_ElectricalFixtures,
        DB.BuiltInCategory.OST_ElectricalEquipment,
    ],
    "Plumbing": [
        DB.BuiltInCategory.OST_PlumbingFixtures,
        DB.BuiltInCategory.OST_PipeCurves,
        DB.BuiltInCategory.OST_PipeFitting,
        DB.BuiltInCategory.OST_PipeAccessory,
    ],
    "Painting": VIRTUAL_PAINT,    # <-- handled specially
    "Wall and Floor Finishes": DB.BuiltInCategory.OST_GenericModel,
    "Furniture": [
        DB.BuiltInCategory.OST_Furniture,
        DB.BuiltInCategory.OST_FurnitureSystems,   # added
    ],
    # Not in order, but you keep it in MAP if you use elsewhere:
    "Ceilings": DB.BuiltInCategory.OST_Ceilings,
}

# ---- sanity check to avoid KeyError on typos/mismatches ----
_missing = [c for c in CATEGORY_ORDER if c not in CATEGORY_MAP]
if _missing:
    from pyrevit import forms
    forms.alert("Missing in CATEGORY_MAP:\n\n- " + "\n- ".join(_missing),
                title="Category mapping error")
    raise SystemExit

# --- Category descriptions (unchanged + Painting) ---
CATEGORY_DESCRIPTIONS = {
    "Block Work in Walls": (
        "Concrete block walls, load-bearing or cavity, plastered both sides and painted to BS 8000-3 masonry workmanship standards, "
        "including all mortar, bed-joint reinforcement, movement provision and finishing to BS 5628-2/-3 quality."
    ),
    "Doors": (
        "Timber or engineered doors with hardwood frames, architraves, ironmongery, seals and painting; installed and fitted as per BS 8214."
    ),
    "Windows": (
        "Aluminium sliding or casement windows with glazing, mosquito nets, stays, handles and fixings; installed per BS 6262 (glazing) and BS 6375."
    ),
    "Structural Foundations": (
        "Mass or reinforced concrete footings, hardcore bedding, DPM and formwork, conforming to BS 8000 (earthworks) and BS 8110 (concrete)."
    ),
    "Structural Framing": (
        "Mild steel beams and trusses, welded or bolted, treated with primer/paint to BS 5493 and fabricated per BS 5950."
    ),
    "Structural Columns": (
        "Concrete/steel columns with starter bars, ties and shuttering; concrete to spec per BS 8110-1, steel primed per BS 5493."
    ),
    "Structural Rebar": (
        "High-yield deformed steel bars (BS 4449 B500B), cut, bent, fixed and supported with chairs/spacers, placed per BS 8666 & BS 8110-1."
    ),
    "Roofs": (
        "0.5 mm IBR/IT4 pre-painted roof sheeting fixed to purlins with screws, complete with ridge capping, insulation and flashings, per BS 5534 & BS 8217."
    ),
    "Ceilings": (
        "Particleboard or PVC tongue-and-groove ceilings, fixed or suspended per BS 5306 and manufacturer instructions."
    ),
    "Wall and Floor Finishes": (
        "Tiling and screed finishes and plaster/paint to walls, following BS 5385 (tiling), BS 8203 (screed) and BS 8000 finishing workmanship standards."
    ),
    "Plumbing": (
        "Sanitary appliances (WC pans, cisterns, basins, sinks, urinals) per BS 6465-3, with associated pipework, fittings, joints, valves, traps and accessories per BS 5572 sanitary drainage."
    ),
    "Electrical": (
        "Steel conduits per BS 4568-1, armoured cables/junction boxes per SANS 1507/BS 7671, with lighting fixtures and switchgear as specified."
    ),
    "Painting": (
        "Measured areas from the Revit Paint tool on wall faces (all sides), grouped by material. Rates use the material 'Cost' if present."
    ),
    # "Furniture": "Loose and built-in furniture items as modeled."
}

# --- Unit conversions ---
FT3_TO_M3 = 0.0283168
FT2_TO_M2 = 0.092903
FT_TO_M   = 0.3048

# --- Workbook setup ---
wb = xlsxwriter.Workbook(xlsx_path)
sheet = wb.add_worksheet("BOQ Export")

# --- Title row (bold & left aligned) ---
def _get_project_title():
    pi = revit.doc.ProjectInformation
    pname = None
    p = pi.get_Parameter(DB.BuiltInParameter.PROJECT_NAME) if pi else None
    if p and p.HasValue:
        pname = p.AsString()
    if not pname:
        try:
            import os as _os
            pname = _os.path.splitext(revit.doc.Title)[0]
        except Exception:
            pname = "PROJECT"
    return pname

_title_text = "BILL OF QUANTITIES (BOQ) FOR THE CONSTRUCTION OF {}".format(_get_project_title().upper())
_title_fmt  = wb.add_format({'bold': True, 'font_name': 'Century Gothic', 'font_size': 14, 'align': 'left'})
sheet.merge_range(0, 0, 0, 5, _title_text, _title_fmt)
sheet.freeze_panes(2, 0)

font = 'Century Gothic'
def col_fmt(bold=False, italic=False, underline=False, wrap=False, num_fmt=None):
    fmt = {'valign':'top','font_name':font,'border':1}
    if bold: fmt['bold']=True
    if italic: fmt['italic']=True
    if underline: fmt['underline']=True
    if wrap: fmt['text_wrap']=True
    if num_fmt: fmt['num_format']=num_fmt
    return wb.add_format(fmt)

fmt_header      = col_fmt(bold=True)
fmt_section     = col_fmt(bold=True)
fmt_description = col_fmt(italic=True, underline=True, wrap=True)
fmt_normal      = col_fmt()
fmt_italic      = col_fmt(italic=True, wrap=True)
fmt_money       = col_fmt(num_fmt='#,##0.00')

# --- Columns ---
headers = ["ITEM","DESCRIPTION","UNIT","QTY","RATE (EUR)","AMOUNT (EUR)"]
for c,h in enumerate(headers): sheet.write(1,c,h,fmt_header)
sheet.set_column(1,1,45); sheet.set_column(4,4,12); sheet.set_column(5,5,16)

row = 2
skipped = 0

# Store each category subtotal cell address for COLLECTION & GRAND TOTAL
category_subtotal_cell = {}

# >>> Category numbering counter
cat_counter = 1

# -------- comment helpers (Fix B) --------------------------------------------
def _is_noise(s):
    s = (s or "").strip()
    if not s:
        return True
    s2 = s.replace(".", "").replace(",", "").replace(" ", "")
    if s2.isdigit():
        return True
    if len(s) < 3:
        return True
    return False
# -----------------------------------------------------------------------------

# ---------------------- PAINTING helper (walls; parts & fallback supported) ---
def _gather_wall_painting(doc):
    grouped = {}

    def _add(material_name, rate, area_ft2):
        key = "Paint - {}".format(material_name or "Paint")
        qty_m2 = float(area_ft2) * FT2_TO_M2
        if key not in grouped:
            grouped[key] = {"qty": 0.0, "rate": float(rate or 0.0), "unit": "m²", "comment": ""}
        grouped[key]["qty"] += qty_m2
        if grouped[key]["rate"] == 0.0 and rate:
            grouped[key]["rate"] = float(rate)

    def _rate_from_material(mat):
        try:
            p = mat.LookupParameter(PARAM_COST) if mat else None
            return float(p.AsDouble()) if (p and p.HasValue) else 0.0
        except:
            return 0.0

    def _collect_from_faces(host_elem, faces):
        for f in faces:
            ref = f.Reference
            if not ref:
                continue
            if not doc.IsPainted(host_elem.Id, ref):
                continue
            mid = doc.GetPaintedMaterial(host_elem.Id, ref)
            if mid == DB.ElementId.InvalidElementId:
                continue
            mat = doc.GetElement(mid)
            _add(mat.Name if mat else "Paint", _rate_from_material(mat), f.Area)

    walls = (DB.FilteredElementCollector(doc)
             .OfCategory(DB.BuiltInCategory.OST_Walls)
             .WhereElementIsNotElementType()
             .ToElements())

    opt = DB.Options()
    opt.ComputeReferences = True
    opt.IncludeNonVisibleObjects = False

    for wall in walls:
        try:
            got_any = False
            try:
                for side in (DB.ShellLayerType.Interior, DB.ShellLayerType.Exterior):
                    refs = DB.HostObjectUtils.GetSideFaces(wall, side)
                    if not refs:
                        continue
                    for ref in refs:
                        if not doc.IsPainted(wall.Id, ref):
                            continue
                        gobj = wall.GetGeometryObjectFromReference(ref)
                        face = gobj if isinstance(gobj, DB.Face) else None
                        if not face:
                            continue
                        mid = doc.GetPaintedMaterial(wall.Id, ref)
                        if mid == DB.ElementId.InvalidElementId:
                            continue
                        mat = doc.GetElement(mid)
                        _add(mat.Name if mat else "Paint", _rate_from_material(mat), face.Area)
                        got_any = True
            except:
                pass
            if got_any:
                continue

            try:
                pids = DB.PartUtils.GetAssociatedParts(doc, wall.Id, True, True)
                if pids and pids.Count > 0:
                    for pid in pids:
                        part = doc.GetElement(pid)
                        geom = part.get_Geometry(opt)
                        if not geom:
                            continue
                        for g in geom:
                            if isinstance(g, DB.Solid) and g.Faces:
                                _collect_from_faces(part, list(g.Faces))
                            elif isinstance(g, DB.GeometryInstance):
                                inst = g.GetInstanceGeometry()
                                for gg in inst:
                                    if isinstance(gg, DB.Solid) and gg.Faces:
                                        _collect_from_faces(part, list(gg.Faces))
                    continue
            except:
                pass

            try:
                geom = wall.get_Geometry(opt)
                if geom:
                    for g in geom:
                        if isinstance(g, DB.Solid) and g.Faces:
                            _collect_from_faces(wall, list(g.Faces))
                        elif isinstance(g, DB.GeometryInstance):
                            inst = g.GetInstanceGeometry()
                            for gg in inst:
                                if isinstance(gg, DB.Solid) and gg.Faces:
                                    _collect_from_faces(wall, list(gg.Faces))
            except:
                pass

        except:
            pass

    for v in grouped.values():
        if abs(v["qty"]) < 1e-6:
            v["qty"] = 0.0
    return grouped
# -------------------------------------------------------------------------------

# --- Loop over categories in specified order ---
for cat_name in CATEGORY_ORDER:
    bic = CATEGORY_MAP.get(cat_name)
    if not bic:
        continue

    # -------- VIRTUAL CATEGORY: Painting --------
    if bic is VIRTUAL_PAINT:
        grouped = _gather_wall_painting(revit.doc)

        if grouped:
            sheet.write(row, 0, str(cat_counter), fmt_section)
            sheet.write(row, 1, cat_name.upper(), fmt_section)
            row += 1
            cat_counter += 1

            if cat_name in CATEGORY_DESCRIPTIONS:
                sheet.write(row, 1, CATEGORY_DESCRIPTIONS[cat_name], fmt_description)
                row += 1

            first_item_row = row
            letters = iter(string.ascii_uppercase)

            for name, data in grouped.items():
                sheet.write(row, 0, next(letters), fmt_normal)
                sheet.write(row, 1, name,           fmt_normal)
                sheet.write(row, 2, data["unit"],   fmt_normal)
                sheet.write(row, 3, round(float(data["qty"]), 2), fmt_normal)
                sheet.write(row, 4, round(float(data["rate"]), 2), fmt_money)
                sheet.write_formula(row, 5, "={}*{}".format(
                    xl_rowcol_to_cell(row, 3), xl_rowcol_to_cell(row, 4)), fmt_money)
                row += 1

            last_item_row = row - 1
            sheet.write(row, 1, cat_name.upper() + " TO COLLECTION", fmt_section)
            if last_item_row >= first_item_row:
                sum_range = "F{}:F{}".format(first_item_row + 1, last_item_row + 1)
                sheet.write_formula(row, 5, "=SUM({})".format(sum_range), fmt_money)
            else:
                sheet.write(row, 5, 0, fmt_money)

            category_subtotal_cell[cat_name.upper()] = xl_rowcol_to_cell(row, 5)
            row += 2
        continue
    # -------- END VIRTUAL CATEGORY --------

    # Collect elements for this (real) category
    if isinstance(bic, list):
        elements = []
        for sub in bic:
            elements += (DB.FilteredElementCollector(revit.doc)
                         .OfCategory(sub)
                         .WhereElementIsNotElementType()
                         .ToElements())
    else:
        elements = (DB.FilteredElementCollector(revit.doc)
                    .OfCategory(bic)
                    .WhereElementIsNotElementType()
                    .ToElements())

    # Group by TYPE/Display name (aggregate qty)
    grouped = {}

    for el in elements:
        try:
            # ------------------------ NAME (robust fallbacks) ------------------------
            el_type = revit.doc.GetElement(el.GetTypeId()) if el.GetTypeId() else None
            name = None
            if el_type:
                p_name = el_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                if p_name and p_name.HasValue:
                    name = p_name.AsString()
            if not name:
                p_ft = el.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)
                if p_ft and p_ft.HasValue:
                    name = p_ft.AsValueString()
            if not name:
                name = getattr(el, "Name", None) or (el.Category.Name if el.Category else "Item")

            # ------------------------ RATE (type first, then instance) --------------
            def _get_cost(o):
                if not o: return 0.0
                try:
                    cp = o.LookupParameter(PARAM_COST)
                    if cp and cp.HasValue:
                        return float(cp.AsDouble())
                except:
                    pass
                return 0.0

            rate = _get_cost(el_type)
            if rate == 0.0:
                rate = _get_cost(el)  # instance fallback (useful for Furniture Systems)

            # ------------------------ QTY / UNIT (category rules) -------------------
            qty  = 1.0
            unit = "No."

            if cat_name == "Block Work in Walls":
                prm = el.get_Parameter(DB.BuiltInParameter.HOST_AREA_COMPUTED) or el.LookupParameter("Area")
                if prm and prm.HasValue:
                    qty = prm.AsDouble() * FT2_TO_M2; unit = "m²"
                else:
                    qty = 0.0
            elif cat_name in ("Doors","Windows"):
                qty = 1
            elif cat_name in ("Wall and Floor Finishes","Roofs","Ceilings"):
                prm = el.LookupParameter("Area")
                if prm and prm.HasValue:
                    qty = prm.AsDouble() * FT2_TO_M2; unit = "m²"
            elif cat_name == "Structural Foundations":
                prm = el.LookupParameter("Volume")
                if prm and prm.HasValue:
                    qty = prm.AsDouble() * FT3_TO_M3; unit = "m³"
            elif cat_name == "Structural Framing":
                prm = el.get_Parameter(DB.BuiltInParameter.CURVE_ELEM_LENGTH)
                if prm and prm.HasValue:
                    qty = prm.AsDouble() * FT_TO_M; unit = "m"
            elif cat_name == "Structural Columns":
                mat_prm  = el.LookupParameter("Structural Material")
                mat_elem = revit.doc.GetElement(mat_prm.AsElementId()) if mat_prm else None
                mname    = (mat_elem.Name if mat_elem else "")
                mclass   = (getattr(mat_elem, "MaterialClass", "") if mat_elem else "")
                low      = (mname + " " + mclass).lower()

                vol_prm = el.get_Parameter(DB.BuiltInParameter.HOST_VOLUME_COMPUTED) or el.LookupParameter("Volume")
                len_prm = (el.get_Parameter(DB.BuiltInParameter.CURVE_ELEM_LENGTH) or
                           el.get_Parameter(DB.BuiltInParameter.INSTANCE_LENGTH_PARAM) or
                           el.get_Parameter(DB.BuiltInParameter.COLUMN_HEIGHT) or
                           el.LookupParameter("Length"))

                if "concrete" in low:
                    if vol_prm and vol_prm.HasValue:
                        qty = vol_prm.AsDouble()*FT3_TO_M3; unit = "m3"
                    elif len_prm and len_prm.HasValue:
                        qty = len_prm.AsDouble()*FT_TO_M; unit = "m"
                elif ("steel" in low) or ("metal" in low):
                    if len_prm and len_prm.HasValue:
                        qty = len_prm.AsDouble()*FT_TO_M; unit = "m"
                    elif vol_prm and vol_prm.HasValue:
                        qty = vol_prm.AsDouble()*FT3_TO_M3; unit = "m3"
                else:
                    if vol_prm and vol_prm.HasValue and vol_prm.AsDouble()>0:
                        qty = vol_prm.AsDouble()*FT3_TO_M3; unit = "m3"
                    elif len_prm and len_prm.HasValue:
                        qty = len_prm.AsDouble()*FT_TO_M; unit = "m"

            # ------------------------ COMMENT (Type Comments only; filter noise) ----
            comment = ""
            if el_type:
                tc = el_type.LookupParameter("Type Comments")
                if tc and tc.HasValue:
                    comment = tc.AsString() or ""
            if _is_noise(comment):
                comment = ""
            if comment.strip().lower() == (name or "").strip().lower():
                comment = ""

            # ------------------------ Aggregate -------------------------------------
            if name not in grouped:
                grouped[name] = {"qty": 0.0, "rate": rate, "unit": unit, "comment": comment}
            grouped[name]["qty"] += qty
            if grouped[name]["rate"] == 0.0 and rate:
                grouped[name]["rate"] = rate
            if comment and not grouped[name].get("comment"):
                grouped[name]["comment"] = comment

        except:
            skipped += 1

    if grouped:
        sheet.write(row, 0, str(cat_counter), fmt_section)   # ITEM column
        sheet.write(row, 1, cat_name.upper(), fmt_section)   # DESCRIPTION column
        row += 1
        cat_counter += 1

        if cat_name in CATEGORY_DESCRIPTIONS:
            sheet.write(row, 1, CATEGORY_DESCRIPTIONS[cat_name], fmt_description)
            row += 1

        first_item_row = row

        letters = iter(string.ascii_uppercase)
        for name, data in grouped.items():
            sheet.write(row, 0, next(letters), fmt_normal)
            sheet.write(row, 1, name,           fmt_normal)
            sheet.write(row, 2, data["unit"],   fmt_normal)
            sheet.write(row, 3, round(float(data["qty"]), 2), fmt_normal)
            sheet.write(row, 4, round(float(data["rate"]), 2), fmt_money)
            sheet.write_formula(row, 5, "={}*{}".format(
                xl_rowcol_to_cell(row, 3), xl_rowcol_to_cell(row, 4)), fmt_money)
            row += 1

            # Only Type Comments (already filtered for noise)
            if data.get("comment"):
                sheet.write(row, 1, u"— " + data["comment"], fmt_italic)
                row += 1

        last_item_row = row - 1
        if last_item_row >= first_item_row:
            sum_range = "F{}:F{}".format(first_item_row + 1, last_item_row + 1)
            sheet.write(row, 1, cat_name.upper() + " TO COLLECTION", fmt_section)
            sheet.write_formula(row, 5, "=SUM({})".format(sum_range), fmt_money)
            category_subtotal_cell[cat_name.upper()] = xl_rowcol_to_cell(row, 5)
            row += 2
        else:
            sheet.write(row, 1, cat_name.upper() + " TO COLLECTION", fmt_section)
            sheet.write(row, 5, 0, fmt_money)
            category_subtotal_cell[cat_name.upper()] = xl_rowcol_to_cell(row, 5)
            row += 2

# --- Totals ---
sheet.write(row, 1, "COLLECTION", fmt_section)
row += 1

# Restart numbering under COLLECTION (title unnumbered)
collect_counter = 1
for cname in CATEGORY_ORDER:
    upper = cname.upper()
    cell  = category_subtotal_cell.get(upper)
    if cell:
        sheet.write(row, 0, str(collect_counter), fmt_normal)
        sheet.write(row, 1, upper,              fmt_normal)
        sheet.write_formula(row, 5, "={}".format(cell), fmt_money)
        row += 1
        collect_counter += 1

# GRAND TOTAL (unnumbered, aligned under DESCRIPTION)
sheet.write_blank(row, 0, None, fmt_section)
sheet.write(row, 1, "GRAND TOTAL", fmt_section)
if category_subtotal_cell:
    sum_cells = ",".join(category_subtotal_cell[k.upper()]
                         for k in CATEGORY_ORDER if k.upper() in category_subtotal_cell)
    sheet.write_formula(row, 5, "=SUM({})".format(sum_cells), fmt_money)
else:
    sheet.write(row, 5, 0, fmt_money)

wb.close()
MessageBox.Show("BOQ export complete!\nSaved to Desktop:\n{}\nSkipped: {}".format(xlsx_path, skipped), "✅ XLSX Export")
