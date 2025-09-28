# -*- coding: utf-8 -*-
import os
import string
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
import clr

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import MessageBox
from pyrevit import revit, DB


# ------------------------------------------------------------------------------
# Save path
# ------------------------------------------------------------------------------
desktop = os.path.expanduser("~/Desktop")
xlsx_path = os.path.join(desktop, "BOQ_Export_From_Model.xlsx")


# ------------------------------------------------------------------------------
# Parameters / constants
# ------------------------------------------------------------------------------
PARAM_COST  = "Cost"        # rate on type / material
PARAM_TOTAL = "Test_1234"   # kept for backward compatibility (not used to write Amount)

# Worksheet tab colors
TAB_COLORS = {
    "COVER":   "#A6A6A6",  # gray
    "BILL1":   "#4472C4",  # blue
    "BILL2":   "#C00000",  # red
    "SUMMARY": "#70AD47",  # green
}

# Ordered Categories (keys must match CATEGORY_MAP exactly)
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
    "Painting",                  # virtual: painted wall faces (all together)
    "Wall and Floor Finishes",
    "Furniture",
]

# Sentinel for virtual categories
VIRTUAL_PAINT = object()

# Categories map
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
    "Painting": VIRTUAL_PAINT,    # handled specially
    "Wall and Floor Finishes": DB.BuiltInCategory.OST_GenericModel,
    "Furniture": [
        DB.BuiltInCategory.OST_Furniture,
        DB.BuiltInCategory.OST_FurnitureSystems,
    ],
    # Not in order, but handy to have available:
    "Ceilings": DB.BuiltInCategory.OST_Ceilings,
}

# Sanity check to avoid KeyError on typos/mismatches
_missing = [c for c in CATEGORY_ORDER if c not in CATEGORY_MAP]
if _missing:
    from pyrevit import forms
    forms.alert("Missing in CATEGORY_MAP:\n\n- " + "\n- ".join(_missing),
                title="Category mapping error")
    raise SystemExit

# Category descriptions
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
}

# Unit conversions
FT3_TO_M3 = 0.0283168
FT2_TO_M2 = 0.092903
FT_TO_M   = 0.3048


# ------------------------------------------------------------------------------
# Workbook & formats
# ------------------------------------------------------------------------------
wb = xlsxwriter.Workbook(xlsx_path, {'constant_memory': True})
try:
    wb.set_calc_on_load()
except AttributeError:
    pass

font = 'Arial Narrow'
def col_fmt(bold=False, italic=False, underline=False, wrap=False, num_fmt=None):
    fmt = {'valign': 'top', 'font_name': font, 'font_size': 12, 'border': 1}
    if bold: fmt['bold'] = True
    if italic: fmt['italic'] = True
    if underline: fmt['underline'] = True
    if wrap: fmt['text_wrap'] = True
    if num_fmt: fmt['num_format'] = num_fmt
    return wb.add_format(fmt)

fmt_header      = col_fmt(bold=True)
fmt_section     = col_fmt(bold=True)
fmt_description = col_fmt(italic=True, underline=True, wrap=True)
fmt_normal      = col_fmt()
fmt_italic      = col_fmt(italic=True, wrap=True)
fmt_money       = col_fmt(num_fmt='#,##0.00')
fmt_title       = wb.add_format({'bold': True, 'font_name': font, 'font_size': 12, 'align':'left'})
fmt_cover_big   = wb.add_format({'bold': True, 'font_name': font, 'font_size': 12, 'align':'center'})

# Additional cover formats
fmt_cover_huge  = wb.add_format({'bold': True, 'font_name': font, 'font_size': 16, 'align': 'center'})
fmt_center      = wb.add_format({'font_name': font, 'font_size': 12, 'align': 'center', 'valign': 'vcenter', 'border': 1})
fmt_text        = wb.add_format({'font_name': font, 'font_size': 12, 'border': 1})
fmt_bold        = wb.add_format({'font_name': font, 'font_size': 12, 'border': 1, 'bold': True})
fmt_wrap        = wb.add_format({'font_name': font, 'font_size': 12, 'border': 1, 'text_wrap': True, 'valign': 'top'})
fmt_percent     = wb.add_format({'font_name': font, 'font_size': 12, 'border': 1, 'num_format': '0.00%'})
fmt_money_right = wb.add_format({'font_name': font, 'font_size': 12, 'border': 1, 'num_format': '#,##0.00', 'align': 'right'})
fmt_noborder    = wb.add_format({'font_name': font, 'font_size': 12})


# ------------------------------------------------------------------------------
# Title and helpers
# ------------------------------------------------------------------------------
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

def _get_project_address():
    pi = revit.doc.ProjectInformation
    addr = None
    p = pi.get_Parameter(DB.BuiltInParameter.PROJECT_ADDRESS) if pi else None
    if p and p.HasValue:
        addr = p.AsString()
    if not addr:
        addr = "PROJECT ADDRESS"
    return addr

TITLE_TEXT = "BILL OF QUANTITIES (BOQ) FOR THE CONSTRUCTION OF {}".format(_get_project_title().upper())

# Excel sheet-name helper (<= 31 chars, remove illegal chars, ensure unique)
def _safe_sheet_name(name, used):
    s = name.replace(u"–", "-").replace(u"—", "-")
    for ch in '[]:*?/\\':
        s = s.replace(ch, "")
    s = s.strip().strip("'")
    s = s[:31]
    base = s
    i = 1
    while s in used:
        suf = "({})".format(i)
        s = (base[:31-len(suf)] + suf)
        i += 1
    used.add(s)
    return s

# Comment noise filter (use Type Comments only; ignore numeric/short)
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

# Safe item label (A..Z then 27→"27", etc.)
def _item_label(idx):
    if idx < 26:
        return string.ascii_uppercase[idx]
    return str(idx + 1)


# ------------------------------------------------------------------------------
# Sheet creation helpers
# ------------------------------------------------------------------------------
def _set_portrait(ws):
    ws.set_paper(9)          # A4
    ws.set_portrait()
    ws.set_margins(left=0.5, right=0.5, top=0.5, bottom=0.8)

def init_bill_sheet(name):
    ws = wb.add_worksheet(name)
    _set_portrait(ws)
    ws.merge_range(0, 0, 0, 5, TITLE_TEXT, fmt_title)
    headers = ["ITEM","DESCRIPTION","UNIT","QTY","RATE (EUR)","AMOUNT (EUR)"]
    for c,h in enumerate(headers):
        ws.write(1, c, h, fmt_header)
    ws.set_column(1,1,45); ws.set_column(4,4,12); ws.set_column(5,5,16)
    ws.freeze_panes(2,0)
    return ws

def init_cover_sheet(name):
    ws = wb.add_worksheet(name)
    _set_portrait(ws)

    # column widths
    ws.set_column("B:D", 50)
    # a few taller rows for spacing
    ws.set_row(8,  28)
    ws.set_row(15, 28)
    ws.set_row(19, 28)
    ws.set_row(21, 24)

    # CONTENT (matches your screenshot)
    ws.merge_range("B9:D9",  "DEPARTMENT OF HOUSING AND INFRASTRUCTURE DEVELOPMENT", fmt_cover_huge)
    ws.merge_range("B15:D15", "BILL OF QUANTITIES", fmt_cover_huge)
    ws.merge_range("B17:D17", "FOR THE", fmt_text)
    ws.merge_range("B19:D19", TITLE_TEXT, fmt_cover_huge)
    ws.merge_range("B21:D21", "AT {}".format(_get_project_address().upper()), fmt_text)

    return ws

def finalize_bill_sheet(ws, row, sheet_cat_order, cat_subtotals):
    # COLLECTION
    ws.write(row, 1, "COLLECTION", fmt_section)
    row += 1
    count = 1
    for cname in sheet_cat_order:
        up = cname.upper()
        cell = cat_subtotals.get(up)
        if cell:
            ws.write(row, 0, str(count), fmt_normal)
            ws.write(row, 1, up, fmt_normal)
            ws.write_formula(row, 5, "={}".format(cell), fmt_money)
            row += 1
            count += 1
    # GRAND TOTAL
    ws.write_blank(row, 0, None, fmt_section)
    ws.write(row, 1, "GRAND TOTAL", fmt_section)
    if cat_subtotals:
        sum_cells = ",".join(cat_subtotals[k.upper()] for k in sheet_cat_order if k.upper() in cat_subtotals)
        ws.write_formula(row, 5, "=SUM({})".format(sum_cells), fmt_money)
    else:
        ws.write(row, 5, 0, fmt_money)
    grand_total_addr = xl_rowcol_to_cell(row, 5)
    return grand_total_addr, row

# Return a reference string (NO leading "=") for use in formulas
def _sheet_ref(name, cell_addr):
    return "'{}'!{}".format(name.replace("'", "''"), cell_addr)


# ------------------------------------------------------------------------------
# Build workbook structure (sanitized names)
# ------------------------------------------------------------------------------
_USED_SHEETS = set()
COVER_NAME   = _safe_sheet_name("COVER", _USED_SHEETS)
BILL1_NAME   = _safe_sheet_name("BILL 1 - SUB & SUPERSTRUCTURE", _USED_SHEETS)
BILL2_NAME   = _safe_sheet_name("BILL 2 - MEP", _USED_SHEETS)
SUMMARY_NAME = _safe_sheet_name("GENERAL SUMMARY", _USED_SHEETS)

cover_ws = init_cover_sheet(COVER_NAME)
cover_ws.set_tab_color(TAB_COLORS["COVER"])

sheets = {
    BILL1_NAME: {"ws": init_bill_sheet(BILL1_NAME), "row": 2, "cat_counter": 1, "cat_subtotals": {}, "order": []},
    BILL2_NAME: {"ws": init_bill_sheet(BILL2_NAME), "row": 2, "cat_counter": 1, "cat_subtotals": {}, "order": []},
}
sheets[BILL1_NAME]["ws"].set_tab_color(TAB_COLORS["BILL1"])
sheets[BILL2_NAME]["ws"].set_tab_color(TAB_COLORS["BILL2"])

# Route categories to bills (default to BILL 1)
BILL_FOR_CATEGORY = {"Electrical": BILL2_NAME, "Plumbing": BILL2_NAME}
def _bill_for(cat):
    return BILL_FOR_CATEGORY.get(cat, BILL1_NAME)


# ------------------------------------------------------------------------------
# Painting helper (walls; parts & fallback supported)
# ------------------------------------------------------------------------------
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
            if not ref: continue
            if not doc.IsPainted(host_elem.Id, ref): continue
            mid = doc.GetPaintedMaterial(host_elem.Id, ref)
            if mid == DB.ElementId.InvalidElementId: continue
            mat = doc.GetElement(mid)
            _add(mat.Name if mat else "Paint", _rate_from_material(mat), f.Area)

    walls = (DB.FilteredElementCollector(doc)
             .OfCategory(DB.BuiltInCategory.OST_Walls)
             .WhereElementIsNotElementType()
             .ToElements())

    opt = DB.Options(); opt.ComputeReferences = True; opt.IncludeNonVisibleObjects = False

    for wall in walls:
        try:
            got_any = False
            try:
                for side in (DB.ShellLayerType.Interior, DB.ShellLayerType.Exterior):
                    refs = DB.HostObjectUtils.GetSideFaces(wall, side) or []
                    for ref in refs:
                        if not doc.IsPainted(wall.Id, ref): continue
                        gobj = wall.GetGeometryObjectFromReference(ref)
                        face = gobj if isinstance(gobj, DB.Face) else None
                        if not face: continue
                        mid = doc.GetPaintedMaterial(wall.Id, ref)
                        if mid == DB.ElementId.InvalidElementId: continue
                        mat = doc.GetElement(mid)
                        _add(mat.Name if mat else "Paint", _rate_from_material(mat), face.Area)
                        got_any = True
            except:
                pass
            if got_any: continue

            try:
                pids = DB.PartUtils.GetAssociatedParts(doc, wall.Id, True, True)
                if pids and pids.Count > 0:
                    for pid in pids:
                        part = doc.GetElement(pid)
                        geom = part.get_Geometry(opt)
                        if not geom: continue
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


# ------------------------------------------------------------------------------
# MAIN: loop over categories and route rows to the right bill
# ------------------------------------------------------------------------------
skipped = 0

for cat_name in CATEGORY_ORDER:
    bill_name = _bill_for(cat_name)
    ctx = sheets[bill_name]
    ws = ctx["ws"]
    row = ctx["row"]
    cat_counter = ctx["cat_counter"]
    cat_subtotals = ctx["cat_subtotals"]

    bic = CATEGORY_MAP.get(cat_name)
    if not bic:
        continue

    # ------ VIRTUAL: Painting ------
    if bic is VIRTUAL_PAINT:
        grouped = _gather_wall_painting(revit.doc)
        if grouped:
            ws.write(row, 0, str(cat_counter), fmt_section)
            ws.write(row, 1, cat_name.upper(), fmt_section); row += 1; cat_counter += 1
            ctx["order"].append(cat_name)

            if cat_name in CATEGORY_DESCRIPTIONS:
                ws.write(row, 1, CATEGORY_DESCRIPTIONS[cat_name], fmt_description)
                row += 1

            first_item_row = row
            item_idx = 0
            for name, data in grouped.items():
                ws.write(row, 0, _item_label(item_idx), fmt_normal)
                ws.write(row, 1, name,           fmt_normal)
                ws.write(row, 2, data["unit"],   fmt_normal)
                ws.write(row, 3, round(float(data["qty"]), 2), fmt_normal)
                ws.write(row, 4, round(float(data["rate"]), 2), fmt_money)
                ws.write_formula(row, 5, "={}*{}".format(
                    xl_rowcol_to_cell(row, 3), xl_rowcol_to_cell(row, 4)), fmt_money)
                row += 1
                item_idx += 1

            last_item_row = row - 1
            ws.write(row, 1, cat_name.upper() + " TO COLLECTION", fmt_section)
            if last_item_row >= first_item_row:
                sum_range = "F{}:F{}".format(first_item_row + 1, last_item_row + 1)
                ws.write_formula(row, 5, "=SUM({})".format(sum_range), fmt_money)
            else:
                ws.write(row, 5, 0, fmt_money)

            cat_subtotals[cat_name.upper()] = xl_rowcol_to_cell(row, 5)
            row += 2

        ctx["row"] = row
        ctx["cat_counter"] = cat_counter
        continue
    # ------ END VIRTUAL ------

    # Collect elements
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

    grouped = {}

    for el in elements:
        try:
            # -------- NAME (robust) --------
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

            # -------- RATE (type first, then instance) --------
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
                rate = _get_cost(el)  # instance fallback useful for Furniture Systems

            # -------- QTY / UNIT (category rules) --------
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
                        qty = vol_prm.AsDouble()*FT3_TO_M3; unit = "m³"
                    elif len_prm and len_prm.HasValue:
                        qty = len_prm.AsDouble()*FT_TO_M; unit = "m"
                elif ("steel" in low) or ("metal" in low):
                    if len_prm and len_prm.HasValue:
                        qty = len_prm.AsDouble()*FT_TO_M; unit = "m"
                    elif vol_prm and vol_prm.HasValue:
                        qty = vol_prm.AsDouble()*FT3_TO_M3; unit = "m³"
                else:
                    if vol_prm and vol_prm.HasValue and vol_prm.AsDouble()>0:
                        qty = vol_prm.AsDouble()*FT3_TO_M3; unit = "m³"
                    elif len_prm and len_prm.HasValue:
                        qty = len_prm.AsDouble()*FT_TO_M; unit = "m"

            # -------- COMMENT (Type Comments only; filter noise) --------
            comment = ""
            if el_type:
                tc = el_type.LookupParameter("Type Comments")
                if tc and tc.HasValue:
                    comment = tc.AsString() or ""
            if _is_noise(comment):
                comment = ""
            if comment.strip().lower() == (name or "").strip().lower():
                comment = ""

            # -------- Aggregate --------
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
        ws.write(row, 0, str(cat_counter), fmt_section)   # ITEM
        ws.write(row, 1, cat_name.upper(), fmt_section)   # DESCRIPTION (section)
        row += 1
        cat_counter += 1
        ctx["order"].append(cat_name)

        if cat_name in CATEGORY_DESCRIPTIONS:
            ws.write(row, 1, CATEGORY_DESCRIPTIONS[cat_name], fmt_description)
            row += 1

        first_item_row = row
        item_idx = 0
        for name, data in grouped.items():
            ws.write(row, 0, _item_label(item_idx), fmt_normal)
            ws.write(row, 1, name,           fmt_normal)
            ws.write(row, 2, data["unit"],   fmt_normal)
            ws.write(row, 3, round(float(data["qty"]), 2), fmt_normal)
            ws.write(row, 4, round(float(data["rate"]), 2), fmt_money)
            ws.write_formula(row, 5, "={}*{}".format(
                xl_rowcol_to_cell(row, 3), xl_rowcol_to_cell(row, 4)), fmt_money)
            row += 1
            item_idx += 1

            if data.get("comment"):
                ws.write(row, 1, data["comment"], fmt_italic)
                row += 1

        last_item_row = row - 1
        ws.write(row, 1, cat_name.upper() + " TO COLLECTION", fmt_section)
        if last_item_row >= first_item_row:
            sum_range = "F{}:F{}".format(first_item_row + 1, last_item_row + 1)
            ws.write_formula(row, 5, "=SUM({})".format(sum_range), fmt_money)
        else:
            ws.write(row, 5, 0, fmt_money)
        ctx["cat_subtotals"][cat_name.upper()] = xl_rowcol_to_cell(row, 5)
        row += 2

    # store back row/counter for this sheet
    ctx["row"] = row
    ctx["cat_counter"] = cat_counter


# ------------------------------------------------------------------------------
# Finalize bills & create GENERAL SUMMARY (portrait, signature pinned at bottom)
# ------------------------------------------------------------------------------
ORDERED_BILLS = [BILL1_NAME, BILL2_NAME]
bill_grand_refs = []

for bill_name in ORDERED_BILLS:
    ctx = sheets[bill_name]
    ws = ctx["ws"]
    grand_addr, _ = finalize_bill_sheet(ws, ctx["row"], ctx["order"], ctx["cat_subtotals"])
    bill_grand_refs.append(_sheet_ref(bill_name, grand_addr))

summary_ws = wb.add_worksheet(SUMMARY_NAME)
_set_portrait(summary_ws)
summary_ws.set_tab_color(TAB_COLORS["SUMMARY"])

summary_ws.set_column(0, 0, 6)    # ITEM
summary_ws.set_column(1, 1, 60)   # DESCRIPTION
summary_ws.set_column(2, 2, 4)    # K
summary_ws.set_column(3, 3, 18)   # AMOUNT

summary_ws.merge_range(0, 0, 0, 3, "GENERAL SUMMARY", fmt_center)
summary_ws.write(1, 0, "ITEM", fmt_header)
summary_ws.write(1, 1, "DESCRIPTION", fmt_header)
summary_ws.write(1, 2, "", fmt_header)
summary_ws.write(1, 3, "AMOUNT (ZMW)", fmt_header)

row = 2
summary_ws.merge_range(row, 1, row, 3, _get_project_title().upper(), fmt_bold)
row += 2

CURRENCY_SYM = "K"

for idx, (bill_name, ref) in enumerate(zip(ORDERED_BILLS, bill_grand_refs), start=1):
    label_tail = bill_name.split(" - ", 1)[-1].upper() if " - " in bill_name else bill_name.upper()
    summary_ws.write(row, 1, "BILL No. {}: {}".format(idx, label_tail), fmt_text)
    summary_ws.write(row, 2, CURRENCY_SYM, fmt_text)
    summary_ws.write_formula(row, 3, "=" + ref, fmt_money_right)
    row += 1

sub1_row = row
summary_ws.write_blank(row, 0, None, fmt_text)
summary_ws.write(row, 1, "Sub total 1", fmt_bold)
summary_ws.write(row, 2, CURRENCY_SYM, fmt_bold)
if bill_grand_refs:
    summary_ws.write_formula(row, 3, "=SUM({})".format(",".join(bill_grand_refs)), fmt_money_right)
else:
    summary_ws.write(row, 3, 0, fmt_money_right)
row += 2

disc_text = ("Should the Contractor desire to make any discount on the above total, "
             "it is to be made here and the amount will be treated as a percentage of "
             "the total as above. The rates inserted by the contractor against the "
             "items throughout this tender will be adjusted accordingly by this "
             "percentage during project execution")
disc_top = row
disc_bottom = row + 5
summary_ws.merge_range(disc_top, 1, disc_bottom, 1, disc_text, fmt_wrap)
summary_ws.write(disc_top, 2, "%", fmt_center)
summary_ws.write(disc_top + 1, 2, 0, fmt_percent)
discount_cell = xl_rowcol_to_cell(disc_top + 1, 2)

row = disc_bottom + 1
sub2_row = row
summary_ws.write_blank(row, 0, None, fmt_text)
summary_ws.write(row, 1, "Sub total 2", fmt_bold)
summary_ws.write(row, 2, CURRENCY_SYM, fmt_bold)
summary_ws.write_formula(row, 3, "={}*(1-{})".format(xl_rowcol_to_cell(sub1_row, 3), discount_cell), fmt_money_right)
row += 1

CONTINGENCY_RATE = 0.05
summary_ws.write(row, 1, "Allow for contingencies @ {}%".format(int(CONTINGENCY_RATE*100)), fmt_text)
summary_ws.write_blank(row, 2, None, fmt_text)
summary_ws.write_formula(row, 3, "={}*{}".format(xl_rowcol_to_cell(sub2_row, 3), CONTINGENCY_RATE), fmt_money_right)
contingency_row = row
row += 1

sub3_row = row
summary_ws.write_blank(row, 0, None, fmt_text)
summary_ws.write(row, 1, "Sub total 3", fmt_bold)
summary_ws.write(row, 2, CURRENCY_SYM, fmt_bold)
summary_ws.write_formula(row, 3, "={}+{}".format(xl_rowcol_to_cell(sub2_row, 3),
                                                 xl_rowcol_to_cell(contingency_row, 3)), fmt_money_right)
row += 1

summary_ws.write(row, 1, "Add VAT OR TOT, whichever is applicable", fmt_text)
summary_ws.write(row, 2, "", fmt_text)
summary_ws.write(row, 3, "Inclusive", fmt_text)
row += 1

summary_ws.write(row, 1, "GRAND TOTAL CARRIED TO FORM OF TENDER", fmt_bold)
summary_ws.write(row, 2, CURRENCY_SYM, fmt_bold)
summary_ws.write_formula(row, 3, "={}".format(xl_rowcol_to_cell(sub3_row, 3)), fmt_money_right)
row += 1

# Signature block pinned to bottom of page 1, with *no borders* above it
FIRST_PAGE_LAST_ROW = 47   # 1-based page-break row
SIG_BLOCK_HEIGHT    = 4
sig_top_row_1based  = FIRST_PAGE_LAST_ROW - SIG_BLOCK_HEIGHT + 1
sig_top_row_0based  = sig_top_row_1based - 1

while row < sig_top_row_0based:
    summary_ws.write_blank(row, 0, None, fmt_noborder)
    summary_ws.write_blank(row, 1, None, fmt_noborder)
    summary_ws.write_blank(row, 2, None, fmt_noborder)
    summary_ws.write_blank(row, 3, None, fmt_noborder)
    row += 1

summary_ws.write(row,   1, "Signature of Contractor .................................................................", fmt_text); row += 1
summary_ws.write(row,   1, "Name of Firm: ..............................................................................", fmt_text); row += 1
summary_ws.write(row,   1, "Address: ...................................................................................", fmt_text); row += 1
summary_ws.write(row,   1, "Date: ......................................................................................", fmt_text); row += 1

summary_ws.set_h_pagebreaks([FIRST_PAGE_LAST_ROW])


# ------------------------------------------------------------------------------
# Close and notify
# ------------------------------------------------------------------------------
wb.close()
MessageBox.Show(
    "BOQ export (multi-sheet) complete!\nSaved to Desktop:\n{}\nSkipped: {}".format(xlsx_path, skipped),
    "✅ XLSX Export"
)
