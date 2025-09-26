# -*- coding: utf-8 -*-
import os
import string
import xlsxwriter
import clr

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import MessageBox
from pyrevit import revit, DB

# --- Save path ---
desktop = os.path.expanduser("~/Desktop")
xlsx_path = os.path.join(desktop, "BOQ_Export_From_Model.xlsx")

# --- Parameters ---
PARAM_COST = "Cost"
PARAM_TOTAL = "Test_1234"

# --- Ordered Categories ---
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
    "Wall and Floor Finishes"
]

CATEGORY_MAP = {
    "Block Work in Walls": DB.BuiltInCategory.OST_Walls,
    "Doors": DB.BuiltInCategory.OST_Doors,
    "Windows": DB.BuiltInCategory.OST_Windows,
    "Structural Foundations": DB.BuiltInCategory.OST_StructuralFoundation,
    "Structural Framing": DB.BuiltInCategory.OST_StructuralFraming,
    "Structural Columns": DB.BuiltInCategory.OST_StructuralColumns,
    "Structural Rebar": DB.BuiltInCategory.OST_Rebar,
    "Roofs": DB.BuiltInCategory.OST_Roofs,
    "Ceilings": DB.BuiltInCategory.OST_Ceilings,
    "Wall and Floor Finishes": DB.BuiltInCategory.OST_GenericModel,
    "Plumbing": [
        DB.BuiltInCategory.OST_PlumbingFixtures,
        DB.BuiltInCategory.OST_PipeCurves,
        DB.BuiltInCategory.OST_PipeFitting,
        DB.BuiltInCategory.OST_PipeAccessory
    ],
    "Electrical": [
        DB.BuiltInCategory.OST_Conduit,
        DB.BuiltInCategory.OST_LightingFixtures,
        DB.BuiltInCategory.OST_LightingDevices,
        DB.BuiltInCategory.OST_ElectricalFixtures,
        DB.BuiltInCategory.OST_ElectricalEquipment
    ]
}

CATEGORY_DESCRIPTIONS = {
    "Block Work in Walls": (
        "Concrete block walls, load‑bearing or cavity, plastered both sides and painted to BS 8000‑3 masonry workmanship standards, "
        "including all mortar, bed‑joint reinforcement, movement provision and finishing to BS 5628‑2/‑3 quality."
    ),
    "Doors": (
        "Timber or engineered doors with hardwood frames, architraves, ironmongery, seals and painting; installed and fitted as per BS 8214."
    ),
    "Windows": (
        "Aluminium sliding or casement windows with glazing, mosquito nets, stays, handles and fixings; installed per BS 6262 (glazing) and BS 6375."
    ),
    "Structural Foundations": (
        "Mass or reinforced concrete footings, hardcore bedding, DPM and formwork, conforming to BS 8000 (earthworks) and BS 8110 (concrete)."
    ),
    "Structural Framing": (
        "Mild steel beams and trusses, welded or bolted, treated with primer/paint to BS 5493 and fabricated per BS 5950."
    ),
    "Structural Columns": (
        "Concrete/steel columns with starter bars, ties and shuttering; concrete to spec per BS 8110‑1, steel primed per BS 5493."
    ),
    "Structural Rebar": (
        "High‑yield deformed steel bars (BS 4449 B500B), cut, bent, fixed and supported with chairs/spacers, placed per BS 8666 & BS 8110‑1."
    ),
    "Roofs": (
        "0.5 mm IBR/IT4 pre‑painted roof sheeting fixed to purlins with screws, complete with ridge capping, insulation and flashings, per BS 5534 & BS 8217."
    ),
    "Ceilings": (
        "Particleboard or PVC tongue‑and‑groove ceilings, fixed or suspended per BS 5306 and manufacturer instructions."
    ),
    "Wall and Floor Finishes": (
        "Tiling and screed finishes and plaster/paint to walls, following BS 5385 (tiling), BS 8203 (screed) and BS 8000 finishing workmanship standards."
    ),
    "Plumbing": (
        "Sanitary appliances (WC pans, cisterns, basins, sinks, urinals) per BS 6465‑3, with associated pipework, fittings, joints, valves, traps and accessories per BS 5572 sanitary drainage."
    ),
    "Electrical": (
        "Steel conduits per BS 4568‑1, armoured cables/junction boxes per SANS 1507/BS 7671, with lighting fixtures and switchgear as specified."
    )
}

# --- Unit conversions ---
FT3_TO_M3 = 0.0283168
FT2_TO_M2 = 0.092903
FT_TO_M = 0.3048
CONCRETE_NAME = "Concrete - Cast-in-Place Concrete"
STEEL_NAME = "Metal - Steel 43-275"

# --- Workbook setup ---
wb = xlsxwriter.Workbook(xlsx_path)
sheet = wb.add_worksheet("BOQ Export")
sheet.freeze_panes(1, 0)

font = 'Century Gothic'
def col_fmt(bold=False, italic=False, underline=False, wrap=False, num_fmt=None):
    fmt = {'valign':'top','font_name':font,'border':1}
    if bold: fmt['bold']=True
    if italic: fmt['italic']=True
    if underline: fmt['underline']=True
    if wrap: fmt['text_wrap']=True
    if num_fmt: fmt['num_format']=num_fmt
    return wb.add_format(fmt)

fmt_header = col_fmt(bold=True)
fmt_section = col_fmt(bold=True)
fmt_description = col_fmt(italic=True, underline=True, wrap=True)
fmt_normal = col_fmt()
fmt_italic = col_fmt(italic=True, wrap=True)
fmt_money = col_fmt(num_fmt='#,##0.00')

# --- Columns ---
headers = ["ITEM","DESCRIPTION","UNIT","QTY","RATE (ZMW)","AMOUNT (ZMW)"]
for c,h in enumerate(headers): sheet.write(0,c,h,fmt_header)
sheet.set_column(1,1,45); sheet.set_column(4,4,12); sheet.set_column(5,5,16)

row = 1
skipped = 0
category_totals = []

# --- Loop over categories in specified order ---
for cat_name in CATEGORY_ORDER:
    bic = CATEGORY_MAP.get(cat_name)
    if not bic: continue

    elements = []
    if isinstance(bic, list):
        for sub in bic:
            elements += DB.FilteredElementCollector(revit.doc).OfCategory(sub).WhereElementIsNotElementType().ToElements()
    else:
        elements = DB.FilteredElementCollector(revit.doc).OfCategory(bic).WhereElementIsNotElementType().ToElements()

    grouped = {}
    letters = iter(string.ascii_uppercase)

    for el in elements:
        try:
            el_type = revit.doc.GetElement(el.GetTypeId())
            name = el_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            cost_p = el_type.LookupParameter(PARAM_COST)
            tot_p = el.LookupParameter(PARAM_TOTAL)
            if not (cost_p and tot_p): skipped+=1; continue

            rate = cost_p.AsDouble()
            total = tot_p.AsDouble()
            qty = 1.0; unit = "No."

            if cat_name == "Block Work in Walls":
                prm = el.get_Parameter(DB.BuiltInParameter.HOST_AREA_COMPUTED) or el.LookupParameter("Area")
                if prm and prm.HasValue:
                    qty = prm.AsDouble()*FT2_TO_M2
                    unit = "m²"
                else:
                    qty = 0.0
            elif cat_name in ("Doors","Windows"):
                qty = 1
            elif cat_name in ("Wall and Floor Finishes","Roofs","Ceilings"):
                prm = el.LookupParameter("Area")
                if prm and prm.HasValue: qty = prm.AsDouble()*FT2_TO_M2; unit="m²"
            elif cat_name=="Structural Foundations":
                prm = el.LookupParameter("Volume")
                if prm and prm.HasValue: qty = prm.AsDouble()*FT3_TO_M3; unit="m³"
            elif cat_name=="Structural Framing":
                prm = el.get_Parameter(DB.BuiltInParameter.CURVE_ELEM_LENGTH)
                if prm and prm.HasValue: qty = prm.AsDouble()*FT_TO_M; unit="m"
            elif cat_name=="Structural Columns":
                mat_prm = el.LookupParameter("Structural Material")
                mat_elem = revit.doc.GetElement(mat_prm.AsElementId()) if mat_prm else None
                mname = mat_elem.Name if mat_elem else ""
                if mname==CONCRETE_NAME:
                    prm = el.LookupParameter("Volume")
                    if prm and prm.HasValue: qty = prm.AsDouble()*FT3_TO_M3; unit="m³"
                elif mname==STEEL_NAME:
                    prm = el.get_Parameter(DB.BuiltInParameter.CURVE_ELEM_LENGTH)
                    if prm and prm.HasValue: qty = prm.AsDouble()*FT_TO_M; unit="m"
            elif cat_name=="Structural Rebar":
                prm = el.LookupParameter("Total Bar Length")
                if prm and prm.HasValue: qty = prm.AsDouble()*FT_TO_M; unit="m"
            elif cat_name=="Plumbing":
                cid = el.Category.Id.IntegerValue
                if cid in (
                    int(DB.BuiltInCategory.OST_PlumbingFixtures),
                    int(DB.BuiltInCategory.OST_PipeFitting),
                    int(DB.BuiltInCategory.OST_PipeAccessory)
                ):
                    qty = 1; unit="No."
                elif cid == int(DB.BuiltInCategory.OST_PipeCurves):
                    prm = el.get_Parameter(DB.BuiltInParameter.CURVE_ELEM_LENGTH)
                    if prm and prm.HasValue: qty = prm.AsDouble()*FT_TO_M; unit="m"
            elif cat_name=="Electrical":
                prm = el.get_Parameter(DB.BuiltInParameter.CURVE_ELEM_LENGTH)
                if prm and prm.HasValue: qty = prm.AsDouble()*FT_TO_M; unit="m"

            comment = ""
            cp = el_type.LookupParameter("Type Comments")
            if cp and cp.HasValue: comment = cp.AsString()

            grouped.setdefault(name,{"qty":0.0,"rate":rate,"total":0.0,"unit":unit,"comment":comment})
            grouped[name]["qty"] += qty
            grouped[name]["total"] += total

        except:
            skipped +=1

    if grouped:
        sheet.write(row,1,cat_name.upper(),fmt_section); row+=1
        if cat_name in CATEGORY_DESCRIPTIONS:
            sheet.write(row,1,CATEGORY_DESCRIPTIONS[cat_name],fmt_description); row+=1

        subtotal = 0.0
        for name,data in grouped.items():
            sheet.write(row,0,next(letters),fmt_normal)
            sheet.write(row,1,name,fmt_normal)
            sheet.write(row,2,data["unit"],fmt_normal)
            sheet.write(row,3,round(data["qty"],2),fmt_normal)
            sheet.write(row,4,round(data["rate"],2),fmt_money)
            sheet.write(row,5,round(data["total"],2),fmt_money)
            row+=1
            if data["comment"]:
                sheet.write(row,1,data["comment"],fmt_italic); row+=1
            subtotal += data["total"]

        sheet.write(row,1,cat_name.upper()+" TO COLLECTION",fmt_section)
        sheet.write(row,5,round(subtotal,2),fmt_money)
        category_totals.append((cat_name.upper(),round(subtotal,2)))
        row+=2

# --- Totals ---
sheet.write(row,0,"COLLECTION",fmt_section); row+=1
for cname in CATEGORY_ORDER:
    upper = cname.upper()
    for item in category_totals:
        if item[0] == upper:
            sheet.write(row,0,item[0],fmt_normal)
            sheet.write(row,5,item[1],fmt_money)
            row+=1

sheet.write(row,0,"GRAND TOTAL",fmt_section)
sheet.write(row,5,sum(t[1] for t in category_totals),fmt_money)

wb.close()
MessageBox.Show("BOQ export complete!\nSaved to Desktop:\n{}\nSkipped: {}".format(xlsx_path, skipped), "✅ XLSX Export")
