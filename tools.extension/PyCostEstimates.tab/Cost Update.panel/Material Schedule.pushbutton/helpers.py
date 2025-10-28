# -*- coding: utf-8 -*-
from __future__ import print_function
import os, re, csv

def norm(s):
    return re.sub(r'\s+', ' ', (s or '').strip()).lower()

def find_column(columns, *candidates):
    cols = [c.strip().lower() for c in columns or []]
    for cand in candidates:
        if cand.lower() in cols:
            return cols.index(cand.lower())
    for cand in candidates:
        for i, c in enumerate(cols):
            if cand.lower() in c:
                return i
    return None

def load_cost_folder(folderpath):
    """Merge all CSV price files in a folder into a lookup."""
    result = {}
    if not folderpath or not os.path.isdir(folderpath):
        return result
    files = [f for f in os.listdir(folderpath) if f.lower().endswith(".csv")]
    for fname in files:
        fpath = os.path.join(folderpath, fname)
        try:
            with open(fpath, "r", encoding="utf-8-sig") as fh:
                rdr = csv.reader(fh)
                rows = list(rdr)
            if not rows:
                continue
            header = rows[0]
            name_i = find_column(header, "item", "product description", "material", "description", "name")
            unit_i = find_column(header, "uom", "unit")
            rate_i = find_column(header, "rate", "unit cost", "price", "cost")
            if name_i is None or rate_i is None:
                continue
            for r in rows[1:]:
                try:
                    nm = (r[name_i] or "").strip()
                    if not nm:
                        continue
                    u = (r[unit_i] or "").strip() if unit_i is not None and unit_i < len(r) else ""
                    rr = (r[rate_i] or "0").replace(",", "")
                    try:
                        rate = float(rr)
                    except Exception:
                        rate = 0.0
                    result[norm(nm)] = {"name": nm, "unit": u, "rate": rate, "src": fname}
                except Exception:
                    continue
        except Exception:
            continue
    return result

def load_recipes(csv_path):
    """Return dict: category -> list of rows."""
    from collections import defaultdict
    recipes = defaultdict(list)
    if not csv_path or not os.path.isfile(csv_path):
        return recipes
    with open(csv_path, "r", encoding="utf-8-sig") as fh:
        rdr = csv.DictReader(fh)
        for r in rdr:
            cat  = (r.get("Category") or "").strip()
            patt = (r.get("FamilyOrTypePattern") or "").strip()
            buni = (r.get("BaseUnit") or "").strip()
            mat  = (r.get("Constituent") or "").strip()
            uom  = (r.get("Unit") or "").strip()
            perb = r.get("QtyPerBase") or "0"
            wast = r.get("Waste%") or "0"
            if not (cat and patt and buni and mat and uom):
                continue
            try: perb = float(str(perb).replace(",", ""))
            except: perb = 0.0
            try: wast = float(str(wast).replace("%","").replace(",", ""))
            except: wast = 0.0
            try:
                rgx = re.compile(patt, re.IGNORECASE)
            except:
                rgx = re.compile(re.escape(patt), re.IGNORECASE)
            recipes[cat].append({
                "regex": rgx,
                "base_unit": buni,
                "material": mat,
                "unit": uom,
                "per_base": perb,
                "waste": wast
            })
    return recipes

def price_lookup(cost_map, material_name):
    key = norm(material_name)
    cm = cost_map.get(key)
    if cm:
        return cm.get("rate", 0.0), cm.get("src",""), cm.get("unit","")
    for k, v in cost_map.items():
        if k in key or key in k:
            return v.get("rate", 0.0), v.get("src",""), v.get("unit","")
    return 0.0, "", ""

def safe_float(x, default=0.0):
    try:
        return float(str(x).replace(",", ""))
    except Exception:
        return default
