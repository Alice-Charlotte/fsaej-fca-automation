#!/usr/bin/env python3
"""SITFP FCA sheet generator.

Reads the input sheet of SITFP_FCA_auto_calc_v2.xlsx and generates one
official-template sheet per Part / Assembly into a copy of the official
FSAEJ FCA template.

Usage:
  python3 build_fca_sheets.py [input.xlsx] [template.xlsx] [output.xlsx]

Defaults (same folder as this script):
  input    = SITFP_FCA_auto_calc_v2.xlsx
  template = FSAEJ_FCA_Template_2026_v0.xlsx
  output   = SITFP_FCA_generated.xlsx

Cost tables are read from the 参照_* sheets inside the input workbook,
so updating the tables in the workbook is enough (no code change).
Unit costs for parameterised items are computed with the same pattern
formulas as the workbook (patterns A1-A16, B1, F=fixed, X=no price).
"""
import sys, os, re, math
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
INPUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "SITFP_FCA_auto_calc_v2.xlsx")
TEMPLATE = sys.argv[2] if len(sys.argv) > 2 else os.path.join(HERE, "FSAEJ_FCA_Template_2026_v0.xlsx")
OUTPUT = sys.argv[3] if len(sys.argv) > 3 else os.path.join(HERE, "SITFP_FCA_generated.xlsx")

def num_or_formula(v):
    """Return numeric value for a cell that may hold an arithmetic formula string."""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if s.startswith("="):
        expr = s[1:]
        if not re.fullmatch(r"[0-9.+\-*/() ]+", expr):
            raise ValueError(f"Unsupported formula: {s}")
        return float(eval(expr, {}, {}))
    return float(s)

def eval_unit(code, fixed, c1, c2, s1, s2):
    if code == "F":
        return fixed
    if code == "X" or code is None:
        return None
    s1 = s1 or 0.0
    s2 = s2 or 0.0
    if code == "A1": return c1 * s1 + c2
    if code == "A2": return c1 * s1 * s2 + c2
    if code == "A3": return c1 * ((s1 ** 2 * s2)) ** c2
    if code == "A4": return c1 * s1 ** 2 + c2
    if code == "A5": return c1 * s1 ** 2 * s2 + c2
    if code == "A6": return c1 * s1
    if code == "A7": return c1 * s1 * s2 + c2 * s1 * s1
    if code == "A8": return (s1 - c2) / c1
    if code == "A12": return c1 * math.exp(c2 * s1)
    if code == "A13": return c1 * s1 * s2 / c2
    if code == "A14": return (c1 * s1 + c2) * s2
    if code == "A16": return c1 * s1 * s2
    if code == "B1": return c1 / 105154 * s1 ** 2 * s2 * math.sqrt(s2) + (c2 * math.exp(0.319 * s1))
    raise ValueError(f"Unknown pattern {code}")

# ---------- read reference tables + input from the v2 workbook ----------
wb_in = openpyxl.load_workbook(INPUT, data_only=False)

def ref_dict(sheet, name_col=1):
    ws = wb_in[sheet]
    d = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[name_col - 1]:
            d[row[name_col - 1]] = row
    return d

MAT = ref_dict("参照_材料")     # name, cat, fixed, c1, c2, pattern, unit1, unit2, expr, comment
FAS = ref_dict("参照_締結")
PRO = ref_dict("参照_工程")     # name, cat, unitcost, unit, tooling, multtype, comment
TOO = ref_dict("参照_工具")     # tool, process, cost, unit, comment
MUL = ref_dict("参照_倍率")     # name, value, use, comment

ws = wb_in["入力_FCA明細"]
headers = [c.value for c in ws[1]]
idx = {h: i for i, h in enumerate(headers)}
lines = []
for row in ws.iter_rows(min_row=2, values_only=True):
    if not row[idx["種別"]]:
        continue
    lines.append({h: row[i] for h, i in idx.items()})

# group lines by target (P/N Base)
targets = {}
order = []
for ln in lines:
    key = str(ln["P/N Base"])
    if key not in targets:
        targets[key] = {"meta": ln, "lines": []}
        order.append(key)
    targets[key]["lines"].append(ln)

def unit_cost(ln):
    cat, item = ln["区分"], ln["項目名"]
    s1 = num_or_formula(ln["Size1"])
    s2 = num_or_formula(ln["Size2"])
    if cat in ("材料", "締結"):
        ref = (MAT if cat == "材料" else FAS).get(item)
        if ref is None:
            raise ValueError(f"Item not in table: {item}")
        return eval_unit(ref[5], ref[2], ref[3], ref[4], s1, s2), ref[6], ref[7]
    if cat == "工程":
        ref = PRO.get(item)
        if ref is None:
            raise ValueError(f"Process not in table: {item}")
        return float(ref[2]), ref[3], None
    if cat == "工具":
        ref = TOO.get(item)
        if ref is None:
            raise ValueError(f"Tool not in table: {item}")
        return float(ref[2]), ref[3], None
    raise ValueError(f"Unknown category {cat}")

def mult_value(name):
    if not name:
        return None
    ref = MUL.get(name)
    if ref is None or not isinstance(ref[1], (int, float)):
        raise ValueError(f"Multiplier not usable: {name}")
    return float(ref[1])

# ---------- template handling ----------
if not os.path.exists(TEMPLATE):
    import glob as _glob
    cands = [c for c in _glob.glob(os.path.join(HERE, "*.xlsx"))
             if "template" in os.path.basename(c).lower() and "generated" not in os.path.basename(c).lower()]
    if len(cands) == 1:
        print(f"Note: template not found under the default name; using {os.path.basename(cands[0])}")
        TEMPLATE = cands[0]
    else:
        print("ERROR: FCA template file not found.")
        print(f"  Looked for: {TEMPLATE}")
        print("  Put the official template (FSAEJ_FCA_Template_2026_v0.xlsx) in the same folder as this script,")
        print("  or pass its path as the 2nd argument: py build_fca_sheets.py input.xlsx template.xlsx output.xlsx")
        sys.exit(1)
wb = openpyxl.load_workbook(TEMPLATE)
tpl_part = wb["Part_1"]
tpl_asm = wb["Assembly_1"]

def clear_block(ws, r0, r1, c0, c1):
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            ws.cell(row=r, column=c).value = None

def fill_material(ws, rows, r0, rmax):
    if len(rows) > rmax:
        raise ValueError(f"Material lines exceed template capacity ({rmax}): {len(rows)}")
    clear_block(ws, r0, r0 + rmax - 1, 1, 13)
    for i, ln in enumerate(rows):
        r = r0 + i
        uc, u1, u2 = unit_cost(ln)
        ws.cell(row=r, column=1, value=(i + 1) * 10)
        ws.cell(row=r, column=2, value=ln["項目名"])
        ws.cell(row=r, column=3, value=ln["用途"])
        ws.cell(row=r, column=4, value=uc)
        s1 = num_or_formula(ln["Size1"]); s2 = num_or_formula(ln["Size2"])
        if s1 is not None: ws.cell(row=r, column=5, value=s1)
        if u1: ws.cell(row=r, column=6, value=u1)
        if s2 is not None: ws.cell(row=r, column=7, value=s2)
        if u2: ws.cell(row=r, column=8, value=u2)
        if ln["AreaName"]: ws.cell(row=r, column=9, value=ln["AreaName"])
        area = ln["Area"]
        if area is not None and area != "":
            ws.cell(row=r, column=10, value=area if isinstance(area, str) and area.startswith("=") else num_or_formula(area))
        if ln["Length"] is not None: ws.cell(row=r, column=11, value=num_or_formula(ln["Length"]))
        if ln["Density"] is not None: ws.cell(row=r, column=12, value=num_or_formula(ln["Density"]))
        ws.cell(row=r, column=13, value=num_or_formula(ln["数量"]))

def fill_process(ws, rows, r0, rmax, sheet_name):
    if len(rows) > rmax:
        raise ValueError(f"Process lines exceed template capacity ({rmax}): {len(rows)}")
    clear_block(ws, r0, r0 + rmax - 1, 1, 8)
    for r in range(r0, r0 + rmax):
        ws.cell(row=r, column=9).value = f'=IF(H{r}<>"",D{r}*F{r}*H{r},D{r}*F{r})'
    for i, ln in enumerate(rows):
        r = r0 + i
        uc, unit, _ = unit_cost(ln)
        ws.cell(row=r, column=1, value=(i + 1) * 10)
        ws.cell(row=r, column=2, value=ln["項目名"])
        ws.cell(row=r, column=3, value=ln["用途"])
        ws.cell(row=r, column=4, value=uc)
        if unit: ws.cell(row=r, column=5, value=unit)
        ws.cell(row=r, column=6, value=num_or_formula(ln["数量"]))
        if ln["倍率名"]:
            ws.cell(row=r, column=7, value=ln["倍率名"])
            ws.cell(row=r, column=8, value=mult_value(ln["倍率名"]))
        ws.cell(row=r, column=9).value = f'=IF(H{r}<>"",D{r}*F{r}*H{r},D{r}*F{r})'

def fill_fastener(ws, rows, r0, rmax):
    if len(rows) > rmax:
        raise ValueError(f"Fastener lines exceed template capacity ({rmax}): {len(rows)}")
    clear_block(ws, r0, r0 + rmax - 1, 1, 9)
    for i, ln in enumerate(rows):
        r = r0 + i
        uc, u1, u2 = unit_cost(ln)
        ws.cell(row=r, column=1, value=(i + 1) * 10)
        ws.cell(row=r, column=2, value=ln["項目名"])
        ws.cell(row=r, column=3, value=ln["用途"])
        ws.cell(row=r, column=4, value=uc)
        s1 = num_or_formula(ln["Size1"]); s2 = num_or_formula(ln["Size2"])
        if s1 is not None: ws.cell(row=r, column=5, value=s1)
        if u1: ws.cell(row=r, column=6, value=u1)
        if s2 is not None: ws.cell(row=r, column=7, value=s2)
        if u2: ws.cell(row=r, column=8, value=u2)
        ws.cell(row=r, column=9, value=num_or_formula(ln["数量"]))

def fill_tooling(ws, rows, r0, rmax):
    if len(rows) > rmax:
        raise ValueError(f"Tooling lines exceed template capacity ({rmax}): {len(rows)}")
    clear_block(ws, r0, r0 + rmax - 1, 1, 8)
    for r in range(r0, r0 + rmax):
        ws.cell(row=r, column=9).value = f'=IF(G{r}<>"",D{r}*F{r}/G{r}*H{r},"")'
    for i, ln in enumerate(rows):
        r = r0 + i
        uc, unit, _ = unit_cost(ln)
        ws.cell(row=r, column=1, value=(i + 1) * 10)
        ws.cell(row=r, column=2, value=ln["項目名"])
        ws.cell(row=r, column=3, value=ln["用途"])
        ws.cell(row=r, column=4, value=uc)
        if unit: ws.cell(row=r, column=5, value=unit)
        ws.cell(row=r, column=6, value=num_or_formula(ln["数量"]))
        ws.cell(row=r, column=7, value=num_or_formula(ln["PVF"]))
        share = num_or_formula(ln["工具共用部品数"]) or 1.0
        ws.cell(row=r, column=8, value=1.0 / share)

def by_cat(t, cat):
    return [ln for ln in t["lines"] if ln["区分"] == cat]

def sheet_name_for(key, meta):
    base = f"{key}_{meta['対象名']}"
    base = re.sub(r'[\\/*?:\[\]]', '-', base)
    return base[:31]

generated = {}
for key in order:
    t = targets[key]
    meta = t["meta"]
    if meta["種別"] == "Part":
        nws = wb.copy_worksheet(tpl_part)
        name = sheet_name_for(key, meta)
        nws.title = name
        generated[key] = name
        nws["B2"] = meta["システム"]
        nws["B3"] = meta["アセンブリ名"]
        nws["B4"] = meta["対象名"]
        nws["B5"] = meta["P/N Base"]
        nws["B6"] = meta["Suffix"]
        nws["B7"] = None
        nws["N2"] = 1
        fill_material(nws, by_cat(t, "材料"), 10, 8)
        fill_process(nws, by_cat(t, "工程"), 21, 8, name)
        fill_fastener(nws, by_cat(t, "締結"), 32, 8)
        fill_tooling(nws, by_cat(t, "工具"), 43, 2)

for key in order:
    t = targets[key]
    meta = t["meta"]
    if meta["種別"] != "Assembly":
        continue
    nws = wb.copy_worksheet(tpl_asm)
    name = sheet_name_for(key, meta)
    nws.title = name
    generated[key] = name
    nws["B2"] = meta["システム"]
    nws["B3"] = meta["アセンブリ名"]
    nws["B4"] = meta["P/N Base"]
    nws["B5"] = meta["Suffix"]
    nws["B6"] = None
    nws["N2"] = num_or_formula(meta["親あたり個数"]) or 1
    parts = [targets[k] for k in order if targets[k]["meta"]["種別"] == "Part" and targets[k]["meta"]["アセンブリ名"] == meta["アセンブリ名"]]
    if len(parts) > 6:
        raise ValueError(f"Parts in assembly exceed template capacity (6): {len(parts)}")
    clear_block(nws, 9, 14, 1, 5)
    for i, pt in enumerate(parts):
        r = 9 + i
        pm = pt["meta"]
        pkey = str(pm["P/N Base"])
        nws.cell(row=r, column=1, value=(i + 1) * 10)
        nws.cell(row=r, column=2, value=pm["対象名"])
        nws.cell(row=r, column=3, value=pm["P/N Base"])
        nws.cell(row=r, column=4).value = f"='{generated[pkey]}'!N1"
        nws.cell(row=r, column=5, value=num_or_formula(pm["親あたり個数"]) or 1)
    fill_material(nws, by_cat(t, "材料"), 18, 6)
    fill_process(nws, by_cat(t, "工程"), 27, 6, name)
    fill_fastener(nws, by_cat(t, "締結"), 36, 6)
    fill_tooling(nws, by_cat(t, "工具"), 45, 2)

for s in ("Part_1", "Part_2", "Assembly_1"):
    if s in wb.sheetnames:
        del wb[s]

wb.save(OUTPUT)
print("generated:", OUTPUT)
print("sheets:", wb.sheetnames)
