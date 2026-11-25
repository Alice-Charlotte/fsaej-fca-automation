#!/usr/bin/env python3
"""Build the FCA input workbook (SITFP_FCA_auto_calc_v2.xlsx) from the official
FSAEJ cost tables (tblMaterials / tblProcesses / tblFasteners / tblTooling /
tblProcessMultipliers) placed in the same folder. Requires openpyxl."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os

BASE = os.path.dirname(os.path.abspath(__file__))

def find_file(pattern):
    import glob
    hits = sorted(glob.glob(os.path.join(BASE, pattern)))
    if not hits:
        raise FileNotFoundError(f"Required table file not found in this folder: {pattern}")
    return hits[0]
OUT = os.path.join(BASE, "SITFP_FCA_auto_calc_v2.xlsx")

HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(color="FFFFFF", bold=True)
TITLE_FONT = Font(bold=True, size=14, color="1F4E79")
SUB_FONT = Font(bold=True, color="1F4E79")
THIN = Side(style="thin", color="B0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY = '#,##0.0000'
GREY = Font(color="808080", size=9)

def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = BORDER

def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w

PATTERNS = {
    '[C1]*[Size1]+[C2]': 'A1',
    '[Size1]*[C1]+[C2]': 'A1',
    '([C1]*[Size1]+[C2])': 'A1',
    '[C1]*[Size1]*[Size2]+[C2]': 'A2',
    '([C1]*[Size1]*[Size2]+[C2])': 'A2',
    '[C1]*(([Size1]^2*[Size2]))^([C2])': 'A3',
    '[C1]*[Size1]^2+[C2]': 'A4',
    '[C1]*([Size1]^2)+[C2]': 'A4',
    '([C1]*[Size1]*[Size1]+(C2))': 'A4',
    '[C1]*[Size1]^2*[Size2]+[C2]': 'A5',
    '[C1]*[Size1]': 'A6',
    '[C1]*[Size1]*[Size2]+[C2]*[Size1]*[Size1]': 'A7',
    '([Size1]-[C2])/[C1]': 'A8',
    '[C1]*EXP([C2]*[Size1])': 'A12',
    '([C1]*EXP([C2]*[Size1]))': 'A12',
    '[C1]*[Size1]*[Size2]/[C2]': 'A13',
    '([C1]*[Size1]+[C2])*[Size2]': 'A14',
    '[C1]*[Size1]*[Size2]': 'A16',
}

def pattern_code(price):
    if isinstance(price, (int, float)):
        return ('F', float(price))
    if price is None:
        return ('X', None)
    s = str(price).strip()
    if s.startswith('='):
        try:
            return ('F', float(eval(s[1:], {}, {})))
        except Exception:
            return ('X', None)
    if '105154' in s:
        return ('B1', None)
    if s in PATTERNS:
        return (PATTERNS[s], None)
    return ('X', None)

# ---- load official tables ----
def load_materials():
    wb = openpyxl.load_workbook(find_file("tblMaterials*.xlsx"), data_only=False)
    ws = wb["tblMaterials"]
    rows = []
    for r in ws.iter_rows(min_row=3):
        name = r[1].value
        if not name:
            continue
        code, fixed = pattern_code(r[4].value)
        rows.append([name, r[3].value, fixed, r[7].value, r[8].value, code, r[5].value, r[6].value, r[4].value if not isinstance(r[4].value, (int, float)) else None, r[13].value])
    return rows

def load_fasteners():
    wb = openpyxl.load_workbook(find_file("tblFasteners*.xlsx"), data_only=False)
    ws = wb["tblFasteners"]
    rows = []
    for r in ws.iter_rows(min_row=4):
        name = r[1].value
        if not name:
            continue
        code, fixed = pattern_code(r[3].value)
        rows.append([name, r[2].value, fixed, r[8].value, r[9].value, code, r[5].value, r[7].value, r[3].value if not isinstance(r[3].value, (int, float)) else None, r[14].value])
    return rows

def load_processes():
    wb = openpyxl.load_workbook(find_file("tblProcesses*.xlsx"), data_only=False)
    ws = wb["tblProcesses"]
    rows = []
    for r in ws.iter_rows(min_row=3):
        name = r[1].value
        if not name:
            continue
        rows.append([name, r[4].value, r[2].value, r[3].value, r[5].value, r[7].value, r[9].value])
    return rows

def load_tooling():
    wb = openpyxl.load_workbook(find_file("tblTooling*.xlsx"), data_only=False)
    ws = wb["tblTooling"]
    rows = []
    for r in ws.iter_rows(min_row=3):
        tool = r[2].value
        if not tool:
            continue
        rows.append([tool, r[1].value, r[3].value, r[4].value, r[6].value])
    return rows

def load_multipliers():
    wb = openpyxl.load_workbook(find_file("tblProcessMultipliers*.xlsx"), data_only=False)
    ws = wb["tblProcessMultipliers"]
    rows = []
    for r in ws.iter_rows(min_row=3):
        name = r[1].value
        if not name:
            continue
        rows.append([name, r[2].value, r[3].value, r[4].value])
    return rows

wb = openpyxl.Workbook()

# ---- README ----
ws = wb.active
ws.title = "README"
widths(ws, {"A": 115})
lines = [
    ("SITFP FCA自動計算 v2（公式2026Jコストテーブル版）", TITLE_FONT),
    ("", None),
    ("プロトタイプからの重要な修正：公式テーブルは「単価＝固定値」ではありませんでした。", Font(bold=True, color="C00000")),
    ("材料1,582件のうち約250件、締結の多くは、Size1／Size2と係数C1／C2から単価を計算するパラメータ式です。", None),
    ("この版では式のパターン（18種＋締結13種）を全部分類し、サイズを入れれば単価が自動計算されるようにしました。", None),
    ("固定価格の項目はそのまま自動で入ります。Deleted／AIR／価格なしの項目は「要確認」と表示されます。", None),
    ("", None),
    ("使い方", SUB_FONT),
    ("1. 入力_FCA明細 に1行ずつ入力する。項目名は参照シートの表記と完全一致させること（英語表記そのまま）", None),
    ("2. 材料で断面積を使う行は、Area欄に数値かExcel式（例：=((15.9/2)*(15.9/2)-(15.9/2-1.6)*(15.9/2-1.6))*3.14）を入れる", None),
    ("3. 単価・単位・倍率値・明細コストは自動。黄色の列だけ入力する", None),
    ("4. 集計_部品別／集計_アセンブリで自動集計、チェックで公式テンプレートのサンプル値と一致するか確認できる", None),
    ("5. FCA生成スクリプト（build_fca_sheets.py）で、公式テンプレート形式のシートを部品・アセンブリごとに自動生成できる", None),
    ("", None),
    ("計算式（公式FCAテンプレートの式と同一）", SUB_FONT),
    ("・材料：単価×数量　ただしAreaがある行は 単価×Area×Length×Density×数量", None),
    ("・工程：単価×数量×倍率値（倍率が無ければ×1）", None),
    ("・締結：単価×数量", None),
    ("・工具：単価×数量÷PVF×FracIncld（FracIncld＝1÷その工具を使う部品数）", None),
    ("", None),
    ("注意", SUB_FONT),
    ("・これは2026年版テーブル／テンプレートで作ってあります。2027年大会用が配布されたら、参照シートを差し替える必要があります", None),
    ("・サンプルとして公式テンプレートの Front A-Arm Assembly 一式（部品2点＋アセンブリ）を入力済みです。実部品に入れ替えて使ってください", None),
    ("・倍率『Repeat **』は値が数値ではないため、10回超の繰り返しは倍率名を使わず数量側で調整してください", None),
]
for i, (text, font) in enumerate(lines, start=1):
    c = ws.cell(row=i, column=1, value=text)
    if font:
        c.font = font
    c.alignment = Alignment(wrap_text=True)

# ---- Ref sheets ----
def make_ref(title, headers, rows, money_cols=()):
    s = wb.create_sheet(title)
    for j, h in enumerate(headers, start=1):
        s.cell(row=1, column=j, value=h)
    style_header(s, 1, len(headers))
    for i, r in enumerate(rows, start=2):
        for j, v in enumerate(r, start=1):
            if v is None:
                continue
            cell = s.cell(row=i, column=j, value=v)
            cell.border = BORDER
            if j in money_cols and isinstance(v, (int, float)):
                cell.number_format = MONEY
    s.freeze_panes = "A2"
    return s

mrows = load_materials()
frows = load_fasteners()
prows = load_processes()
trows = load_tooling()
xrows = load_multipliers()

s = make_ref("参照_材料", ["項目名", "カテゴリ", "固定価格", "C1", "C2", "パターン", "Unit1", "Unit2", "式（原文）", "コメント"], mrows, money_cols=(3,))
widths(s, {"A": 46, "B": 18, "C": 11, "D": 10, "E": 10, "F": 9, "G": 8, "H": 8, "I": 44, "J": 60})
s = make_ref("参照_締結", ["項目名", "カテゴリ", "固定価格", "C1", "C2", "パターン", "Unit1", "Unit2", "式（原文）", "コメント"], frows, money_cols=(3,))
widths(s, {"A": 46, "B": 16, "C": 11, "D": 10, "E": 10, "F": 9, "G": 8, "H": 8, "I": 52, "J": 52})
s = make_ref("参照_工程", ["項目名", "カテゴリ", "単価", "Unit", "ToolingRequired", "MultiplierType", "コメント"], prows, money_cols=(3,))
widths(s, {"A": 46, "B": 16, "C": 11, "D": 9, "E": 13, "F": 26, "G": 60})
s = make_ref("参照_工具", ["Tool名", "Process", "単価", "Unit", "コメント"], trows, money_cols=(3,))
widths(s, {"A": 30, "B": 28, "C": 11, "D": 9, "E": 70})
s = make_ref("参照_倍率", ["倍率名", "倍率値", "Use", "コメント"], xrows)
widths(s, {"A": 36, "B": 10, "C": 20, "D": 44})

# ---- Input ----
inp = wb.create_sheet("入力_FCA明細")
headers = ["種別", "システム", "アセンブリ名", "対象名", "P/N Base", "Suffix", "親あたり個数", "区分", "項目名", "用途",
           "Size1", "Size2", "数量", "倍率名", "Area", "Length", "Density", "AreaName", "PVF", "工具共用部品数",
           "単価（自動）", "Unit1（自動）", "Unit2（自動）", "倍率値（自動）", "明細コスト（自動）",
           "C1（自動）", "C2（自動）", "パターン（自動）", "固定価格（自動）"]
for j, h in enumerate(headers, start=1):
    inp.cell(row=1, column=j, value=h)
style_header(inp, 1, len(headers))
widths(inp, {"A": 9, "B": 12, "C": 24, "D": 22, "E": 10, "F": 8, "G": 10, "H": 7, "I": 30, "J": 30,
             "K": 8, "L": 8, "M": 9, "N": 16, "O": 14, "P": 9, "Q": 11, "R": 12, "S": 8, "T": 11,
             "U": 12, "V": 8, "W": 8, "X": 10, "Y": 13, "Z": 9, "AA": 9, "AB": 10, "AC": 11})
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")

ASM = "Front A-Arm Assembly - Upper"
SYS = "Suspension"
sample = [
    # type, sys, asm, target, pn, suffix, perparent, cat, item, use, s1, s2, qty, mult, area, length, density, areaname, pvf, toolshare
    ("Part", SYS, ASM, "Front A-ARM - Upper", 70901, "AA", 1, "材料", "Steel, Mild", "For Arm Rods - A φ15.9 (t=1.6 L=420)", None, None, 2, None, "=((15.9/2)*(15.9/2)-(15.9/2-1.6)*(15.9/2-1.6))*3.14", 420, 0.00000785, "Round pipe", None, None),
    ("Part", SYS, ASM, "Front A-ARM - Upper", 70901, "AA", 1, "材料", "Steel, Mild", "For Arm Rod - B φ15.9(t=1.2 L=480)", None, None, 1, None, "=((15.9/2)*(15.9/2)-(15.9/2-1.2)*(15.9/2-1.2))*3.14", 480, 0.00000785, "Round pipe", None, None),
    ("Part", SYS, ASM, "Front A-ARM - Upper", 70901, "AA", 1, "材料", "Steel, Mild", "For Rod Ends Housing φ19 (t=2.0 L =20)", None, None, 3, None, "=((19/2)*(19/2)-(19/2-2)*(19/2-2))*3.14", 20, 0.00000785, "Round pipe", None, None),
    ("Part", SYS, ASM, "Front A-ARM - Upper", 70901, "AA", 1, "材料", "Paint", "A-ARM Coating", None, None, 0.07, None, None, None, None, None, None, None),
    ("Part", SYS, ASM, "Front A-ARM - Upper", 70901, "AA", 1, "工程", "Tube cut", "For Arm Rods", None, None, "=15.9/10", "Repeat 3", None, None, None, None, None, None),
    ("Part", SYS, ASM, "Front A-ARM - Upper", 70901, "AA", 1, "工程", "Tube cut", "For Rod Ends Housing", None, None, "=19/10", "Repeat 3", None, None, None, None, None, None),
    ("Part", SYS, ASM, "Front A-ARM - Upper", 70901, "AA", 1, "工程", "Tube end preparation for welding", "For Arm Rod to Rod Ends Housing", None, None, 1, "Repeat 6", None, None, None, None, None, None),
    ("Part", SYS, ASM, "Front A-ARM - Upper", 70901, "AA", 1, "工程", "Weld - Round Tubing", "Arm Rods to Rod Ends Housing", None, None, "=15.9/10", "Repeat 6", None, None, None, None, None, None),
    ("Part", SYS, ASM, "Front A-ARM - Upper", 70901, "AA", 1, "工程", "Aerosol Apply", "A-ARM Coating", None, None, 0.07, None, None, None, None, None, None, None),
    ("Part", SYS, ASM, "Front A-ARM - Upper", 70901, "AA", 1, "工具", "Welding Fixture", "Arm Rods to Rod Ends Housing", None, None, 6, None, None, None, None, None, 3000, 2),
    ("Part", SYS, ASM, "Rod End - C", 70902, "AA", 3, "材料", "Rod End, Suspension", "A-Arm ends", 15, None, 1, None, None, None, None, None, None, None),
    ("Assembly", SYS, ASM, ASM, "A7090", "AA", 2, "工程", "Assemble, 1 kg, Interference", "Rod Ends assemble to A-ARM", None, None, 3, None, None, None, None, None, None, None),
]
LAST = 201
for r in range(2, LAST + 1):
    if r - 2 < len(sample):
        vals = sample[r - 2]
        for j, v in enumerate(vals, start=1):
            if v is not None:
                inp.cell(row=r, column=j, value=v)
    # helpers Z C1, AA C2, AB pattern, AC fixed
    inp.cell(row=r, column=26).value = f'=IF($A{r}="","",IF($H{r}="材料",IFERROR(INDEX(参照_材料!$D:$D,MATCH($I{r},参照_材料!$A:$A,0)),""),IF($H{r}="締結",IFERROR(INDEX(参照_締結!$D:$D,MATCH($I{r},参照_締結!$A:$A,0)),""),"")))'
    inp.cell(row=r, column=27).value = f'=IF($A{r}="","",IF($H{r}="材料",IFERROR(INDEX(参照_材料!$E:$E,MATCH($I{r},参照_材料!$A:$A,0)),""),IF($H{r}="締結",IFERROR(INDEX(参照_締結!$E:$E,MATCH($I{r},参照_締結!$A:$A,0)),""),"")))'
    inp.cell(row=r, column=28).value = f'=IF($A{r}="","",IF($H{r}="材料",IFERROR(INDEX(参照_材料!$F:$F,MATCH($I{r},参照_材料!$A:$A,0)),"X"),IF($H{r}="締結",IFERROR(INDEX(参照_締結!$F:$F,MATCH($I{r},参照_締結!$A:$A,0)),"X"),"F")))'
    inp.cell(row=r, column=29).value = f'=IF($A{r}="","",IF($H{r}="材料",IFERROR(INDEX(参照_材料!$C:$C,MATCH($I{r},参照_材料!$A:$A,0)),""),IF($H{r}="締結",IFERROR(INDEX(参照_締結!$C:$C,MATCH($I{r},参照_締結!$A:$A,0)),""),IF($H{r}="工程",IFERROR(INDEX(参照_工程!$C:$C,MATCH($I{r},参照_工程!$A:$A,0)),""),IF($H{r}="工具",IFERROR(INDEX(参照_工具!$C:$C,MATCH($I{r},参照_工具!$A:$A,0)),""),"")))))'
    # U unit cost
    inp.cell(row=r, column=21).value = (
        f'=IF($A{r}="","",IF($AB{r}="X","要確認",IF($AB{r}="F",$AC{r},'
        f'IFERROR(IF($AB{r}="A1",$Z{r}*$K{r}+$AA{r},IF($AB{r}="A2",$Z{r}*$K{r}*$L{r}+$AA{r},'
        f'IF($AB{r}="A3",$Z{r}*(($K{r}^2*$L{r}))^($AA{r}),IF($AB{r}="A4",$Z{r}*$K{r}^2+$AA{r},'
        f'IF($AB{r}="A5",$Z{r}*$K{r}^2*$L{r}+$AA{r},IF($AB{r}="A6",$Z{r}*$K{r},'
        f'IF($AB{r}="A7",$Z{r}*$K{r}*$L{r}+$AA{r}*$K{r}*$K{r},IF($AB{r}="A8",($K{r}-$AA{r})/$Z{r},'
        f'IF($AB{r}="A12",$Z{r}*EXP($AA{r}*$K{r}),IF($AB{r}="A13",$Z{r}*$K{r}*$L{r}/$AA{r},'
        f'IF($AB{r}="A14",($Z{r}*$K{r}+$AA{r})*$L{r},IF($AB{r}="A16",$Z{r}*$K{r}*$L{r},'
        f'IF($AB{r}="B1",$Z{r}/105154*$K{r}^2*$L{r}*SQRT($L{r})+($AA{r}*EXP(0.319*$K{r})),""))))))))))))),"要確認")))'
    )
    # V Unit1, W Unit2
    inp.cell(row=r, column=22).value = f'=IF($A{r}="","",IF($H{r}="材料",IFERROR(INDEX(参照_材料!$G:$G,MATCH($I{r},参照_材料!$A:$A,0)),""),IF($H{r}="締結",IFERROR(INDEX(参照_締結!$G:$G,MATCH($I{r},参照_締結!$A:$A,0)),""),IF($H{r}="工程",IFERROR(INDEX(参照_工程!$D:$D,MATCH($I{r},参照_工程!$A:$A,0)),""),IF($H{r}="工具",IFERROR(INDEX(参照_工具!$D:$D,MATCH($I{r},参照_工具!$A:$A,0)),""),"")))))'
    inp.cell(row=r, column=23).value = f'=IF($A{r}="","",IF($H{r}="材料",IFERROR(INDEX(参照_材料!$H:$H,MATCH($I{r},参照_材料!$A:$A,0)),""),IF($H{r}="締結",IFERROR(INDEX(参照_締結!$H:$H,MATCH($I{r},参照_締結!$A:$A,0)),""),"")))'
    # X multiplier value
    inp.cell(row=r, column=24).value = f'=IF($A{r}="","",IF($N{r}="", "",IFERROR(INDEX(参照_倍率!$B:$B,MATCH($N{r},参照_倍率!$A:$A,0)),"要確認")))'
    # Y line cost (same formulas as official template)
    inp.cell(row=r, column=25).value = (
        f'=IF(OR($A{r}="",$U{r}="要確認",$U{r}="",$X{r}="要確認"),"",'
        f'IF($H{r}="材料",IF($O{r}="",$U{r}*$M{r},$U{r}*$O{r}*$P{r}*$Q{r}*$M{r}),'
        f'IF($H{r}="工程",$U{r}*$M{r}*IF($X{r}="",1,$X{r}),'
        f'IF($H{r}="締結",$U{r}*$M{r},'
        f'IF($H{r}="工具",IF($S{r}="","要確認",$U{r}*$M{r}/$S{r}*(1/IF($T{r}="",1,$T{r}))),"")))))'
    )
    for j in range(1, 30):
        c = inp.cell(row=r, column=j)
        c.border = BORDER
        if j <= 20:
            c.fill = INPUT_FILL
        if j in (21, 25):
            c.number_format = MONEY
        if j >= 26:
            c.font = GREY
inp.freeze_panes = "A2"

# ---- Parts rollup ----
parts = wb.create_sheet("集計_部品別")
ph = ["P/N Base", "種別", "システム", "アセンブリ名", "対象名", "親あたり個数", "材料費", "工程費", "締結費", "工具費", "単価（対象1つ）", "拡張（単価×親あたり）"]
for j, h in enumerate(ph, start=1):
    parts.cell(row=1, column=j, value=h)
style_header(parts, 1, len(ph))
widths(parts, {"A": 11, "B": 9, "C": 12, "D": 24, "E": 22, "F": 11, "G": 11, "H": 11, "I": 11, "J": 11, "K": 14, "L": 15})
targets = [(70901, "Part", SYS, ASM, "Front A-ARM - Upper", 1), (70902, "Part", SYS, ASM, "Rod End - C", 3), ("A7090", "Assembly", SYS, ASM, ASM, 2)]
PLAST = 61
for r in range(2, PLAST + 1):
    if r - 2 < len(targets):
        for j, v in enumerate(targets[r - 2], start=1):
            parts.cell(row=r, column=j, value=v)
    for cat_col, cat in [(7, "材料"), (8, "工程"), (9, "締結"), (10, "工具")]:
        parts.cell(row=r, column=cat_col).value = f'=IF($A{r}="","",SUMIFS(入力_FCA明細!$Y:$Y,入力_FCA明細!$E:$E,$A{r},入力_FCA明細!$H:$H,"{cat}"))'
        parts.cell(row=r, column=cat_col).number_format = MONEY
    parts.cell(row=r, column=11).value = f'=IF($A{r}="","",SUM(G{r}:J{r}))'
    parts.cell(row=r, column=11).number_format = MONEY
    parts.cell(row=r, column=12).value = f'=IF($A{r}="","",$K{r}*$F{r})'
    parts.cell(row=r, column=12).number_format = MONEY
    for j in range(1, 13):
        parts.cell(row=r, column=j).border = BORDER
parts.freeze_panes = "A2"

# ---- Assembly rollup ----
asm = wb.create_sheet("集計_アセンブリ")
for j, h in enumerate(["アセンブリ名", "部品合計", "アセンブリ自体の明細合計", "アセンブリ単価", "車両あたり個数", "車両換算額"], start=1):
    asm.cell(row=1, column=j, value=h)
style_header(asm, 1, 6)
widths(asm, {"A": 30, "B": 13, "C": 16, "D": 13, "E": 12, "F": 14})
asm.cell(row=2, column=1, value=ASM)
asm.cell(row=2, column=2).value = '=SUMIFS(集計_部品別!$L:$L,集計_部品別!$D:$D,$A2,集計_部品別!$B:$B,"Part")'
asm.cell(row=2, column=3).value = '=SUMIFS(集計_部品別!$K:$K,集計_部品別!$D:$D,$A2,集計_部品別!$B:$B,"Assembly")'
asm.cell(row=2, column=4).value = '=B2+C2'
asm.cell(row=2, column=5).value = f'=IFERROR(INDEX(集計_部品別!$F:$F,MATCH($A2,集計_部品別!$E:$E,0)),"")'
asm.cell(row=2, column=6).value = '=D2*E2'
for j in range(1, 7):
    c = asm.cell(row=2, column=j)
    c.border = BORDER
    if j in (2, 3, 4, 6):
        c.number_format = MONEY
asm.cell(row=4, column=1, value="※アセンブリが増えたら行を追加し、B〜Fの式をコピーしてください。").font = Font(italic=True, size=9)

# ---- Check ----
chk = wb.create_sheet("チェック")
widths(chk, {"A": 44, "B": 16, "C": 16, "D": 44})
for j, h in enumerate(["チェック項目", "このブックの値", "公式テンプレートのサンプル値", "意味"], start=1):
    chk.cell(row=1, column=j, value=h)
style_header(chk, 1, 4)
rows_chk = [
    ("Part_1（P/N 70901）単価", '=IFERROR(INDEX(集計_部品別!$K:$K,MATCH(70901,集計_部品別!$A:$A,0)),"")', 12.9118335756, "公式テンプレート Part_1 の Part Cost と一致すること"),
    ("Part_2（P/N 70902）単価", '=IFERROR(INDEX(集計_部品別!$K:$K,MATCH(70902,集計_部品別!$A:$A,0)),"")', 16.25, "公式テンプレート Part_2 の Part Cost と一致すること"),
    ("Assembly_1 単価", "=集計_アセンブリ!D2", 62.2243335756, "公式テンプレート Assembly_1 の Asm Cost と一致すること"),
    ("Assembly_1 車両換算額", "=集計_アセンブリ!F2", 124.4486671512, "公式テンプレート Assembly_1 の Extended Cost と一致すること"),
    ("差額合計（0であること）", "=SUMPRODUCT(ABS(B2:B5-C2:C5))", 0, "サンプルを入れ替えた後は、この行は消して実データ用チェックに使ってください"),
    ("単価が要確認の行数（0であること）", f'=COUNTIF(入力_FCA明細!U2:U{LAST},"要確認")', 0, "項目名のタイプミス／Deleted・AIR等の価格なし項目"),
    ("倍率が要確認の行数（0であること）", f'=COUNTIF(入力_FCA明細!X2:X{LAST},"要確認")', 0, "倍率名が参照_倍率に無い"),
    ("明細コストが要確認の行数（0であること）", f'=COUNTIF(入力_FCA明細!Y2:Y{LAST},"要確認")', 0, "工具行でPVFが未入力など"),
]
for i, (label, formula, bench, note) in enumerate(rows_chk, start=2):
    chk.cell(row=i, column=1, value=label).border = BORDER
    c = chk.cell(row=i, column=2, value=formula)
    c.border = BORDER
    c.number_format = MONEY
    c3 = chk.cell(row=i, column=3, value=bench)
    c3.border = BORDER
    c3.number_format = MONEY
    chk.cell(row=i, column=4, value=note).border = BORDER

wb.save(OUT)
print("saved", OUT)
print("materials", len(mrows), "fasteners", len(frows), "processes", len(prows), "tooling", len(trows), "multipliers", len(xrows))
from collections import Counter
print("material patterns:", Counter(r[5] for r in mrows))
print("fastener patterns:", Counter(r[5] for r in frows))
