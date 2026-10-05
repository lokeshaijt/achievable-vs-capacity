"""
Capacity Vs Production (Achievable Vs Produced) — report generator
===================================================================
Production Register  ->  container quantities  ->  one report workbook in the finalized format.

Workbook produced
-----------------
 Capacity Vs Production   (visible)  the report: capacity, achieved containers, %, Blend and Brand
                                     changeover, week by week for every machine line
 Source Calculations      (hidden)   how every capacity and every achieved figure is derived
 Changeover Calc / Changeover Detail / CO Routing / FG Blend Map   (hidden)
                                     the changeover calculation (see changeover.py)
"""

import datetime
import io
import os
import re
from collections import defaultdict

import openpyxl
from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import changeover

# ── Finalized format ─────────────────────────────────────────────────────────
SHEET_NAME = "Capacity Vs Production"
SOURCE_SHEET = "Source Calculations"
CAPACITY_HEADER = "Capacity(In\nFCL)"
EFFICIENCY = 0.9          # the capacity column uses 90% for every line

COL_WIDTHS = {"A": 9.42578125, "B": 9.140625, "C": 23.7109375}   # D .. last column: WEEK_COL_WIDTH
WEEK_COL_WIDTH = 7.42578125
HELPER_WIDTH = 9.0

# ── Layout: region, sub-group, machine line, machines, shifts ────────────────
# (machines / shifts are the yellow input cells of the report)
LAYOUT = [
    ("AFRICA",    "DC",         "C250 AF",                  2,   2),
    (None,        None,         "CONSTANTA ENV",            32,  2),
    (None,        None,         "PERFECTA ENV",             3,   2),
    (None,        None,         "MD20 ENV",                 2,   2),
    ("__TOTAL__", None,         None,                       None, None),
    (None,        "SC",         "MAISA TAGLESS",            7,   1),
    ("__TOTAL__", None,         None,                       None, None),
    (None,        "PREMIX",     "INSTANT COFFEE STICK PACK", 1,  1),
    (None,        None,         "Instant Tea",              1,   1),
    (None,        None,         "PEARL PACK-PREMIX",        2,   1),
    (None,        None,         "VIKING-MULTITRACK",        1,   1),
    ("__TOTAL__", None,         None,                       None, None),
    ("AUSTRALIA", "DC TAG",     "CONSTANTA TAG (D)",        5,   1),
    ("__TOTAL__", None,         None,                       None, None),
    ("EUROPE",    "DC SF CRIMP", "C250",                     3,   2),
    (None,        None,         "VARIETIES PACK",           1,   2),
    (None,        None,         "MD20 ENV GARANT",          2,   2),
    ("__TOTAL__", None,         None,                       None, None),
    (None,        "DC TAG",     "PERFECTA  K 45 TAG",       1,   2),
    (None,        None,         "MD20 4GM TAG",             1,   1),
    ("__TOTAL__", None,         None,                       None, None),
    (None,        "SC",         "UNIVERSAL POT BAG",        1,   1),
    (None,        None,         "UNIVERSAL TWIN BAG",       1,   1),
    ("__TOTAL__", None,         None,                       None, None),
    (None,        "MISC",       "C250F3 HS",                1,   0),
    (None,        None,         "MD20 HS",                  2,   2),
    ("__TOTAL__", None,         None,                       None, None),
    ("RUSSIA",    "DC TAG",     "CONSTANTA TAG",            3,   1),
    ("__TOTAL__", None,         None,                       None, None),
    ("USA",       "TREE HOUSE", "BREW MAGIC",               1,   1),
    (None,        None,         "C250-A",                   3,   2),
    (None,        None,         "MAISA ENV",                2,   1),
    (None,        None,         "PERFACTA TAG",             2,   2),
    (None,        None,         "CONSTANTA FAMILY SIZE",    2,   2),
    (None,        "HF Co",      "Pearl Pack-FFS POUCH",     1.5, 1),
    (None,        None,         "C250F3 ENV",               1,   0),
    ("__TOTAL__", None,         None,                       None, None),
]

# ── Capacity factors per machine line (New capacity — Arul Sir) ──────────────
MC_FACTORS = {
    "BREW MAGIC":               {"rate": 8,   "shiftmin": 540, "tbgs_cfc": 32,   "cfc_cntr": 3600},
    "C250":                     {"rate": 250, "shiftmin": 524, "tbgs_cfc": 480,  "cfc_cntr": 3600},
    "C250 AF":                  {"rate": 250, "shiftmin": 524, "tbgs_cfc": 1200, "cfc_cntr": 3340},
    "C250-A":                   {"rate": 250, "shiftmin": 524, "tbgs_cfc": 480,  "cfc_cntr": 3600},
    "C250F3 ENV":               {"rate": 250, "shiftmin": 520, "tbgs_cfc": 240,  "cfc_cntr": 7200},
    "C250F3 HS":                {"rate": 250, "shiftmin": 522, "tbgs_cfc": 240,  "cfc_cntr": 7200},
    "CONSTANTA ENV":            {"rate": 125, "shiftmin": 553, "tbgs_cfc": 1125, "cfc_cntr": 3340},
    "CONSTANTA FAMILY SIZE":    {"rate": 95,  "shiftmin": 540, "tbgs_cfc": 144,  "cfc_cntr": 5720},
    "CONSTANTA TAG":            {"rate": 120, "shiftmin": 543, "tbgs_cfc": 1200, "cfc_cntr": 2688},
    "CONSTANTA TAG (D)":        {"rate": 120, "shiftmin": 543, "tbgs_cfc": 1200, "cfc_cntr": 2688},
    "INSTANT COFFEE STICK PACK": {"rate": 40, "shiftmin": 540, "tbgs_cfc": 1350, "cfc_cntr": 3000},
    "Instant Tea":              {"rate": 14,  "shiftmin": 550, "tbgs_cfc": 60,   "cfc_cntr": 15120},
    "MAISA ENV":                {"rate": 100, "shiftmin": 544, "tbgs_cfc": 1200, "cfc_cntr": 1656},
    "MAISA TAGLESS":            {"rate": 120, "shiftmin": 555, "tbgs_cfc": 2400, "cfc_cntr": 2800},
    "MD20 4GM TAG":             {"rate": 125, "shiftmin": 548, "tbgs_cfc": 480,  "cfc_cntr": 4032},
    "MD20 ENV":                 {"rate": 185, "shiftmin": 528, "tbgs_cfc": 1200, "cfc_cntr": 3340},
    "MD20 ENV GARANT":          {"rate": 185, "shiftmin": 480, "tbgs_cfc": 300,  "cfc_cntr": 3250},
    "MD20 HS":                  {"rate": 170, "shiftmin": 526, "tbgs_cfc": 300,  "cfc_cntr": 6500},
    "Pearl Pack-FFS POUCH":     {"rate": 24,  "shiftmin": 550, "tbgs_cfc": 24,   "cfc_cntr": 1200},
    "PEARL PACK-PREMIX":        {"rate": 30,  "shiftmin": 550, "tbgs_cfc": 120,  "cfc_cntr": 6000},
    "PERFACTA TAG":             {"rate": 300, "shiftmin": 525, "tbgs_cfc": 600,  "cfc_cntr": 4620},
    "PERFECTA  K 45 TAG":       {"rate": 250, "shiftmin": 525, "tbgs_cfc": 600,  "cfc_cntr": 5040},
    "PERFECTA ENV":             {"rate": 280, "shiftmin": 539, "tbgs_cfc": 1200, "cfc_cntr": 3340},
    "UNIVERSAL POT BAG":        {"rate": 272, "shiftmin": 540, "tbgs_cfc": 960,  "cfc_cntr": 3220},
    "UNIVERSAL TWIN BAG":       {"rate": 420, "shiftmin": 541, "tbgs_cfc": 960,  "cfc_cntr": 3220},
    "VARIETIES PACK":           {"rate": 250, "shiftmin": 524, "tbgs_cfc": 480,  "cfc_cntr": 3600},
    "VIKING-MULTITRACK":        {"rate": 56,  "shiftmin": 550, "tbgs_cfc": 60,   "cfc_cntr": 15120},
}

# ── Routing map: work-center name -> machine line ────────────────────────────
ROUTING_MAP = {
    "BREW MAGIC":           "BREW MAGIC",
    "C250":                 "C250",
    "C250 AF":              "C250 AF",
    "C250-A":               "C250-A",
    "C250F3 ENV":           "C250F3 ENV",
    "C250F3 HS":            "C250F3 HS",
    "CONSTANTA ENV":        "CONSTANTA ENV",
    "CONSTANTA ENV BT":     "CONSTANTA ENV",
    "CONSTANTA ENV BT 25":  "CONSTANTA ENV",
    "CONSTANTA ENV F":      "CONSTANTA ENV",
    "CONSTANTA ENV G":      "CONSTANTA ENV",
    "CONSTANTA ENV H":      "CONSTANTA ENV",
    "CONSTANTA ENV HF":     "CONSTANTA ENV",
    "CONSTANTA ENV M":      "CONSTANTA ENV",
    "CONSTANTA FAMILY SIZE": "CONSTANTA FAMILY SIZE",
    "CONSTANTA TAG":        "CONSTANTA TAG",
    "CONSTANTA TAG (D)":    "CONSTANTA TAG (D)",
    "INSTANT COFFEE STICK PACK": "INSTANT COFFEE STICK PACK",
    "Instant Tea":          "Instant Tea",
    "MAISA ENV":            "MAISA ENV",
    "MAISA TAGLESS":        "MAISA TAGLESS",
    "MD20 4GM TAG":         "MD20 4GM TAG",
    "MD20 ENV":             "MD20 ENV",
    "MD20 ENV GARANT":      "MD20 ENV GARANT",
    "MD20 HS":              "MD20 HS",
    "PEARL PACK-NANO":      "PEARL PACK-NANO",
    "PEARL PACK-PREMIX":    "PEARL PACK-PREMIX",
    "PEARL PACK-PREMIX 2":  "PEARL PACK-PREMIX",
    "PERFACTA TAG":         "PERFACTA TAG",
    "PERFECTA  K 45 TAG":   "PERFECTA  K 45 TAG",
    "PERFECTA ENV":         "PERFECTA ENV",
    "Pearl Pack-FFS POUCH": "Pearl Pack-FFS POUCH",
    "UNIVERSAL POT BAG":    "UNIVERSAL POT BAG",
    "UNIVERSAL TWIN BAG":   "UNIVERSAL TWIN BAG",
    "VARIETIES PACK":       "VARIETIES PACK",
    "VIKING-MULTITRACK":    "VIKING-MULTITRACK",
}

# Work centers that are not packing lines / cannot be converted to containers
SKIP_WORK_CENTERS = {
    "CFC FORMING", "RECLAIM", "MANUAL", "UNFOLD", "BAG FORMING",
    "MISC", "PAKONA 250GMS/500GMS/1000GMS", "FFS POUCH", "FFS POUCH - 13G- 40G",
}
SKIP_UOMS = {"KGS", "BAG", "POUCH", "ENV", "PKT", "BOX", "CASE", "PCS"}

_HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_FG_BLEND_FILE = os.path.join(_HERE, "FG_Blend_Names.xlsx")


# ── Production Register extraction ───────────────────────────────────────────

def extract_prod_register(pr_bytes):
    """Production Register -> list of (date, group, item_name, qty, uom, work_center)."""
    wb = openpyxl.load_workbook(io.BytesIO(pr_bytes), data_only=True)
    ws = wb["Default"] if "Default" in wb.sheetnames else wb[wb.sheetnames[0]]
    rows = []
    for r in range(8, ws.max_row + 1):
        pdate = ws.cell(r, 1).value
        wc = ws.cell(r, 3).value
        item_name = ws.cell(r, 6).value
        qty = ws.cell(r, 8).value
        uom = ws.cell(r, 10).value
        if not pdate or not wc:
            continue
        if wc in SKIP_WORK_CENTERS or uom in SKIP_UOMS:
            continue
        if uom not in ("CTN", "CFC"):
            continue
        group = ROUTING_MAP.get(wc)
        if not group:
            continue
        d = pdate.date() if isinstance(pdate, datetime.datetime) else pdate
        rows.append((d, group, item_name or "", qty or 0, uom, wc))
    return rows


def register_date_label(pr_bytes):
    """'28-Sep-2026' from the register header (Report Generation Date), or ''."""
    try:
        wb = openpyxl.load_workbook(io.BytesIO(pr_bytes), data_only=True, read_only=True)
        ws = wb["Default"] if "Default" in wb.sheetnames else wb[wb.sheetnames[0]]
        for row in ws.iter_rows(min_row=1, max_row=8, max_col=2, values_only=True):
            if row[0] and "Generation Date" in str(row[0]) and row[1]:
                v = row[1]
                if isinstance(v, (datetime.datetime, datetime.date)):
                    return v.strftime("%d-%b-%Y")
                return datetime.datetime.strptime(str(v).strip(), "%d-%m-%Y").strftime("%d-%b-%Y")
    except Exception:
        pass
    return ""


def load_tbgs_per_ctn(ref_bytes):
    """TBGS PER CTN lookup from the reference workbook."""
    wb = openpyxl.load_workbook(io.BytesIO(ref_bytes), data_only=True)
    ws = wb["TBGS PER CTN"]
    table = {}
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, 2).value
        val = ws.cell(r, 3).value
        if name and val:
            table[name] = val
    return table


# ── Week numbering (Excel WEEKNUM type 2: Monday start, week 1 holds 1 January) ─

# Weeks are numbered continuously from the Monday of the week holding 1 Jan 2026 (= Mon 29-Dec-2025),
# so they never restart at 1 — the same numbering as the report headers (W #36 = 31-Aug-2026).
_JAN1_2026 = datetime.date(2026, 1, 1)
WEEK1_MONDAY = _JAN1_2026 - datetime.timedelta(days=_JAN1_2026.weekday())


def weeknum(d):
    monday = d - datetime.timedelta(days=d.weekday())
    return ((monday - WEEK1_MONDAY).days // 7) + 1


def week_label(wn):
    return f"W #{wn}"


# ── Container conversion ─────────────────────────────────────────────────────

def prod_entries(prod_rows, tbgs_per_ctn, target_weeks):
    """One dict per production entry that can be converted, plus the names that cannot."""
    entries, skipped = [], set()
    for d, group, item_name, qty, uom, wc in prod_rows:
        wn = weeknum(d)
        if wn not in target_weeks:
            continue
        f = MC_FACTORS.get(group, {})
        tpc, cpc = f.get("tbgs_cfc", 0), f.get("cfc_cntr", 0)
        if not tpc or not cpc:
            continue
        if uom == "CTN":
            tbgs_ctn = tbgs_per_ctn.get(item_name)
            source = "table"
            if not tbgs_ctn:
                m = re.search(r"\b(\d+)\s+DC\b", item_name or "")      # "20 DC SF ENV" -> 20
                tbgs_ctn = int(m.group(1)) if m else None
                source = "name"
            if not tbgs_ctn:
                skipped.add(item_name)
                continue
            cfc = qty * tbgs_ctn / tpc
        else:
            tbgs_ctn, source, cfc = None, "CFC", qty
        entries.append(dict(date=d, wn=wn, group=group, wc=wc, uom=uom, qty=qty, tbgs_ctn=tbgs_ctn,
                            source=source, tpc=tpc, cfc=cfc, cpc=cpc, containers=cfc / cpc))
    return entries, skipped


def prod_to_containers(prod_rows, tbgs_per_ctn, target_weeks):
    """Containers per (machine line, week number)."""
    entries, skipped = prod_entries(prod_rows, tbgs_per_ctn, set(target_weeks))
    result = defaultdict(float)
    for e in entries:
        result[(e["group"], e["wn"])] += e["containers"]
    return dict(result), skipped


def compute_capacity(line, machines, shifts):
    """Weekly capacity in containers (same arithmetic as the capacity column formula)."""
    f = MC_FACTORS.get(line, {})
    rate, shiftmin = f.get("rate", 0), f.get("shiftmin", 0)
    tbgs_cfc, cfc_cntr = f.get("tbgs_cfc", 0), f.get("cfc_cntr", 0)
    if not (rate and shiftmin and tbgs_cfc and cfc_cntr and machines and shifts):
        return 0.0
    return rate * shiftmin * machines * shifts * 6 * EFFICIENCY / tbgs_cfc / cfc_cntr


# ── Styles ───────────────────────────────────────────────────────────────────

def _styles():
    FN = "Calibri"
    thin = Side(style="thin")                       # automatic (black) colour, as in the finalized file
    grey = Side(style="thin", color="FFB0B0B0")
    return {
        "FN": FN,
        "BORDER": Border(left=thin, right=thin, top=thin, bottom=thin),
        "BORDER_GREY": Border(left=grey, right=grey, top=grey, bottom=grey),
        "TITLE_L_FILL": PatternFill(start_color="FFE8D5B0", end_color="FFE8D5B0", fill_type="solid"),
        "TITLE_C_FILL": PatternFill(fgColor=Color(theme=5, tint=-0.5), fill_type="solid"),
        "TITLE_R_FILL": PatternFill(fgColor=Color(theme=7, tint=-0.5), fill_type="solid"),
        "HDR_FILL":     PatternFill(start_color="FFB8C4D6", end_color="FFB8C4D6", fill_type="solid"),
        "CNTR_HDR_FILL": PatternFill(fgColor=Color(theme=3, tint=-0.25), fill_type="solid"),
        "PCT_HDR_FILL": PatternFill(fgColor=Color(theme=6, tint=-0.5), fill_type="solid"),
        "PCT_BODY_FILL": PatternFill(start_color="FF92D050", end_color="FF92D050", fill_type="solid"),
        "TOTAL_FILL":   PatternFill(start_color="FFCDC4A0", end_color="FFCDC4A0", fill_type="solid"),
        "REGION_FILL":  PatternFill(start_color="FFD6D6D6", end_color="FFD6D6D6", fill_type="solid"),
        "TITLE_L_FONT": Font(name=FN, bold=True, size=14, color="FF8B1A1A"),
        "TITLE_C_FONT": Font(name=FN, bold=True, size=12, color="FFFFFFFF"),
        "HDR_FONT":     Font(name=FN, bold=True, size=10),
        "WEEK_HDR_FONT": Font(name=FN, bold=True, size=10, color=Color(theme=0)),
        "NORM":         Font(name=FN, size=10),
        "BOLD":         Font(name=FN, bold=True, size=10),
        "OVER_FONT":    Font(name=FN, size=10, color="FFCC3300"),
        "HELPER_FONT":  Font(name=FN, size=9, color="FFAAAAAA"),
        "NUMFMT_CNTR":  "_ * #,##0.0_ ;_ * \\-#,##0.0_ ;_ * \\-??_ ;_ @_ ",
        "NUMFMT_PCT":   "0.0\\%",
    }


# ── Main report builder ──────────────────────────────────────────────────────

def generate_avp_report(pr_bytes, ref_bytes, target_weeks, fg_blend_bytes=None):
    """
    Parameters
    ----------
    pr_bytes       : bytes  Production Register xlsx
    ref_bytes      : bytes  Reference workbook (for TBGS PER CTN)
    target_weeks   : list   e.g. [36, 37, 38, 39]
    fg_blend_bytes : bytes  FG Blend Names xlsx (defaults to the FG_Blend_Names.xlsx next to this file)

    Returns
    -------
    (bytes, set, dict|None) — xlsx bytes, product names that could not be converted,
                              changeover summary (None when no FG Blend Names file is available)
    """
    target_weeks = list(target_weeks)
    if fg_blend_bytes is None and os.path.exists(DEFAULT_FG_BLEND_FILE):
        with open(DEFAULT_FG_BLEND_FILE, "rb") as fh:
            fg_blend_bytes = fh.read()
    with_co = fg_blend_bytes is not None

    tbgs_per_ctn = load_tbgs_per_ctn(ref_bytes)
    prod_rows = extract_prod_register(pr_bytes)
    entries, skipped = prod_entries(prod_rows, tbgs_per_ctn, set(target_weeks))
    achieved = defaultdict(float)
    for e in entries:
        achieved[(e["group"], e["wn"])] += e["containers"]

    n_weeks = len(target_weeks)
    CNTR_FIRST = 7                                  # G
    PCT_FIRST = CNTR_FIRST + n_weeks                # K
    CO_FIRST = PCT_FIRST + n_weeks                  # O  (Blend block, then Brand block)
    LAST_COL = CO_FIRST + 2 * n_weeks - 1 if with_co else PCT_FIRST + n_weeks - 1
    LKP_COL = LAST_COL + 3                          # hidden capacity factors

    S = _styles()
    BORDER = S["BORDER"]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET_NAME

    # ── Row 1: title and block headers ──────────────────────────────────────
    ws.merge_cells("A1:F1")
    ws["A1"] = f"Capacity Vs Produced in last {n_weeks} Weeks"
    ws["A1"].font, ws["A1"].fill = S["TITLE_L_FONT"], S["TITLE_L_FILL"]
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    for first, text, fill in ((CNTR_FIRST, "By container", S["TITLE_C_FILL"]),
                              (PCT_FIRST, "By Percentage  ", S["TITLE_R_FILL"])):
        ws.merge_cells(start_row=1, start_column=first, end_row=1, end_column=first + n_weeks - 1)
        c = ws.cell(1, first, text)
        c.font, c.fill = S["TITLE_C_FONT"], fill
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    # ── Row 2: column headers ───────────────────────────────────────────────
    ws.row_dimensions[2].height = 36
    for c, h in {1: "Region", 2: None, 3: "Machine Line", 4: "No. of\nMachines",
                 5: "No. of\nShifts", 6: CAPACITY_HEADER}.items():
        cell = ws.cell(2, c, h)
        cell.font, cell.fill, cell.border = S["HDR_FONT"], S["HDR_FILL"], BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, wn in enumerate(target_weeks):
        for col, fill in ((CNTR_FIRST + i, S["CNTR_HDR_FILL"]), (PCT_FIRST + i, S["PCT_HDR_FILL"])):
            cell = ws.cell(2, col, week_label(wn))
            cell.font, cell.fill, cell.border = S["WEEK_HDR_FONT"], fill, BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # ── Column widths ───────────────────────────────────────────────────────
    for col, w in COL_WIDTHS.items():
        ws.column_dimensions[col].width = w
    for c in range(4, LAST_COL + 1):
        ws.column_dimensions[get_column_letter(c)].width = WEEK_COL_WIDTH

    # ── Hidden capacity-factor columns ──────────────────────────────────────
    for i, h in enumerate(["Rate", "ShMin", "TBGS/CFC", "CFC/Cntr"]):
        ws.cell(2, LKP_COL + i, h).font = S["HELPER_FONT"]
        ws.column_dimensions[get_column_letter(LKP_COL + i)].hidden = True
        ws.column_dimensions[get_column_letter(LKP_COL + i)].width = HELPER_WIDTH
    Qc, Rc, Sc, Tc = (get_column_letter(LKP_COL + i) for i in range(4))

    # ── Data rows ───────────────────────────────────────────────────────────
    row = 3
    pending_merges = []
    cur_a_start = cur_b_start = None
    cur_region = cur_sub = None
    grp_start = row
    line_rows, total_groups = {}, {}

    for region, sub, line, nm, ns in LAYOUT:
        ws.row_dimensions[row].height = 18

        if region == "__TOTAL__":
            total_groups[row] = (grp_start, row - 1)
            for c in range(1, PCT_FIRST + n_weeks):
                cell = ws.cell(row, c)
                cell.border = BORDER
                if c >= 3:
                    cell.fill, cell.font = S["TOTAL_FILL"], S["BOLD"]
            ws.cell(row, 3, "TOTAL")
            f_cell = ws.cell(row, 6, f"=SUM(F{grp_start}:F{row - 1})")
            f_cell.number_format = S["NUMFMT_CNTR"]
            for i in range(n_weeks):
                col = get_column_letter(CNTR_FIRST + i)
                a_cell = ws.cell(row, CNTR_FIRST + i, f"=SUM({col}{grp_start}:{col}{row - 1})")
                a_cell.number_format = S["NUMFMT_CNTR"]
                p_cell = ws.cell(row, PCT_FIRST + i,
                                 f'=IFERROR(ROUND(SUM({col}{grp_start}:{col}{row - 1})/'
                                 f'SUMIF(F{grp_start}:F{row - 1},">0")*100,1),0)')
                p_cell.number_format = S["NUMFMT_PCT"]
                p_cell.fill = S["PCT_BODY_FILL"]
            grp_start = row + 1
            row += 1
            if cur_b_start is not None:
                pending_merges.append(("B", cur_b_start, row - 2))
                cur_b_start = None
                cur_sub = None
            continue

        # region / sub-group labels
        if region and region != cur_region:
            if cur_a_start is not None:
                pending_merges.append(("A", cur_a_start, row - 1))
            cur_region, cur_a_start = region, row
            ws.cell(row, 1, region).font = S["BOLD"]
            ws.cell(row, 1).fill = S["REGION_FILL"]
        if sub and sub != cur_sub:
            if cur_b_start is not None:
                pending_merges.append(("B", cur_b_start, row - 1))
            cur_sub, cur_b_start = sub, row
            ws.cell(row, 2, sub).font = S["NORM"]
            if len(sub) > 10:
                ws.cell(row, 2).alignment = Alignment(wrap_text=True)
        ws.cell(row, 1).border = BORDER
        ws.cell(row, 2).border = BORDER

        line_rows[row] = line
        ws.cell(row, 3, line).font = S["NORM"]
        ws.cell(row, 3).border = BORDER
        for c, v in ((4, nm), (5, ns)):                               # editable inputs
            cell = ws.cell(row, c, v)
            cell.font, cell.border = S["NORM"], BORDER
            cell.alignment = Alignment(horizontal="right", vertical="center")

        f = MC_FACTORS.get(line, {})
        for i, key in enumerate(["rate", "shiftmin", "tbgs_cfc", "cfc_cntr"]):
            ws.cell(row, LKP_COL + i, f.get(key, 0)).font = S["HELPER_FONT"]

        zero_cap = (ns == 0)
        # F and K-N are always live formulas (never a static "-"), so changing
        # Machines/Shifts for a line that started at 0 shifts still recalculates.
        f_cell = ws.cell(row, 6)
        f_cell.border, f_cell.font = BORDER, S["NORM"]
        f_cell.alignment = Alignment(horizontal="right", vertical="center")
        f_cell.number_format = S["NUMFMT_CNTR"]
        f_cell.value = (
            f'=IFERROR(IF({Tc}{row}=0,"-",'
            f'ROUND({Qc}{row}*{Rc}{row}*$D{row}*$E{row}*6*0.9/{Sc}{row}/{Tc}{row},3)),"-")')

        cap_val = round(compute_capacity(line, nm, ns), 3)
        for i, wn in enumerate(target_weeks):
            ach = achieved.get((line, wn), 0)
            over = bool(cap_val and ach > cap_val)
            c_cell = ws.cell(row, CNTR_FIRST + i)
            c_cell.border = BORDER
            c_cell.alignment = Alignment(horizontal="right", vertical="center")
            if not zero_cap:
                c_cell.number_format = S["NUMFMT_CNTR"]
            c_cell.font = S["OVER_FONT"] if over else S["NORM"]
            c_cell.value = "-" if (zero_cap or not ach) else round(ach, 3)

            acol = get_column_letter(CNTR_FIRST + i)
            p_cell = ws.cell(row, PCT_FIRST + i)
            p_cell.border, p_cell.fill = BORDER, S["PCT_BODY_FILL"]
            p_cell.alignment = Alignment(horizontal="right", vertical="center")
            p_cell.number_format = S["NUMFMT_PCT"]
            p_cell.font = S["OVER_FONT"] if over else S["NORM"]
            p_cell.value = (
                f'=IFERROR(IF($F{row}="-","-",IF($F{row}=0,0,ROUND({acol}{row}/$F{row}*100,1))),"-")')
        row += 1

    if cur_a_start is not None:
        pending_merges.append(("A", cur_a_start, row - 1))
    if cur_b_start is not None:
        pending_merges.append(("B", cur_b_start, row - 1))
    for cl, ds, de in pending_merges:
        if de > ds:
            ws.merge_cells(f"{cl}{ds}:{cl}{de}")
        wrap = cl == "B" and len(str(ws[f"{cl}{ds}"].value or "")) > 10      # e.g. "DC SF CRIMP"
        ws[f"{cl}{ds}"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True if wrap else None)

    ws.freeze_panes = f"{get_column_letter(CNTR_FIRST)}3"
    ws.sheet_view.showGridLines = False

    # ── Blend / Brand changeover ────────────────────────────────────────────
    co_info, model = None, None
    if with_co:
        model = changeover.add_changeover(
            wb, ws, pr_bytes, fg_blend_bytes, ROUTING_MAP, list(dict.fromkeys(line_rows.values())),
            target_weeks, CO_FIRST, line_rows, total_groups, footer_row=row, week1_monday=WEEK1_MONDAY)
        blend_tot, brand_tot = model.totals
        co_info = {
            "blend_total": blend_tot,
            "brand_total": brand_tot,
            "unknown_items": model.unknown_items,          # FG items with no blend in the FG list
            "not_in_report": model.not_in_report(),        # work centers with no report line
            "machine_shifts": len(model.calc),
        }

    for c in range(1, LAST_COL + 1):                       # thin frame round the title / block headers
        ws.cell(1, c).border = BORDER

    # ── Source Calculations (hidden) ────────────────────────────────────────
    _write_source_calculations(wb, entries, target_weeks, S, register_date_label(pr_bytes))
    for name in (SOURCE_SHEET, changeover.SHEET_CALC, changeover.SHEET_DETAIL,
                 changeover.SHEET_ROUTING, changeover.SHEET_FG):
        if name in wb.sheetnames:
            wb[name].sheet_state = "hidden"
    order = [SHEET_NAME, SOURCE_SHEET, changeover.SHEET_CALC, changeover.SHEET_DETAIL,
             changeover.SHEET_ROUTING, changeover.SHEET_FG]
    wb._sheets = [wb[n] for n in order if n in wb.sheetnames]
    wb.active = 0

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue(), skipped, co_info


# ── Source Calculations sheet ────────────────────────────────────────────────

def _write_source_calculations(wb, entries, target_weeks, S, register_label):
    FN = S["FN"]
    BORDER = S["BORDER_GREY"]          # Section A (header + lines)
    ws = wb.create_sheet(SOURCE_SHEET)
    wk_text = ", ".join(f"W#{w}" for w in target_weeks)
    title_fill = S["TITLE_L_FILL"]
    hdr_a = PatternFill(start_color="FF1F3864", end_color="FF1F3864", fill_type="solid")
    hdr_b = PatternFill(start_color="FF38761D", end_color="FF38761D", fill_type="solid")
    cap_fill = PatternFill(start_color="FFDCE6F1", end_color="FFDCE6F1", fill_type="solid")
    tot_fill = S["TOTAL_FILL"]
    white_b = Font(name=FN, bold=True, size=10, color="FFFFFFFF")
    norm = Font(name=FN, size=10)
    bold = Font(name=FN, bold=True, size=10)
    grey9 = Font(name=FN, size=9, color="FF888888")
    grey10 = Font(name=FN, size=10, color="FF888888")
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    center = Alignment(horizontal="center", vertical="center")
    right = Alignment(horizontal="right", vertical="center")

    ws.merge_cells("A1:O1")
    ws["A1"] = f"Source Calculations — {wk_text}"
    ws["A1"].font = Font(name=FN, bold=True, size=14, color="FF8B1A1A")
    ws["A1"].fill, ws["A1"].alignment = title_fill, left
    ws.row_dimensions[1].height = 18
    ws.merge_cells("A2:O2")
    ws["A2"] = ("CAPACITY: Rate(TBGS/min) × ShiftMin × Machines × Shifts × 6 days × 90% ÷ TBGS/CFC ÷ CFC/Container  |  "
                "PRODUCTION: Prod Qty(CTN) × TBGS/CTN ÷ TBGS/CFC ÷ CFC/Container  |  CFC entries: qty ÷ CFC/Container")
    ws["A2"].font = Font(name=FN, size=9, italic=True, color="FF444444")
    ws["A2"].alignment = left
    ws.row_dimensions[2].height = 14.25

    # Section A — capacity
    ws.merge_cells("A3:M3")
    ws["A3"] = "  SECTION A — ACHIEVABLE CAPACITY (New Capacity — Arul Sir Factors)"
    ws["A3"].font, ws["A3"].fill, ws["A3"].alignment = white_b, hdr_a, left
    ws.row_dimensions[3].height = 14.25
    heads_a = ["Machine Line", "No. of\nMachines", "No. of\nShifts", "Rate\n(TBGS/min)", "Shift Min\n(active)",
               "TBGS/CFC", "CFC/Cntr", "TBGS/Shift\n(Rate×ShMin×M)", "TBGS/Week\n(×S×6)", "Achievable\n(×90%)",
               "CFC/Week\n(TBGS÷TBGS/CFC)", "Cap(containers)\n(CFC÷CFC/Cntr)", "Capacity Formula"]
    ws.row_dimensions[4].height = 54.75
    for c, h in enumerate(heads_a, 1):
        cell = ws.cell(4, c, h)
        cell.font, cell.fill, cell.border = white_b, hdr_a, BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    r = 5
    for region, sub, line, nm, ns in LAYOUT:
        if region == "__TOTAL__":
            continue
        f = MC_FACTORS.get(line, {})
        rate, shmin = f.get("rate", 0), f.get("shiftmin", 0)
        tpc, cpc = f.get("tbgs_cfc", 0), f.get("cfc_cntr", 0)
        zero = (ns == 0)
        if not zero and rate and shmin and tpc and cpc:
            tbgs_shift = rate * shmin * nm
            tbgs_week = tbgs_shift * ns * 6
            achievable = round(tbgs_week * EFFICIENCY, 0)
            cfc_week = round(achievable / tpc, 3)
            cap = round(cfc_week / cpc, 4)
            vals = [line, nm, ns, rate, shmin, tpc, cpc, int(tbgs_shift), int(tbgs_week), int(achievable),
                    round(cfc_week, 2), cap, f"({rate}×{shmin}×{nm}×{ns}×6×90%) ÷ {tpc} ÷ {cpc} = {cap}"]
        else:
            vals = [line, nm, ns, rate, shmin, tpc, cpc, "-", "-", "-", "-", "-",
                    "No shifts — not running" if zero else "Missing factors"]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v)
            cell.font = norm
            cell.border = BORDER
            cell.alignment = left if c in (1, 13) else right
            if 4 <= c <= 12:
                cell.fill = cap_fill
            if zero and c > 2 and c < 13:
                cell.font = grey10
            if c == 13:
                cell.font = grey10 if zero else grey9
        r += 1

    # Section B — achieved production
    r += 1
    section_b_row = r
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=15)
    label = f"Production Register {register_label} — " if register_label else "Production Register — "
    ws.cell(r, 1, f"  SECTION B — ACHIEVED PRODUCTION ({label}W#{target_weeks[0]} to W#{target_weeks[-1]})")
    ws.cell(r, 1).font, ws.cell(r, 1).fill, ws.cell(r, 1).alignment = white_b, hdr_b, left
    ws.row_dimensions[r].height = 14.25
    r += 1
    heads_b = ["#", "Group\n(Cap. Line)", "Date", "Wk", "Work Center", "UOM", "Prod Qty\n(A)", "TBGS/CTN\n(B)",
               "Source", "TBGS/CFC\n(C)", "CFC Qty\n(A×B÷C)", "CFC/Cntr\n(D)", "Containers\n(CFC÷D)", "Full Formula"]
    ws.row_dimensions[r].height = 27
    for c, h in enumerate(heads_b, 1):
        cell = ws.cell(r, c, h)
        cell.font, cell.fill = white_b, hdr_b
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    r += 1

    wk_fill = {36: "FFDCE6F1", 37: "FFD9EAD3", 38: "FFFFF2CC", 39: "FFFCE5CD"}
    palette = ["FFDCE6F1", "FFD9EAD3", "FFFFF2CC", "FFFCE5CD"]
    wk_fills = {w: PatternFill(start_color=wk_fill.get(w, palette[i % 4]),
                               end_color=wk_fill.get(w, palette[i % 4]), fill_type="solid")
                for i, w in enumerate(target_weeks)}
    wk_font = Font(name=FN, bold=True, size=10, color="FF1F3864")

    blocks = defaultdict(list)
    for e in entries:
        blocks[(e["group"], e["wn"])].append(e)
    idx = 1
    for (group, wn), rows in sorted(blocks.items(), key=lambda kv: (kv[0][0], kv[0][1])):
        total = 0.0
        for e in sorted(rows, key=lambda x: str(x["date"])):
            shown = round(e["containers"], 6)
            if e["uom"] == "CTN":
                formula = (f'({e["qty"]:,}CTN×{e["tbgs_ctn"]}TBGS/CTN÷{e["tpc"]}TBGS/CFC)'
                           f'÷{e["cpc"]}CFC/Cntr = {shown:.4f}')
            else:
                formula = f'{e["qty"]:,}CFC÷{e["cpc"]}CFC/Cntr = {shown:.4f}'
            vals = [idx, group, e["date"], f"W#{wn}", e["wc"], e["uom"], e["qty"],
                    e["tbgs_ctn"] if e["tbgs_ctn"] else "—", e["source"], e["tpc"], round(e["cfc"], 4),
                    e["cpc"], round(e["containers"], 6), formula]
            for c, v in enumerate(vals, 1):
                cell = ws.cell(r, c, v)
                cell.font = norm
                cell.alignment = (right if c in (1, 7, 8, 10, 11, 12, 13) else
                                  center if c in (4, 6, 9) else left)
                if c == 3:
                    cell.number_format = "d-mmm-yy"
                    cell.alignment = center
                if c == 4:
                    cell.fill, cell.font = wk_fills[wn], wk_font
            total += e["containers"]
            idx += 1
            r += 1
        for c in range(1, 15):
            ws.cell(r, c).fill = tot_fill
        ws.cell(r, 2, f"{group}  —  W#{wn} TOTAL").font = bold
        ws.cell(r, 4, f"W#{wn}")
        ws.cell(r, 4).fill, ws.cell(r, 4).font, ws.cell(r, 4).alignment = wk_fills[wn], wk_font, center
        ws.cell(r, 6, "entries").font = norm
        ws.cell(r, 7, len(rows)).font = bold
        ws.cell(r, 13, round(total, 4)).font = bold
        ws.cell(r, 14, f"Sum of {len(rows)} production entries = {round(total, 4)} containers").font = grey9
        r += 1

    black = S["BORDER"]                                    # Section B: black frame over the whole A..O region
    for rr in range(section_b_row, r):
        for c in range(1, 16):
            ws.cell(rr, c).border = black

    for col, w in {'A': 23.42578125, 'B': 12.140625, 'C': 9.140625, 'D': 5.5703125, 'E': 23.42578125, 'F': 6.28515625, 'G': 8.0, 'H': 9.0, 'I': 6.28515625, 'J': 8.7109375, 'K': 9.0, 'L': 8.0, 'M': 31.7109375, 'N': 54.28515625}.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A5"
    ws.sheet_view.showGridLines = False
