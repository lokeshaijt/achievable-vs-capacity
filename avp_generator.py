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

COL_WIDTHS = {"A": 9.0, "B": 9.42578125, "C": 9.140625, "D": 23.7109375}   # E .. last column: WEEK_COL_WIDTH
WEEK_COL_WIDTH = 7.42578125
HELPER_WIDTH = 9.0

# ── Layout: zone, region, sub-group, machine line, machines, shifts ──────────
# (machines / shifts are the editable input cells of the report; grouped and
#  sorted Zone 1 -> Zone 2 -> Zone 3, region/classification order inside each)
LAYOUT = [
    ("Zone 1", "AFRICA",    "DC",          "PERFECTA ENV(Black Tea)",   3,  2),
    (None,     None,        None,          "C250 AF(Black Tea)",        2,  2),
    (None,     None,        None,          "MD20 ENV(Black Tea)",       2,  2),
    (None,     None,        None,          "CONSTANTA ENV(Flavour Tea)",23, 2),
    (None,     None,        None,          "CONSTANTA ENV (HERBAL)",    9,  2),
    ("__TOTAL__", None, None, None, None, None),
    (None,     "AUSTRALIA", "DC TAG",      "CONSTANTA TAG",             6,  1),
    ("__TOTAL__", None, None, None, None, None),
    (None,     "EUROPE",    "DC SF",       "C250",                      3,  2),
    (None,     None,        None,          "VARIETIES PACK",            1,  2),
    (None,     None,        None,          "MD20 ENV GARANT",           2,  2),
    ("__TOTAL__", None, None, None, None, None),
    (None,     None,        "DC TAG",      "MD20 4GM TAG",              1,  1),
    (None,     None,        None,          "PERFECTA  K 45 TAG",        1,  1),
    ("__TOTAL__", None, None, None, None, None),
    (None,     "RUSSIA",    "DC TAG",      "MD20 HS",                   2,  1),
    ("__TOTAL__", None, None, None, None, None),
    (None,     "USA",       "TREE HOUSE",  "CONSTANTA FAMILY SIZE",     2,  0),
    (None,     None,        None,          "PERFACTA TAG",              2,  1),
    (None,     None,        None,          "C250-A",                   3,  2),
    ("__TOTAL__", None, None, None, None, None),
    ("Zone 2", "AFRICA",    "SC",          "MAISA TAGLESS",             7,  1),
    ("__TOTAL__", None, None, None, None, None),
    (None,     "EUROPE",    "SC",          "UNIVERSAL POT BAG",         1,  0),
    (None,     None,        None,          "UNIVERSAL TWIN BAG",        1,  1),
    ("__TOTAL__", None, None, None, None, None),
    (None,     "USA",       "TREE HOUSE",  "BREW MAGIC",                1,  1),
    (None,     None,        None,          "MAISA ENV",                 2,  1),
    ("__TOTAL__", None, None, None, None, None),
    ("Zone 3", "AFRICA",    "PREMIX",      "INSTANT COFFEE STICK PACK", 1,  1),
    (None,     None,        None,          "PEARL PACK-PREMIX",         1,  1),
    (None,     None,        None,          "PEARL PACK-PREMIX - cappuccino", 2, 1),
    ("__TOTAL__", None, None, None, None, None),
    (None,     "USA",       "TREE HOUSE",  "Instant Tea",               1,  0),
    (None,     None,        None,          "VIKING-MULTITRACK",         1,  0),
    ("__TOTAL__", None, None, None, None, None),
    ("Zone 4", "INDIA",     "HF Co",       "Shubam-FFS POUCH",          1,  1),
    ("__TOTAL__", None, None, None, None, None),
    (None,     "USA",       "HF Co",       "Pearl Pack-FFS POUCH",      0.5, 1),
    ("__TOTAL__", None, None, None, None, None),
]

# ── Capacity factors per machine line (New capacity — Arul Sir) ──────────────
MC_FACTORS = {
    "BREW MAGIC":                     {"rate": 8,   "shiftmin": 540, "tbgs_cfc": 32,   "cfc_cntr": 3600},
    "C250":                           {"rate": 250, "shiftmin": 494, "tbgs_cfc": 480,  "cfc_cntr": 7200},
    "C250 AF(Black Tea)":             {"rate": 250, "shiftmin": 524, "tbgs_cfc": 1200, "cfc_cntr": 3340},
    "C250-A":                         {"rate": 250, "shiftmin": 502, "tbgs_cfc": 480,  "cfc_cntr": 3600},
    "VARIETIES PACK":                 {"rate": 250, "shiftmin": 503, "tbgs_cfc": 480,  "cfc_cntr": 7200},
    "INSTANT COFFEE STICK PACK":      {"rate": 40,  "shiftmin": 540, "tbgs_cfc": 1350, "cfc_cntr": 3000},
    "CONSTANTA ENV(Flavour Tea)":     {"rate": 125, "shiftmin": 542, "tbgs_cfc": 600,  "cfc_cntr": 6375},
    "CONSTANTA ENV (HERBAL)":         {"rate": 125, "shiftmin": 553, "tbgs_cfc": 1125, "cfc_cntr": 3340},
    "CONSTANTA FAMILY SIZE":          {"rate": 95,  "shiftmin": 540, "tbgs_cfc": 144,  "cfc_cntr": 5720},
    "CONSTANTA TAG":                  {"rate": 120, "shiftmin": 541, "tbgs_cfc": 1200, "cfc_cntr": 2688},
    "Instant Tea":                    {"rate": 14,  "shiftmin": 550, "tbgs_cfc": 60,   "cfc_cntr": 15120},
    "MAISA ENV":                      {"rate": 100, "shiftmin": 539, "tbgs_cfc": 1200, "cfc_cntr": 1656},
    "MAISA TAGLESS":                  {"rate": 120, "shiftmin": 555, "tbgs_cfc": 2400, "cfc_cntr": 2800},
    "MD20 4GM TAG":                   {"rate": 125, "shiftmin": 548, "tbgs_cfc": 480,  "cfc_cntr": 4032},
    "MD20 ENV(Black Tea)":            {"rate": 185, "shiftmin": 528, "tbgs_cfc": 1200, "cfc_cntr": 3340},
    "MD20 ENV GARANT":                {"rate": 185, "shiftmin": 514, "tbgs_cfc": 1200, "cfc_cntr": 3340},
    "MD20 HS":                        {"rate": 170, "shiftmin": 516, "tbgs_cfc": 300,  "cfc_cntr": 6500},
    "Pearl Pack-FFS POUCH":           {"rate": 10,  "shiftmin": 521, "tbgs_cfc": 24,   "cfc_cntr": 1200},
    "PEARL PACK-PREMIX":              {"rate": 30,  "shiftmin": 550, "tbgs_cfc": 120,  "cfc_cntr": 6000},
    "PEARL PACK-PREMIX - cappuccino": {"rate": 30,  "shiftmin": 550, "tbgs_cfc": 120,  "cfc_cntr": 4800},
    "PERFACTA TAG":                   {"rate": 300, "shiftmin": 525, "tbgs_cfc": 600,  "cfc_cntr": 4620},
    "PERFECTA  K 45 TAG":             {"rate": 250, "shiftmin": 533, "tbgs_cfc": 600,  "cfc_cntr": 5040},
    "PERFECTA ENV(Black Tea)":        {"rate": 280, "shiftmin": 523, "tbgs_cfc": 1200, "cfc_cntr": 3340},
    "Shubam-FFS POUCH":               {"rate": 190, "shiftmin": 530, "tbgs_cfc": 0,    "cfc_cntr": 0},
    "UNIVERSAL POT BAG":              {"rate": 272, "shiftmin": 540, "tbgs_cfc": 960,  "cfc_cntr": 0},
    "UNIVERSAL TWIN BAG":             {"rate": 420, "shiftmin": 537, "tbgs_cfc": 960,  "cfc_cntr": 3320},
    "VIKING-MULTITRACK":              {"rate": 56,  "shiftmin": 550, "tbgs_cfc": 60,   "cfc_cntr": 15120},
}

# ── Routing map: work-center name -> machine line (fallback when an item has
#    no exact match in the MC Routing Master, see load_item_routing below) ───
ROUTING_MAP = {
    "BREW MAGIC":                "BREW MAGIC",
    "C250":                      "C250",
    "C250 AF":                   "C250 AF(Black Tea)",
    "C250-A":                    "C250-A",
    "CONSTANTA TAG (D)":         "CONSTANTA TAG",
    "INSTANT COFFEE STICK PACK": "INSTANT COFFEE STICK PACK",
    "Instant Tea":               "Instant Tea",
    "MAISA ENV":                 "MAISA ENV",
    "MAISA TAGLESS":             "MAISA TAGLESS",
    "MD20 4GM TAG":              "MD20 4GM TAG",
    "MD20 ENV":                  "MD20 ENV(Black Tea)",
    "MD20 ENV GARANT":           "MD20 ENV GARANT",
    "MD20 HS":                   "MD20 HS",
    "PERFACTA TAG":              "PERFACTA TAG",
    "PERFECTA  K 45 TAG":        "PERFECTA  K 45 TAG",
    "PERFECTA ENV":              "PERFECTA ENV(Black Tea)",
    "FFS POUCH":                 "Pearl Pack-FFS POUCH",
    # "FFS POUCH - 13G- 40G" intentionally not routed here — excluded from
    # the Pearl Pack-FFS POUCH line, falls through SKIP_WORK_CENTERS-style drop.
    "UNIVERSAL POT BAG":         "UNIVERSAL POT BAG",
    "UNIVERSAL TWIN BAG":        "UNIVERSAL TWIN BAG",
    "VARIETIES PACK":            "VARIETIES PACK",
    "VIKING-MULTITRACK":         "VIKING-MULTITRACK",
    "CONSTANTA ENV":             "CONSTANTA ENV(Flavour Tea)",
    "CONSTANTA ENV BT":          "CONSTANTA ENV(Flavour Tea)",
    "CONSTANTA ENV BT 25":       "CONSTANTA ENV(Flavour Tea)",
    "CONSTANTA ENV F":           "CONSTANTA ENV(Flavour Tea)",
    "CONSTANTA ENV G":           "CONSTANTA ENV(Flavour Tea)",
    "CONSTANTA ENV H":           "CONSTANTA ENV(Flavour Tea)",
    "CONSTANTA ENV HF":          "CONSTANTA ENV(Flavour Tea)",
    "CONSTANTA ENV M":           "CONSTANTA ENV(Flavour Tea)",
    "PEARL PACK-PREMIX":         "PEARL PACK-PREMIX",
    "PEARL PACK-PREMIX 2":       "PEARL PACK-PREMIX",
}

# Items with no herbal/cappuccino signal in the MC Routing Master are split by
# name as a fallback (see load_item_routing / extract_prod_register).
HERBAL_ITEMS = {
    "KHAIMA CRANBERRY APPLE 25 DC ENV TBGS",
    "GV CRANBERRY APPLE 25 DC ENV TBGS - 2GM",
}

# MC Routing Master line names -> this file's LAYOUT line names (None = no
# home in the current report, so items routed there are dropped).
ITEM_MASTER_LINE_MAP = {
    "CONSTANTA ENV":                "CONSTANTA ENV(Flavour Tea)",
    "CONSTANTA ENV H":              "CONSTANTA ENV (HERBAL)",
    "C250":                         "C250",
    "C250 AF":                      "C250 AF(Black Tea)",
    "C250-A":                       "C250-A",
    "MD20 HS":                      "MD20 HS",
    "Pearl Pack-FFS POUCH":         "Pearl Pack-FFS POUCH",
    "MD20 ENV":                     "MD20 ENV(Black Tea)",
    "PERFACTA TAG":                 "PERFACTA TAG",
    "PERFECTA ENV":                 "PERFECTA ENV(Black Tea)",
    "VARIETIES PACK":               "VARIETIES PACK",
    "Pearl Pack-Premix":            "PEARL PACK-PREMIX",
    "Pearl Pack Premix":            "PEARL PACK-PREMIX",
    "PERFECTA  K 45 TAG":           "PERFECTA  K 45 TAG",
    "MAISA ENV":                    "MAISA ENV",
    "VIKING-MULTITRACK":            "VIKING-MULTITRACK",
    "BREW MAGIC":                   "BREW MAGIC",
    "CONSTANTA TAG D":              "CONSTANTA TAG",
    "INSTANT COFFEE STICK PACK":    "INSTANT COFFEE STICK PACK",
    "MD 20 4GM":                    "MD20 4GM TAG",
    "UNIVERSAL TWIN BAG":           "UNIVERSAL TWIN BAG",
    "MAISA TAGLESS":                "MAISA TAGLESS",
    "PAKONA 250GMS/500GMS/1000GMS": None,
    "MD20 ENV GARANT":              "MD20 ENV GARANT",
    "C250F3 HS":                    None,
}

# Work centers that are not packing lines / cannot be converted to containers
SKIP_WORK_CENTERS = {
    "CFC FORMING", "RECLAIM", "MANUAL", "UNFOLD", "BAG FORMING",
    "MISC", "PAKONA 250GMS/500GMS/1000GMS",
}
# KGS is a weight, not a piece/tea-bag count, so it can't be converted the
# same way as the others — everything else (CTN, CFC, ENV, POUCH, PKT, BOX,
# CASE, PCS) is allowed through.
SKIP_UOMS = {"KGS"}

_HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_FG_BLEND_FILE = os.path.join(_HERE, "FG_Blend_Names.xlsx")
DEFAULT_ITEM_ROUTING_FILE = os.path.join(_HERE, "MC_Routing_Master.xlsx")


def load_item_routing(path_or_bytes):
    """MC Routing Master (item name -> report line, normalized via
    ITEM_MASTER_LINE_MAP). A line not in ITEM_MASTER_LINE_MAP is left out so
    unrecognized lines fall back to the work-center ROUTING_MAP instead."""
    wb = openpyxl.load_workbook(
        io.BytesIO(path_or_bytes) if isinstance(path_or_bytes, (bytes, bytearray)) else path_or_bytes,
        data_only=True)
    ws = wb[wb.sheetnames[0]]
    out = {}
    for r in range(2, ws.max_row + 1):
        name, line = ws.cell(r, 1).value, ws.cell(r, 2).value
        if not name or not line:
            continue
        line = str(line).strip()
        if line in ITEM_MASTER_LINE_MAP:
            out[str(name).strip()] = ITEM_MASTER_LINE_MAP[line]      # may be None -> explicit drop
    return out


# ── Production Register extraction ───────────────────────────────────────────

def extract_prod_register(pr_bytes, item_routing=None):
    """Production Register -> list of (date, group, item_name, qty, uom, work_center).

    item_routing : {item name: report line or None} from the MC Routing Master
    (load_item_routing). An exact item-name match there wins over the
    work-center ROUTING_MAP; a value of None explicitly drops that item.

    Columns are located by header name (not fixed position) — the export's
    column layout has changed between registers (e.g. extra "BOM Ver Id" /
    "Version_No" columns inserted), so a fixed-offset read would silently
    pull from the wrong columns and extract nothing.
    """
    wb = openpyxl.load_workbook(io.BytesIO(pr_bytes), data_only=True)
    ws = wb["Default"] if "Default" in wb.sheetnames else wb[wb.sheetnames[0]]
    item_routing = item_routing or {}

    hdr_row = next((r for r in range(1, 15) if str(ws.cell(r, 1).value).strip() == "Process_Date"), None)
    if hdr_row is None:
        raise ValueError("Production Register: could not find the header row (Process_Date).")
    cols = {str(ws.cell(hdr_row, c).value).strip(): c
            for c in range(1, ws.max_column + 1) if ws.cell(hdr_row, c).value is not None}
    need = ["Process_Date", "Work Center", "Item Name", "Production Qty", "UOM"]
    missing = [n for n in need if n not in cols]
    if missing:
        raise ValueError(f"Production Register is missing column(s): {', '.join(missing)}")
    c_date, c_wc, c_item, c_qty, c_uom = (cols[n] for n in need)
    c_txn = cols.get("Transaction Type")        # not present in every export vintage

    rows = []
    for r in range(hdr_row + 1, ws.max_row + 1):
        pdate = ws.cell(r, c_date).value
        wc = ws.cell(r, c_wc).value
        item_name = ws.cell(r, c_item).value or ""
        qty = ws.cell(r, c_qty).value
        uom = ws.cell(r, c_uom).value
        if not pdate or not wc:
            continue
        if c_txn and str(ws.cell(r, c_txn).value).strip() != "Production":
            continue
        if wc in SKIP_WORK_CENTERS or uom in SKIP_UOMS:
            continue

        if item_name in item_routing:
            group = item_routing[item_name]                        # may be None -> explicit drop
        elif wc in ("PEARL PACK-PREMIX", "PEARL PACK-PREMIX 2") and "CAPPUCCINO" in item_name.upper():
            group = "PEARL PACK-PREMIX - cappuccino"
        elif wc and wc.startswith("CONSTANTA ENV") and item_name in HERBAL_ITEMS:
            group = "CONSTANTA ENV (HERBAL)"
        else:
            group = ROUTING_MAP.get(wc)
        if not group:
            continue

        d = pdate.date() if isinstance(pdate, datetime.datetime) else pdate
        rows.append((d, group, item_name, qty or 0, uom, wc))
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


def _month_groups(weeks, week1_monday):
    """[[month_label, [(week_number, year, month, days_in_this_month), ...]], ...]
    for the given weeks. A week whose 7 days (Mon-Sun) fall across a calendar
    month boundary appears in BOTH months' groups — once per month it
    touches — each carrying only the day-count actually in that month, so
    callers can split that week's column (and its production) accordingly."""
    groups, index = [], {}
    for wn in weeks:
        monday = week1_monday + datetime.timedelta(days=(wn - 1) * 7)
        days_in_month = defaultdict(int)
        for i in range(7):
            d = monday + datetime.timedelta(days=i)
            days_in_month[(d.year, d.month)] += 1
        for (y, m) in sorted(days_in_month):
            key = (y, m)
            if key not in index:
                index[key] = len(groups)
                groups.append([datetime.date(y, m, 1).strftime("%b").upper(), []])
            groups[index[key]][1].append((wn, y, m, days_in_month[key]))
    return groups


def _block_columns(start_col, month_groups):
    """Lay out week + month-total columns for one data block, left to right.
    A week split across two months gets one column per month (same week
    number, different (year, month) key).
    Returns (week_col: {(wn, y, m): col}, month_span: {label: (first_week_col, total_col)}, next_free_col)."""
    week_col, month_span = {}, {}
    c = start_col
    for label, weeks in month_groups:
        span_start = c
        for (wn, y, m, _daycount) in weeks:
            week_col[(wn, y, m)] = c
            c += 1
        month_span[label] = (span_start, c)      # c = this month's total column
        c += 1
    return week_col, month_span, c


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
        elif uom == "CFC":
            tbgs_ctn, source, cfc = None, "CFC", qty            # already a CFC count
        else:
            # ENV/POUCH/PKT/BOX/CASE/PCS: qty is a TBGS (tea-bag-equivalent)
            # count directly, same as CTN's qty x TBGS/CTN, minus that step.
            tbgs_ctn, source, cfc = None, uom, qty / tpc
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


def compute_capacity_tbgs(line, machines, shifts):
    """Weekly capacity in TBGS (tea-bag-equivalent) — same as compute_capacity
    but without the ÷TBGS/CFC÷CFC/Cntr step."""
    f = MC_FACTORS.get(line, {})
    rate, shiftmin = f.get("rate", 0), f.get("shiftmin", 0)
    if not (rate and shiftmin and machines and shifts):
        return 0.0
    return rate * shiftmin * machines * shifts * 6 * EFFICIENCY


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
        # More decimal places, for a zero-shift line that still shows a real
        # (usually tiny) achieved value — 0.0 at one decimal would round it
        # away to looking like nothing ran.
        "NUMFMT_CNTR_PRECISE": "_ * #,##0.000_ ;_ * \\-#,##0.000_ ;_ * \\-??_ ;_ @_ ",
        "NUMFMT_PCT":   "0.0\\%",
    }


# ── Generic month+week grouped sheet writer ──────────────────────────────────

def _write_group_sheet(wb, title, month_groups, S, value_fn, *, with_capacity,
                        use_active=False, footer_fn=None, grand_total=False, co_ratio_week=None,
                        tbgs_capacity=False, decimals=3, scale=1, capacity_header=None,
                        full_week_fn=None):
    """One sheet: Zone/Region/Classification/Machine Line/Machines/Shifts, then
    (optionally) a live Capacity column, then value_fn(line, week, year, month)
    grouped by month with a per-month SUM column. A week split across two
    calendar months gets one column per month, both calling value_fn with the
    same week number but a different (year, month) — value_fn must itself
    return only that month's share. Returns ({row: line name}, next_free_row).

    grand_total     : add a trailing "Grand Total" row, summing only the
                       per-classification TOTAL rows (never the data rows
                       themselves) so nothing is double-counted.
    co_ratio_week   : if given, add a trailing "Changeover/MC/Shift
                       (Week N)" column = that week's value ÷ Machines ÷ 12
                       (blank "-" wherever Machines is blank, i.e. on TOTAL
                       rows).
    tbgs_capacity   : capacity formula stops at Rate×ShiftMin×Machines×
                       Shifts×6×90% — no ÷TBGS/CFC÷CFC/Cntr — so the
                       Capacity/Month Capacity columns come out in TBGS
                       instead of containers (value_fn must then also
                       return TBGS amounts).
    decimals        : rounding applied to the achieved-value cells and to
                       the capacity formula.
    scale           : divide the Capacity formula and value_fn's results by
                       this before rounding (e.g. 100000 to show TBGS in
                       lakhs instead of raw tea-bag counts).
    capacity_header : text for the Capacity column header; defaults to
                       CAPACITY_HEADER.
    full_week_fn    : full_week_fn(line, week) returning that week's FULL
                       (unsplit) value — used only for the over-capacity red
                       highlight, so a month-fragment cell is compared
                       against the whole week's total, not just its own
                       partial share. Defaults to value_fn's own column
                       value when not given (fine for weeks that never
                       split)."""
    BORDER = S["BORDER"]
    VALUE_START = 9 if with_capacity else 7

    ws = wb.active if use_active else wb.create_sheet(title)
    ws.title = title

    week_col, month_span, next_col = _block_columns(VALUE_START, month_groups)
    LAST_COL = next_col - 1

    # ── Row 1: title + month headers ─────────────────────────────────────
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=VALUE_START - 1)
    ws.cell(1, 1, title)
    ws.cell(1, 1).font, ws.cell(1, 1).fill = S["TITLE_L_FONT"], S["TITLE_L_FILL"]
    ws.cell(1, 1).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    for label, (span_start, total_col) in month_span.items():
        ws.merge_cells(start_row=1, start_column=span_start, end_row=1, end_column=total_col)
        c = ws.cell(1, span_start, label)
        c.font, c.fill = S["TITLE_C_FONT"], S["TITLE_C_FILL"]
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    # ── Row 2: column headers ────────────────────────────────────────────
    ws.row_dimensions[2].height = 36
    headers = {1: "Zone", 2: "Region", 3: None, 4: "Machine Line",
               5: "No. of\nMachines", 6: "No. of\nShifts"}
    if with_capacity:
        headers[7] = capacity_header or CAPACITY_HEADER
        headers[8] = "Month\nCapacity"
    for c, h in headers.items():
        cell = ws.cell(2, c, h)
        cell.font, cell.fill, cell.border = S["HDR_FONT"], S["HDR_FILL"], BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for (wn, y, m), col in week_col.items():
        cell = ws.cell(2, col, week_label(wn))
        cell.font, cell.fill, cell.border = S["WEEK_HDR_FONT"], S["CNTR_HDR_FILL"], BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for label, (_, total_col) in month_span.items():
        cell = ws.cell(2, total_col, label)
        cell.font, cell.fill, cell.border = S["BOLD"], S["TOTAL_FILL"], BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # ── Column widths ────────────────────────────────────────────────────
    for col, w in COL_WIDTHS.items():
        ws.column_dimensions[col].width = w
    for c in range(5, LAST_COL + 1):
        ws.column_dimensions[get_column_letter(c)].width = WEEK_COL_WIDTH

    if with_capacity:
        LKP_COL = LAST_COL + 3
        for i, h in enumerate(["Rate", "ShMin", "TBGS/CFC", "CFC/Cntr"]):
            ws.cell(2, LKP_COL + i, h).font = S["HELPER_FONT"]
            ws.column_dimensions[get_column_letter(LKP_COL + i)].hidden = True
            ws.column_dimensions[get_column_letter(LKP_COL + i)].width = HELPER_WIDTH
        Qc, Rc, Sc, Tc = (get_column_letter(LKP_COL + i) for i in range(4))

    def month_total_formula(r):
        return {total_col: f"=SUM({get_column_letter(span_start)}{r}:{get_column_letter(total_col - 1)}{r})"
                for label, (span_start, total_col) in month_span.items()}

    # ── Data rows ────────────────────────────────────────────────────────
    row = 3
    pending_merges = []
    cur_a_start = cur_b_start = cur_c_start = None
    cur_zone = cur_region = cur_sub = None
    grp_start = row
    line_rows, total_groups = {}, {}

    for zone, region, sub, line, nm, ns in LAYOUT:
        ws.row_dimensions[row].height = 18

        if zone == "__TOTAL__":
            total_groups[row] = (grp_start, row - 1)
            for c in range(1, LAST_COL + 1):
                cell = ws.cell(row, c)
                cell.border = BORDER
                if c >= 4:
                    cell.fill, cell.font = S["TOTAL_FILL"], S["BOLD"]
            ws.cell(row, 4, "TOTAL")
            if with_capacity:
                f_cell = ws.cell(row, 7, f"=SUM(G{grp_start}:G{row - 1})")
                f_cell.number_format = S["NUMFMT_CNTR"]
                mc_cell = ws.cell(row, 8, f"=G{row}/6*26")
                mc_cell.number_format = S["NUMFMT_CNTR"]
            for col in list(week_col.values()) + [tc for _, tc in month_span.values()]:
                cl = get_column_letter(col)
                a_cell = ws.cell(row, col, f"=SUM({cl}{grp_start}:{cl}{row - 1})")
                a_cell.number_format = S["NUMFMT_CNTR"]
            grp_start = row + 1
            row += 1
            if cur_c_start is not None:
                pending_merges.append(("C", cur_c_start, row - 2))
                cur_c_start = None
                cur_sub = None
            continue

        # zone / region / classification labels
        if zone and zone != cur_zone:
            if cur_a_start is not None:
                pending_merges.append(("A", cur_a_start, row - 1))
            cur_zone, cur_a_start = zone, row
            ws.cell(row, 1, zone).font = S["BOLD"]
            ws.cell(row, 1).fill = S["REGION_FILL"]
        if region and region != cur_region:
            if cur_b_start is not None:
                pending_merges.append(("B", cur_b_start, row - 1))
            cur_region, cur_b_start = region, row
            ws.cell(row, 2, region).font = S["BOLD"]
        if sub and sub != cur_sub:
            if cur_c_start is not None:
                pending_merges.append(("C", cur_c_start, row - 1))
            cur_sub, cur_c_start = sub, row
            ws.cell(row, 3, sub).font = S["NORM"]
            if len(sub) > 10:
                ws.cell(row, 3).alignment = Alignment(wrap_text=True)
        ws.cell(row, 1).border = BORDER
        ws.cell(row, 2).border = BORDER
        ws.cell(row, 3).border = BORDER

        line_rows[row] = line
        ws.cell(row, 4, line).font = S["NORM"]
        ws.cell(row, 4).border = BORDER
        for c, v in ((5, nm), (6, ns)):
            cell = ws.cell(row, c, v)
            cell.font, cell.border = S["NORM"], BORDER
            cell.alignment = Alignment(horizontal="right", vertical="center")

        cap_val = round((compute_capacity_tbgs if tbgs_capacity else compute_capacity)(line, nm, ns) / scale, decimals)
        zero_cap = (ns == 0)

        if with_capacity:
            f = MC_FACTORS.get(line, {})
            for i, key in enumerate(["rate", "shiftmin", "tbgs_cfc", "cfc_cntr"]):
                ws.cell(row, LKP_COL + i, f.get(key, 0)).font = S["HELPER_FONT"]
            # Capacity is always a live formula (never a static "-"), so changing
            # Machines/Shifts for a line that started at 0 shifts still recalculates.
            f_cell = ws.cell(row, 7)
            f_cell.border, f_cell.font = BORDER, S["NORM"]
            f_cell.alignment = Alignment(horizontal="right", vertical="center")
            f_cell.number_format = S["NUMFMT_CNTR"]
            if tbgs_capacity:
                f_cell.value = (
                    f'=IF($F{row}=0,"-",ROUND({Qc}{row}*{Rc}{row}*$E{row}*$F{row}*6*0.9/{scale},{decimals}))')
            else:
                f_cell.value = (
                    f'=IFERROR(IF({Tc}{row}=0,"-",'
                    f'ROUND({Qc}{row}*{Rc}{row}*$E{row}*$F{row}*6*0.9/{Sc}{row}/{Tc}{row},{decimals})),"-")')

            mc_cell = ws.cell(row, 8, f'=IF(G{row}="-","-",G{row}/6*26)')
            mc_cell.border, mc_cell.font, mc_cell.fill = BORDER, S["NORM"], S["REGION_FILL"]
            mc_cell.alignment = Alignment(horizontal="right", vertical="center")
            mc_cell.number_format = S["NUMFMT_CNTR"]

        # A zero-shift line is nominally not running, but production can still
        # show up against it (e.g. a small carry-over run) — that must still
        # display, at higher precision so a tiny amount doesn't round to 0.0.
        row_numfmt = S["NUMFMT_CNTR_PRECISE"] if (with_capacity and zero_cap) else S["NUMFMT_CNTR"]

        for (wn, y, m), col in week_col.items():
            val = value_fn(line, wn, y, m)
            full_val = full_week_fn(line, wn) if full_week_fn else val
            scaled_full = full_val / scale if full_val else full_val
            over = bool(with_capacity and cap_val and scaled_full and scaled_full > cap_val)
            cell = ws.cell(row, col)
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.font = S["OVER_FONT"] if over else S["NORM"]
            if not val:
                cell.value = "-"
            else:
                cell.value = round(val / scale, decimals)
                cell.number_format = row_numfmt

        for total_col, formula in month_total_formula(row).items():
            cell = ws.cell(row, total_col, formula)
            cell.border = BORDER
            cell.font = S["NORM"]
            cell.number_format = row_numfmt
            cell.alignment = Alignment(horizontal="right", vertical="center")
        row += 1

    if cur_a_start is not None:
        pending_merges.append(("A", cur_a_start, row - 1))
    if cur_b_start is not None:
        pending_merges.append(("B", cur_b_start, row - 1))
    if cur_c_start is not None:
        pending_merges.append(("C", cur_c_start, row - 1))
    for cl, ds, de in pending_merges:
        if de > ds:
            ws.merge_cells(f"{cl}{ds}:{cl}{de}")
        wrap = cl == "C" and len(str(ws[f"{cl}{ds}"].value or "")) > 10      # e.g. "DC SF CRIMP"
        ws[f"{cl}{ds}"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True if wrap else None)

    data_end_row = row          # one past the last data/TOTAL row (before any footer/grand-total)

    # ── Trailing "Changeover/MC/Shift (Week N)" column ────────────────────
    co_cols = sorted(col for (wn, y, m), col in week_col.items() if wn == co_ratio_week)
    if co_ratio_week is not None and co_cols:
        extra_col = LAST_COL + 1
        # A split week has two columns (one per month) — sum them to get the
        # week's full changeover count before dividing.
        numerator = "+".join(get_column_letter(c) + "{row}" for c in co_cols)
        header_cell = ws.cell(1, extra_col, f"Changeover/MC/Shift\n(Week {co_ratio_week})")
        ws.merge_cells(start_row=1, start_column=extra_col, end_row=2, end_column=extra_col)
        header_cell.font, header_cell.fill, header_cell.border = S["HDR_FONT"], S["HDR_FILL"], BORDER
        header_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.cell(2, extra_col).border = BORDER
        ws.column_dimensions[get_column_letter(extra_col)].width = 14
        for r in range(3, data_end_row):
            num = numerator.format(row=r)
            cell = ws.cell(r, extra_col, f'=IFERROR(ROUND(({num})/$E{r}/12,2),"-")')
            cell.border = BORDER
            cell.font = S["NORM"]
            cell.alignment = Alignment(horizontal="right", vertical="center")
        LAST_COL = extra_col

    # ── Footer row: all-machines total, including work centers with no line ──
    if footer_fn is not None:
        ws.row_dimensions[row].height = 18
        for (wn, y, m), col in week_col.items():
            val = footer_fn(wn, y, m)
            cell = ws.cell(row, col, "-" if not val else round(val, 3))
            cell.font, cell.border, cell.number_format = S["NORM"], BORDER, S["NUMFMT_CNTR"]
            cell.alignment = Alignment(horizontal="right", vertical="center")
        for total_col, formula in month_total_formula(row).items():
            cell = ws.cell(row, total_col, formula)
            cell.border, cell.font, cell.number_format = BORDER, S["NORM"], S["NUMFMT_CNTR"]
            cell.alignment = Alignment(horizontal="right", vertical="center")
        row += 1

    # ── Grand Total row: sums only the per-classification TOTAL rows, so
    #    nothing already counted inside a TOTAL row is counted twice ──────
    if grand_total:
        ws.row_dimensions[row].height = 18
        total_rows = sorted(total_groups.keys())
        label_cell = ws.cell(row, 1, "Grand Total")
        label_cell.font = S["BOLD"]
        for c in range(1, LAST_COL + 1):
            cell = ws.cell(row, c)
            cell.border = BORDER
            if c >= 4:
                cell.fill, cell.font = S["TOTAL_FILL"], S["BOLD"]
        cols_to_sum = ([7, 8] if with_capacity else []) + list(week_col.values()) + [tc for _, tc in month_span.values()]
        for c in cols_to_sum:
            cl = get_column_letter(c)
            formula = "=" + "+".join(f"{cl}{r}" for r in total_rows)
            cell = ws.cell(row, c, formula)
            cell.font = S["BOLD"]
            cell.number_format = S["NUMFMT_CNTR"]
            cell.alignment = Alignment(horizontal="right", vertical="center")
        row += 1

    ws.freeze_panes = f"{get_column_letter(VALUE_START)}3"
    ws.sheet_view.showGridLines = False
    for c in range(1, LAST_COL + 1):                       # thin frame round the title / month headers
        ws.cell(1, c).border = BORDER

    return line_rows, row


def _write_efficiency_sheet(wb, title, month_groups, S, main_sheet, main_week_col, main_month_span):
    """Capacity Vs Production, in %. Every cell is a live formula referencing
    the main sheet's Capacity (col G) and achieved columns, so it recalculates
    when Machines/Shifts change there — no independent achieved computation."""
    BORDER = S["BORDER"]
    VALUE_START = 7
    ws = wb.create_sheet(title)

    week_col, month_span, next_col = _block_columns(VALUE_START, month_groups)
    LAST_COL = next_col - 1

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=VALUE_START - 1)
    ws.cell(1, 1, title)
    ws.cell(1, 1).font, ws.cell(1, 1).fill = S["TITLE_L_FONT"], S["TITLE_L_FILL"]
    ws.cell(1, 1).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    for label, (span_start, total_col) in month_span.items():
        ws.merge_cells(start_row=1, start_column=span_start, end_row=1, end_column=total_col)
        c = ws.cell(1, span_start, label)
        c.font, c.fill = S["TITLE_C_FONT"], S["TITLE_C_FILL"]
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    ws.row_dimensions[2].height = 36
    for c, h in {1: "Zone", 2: "Region", 3: None, 4: "Machine Line",
                 5: "No. of\nMachines", 6: "No. of\nShifts"}.items():
        cell = ws.cell(2, c, h)
        cell.font, cell.fill, cell.border = S["HDR_FONT"], S["HDR_FILL"], BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for (wn, y, m), col in week_col.items():
        cell = ws.cell(2, col, week_label(wn))
        cell.font, cell.fill, cell.border = S["WEEK_HDR_FONT"], S["CNTR_HDR_FILL"], BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for label, (_, total_col) in month_span.items():
        cell = ws.cell(2, total_col, label)
        cell.font, cell.fill, cell.border = S["BOLD"], S["TOTAL_FILL"], BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for col, w in COL_WIDTHS.items():
        ws.column_dimensions[col].width = w
    for c in range(5, LAST_COL + 1):
        ws.column_dimensions[get_column_letter(c)].width = WEEK_COL_WIDTH

    row = 3
    pending_merges = []
    cur_a_start = cur_b_start = cur_c_start = None
    cur_zone = cur_region = cur_sub = None
    grp_start = row

    # weeks as (wn, y, m) keys (dropping the day-count), plus the equivalent
    # whole-week weight for that month (1.0 for a week entirely in it, a
    # fraction like 2/7 or 5/7 for a week split across the month boundary —
    # used so a split week doesn't get double-weighted as a full week in
    # both months' averages).
    month_weeks = {label: [(wn, y, m) for (wn, y, m, _dc) in weeks] for label, weeks in month_groups}
    month_week_weight = {label: sum(dc / 7 for (_wn, _y, _m, dc) in weeks) for label, weeks in month_groups}

    for zone, region, sub, line, nm, ns in LAYOUT:
        ws.row_dimensions[row].height = 18

        if zone == "__TOTAL__":
            for c in range(1, LAST_COL + 1):
                cell = ws.cell(row, c)
                cell.border = BORDER
                if c >= 4:
                    cell.fill, cell.font = S["TOTAL_FILL"], S["BOLD"]
            ws.cell(row, 4, "TOTAL")
            for (wn, y, m), col in week_col.items():
                mcol = main_week_col[(wn, y, m)]
                cell = ws.cell(row, col,
                    f'=IFERROR(ROUND(SUM(\'{main_sheet}\'!{mcol}{grp_start}:{mcol}{row - 1})/'
                    f'SUMIF(\'{main_sheet}\'!G{grp_start}:G{row - 1},">0")*100,1),0)')
                cell.number_format = S["NUMFMT_PCT"]
            for label, (span_start, total_col) in month_span.items():
                weeks = month_weeks[label]
                weight = month_week_weight[label]
                c1, c2 = main_week_col[weeks[0]], main_week_col[weeks[-1]]
                cell = ws.cell(row, total_col,
                    f'=IFERROR(ROUND(SUM(\'{main_sheet}\'!{c1}{grp_start}:{c2}{row - 1})/'
                    f'(SUMIF(\'{main_sheet}\'!G{grp_start}:G{row - 1},">0")*{weight})*100,1),0)')
                cell.number_format = S["NUMFMT_PCT"]
            grp_start = row + 1
            row += 1
            if cur_c_start is not None:
                pending_merges.append(("C", cur_c_start, row - 2))
                cur_c_start = None
                cur_sub = None
            continue

        if zone and zone != cur_zone:
            if cur_a_start is not None:
                pending_merges.append(("A", cur_a_start, row - 1))
            cur_zone, cur_a_start = zone, row
            ws.cell(row, 1, zone).font = S["BOLD"]
            ws.cell(row, 1).fill = S["REGION_FILL"]
        if region and region != cur_region:
            if cur_b_start is not None:
                pending_merges.append(("B", cur_b_start, row - 1))
            cur_region, cur_b_start = region, row
            ws.cell(row, 2, region).font = S["BOLD"]
        if sub and sub != cur_sub:
            if cur_c_start is not None:
                pending_merges.append(("C", cur_c_start, row - 1))
            cur_sub, cur_c_start = sub, row
            ws.cell(row, 3, sub).font = S["NORM"]
            if len(sub) > 10:
                ws.cell(row, 3).alignment = Alignment(wrap_text=True)
        ws.cell(row, 1).border = BORDER
        ws.cell(row, 2).border = BORDER
        ws.cell(row, 3).border = BORDER

        ws.cell(row, 4, line).font = S["NORM"]
        ws.cell(row, 4).border = BORDER
        for c, v in ((5, nm), (6, ns)):
            cell = ws.cell(row, c, v)
            cell.font, cell.border = S["NORM"], BORDER
            cell.alignment = Alignment(horizontal="right", vertical="center")

        mcap = f"'{main_sheet}'!G{row}"
        for (wn, y, m), col in week_col.items():
            mcol = main_week_col[(wn, y, m)]
            cell = ws.cell(row, col,
                f'=IFERROR(IF({mcap}="-","-",IF({mcap}=0,0,'
                f'ROUND(\'{main_sheet}\'!{mcol}{row}/{mcap}*100,1))),"-")')
            cell.border = BORDER
            cell.font = S["NORM"]
            cell.number_format = S["NUMFMT_PCT"]
            cell.alignment = Alignment(horizontal="right", vertical="center")
        for label, (span_start, total_col) in month_span.items():
            weeks = month_weeks[label]
            weight = month_week_weight[label]
            c1, c2 = main_week_col[weeks[0]], main_week_col[weeks[-1]]
            cell = ws.cell(row, total_col,
                f'=IFERROR(ROUND(SUM(\'{main_sheet}\'!{c1}{row}:{c2}{row})/'
                f'({mcap}*{weight})*100,1),"-")')
            cell.border = BORDER
            cell.font = S["NORM"]
            cell.number_format = S["NUMFMT_PCT"]
            cell.alignment = Alignment(horizontal="right", vertical="center")
        row += 1

    if cur_a_start is not None:
        pending_merges.append(("A", cur_a_start, row - 1))
    if cur_b_start is not None:
        pending_merges.append(("B", cur_b_start, row - 1))
    if cur_c_start is not None:
        pending_merges.append(("C", cur_c_start, row - 1))
    for cl, ds, de in pending_merges:
        if de > ds:
            ws.merge_cells(f"{cl}{ds}:{cl}{de}")
        wrap = cl == "C" and len(str(ws[f"{cl}{ds}"].value or "")) > 10
        ws[f"{cl}{ds}"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True if wrap else None)

    ws.freeze_panes = f"{get_column_letter(VALUE_START)}3"
    ws.sheet_view.showGridLines = False
    for c in range(1, LAST_COL + 1):
        ws.cell(1, c).border = BORDER


# ── Main report builder ──────────────────────────────────────────────────────

def generate_avp_report(pr_bytes, ref_bytes, target_weeks, fg_blend_bytes=None,
                         prod_rows_override=None, changeover_rows_override=None):
    """
    Parameters
    ----------
    pr_bytes       : bytes  Production Register xlsx
    ref_bytes      : bytes  Reference workbook (for TBGS PER CTN)
    target_weeks   : list   e.g. [36, 37, 38, 39]
    fg_blend_bytes : bytes  FG Blend Names xlsx (defaults to the FG_Blend_Names.xlsx next to this file)
    prod_rows_override       : pre-extracted extract_prod_register() rows to use
                                instead of extracting pr_bytes (e.g. merged from more
                                than one Production Register) — pr_bytes is still
                                needed for register_date_label().
    changeover_rows_override : pre-extracted changeover.read_production_rows() rows
                                to use instead of reading pr_bytes, for the same reason.

    Returns
    -------
    (bytes, set, dict|None) — xlsx bytes, product names that could not be converted,
                              changeover summary (None when no FG Blend Names file is available)

    Produces five visible sheets — Capacity Vs Production, Capacity Vs
    Production (in TBGS), Efficiency, Blend Changeover, Brand Changeover —
    each grouped Zone/Region/Classification/Machine Line, with weeks grouped
    under their month and a per-month SUM column. Efficiency is capacity vs
    achieved in %, as live formulas referencing the Capacity Vs Production
    sheet.
    """
    target_weeks = list(target_weeks)
    if fg_blend_bytes is None and os.path.exists(DEFAULT_FG_BLEND_FILE):
        with open(DEFAULT_FG_BLEND_FILE, "rb") as fh:
            fg_blend_bytes = fh.read()
    with_co = fg_blend_bytes is not None

    item_routing = {}
    if os.path.exists(DEFAULT_ITEM_ROUTING_FILE):
        item_routing = load_item_routing(DEFAULT_ITEM_ROUTING_FILE)

    tbgs_per_ctn = load_tbgs_per_ctn(ref_bytes)
    prod_rows = prod_rows_override if prod_rows_override is not None else extract_prod_register(pr_bytes, item_routing)
    entries, skipped = prod_entries(prod_rows, tbgs_per_ctn, set(target_weeks))

    # Full-week totals (for the over-capacity check and Observations) plus a
    # split-by-month version (for the actual per-column values — a week that
    # spans two calendar months gets its own column, and its own slice of
    # production, in each).
    achieved = defaultdict(float)
    achieved_split = defaultdict(float)
    achieved_tbgs = defaultdict(float)
    achieved_tbgs_split = defaultdict(float)
    for e in entries:
        achieved[(e["group"], e["wn"])] += e["containers"]
        achieved_split[(e["group"], e["wn"], e["date"].year, e["date"].month)] += e["containers"]
        tbgs = e["cfc"] * e["tpc"]
        achieved_tbgs[(e["group"], e["wn"])] += tbgs
        achieved_tbgs_split[(e["group"], e["wn"], e["date"].year, e["date"].month)] += tbgs

    month_groups = _month_groups(target_weeks, WEEK1_MONDAY)
    S = _styles()
    wb = openpyxl.Workbook()

    line_rows, cap_end_row = _write_group_sheet(
        wb, SHEET_NAME, month_groups, S,
        value_fn=lambda line, wn, y, m: achieved_split.get((line, wn, y, m), 0),
        full_week_fn=lambda line, wn: achieved.get((line, wn), 0),
        with_capacity=True, use_active=True, grand_total=True)

    _write_group_sheet(
        wb, "Capacity Vs Production (TBGS)", month_groups, S,
        value_fn=lambda line, wn, y, m: achieved_tbgs_split.get((line, wn, y, m), 0),
        full_week_fn=lambda line, wn: achieved_tbgs.get((line, wn), 0),
        with_capacity=True, grand_total=True, tbgs_capacity=True, decimals=2,
        scale=100000, capacity_header="Capacity\n(TBGS ÷ 1,00,000)")

    main_week_col_n, main_month_span_n, _ = _block_columns(9, month_groups)
    main_week_col = {k: get_column_letter(c) for k, c in main_week_col_n.items()}
    main_month_span = {label: (get_column_letter(s), get_column_letter(e))
                        for label, (s, e) in main_month_span_n.items()}
    _write_efficiency_sheet(wb, "Efficiency", month_groups, S, SHEET_NAME, main_week_col, main_month_span)

    co_info, model, by_lw = None, None, None
    if with_co:
        report_lines = list(dict.fromkeys(line_rows.values()))
        if changeover_rows_override is not None:
            model = changeover.build_model_from_rows(changeover_rows_override, fg_blend_bytes, ROUTING_MAP, report_lines)
        else:
            model = changeover.build_model(pr_bytes, fg_blend_bytes, ROUTING_MAP, report_lines)
        changeover.write_changeover_sheets(wb, model, report_lines, WEEK1_MONDAY)

        by_lw = model.by_line_week(weeknum)                    # full-week (Observations use this)
        by_w = model.by_week(weeknum)
        by_lwm = model.by_line_week_month(weeknum)             # split by month (per-column values)
        by_wm = model.by_week_month(weeknum)
        _, blend_end_row = _write_group_sheet(
            wb, "Blend Changeover", month_groups, S,
            value_fn=lambda line, wn, y, m: by_lwm.get((line, wn, y, m), (0, 0))[0],
            full_week_fn=lambda line, wn: by_lw.get((line, wn), (0, 0))[0],
            with_capacity=False, footer_fn=lambda wn, y, m: by_wm.get((wn, y, m), (0, 0))[0],
            co_ratio_week=target_weeks[-1])
        _, brand_end_row = _write_group_sheet(
            wb, "Brand Changeover", month_groups, S,
            value_fn=lambda line, wn, y, m: by_lwm.get((line, wn, y, m), (0, 0))[1],
            full_week_fn=lambda line, wn: by_lw.get((line, wn), (0, 0))[1],
            with_capacity=False, footer_fn=lambda wn, y, m: by_wm.get((wn, y, m), (0, 0))[1],
            co_ratio_week=target_weeks[-1])

        blend_tot, brand_tot = model.totals
        co_info = {
            "blend_total": blend_tot,
            "brand_total": brand_tot,
            "unknown_items": model.unknown_items,          # FG items with no blend in the FG list
            "not_in_report": model.not_in_report(),        # work centers with no report line
            "machine_shifts": len(model.calc),
        }

    last_week = target_weeks[-1]
    _append_efficiency_observations(wb[SHEET_NAME], S, achieved, last_week, cap_end_row)
    if with_co:
        _append_changeover_observations(wb["Blend Changeover"], S, by_lw, 0, "Blend", last_week, blend_end_row)
        _append_changeover_observations(wb["Brand Changeover"], S, by_lw, 1, "Brand", last_week, brand_end_row)

    # ── Source Calculations (hidden) ────────────────────────────────────────
    _write_source_calculations(wb, entries, target_weeks, S, register_date_label(pr_bytes))
    hidden_sheets = [SOURCE_SHEET]
    order = [SHEET_NAME, "Capacity Vs Production (TBGS)", "Efficiency", "Blend Changeover", "Brand Changeover", SOURCE_SHEET]
    if with_co:
        hidden_sheets += [changeover.SHEET_CALC, changeover.SHEET_DETAIL,
                          changeover.SHEET_ROUTING, changeover.SHEET_FG]
        order += [changeover.SHEET_CALC, changeover.SHEET_DETAIL,
                 changeover.SHEET_ROUTING, changeover.SHEET_FG]
    for name in hidden_sheets:
        if name in wb.sheetnames:
            wb[name].sheet_state = "hidden"
    wb._sheets = [wb[n] for n in order if n in wb.sheetnames]
    wb.active = 0

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue(), skipped, co_info


# ── Observations sheet ───────────────────────────────────────────────────────

def _append_observation_block(ws, S, start_row, heading, points, empty_msg):
    """Write one observation section (heading + numbered points) starting
    one blank row below start_row (the sheet's own last used row)."""
    row = start_row + 1
    ws.merge_cells(f"A{row}:F{row}")
    cell = ws.cell(row, 1, heading)
    cell.font = Font(name=S["FN"], italic=True, size=10, color="FF444444")
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row].height = 18
    row += 1
    if not points:
        ws.merge_cells(f"A{row}:F{row}")
        ws.cell(row, 1, empty_msg).font = Font(name=S["FN"], size=11)
        ws.row_dimensions[row].height = 18
        row += 1
    else:
        for i, text in enumerate(points, 1):
            ws.merge_cells(f"A{row}:F{row}")
            cell = ws.cell(row, 1, f"{i}. {text}")
            cell.font = Font(name=S["FN"], size=11)
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            ws.row_dimensions[row].height = 18
            row += 1
    return row


def _append_efficiency_observations(ws, S, achieved, last_week, end_row):
    """5 points on the worst-performing ran machines (lowest % of capacity)
    in last_week only, appended below the Capacity Vs Production table."""
    stats = []
    for zone, region, sub, line, nm, ns in LAYOUT:
        if zone == "__TOTAL__":
            continue
        cap_week = compute_capacity(line, nm, ns)
        ach = achieved.get((line, last_week), 0)
        if not cap_week or ach <= 0:
            continue
        pct = ach / cap_week * 100
        stats.append((line, pct, ach, cap_week))
    stats.sort(key=lambda t: t[1])
    worst = stats[:5]

    _append_observation_block(
        ws, S, end_row,
        f"Poorest-performing machines, Week {last_week} (ran machines only, by % of capacity)",
        [f"{line} ran at only {pct:.1f}% of capacity in Week {last_week} "
         f"({ach:.1f} of {cap_week:.1f} containers achievable)."
         for line, pct, ach, cap_week in worst],
        "No machine ran below capacity in this week.")


def _append_changeover_observations(ws, S, by_lw, idx, kind, last_week, end_row):
    """Top 3 lines by Blend (idx=0) or Brand (idx=1) changeover count in
    last_week only, appended below the changeover table."""
    totals = defaultdict(int)
    for zone, region, sub, line, nm, ns in LAYOUT:
        if zone == "__TOTAL__":
            continue
        totals[line] += by_lw.get((line, last_week), (0, 0))[idx]
    co_list = sorted(((line, c) for line, c in totals.items() if c > 0), key=lambda t: -t[1])[:3]

    points = []
    for i, (line, count) in enumerate(co_list):
        tag = " — the most of any line" if i == 0 else ""
        points.append(f"{line} had {count} {kind} changeovers in Week {last_week}{tag}.")

    _append_observation_block(
        ws, S, end_row,
        f"Highest {kind} changeovers, Week {last_week}",
        points, f"No {kind} changeovers in this week.")


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
    for zone, region, sub, line, nm, ns in LAYOUT:
        if zone == "__TOTAL__":
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
            elif e["uom"] == "CFC":
                formula = f'{e["qty"]:,}CFC÷{e["cpc"]}CFC/Cntr = {shown:.4f}'
            else:
                formula = (f'{e["qty"]:,}{e["uom"]}(TBGS)÷{e["tpc"]}TBGS/CFC'
                           f'÷{e["cpc"]}CFC/Cntr = {shown:.4f}')
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
