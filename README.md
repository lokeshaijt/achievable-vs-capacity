# Capacity vs 4 Weeks production — Report Generator

Streamlit app: upload the Production Register → download the **Capacity Vs Produced in last N weeks**
report in the finalized format (sheet name, title, column widths and colour palette are fixed in the code).

## The workbook

| Sheet | Visible | What it is |
|---|---|---|
| `Capacity Vs Production` | yes | capacity vs achieved containers |
| `Blend Changeover` | yes | blend changeovers per machine line |
| `Brand Changeover` | yes | brand changeovers per machine line |
| `Source Calculations` | hidden | how every capacity and every achieved figure is derived |
| `Changeover Calc` · `Changeover Detail` · `CO Routing` · `FG Blend Map` | hidden | the changeover calculation (right-click a tab → Unhide) |

### Each visible sheet

Rows: Zone / Region / Classification / Machine Line / Machines / Shifts (+ **Capacity (In FCL)** on the
Capacity Vs Production sheet only).

Columns: weeks are grouped under their calendar month, each month followed by a SUM column for that
month (e.g. `W #39 | W #40 | SEP`). The Blend/Brand Changeover sheets also carry an unlabeled footer row —
the changeover total for **all machines**, including work centers that have no line in the report.

- Rows are grouped and sorted **Zone 1 → Zone 2 → Zone 3**; within a zone, regions and classifications
  keep the order set in `LAYOUT` (`avp_generator.py`) — identical row order on all three sheets.
- Machines and Shifts are the editable inputs. **Capacity** is always a live formula — including lines
  that start at 0 Shifts — so it, and every TOTAL/month-sum cell, recalculate when you change them.
- TOTAL rows and month-sum columns are `SUM` formulas.
- Values over capacity are shown in red (Capacity Vs Production only).
- Capacity factors are hidden in the columns just past the last month-sum column.

## Blend & Brand changeover

Automates the manual Excel process (Sheet1 → Sheet6 + the two pivots):

```
1. Production Register        keep Transaction Type = "Production"
2. Unique Date / Shift / Work Center / Machine No / Item Name
3. Blend name                 lookup of Item Name in FG Blend Names (first match, like the VLOOKUP)
4. Per Date / Shift / Work Center / Machine No
       Brand changeover = distinct FG items  − 1
       Blend changeover = distinct blends    − 1
5. Sum by machine line and week
```

- **New FG items:** an FG item that is not in `FG_Blend_Names.xlsx` is treated as its own blend (it can only add a
  changeover, never hide one) and is listed in a warning. Replace the bundled file, or upload the latest one in the app.
- **Work centers with no report line** (currently PAKONA, MD20 ENV GARANT, C250F3 HS, …) are counted only in the
  footer row of the two changeover sheets. To give one a line, unhide `CO Routing` and change its *Report Line* —
  everything recalculates.
- The count follows the manual method exactly: a machine is identified by Date + Shift + Work Center + Machine No,
  so a machine that changes work center inside a shift is not counted as a changeover.

## Email the report

After generating a report, expand **📧 Send this report by email**, enter one
or more recipient addresses (comma-separated), and click **Send Email** — the
report is attached and sent via SMTP.

This needs SMTP credentials configured as Streamlit secrets (not the
recipient's email — that's just typed into the box each time):

1. Copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml` locally,
   or paste its contents into the app's **Secrets** settings on Streamlit Cloud.
2. Fill in `server`, `port`, `username`, `password` (and optionally `sender`).
   For Gmail, use an **App Password** (Google Account → Security → 2-Step
   Verification → App passwords), not your normal password.
3. Never commit the real `secrets.toml` — it's already in `.gitignore`.

## Files

```
app.py                          — Streamlit UI (JAY branding, generate, download, email)
avp_generator.py                — layout, capacity factors, routing, report writer
changeover.py                   — Blend / Brand changeover logic + hidden sheets
mailer.py                       — SMTP email delivery for the generated report
reference_workbook.xlsx         — TBGS PER CTN lookup table
FG_Blend_Names.xlsx             — FG item → Blend name (used for changeovers)
MC_Routing_Master.xlsx          — item name → machine line (overrides work-center routing)
assets/jay_logo.jpg             — JAY logo (header + favicon)
.streamlit/config.toml          — black/gold theme
.streamlit/secrets.toml.example — Template for SMTP credentials
requirements.txt
README.md
```

## Deploy on Streamlit Cloud

1. Push this repo to GitHub
2. [share.streamlit.io](https://share.streamlit.io) → **New app** → branch `main`, main file `app.py`
3. Add your SMTP secrets under **Advanced settings → Secrets** (see above)
4. **Deploy**

## How to use

1. Upload the **Production Register** (.xlsx)
2. Set the As-Of Date and the number of past weeks (default 4)
3. (Optional) upload a newer **FG Blend Names** file if new FG items were added
4. **Generate Report** → download `Capacity_vs_4_weeks_production_W<first>_<last>.xlsx`
   (name follows the weeks chosen), or email it directly

## Conversion chain

```
Production Register (CTN)
  × TBGS per CTN  (reference_workbook TBGS PER CTN, or read from the product name e.g. "20 DC SF" → 20)
  ÷ TBGS per CFC  (MC_FACTORS)
  ÷ CFC per Container
= Achieved containers / week

Capacity = Rate × ShiftMin × Machines × Shifts × 6 days × 90% ÷ TBGS/CFC ÷ CFC/Container
% = Achieved ÷ Capacity × 100  (one decimal)
```

## Updating master data

- **Machines / shifts / lines:** `LAYOUT` in `avp_generator.py`.
- **Capacity factors:** `MC_FACTORS` in `avp_generator.py`.
- **Item name → line (precise, takes priority):** replace `MC_Routing_Master.xlsx` — a "PRODUCT NAME"/"MC ROUTING"
  sheet. A routing name not recognized by `ITEM_MASTER_LINE_MAP` in `avp_generator.py` is ignored for that item,
  which then falls back to the work-center map below.
- **Work center → line (fallback):** `ROUTING_MAP` in `avp_generator.py`.
- **Look and feel** (sheet name, title, widths, colours): the constants at the top of `avp_generator.py` and `_styles()`.

Push the change — Streamlit Cloud redeploys within a minute.
