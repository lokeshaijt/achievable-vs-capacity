# Achievable Vs Produced — Report Generator

Streamlit app: upload the Production Register → download the **Achievable Vs Produced** report.

## What the report contains

Single sheet (`achievable vs produced`) with two side-by-side analyses:

| Cols A–F | Cols G–J | Cols K–N |
|---|---|---|
| Region / Sub / Machine Line / Machines / Shifts / Capacity | **By container** — achieved production in containers | **By Percentage** — achieved ÷ capacity × 100% |

- **Capacity column (F)** is a live formula — change D (Machines) or E (Shifts) and F recalculates instantly
- **Percentage columns (K–N)** are live formulas referencing F — they update automatically when capacity changes
- **TOTAL rows** use weighted average: SUM(achieved) ÷ SUM(capacity) × 100
- **Over-capacity** values shown in orange-red font

## Files

```
app.py                          — Streamlit UI
avp_generator.py                — All generation logic
mailer.py                       — SMTP email delivery for the generated report
reference_workbook.xlsx         — TBGS PER CTN lookup table (1549 entries)
.streamlit/secrets.toml.example — Template for SMTP credentials
requirements.txt
README.md
```

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

## Deploy on Streamlit Cloud

1. Push this repo to GitHub (public, or connect your account)
2. Go to [share.streamlit.io](https://share.streamlit.io) → **New app**
3. Select repo, branch `main`, main file `app.py`
4. Add your SMTP secrets under **Advanced settings → Secrets** (see above)
5. Click **Deploy**

## How to use

1. Open the app
2. Upload today's **Production Register** (.xlsx)
3. Set As-Of Date and number of past weeks (default: 4)
4. Click **Generate Report** → download, or email it directly

## Conversion chain

```
Production Register (CTN)
  × TBGS per CTN  (from reference_workbook TBGS PER CTN, or derived from product name)
  ÷ TBGS per CFC  (from MC_FACTORS in avp_generator.py)
  ÷ CFC per Container
= Achieved containers / week

Achievable Capacity
  = Rate × ShiftMin × Machines × Shifts × 6 days × 90%
  ÷ TBGS/CFC ÷ CFC/Container

% = Achieved ÷ Capacity × 100  (whole number, rounded)
```

## Updating master data

If machine-line factors (rate, shiftmin, TBGS/CFC, CFC/Container) change,
update the `MC_FACTORS` dictionary in `avp_generator.py` and push — Streamlit Cloud
auto-redeploys within a minute.

If new products need TBGS/CTN values, replace `reference_workbook.xlsx` with
the latest master workbook.
