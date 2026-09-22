"""
Achievable Vs Produced — Streamlit App
Upload Production Register → download report
"""

import datetime
import pathlib

import streamlit as st

from avp_generator import generate_avp_report, weeknum, week_label

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Achievable Vs Produced Report",
    page_icon="📊",
    layout="centered",
)

# ── Reference workbook (bundled with repo) ────────────────────────────────────
REF_PATH = pathlib.Path(__file__).parent / "reference_workbook.xlsx"

@st.cache_data(show_spinner=False)
def _load_ref():
    return REF_PATH.read_bytes()

ref_bytes = _load_ref()

# ── UI ────────────────────────────────────────────────────────────────────────
st.title("📊 Achievable Vs Produced Report")
st.markdown(
    "Upload the **Production Register** to generate the weekly achieved "
    "capacity report — by container and by percentage."
)

col1, col2 = st.columns([2, 1])
with col1:
    pr_file = st.file_uploader(
        "Upload Production Register (.xlsx)",
        type=["xlsx"],
        key="pr",
    )
with col2:
    as_of = st.date_input(
        "As-Of Date",
        value=datetime.date.today(),
        help="Used to determine the current week number.",
    )

# ── Week selector ─────────────────────────────────────────────────────────────
current_wn = weeknum(as_of)

with st.expander("⚙️ Week selection", expanded=True):
    n_weeks = st.slider(
        "Number of past weeks to include",
        min_value=2, max_value=8, value=4, step=1,
    )
    first_week = current_wn - n_weeks
    target_weeks = list(range(first_week, current_wn))

    st.markdown(
        "**Weeks included:** " +
        "  ·  ".join(f"W #{wn}" for wn in target_weeks)
    )

# ── Generate ──────────────────────────────────────────────────────────────────
st.divider()
ready = pr_file is not None
btn = st.button("🚀 Generate Report", disabled=not ready, type="primary")

if not ready:
    st.info("Please upload the Production Register to enable the button.")

if btn and ready:
    with st.spinner("Converting production data → containers → report…"):
        try:
            result_bytes, skipped = generate_avp_report(
                pr_bytes=pr_file.read(),
                ref_bytes=ref_bytes,
                target_weeks=target_weeks,
            )

            filename = f"Achievable_Vs_Produced_{as_of.strftime('%d-%b-%Y')}.xlsx"

            st.success("✅ Report generated successfully!")

            st.download_button(
                label="⬇️ Download Report",
                data=result_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            if skipped:
                skipped_sorted = sorted(skipped)
                st.warning(
                    f"⚠️ **{len(skipped_sorted)} product(s) could not be converted** "
                    f"(no TBGS/CTN value found — excluded from achieved totals):\n\n"
                    + "\n".join(f"- {n}" for n in skipped_sorted[:20])
                    + ("\n- …and more" if len(skipped_sorted) > 20 else "")
                )

        except Exception as exc:
            st.error(f"❌ Error generating report:\n\n```\n{exc}\n```")
            raise

# ── Footer ────────────────────────────────────────────────────────────────────
st.divider()
st.caption(
    "**Conversion chain:** Production Qty (CTN) × TBGS/CTN ÷ TBGS/CFC ÷ CFC/Container "
    "= Achieved Containers.  "
    "**Capacity** = Rate × ShiftMin × Machines × Shifts × 6 days × 90% ÷ TBGS/CFC ÷ CFC/Container.  "
    "Machines and Shifts are editable in the downloaded report — Capacity and % update automatically."
)
