"""
Capacity vs 4 Weeks production — Streamlit App
Upload Production Register → download report (finalized format)
"""

import base64
import datetime
import pathlib

import streamlit as st
from PIL import Image

from avp_generator import generate_avp_report, weeknum, week_label
from mailer import EmailNotConfigured, send_report_email

# ── Brand assets ──────────────────────────────────────────────────────────────
LOGO_PATH = pathlib.Path(__file__).parent / "assets" / "jay_logo.jpg"

@st.cache_data(show_spinner=False)
def _load_logo_b64():
    return base64.b64encode(LOGO_PATH.read_bytes()).decode()

logo_b64 = _load_logo_b64()

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Capacity vs 4 Weeks production",
    page_icon=Image.open(LOGO_PATH),
    layout="centered",
)

# ── Brand styling (JAY black / gold palette) ──────────────────────────────────
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;700&display=swap');

    html, body, [class*="css"] { font-family: 'Poppins', sans-serif; }

    .jay-header {
        display: flex;
        align-items: center;
        gap: 20px;
        background: linear-gradient(135deg, #1E1B12 0%, #000000 100%);
        border: 1px solid #D9A526;
        border-radius: 16px;
        padding: 18px 26px;
        margin-bottom: 24px;
        box-shadow: 0 0 20px rgba(217, 165, 38, 0.15);
    }
    .jay-header .jay-logo-frame {
        background: #ffffff;
        border-radius: 50%;
        padding: 4px;
        display: flex;
        box-shadow: 0 0 0 2px #D9A526;
        flex-shrink: 0;
    }
    .jay-header .jay-logo-frame img {
        width: 56px;
        height: 56px;
        border-radius: 50%;
        display: block;
    }
    .jay-header h1 {
        color: #F2C94C;
        font-size: 1.6rem;
        font-weight: 700;
        margin: 0;
        line-height: 1.25;
    }
    .jay-header p {
        color: #E8E1CF;
        font-size: 0.92rem;
        margin: 4px 0 0 0;
    }

    hr { border-top: 1px solid #3A331E; }

    div[data-testid="stButton"] button,
    div[data-testid="stDownloadButton"] button {
        background: linear-gradient(135deg, #F2C94C 0%, #D9A526 100%);
        color: #141414;
        font-weight: 600;
        border: none;
        border-radius: 8px;
    }
    div[data-testid="stButton"] button:hover,
    div[data-testid="stDownloadButton"] button:hover {
        background: linear-gradient(135deg, #F7D96B 0%, #E6B62C 100%);
        color: #000000;
    }
    div[data-testid="stButton"] button:disabled {
        background: #2A2517;
        color: #8C7F55;
        border: 1px solid #4A4126;
        opacity: 1;
    }

    div[data-testid="stSlider"] > div > div > div {
        background: #4A4126 !important;
    }
    div[data-testid="stSlider"] > div > div > div > div {
        background: #D9A526 !important;
    }

    div[data-testid="stExpander"] {
        border: 1px solid #3A331E;
        border-radius: 10px;
    }

    div[data-testid="stFileUploaderDropzone"],
    div[data-testid="stDateInput"] input {
        border: 1px solid #3A331E !important;
    }

    .jay-footer {
        color: #A99B6E;
        font-size: 0.85rem;
        border-top: 1px solid #3A331E;
        padding-top: 12px;
        margin-top: 8px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ── Reference workbook (bundled with repo) ────────────────────────────────────
REF_PATH = pathlib.Path(__file__).parent / "reference_workbook.xlsx"

@st.cache_data(show_spinner=False)
def _load_ref():
    return REF_PATH.read_bytes()

ref_bytes = _load_ref()

# ── FG Blend Names (bundled; can be replaced by an upload) ────────────────────
FG_PATH = pathlib.Path(__file__).parent / "FG_Blend_Names.xlsx"

@st.cache_data(show_spinner=False)
def _load_fg():
    return FG_PATH.read_bytes()

# ── UI: branded header ─────────────────────────────────────────────────────────
st.markdown(
    f"""
    <div class="jay-header">
        <div class="jay-logo-frame">
            <img src="data:image/jpeg;base64,{logo_b64}" alt="JAY logo" />
        </div>
        <div>
            <h1>📊 Capacity vs 4 Weeks production</h1>
            <p>Upload the Production Register to generate the <i>Capacity Vs Produced</i> report —
            capacity, achieved containers and % for each machine line and week — with the
            <b>Blend</b> and <b>Brand changeovers</b>.</p>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
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

with st.expander("🧪 Blend names for changeovers (FG Blend Names)", expanded=False):
    st.markdown(
        "Blend names come from the **FG Blend Names** file bundled with the app. "
        "If new FG items were added, upload the latest file here to use it instead."
    )
    fg_file = st.file_uploader("FG Blend Names (.xlsx) — optional", type=["xlsx"], key="fg")

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
            result_bytes, skipped, co = generate_avp_report(
                pr_bytes=pr_file.read(),
                ref_bytes=ref_bytes,
                target_weeks=target_weeks,
                fg_blend_bytes=fg_file.read() if fg_file is not None else _load_fg(),
            )

            st.session_state["report_bytes"] = result_bytes
            st.session_state["report_filename"] = (
                f"Capacity_vs_{len(target_weeks)}_weeks_production_"
                f"W{target_weeks[0]}_{target_weeks[-1]}.xlsx"
            )
            st.session_state["report_skipped"] = skipped
            st.session_state["report_co"] = co

        except Exception as exc:
            st.error(f"❌ Error generating report:\n\n```\n{exc}\n```")
            raise

# ── Result: download + email ──────────────────────────────────────────────────
if "report_bytes" in st.session_state:
    st.success("✅ Report generated successfully!")

    st.download_button(
        label="⬇️ Download Report",
        data=st.session_state["report_bytes"],
        file_name=st.session_state["report_filename"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    co = st.session_state.get("report_co")
    if co:
        m1, m2, m3 = st.columns(3)
        m1.metric("Blend changeovers (all weeks in register)", co["blend_total"])
        m2.metric("Brand changeovers (all weeks in register)", co["brand_total"])
        m3.metric("Machine-shifts analysed", co["machine_shifts"])
        if co["unknown_items"]:
            st.warning(
                f"⚠️ **{len(co['unknown_items'])} FG item(s) have no blend in the FG Blend Names file** — "
                "each is treated as its own blend (so it counts as a changeover). "
                "Upload an updated FG Blend Names file:\n\n"
                + "\n".join(f"- {n}" for n in co["unknown_items"][:20])
                + ("\n- …and more" if len(co["unknown_items"]) > 20 else "")
            )
        unmapped = {wc: v for wc, v in co["not_in_report"].items() if v != (0, 0)}
        if unmapped:
            st.info(
                "ℹ️ Changeovers on work centers that have no machine line in the report are only in the "
                "all-machines total row under the changeover columns. To give them a line, unhide the "
                "**CO Routing** sheet (right-click a tab → Unhide) and re-map them:\n\n"
                + "\n".join(f"- {wc}: blend {b}, brand {r}" for wc, (b, r) in unmapped.items())
            )

    skipped = st.session_state.get("report_skipped")
    if skipped:
        skipped_sorted = sorted(skipped)
        st.warning(
            f"⚠️ **{len(skipped_sorted)} product(s) could not be converted** "
            f"(no TBGS/CTN value found — excluded from achieved totals):\n\n"
            + "\n".join(f"- {n}" for n in skipped_sorted[:20])
            + ("\n- …and more" if len(skipped_sorted) > 20 else "")
        )

    with st.expander("📧 Send this report by email"):
        to_email = st.text_input(
            "Recipient email(s)",
            placeholder="name@example.com, name2@example.com",
            help="Comma-separate multiple addresses.",
            key="to_email",
        )
        send_clicked = st.button("📧 Send Email")

        if send_clicked:
            if not to_email.strip():
                st.error("Enter at least one recipient email address.")
            else:
                with st.spinner("Sending email…"):
                    try:
                        sent_to = send_report_email(
                            to_addrs=to_email,
                            subject=(
                                f"Capacity vs {len(target_weeks)} Weeks production — "
                                f"W{target_weeks[0]}-{target_weeks[-1]}"
                            ),
                            body=(
                                "Please find attached the Capacity vs 4 Weeks production report "
                                f"for weeks {target_weeks[0]}–{target_weeks[-1]}."
                            ),
                            attachment_bytes=st.session_state["report_bytes"],
                            attachment_filename=st.session_state["report_filename"],
                        )
                        st.success(f"✅ Email sent to {', '.join(sent_to)}")
                    except EmailNotConfigured as exc:
                        st.error(f"❌ {exc}")
                    except Exception as exc:
                        st.error(f"❌ Failed to send email:\n\n```\n{exc}\n```")

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown(
    """
    <div class="jay-footer">
    The workbook also holds the hidden sheets <i>Source Calculations</i> and the changeover sheets
    (right-click a tab → Unhide). Machines and Shifts are editable in the downloaded report —
    Capacity, % and the TOTAL rows update automatically.
    <b>Changeovers:</b> per machine, per shift — Brand = distinct FG items − 1,
    Blend = distinct blends − 1 (Production rows only), summed by machine line and week.
    </div>
    """,
    unsafe_allow_html=True,
)
