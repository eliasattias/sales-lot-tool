from pathlib import Path
from tempfile import NamedTemporaryFile
import base64

import streamlit as st

from sales_tool import process_workbook


LOGO_PATH = Path(__file__).parent / "assets" / "sensimedical-logo.png"

SENSIMEDICAL_CSS = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');

    /* ─── Reset & Base ───────────────────────────────────── */
    html, body, .stApp {
        background-color: #f4f6f9 !important;
        font-family: 'DM Sans', sans-serif !important;
    }

    /* ─── Hide Streamlit chrome ──────────────────────────── */
    #MainMenu, footer, header { visibility: hidden; }
    [data-testid="stSidebar"] { display: none; }
    [data-testid="stSidebar"] ~ div { margin-left: 0 !important; }
    [data-testid="stDecoration"] { display: none; }

    /* ─── Nuke ALL white box backgrounds Streamlit injects ── */
    .stApp > div,
    [data-testid="stAppViewContainer"],
    [data-testid="stAppViewBlockContainer"],
    [data-testid="stVerticalBlock"],
    [data-testid="stVerticalBlockBorderWrapper"],
    [data-testid="stHorizontalBlock"],
    [data-testid="column"],
    section[data-testid="stSidebar"],
    div[class*="block-container"] {
        background: transparent !important;
        box-shadow: none !important;
        border: none !important;
    }

    /* ─── Top Navigation Bar ─────────────────────────────── */
    .sm-navbar {
        position: fixed;
        top: 0; left: 0; right: 0;
        z-index: 999;
        background: #0c1f3a;
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 0 2rem;
        height: 68px;
        border-bottom: 1px solid rgba(255,255,255,0.06);
        box-shadow: 0 2px 16px rgba(0,0,0,0.25);
    }
    .sm-navbar-brand {
        display: flex;
        align-items: center;
        gap: 10px;
    }
    /* Force logo white so it blends with navy */
    .sm-navbar-brand img {
        height: 48px;
        width: auto;
        filter: brightness(0) invert(1);
    }
    .sm-navbar-title {
        font-family: 'DM Sans', sans-serif;
        font-weight: 600;
        font-size: 0.85rem;
        letter-spacing: 0.12em;
        text-transform: uppercase;
        color: rgba(255,255,255,0.55);
    }
    .sm-navbar-center {
        position: absolute;
        left: 50%;
        transform: translateX(-50%);
        color: #ffffff;
        font-family: 'DM Sans', sans-serif;
        font-weight: 600;
        font-size: 1rem;
        letter-spacing: -0.01em;
        pointer-events: none;
    }

    /* ─── Main content offset for fixed nav ─────────────── */
    .main .block-container {
        padding-top: 5.5rem !important;
        padding-left: 2.5rem !important;
        padding-right: 2.5rem !important;
        max-width: 1400px !important;
    }

    /* ─── Page heading ───────────────────────────────────── */
    .sm-page-header {
        margin-bottom: 1.4rem;
        border-bottom: 1px solid #e2e8f0;
        padding-bottom: 0.9rem;
    }
    .sm-page-header h1 {
        font-family: 'DM Sans', sans-serif !important;
        font-size: 1.65rem !important;
        font-weight: 600 !important;
        color: #0c1f3a !important;
        margin: 0 0 0.2rem 0 !important;
        letter-spacing: -0.02em;
    }
    .sm-page-header p {
        color: #64748b;
        font-size: 0.88rem;
        margin: 0;
    }

    /* ─── Cards — applied to column wrappers ─────────────── */
    .sm-col-card {
        background: #ffffff !important;
        border: 1px solid #e2e8f0 !important;
        border-radius: 10px !important;
        padding: 1.4rem 1.6rem !important;
        box-shadow: 0 1px 4px rgba(0,0,0,0.05) !important;
    }
    .sm-col-card p, .sm-col-card ul, .sm-col-card li {
        font-family: 'DM Sans', sans-serif;
    }
    .sm-card-title {
        font-size: 0.95rem;
        font-weight: 600;
        color: #0c1f3a;
        margin: 0 0 0.35rem 0;
    }
    .sm-card-text {
        font-size: 0.85rem;
        color: #64748b;
        margin: 0 0 0.9rem 0;
        line-height: 1.55;
    }
    .sm-card-subtitle {
        font-size: 0.72rem;
        font-weight: 600;
        letter-spacing: 0.09em;
        text-transform: uppercase;
        color: #94a3b8;
        margin: 1rem 0 0.4rem 0;
    }
    .sm-list {
        margin: 0 0 0.5rem 0;
        padding-left: 1.1rem;
        color: #334155;
        font-size: 0.85rem;
        line-height: 1.75;
    }
    .sm-list li code {
        background: #f1f5f9;
        border-radius: 4px;
        padding: 1px 6px;
        font-family: 'DM Mono', monospace;
        font-size: 0.8rem;
        color: #0c1f3a;
    }

    /* ─── File uploader ──────────────────────────────────── */
    [data-testid="stFileUploaderDropzone"] {
        background: #f8fafc !important;
        border: 2px dashed #cbd5e1 !important;
        border-radius: 8px !important;
    }
    [data-testid="stFileUploaderDropzone"]:hover {
        border-color: #0d9488 !important;
    }

    /* ─── Buttons ────────────────────────────────────────── */
    .stButton > button,
    .stDownloadButton > button {
        font-family: 'DM Sans', sans-serif !important;
        font-weight: 600 !important;
        font-size: 0.85rem !important;
        border-radius: 8px !important;
        transition: all 0.2s ease !important;
    }
    .stDownloadButton > button[kind="primary"],
    .stButton > button[kind="primary"] {
        background: #0c1f3a !important;
        color: white !important;
        border: none !important;
        box-shadow: 0 2px 8px rgba(12,31,58,0.25) !important;
    }
    .stDownloadButton > button[kind="primary"]:hover,
    .stButton > button[kind="primary"]:hover {
        background: #0d9488 !important;
        box-shadow: 0 4px 16px rgba(13,148,136,0.3) !important;
        transform: translateY(-1px) !important;
    }
    .stDownloadButton > button[kind="secondary"],
    .stButton > button[kind="secondary"] {
        background: #ffffff !important;
        color: #0c1f3a !important;
        border: 1px solid #e2e8f0 !important;
    }
    .stDownloadButton > button[kind="secondary"]:hover,
    .stButton > button[kind="secondary"]:hover {
        border-color: #0d9488 !important;
        color: #0d9488 !important;
    }

    /* ─── Alerts ─────────────────────────────────────────── */
    [data-testid="stSuccess"] {
        background: #f0fdf9 !important;
        border-left: 3px solid #0d9488 !important;
        border-radius: 8px !important;
    }
    [data-testid="stError"] {
        background: #fef2f2 !important;
        border-left: 3px solid #ef4444 !important;
        border-radius: 8px !important;
    }

    /* ─── Caption ────────────────────────────────────────── */
    .stCaption, [data-testid="stCaptionContainer"] {
        font-family: 'DM Sans', sans-serif !important;
        color: #94a3b8 !important;
        font-size: 0.78rem !important;
    }

    h1:first-of-type { display: none; }
    h2, h3 { color: #0c1f3a !important; font-weight: 600 !important; }
</style>
"""


def render_navbar() -> None:
    logo_src = ""
    if LOGO_PATH.exists():
        raw = LOGO_PATH.read_bytes()
        b64 = base64.b64encode(raw).decode()
        logo_src = f"data:image/png;base64,{b64}"

    logo_html = (
        f'<img src="{logo_src}" alt="SensiMedical" />'
        if logo_src
        else '<span style="color:white;font-weight:700;font-size:1rem;">SensiMedical</span>'
    )

    st.markdown(
        f"""
        <div class="sm-navbar">
            <div class="sm-navbar-brand">
                {logo_html}
            </div>
            <div class="sm-navbar-center">Sales &amp; Inventory Lot Tool</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def main() -> None:
    st.set_page_config(
        page_title="SensiMedical – Sales & Inventory Lots",
        layout="wide",
        page_icon="📊",
    )
    st.markdown(SENSIMEDICAL_CSS, unsafe_allow_html=True)
    render_navbar()

    # ─── Page header ─────────────────────────────────────────
    st.markdown(
        """
        <div class="sm-page-header">
            <h1>Sales &amp; Inventory Lot Intelligence</h1>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ─── Two-column layout ────────────────────────────────────
    left, right = st.columns([1.4, 1])

    with left:
        # Card header rendered as HTML above the Streamlit widgets
        st.markdown(
            """
            <div class="sm-col-card">
                <p class="sm-card-title">Upload monthly sales workbook</p>
                <p class="sm-card-text">
                    Use the standard ERP export (.xlsx). We will explode lots, allocate sales
                    and costs, and write back clean tabs ready for analysis.
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        uploaded = st.file_uploader(
            "Upload Excel file (.xlsx)",
            type=["xlsx"],
            help="Drag in the monthly Sales & Inventory export from your ERP.",
            label_visibility="collapsed",
        )
        st.caption("Accepted format: Excel .xlsx · one file at a time")

        sample_path = Path(__file__).parent / "assets" / "sample.xlsx"
        if sample_path.exists():
            try:
                with open(sample_path, "rb") as f:
                    sample_bytes = f.read()
                st.download_button(
                    label="⬇  Download sample workbook",
                    data=sample_bytes,
                    file_name="sample_sales_inventory_lots.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="secondary",
                )
            except OSError:
                st.caption("Sample workbook is present but could not be read.")
        else:
            st.caption("Sample workbook not available on this deployment.")

    with right:
        st.markdown(
            """
            <div class="sm-col-card">
                <p class="sm-card-title">What this tool does</p>
                <p class="sm-card-text">Turn the standard monthly export into clean, lot-level SensiMedical tabs in one upload.</p>
                <ul class="sm-list">
                    <li>Explodes lot numbers into one row per lot.</li>
                    <li>Allocates Ext Sales, Ext Cost, Weight, and fees by lot.</li>
                    <li>Separates Real Sales, Samples, and Credit Memos.</li>
                    <li>Builds a clean Summary tab for (MPN, Lot) totals.</li>
                </ul>
                <p class="sm-card-subtitle">Required columns (first sheet)</p>
                <p class="sm-card-text" style="margin-bottom:0.4rem;">
                    Column headers must exist exactly with these names (order does not matter):
                </p>
                <ul class="sm-list">
                    <li><code>Lot Numbers(QTY)</code></li>
                    <li><code>Qty Shipped</code></li>
                    <li><code>Ext Sales</code></li>
                    <li><code>Sales Price</code></li>
                </ul>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # ─── Processing ───────────────────────────────────────────
    if uploaded is not None:
        with st.spinner("Processing workbook… this can take a moment for large files."):
            with NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp.write(uploaded.getbuffer())
                tmp_path = Path(tmp.name)

            try:
                output_path = process_workbook(tmp_path)
                with open(output_path, "rb") as f:
                    data = f.read()

                st.success("✓ Processing complete. Download your updated workbook below.")
                st.download_button(
                    label="⬇  Download processed workbook",
                    data=data,
                    file_name=f"processed_{uploaded.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
            except Exception as exc:
                st.error(f"Error processing file: {exc}")


if __name__ == "__main__":
    main()
