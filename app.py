from io import BytesIO
from pathlib import Path
from tempfile import NamedTemporaryFile
import base64

import streamlit as st

from sales_tool import process_workbook


st.set_page_config(
    page_title="SensiMedical – Sales & Inventory Lots",
    layout="wide",
)


def render_header() -> None:
    """Top branding bar with SensiMedical styling."""
    logo_path = Path(__file__).parent / "assets" / "sensimedical-logo.png"

    # Inject simple SensiMedical‑style theming via CSS
    st.markdown(
        """
        <style>
        .sensi-header {
            background: linear-gradient(90deg, #19d0c8 0%, #1f66d3 100%);
            padding: 1.2rem 2rem;
            border-radius: 0 0 18px 18px;
            display: flex;
            align-items: center;
            gap: 1.25rem;
            box-shadow: 0 8px 18px rgba(0, 0, 0, 0.12);
        }
        .sensi-title {
            color: #ffffff;
            font-weight: 700;
            font-size: 1.6rem;
            margin: 0;
        }
        .sensi-subtitle {
            color: #e9f3ff;
            margin: 0.15rem 0 0;
            font-size: 0.95rem;
        }
        .sensi-main {
            padding: 1.6rem 1.8rem 2rem 1.8rem;
        }
        .sensi-card {
            background: #ffffff;
            border-radius: 18px;
            padding: 1.2rem 1.5rem 1.4rem;
            box-shadow: 0 10px 24px rgba(7, 26, 71, 0.08);
            border: 1px solid rgba(16, 50, 91, 0.04);
        }
        .sensi-card-title {
            font-size: 1.05rem;
            font-weight: 600;
            color: #10325b;
            margin-bottom: 0.4rem;
        }
        .sensi-card-text {
            font-size: 0.9rem;
            color: #4c607a;
            margin-bottom: 0.8rem;
        }
        .sensi-steps li {
            margin-bottom: 0.3rem;
        }
        .sensi-logo {
            height: 32px;
            display: block;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    # Load logo as base64 so it can sit inside the gradient band
    logo_html: str
    if logo_path.exists():
        with open(logo_path, "rb") as f:
            encoded = base64.b64encode(f.read()).decode("utf-8")
        logo_html = (
            f'<img src="data:image/png;base64,{encoded}" '
            f'alt="SensiMedical logo" class="sensi-logo"/>'
        )
    else:
        logo_html = '<span class="sensi-title">SensiMedical</span>'

    st.markdown(
        f"""
        <div class="sensi-header">
            <div style="display:flex; align-items:center; gap:1rem;">
                {logo_html}
                <div>
                    <p class="sensi-title">Sales &amp; Inventory Lot Intelligence</p>
                    <p class="sensi-subtitle">
                        Turn the standard monthly export into clean, lot-level SensiMedical tabs in one upload.
                    </p>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def main() -> None:
    render_header()

    st.markdown('<div class="sensi-main">', unsafe_allow_html=True)

    left, right = st.columns([1.4, 1])

    with left:
        st.markdown('<div class="sensi-card">', unsafe_allow_html=True)
        st.markdown(
            """
            <p class="sensi-card-title">Upload monthly sales workbook</p>
            <p class="sensi-card-text">
                Use the standard ERP export (.xlsx). We will explode lots, allocate sales and costs,
                and write back clean tabs ready for analysis.
            </p>
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
                    label="Download sample workbook",
                    data=sample_bytes,
                    file_name="sample_sales_inventory_lots.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                    type="secondary",
                )
            except OSError:
                st.caption("Sample workbook is present but could not be read.")
        else:
            st.caption("Sample workbook not available on this machine.")
        st.markdown("</div>", unsafe_allow_html=True)

    with right:
        st.markdown('<div class="sensi-card">', unsafe_allow_html=True)
        st.markdown(
            """
            <p class="sensi-card-title">What this tool does</p>
            <ul class="sensi-steps">
                <li>Explodes lot numbers into one row per lot.</li>
                <li>Allocates Ext Sales, Ext Cost, Weight, and fees by lot.</li>
                <li>Separates Real Sales, Samples, and Credit Memos.</li>
                <li>Builds a clean Summary tab for (MPN, Lot) totals.</li>
            </ul>
            <p class="sensi-card-title" style="margin-top:0.9rem;">Required columns (first sheet)</p>
            <p class="sensi-card-text" style="margin-bottom:0.4rem;">
                Column headers must exist exactly with these names (order does not matter):
            </p>
            <ul class="sensi-steps">
                <li><code>Lot Numbers(QTY)</code></li>
                <li><code>Qty Shipped</code></li>
                <li><code>Ext Sales</code></li>
                <li><code>Sales Price</code></li>
            </ul>
            """,
            unsafe_allow_html=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)

    if uploaded is not None:
        with st.spinner("Processing workbook… this can take a moment for large files."):
            # Save uploaded file to a temporary location
            with NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp.write(uploaded.getbuffer())
                tmp_path = Path(tmp.name)

            try:
                output_path = process_workbook(tmp_path)

                # Read processed workbook back into memory for download
                with open(output_path, "rb") as f:
                    data = f.read()

                st.success("Processing complete. Download your updated workbook below.")

                st.download_button(
                    label="Download processed workbook",
                    data=data,
                    file_name=f"processed_{uploaded.name}",
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                    type="primary",
                )
            except Exception as exc:
                st.error(f"Error processing file: {exc}")

    st.markdown("</div>", unsafe_allow_html=True)


if __name__ == "__main__":
    main()

