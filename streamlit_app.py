"""Streamlit app for AMP Laydowns deck generation."""

import tempfile
from datetime import datetime
from pathlib import Path

import streamlit as st

st.set_page_config(
    page_title="AMP Laydowns Generator",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="collapsed",
)

from amp_automation.config import load_master_config
from amp_automation.presentation.assembly import build_presentation


def get_project_root() -> Path:
    return Path(__file__).resolve().parent


def main():
    # Load config
    project_root = get_project_root()
    config = load_master_config(project_root / "config" / "master_config.json")
    template_location = config.section("template")["location"]
    template_path = project_root / template_location
    table_config = config.section("presentation")["table"]
    smart_pagination = table_config.get("smart_pagination_enabled", False)
    max_rows = table_config.get("max_rows_per_slide", 40)

    # Header
    st.title("AMP Laydowns Generator")
    st.caption("Transform Lumina Excel exports into PowerPoint presentations")

    # Status bar
    col1, col2, col3 = st.columns(3)
    with col1:
        if template_path.exists():
            st.success("Template Ready", icon="✅")
        else:
            st.error("Template Missing", icon="❌")
            return
    with col2:
        st.info(f"Max {max_rows} rows/slide", icon="📄")
    with col3:
        if smart_pagination:
            st.success("Smart Pagination ON", icon="🧠")
        else:
            st.warning("Smart Pagination OFF", icon="⚠️")

    st.divider()

    # File upload
    uploaded_file = st.file_uploader(
        "Upload BulkPlanData Excel file",
        type=["xlsx"],
        help="Lumina export file (BulkPlanData_YYYY_MM_DD.xlsx)"
    )

    # Output filename
    output_name = st.text_input(
        "Output filename (optional)",
        placeholder="Leave blank for auto-generated name"
    )

    if uploaded_file:
        st.divider()

        # File info
        file_size_mb = uploaded_file.size / 1024 / 1024
        st.write(f"**Selected:** {uploaded_file.name} ({file_size_mb:.2f} MB)")

        # Generate button
        if st.button("Generate Presentation", type="primary", use_container_width=True):

            with st.status("Generating presentation...", expanded=True) as status:
                try:
                    with tempfile.TemporaryDirectory() as temp_dir:
                        temp_path = Path(temp_dir)

                        # Save uploaded file
                        st.write("📁 Saving uploaded file...")
                        input_path = temp_path / uploaded_file.name
                        with open(input_path, "wb") as f:
                            f.write(uploaded_file.getbuffer())

                        # Generate output path
                        if output_name:
                            out_filename = output_name if output_name.endswith(".pptx") else f"{output_name}.pptx"
                        else:
                            date_str = datetime.now().strftime("%d%m%y")
                            out_filename = f"AMP_Laydowns_{date_str}.pptx"

                        output_path = temp_path / out_filename

                        # Build presentation
                        st.write("🔄 Processing data and building slides...")
                        build_presentation(
                            template_path=str(template_path),
                            excel_path=str(input_path),
                            output_path=str(output_path),
                        )

                        st.write("✅ Generation complete!")

                        # Read generated file
                        with open(output_path, "rb") as f:
                            pptx_bytes = f.read()

                        status.update(label="Complete!", state="complete", expanded=False)

                    # Success and download
                    file_size_kb = len(pptx_bytes) / 1024
                    st.success(f"Generated: {out_filename} ({file_size_kb:.1f} KB)")

                    st.download_button(
                        label="Download Presentation",
                        data=pptx_bytes,
                        file_name=out_filename,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        type="primary",
                        use_container_width=True,
                    )

                except Exception as e:
                    status.update(label="Failed", state="error")
                    st.error(f"Error: {str(e)}")
                    with st.expander("Details"):
                        st.exception(e)

    else:
        # Instructions when no file
        st.divider()

        with st.expander("ℹ️ How to use", expanded=True):
            st.markdown("""
            1. Export **BulkPlanData** from Lumina
            2. Delete **"Media Plans"** and **"Vendor Detail"** tabs (keep only **"Flight"**)
            3. Upload the Excel file above
            4. Click **Generate Presentation**
            5. Download the PowerPoint deck
            """)

        with st.expander("📋 Data transformations applied"):
            st.markdown("""
            - **Expert campaigns** excluded automatically
            - **Geography** normalized (FWA→FSA, GINE→GNE, etc.)
            - **Panadol** split into Pain and C&F brands
            - **GNE Pan Asian TV** rows filtered
            """)

        with st.expander("📄 Required columns"):
            st.code("""
Country / Geography
Global Masterbrand
Plan Name
Media Type
Net Cost
Flight Start Date
            """)


if __name__ == "__main__":
    main()
