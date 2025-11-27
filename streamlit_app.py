"""Streamlit app for AMP Laydowns deck generation."""

import logging
import re
import tempfile
import threading
import time
from datetime import datetime
from pathlib import Path
from queue import Queue

import streamlit as st

st.set_page_config(
    page_title="AMP Laydowns Generator",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="collapsed",
)

from amp_automation.config import load_master_config
from amp_automation.presentation.assembly import build_presentation


class ProgressHandler(logging.Handler):
    """Custom logging handler to capture progress from build_presentation."""

    def __init__(self, queue: Queue):
        super().__init__()
        self.queue = queue
        # Pattern to match "Processing combination X/Y: ..."
        self.combination_pattern = re.compile(r"Processing combination (\d+)/(\d+)")
        # Pattern to match "Found X unique ... combinations"
        self.total_pattern = re.compile(r"Found (\d+) unique .* combinations")

    def emit(self, record):
        msg = record.getMessage()

        # Check for total combinations
        total_match = self.total_pattern.search(msg)
        if total_match:
            self.queue.put(("total", int(total_match.group(1))))
            return

        # Check for combination progress
        combo_match = self.combination_pattern.search(msg)
        if combo_match:
            current = int(combo_match.group(1))
            total = int(combo_match.group(2))
            self.queue.put(("progress", current, total, msg))


def format_time(seconds: float) -> str:
    """Format seconds into human-readable time."""
    if seconds < 60:
        return f"{int(seconds)}s"
    minutes = int(seconds // 60)
    secs = int(seconds % 60)
    return f"{minutes}m {secs}s"


def get_project_root() -> Path:
    return Path(__file__).resolve().parent


def run_generation(template_path: str, input_path: str, output_path: str, queue: Queue):
    """Run presentation generation in a thread, sending progress to queue."""
    try:
        # Add our progress handler to the logger
        logger = logging.getLogger()
        handler = ProgressHandler(queue)
        handler.setLevel(logging.INFO)
        logger.addHandler(handler)

        queue.put(("status", "Starting generation..."))

        build_presentation(
            template_path=template_path,
            excel_path=input_path,
            output_path=output_path,
        )

        queue.put(("done", None))

    except Exception as e:
        queue.put(("error", str(e)))

    finally:
        logger.removeHandler(handler)


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

            try:
                with tempfile.TemporaryDirectory() as temp_dir:
                    temp_path = Path(temp_dir)

                    # Save uploaded file
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

                    # Progress UI
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    time_text = st.empty()

                    status_text.markdown("**Status:** Initializing...")

                    # Create queue for thread communication
                    queue: Queue = Queue()

                    # Start generation in background thread
                    thread = threading.Thread(
                        target=run_generation,
                        args=(str(template_path), str(input_path), str(output_path), queue)
                    )
                    thread.start()

                    # Track progress
                    start_time = time.time()
                    total_combinations = 0
                    current_combination = 0
                    last_message = ""
                    done = False
                    error = None

                    while thread.is_alive() or not queue.empty():
                        try:
                            msg = queue.get(timeout=0.1)

                            if msg[0] == "total":
                                total_combinations = msg[1]

                            elif msg[0] == "progress":
                                current_combination = msg[1]
                                total_combinations = msg[2]
                                last_message = msg[3]

                                # Update progress bar (reserve 10% for finalization)
                                progress = min(0.9, current_combination / total_combinations) if total_combinations > 0 else 0
                                progress_bar.progress(progress)

                                # Calculate ETA
                                elapsed = time.time() - start_time
                                if current_combination > 0:
                                    rate = elapsed / current_combination
                                    remaining = (total_combinations - current_combination) * rate
                                    eta_str = format_time(remaining)
                                    time_text.markdown(f"⏱️ **Elapsed:** {format_time(elapsed)} | **ETA:** {eta_str}")

                                # Extract brand info from message
                                brand_match = re.search(r": (.+) - (\d+)$", last_message)
                                if brand_match:
                                    status_text.markdown(f"**Processing:** {current_combination}/{total_combinations} — {brand_match.group(1)}")
                                else:
                                    status_text.markdown(f"**Processing:** {current_combination}/{total_combinations}")

                            elif msg[0] == "status":
                                status_text.markdown(f"**Status:** {msg[1]}")

                            elif msg[0] == "done":
                                done = True

                            elif msg[0] == "error":
                                error = msg[1]

                        except:
                            pass  # Queue timeout, continue loop

                    thread.join()

                    if error:
                        raise Exception(error)

                    # Finalization
                    progress_bar.progress(1.0)
                    elapsed = time.time() - start_time
                    status_text.markdown("**Status:** ✅ Complete!")
                    time_text.markdown(f"⏱️ **Total time:** {format_time(elapsed)}")

                    # Read generated file
                    with open(output_path, "rb") as f:
                        pptx_bytes = f.read()

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
