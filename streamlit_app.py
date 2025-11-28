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

# Custom CSS - dark theme styling with animations
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;600&family=Inter:wght@400;500;600&display=swap');

    /* Dark theme colors */
    :root {
        --emerald: #10b981;
        --sky: #38bdf8;
        --rose: #f472b6;
        --amber: #fbbf24;
        --surface: rgba(20, 20, 22, 0.8);
        --surface-hover: #252525;
        --border: rgba(255, 255, 255, 0.06);
    }

    /* Animated gradient background */
    .stApp {
        background: linear-gradient(
            135deg,
            #0a0a0c 0%,
            #0d1117 25%,
            #0f0f14 50%,
            #0a0f18 75%,
            #0a0a0c 100%
        );
        background-size: 400% 400%;
        animation: gradientShift 20s ease infinite;
    }

    .stApp::before {
        content: '';
        position: fixed;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background:
            radial-gradient(ellipse 80% 60% at 15% 10%, rgba(16, 185, 129, 0.12) 0%, transparent 60%),
            radial-gradient(ellipse 70% 50% at 85% 85%, rgba(56, 189, 248, 0.1) 0%, transparent 55%),
            radial-gradient(ellipse 60% 40% at 50% 50%, rgba(244, 114, 182, 0.08) 0%, transparent 50%);
        pointer-events: none;
        z-index: 0;
        animation: auroraMove 8s ease-in-out infinite alternate;
    }

    .stApp::after {
        content: '';
        position: fixed;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background:
            radial-gradient(ellipse 50% 40% at 70% 20%, rgba(168, 85, 247, 0.08) 0%, transparent 50%),
            radial-gradient(ellipse 40% 30% at 30% 70%, rgba(251, 191, 36, 0.06) 0%, transparent 45%);
        pointer-events: none;
        z-index: 0;
        animation: auroraMove 10s ease-in-out infinite alternate-reverse;
    }

    @keyframes gradientShift {
        0%, 100% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
    }

    @keyframes auroraMove {
        0% {
            opacity: 0.7;
            transform: translateX(-3%) translateY(-2%) scale(1);
        }
        50% {
            opacity: 1;
            transform: translateX(3%) translateY(2%) scale(1.05);
        }
        100% {
            opacity: 0.8;
            transform: translateX(-2%) translateY(3%) scale(0.98);
        }
    }

    /* Ensure content is above background */
    .stApp > * {
        position: relative;
        z-index: 1;
    }

    /* Animations */
    @keyframes fadeSlideIn {
        from {
            opacity: 0;
            transform: translateY(-10px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }

    @keyframes slideInLeft {
        from {
            opacity: 0;
            transform: translateX(-15px);
        }
        to {
            opacity: 1;
            transform: translateX(0);
        }
    }

    @keyframes chipPop {
        from {
            opacity: 0;
            transform: scale(0.9);
        }
        to {
            opacity: 1;
            transform: scale(1);
        }
    }

    @keyframes pulse {
        0%, 100% { opacity: 1; }
        50% { opacity: 0.6; }
    }

    @keyframes shimmer {
        0% { background-position: -200% 0; }
        100% { background-position: 200% 0; }
    }

    /* Instructions box */
    .guide-box {
        background: var(--surface);
        border: 1px solid var(--border);
        border-radius: 12px;
        padding: 1.5rem 1.75rem;
        margin: 0.5rem 0 1.5rem 0;
        animation: fadeSlideIn 0.5s ease-out;
    }

    .guide-title {
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.7rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.2em;
        color: var(--emerald);
        margin-bottom: 1.25rem;
        display: flex;
        align-items: center;
        gap: 0.625rem;
    }

    .guide-title::before {
        content: '';
        width: 8px;
        height: 8px;
        background: var(--emerald);
        border-radius: 50%;
        animation: pulse 2s ease-in-out infinite;
        box-shadow: 0 0 8px rgba(16, 185, 129, 0.5);
    }

    .guide-step {
        display: flex;
        align-items: baseline;
        gap: 1rem;
        padding: 0.5rem 0;
        font-size: 0.9rem;
        font-family: 'Inter', sans-serif;
        color: #d4d4d4;
        letter-spacing: 0.01em;
        line-height: 1.6;
        animation: slideInLeft 0.4s ease-out backwards;
    }

    .guide-step:nth-child(2) { animation-delay: 0.1s; }
    .guide-step:nth-child(3) { animation-delay: 0.2s; }
    .guide-step:nth-child(4) { animation-delay: 0.3s; }
    .guide-step:nth-child(5) { animation-delay: 0.4s; }

    .guide-step .num {
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.8rem;
        font-weight: 500;
        color: #525252;
        min-width: 1.75rem;
        letter-spacing: 0;
    }

    .guide-step .hl {
        color: var(--amber);
        font-weight: 600;
        letter-spacing: 0.02em;
        padding: 0.125rem 0.375rem;
        background: rgba(251, 191, 36, 0.1);
        border-radius: 4px;
        margin: 0 0.125rem;
    }

    /* Status chips */
    .chip-row {
        display: flex;
        gap: 0.75rem;
        flex-wrap: wrap;
        margin: 1.25rem 0 1.5rem 0;
    }

    .chip {
        display: inline-flex;
        align-items: center;
        gap: 0.5rem;
        padding: 0.75rem 1.125rem;
        border-radius: 10px;
        font-size: 0.825rem;
        font-weight: 500;
        font-family: 'Inter', sans-serif;
        letter-spacing: 0.015em;
        animation: chipPop 0.4s ease-out backwards;
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }

    .chip:nth-child(1) { animation-delay: 0.5s; }
    .chip:nth-child(2) { animation-delay: 0.6s; }
    .chip:nth-child(3) { animation-delay: 0.7s; }

    .chip:hover {
        transform: translateY(-2px);
    }

    .chip.green {
        background: rgba(16, 185, 129, 0.15);
        color: var(--emerald);
        border: 1px solid rgba(16, 185, 129, 0.3);
    }

    .chip.green:hover {
        box-shadow: 0 4px 12px rgba(16, 185, 129, 0.2);
    }

    .chip.blue {
        background: rgba(56, 189, 248, 0.15);
        color: var(--sky);
        border: 1px solid rgba(56, 189, 248, 0.3);
    }

    .chip.blue:hover {
        box-shadow: 0 4px 12px rgba(56, 189, 248, 0.2);
    }

    .chip.pink {
        background: rgba(244, 114, 182, 0.15);
        color: var(--rose);
        border: 1px solid rgba(244, 114, 182, 0.3);
    }

    .chip.pink:hover {
        box-shadow: 0 4px 12px rgba(244, 114, 182, 0.2);
    }

    /* File badge */
    .file-badge {
        display: inline-flex;
        align-items: center;
        gap: 0.5rem;
        background: rgba(16, 185, 129, 0.1);
        border: 1px solid rgba(16, 185, 129, 0.25);
        border-radius: 8px;
        padding: 0.75rem 1.125rem;
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.825rem;
        color: var(--emerald);
        margin: 0.5rem 0;
        letter-spacing: 0.01em;
        animation: fadeSlideIn 0.4s ease-out;
    }

    /* Divider */
    .divider-line {
        height: 1px;
        background: linear-gradient(90deg, transparent, var(--border), transparent);
        margin: 1.5rem 0;
    }

    /* Footer */
    .app-footer {
        font-size: 0.75rem;
        color: #525252;
        text-align: center;
        margin-top: 2.5rem;
        padding-top: 1.25rem;
        border-top: 1px solid var(--border);
        letter-spacing: 0.02em;
    }

    /* Animated title */
    .fancy-title {
        font-size: 2.5rem;
        font-weight: 700;
        font-family: 'Inter', sans-serif;
        letter-spacing: -0.02em;
        background: linear-gradient(
            90deg,
            #10b981 0%,
            #38bdf8 25%,
            #f472b6 50%,
            #fbbf24 75%,
            #10b981 100%
        );
        background-size: 200% auto;
        -webkit-background-clip: text;
        background-clip: text;
        -webkit-text-fill-color: transparent;
        animation: titleGradient 4s linear infinite, titleGlow 2s ease-in-out infinite alternate;
        filter: drop-shadow(0 0 20px rgba(16, 185, 129, 0.3));
        margin-bottom: 0.5rem;
    }

    @keyframes titleGradient {
        0% { background-position: 0% center; }
        100% { background-position: 200% center; }
    }

    @keyframes titleGlow {
        0% {
            filter: drop-shadow(0 0 15px rgba(16, 185, 129, 0.4));
        }
        100% {
            filter: drop-shadow(0 0 25px rgba(56, 189, 248, 0.5));
        }
    }
</style>
""", unsafe_allow_html=True)

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

    # Header with animated title
    st.markdown('<h1 class="fancy-title">AMP Laydowns Generator</h1>', unsafe_allow_html=True)
    st.caption("Transform Lumina Excel exports into PowerPoint presentations")

    # Instructions box at TOP
    st.markdown("""
    <div class="guide-box">
        <div class="guide-title">How to Use</div>
        <div class="guide-step"><span class="num">1.</span> Export <span class="hl">BulkPlanData</span> from Lumina</div>
        <div class="guide-step"><span class="num">2.</span> Delete <span class="hl">Media Plans</span> and <span class="hl">Vendor Detail</span> tabs (keep only <span class="hl">Flight</span>)</div>
        <div class="guide-step"><span class="num">3.</span> Upload the Excel file below</div>
        <div class="guide-step"><span class="num">4.</span> Click <span class="hl">Generate Presentation</span> and download your deck</div>
    </div>
    """, unsafe_allow_html=True)

    # Status chips
    template_ok = template_path.exists()
    st.markdown(f"""
    <div class="chip-row">
        <span class="chip {'green' if template_ok else 'pink'}">{'✓' if template_ok else '✗'} {'Template Ready' if template_ok else 'Template Missing'}</span>
        <span class="chip blue">📄 Max {max_rows} rows/slide</span>
        <span class="chip {'pink' if smart_pagination else 'blue'}">{'🧠' if smart_pagination else '○'} Smart Pagination {'ON' if smart_pagination else 'OFF'}</span>
    </div>
    """, unsafe_allow_html=True)

    if not template_path.exists():
        st.error("Template file not found. Please ensure the template is in place.")
        return

    st.markdown('<div class="divider-line"></div>', unsafe_allow_html=True)

    # File upload
    uploaded_file = st.file_uploader(
        "Upload BulkPlanData Excel file",
        type=["xlsx"],
        help="Lumina export file (BulkPlanData_YYYY_MM_DD.xlsx)"
    )

    if uploaded_file:
        st.markdown('<div class="divider-line"></div>', unsafe_allow_html=True)

        # File info badge
        file_size_mb = uploaded_file.size / 1024 / 1024
        st.markdown(f'<div class="file-badge">📁 {uploaded_file.name} — {file_size_mb:.2f} MB</div>', unsafe_allow_html=True)

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
        # Additional info when no file uploaded (instructions already at top)
        st.markdown('<div class="divider-line"></div>', unsafe_allow_html=True)

        with st.expander("📋 Data transformations applied automatically"):
            st.markdown("""
            - **Expert campaigns** excluded
            - **Geography** normalized (FWA→FSA, GINE→GNE, etc.)
            - **Panadol** split into Pain and C&F brands
            - **GNE Pan Asian TV** rows filtered
            """)

        with st.expander("📄 Required Excel columns"):
            st.code("""Country / Geography
Global Masterbrand
Plan Name
Media Type
Net Cost
Flight Start Date""", language=None)

    # Footer (always visible - outside if/else block)
    st.markdown("""
    <div class="app-footer">
        AMP Laydowns Automation • v1.0 • Python + Streamlit
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
