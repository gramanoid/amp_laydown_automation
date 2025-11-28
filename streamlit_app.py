"""Streamlit app for AMP Laydowns deck generation."""

import html
import logging
import re
import tempfile
import threading
import time
from datetime import datetime
from pathlib import Path
from queue import Empty, Queue

import streamlit as st
import streamlit.components.v1 as components

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

    /* ═══════════════════════════════════════════════════════════════
       PROGRESS UI - Enhanced Generation Experience
       ═══════════════════════════════════════════════════════════════ */

    /* Main progress container */
    .progress-container {
        background: var(--surface);
        border: 1px solid var(--border);
        border-radius: 16px;
        padding: 1.5rem;
        margin: 1rem 0;
        animation: fadeSlideIn 0.4s ease-out;
    }

    /* Stage indicators */
    .stage-indicators {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 1.5rem;
        position: relative;
    }

    .stage-indicators::before {
        content: '';
        position: absolute;
        top: 50%;
        left: 15%;
        right: 15%;
        height: 2px;
        background: var(--border);
        transform: translateY(-50%);
        z-index: 0;
    }

    .stage {
        display: flex;
        flex-direction: column;
        align-items: center;
        gap: 0.5rem;
        z-index: 1;
    }

    .stage-dot {
        width: 32px;
        height: 32px;
        border-radius: 50%;
        background: var(--surface);
        border: 2px solid var(--border);
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 0.75rem;
        color: #525252;
        transition: all 0.4s ease;
    }

    .stage-dot.active {
        border-color: var(--emerald);
        background: rgba(16, 185, 129, 0.15);
        color: var(--emerald);
        animation: stagePulse 1.5s ease-in-out infinite;
        box-shadow: 0 0 15px rgba(16, 185, 129, 0.4);
    }

    .stage-dot.complete {
        border-color: var(--emerald);
        background: var(--emerald);
        color: #0a0a0c;
    }

    .stage-label {
        font-size: 0.7rem;
        font-family: 'JetBrains Mono', monospace;
        text-transform: uppercase;
        letter-spacing: 0.1em;
        color: #525252;
        transition: color 0.3s ease;
    }

    .stage.active .stage-label,
    .stage.complete .stage-label {
        color: var(--emerald);
    }

    @keyframes stagePulse {
        0%, 100% {
            transform: scale(1);
            box-shadow: 0 0 15px rgba(16, 185, 129, 0.4);
        }
        50% {
            transform: scale(1.1);
            box-shadow: 0 0 25px rgba(16, 185, 129, 0.6);
        }
    }

    /* Custom progress bar */
    .progress-track {
        height: 8px;
        background: rgba(255, 255, 255, 0.05);
        border-radius: 4px;
        overflow: hidden;
        position: relative;
        margin-bottom: 1rem;
    }

    .progress-fill {
        height: 100%;
        background: linear-gradient(
            90deg,
            var(--emerald) 0%,
            var(--sky) 50%,
            var(--emerald) 100%
        );
        background-size: 200% 100%;
        border-radius: 4px;
        transition: width 0.3s ease-out;
        animation: progressShimmer 2s linear infinite;
        position: relative;
    }

    .progress-fill::after {
        content: '';
        position: absolute;
        top: 0;
        right: 0;
        bottom: 0;
        width: 100px;
        background: linear-gradient(90deg, transparent, rgba(255, 255, 255, 0.3), transparent);
        animation: progressGlow 1.5s ease-in-out infinite;
    }

    @keyframes progressShimmer {
        0% { background-position: 200% 0; }
        100% { background-position: -200% 0; }
    }

    @keyframes progressGlow {
        0%, 100% { opacity: 0; transform: translateX(-100%); }
        50% { opacity: 1; transform: translateX(100%); }
    }

    /* Stats grid */
    .stats-grid {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 1rem;
        margin-top: 1rem;
    }

    .stat-card {
        background: rgba(255, 255, 255, 0.02);
        border: 1px solid var(--border);
        border-radius: 10px;
        padding: 0.875rem;
        text-align: center;
        transition: all 0.3s ease;
    }

    .stat-card:hover {
        background: rgba(255, 255, 255, 0.04);
        border-color: rgba(255, 255, 255, 0.1);
    }

    .stat-value {
        font-family: 'JetBrains Mono', monospace;
        font-size: 1.25rem;
        font-weight: 600;
        color: #e5e5e5;
        margin-bottom: 0.25rem;
    }

    .stat-value.highlight {
        color: var(--emerald);
    }

    .stat-label {
        font-size: 0.65rem;
        font-family: 'JetBrains Mono', monospace;
        text-transform: uppercase;
        letter-spacing: 0.15em;
        color: #525252;
    }

    /* Current item display */
    .current-item {
        display: flex;
        align-items: center;
        gap: 0.75rem;
        padding: 0.875rem 1rem;
        background: rgba(16, 185, 129, 0.08);
        border: 1px solid rgba(16, 185, 129, 0.2);
        border-radius: 10px;
        margin-top: 1rem;
    }

    .pulse-dot {
        width: 10px;
        height: 10px;
        background: var(--emerald);
        border-radius: 50%;
        animation: pulseDot 1s ease-in-out infinite;
        flex-shrink: 0;
    }

    @keyframes pulseDot {
        0%, 100% {
            transform: scale(1);
            opacity: 1;
        }
        50% {
            transform: scale(1.3);
            opacity: 0.7;
        }
    }

    .current-item-text {
        font-family: 'Inter', sans-serif;
        font-size: 0.85rem;
        color: var(--emerald);
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }

    /* Completion state */
    .progress-complete {
        text-align: center;
        padding: 1rem 0;
    }

    .completion-icon {
        width: 64px;
        height: 64px;
        background: linear-gradient(135deg, var(--emerald), var(--sky));
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        margin: 0 auto 1rem auto;
        animation: completionPop 0.5s cubic-bezier(0.68, -0.55, 0.265, 1.55);
        box-shadow: 0 0 30px rgba(16, 185, 129, 0.5);
    }

    .completion-icon svg {
        width: 32px;
        height: 32px;
        color: #0a0a0c;
    }

    @keyframes completionPop {
        0% {
            transform: scale(0);
            opacity: 0;
        }
        50% {
            transform: scale(1.2);
        }
        100% {
            transform: scale(1);
            opacity: 1;
        }
    }

    .completion-text {
        font-family: 'Inter', sans-serif;
        font-size: 1.1rem;
        font-weight: 600;
        color: var(--emerald);
        margin-bottom: 0.25rem;
    }

    .completion-subtext {
        font-size: 0.8rem;
        color: #737373;
    }

    /* Download button enhancement */
    .download-section {
        margin-top: 1.25rem;
        padding-top: 1.25rem;
        border-top: 1px solid var(--border);
    }

    /* Hide default streamlit progress bar */
    .stProgress {
        display: none !important;
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

                    # Enhanced Progress UI Container
                    progress_container = st.empty()

                    # Auto-scroll to progress section using components.html (actually executes JS)
                    components.html('''<script>
                    window.parent.document.querySelector('[data-testid="stVerticalBlock"]').scrollTo({top: 9999, behavior: 'smooth'});
                    </script>''', height=0)

                    def render_progress(stage: int, progress_pct: float, elapsed: float, eta: float, current: int, total: int, current_item: str = ""):
                        """Render the enhanced progress UI with stages, stats, and animations."""
                        stage_classes = ["", "", ""]
                        if stage >= 1:
                            stage_classes[0] = "complete" if stage > 1 else "active"
                        if stage >= 2:
                            stage_classes[1] = "complete" if stage > 2 else "active"
                        if stage >= 3:
                            stage_classes[2] = "active"

                        progress_width = f"{progress_pct * 100:.1f}%"
                        elapsed_str = format_time(elapsed)
                        eta_str = format_time(eta) if eta > 0 else "--"
                        pct_str = f"{progress_pct * 100:.0f}%"

                        # Escape user-controlled data
                        safe_item = html.escape(current_item) if current_item else ""
                        current_item_html = f'<div class="current-item"><div class="pulse-dot"></div><span class="current-item-text">Processing {current}/{total}: {safe_item}</span></div>' if current_item else ''

                        markup = f'''<div class="progress-container"><div class="stage-indicators"><div class="stage {stage_classes[0]}"><div class="stage-dot {stage_classes[0]}">1</div><span class="stage-label">Loading</span></div><div class="stage {stage_classes[1]}"><div class="stage-dot {stage_classes[1]}">2</div><span class="stage-label">Processing</span></div><div class="stage {stage_classes[2]}"><div class="stage-dot {stage_classes[2]}">3</div><span class="stage-label">Finalizing</span></div></div><div class="progress-track"><div class="progress-fill" style="width: {progress_width}"></div></div><div class="stats-grid"><div class="stat-card"><div class="stat-value">{elapsed_str}</div><div class="stat-label">Elapsed</div></div><div class="stat-card"><div class="stat-value highlight">{pct_str}</div><div class="stat-label">Progress</div></div><div class="stat-card"><div class="stat-value">{eta_str}</div><div class="stat-label">Remaining</div></div></div>{current_item_html}</div>'''
                        progress_container.markdown(markup, unsafe_allow_html=True)

                    def render_completion(elapsed: float, total_items: int, file_size_kb: float):
                        """Render the completion celebration UI."""
                        elapsed_str = format_time(elapsed)
                        markup = f'''<div class="progress-container"><div class="progress-complete"><div class="completion-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"></polyline></svg></div><div class="completion-text">Generation Complete!</div><div class="completion-subtext">{total_items} slides • {file_size_kb:.1f} KB • {elapsed_str}</div></div><div class="stats-grid"><div class="stat-card"><div class="stat-value highlight">{total_items}</div><div class="stat-label">Combinations</div></div><div class="stat-card"><div class="stat-value">{elapsed_str}</div><div class="stat-label">Total Time</div></div><div class="stat-card"><div class="stat-value">{file_size_kb:.1f} KB</div><div class="stat-label">File Size</div></div></div></div>'''
                        progress_container.markdown(markup, unsafe_allow_html=True)

                    # Initial loading state
                    render_progress(1, 0.05, 0, 0, 0, 0, "Initializing...")

                    # Create queue for thread communication
                    queue: Queue = Queue()

                    # Start generation in background thread
                    thread = threading.Thread(
                        target=run_generation,
                        args=(str(template_path), str(input_path), str(output_path), queue)
                    )
                    thread.start()

                    # Track progress state
                    start_time = time.time()
                    total_combinations = 0
                    current_combination = 0
                    last_message = ""
                    current_brand = "Initializing..."
                    current_stage = 1
                    current_progress = 0.05
                    error = None

                    while thread.is_alive() or not queue.empty():
                        try:
                            msg = queue.get(timeout=0.2)

                            if msg[0] == "total":
                                total_combinations = msg[1]
                                current_stage = 2
                                current_progress = 0.1
                                current_brand = f"Found {total_combinations} combinations"

                            elif msg[0] == "progress":
                                current_combination = msg[1]
                                total_combinations = msg[2]
                                last_message = msg[3]
                                current_stage = 2

                                # Calculate progress (reserve 10% for finalization)
                                current_progress = min(0.9, 0.1 + (current_combination / total_combinations) * 0.8) if total_combinations > 0 else 0.1

                                # Extract brand info from message
                                brand_match = re.search(r": (.+) - (\d+)$", last_message)
                                if brand_match:
                                    current_brand = brand_match.group(1)
                                else:
                                    current_brand = f"Item {current_combination}"

                            elif msg[0] == "status":
                                current_brand = msg[1]

                            elif msg[0] == "done":
                                current_stage = 3
                                current_progress = 1.0
                                current_brand = "Finalizing..."

                            elif msg[0] == "error":
                                error = msg[1]

                        except Empty:
                            pass  # Queue timeout, will still update UI below

                        # Always update UI with current elapsed time
                        elapsed = time.time() - start_time
                        eta = 0
                        if current_combination > 0 and total_combinations > 0:
                            rate = elapsed / current_combination
                            eta = (total_combinations - current_combination) * rate

                        render_progress(current_stage, current_progress, elapsed, eta, current_combination, total_combinations, current_brand)

                    thread.join()

                    if error:
                        raise Exception(error)

                    # Read generated file
                    with open(output_path, "rb") as f:
                        pptx_bytes = f.read()

                    # Finalization stage
                    elapsed = time.time() - start_time
                    file_size_kb = len(pptx_bytes) / 1024

                    # Show completion celebration
                    render_completion(elapsed, total_combinations, file_size_kb)

                    # Download button with styled wrapper
                    st.markdown('<div class="download-section">', unsafe_allow_html=True)
                    st.download_button(
                        label="Download Presentation",
                        data=pptx_bytes,
                        file_name=out_filename,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        type="primary",
                        use_container_width=True,
                    )
                    st.markdown('</div>', unsafe_allow_html=True)

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
