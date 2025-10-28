<!-- OPENSPEC:START -->
# OpenSpec Instructions

These instructions are for AI assistants working in this project.

## Quick Project Recap (28 Oct 2025)
Last verified on 28-10-25 (evening)
- **Mission:** Clone-based generation of AMP laydown decks that match `Template_V4_FINAL_071025.pptx` pixel-for-pixel while binding Lumina Excel data; AutoPPTX remains disabled except for negative tests.
- **Latest Baseline Deck:** `output/presentations/run_20251028_163719/AMP_Laydowns_281025.pptx` (127 slides, with all formatting/output improvements).
- **CRITICAL ARCHITECTURE DECISION (24 Oct 2025):** PowerPoint COM automation for bulk operations is PROHIBITED due to catastrophic performance issues (10+ hours vs Python's 10 minutes - 60x difference). See `docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md`.
- **Session 28 Oct Status (ALL TASKS COMPLETED & ARCHIVED):**
  - ✅ 6 formatting improvements (bold columns, merged percentages, quarterly formatting, output naming, footer dates)
  - ✅ Test suite rehydration (16 regression tests: 8 formatting, 3 structural, 5 footer extraction)
  - ✅ validate_structure.py PROJECT_ROOT bug fix
  - 🗂️ **Archived/Deferred (Phase 4+):** Slide 1 geometry parity, visual diff workflow, automated regression scripts, Python normalization expansion, smoke tests with additional markets
- **Production Status:** All 6 improvements validated in `run_20251028_163719/AMP_Laydowns_281025.pptx` (127 slides). Pipeline complete.

Always open `@/openspec/AGENTS.md` when the request:
- Mentions planning or proposals (words like proposal, spec, change, plan)
- Introduces new capabilities, breaking changes, architecture shifts, or big performance/security work
- Sounds ambiguous and you need the authoritative spec before coding

Use `@/openspec/AGENTS.md` to learn:
- How to create and apply change proposals
- Spec format and conventions
- Project structure and guidelines

Keep this managed block so 'openspec update' can refresh the instructions.

<!-- OPENSPEC:END -->
