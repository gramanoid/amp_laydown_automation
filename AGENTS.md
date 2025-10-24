<!-- OPENSPEC:START -->
# OpenSpec Instructions

These instructions are for AI assistants working in this project.

## Quick Project Recap (24 Oct 2025)
Last verified on 24-10-25
- **Mission:** Clone-based generation of AMP laydown decks that match `Template_V4_FINAL_071025.pptx` pixel-for-pixel while binding Lumina Excel data; AutoPPTX remains disabled except for negative tests.
- **Latest Baseline Deck:** `output/presentations/run_20251024_161355/presentations.pptx` (88 slides, 63 brand/market combinations), passing structural validation.
- **CRITICAL ARCHITECTURE DECISION (24 Oct 2025):** PowerPoint COM automation for bulk operations is PROHIBITED due to catastrophic performance issues (10+ hours vs Python's 10 minutes - 60x difference). See `docs/ARCHITECTURE_DECISION_COM_PROHIBITION.md`.
- **Current Focus:** Migrating post-processing from PowerShell COM to Python (python-pptx). Table normalization complete; merge operations and span resets in progress. Target: <20 minute end-to-end pipeline.

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
