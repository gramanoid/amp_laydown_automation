<!-- OPENSPEC:START -->
# OpenSpec Instructions

These instructions are for AI assistants working in this project.

## Quick Project Recap (21 Oct 2025)
- **Mission:** Clone-based generation of AMP laydown decks that match `Template_V4_FINAL_071025.pptx` pixel-for-pixel while binding Lumina Excel data; AutoPPTX remains disabled except for negative tests.
- **Latest Baseline Deck:** `output/presentations/run_20251021_102357/GeneratedDeck_20251021_102357.pptx`, passing structural validation; visual diff still flags Slide 1 due to row-height/alignment/legend differences.
- **Current Focus:** Reset table row heights/column widths to template EMUs, center-align all cells, drop the synthesized legend on Slide 1, capture multi-slide template baseline, rerun visual diff + Zen MCP/Compare, extend pytest coverage, and execute multi-market pipeline smoke tests.

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
