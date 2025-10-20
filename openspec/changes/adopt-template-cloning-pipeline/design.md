## Overview
Switching to a clone-and-populate strategy ensures every generated slide inherits the master template’s geometry. We duplicate template shapes, update only text/content portions, and step away from rebuilding layout primitives. As of 20 Oct 2025 the clone helpers append to `p:spTree` with relationship remapping, eliminating prior PowerPoint repair prompts.

## Key Decisions
- Shape cloning strategy: XML deep-copy appending to `p:spTree` while remapping embedded relationship IDs (`r:embed`, `r:link`, `r:id`); table + tiles cloned before population. Implemented in `template_clone.py::_clone_element` and used by table/tile helpers.
- Data binding: deterministic cell/shape name lookups; update text runs only; geometry remains template-defined.
- Fallback behavior: AutoPPTX kept optional; legacy builder removed from main path. Config toggle still pending (OpenSpec task 1.4).
- Validation: per-slide visual diff with thresholds, structural assertions in tests, `tools/inspect_generated_deck.py` for corruption diagnostics. Table border styling now uses minimal DrawingML so clone decks open cleanly via PowerPoint COM; need full template baseline to finish diff comparison.

## Open Questions & Follow-ups
- Partial rebuild capabilities for dynamic layouts (e.g., removing empty rows) still unresolved; current approach relies on template duplication with selective population.
- Chart regeneration remains disabled; decision pending data requirements.
- Template baseline imagery: template export still produces a single master-slide PNG, so diff thresholds fail. Need 114-slide reference set (golden deck or rendered template) before parity metrics are meaningful.
- Configuration toggle design: where to surface user control (CLI flag vs. config file) and how to ensure safe fallback.
