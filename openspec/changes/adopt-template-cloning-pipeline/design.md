## Overview
Switching to a clone-and-populate strategy ensures every generated slide inherits the master template's geometry. We duplicate template shapes, update only text/content portions, and step away from rebuilding layout primitives. As of 20 Oct 2025 the clone helpers append to `p:spTree` with relationship remapping, eliminating prior PowerPoint repair prompts.

## Key Decisions
- Shape cloning strategy: XML deep-copy appending to `p:spTree` while remapping embedded relationship IDs (`r:embed`, `r:link`, `r:id`); tables, tiles, and legend elements are cloned before population. Implementation lives in `template_clone.py::_clone_element` and helper wrappers in `assembly.py`.
- Styling fidelity: `tables.py::style_table_cell` mirrors Template V4 theme fills, summary tiles/footer reuse template text frames, and `_ensure_legend_shapes` rebuilds colour chips/text when grouped shapes are absent (falling back to measurement-driven rectangles). Palette and typography are config-driven (`fonts.legend_family`).
- Data binding: deterministic cell/shape name lookups; we update text runs only while geometry stays template-defined.
- Fallback behaviour: AutoPPTX is retained only as a manual contingency. Runtime uses the clone pipeline exclusively (`tooling.autopptx.enabled = false` by default); `_generate_autopptx_only` logs a warning but still honours explicit toggles. Guard test `tests/test_autopptx_fallback.py` protects the legacy path.
- Validation: structural assertions via `tools/validate_structure.py` (current clone deck passes); targeted pytest suite covers tables, assembly splitting, and AutoPPTX fallback. Visual diff tooling (`tools/visual_diff.py`) exports both decks, but full-deck comparison is blocked by missing template imagery. Zen MCP + PowerPoint Review > Compare remain in the validation plan once imagery lands.
- Structural contract (20 Oct 2025): Slide frame, column headers, media ordering, metric rows, styling palette, subtotal placement, footer tiles, legend, and footnotes are immutable. Campaign rows/values stay data-driven. Media sections appear only when populated; columns never vanish. Source footnote date derives from the Lumina export filename.

## Open Questions & Follow-ups
- Capture a multi-slide template baseline so visual diff and Zen MCP can assert parity across the deck.
- Confirm final styling on Slide 1 (alternating greys, legend/summary typography, footer spacing) once baseline imagery is available.
- Decide whether partial rebuild support (e.g., removing empty rows beyond cloning) is required in the near term.
- Chart regeneration remains disabled; revisit once business requirements surface.
