## 1. Implementation
- [x] 1.1 Audit template contract requirements and extend config for per-cell styling metadata.
- [x] 1.2 Update table assembly to apply Template V4 fonts, alignments, fills, and dual-line rules.
- [x] 1.3 Enhance continuation-slide builder to retain headers, banding, and carried totals across splits.
- [x] 1.4 Regenerate or extend reconciliation checks and pytest coverage for styling + continuation behavior.
- [x] 1.5 Re-run `PYTEST_DISABLE_PLUGIN_AUTOLOAD=1 python -m pytest tests/test_template_validation.py`.

## 2. Validation
- [x] 2.1 `openspec validate update-table-styling-continuations --strict`

## 3. Slide 1 Parity Follow-up (21 Oct 2025)
- [ ] 3.1 Snap Slide 1 table row heights/column widths to template EMUs.
- [ ] 3.2 Center-align all table cell paragraphs post-population.
- [ ] 3.3 Prune synthesized legend shapes on slides where the template omits them.
- [ ] 3.4 Regenerate deck, rerun diff/validator/tests, and archive updated evidence.
