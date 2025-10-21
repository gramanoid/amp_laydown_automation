**Goal**
Upgrade the deck automation workflow with the 2025-era toolkits that improve template fidelity and QA.

**Plan**
1. **AutoPPTX integration**
   - Add AutoPPTX as a dependency and scaffold a JSON-driven presentation builder (Excel → normalized JSON → AutoPPTX run).
   - Map existing template placeholders (title blocks, summary tiles, table frames) to AutoPPTX schema; keep our python-pptx table assembly as a custom handler.
   - Expose CLI flags to choose between “legacy” and “autopptx” rendering paths.

2. **Aspose.Slides pixel renderer**
   - Introduce an optional module using `aspose.slides` (via Python/Java) to regenerate the final PPTX and export slide PNGs with Office-grade fidelity.
   - Provide configuration for license key (env var) and fall back gracefully if unavailable.

3. **DocStrange QA utilities**
   - Add a `tools/ppt_validate.py` helper that leverages DocStrange to extract tables/text from both template and generated slides.
   - Compare extracted structures to highlight visual or data drift; integrate into CI-style check.

4. **Pipeline glue & docs**
   - Update CLI workflow to orchestrate: Excel → JSON → AutoPPTX → optional Aspose export → DocStrange diff.
   - Document configuration toggles and provide sanity tests.
