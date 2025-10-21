# Tooling Research Reset (Current Snapshot)
**Last Updated:** 2025-10-15 18:05

## Session Overview
- Re-evaluated third-party PPTX automation libraries after discovering AutoPPTX’s GitHub source is unavailable.
- Compiled a comprehensive tech-stack comparison covering Python, Node.js, .NET/OpenXML, and native PowerPoint automation options.
- Documented findings in `pptx_automation_options.md` for long-term reference and decision-making.

## Work Completed
- Performed deep EXA and web searches to validate availability, licensing, and maintenance status of candidate libraries (office-templates, pptx-template, pptx-automizer, docxtemplater, PPTXTemplater, Clippit, PowerPoint MCP Server, etc.).
  - Artifacts: `D:\OneDrive - Publicis Groupe\work\(.) AMP Laydowns Automation\docs\15_10_25\pptx_automation_options.md`
- Assessed AutoPPTX risk posture (PyPI-only distribution) and confirmed adapter gracefully degrades when dependency missing.
  - Artifacts: `amp_automation\tooling\autopptx_adapter.py`

## Current State
- AutoPPTX remains installed but is now treated as deprecated because upstream source vanished.
- Open-source alternatives are cataloged with pros/cons; Python stack identified as primary migration path.
- Node.js and .NET options logged for supplemental use cases; COM bridge flagged as last-resort.

## Purpose
- Ensure presentation generation pipeline relies on sustainable, privacy-preserving, and actively maintained tooling.

## Next Steps
- [ ] Prototype `office-templates` integration to replace AutoPPTX placeholder workflow.
- [ ] Spike Node.js `pptx-automizer` integration for advanced templating scenarios (conditional slides, chart merges).
- [ ] Produce recommendation memo selecting final toolkit mix for production adoption.
- [ ] Remove AutoPPTX dependency once replacement path is validated.

## Important Notes
- GPLv3 tooling (docxtemplater) is acceptable only while outputs remain internal; distribution would require commercial licensing.
- Preserve a frozen copy of `autopptx==1.0.0` wheel for reproducibility until migration completes.
- Python-based templaters align best with existing pipeline, but .NET/OpenXML tools provide highest fidelity when needed.

## Session Metadata
- Date: 2025-10-15
- Location: `D:\OneDrive - Publicis Groupe\work\(.) AMP Laydowns Automation\docs\15_10_25`
- Tools: Exa search, web browser, ApplyPatch, python
