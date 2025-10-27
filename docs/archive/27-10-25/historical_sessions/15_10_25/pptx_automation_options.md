# PPTX Automation Options & Tech Stack Analysis (October 2025)

## Overview
This note summarizes current open-source and self-hostable choices for programmatic PowerPoint generation and highlights trade-offs between the main technology stacks available to the AMP automation workflow.

## Option Catalogue

- **AutoPPTX** *(PyPI-only, repo gone)* — Placeholder replacement toolkit formerly at `github.com/chenzhex/AutoPPTX`; wheel `autopptx==1.0.0` still downloadable but upstream is missing, so long-term maintenance is uncertain.
- **office-templates** *(Python, Apache-2.0)* — Jinja-style templating over native PPTX with automatic list/table expansion and chart hooks.
- **pptx-template (m3dev)** *(Python)* — Data-driven slide cloning system designed for BI decks and YAML/JSON payloads.
- **python-pptx-templater (kwlo)** *(Python)* — Lightweight JSON-to-layout mapper for deterministic table/text population.
- **gramex/pptgen2** *(Python)* — Rule-engine built atop python-pptx with strong Pandas integration and chart/image helpers.
- **pptx-automizer** *(Node.js, MIT)* — Combines template merging with PptxGenJS for chart/table creation; npm package is actively maintained (Feb 2025 release cadence).
- **docxtemplater + SlidesModule** *(Node.js, GPLv3/commercial)* — Mature placeholder DSL with loops/conditionals supporting DOCX/PPTX. GPLv3 acceptable for internal-only distribution.
- **cg123/pptx-api / ltc6539/mcp-ppt** *(Agent endpoints)* — HTTP/MCP services exposing JSON-driven slide generation for LLM agents.
- **PPTXTemplater** *(C#/.NET, MIT)* — OpenXML-based templating with slide cloning and precise shape control.
- **DefaultOpenXmlTemplateEngine & Clippit** *(C#/.NET, MIT)* — Low-level OpenXML manipulation libraries enabling granular styling and animation edits.
- **PowerPoint MCP Server** *(pywin32 COM bridge)* — Automates a local PowerPoint instance for 100% fidelity updates; Windows + Office dependency.

## Stack Comparisons

### Python-centric Stack
- **Pros:** Runs fully offline, aligns with our existing python-pptx code, integrates cleanly with Pandas/NumPy pipelines, permissive licenses (MIT/Apache), easy to containerize.
- **Cons:** Rendering fidelity limited to python-pptx capabilities (no native animation rendering), complex layouts may require additional scripting, fewer high-level abstractions for transitions/animation.
- **Fit:** Recommended baseline stack—migrate away from AutoPPTX toward `office-templates` + `pptx-template` + `python-pptx-templater`, with `gramex/pptgen2` for rule-heavy decks.

### Node.js / JavaScript Stack
- **Pros:** Rich templating DSLs (loops, conditionals), strong community momentum, integrations with Excel/CSV ingestion (`pptx-automizer` + `automizer-data`), easy pairing with LLM services, cross-platform.
- **Cons:** Requires Node toolchain alongside Python stack, docxtemplater GPLv3 imposes internal-use-only unless commercial license, relies on PptxGenJS approximations (slightly less precise than native PowerPoint).
- **Fit:** Ideal when we need declarative templates with heavy data binding or when Node infrastructure is already present; keep for advanced templating features and interplay with web tooling.

### .NET / OpenXML Stack
- **Pros:** Direct OpenXML manipulation gives pixel-perfect fidelity, strong control over animations, shapes, and themes, performant on Windows, permissive MIT licenses.
- **Cons:** Windows-first development story, steeper learning curve for OpenXML schema, fewer off-the-shelf helpers for data-driven pipelines, requires .NET runtime orchestration.
- **Fit:** Use when we must match template geometry exactly or control advanced PowerPoint features that python-pptx cannot reach; integrate via PowerShell or .NET service wrappers.

### Native PowerPoint Automation (COM/MCP)
- **Pros:** Runs the actual PowerPoint desktop renderer, so every feature (animations, charts, SmartArt) behaves exactly as in the UI; useful for QA diffs or last-mile fixes.
- **Cons:** Hard dependency on Microsoft Office installations, brittle automation via COM, not container-friendly, limited scalability, requires Windows worker nodes.
- **Fit:** Reserve for specialized cases needing native rendering validation or slides only PowerPoint itself can produce rather than day-to-day generation.

## Recommendations
- **Primary Path:** Standardize on the Python templating set (`office-templates`, `pptx-template`, `python-pptx-templater`) and retire AutoPPTX once replacements are scripted.
- **Advanced Templates:** Adopt `pptx-automizer` (MIT) when Node is acceptable; consider docxtemplater SlidesModule for internal-only GPLv3 use if we need its DSL.
- **High-Fidelity Edge Cases:** Keep .NET/OpenXML tools and the PowerPoint MCP bridge available for scenarios where python-pptx approximations fall short.
