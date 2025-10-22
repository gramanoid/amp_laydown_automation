Summary: Built and ran the COM post-processor to enforce deck fonts/percents while tracing the lingering row-height variance caused by non-idempotent monthly-total merges.
Docs Updated:
- docs/21_10_25/BRAIN_RESET_211025.md: refreshed snapshot, current position, and phased plan
- docs/21_10_25/21_10_25.md: rewritten end-of-day summary, blockers, and repository map
- docs/daily-logs/10-21-25.md: logged daily highlights through next steps under the standard headings
Outstanding:
- Now: Implement the column 1–3 unmerge pre-pass in tools/PostProcessCampaignMerges.ps1 and rerun the COM probe on the latest deck.
- Next: Add automated merge/height validation and regenerate Slide 1 visual diff artefacts once geometry stabilises.
- Later: Restore the pytest suite, expand visual diff masking, and document the COM workflow in the runbook.
Insights: PowerPoint COM retains prior merge geometry; automation must reset tables before applying new spans to stay deterministic.
Validation:
- Tests: Not run today; pytest suite currently absent and needs restoration before automation can signal regressions.
- Deploy: Not applicable; no production deployment attempted.
Git: Working tree dirty (modified/deleted files, removed tests, new docs); not ready for /12-[git-stage]-git-commit or /14-[git-push]-git-push until the COM fix lands and changes are reviewed. Next command: git status.
Tomorrow: /06-[delivery]-next-task to jump straight into the unmerge pre-pass implementation and verification work.
STATUS: OK | NEXT: Implement COM unmerge pre-pass then rerun row-height probe
