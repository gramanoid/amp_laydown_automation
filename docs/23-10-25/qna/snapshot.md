Deck regeneration and merge sanitization are underway to restore compliant AMP baseline output.
• Project: Clone Lumina Excel into template-accurate decks via python pipeline plus COM post-processing (`tools/PostProcessCampaignMerges.ps1`, new `tools/SanitizePrimaryColumns.ps1`).
• Now: Finish sanitize-first pass—rerun the post-processor on `run_20251023_132502` with no residual span warnings, log the attempt, and capture row-height probe CSV showing ≤8.5 pt before updating docs.
• Risk: Post-process may still throw `Cell.Split` failures if spans linger; mitigate by ensuring sanitize-first outputs used and clearing POWERPNT sessions before retries.
