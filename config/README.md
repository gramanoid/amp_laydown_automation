# Configuration Folder

**Last Updated**: October 7, 2025  
**Configuration Version**: 2.0.0

## Single Master Configuration File

**master_config.json** - Complete configuration for AMP automation system

### Sections:
- Project metadata
- Paths (input, output, logs, backup, temp) with auto-creation
- Script and template settings
- Presentation formatting (table, fonts, colors, charts)
- Data column mappings and processing rules
- Performance thresholds and monitoring
- Logging configuration
- Output file naming and organization
- Error handling and recovery
- Feature flags
- Platform settings
- Startup/shutdown behaviors
- Development options

### Deleted Files (Oct 7, 2025):
- automation_config.json (6.9 KB) - Obsolete router config
- performance_thresholds.json (1.4 KB) - Consolidated
- presentation_config.json (731 B) - Consolidated
- platform_config.json (11.1 KB) - Reference only
- tasks.json (14.4 KB) - Misplaced task list

**Result**: 6 files → 1 file, 34 KB saved, zero redundancies

### Usage:
Edit master_config.json for all configuration changes. All paths support auto-creation. Performance thresholds can be adjusted per environment.

