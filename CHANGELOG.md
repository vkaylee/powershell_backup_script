# Changelog

All notable changes to this project will be documented in this file.

## [1.2.0] - 2026-01-14
### Added
- **JSONC Support:** Configuration file now supports C-style comments (`//`, `/* */`).
- **Default Config Rename:** Default configuration file renamed to `config.jsonc` for better VS Code compatibility.
- **Granular Backup Modes:** Added `Root` mode (backup folder as one unit) vs `SubDirectories` mode (backup each subfolder).
- **Configuration Templates:** Added `config.template.simple.json` and `config.template.advanced.json` with documentation.
- **Enhanced Testing:** Expanded Pester test suite covering logic, config, and cleanup.

### Changed
- **Modular Refactoring:** Extracted monolithic `Start-BackupProcess` into `Get-BackupItems` and `Execute-BackupItem` for better testability.
- **Improved Path Resolution:** Script now checks for `config.jsonc` first and falls back to `config.json` if necessary.

### Fixed
- Fixed array unrolling issue in `Get-BackupItems` when returning a single item.

## [1.1.0] - 2026-01-13
### Added
- Initial VSS Snapshot integration.
- Robocopy error handling and exit code mapping.
- Retention policy implementation.
- JSON history logging.
