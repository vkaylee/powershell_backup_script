# Changelog

All notable changes to this project will be documented in this file.

## [1.2.0] - 2026-01-14
### Added
- **Usage Guide:** Script now displays a helpful usage guide and examples if executed without any command-line arguments.
- **Pre-flight Diagnostic Check:** New `Test-BackupPrerequisites` function ensures `robocopy.exe` exists, Admin privileges are active (for VSS), and paths are valid before running.
- **Diagnostic Mode:** Added `-CheckOnly` parameter to run only the system checks and exit.
- **JSONC Support:** Configuration file now supports C-style comments (`//`, `/* */`).
- **Default Config Rename:** Default configuration file renamed to `config.jsonc` for better VS Code compatibility.
- **Granular Backup Modes:** Added `Root` mode (backup folder as one unit) vs `SubDirectories` mode (backup each subfolder).
- **Configuration Templates:** Added `config.template.simple.json` and `config.template.advanced.json` with documentation.
- **Enhanced Testing:** Expanded Pester test suite to 25 tests covering logic, config, and junction workflows.
- **VSS Junction Mapping:** Implemented automatic directory junction creation (`mklink /j`) for VSS snapshots. This resolves persistent "Syntax Incorrect" (Error 123) and "Path Not Found" (Error 53) issues in Robocopy by providing a standard local path to the snapshot data.
- **VSS Integration Testing:** Implemented comprehensive mocks for VSS snapshots and junctions, ensuring lifecycle management is verified in CI/CD.
- **Project-Standard Test Paths:** Tests now use project-standard drive letters (`D:`, `E:`) and mock the filesystem to ensure compatibility across different environments.
- **High-Precision Timestamps:** Switched to millisecond-precision timestamps (`yyyyMMdd_HHmmss_fff`) for backup folders to prevent data corruption/purging during rapid subfolder backups under `/MIR`.
- **Reliability Fixes:** Resolved Robocopy Error 16 (parameter parsing), Error 53 (path not found), and Error 123 (invalid syntax) by refactoring argument passing, ensuring mandatory `\\?\` prefixes, and normalizing trailing slashes.
- **Robust Option Filtering:** Implemented a "Protected Flags" system that strips conflicting user-provided options (`/R`, `/W`, `/MT`, `/LOG`) from the configuration to ensure script stability.
- **Fail-Fast Enforcement:** Hardcoded `/R:1 /W:1` into core operations to prevent indefinite hangs on inaccessible network or VSS paths.
- **High-Resolution Logging:** Added detailed diagnostic logging with millisecond-precision timestamps throughout the script for better troubleshooting.
- **Performance Optimization:** Optimized default Robocopy switches with Multi-threading (`/MT:8`) and Unbuffered I/O (`/J`) for faster 300MB+ transfers.
- **Timestamped History:** Prepend human-readable timestamps to `backup-history.log` entries for better readability.

### Changed
- **Config Path Mandatory:** The `-ConfigFilePath` parameter is now required. The script will display a usage guide if it is omitted.
- **Modular Refactoring:** Extracted monolithic `Start-BackupProcess` into `Get-BackupItems` and `Execute-BackupItem` for better testability.
- **Improved Path Resolution:** Script now checks for `config.jsonc` first and falls back to `config.json` if necessary.

### Fixed
- **Array Unrolling Bug:** Fixed a critical issue where single-item arrays returned by `Get-BackupItems` were being unrolled, causing processing failures.
- **Discovery Logic Fix:** Resolved an issue where the script would inadvertently backup the current working directory if certain variables were undefined.

## [1.1.0] - 2026-01-13
### Added
- Initial VSS Snapshot integration.
- Robocopy error handling and exit code mapping.
- Retention policy implementation.
- JSON history logging.
