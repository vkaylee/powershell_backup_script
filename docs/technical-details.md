# Technical Details

This document explains the internal architecture and workflow of the Snapshot Backup Script.

## Architecture

The script has been refactored into modular functions to improve maintainability and testability:

1.  **Configuration Management (`Get-Configuration`):**
    *   Loads `config.jsonc` (default).
    *   **JSONC Support:** Uses Regex to strip single-line (`//`) and multi-line (`/* */`) comments before parsing with `ConvertFrom-Json`.
2.  **Pre-flight Diagnostic Check (`Test-BackupPrerequisites`):**
    *   Validates system readiness before attempting any operations.
    *   Checks for `robocopy.exe`, Administrator privileges (if VSS enabled), and basic path accessibility.
    *   **Fail-safe:** If critical checks fail, the script terminates immediately to prevent inconsistent backups.
3.  **VSS Management:**
    *   `New-ShadowCopy`: Uses WMI `Win32_ShadowCopy` to create a point-in-time volume snapshot.
    *   `Remove-ShadowCopy`: Ensures cleanup of snapshots after use.
3.  **Backup Item Discovery (`Get-BackupItems`):**
    *   Supports **Granular Modes**:
        *   `SubDirectories`: Backs up each immediate child folder separately.
        *   `Root`: Backs up the source folder as a single unit.
    *   Translates live paths to VSS snapshot paths (`\\?\GLOBALROOT\Device\...`).
4.  **Execution Engine (`Execute-BackupItem`):**
    *   Wraps `Invoke-RobocopyBackup` with timestamping and logging logic.
    *   Ensures consistent destination folder structure: `DestinationPath\SourceShareName\ItemName_TIMESTAMP`.
5.  **Logging & Cleanup:**
    *   `Write-BackupHistory`: Appends results to `backup-history.log`. Each line is prepended with a human-readable timestamp `[YYYY-MM-DD HH:MM:SS]` followed by the JSON execution details.
    *   `Clean-OldBackups`: Enforces retention policies for both data and logs.

## VSS Snapshot Logic

1.  Identify the root volume (e.g., `D:\`).
2.  Create a "ClientAccessible" snapshot via WMI.
3.  Retrieve the `DeviceObject` (e.g., `\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy1`).
4.  Map the source relative path onto this device path to access frozen data.

## Robocopy Integration

Switches used:
- `/MIR`: Mirror (Copy everything, purge destination).
- `/DCOPY:DA`: Copy Directory Attributes (timestamps).
- `/COPY:DAT`: Copy Data, Attributes, Timestamps.
- `/IPG:n`: (Optional) Inter-Packet Gap for bandwidth throttling.

## Performance Tuning

- **Throttling:** Use `RobocopyInterPacketGapMs` in config to prevent 100% disk/network saturation during production hours.
- **Fail-safe:** If `UseVSS` is `false`, the script skips snapshot creation and copies directly from the source.

## Testing Strategy

The script uses **Pester (3.4.0)** for automated unit and integration testing. To ensure tests are reliable and runnable on standard developer machines (without Administrator privileges or specific drive configurations), the following strategy is employed:

- **VSS Mocking:** WMI calls for snapshot creation (`Win32_ShadowCopy`) are mocked to return fake `DeviceObject` paths. This verifies the script's path translation logic and snapshot lifecycle management.
- **Filesystem Isolation:** Core PowerShell cmdlets like `Join-Path`, `Test-Path`, and `New-Item` are mocked when testing logic that involves non-existent drives (like `E:\` or UNC paths). This allows the test suite to use project-standard paths without dependency on the physical host setup.
- **External Binary Mocking:** `robocopy.exe` and `Get-Command` are mocked to verify that the script correctly constructs command-line arguments and handles exit codes, without actually executing file copies.