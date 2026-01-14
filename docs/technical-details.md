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
    *   **Collision Prevention:** Uses high-precision timestamps (`_fff`) to ensure unique destination folders for each item, even when multiple subfolders are processed within the same second.
    *   Ensures consistent destination folder structure: `DestinationPath\SourceShareName\ItemName_TIMESTAMP`.
5.  **Logging & Cleanup:**
    *   `Write-BackupHistory`: Appends results to `backup-history.log`. Each line is prepended with a human-readable timestamp `[YYYY-MM-DD HH:MM:SS]` followed by the JSON execution details.
    *   **High-Resolution Diagnostic Logging:** The script outputs detailed progress logs with millisecond-precision timestamps (`HH:mm:ss.fff`) to the console, allowing for precise identification of performance bottlenecks or hangs.
    *   `Clean-OldBackups`: Enforces retention policies for both data and logs. Supports regex-based identification of high-precision timestamp folders.

## VSS Snapshot Logic
The script uses the following workflow to ensure data consistency:
1.  Identify the volume root (e.g., `D:\`) for the given source path.
2.  Create a VSS Snapshot using WMI (`Win32_ShadowCopy`).
3.  **Junction Mapping:** To avoid Robocopy Error 123/53 (common with `\\?\GLOBALROOT` paths), the script creates a temporary Directory Junction (`mklink /j`) pointing to the snapshot's `DeviceObject`.
4.  Translate the source path to use the junction (e.g., `D:\VssJunction_{ID}\path\to\data`).
5.  Execute backup.
6.  **Cleanup:** Remove the junction (`rd /q`) and delete the VSS snapshot.

### VSS Junction Strategy
Directory Junctions are used because they are handled at the filesystem level and are fully compatible with standard Win32 tools like Robocopy and `cmd.exe`, whereas direct device namespace paths are often rejected by these tools.

## Robocopy Integration
...

### Array Unrolling Management
A critical fix was implemented in `Get-BackupItems` to handle PowerShell's automatic array unrolling. When returning a single backup item, the script uses the unary comma operator (`,$ItemsToBackup`) to ensure the calling function receives a consistent array object, preventing iteration errors.

1.  Identify the root volume (e.g., `D:\`).
2.  Create a "ClientAccessible" snapshot via WMI.
3.  Retrieve the `DeviceObject` (e.g., `\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy1`).
4.  Map the source relative path onto this device path to access frozen data.
    *   **Crucial:** The mandatory `\\?\` prefix is preserved to ensure absolute Win32 device path resolution.

## Robocopy Integration

### Robust Option Filtering
To prevent user configuration from breaking the script's core logic, `Invoke-RobocopyBackup` implements a filtering layer. Protected flags are automatically stripped from the `RobocopyOptions` string provided in the config file:
- **Protected:** `/R`, `/W` (Retries), `/MT` (Multi-threading), `/LOG`, `/LOG+` (Logging).
- **Mandatory:** The script enforces its own values for these flags (`/R:1 /W:1`, `/MT:8`, and timestamped log paths) to guarantee stability and prevent hangs.

### Path Normalization
Win32 device paths (VSS) are extremely sensitive to syntax. The script ensures:
1.  **Mandatory Prefix:** `\\?\` is preserved for absolute device resolution.
2.  **Trailing Slash Logic:** Device paths ending in a subfolder (e.g., `\\?\...\workdirs\ADM`) must **not** have a trailing backslash when quoted, or Robocopy will throw Error 123. The script automatically sanitizes these paths before execution.

Switches used:
- `/MIR`: Mirror (Copy everything, purge destination).
- `/DCOPY:DA`: Copy Directory Attributes (timestamps).
- `/COPY:DAT`: Copy Data, Attributes, Timestamps.
- `/MT:8`: Multi-threaded copy (8 threads).
- `/J`: Unbuffered I/O (faster for large/numerous files).
- `/IPG:n`: (Optional) Inter-Packet Gap for bandwidth throttling.
- `/R:1 /W:1`: Rapid retry logic to prevent long hangs on inaccessible VSS paths.

## Performance Tuning

- **Throttling:** Use `RobocopyInterPacketGapMs` in config to prevent 100% disk/network saturation during production hours.
- **Fail-safe:** If `UseVSS` is `false`, the script skips snapshot creation and copies directly from the source.

## Testing Strategy

The script uses **Pester (3.4.0)** for automated unit and integration testing. To ensure tests are reliable and runnable on standard developer machines (without Administrator privileges or specific drive configurations), the following strategy is employed:

- **VSS Mocking:** WMI calls for snapshot creation (`Win32_ShadowCopy`) are mocked to return fake `DeviceObject` paths. This verifies the script's path translation logic and snapshot lifecycle management.
- **Filesystem Isolation:** Core PowerShell cmdlets like `Join-Path`, `Test-Path`, and `New-Item` are mocked when testing logic that involves non-existent drives (like `E:\` or UNC paths). This allows the test suite to use project-standard paths without dependency on the physical host setup.
- **External Binary Mocking:** `robocopy.exe` and `Get-Command` are mocked to verify that the script correctly constructs command-line arguments and handles exit codes, without actually executing file copies.