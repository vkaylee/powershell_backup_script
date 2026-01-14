# User Guide

## Configuration (`config.jsonc`)

The script is controlled entirely by `config.jsonc` (formerly `config.json`). Using the `.jsonc` extension allows VS Code to support comments without displaying errors.

| Setting | Type | Description | Default |
| :--- | :--- | :--- | :--- |
| `SourcePaths` | Array | List of directory paths (Strings) OR configuration objects. | `[]` |
| `RobocopyOptions` | String | Extra flags passed to Robocopy. (See note below). | `/MIR /MT:8...` |

### Robocopy Options & Protected Flags
While you can customize `RobocopyOptions` (e.g., adding `/XF` to exclude files or `/XD` to exclude directories), certain flags are **Protected** to ensure script reliability:
*   **Overridden:** `/R`, `/W`, `/MT`, `/LOG`, `/LOG+`.
*   The script automatically strips these from your configuration and enforces its own safe defaults (`/R:1 /W:1`, `/MT:8`, and internal log management).

### Comments Support
The configuration file supports standard C-style comments (`//` for single line, `/* ... */` for blocks). This allows you to document your settings directly in the JSON file.

### Granular Backup Modes
... (rest of the file)
You can specify how each source path is backed up:
*   **SubDirectories (Default):** Iterates through immediate subdirectories and backups each one separately.
*   **Root:** Backups the specified folder itself as a single unit.

**Advanced Configuration Example:**
```json
{
  "SourcePaths": [
    "D:\\SimplePath", 
    {
      "Path": "D:\\Shares\\Departments",
      "Mode": "SubDirectories" 
    },
    {
      "Path": "D:\\Shares\\CriticalApp",
      "Mode": "Root" 
    }
  ],
  ...
}
```

### Example `config.json`
```json
{
  "SourcePaths": [
    "D:\\DepartmentShares\\Finance",
    "D:\\DepartmentShares\\HR"
  ],
  "DestinationPath": "\\\\NAS01\\Backups\\Daily",
  "RetentionDays": 60,
  "LogRetentionDays": 180,
  "UseVSS": true,
  "RobocopyInterPacketGapMs": 20
}
```

---

## Verifying Configuration
Before running a full backup, you can use the `-CheckOnly` switch to verify your environment. This checks for:
1.  **Robocopy Availability:** Ensures `robocopy.exe` is in the system PATH.
2.  **Permissions:** Checks if the script is running with necessary privileges (Administrator required for VSS).
3.  **Paths:** Validates accessibility of source and destination paths.

**Command:**
```powershell
.\backup-script.ps1 -ConfigFilePath config.jsonc -CheckOnly
```

**Output Example:**
```text
Running Pre-flight Diagnostic Checks...
  [OK] Robocopy detected.
  [OK] Administrator privileges confirmed for VSS.
  [OK] Destination path is accessible.

[SUCCESS] System is ready for backup.
```

## Scheduling (Windows Task Scheduler)

To run this script automatically every day:

1.  Open **Task Scheduler** (`taskschd.msc`).
2.  Click **Create Task**.
3.  **General Tab:**
    *   Name: `Daily Snapshot Backup`
    *   **IMPORTANT:** Check **"Run with highest privileges"** (Required for VSS).
    *   Select **"Run whether user is logged on or not"**.
4.  **Triggers Tab:**
    *   New... -> Daily -> Set time (e.g., 23:00).
5.  **Actions Tab:**
    *   New... -> Start a program.
    *   **Program/script:** `powershell.exe`
    *   **Add arguments:** `-ExecutionPolicy Bypass -File "D:\tuanlee\snapshot_backup_script\backup-script.ps1" -ConfigFilePath "config.jsonc"`
    *   **Start in:** `D:\tuanlee\snapshot_backup_script\` (Directory where script lives).

---

## Restoring Data

Backups are standard file folders, making restoration easy.

1.  Navigate to your `DestinationPath`.
2.  Open the folder matching your source share name (e.g., `Finance`).
3.  Find the folder with the desired timestamp (e.g., `Reports_20260114_230000`).
4.  Copy the files you need back to the original location.
