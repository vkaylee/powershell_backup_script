# PowerShell Snapshot Backup Script

A robust, enterprise-grade PowerShell script for performing daily snapshot backups of Windows SMB shares (or local directories). It leverages **Volume Shadow Copy Service (VSS)** to ensure data consistency for open files and **Robocopy** for efficient, resilient file transfer.

## 🚀 Features

- **VSS Snapshot Integration:** Creates a consistent "point-in-time" snapshot of the source volume before copying.
- **VSS Junction Mapping:** Automatically creates temporary directory junctions to VSS snapshots, ensuring 100% compatibility with Robocopy and resolving common Win32 pathing errors (Error 123/53).
- **Efficient File Copy:** Uses `Robocopy` for multi-threaded, resume-supported file transfers.
- **Configurable Throttling:** Built-in support for Robocopy's `/IPG` (Inter-Packet Gap) to prevent network/server overload.
- **Pre-flight Diagnostics:** Auto-detects missing tools or permissions and aborts to prevent failures. Use `-CheckOnly` to verify environment readiness.
- **High-Resolution Logging:** Outputs detailed progress with millisecond-precision timestamps (`HH:mm:ss.fff`) for precise troubleshooting and performance monitoring.
- **Timestamped Backups:** Creates isolated backup folders for each run with high-precision (ms) timestamps to ensure uniqueness (e.g., `ProjectA_20260114_080000_123`).
- **Automated Retention:** Automatically deletes backups and logs older than X days (configurable).
- **Self-Documenting Configuration:** Supports **C-style comments** (`//` and `/* */`) in the `config.jsonc` file for better documentation.
- **Detailed History:** Maintains a JSON-based `backup-history.log` and granular logs for each Robocopy operation.
- **Fail-safe Mode:** Can fallback to direct copying if VSS is disabled/unavailable (configurable).

## 📋 Prerequisites

- **OS:** Windows Server or Windows 10/11.
- **Permissions:** **Administrator privileges** are required to create VSS Snapshots.
- **PowerShell:** Windows PowerShell 5.1 or PowerShell Core.

## 🛠️ Quick Start

1.  **Configure:**
    Edit `config.jsonc` to set your source and destination paths.
    ```jsonc
    {
      // My Source Folders
      "SourcePaths": ["D:\\Shares\\Data"],
      "DestinationPath": "\\\\BackupServer\\\\Store",
      "RetentionDays": 30,
      "UseVSS": true
    }
    ```

2.  **Run:**
    Open PowerShell as **Administrator** and execute:
    ```powershell
    .\backup-script.ps1 -ConfigFilePath config.jsonc
    ```

## 🧪 Testing

The project includes a comprehensive Pester test suite. To run tests:

```powershell
Invoke-Pester .\backup-script.Tests.ps1
```

The tests cover configuration parsing, VSS logic (mocked), Robocopy execution, and retention cleanup.

## 📂 Documentation

- [User Guide](docs/user-guide.md) - Detailed configuration and scheduling via Task Scheduler.
- [Technical Details](docs/technical-details.md) - How VSS and Robocopy logic works under the hood.
- [Changelog](CHANGELOG.md) - Recent updates and fixes.

## 📝 License

Internal Use / Open Source.
