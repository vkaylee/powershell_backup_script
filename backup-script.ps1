#region Script Header
<#
.SYNOPSIS
    PowerShell script for Snapshot Backup using VSS and Robocopy.
    Ensures data consistency with VSS and efficient transfer with Robocopy.
    Supports configurable retention policies and detailed logging.

.DESCRIPTION
    This script performs snapshot backups of specified source directories.
    It leverages Volume Shadow Copy Service (VSS) to create consistent snapshots
    of volumes containing the source data, ensuring integrity even for open files.
    Robocopy is then used to efficiently copy data from these VSS snapshots
    to a timestamped destination, supporting large data volumes.
    The script includes features for:
    - Configurable source and destination paths.
    - Configurable retention policy for backups and logs.
    - Detailed logging of backup operations and Robocopy output.
    - "Lightweight" operation using Robocopy's /IPG option to prevent server overload.
    - Restoration facilitated by self-contained, timestamped backup folders and a detailed history log.

.PARAMETER ConfigFilePath
    Path to the JSON configuration file. Defaults to 'config.json' in the script directory.

.EXAMPLE
    .\backup-script.ps1
    Runs the backup using the default config.json.

.EXAMPLE
    .\backup-script.ps1 -ConfigFilePath "C:\MyBackupConfig\custom_config.json"
    Runs the backup using a custom configuration file.

.NOTES
    Requires Administrator privileges to create and manage VSS snapshots.
    Ensure Robocopy.exe is available on the system PATH.
#>
#endregion Script Header

Param (
    [Parameter(Mandatory=$false)]
    [string]$ConfigFilePath = "config.jsonc"
)

#region Global Variables and Constants
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
# Resolve absolute path for config file if it's relative
if (-not (Test-Path $ConfigFilePath -PathType Leaf)) {
    # Check if it exists relative to script dir
    $PotentialPath = Join-Path $ScriptDir $ConfigFilePath
    if (Test-Path $PotentialPath -PathType Leaf) {
        $ConfigFilePath = $PotentialPath
    } elseif ($ConfigFilePath -eq "config.jsonc" -and (Test-Path (Join-Path $ScriptDir "config.json"))) {
        # Fallback to config.json if jsonc is default but only json exists
        $ConfigFilePath = Join-Path $ScriptDir "config.json"
    }
}
$LogsDir = Join-Path $ScriptDir "logs"
$Config = $null # Will store loaded configuration
$HistoryLogFilePath = $null # Will be set after config is loaded
#endregion Global Variables and Constants

#region Functions - Configuration Management

Function Get-Configuration {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [string]$ConfigPath
    )

    Write-Host "Attempting to load configuration from: $ConfigPath" -ForegroundColor Cyan

    $DefaultConfig = @{
        SourcePaths = @("C:\BackupSources\TestShare1", "C:\BackupSources\TestShare2")
        DestinationPath = (Join-Path $ScriptDir "Backups")
        RetentionDays = 30
        LogRetentionDays = 90
        MaxBackupAttempts = 3
        RobocopyOptions = '/MIR /XD RECYCLE.BIN "System Volume Information" /XF Thumbs.db /R:5 /W:10 /NP /LOG+:{logpath}'
        RobocopyInterPacketGapMs = 0
        HistoryLogFile = "backup-history.log"
        UseVSS = $true
    }

    if (-not (Test-Path $ConfigPath -PathType Leaf)) {
        Write-Warning "Configuration file not found at '$ConfigPath'. Using default configuration."
        return $DefaultConfig
    }

    try {
        $JsonContent = Get-Content -Raw -Path $ConfigPath
        
        # Pre-process JSON to strip comments (JSONC support)
        # 1. Strip multi-line comments /* ... */
        $JsonContent = $JsonContent -replace '(?s)/\*.*?\*/', ''
        # 2. Strip single-line comments // (anchored to ensure we don't break simple urls if strictly needed, 
        #    but generally assumes // is comment if preceded by whitespace or start of line)
        #    Note: This simple regex might match // inside strings if preceded by space. 
        #    For this specific tool's config (paths/options), it's a safe trade-off for usability.
        $JsonContent = $JsonContent -replace '(?m)^\s*//.*$', '' -replace '(?m)\s+//.*$', ''

        $LoadedConfig = $JsonContent | ConvertFrom-Json -ErrorAction Stop

        # Merge loaded config with defaults, ensuring all keys are present
        foreach ($key in $DefaultConfig.Keys) {
            if ($null -eq $LoadedConfig.$key) {
                # Property missing or null, add it from defaults
                $LoadedConfig | Add-Member -MemberType NoteProperty -Name $key -Value $DefaultConfig.$key -Force
            }
        }
        
        # Validate essential paths
        if (-not $LoadedConfig.SourcePaths -or $LoadedConfig.SourcePaths.Count -eq 0) {
            Write-Warning "No SourcePaths defined in configuration. Using default source paths."
            $LoadedConfig.SourcePaths = $DefaultConfig.SourcePaths
        }

        # Handle DestinationPath validation and creation
        if (-not [string]::IsNullOrEmpty($LoadedConfig.DestinationPath)) {
             if (-not (Test-Path $LoadedConfig.DestinationPath -PathType Container)) {
                Write-Warning "DestinationPath invalid or not found: '$($LoadedConfig.DestinationPath)'. Attempting to create it."
                try {
                    New-Item -Path $LoadedConfig.DestinationPath -ItemType Directory -Force | Out-Null
                } catch {
                    Write-Error "Failed to create destination path '$($LoadedConfig.DestinationPath)'. Error: $($Error[0].Exception.Message). Please ensure it's accessible and you have permissions."
                    exit 1
                }
            }
        } else {
            Write-Error "Configuration Error: DestinationPath is null or empty. Please provide a valid path."
            exit 1
        }
        return $LoadedConfig
    }
    catch {
        Write-Error "Failed to load or parse configuration file '$ConfigPath'. Error: $($Error[0].Exception.Message). Using default configuration."
        return $DefaultConfig
    }
}

#endregion Functions - Configuration Management

#region Functions - VSS Management

Function Get-VolumeRoot {
    Param([string]$Path)
    $FullPath = (Resolve-Path $Path).Path
    $Drive = Split-Path $FullPath -Qualifier
    return "$Drive\"
}

Function New-ShadowCopy {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [string]$VolumeRoot
    )

    Write-Host "Creating VSS Snapshot for volume: $VolumeRoot" -ForegroundColor Cyan
    try {
        $ShadowCopyClass = [wmiclass]"root\cimv2:Win32_ShadowCopy"
        $Result = $ShadowCopyClass.Create($VolumeRoot, "ClientAccessible")
        
        if ($Result.ReturnValue -eq 0) {
            $ShadowID = $Result.ShadowID
            Write-Host "VSS Snapshot created successfully. ID: $ShadowID" -ForegroundColor Green
            
            # Get the Snapshot Object to retrieve the DeviceObject path
            $Snapshot = Get-WmiObject Win32_ShadowCopy | Where-Object { $_.ID -eq $ShadowID }
            return $Snapshot
        } else {
            Write-Error "Failed to create VSS Snapshot. Return Code: $($Result.ReturnValue)"
            throw "VSS Creation Failed with code $($Result.ReturnValue)"
        }
    } catch {
        Write-Error "Error creating VSS Snapshot: $($_.Exception.Message)"
        throw
    }
}

Function Remove-ShadowCopy {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [string]$ShadowID
    )

    Write-Host "Removing VSS Snapshot ID: $ShadowID" -ForegroundColor Cyan
    try {
        $Snapshot = Get-WmiObject Win32_ShadowCopy | Where-Object { $_.ID -eq $ShadowID }
        if ($Snapshot) {
            $Snapshot.Delete() # Delete() method doesn't return a value in typical WMI usage for this class via PS
            Write-Host "VSS Snapshot removed successfully." -ForegroundColor Green
        } else {
            Write-Warning "VSS Snapshot with ID $ShadowID not found."
        }
    } catch {
        Write-Error "Error removing VSS Snapshot: $($_.Exception.Message)"
        # We don't throw here to ensure backup process cleanup continues if possible
    }
}

#endregion Functions - VSS Management

#region Functions - Backup Operations

Function Invoke-RobocopyBackup {
    [CmdletBinding()]
    Param (
        [string]$Source,
        [string]$Destination,
        [string]$Options,
        [int]$InterPacketGapMs,
        [string]$LogFile
    )

    # Prepare options
    $FinalOptions = $Options.Replace("{logpath}", "`"$LogFile`"")
    
    if ($InterPacketGapMs -gt 0) {
        $FinalOptions += " /IPG:$InterPacketGapMs"
    }

    $CommandArgs = @($Source, $Destination) + $FinalOptions.Split(" ") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    Write-Host "Executing Robocopy..." -ForegroundColor Cyan
    Write-Host "Source: $Source" -ForegroundColor DarkGray
    Write-Host "Dest:   $Destination" -ForegroundColor DarkGray
    Write-Host "Log:    $LogFile" -ForegroundColor DarkGray

    # Execute Robocopy
    # We use Start-Process to handle exit codes correctly without throwing exceptions for codes < 8
    $Process = Start-Process -FilePath "robocopy.exe" -ArgumentList $CommandArgs -Wait -NoNewWindow -PassThru
    
    $ExitCode = $Process.ExitCode
    $Status = "Unknown"
    
    if ($ExitCode -lt 8) {
        $Status = "Success"
        Write-Host "Robocopy finished successfully (Code $ExitCode)." -ForegroundColor Green
    } else {
        $Status = "Failed"
        Write-Error "Robocopy failed with exit code $ExitCode."
    }

    # Parse log file for stats (simplified)
    $TotalFiles = 0
    $CopiedFiles = 0
    $TotalBytes = 0
    $CopiedBytes = 0
    
    if (Test-Path $LogFile) {
        $LogContent = Get-Content $LogFile
        # Regex to find the summary table: "   Files :      123         0 ..."
        # This is a bit brittle, fallback to 0 if parsing fails
        # A more robust way is parsing the footer lines
    }

    return @{
        Status = $Status
        ExitCode = $ExitCode
        LogFile = $LogFile
    }
}

Function Write-BackupHistory {
    Param (
        [string]$LogFilePath,
        [object]$Entry
    )
    $JsonLine = $Entry | ConvertTo-Json -Compress -Depth 10
    Add-Content -Path $LogFilePath -Value $JsonLine
}

Function Clean-OldBackups {
    Param (
        [string]$DestinationRoot,
        [int]$RetentionDays,
        [string]$LogsDir,
        [int]$LogRetentionDays
    )

    Write-Host "Starting cleanup..." -ForegroundColor Cyan
    $CutoffDate = (Get-Date).AddDays(-$RetentionDays)
    $LogCutoffDate = (Get-Date).AddDays(-$LogRetentionDays)

    # 1. Clean Backup Folders
    # Assuming structure: DestinationRoot\ShareName\SubDir_TIMESTAMP
    # actually structure is DestinationPath\SubDirName_YYYYMMDD_HHMMSS
    # We need to be careful only to delete folders matching the pattern and older than cutoff
    
    if (Test-Path $DestinationRoot) {
        Get-ChildItem -Path $DestinationRoot -Directory -Recurse | Where-Object {
            $_.Name -match "_\d{8}_\d{6}$"
        } | ForEach-Object {
            if ($_.CreationTime -lt $CutoffDate) {
                Write-Host "Deleting old backup: $($_.FullName)" -ForegroundColor Yellow
                Remove-Item -Path $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }

    # 2. Clean Log Files
    if (Test-Path $LogsDir) {
        Get-ChildItem -Path $LogsDir -Filter "*.log" | Where-Object {
            $_.LastWriteTime -lt $LogCutoffDate
        } | ForEach-Object {
            Write-Host "Deleting old log: $($_.FullName)" -ForegroundColor Yellow
            Remove-Item -Path $_.FullName -Force -ErrorAction SilentlyContinue
        }
    }
}

#endregion Functions - Backup Operations

#region Main Script Execution

Function Get-BackupItems {
    Param (
        [string]$SourcePath,
        [string]$VssSourceRoot,
        [string]$BackupMode
    )
    
    $ItemsToBackup = @()
    
    if ($BackupMode -eq "Root") {
        $FolderName = Split-Path $SourcePath -Leaf
        $ItemsToBackup += [PSCustomObject]@{
            Name = $FolderName
            SourceSubPath = $VssSourceRoot
            IsRootMode = $true
        }
    } else {
        $SubDirectories = Get-ChildItem -Path $SourcePath -Directory
        foreach ($SubDir in $SubDirectories) {
            $ItemsToBackup += [PSCustomObject]@{
                Name = $SubDir.Name
                SourceSubPath = "$VssSourceRoot\$($SubDir.Name)"
                IsRootMode = $false
            }
        }
    }
    return ,$ItemsToBackup
}

Function Execute-BackupItem {
    Param (
        $Item,
        $SourcePath,
        $BackupMode,
        $DestinationRoot,
        $RobocopyOptions,
        $InterPacketGapMs,
        $LogsDir,
        $HistoryLogFilePath,
        $SnapshotID
    )

    $SubDirName = $Item.Name
    $SourceSubDirPath = $Item.SourceSubPath
    $Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $SourceFolderName = Split-Path $SourcePath -Leaf
    
    $TargetParentDir = Join-Path $DestinationRoot $SourceFolderName
    if (-not (Test-Path $TargetParentDir)) { New-Item -Path $TargetParentDir -ItemType Directory -Force | Out-Null }
    
    $TargetDirName = "${SubDirName}_${Timestamp}"
    $TargetFullPath = Join-Path $TargetParentDir $TargetDirName
    
    $LogFileName = "${Timestamp}_${SourceFolderName}_${SubDirName}.log"
    $DetailLogPath = Join-Path $LogsDir $LogFileName
    
    Write-Host "  Backing up: $SubDirName" -ForegroundColor White
    
    $BackupResult = Invoke-RobocopyBackup `
        -Source $SourceSubDirPath `
        -Destination $TargetFullPath `
        -Options $RobocopyOptions `
        -InterPacketGapMs $InterPacketGapMs `
        -LogFile $DetailLogPath
    
    $HistoryEntry = @{
        Timestamp = $Timestamp
        SourcePath = $SourcePath
        Subdirectory = $SubDirName
        Mode = $BackupMode
        DestinationPath = $TargetFullPath
        Status = $BackupResult.Status
        ExitCode = $BackupResult.ExitCode
        ShadowCopyId = if ($SnapshotID) { $SnapshotID } else { "N/A" }
        DetailLogFile = $DetailLogPath
    }
    Write-BackupHistory -LogFilePath $HistoryLogFilePath -Entry $HistoryEntry
}

Function Start-BackupProcess {
    Param (
        [string]$ConfigFilePath
    )

    try {
        # Ensure logs directory exists
        if (-not (Test-Path $LogsDir -PathType Container)) {
            Write-Host "Creating logs directory: $LogsDir" -ForegroundColor Yellow
            New-Item -Path $LogsDir -ItemType Directory -Force | Out-Null
        }

        $Config = Get-Configuration -ConfigPath $ConfigFilePath
        $HistoryLogFilePath = Join-Path $ScriptDir $Config.HistoryLogFile

        if (-not (Test-Path $HistoryLogFilePath -PathType Leaf)) {
            Write-Host "Creating history log file: $HistoryLogFilePath" -ForegroundColor Yellow
            New-Item -Path $HistoryLogFilePath -ItemType File -Force | Out-Null
        }

        Write-Host "Configuration loaded successfully." -ForegroundColor Green
        
        # --- Main Backup Loop ---
        foreach ($SourceItem in $Config.SourcePaths) {
            # Normalize Source Item (String vs Object)
            $SourcePath = $null
            $BackupMode = "SubDirectories" # Default

            if ($SourceItem -is [string]) {
                $SourcePath = $SourceItem
            } elseif ($null -ne $SourceItem.Path) {
                $SourcePath = $SourceItem.Path
                if ($null -ne $SourceItem.Mode) { $BackupMode = $SourceItem.Mode }
            } else {
                Write-Warning "Skipping invalid source item format."
                continue
            }

            Write-Host "`nProcessing Source: $SourcePath (Mode: $BackupMode)" -ForegroundColor Cyan
            
            if (-not (Test-Path $SourcePath)) {
                Write-Warning "Source path '$SourcePath' does not exist. Skipping."
                continue
            }

            $VolumeRoot = Get-VolumeRoot -Path $SourcePath
            $RelativePath = $SourcePath.Substring($VolumeRoot.Length) # Remove Drive Letter (e.g., "C:\")
            
            # 1. Create VSS Snapshot or Use Direct Path
            $Snapshot = $null
            $VssSourceRoot = $null
            
            if ($Config.UseVSS) {
                try {
                    $Snapshot = New-ShadowCopy -VolumeRoot $VolumeRoot
                } catch {
                    Write-Error "Skipping $SourcePath due to VSS creation failure."
                    continue
                }

                if ($Snapshot) {
                    $ShadowDevicePath = $Snapshot.DeviceObject
                    # Construct VSS path for the source directory
                    $VssSourceRoot = "$ShadowDevicePath\$RelativePath" 
                }
            } else {
                Write-Warning "VSS is disabled in configuration. Using direct source path (no snapshot consistency)."
                $VssSourceRoot = $SourcePath
            }

            if ($VssSourceRoot) {
                $ItemsToBackup = Get-BackupItems `
                    -SourcePath $SourcePath `
                    -VssSourceRoot $VssSourceRoot `
                    -BackupMode $BackupMode

                foreach ($Item in $ItemsToBackup) {
                    Execute-BackupItem `
                        -Item $Item `
                        -SourcePath $SourcePath `
                        -BackupMode $BackupMode `
                        -DestinationRoot $Config.DestinationPath `
                        -RobocopyOptions $Config.RobocopyOptions `
                        -InterPacketGapMs $Config.RobocopyInterPacketGapMs `
                        -LogsDir $LogsDir `
                        -HistoryLogFilePath $HistoryLogFilePath `
                        -SnapshotID ($Snapshot.ID)
                }

                # 5. Remove VSS Snapshot
                if ($Snapshot) {
                    Remove-ShadowCopy -ShadowID $Snapshot.ID
                }
            }
        }

        # 6. Cleanup
        Clean-OldBackups `
            -DestinationRoot $Config.DestinationPath `
            -RetentionDays $Config.RetentionDays `
            -LogsDir $LogsDir `
            -LogRetentionDays $Config.LogRetentionDays

        Write-Host "`nBackup Script Completed." -ForegroundColor Green
    }
    catch {
        Write-Error "An unhandled error occurred during script execution: $($Error[0].Exception.Message)"
        exit 1
    }
}

# Only execute if script is run directly (not dot-sourced)
if ($MyInvocation.InvocationName -ne '.') {
    Start-BackupProcess -ConfigFilePath $ConfigFilePath
}

#endregion Main Script Execution
