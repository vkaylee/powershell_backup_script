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
    Path to the JSON configuration file. Defaults to 'config.jsonc' in the script directory.

.PARAMETER CheckOnly
    If specified, runs only the pre-flight diagnostic checks and exits. Useful for verifying environment readiness.

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
    [Parameter(Mandatory = $false)]
    [string]$ConfigFilePath,

    [Parameter(Mandatory = $false)]
    [switch]$CheckOnly
)

#region Global Variables and Constants
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
# Resolve absolute path for config file if it's relative
if ($ConfigFilePath -and -not (Test-Path $ConfigFilePath -PathType Leaf)) {
    # Check if it exists relative to script dir
    $PotentialPath = Join-Path $ScriptDir $ConfigFilePath
    if (Test-Path $PotentialPath -PathType Leaf) {
        $ConfigFilePath = $PotentialPath
    }
}
$LogsDir = Join-Path $ScriptDir "logs"
$Config = $null # Will store loaded configuration
$HistoryLogFilePath = $null # Will be set after config is loaded
#endregion Global Variables and Constants

#region Functions - Configuration Management

Function Show-Usage {
    Write-Host @"
============================================================
Snapshot Backup Script - Usage Guide
============================================================
This script performs consistent backups using VSS and Robocopy.

PARAMETERS:
  -ConfigFilePath <path>  [REQUIRED] Path to the JSONC configuration file.
  
  -CheckOnly              Runs only pre-flight diagnostic checks.
                          Does NOT perform any backup operations.
                          (Also requires -ConfigFilePath to validate config)

EXAMPLES:
  1. Run diagnostic check:
     .\backup-script.ps1 -ConfigFilePath config.jsonc -CheckOnly

  2. Run backup:
     .\backup-script.ps1 -ConfigFilePath config.jsonc

NOTES:
  - Requires Administrator privileges for VSS snapshots.
  - Supports .jsonc (JSON with comments) for configuration.

For full documentation, see: .\docs\user-guide.md
============================================================
"@ -ForegroundColor Cyan
}

Function Test-BackupPrerequisites {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [object]$Config
    )

    $Passed = $true
    Write-Host "Running Pre-flight Diagnostic Checks..." -ForegroundColor Cyan

    # 1. Check for Robocopy
    if (-not (Get-Command "robocopy.exe" -ErrorAction SilentlyContinue)) {
        Write-Error "CRITICAL: 'robocopy.exe' was not found in the system PATH. This tool is required for backup operations."
        $Passed = $false
    }
    else {
        Write-Host "  [OK] Robocopy detected." -ForegroundColor Green
    }

    # 2. Check for Administrator Privileges (Required for VSS)
    $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    
    if ($Config.UseVSS -and -not $IsAdmin) {
        Write-Error "CRITICAL: Administrator privileges are required when 'UseVSS' is enabled. Please run PowerShell as Administrator."
        $Passed = $false
    }
    elseif ($Config.UseVSS) {
        Write-Host "  [OK] Administrator privileges confirmed for VSS." -ForegroundColor Green
    }
    else {
        Write-Host "  [SKIP] VSS is disabled; skipping Administrator check." -ForegroundColor Gray
    }

    # 3. Basic Path Validation
    if (-not (Test-Path $Config.DestinationPath)) {
        Write-Warning "DestinationPath '$($Config.DestinationPath)' is currently inaccessible. The script will attempt to create it later, but please verify permissions."
    }
    else {
        Write-Host "  [OK] Destination path is accessible." -ForegroundColor Green
    }

    return $Passed
}

Function Get-Configuration {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )

    Write-Host "Attempting to load configuration from: $ConfigPath" -ForegroundColor Cyan

    $DefaultConfig = @{
        SourcePaths              = @("C:\BackupSources\TestShare1", "C:\BackupSources\TestShare2")
        DestinationPath          = (Join-Path $ScriptDir "Backups")
        RetentionDays            = 30
        LogRetentionDays         = 90
        MaxBackupAttempts        = 3
        RobocopyOptions          = '/MIR /MT:8 /R:1 /W:1 /J /XD RECYCLE.BIN "System Volume Information" /XF Thumbs.db /NP /LOG+:{logpath}'
        RobocopyInterPacketGapMs = 0
        HistoryLogFile           = "backup-history.log"
        UseVSS                   = $true
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
                }
                catch {
                    throw "Failed to create destination path '$($LoadedConfig.DestinationPath)'. Error: $($_.Exception.Message). Please ensure it's accessible and you have permissions."
                }
            }
        }
        else {
            throw "Configuration Error: DestinationPath is null or empty. Please provide a valid path."
        }
        return $LoadedConfig
    }
    catch {
        Write-Error "Failed to load or parse configuration file '$ConfigPath'. Error: $($_.Exception.Message). Using default configuration."
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
        [Parameter(Mandatory = $true)]
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
        }
        else {
            Write-Error "Failed to create VSS Snapshot. Return Code: $($Result.ReturnValue)"
            throw "VSS Creation Failed with code $($Result.ReturnValue)"
        }
    }
    catch {
        Write-Error "Error creating VSS Snapshot: $($_.Exception.Message)"
        throw
    }
}

Function Remove-ShadowCopy {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [string]$ShadowID
    )

    Write-Host "Removing VSS Snapshot ID: $ShadowID" -ForegroundColor Cyan
    try {
        $Snapshot = Get-WmiObject Win32_ShadowCopy | Where-Object { $_.ID -eq $ShadowID }
        if ($Snapshot) {
            $Snapshot.Delete() # Delete() method doesn't return a value in typical WMI usage for this class via PS
            Write-Host "VSS Snapshot removed successfully." -ForegroundColor Green
        }
        else {
            Write-Warning "VSS Snapshot with ID $ShadowID not found."
        }
    }
    catch {
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

    # 1. Ensure the directory for the log file exists
    $LogDir = Split-Path $LogFile -Parent
    if (-not (Test-Path $LogDir)) {
        New-Item -Path $LogDir -ItemType Directory -Force | Out-Null
    }

    # 2. Build Argument List
    if ([string]::IsNullOrWhiteSpace($Source) -or [string]::IsNullOrWhiteSpace($Destination)) {
        Write-Error "CRITICAL: Robocopy invoked with missing Source or Destination path."
        return @{ Status = "Failed"; ExitCode = 16; LogFile = $LogFile }
    }

    # Robocopy is extremely sensitive to trailing slashes in quoted strings.
    # We remove them to prevent escaping the closing quote.
    $SourceArg = $Source.TrimEnd('\')
    $DestArg = $Destination.TrimEnd('\')

    # 3. Process Options
    # We must exclude options that the script handles internally to prevent conflicts.
    # Protected: /LOG, /LOG+, /R, /W, /MT
    $ProtectedRegex = "^/(LOG|LOG\+|R|W|MT)(:.*)?$"
    $CleanOptions = $Options -replace '/LOG\+?:\S+', '' -replace '\{logpath\}', ''
    
    # Split options into an array, respecting quotes, and filter out protected ones
    # We do NOT trim quotes here because they might be needed for paths with spaces.
    $OptionArray = [regex]::Matches($CleanOptions, '(?:[^\s"]+|"[^"]*")+') | 
        ForEach-Object { $_.Value } |
        Where-Object { $_ -notmatch $ProtectedRegex }

    # 4. Execute using Call Operator (&)
    Write-Host "Executing Robocopy..." -ForegroundColor Cyan
    Write-Host "Source: $SourceArg" -ForegroundColor DarkGray
    Write-Host "Dest:   $DestArg" -ForegroundColor DarkGray

    # Construct the final argument list carefully for Win32 compatibility
    $RobocopyParams = @($SourceArg, $DestArg)
    foreach ($Opt in $OptionArray) { if ($Opt) { $RobocopyParams += $Opt } }
    
    # Add mandatory log, throttle, multi-threading, and fail-fast retries
    $RobocopyParams += "/LOG+:$LogFile"
    $RobocopyParams += "/R:1"
    $RobocopyParams += "/W:1"
    $RobocopyParams += "/MT:8"
    if ($InterPacketGapMs -gt 0) { $RobocopyParams += "/IPG:$InterPacketGapMs" }

    # Execute
    $LASTEXITCODE = 0
    & robocopy.exe $RobocopyParams
    $ExitCode = $LASTEXITCODE
    
    $Status = "Unknown"
    # Success codes are < 8
    if ($ExitCode -lt 8) {
        $Status = "Success"
        Write-Host "Robocopy finished successfully (Code $ExitCode)." -ForegroundColor Green
    }
    else {
        $Status = "Failed"
        Write-Error "Robocopy failed with exit code $ExitCode."
    }

    return @{
        Status   = $Status
        ExitCode = $ExitCode
        LogFile  = $LogFile
    }
}

Function Write-BackupHistory {
    Param (
        [string]$LogFilePath,
        [object]$Entry
    )
    $JsonLine = $Entry | ConvertTo-Json -Compress -Depth 10
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogLine = "[$Timestamp] $JsonLine"
    Add-Content -Path $LogFilePath -Value $LogLine
}

Function Remove-OldBackups {
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
            $_.Name -match "_\d{8}_\d{6}(_\d{3})?$"
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
        # Split-Path -Leaf returns empty if path ends with \. Trim it.
        $FolderName = Split-Path ($SourcePath.TrimEnd('\')) -Leaf
        $ItemsToBackup += [PSCustomObject]@{
            Name          = $FolderName
            SourceSubPath = $VssSourceRoot
            IsRootMode    = $true
        }
    }
    else {
        $SubDirectories = Get-ChildItem -Path $SourcePath -Directory
        foreach ($SubDir in $SubDirectories) {
            $ItemsToBackup += [PSCustomObject]@{
                Name          = $SubDir.Name
                SourceSubPath = Join-Path $VssSourceRoot $SubDir.Name
                IsRootMode    = $false
            }
        }
    }
    # Write-Error "[DEBUG] Get-BackupItems: Returning $($ItemsToBackup.Count) items"
    # Use unary comma operator to prevent array unrolling when returning single item
    return , $ItemsToBackup
}

Function Invoke-BackupItem {
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
    $SourceSubPath = $Item.SourceSubPath
    # Use high-precision timestamp to ensure unique folders for rapidly succeeding items
    $Timestamp = Get-Date -Format "yyyyMMdd_HHmmss_fff"
    $SourceFolderName = Split-Path $SourcePath -Leaf
    
    $TargetParentDir = Join-Path $DestinationRoot $SourceFolderName
    if (-not (Test-Path $TargetParentDir)) { New-Item -Path $TargetParentDir -ItemType Directory -Force | Out-Null }
    
    $TargetDirName = "${SubDirName}_${Timestamp}"
    $TargetFullPath = Join-Path $TargetParentDir $TargetDirName
    
    # Ensure LogsDir exists before using it (extra safety)
    if (-not (Test-Path $LogsDir)) { New-Item -Path $LogsDir -ItemType Directory -Force | Out-Null }
    
    $LogFileName = "${Timestamp}_${SourceFolderName}_${SubDirName}.log"
    $DetailLogPath = Join-Path $LogsDir $LogFileName
    
    Write-Host "  Backing up: $SubDirName" -ForegroundColor White
    
    $BackupResult = Invoke-RobocopyBackup `
        -Source $SourceSubPath `
        -Destination $TargetFullPath `
        -Options $RobocopyOptions `
        -InterPacketGapMs $InterPacketGapMs `
        -LogFile $DetailLogPath
    
    $HistoryEntry = @{
        Timestamp       = $Timestamp
        SourcePath      = $SourceSubPath # Log the actual VSS path used
        Subdirectory    = $SubDirName
        Mode            = $BackupMode
        DestinationPath = $TargetFullPath
        Status          = $BackupResult.Status
        ExitCode        = $BackupResult.ExitCode
        ShadowCopyId    = if ($SnapshotID) { $SnapshotID } else { "N/A" }
        DetailLogFile   = $DetailLogPath
    }
    Write-BackupHistory -LogFilePath $HistoryLogFilePath -Entry $HistoryEntry
}

Function Start-BackupProcess {
    Param (
        [string]$ConfigFilePath,
        [switch]$CheckOnly
    )

    $ScriptStartTime = Get-Date
    Write-Host "Backup process started at $($ScriptStartTime.ToString('HH:mm:ss.fff'))" -ForegroundColor Gray

    try {
        # Ensure logs directory exists (skip if CheckOnly, we might not need it yet, but good to check permissions)
        if (-not (Test-Path $LogsDir -PathType Container)) {
            Write-Host "Creating logs directory: $LogsDir" -ForegroundColor Yellow
            New-Item -Path $LogsDir -ItemType Directory -Force | Out-Null
        }

        $Config = Get-Configuration -ConfigPath $ConfigFilePath
        $HistoryLogFilePath = Join-Path $ScriptDir $Config.HistoryLogFile

        # --- Pre-flight Checks ---
        $CheckResult = Test-BackupPrerequisites -Config $Config
        
        if ($CheckOnly) {
            if ($CheckResult) {
                Write-Host "`n[SUCCESS] System is ready for backup." -ForegroundColor Green
            }
            else {
                Write-Host "`n[FAILED] System is NOT ready for backup. Please fix the errors above." -ForegroundColor Red
            }
            return # Exit successfully (diagnostic mode)
        }

        if (-not $CheckResult) {
            Write-Error "Pre-flight checks failed. Aborting backup process to prevent inconsistent data."
            return # Exit function
        }

        if (-not (Test-Path $HistoryLogFilePath -PathType Leaf)) {
            Write-Host "Creating history log file: $HistoryLogFilePath" -ForegroundColor Yellow
            New-Item -Path $HistoryLogFilePath -ItemType File -Force | Out-Null
        }

        Write-Host "Configuration loaded successfully." -ForegroundColor Green
        
        # --- Main Backup Loop ---
        foreach ($SourceItem in $Config.SourcePaths) {
            $LoopStartTime = Get-Date
            Write-Host "Processing Source Item at $($LoopStartTime.ToString('HH:mm:ss.fff'))" -ForegroundColor Gray
            # Normalize Source Item (String vs Object)
            $SourcePath = $null
            $BackupMode = "SubDirectories" # Default

            if ($SourceItem -is [string]) {
                $SourcePath = $SourceItem
            }
            elseif ($null -ne $SourceItem.Path) {
                $SourcePath = $SourceItem.Path
                if ($null -ne $SourceItem.Mode) {
                    $ValidModes = @("SubDirectories", "Root")
                    if ($SourceItem.Mode -notin $ValidModes) {
                        Write-Warning "Invalid Mode '$($SourceItem.Mode)' for source '$($SourceItem.Path)'. Valid modes: $($ValidModes -join ', '). Defaulting to 'SubDirectories'."
                    }
                    else {
                        $BackupMode = $SourceItem.Mode
                    }
                }
            }
            else {
                Write-Warning "Skipping invalid source item format."
                continue
            }

            # 1. Resolve to Absolute Path immediately
            $ResolvedPath = $null
            try {
                if (Test-Path $SourcePath) {
                    $ResolvedPath = (Resolve-Path $SourcePath).Path
                }
            }
            catch {
                Write-Verbose "Path resolution failed for '$SourcePath': $($_.Exception.Message)"
            }

            if (-not $ResolvedPath) {
                Write-Warning "Source path '$SourcePath' does not exist or is inaccessible. Skipping."
                continue
            }
            $SourcePath = $ResolvedPath

            Write-Host "`nProcessing Source: $SourcePath (Mode: $BackupMode)" -ForegroundColor Cyan
            
            $VolumeRoot = Get-VolumeRoot -Path $SourcePath
            # Calculate relative path from volume root. Handle casing and trailing slashes.
            $RelativePath = ""
            if ($SourcePath.Length -gt $VolumeRoot.Length) {
                $RelativePath = $SourcePath.Substring($VolumeRoot.Length).TrimStart('\')
            }
            
            # 2. Create VSS Snapshot or Use Direct Path
            $Snapshot = $null
            $VssSourceRoot = $null
            
            if ($Config.UseVSS) {
                Write-Host "Creating VSS Snapshot for volume: $VolumeRoot at $((Get-Date).ToString('HH:mm:ss.fff'))" -ForegroundColor Cyan
                try {
                    $Snapshot = New-ShadowCopy -VolumeRoot $VolumeRoot
                    $SnapshotID = $Snapshot.ID
                    Write-Host "VSS Snapshot created successfully at $((Get-Date).ToString('HH:mm:ss.fff'))" -ForegroundColor Green
                    # Wait for device object stability
                    Start-Sleep -Seconds 2
                }
                catch {
                    Write-Error "Skipping $SourcePath due to VSS creation failure."
                    continue
                }

                if ($Snapshot) {
                    $ShadowDevicePath = $Snapshot.DeviceObject
                    
                    # DIRECT \\?\GLOBALROOT paths often fail in Robocopy/PowerShell with Error 123.
                    # FIX: Create a temporary directory junction to the snapshot.
                    $VssJunctionPath = Join-Path $ScriptDir ("VssJunction_" + $SnapshotID.Replace("{", "").Replace("}", ""))
                    if (Test-Path $VssJunctionPath) { cmd.exe /c "rd /q `"$VssJunctionPath`"" }
                    
                    $JunctionTarget = $ShadowDevicePath.TrimEnd('\') + "\"
                    Write-Host "Creating VSS Junction: $VssJunctionPath -> $JunctionTarget" -ForegroundColor Gray
                    cmd.exe /c "mklink /j `"$VssJunctionPath`" `"$JunctionTarget`" 2>&1" | Out-Null
                    
                    if (-not (Test-Path $VssJunctionPath)) {
                        Write-Error "Failed to create VSS Junction. Robocopy might fail."
                        $VssSourceRoot = $ShadowDevicePath.TrimEnd('\')
                    }
                    else {
                        # Map the relative path within the junction
                        if ([string]::IsNullOrWhiteSpace($RelativePath)) {
                            $VssSourceRoot = $VssJunctionPath
                        }
                        else {
                            $VssSourceRoot = Join-Path $VssJunctionPath ($RelativePath.TrimStart('\'))
                        }
                    }
                }
            }
            else {
                Write-Warning "VSS is disabled in configuration. Using direct source path (no snapshot consistency)."
                $VssSourceRoot = $SourcePath
            }

            if ($VssSourceRoot) {
                Write-Host "Item discovery started at $((Get-Date).ToString('HH:mm:ss.fff'))" -ForegroundColor Gray
                $Items = Get-BackupItems -SourcePath $SourcePath -VssSourceRoot $VssSourceRoot -BackupMode $BackupMode
                Write-Host "Found $($Items.Count) items to backup." -ForegroundColor Gray

                foreach ($Item in $Items) {
                    Write-Host "  Executing backup for $($Item.Name) at $((Get-Date).ToString('HH:mm:ss.fff'))" -ForegroundColor Gray
                    Invoke-BackupItem `
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
            }

            # 5. Remove VSS Snapshot
            if ($VssJunctionPath -and (Test-Path $VssJunctionPath)) {
                Write-Host "Removing VSS Junction: $VssJunctionPath" -ForegroundColor Gray
                cmd.exe /c "rd /q `"$VssJunctionPath`""
            }

            if ($Snapshot) {
                Remove-ShadowCopy -ShadowID $Snapshot.ID
            }
        }

        # 6. Cleanup
        Remove-OldBackups `
            -DestinationRoot $Config.DestinationPath `
            -RetentionDays $Config.RetentionDays `
            -LogsDir $LogsDir `
            -LogRetentionDays $Config.LogRetentionDays

    }
    catch {
        Write-Error "CRITICAL ERROR: $($_.Exception.Message)"
    }
    finally {
        $EndTime = Get-Date
        $Duration = $EndTime - $ScriptStartTime
        Write-Host "`nBackup Script Completed at $($EndTime.ToString('HH:mm:ss.fff'))" -ForegroundColor Green
        Write-Host "Total Duration: $($Duration.TotalSeconds.ToString('F2')) seconds" -ForegroundColor Gray
    }
}

# Only execute if script is run directly (not dot-sourced)
if ($MyInvocation.InvocationName -ne '.') {
    if ([string]::IsNullOrWhiteSpace($ConfigFilePath)) {
        Show-Usage
        exit 1 # Exit with error code if required arg is missing
    }
    
    Start-BackupProcess -ConfigFilePath $ConfigFilePath -CheckOnly:$CheckOnly
}

#endregion Main Script Execution
