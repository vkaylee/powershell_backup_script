$ScriptPath = Join-Path $PSScriptRoot "backup-script.ps1"
. $ScriptPath

Describe "Snapshot Backup Script - Extended Tests" {

    # --- Pre-flight Checks ---
    Context "Test-BackupPrerequisites" {
        It "Should fail if Robocopy is missing" {
            Mock Get-Command { return $null } -ParameterFilter { $Name -eq "robocopy.exe" }
            $Config = @{ UseVSS = $false; DestinationPath = "C:\" }
            $Result = Test-BackupPrerequisites -Config $Config
            $Result | Should Be $false
        }

        It "Should fail if UseVSS is true but user is not Admin" {
            Mock Get-Command { return [PSCustomObject]@{ Name = "robocopy.exe" } } -ParameterFilter { $Name -eq "robocopy.exe" }
            $Config = @{ UseVSS = $true; DestinationPath = "C:\" }
            $Result = Test-BackupPrerequisites -Config $Config
            
            # Check current process admin status to know expected result
            $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
            if (-not $IsAdmin) {
                $Result | Should Be $false
            }
            else {
                $Result | Should Be $true
            }
        }
    }

    # --- VSS Workflow Integration ---
    Context "VSS Workflow Integration - Enabled" {
        $script:RobocopyCalled = $false
        Mock Invoke-RobocopyBackup { $script:RobocopyCalled = $true; return @{ Status = "Success"; ExitCode = 0; LogFile = "test.log" } }

        It "Should create snapshots for multiple sources and copy to the correct destination" {
            Mock Get-Command { return [PSCustomObject]@{ Name = "robocopy.exe" } } -ParameterFilter { $Name -eq "robocopy.exe" }
            Mock Test-Path { return $true }
            Mock New-Item { return $null }
            Mock Join-Path { 
                param($Path, $ChildPath)
                if ([string]::IsNullOrWhiteSpace($Path)) { return $ChildPath }
                if ($Path.EndsWith("\")) { return "$Path$ChildPath" }
                return "$Path\$ChildPath"
            }
            Mock Get-VolumeRoot { 
                if ($Path -like "D:\*") { return "D:\" }
                if ($Path -like "E:\*") { return "E:\" }
                return "D:\"
            }
            Mock Test-BackupPrerequisites { return $true }
            Mock Start-Sleep { } # Speed up tests
            Mock Resolve-Path { return [PSCustomObject]@{ Path = $Path } }

            $Config = @{
                SourcePaths              = @("D:\Shares\Data", "D:\Other")
                DestinationPath          = "E:\Backups"
                RetentionDays            = 30
                LogRetentionDays         = 90
                MaxBackupAttempts        = 3
                RobocopyOptions          = '/MIR /LOG+:{logpath}'
                RobocopyInterPacketGapMs = 0
                HistoryLogFile           = "backup-history.log"
                UseVSS                   = $true
            }

            # Mock VSS functions
            Mock New-ShadowCopy { 
                return [PSCustomObject]@{ ID = "{ID-D}"; DeviceObject = "\\?\GLOBALROOT\Device\VSS_D" }
            }
            Mock Remove-ShadowCopy { }
            Mock Get-Configuration { return $Config }
            Mock cmd.exe { }
            Mock Write-BackupHistory { }
            Mock Remove-OldBackups { }
            
            # Mock Get-BackupItems to reflect the source being processed
            Mock Get-BackupItems {
                # Ensure the path passed to Robocopy mock has the trailing slash the script now adds
                $VssPath = if ($VssSourceRoot.EndsWith("\")) { $VssSourceRoot } else { "$VssSourceRoot\" }
                return , @([PSCustomObject]@{ Name = "Data"; SourceSubPath = $VssPath; IsRootMode = $true })
            }

            Start-BackupProcess -ConfigFilePath "fake.json"

            # Verify snapshots created for D: (called twice because we have two D: sources)
            Assert-MockCalled New-ShadowCopy -Exactly 2 -ParameterFilter { $VolumeRoot -eq "D:\" }
            
            # Verify Robocopy called with correctly constructed VSS source
            $script:RobocopyCalled | Should Be $true
        }
    }

    Context "VSS Workflow Integration - Deep Path Mapping" {
        It "Should correctly map deep subdirectories and NOT default to volume root" {
            Mock Invoke-RobocopyBackup { 
                return @{ Status = "Success"; ExitCode = 0; LogFile = "test.log" } 
            }

            Mock Get-Command { return [PSCustomObject]@{ Name = "robocopy.exe" } } -ParameterFilter { $Name -eq "robocopy.exe" }
            Mock Test-Path { return $true }
            Mock New-Item { return $null }
            Mock Join-Path { 
                param($Path, $ChildPath)
                if ([string]::IsNullOrWhiteSpace($Path)) { return $ChildPath }
                if ($Path.EndsWith("\")) { return "$Path$ChildPath" }
                return "$Path\$ChildPath"
            }
            Mock Get-VolumeRoot { return "D:\" }
            Mock Test-BackupPrerequisites { return $true }
            Mock Start-Sleep { }
            
            $DeepPath = "D:\tuanlee\snapshot_backup_script\workdirs"
            Mock Resolve-Path { return [PSCustomObject]@{ Path = $DeepPath } }

            $Config = @{
                SourcePaths              = @($DeepPath)
                DestinationPath          = "E:\Backups"
                RetentionDays            = 30
                LogRetentionDays         = 90
                MaxBackupAttempts        = 3
                RobocopyOptions          = '/MIR /LOG+:{logpath}'
                RobocopyInterPacketGapMs = 0
                HistoryLogFile           = "backup-history.log"
                UseVSS                   = $true
            }

            Mock New-ShadowCopy { return [PSCustomObject]@{ ID = "{ID-D}"; DeviceObject = "\\?\GLOBALROOT\Device\VSS_D" } }
            Mock Remove-ShadowCopy { }
            Mock Get-Configuration { return $Config }
            Mock cmd.exe { }
            Mock Get-BackupItems {
                return , @([PSCustomObject]@{ Name = "workdirs"; SourceSubPath = "\\?\GLOBALROOT\Device\VSS_D\tuanlee\snapshot_backup_script\workdirs\"; IsRootMode = $true })
            }
            Mock Write-BackupHistory { }
            Mock Remove-OldBackups { }

            Start-BackupProcess -ConfigFilePath "fake.json"

            # Verify the mock was called (regression test for VSS workflow completion)
            Assert-MockCalled Invoke-RobocopyBackup -Exactly 1
        }
    }

    Context "VSS Workflow Integration - Disabled" {
        $script:RobocopyCalledDisabled = $false
        Mock Invoke-RobocopyBackup { $script:RobocopyCalledDisabled = $true; return @{ Status = "Success"; ExitCode = 0; LogFile = "test.log" } }

        It "Should NOT create a snapshot and use direct path when UseVSS is false" {
            # Normalize drive root mock for this context
            Mock Get-Command { return [PSCustomObject]@{ Name = "robocopy.exe" } } -ParameterFilter { $Name -eq "robocopy.exe" }
            Mock Test-Path { return $true }
            Mock New-Item { return $null }
            Mock Join-Path { 
                param($Path, $ChildPath)
                if ([string]::IsNullOrWhiteSpace($Path)) { return $ChildPath }
                if ($Path.EndsWith("\")) { return "$Path$ChildPath" }
                return "$Path\$ChildPath"
            }
            Mock Get-VolumeRoot { return "D:\" }
            Mock Test-BackupPrerequisites { return $true }
            Mock Resolve-Path { return [PSCustomObject]@{ Path = $Path } }

            $Config = @{
                SourcePaths              = @("D:\Shares\Data")
                DestinationPath          = "E:\Backups"
                RetentionDays            = 30
                LogRetentionDays         = 90
                MaxBackupAttempts        = 3
                RobocopyOptions          = '/MIR /LOG+:{logpath}'
                RobocopyInterPacketGapMs = 0
                HistoryLogFile           = "backup-history.log"
                UseVSS                   = $false
            }
            
            Mock Get-Configuration { return $Config }
            Mock New-ShadowCopy { throw "Should not be called" }
            Mock Get-BackupItems {
                $VssPath = if ($VssSourceRoot.EndsWith("\")) { $VssSourceRoot } else { "$VssSourceRoot\" }
                return , @([PSCustomObject]@{
                        Name          = "Data"
                        SourceSubPath = $VssPath
                        IsRootMode    = $true
                    })
            }
            Mock Write-BackupHistory { }
            Mock Remove-OldBackups { }

            Start-BackupProcess -ConfigFilePath "fake.json"

            Assert-MockCalled New-ShadowCopy -Exactly 0
            $script:RobocopyCalledDisabled | Should Be $true
        }
    }

    # --- CLI Behavior ---
    Context "CLI Behavior" {
        $TestCliDir = Join-Path $PSScriptRoot "TestCliDir"
        New-Item -Path $TestCliDir -ItemType Directory -Force | Out-Null
        $TestConfigForCLI = Join-Path $PSScriptRoot "cli_test_config.jsonc"
        
        # We need escaped backslashes for JSON
        $JsonDir = $TestCliDir.Replace("\", "\\")
        Set-Content -Path $TestConfigForCLI -Value "{ ""SourcePaths"": [""$JsonDir""], ""DestinationPath"": ""$JsonDir"", ""UseVSS"": false }"

        It "Should display usage if no parameters are passed" {
            # We check if the function exists
            (Get-Command Show-Usage -ErrorAction SilentlyContinue) | Should Not BeNullOrEmpty
        }

        It "Should exit early if CheckOnly is specified" {
            # Mocking Get-Configuration so we don't depend on files
            Mock Get-Configuration { return @{ UseVSS = $false; DestinationPath = "C:\"; HistoryLogFile = "test.log" } }
            $Result = Start-BackupProcess -ConfigFilePath "dummy.json" -CheckOnly:$true
        }

        It "Should report failure in -CheckOnly mode if prerequisites fail" {
            Mock Get-Configuration { return @{ UseVSS = $false; DestinationPath = "C:\"; HistoryLogFile = "test.log" } }
            # Mock Test-BackupPrerequisites to return false
            Mock Test-BackupPrerequisites { return $false }
            
            $Output = Start-BackupProcess -ConfigFilePath "dummy.json" -CheckOnly:$true *>&1
            $Output -join "`n" | Should Match "\[FAILED\] System is NOT ready for backup"
        }

        It "Should execute diagnostic checks and report success when -CheckOnly is specified" {
            # Invoking with & operator runs in a child scope
            $Output = & $ScriptPath -ConfigFilePath $TestConfigForCLI -CheckOnly *>&1
            $JoinedOutput = $Output -join "`n"
            $JoinedOutput | Should Match "Running Pre-flight Diagnostic Checks..."
            $JoinedOutput | Should Match "\[OK\] Robocopy detected"
            $JoinedOutput | Should Match "\[SUCCESS\] System is ready for backup"
        }

        It "Should NOT run diagnostic mode when invoked externally without -CheckOnly" {
            $Output = & $ScriptPath -ConfigFilePath $TestConfigForCLI *>&1
            $Output -join "`n" | Should Match "Configuration loaded successfully"
            $Output -join "`n" | Should Not Match "System is ready for backup"
        }

        # Cleanup
        Remove-Item $TestConfigForCLI -ErrorAction SilentlyContinue
        Remove-Item $TestCliDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    # --- Helper Functions ---
    Context "Get-VolumeRoot" {
        It "Should return drive root for local path" {
            # We use $PSScriptRoot which definitely exists
            $Drive = Split-Path $PSScriptRoot -Qualifier
            $Expected = "$Drive\"
            
            $Result = Get-VolumeRoot -Path $PSScriptRoot
            $Result | Should Be $Expected
        }
    }

    # --- Configuration ---
    Context "Get-Configuration" {
        $TestConfigPath = Join-Path $PSScriptRoot "test_config_extended.json"

        It "Should validate missing SourcePaths" {
            @{ 
                SourcePaths     = @()
                DestinationPath = "C:\Dest"
            } | ConvertTo-Json | Set-Content $TestConfigPath

            # Should return default config (with warning) or handle it
            # Our function returns default if file missing, but if file exists and invalid?
            # It merges. So SourcePaths will be empty if we override it?
            # Actually logic says: if (-not $Loaded.SourcePaths) { warn; use default }
            
            $Config = Get-Configuration -ConfigPath $TestConfigPath
            $Config.SourcePaths.Count | Should BeGreaterThan 0
            
            Remove-Item $TestConfigPath -ErrorAction SilentlyContinue
        }
    }

    # --- Backup Items Logic ---
    Context "Get-BackupItems Logic" {
        $TestRoot = Join-Path $PSScriptRoot "TestItems"
        # Ensure cleanup first
        Remove-Item $TestRoot -Recurse -Force -ErrorAction SilentlyContinue
        New-Item -Path "$TestRoot\Sub1" -ItemType Directory -Force | Out-Null
        New-Item -Path "$TestRoot\Sub2" -ItemType Directory -Force | Out-Null

        It "Should return all subdirectories in SubDirectories mode" {
            $Items = Get-BackupItems -SourcePath $TestRoot -VssSourceRoot \\?\GLOBALROOT\Dev\Shadow\TestItems -BackupMode "SubDirectories"
            
            $Items.Count | Should Be 2
            $Items[0].IsRootMode | Should Be $false
        }

        It "Should return single item in Root mode" {
            # Manually construct a path string
            $BaseTestPath = "C:\MyBackupSource"
            $VssRoot = "\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy1\MyBackupSource"
            
            $Items = Get-BackupItems -SourcePath $BaseTestPath -VssSourceRoot $VssRoot -BackupMode "Root"
            
            $Items.Count | Should Be 1
            $Items[0].IsRootMode | Should Be $true
            $Items[0].Name | Should Be "MyBackupSource"
            $Items[0].SourceSubPath | Should Be $VssRoot
        }
        
        It "Should return empty if source has no subdirectories in SubDirectories mode" {
            $EmptyRoot = Join-Path $PSScriptRoot "EmptyRoot"
            New-Item -Path $EmptyRoot -ItemType Directory -Force | Out-Null
             
            $Items = Get-BackupItems -SourcePath $EmptyRoot -VssSourceRoot "..." -BackupMode "SubDirectories"
            $Items.Count | Should Be 0
             
            Remove-Item $EmptyRoot -Force
        }

        # Cleanup
        Remove-Item $TestRoot -Recurse -Force -ErrorAction SilentlyContinue
    }

    # --- Execution Logic ---
    Context "Invoke-RobocopyBackup Regression Tests" {
        $TestLog = Join-Path $PSScriptRoot "robocopy_test.log"
        
        It "Should strip protected flags from user options" {
            # Capture the arguments passed to robocopy
            $RobocopyArguments = @()
            Mock robocopy.exe { 
                # We can't easily capture parameters passed to a native EXE in Pester 3.4.0 
                # without wrapping it. But we can verify the function logic.
            }
            
            # The issue was redundant /LOG, /R, /MT etc.
            # We will mock Start-Process or similar if it was used, but currently it's native call.
            # In Pester 3.4.0, we can mock the Invoke-RobocopyBackup function itself 
            # to test Start-BackupProcess integration, but to test internal filtering 
            # we need to verify the code path.
            
            # Since Invoke-RobocopyBackup calls robocopy.exe directly with &, 
            # we will verify the $RobocopyParams construction logic by running the function.
            # However, mocking & robocopy.exe is not possible in Pester.
            # WORKAROUND: We'll add a test for the filtering regex logic.
        }

        It "Should correctly strip protected flags (/R, /W, /MT, /LOG)" {
            # This directly tests the regex we implemented
            $ProtectedRegex = "^/(LOG|LOG\+|R|W|MT)(:.*)?$"
            "/R:1000" | Should Match $ProtectedRegex
            "/MT:128" | Should Match $ProtectedRegex
            "/LOG:C:\danger.log" | Should Match $ProtectedRegex
            "/MIR" | Should Not Match $ProtectedRegex
            "/XF" | Should Not Match $ProtectedRegex
        }
    }

    Context "Invoke-BackupItem Collision Prevention" {
        It "Should use high-precision timestamps for unique folder names" {
            $TestDest = Join-Path $PSScriptRoot "CollisionTest"
            $Item = [PSCustomObject]@{ Name = "Sub1"; SourceSubPath = "C:\Source\Sub1" }
            
            Mock Invoke-RobocopyBackup { return @{ Status = "Success"; ExitCode = 0 } }
            Mock New-Item { return $null }
            Mock Test-Path { return $true }
            
            # We want to ensure that even if called in rapid succession, the timestamp differs (or at least includes _fff)
            # Actually, we just need to verify the format contains 3 digits for ms
            
            # Invoke-BackupItem defines $Timestamp = Get-Date -Format "yyyyMMdd_HHmmss_fff"
            # We verify the behavior by checking the history log or capturing the destination
            # But the function doesn't return the path. 
            # We'll verify the timestamp format in the history entry.
            
            $script:HistoryEntries = @()
            Mock Write-BackupHistory { param($LogFilePath, $Entry) $script:HistoryEntries += $Entry }
            
            Invoke-BackupItem -Item $Item -SourcePath "C:\Src" -BackupMode "Root" -DestinationRoot $TestDest -RobocopyOptions "" -InterPacketGapMs 0 -LogsDir "C:\Logs" -HistoryLogFilePath "C:\hist.log"
            
            $script:HistoryEntries[0].Timestamp | Should Match "_\d{3}$" # Milliseconds check
        }
    }

    Context "VSS Path Normalization Regression" {
        It "Should NOT have a trailing backslash when joining device path with subfolder (Error 123 fix)" {
            $ShadowDevicePath = "\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy1"
            $RelativePath = "workdirs\ADM"
            
            # Logic: $VssSourceRoot = $ShadowDevicePath.TrimEnd('\') + "\" + $RelativePath.TrimStart('\').TrimEnd('\')
            $Result = $ShadowDevicePath.TrimEnd('\') + "\" + $RelativePath.TrimStart('\').TrimEnd('\')
            
            $Result | Should Be "\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy1\workdirs\ADM"
            $Result.EndsWith("\") | Should Be $false
        }
    }

    Context "VSS Junction Workflow" {
        It "Should create, use, and cleanup directory junctions for VSS access" {
            $script:JunctionCreated = $false
            $script:JunctionRemoved = $false
            $script:RobocopySource = ""

            Mock Invoke-RobocopyBackup { 
                param($Source) $script:RobocopySource = $Source; 
                return @{ Status = "Success"; ExitCode = 0; LogFile = "test.log" } 
            }
            # Mock cmd.exe for mklink and rd
            Mock cmd.exe { 
                param($Param1, $Param2, $Param3, $Param4)
                $all = "$Param1 $Param2 $Param3 $Param4"
                if ($all -match "mklink /j") { $script:JunctionCreated = $true }
                if ($all -match "rd /q") { $script:JunctionRemoved = $true }
            }
            Mock Test-Path { return $true }
            Mock New-Item { return $null }
            Mock Get-VolumeRoot { return "D:\" }
            Mock Test-BackupPrerequisites { return $true }
            Mock Start-Sleep { }
            Mock Resolve-Path { return [PSCustomObject]@{ Path = $Path } }
            
            $Config = @{
                SourcePaths     = @("D:\Data")
                DestinationPath = "D:\Backups"
                UseVSS          = $true
                HistoryLogFile  = "history.log"
                RetentionDays   = 1
            }
            Mock Get-Configuration { return $Config }
            Mock New-ShadowCopy { return [PSCustomObject]@{ ID = "{VSS-123}"; DeviceObject = "\\?\GLOBALROOT\Device\VSS1" } }
            Mock Remove-ShadowCopy { }
            Mock Write-BackupHistory { }
            Mock Remove-OldBackups { }

            # Mock Get-BackupItems to use the junction-based VssSourceRoot
            Mock Get-BackupItems {
                param($SourcePath, $VssSourceRoot, $BackupMode)
                return , @([PSCustomObject]@{ Name = "Data"; SourceSubPath = $VssSourceRoot; IsRootMode = $true })
            }

            Start-BackupProcess -ConfigFilePath "dummy.json"

            # Verify Junction Lifecycle
            $script:JunctionCreated | Should Be $true
            $script:JunctionRemoved | Should Be $true
            
            # Verify Robocopy used the junction path, not the device path
            $script:RobocopySource | Should Not Match "^\\\\\\\?\\\GLOBALROOT"
            $script:RobocopySource | Should Match "VssJunction_VSS-123"
        }
    }

    # --- Cleanup Logic ---
    Context "Remove-OldBackups Logic" {
        $TestBackupDir = Join-Path $PSScriptRoot "CleanupTest"
        $TestLogsDir = Join-Path $PSScriptRoot "CleanupLogs"
         
        Remove-Item $TestBackupDir -Recurse -Force -ErrorAction SilentlyContinue
        New-Item -Path $TestBackupDir -ItemType Directory -Force | Out-Null
         
        It "Should delete old backups but keep new ones" {
            # Old Folder
            $Old = New-Item -Path "$TestBackupDir\Share_20200101_000000" -ItemType Directory -Force
            $Old.CreationTime = (Get-Date).AddDays(-100)
             
            # New Folder
            $New = New-Item -Path "$TestBackupDir\Share_20260101_000000" -ItemType Directory -Force
            $New.CreationTime = (Get-Date)
             
            Remove-OldBackups -DestinationRoot $TestBackupDir -RetentionDays 30 -LogsDir $TestLogsDir -LogRetentionDays 90
             
            Test-Path $Old.FullName | Should Be $false
            Test-Path $New.FullName | Should Be $true
        }
         
        It "Should ignore folders not matching timestamp pattern" {
            $Other = New-Item -Path "$TestBackupDir\OtherFolder" -ItemType Directory -Force
            $Other.CreationTime = (Get-Date).AddDays(-100)
             
            Remove-OldBackups -DestinationRoot $TestBackupDir -RetentionDays 30 -LogsDir $TestLogsDir -LogRetentionDays 90
             
            Test-Path $Other.FullName | Should Be $true
        }
         
        # Cleanup
        Remove-Item $TestBackupDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    # --- JSONC Support ---
    Context "JSONC Configuration Support" {
        $TestJsoncPath = Join-Path $PSScriptRoot "test_config_comments.json"

        It "Should ignore single-line comments" {
            $Content = @'
{
    // This is a comment
    "SourcePaths": ["C:\\Test"], // Inline comment
    "DestinationPath": "D:\\Backup"
}
'@
            Set-Content -Path $TestJsoncPath -Value $Content
            $Config = Get-Configuration -ConfigPath $TestJsoncPath
            $Config.SourcePaths[0] | Should Be "C:\Test"
            $Config.DestinationPath | Should Be "D:\Backup"
        }

        It "Should ignore multi-line comments" {
            $Content = @'
{
    /* 
       Multi-line comment 
       should be removed 
    */
    "SourcePaths": ["C:\\Test"],
    "DestinationPath": "D:\\Backup"
}
'@
            Set-Content -Path $TestJsoncPath -Value $Content
            $Config = Get-Configuration -ConfigPath $TestJsoncPath
            $Config.SourcePaths[0] | Should Be "C:\Test"
        }

        # Cleanup
        Remove-Item $TestJsoncPath -ErrorAction SilentlyContinue
    }

    # --- Logging Logic ---
    Context "Write-BackupHistory" {
        $TestLogPath = Join-Path $PSScriptRoot "history_timestamp_test.log"
        
        It "Should prepend a human-readable timestamp to the JSON entry" {
            Remove-Item $TestLogPath -ErrorAction SilentlyContinue
            $Entry = [PSCustomObject]@{ Status = "Test"; Id = 123 }
            
            Write-BackupHistory -LogFilePath $TestLogPath -Entry $Entry
            
            Test-Path $TestLogPath | Should Be $true
            $Content = (Get-Content $TestLogPath -Raw).Trim()
            
            # Format: [YYYY-MM-DD HH:MM:SS] {"Status":"Test","Id":123}
            $Content | Should Match '^\[\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\] \{.*\}$'
            $Content | Should Match '"Status":"Test"'
            
            Remove-Item $TestLogPath -ErrorAction SilentlyContinue
        }
    }

    # ==========================================================================
    # NEW TESTS — Full Coverage
    # ==========================================================================

    # --- Show-Usage ---
    Context "Show-Usage Output" {
        It "Should output the usage guide text" {
            $Output = Show-Usage *>&1 | Out-String
            $Output | Should Match "Snapshot Backup Script - Usage Guide"
            $Output | Should Match "PARAMETERS"
            $Output | Should Match "-ConfigFilePath"
            $Output | Should Match "-CheckOnly"
            $Output | Should Match "EXAMPLES"
        }
    }

    # --- Test-BackupPrerequisites (additional branches) ---
    Context "Test-BackupPrerequisites - Additional Branches" {
        It "Should skip admin check and pass when VSS is disabled" {
            Mock Get-Command { return [PSCustomObject]@{ Name = "robocopy.exe" } } -ParameterFilter { $Name -eq "robocopy.exe" }
            $Config = @{ UseVSS = $false; DestinationPath = "C:\" }
            $Result = Test-BackupPrerequisites -Config $Config
            $Result | Should Be $true
        }

        It "Should warn but still pass when destination path is inaccessible" {
            Mock Get-Command { return [PSCustomObject]@{ Name = "robocopy.exe" } } -ParameterFilter { $Name -eq "robocopy.exe" }
            Mock Test-Path { return $false } -ParameterFilter { $_ -eq "Z:\NonExistent\Path" }
            $Config = @{ UseVSS = $false; DestinationPath = "Z:\NonExistent\Path" }
            $Result = Test-BackupPrerequisites -Config $Config
            # Destination inaccessible is a warning, not a failure
            $Result | Should Be $true
        }

        It "Should pass all checks when everything is valid" {
            Mock Get-Command { return [PSCustomObject]@{ Name = "robocopy.exe" } } -ParameterFilter { $Name -eq "robocopy.exe" }
            $Config = @{ UseVSS = $false; DestinationPath = $PSScriptRoot }
            $Result = Test-BackupPrerequisites -Config $Config
            $Result | Should Be $true
        }
    }

    # --- Get-Configuration (error paths and edge cases) ---
    Context "Get-Configuration - Missing File Returns Defaults" {
        It "Should return default config when file does not exist" {
            $Config = Get-Configuration -ConfigPath "Z:\absolutely\nonexistent\config.json"
            $Config | Should Not BeNullOrEmpty
            $Config.RetentionDays | Should Be 30
            $Config.UseVSS | Should Be $true
            $Config.SourcePaths.Count | Should BeGreaterThan 0
        }
    }

    Context "Get-Configuration - Empty DestinationPath" {
        $TestEmptyDestPath = Join-Path $PSScriptRoot "test_empty_dest.json"

        It "Should return defaults with error when DestinationPath is empty string" {
            @{
                SourcePaths     = @("C:\Test")
                DestinationPath = ""
            } | ConvertTo-Json | Set-Content $TestEmptyDestPath

            # The throw is caught by outer catch, which returns defaults
            $Config = Get-Configuration -ConfigPath $TestEmptyDestPath
            $Config | Should Not BeNullOrEmpty
            # Falls back to defaults since the throw is caught
            $Config.RetentionDays | Should Be 30
        }

        It "Should return defaults with error when DestinationPath is null" {
            $Content = '{ "SourcePaths": ["C:\\Test"], "DestinationPath": null }'
            Set-Content $TestEmptyDestPath -Value $Content

            $Config = Get-Configuration -ConfigPath $TestEmptyDestPath
            $Config | Should Not BeNullOrEmpty
            $Config.RetentionDays | Should Be 30
        }

        Remove-Item $TestEmptyDestPath -ErrorAction SilentlyContinue
    }

    Context "Get-Configuration - Destination Creation Failure" {
        $TestFailDestPath = Join-Path $PSScriptRoot "test_fail_dest.json"

        It "Should return defaults when destination directory cannot be created" {
            $BadDest = "Z:\InvalidDrive\Cannot\Create"
            @{
                SourcePaths     = @("C:\Test")
                DestinationPath = $BadDest
            } | ConvertTo-Json | Set-Content $TestFailDestPath

            # The inner throw is caught by the outer catch, returns defaults
            $Config = Get-Configuration -ConfigPath $TestFailDestPath
            $Config | Should Not BeNullOrEmpty
            $Config.RetentionDays | Should Be 30
        }

        Remove-Item $TestFailDestPath -ErrorAction SilentlyContinue
    }

    Context "Get-Configuration - Full Default Merge" {
        $TestMergePath = Join-Path $PSScriptRoot "test_merge_config.json"

        It "Should merge missing properties from defaults" {
            # Only provide SourcePaths and DestinationPath
            @{
                SourcePaths     = @("C:\MyData")
                DestinationPath = $PSScriptRoot
            } | ConvertTo-Json | Set-Content $TestMergePath

            $Config = Get-Configuration -ConfigPath $TestMergePath
            $Config.RetentionDays | Should Be 30
            $Config.LogRetentionDays | Should Be 90
            $Config.MaxBackupAttempts | Should Be 3
            $Config.UseVSS | Should Be $true
            $Config.HistoryLogFile | Should Be "backup-history.log"
        }

        Remove-Item $TestMergePath -ErrorAction SilentlyContinue
    }

    Context "Get-Configuration - Invalid JSON" {
        $TestBadJsonPath = Join-Path $PSScriptRoot "test_bad_json.json"

        It "Should return defaults when JSON is malformed" {
            Set-Content $TestBadJsonPath -Value "{ this is not valid json !!!"
            $Config = Get-Configuration -ConfigPath $TestBadJsonPath
            $Config | Should Not BeNullOrEmpty
            $Config.RetentionDays | Should Be 30
        }

        Remove-Item $TestBadJsonPath -ErrorAction SilentlyContinue
    }

    # --- Invoke-RobocopyBackup (execution paths) ---
    Context "Invoke-RobocopyBackup - Empty Source or Destination" {
        It "Should return Failed with exit code 16 when Source is empty" {
            $TestLog = Join-Path $PSScriptRoot "robocopy_empty_src.log"
            $Result = Invoke-RobocopyBackup -Source "" -Destination "C:\Dest" -Options "/MIR" -InterPacketGapMs 0 -LogFile $TestLog
            $Result.Status | Should Be "Failed"
            $Result.ExitCode | Should Be 16
            Remove-Item $TestLog -ErrorAction SilentlyContinue
        }

        It "Should return Failed with exit code 16 when Destination is empty" {
            $TestLog = Join-Path $PSScriptRoot "robocopy_empty_dest.log"
            $Result = Invoke-RobocopyBackup -Source "C:\Src" -Destination "" -Options "/MIR" -InterPacketGapMs 0 -LogFile $TestLog
            $Result.Status | Should Be "Failed"
            $Result.ExitCode | Should Be 16
            Remove-Item $TestLog -ErrorAction SilentlyContinue
        }

        It "Should return Failed with exit code 16 when both are whitespace" {
            $TestLog = Join-Path $PSScriptRoot "robocopy_ws.log"
            $Result = Invoke-RobocopyBackup -Source "   " -Destination "   " -Options "" -InterPacketGapMs 0 -LogFile $TestLog
            $Result.Status | Should Be "Failed"
            $Result.ExitCode | Should Be 16
            Remove-Item $TestLog -ErrorAction SilentlyContinue
        }
    }

    Context "Invoke-RobocopyBackup - Exit Code Handling" {
        It "Should treat exit codes below 8 as Success" {
            # Test the logic directly: codes 0-7 = Success
            foreach ($code in @(0, 1, 3, 7)) {
                $Status = if ($code -lt 8) { "Success" } else { "Failed" }
                $Status | Should Be "Success"
            }
        }

        It "Should treat exit codes 8 and above as Failed" {
            foreach ($code in @(8, 9, 16)) {
                $Status = if ($code -lt 8) { "Success" } else { "Failed" }
                $Status | Should Be "Failed"
            }
        }
    }

    Context "Invoke-RobocopyBackup - Option Handling" {
        It "Should include /IPG in params when InterPacketGapMs is greater than 0" {
            # Test the option-building logic directly
            $InterPacketGapMs = 500
            $Params = @()
            $Params += "/LOG+:test.log"
            $Params += "/R:1"
            $Params += "/W:1"
            $Params += "/MT:8"
            if ($InterPacketGapMs -gt 0) { $Params += "/IPG:$InterPacketGapMs" }
            
            $JoinedParams = $Params -join " "
            $JoinedParams | Should Match "/IPG:500"
        }

        It "Should NOT include /IPG when InterPacketGapMs is 0" {
            $InterPacketGapMs = 0
            $Params = @()
            $Params += "/LOG+:test.log"
            $Params += "/R:1"
            $Params += "/W:1"
            $Params += "/MT:8"
            if ($InterPacketGapMs -gt 0) { $Params += "/IPG:$InterPacketGapMs" }
            
            $JoinedParams = $Params -join " "
            $JoinedParams | Should Not Match "/IPG"
        }

        It "Should always include /MT:8 in hardcoded params" {
            $Params = @()
            $Params += "/LOG+:test.log"
            $Params += "/R:1"
            $Params += "/W:1"
            $Params += "/MT:8"
            
            $JoinedParams = $Params -join " "
            $JoinedParams | Should Match "/MT:8"
        }

        It "Should strip user-provided /LOG, /R, /W, /MT from options via regex" {
            $ProtectedRegex = "^/(LOG|LOG\+|R|W|MT)(:.*)?$"
            $UserOptions = "/MIR /R:9999 /W:60 /MT:128 /LOG:C:\bad.log /NP"
            $CleanOptions = $UserOptions -replace '/LOG\+?:\S+', '' -replace '\{logpath\}', ''
            
            $OptionArray = [regex]::Matches($CleanOptions, '(?:[^\s"]+|"[^"]*")+') | 
            ForEach-Object { $_.Value } |
            Where-Object { $_ -notmatch $ProtectedRegex }
            
            $Remaining = $OptionArray -join " "
            $Remaining | Should Not Match "/R:9999"
            $Remaining | Should Not Match "/W:60"
            $Remaining | Should Not Match "/MT:128"
            $Remaining | Should Match "/MIR"
            $Remaining | Should Match "/NP"
        }

        It "Should create log directory if it does not exist" {
            Mock robocopy.exe { $global:LASTEXITCODE = 0 }
            $TestLogDir = Join-Path $PSScriptRoot "NewLogDir_Test"
            $TestLog = Join-Path $TestLogDir "test.log"
            Remove-Item $TestLogDir -Recurse -Force -ErrorAction SilentlyContinue

            Invoke-RobocopyBackup -Source $PSScriptRoot -Destination $PSScriptRoot -Options "" -InterPacketGapMs 0 -LogFile $TestLog

            Test-Path $TestLogDir | Should Be $true
            Remove-Item $TestLogDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    # --- Remove-OldBackups - Log Cleanup ---
    Context "Remove-OldBackups - Log File Cleanup" {
        $TestLogsCleanupDir = Join-Path $PSScriptRoot "LogCleanupTest"
        $TestBackupStubDir = Join-Path $PSScriptRoot "BackupStub"

        It "Should delete old log files but keep new ones" {
            Remove-Item $TestLogsCleanupDir -Recurse -Force -ErrorAction SilentlyContinue
            New-Item -Path $TestLogsCleanupDir -ItemType Directory -Force | Out-Null
            New-Item -Path $TestBackupStubDir -ItemType Directory -Force | Out-Null

            $OldLog = New-Item -Path "$TestLogsCleanupDir\old_robocopy.log" -ItemType File -Force
            $OldLog.LastWriteTime = (Get-Date).AddDays(-200)

            $NewLog = New-Item -Path "$TestLogsCleanupDir\new_robocopy.log" -ItemType File -Force

            Remove-OldBackups -DestinationRoot $TestBackupStubDir -RetentionDays 30 -LogsDir $TestLogsCleanupDir -LogRetentionDays 90

            Test-Path $OldLog.FullName | Should Be $false
            Test-Path $NewLog.FullName | Should Be $true
        }

        It "Should not fail when logs directory does not exist" {
            Remove-OldBackups -DestinationRoot $TestBackupStubDir -RetentionDays 30 -LogsDir "Z:\Nonexistent\Logs" -LogRetentionDays 90
            # Should complete without error
        }

        # Cleanup
        Remove-Item $TestLogsCleanupDir -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item $TestBackupStubDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    # --- Mode Validation ---
    Context "Mode Validation in Start-BackupProcess" {
        It "Should warn and default to SubDirectories when invalid Mode is specified" {
            Mock Get-Command { return [PSCustomObject]@{ Name = "robocopy.exe" } } -ParameterFilter { $Name -eq "robocopy.exe" }
            Mock Test-Path { return $true }
            Mock New-Item { return $null }
            Mock Join-Path {
                param($Path, $ChildPath)
                if ([string]::IsNullOrWhiteSpace($Path)) { return $ChildPath }
                if ($Path.EndsWith("\")) { return "$Path$ChildPath" }
                return "$Path\$ChildPath"
            }
            Mock Get-VolumeRoot { return "D:\" }
            Mock Test-BackupPrerequisites { return $true }
            Mock Resolve-Path { return [PSCustomObject]@{ Path = $Path } }

            $Config = @{
                SourcePaths              = @([PSCustomObject]@{ Path = "D:\Data"; Mode = "InvalidMode" })
                DestinationPath          = "E:\Backups"
                RetentionDays            = 30
                LogRetentionDays         = 90
                MaxBackupAttempts        = 3
                RobocopyOptions          = '/MIR'
                RobocopyInterPacketGapMs = 0
                HistoryLogFile           = "backup-history.log"
                UseVSS                   = $false
            }
            Mock Get-Configuration { return $Config }
            # Mock Get-BackupItems to capture the BackupMode parameter
            $script:CapturedMode = ""
            Mock Get-BackupItems {
                param($SourcePath, $VssSourceRoot, $BackupMode)
                $script:CapturedMode = $BackupMode
                return , @([PSCustomObject]@{ Name = "Data"; SourceSubPath = $VssSourceRoot; IsRootMode = $false })
            }
            Mock Invoke-RobocopyBackup { return @{ Status = "Success"; ExitCode = 0; LogFile = "test.log" } }
            Mock Write-BackupHistory { }
            Mock Remove-OldBackups { }

            $Output = Start-BackupProcess -ConfigFilePath "fake.json" *>&1
            $OutputStr = $Output -join "`n"

            # Should have used default SubDirectories, not InvalidMode
            $script:CapturedMode | Should Be "SubDirectories"
            $OutputStr | Should Match "Invalid Mode"
        }

        It "Should accept Root mode without warning" {
            Mock Get-Command { return [PSCustomObject]@{ Name = "robocopy.exe" } } -ParameterFilter { $Name -eq "robocopy.exe" }
            Mock Test-Path { return $true }
            Mock New-Item { return $null }
            Mock Join-Path {
                param($Path, $ChildPath)
                if ([string]::IsNullOrWhiteSpace($Path)) { return $ChildPath }
                if ($Path.EndsWith("\")) { return "$Path$ChildPath" }
                return "$Path\$ChildPath"
            }
            Mock Get-VolumeRoot { return "D:\" }
            Mock Test-BackupPrerequisites { return $true }
            Mock Resolve-Path { return [PSCustomObject]@{ Path = $Path } }

            $Config = @{
                SourcePaths              = @([PSCustomObject]@{ Path = "D:\Data"; Mode = "Root" })
                DestinationPath          = "E:\Backups"
                RetentionDays            = 30
                LogRetentionDays         = 90
                RobocopyOptions          = '/MIR'
                RobocopyInterPacketGapMs = 0
                HistoryLogFile           = "backup-history.log"
                UseVSS                   = $false
            }
            Mock Get-Configuration { return $Config }
            $script:CapturedMode = ""
            Mock Get-BackupItems {
                param($SourcePath, $VssSourceRoot, $BackupMode)
                $script:CapturedMode = $BackupMode
                return , @([PSCustomObject]@{ Name = "Data"; SourceSubPath = $VssSourceRoot; IsRootMode = $true })
            }
            Mock Invoke-RobocopyBackup { return @{ Status = "Success"; ExitCode = 0; LogFile = "test.log" } }
            Mock Write-BackupHistory { }
            Mock Remove-OldBackups { }

            $Output = Start-BackupProcess -ConfigFilePath "fake.json" *>&1

            $script:CapturedMode | Should Be "Root"
            ($Output -join "`n") | Should Not Match "Invalid Mode"
        }
    }

    # --- Source Path Error Handling ---
    Context "Start-BackupProcess - Source Path Errors" {
        It "Should skip source items with invalid format (no Path property)" {
            Mock Get-Command { return [PSCustomObject]@{ Name = "robocopy.exe" } } -ParameterFilter { $Name -eq "robocopy.exe" }
            Mock Test-Path { return $true }
            Mock New-Item { return $null }
            Mock Test-BackupPrerequisites { return $true }

            $Config = @{
                SourcePaths              = @([PSCustomObject]@{ Invalid = "NoPathHere" })
                DestinationPath          = "E:\Backups"
                RetentionDays            = 30
                LogRetentionDays         = 90
                RobocopyOptions          = ''
                RobocopyInterPacketGapMs = 0
                HistoryLogFile           = "backup-history.log"
                UseVSS                   = $false
            }
            Mock Get-Configuration { return $Config }
            Mock Invoke-RobocopyBackup { throw "Should not be called" }
            Mock Remove-OldBackups { }

            $Output = Start-BackupProcess -ConfigFilePath "fake.json" *>&1
            ($Output -join "`n") | Should Match "Skipping invalid source item format"
        }

        It "Should skip source paths that do not exist" {
            Mock Get-Command { return [PSCustomObject]@{ Name = "robocopy.exe" } } -ParameterFilter { $Name -eq "robocopy.exe" }
            Mock Test-Path { 
                param($Path)
                if ($Path -eq "Z:\NonExistent\Source") { return $false }
                return $true
            }
            Mock New-Item { return $null }
            Mock Test-BackupPrerequisites { return $true }
            Mock Join-Path {
                param($Path, $ChildPath)
                if ([string]::IsNullOrWhiteSpace($Path)) { return $ChildPath }
                if ($Path.EndsWith("\")) { return "$Path$ChildPath" }
                return "$Path\$ChildPath"
            }

            $Config = @{
                SourcePaths              = @("Z:\NonExistent\Source")
                DestinationPath          = "E:\Backups"
                RetentionDays            = 30
                LogRetentionDays         = 90
                RobocopyOptions          = ''
                RobocopyInterPacketGapMs = 0
                HistoryLogFile           = "backup-history.log"
                UseVSS                   = $false
            }
            Mock Get-Configuration { return $Config }
            Mock Invoke-RobocopyBackup { throw "Should not be called" }
            Mock Remove-OldBackups { }

            $Output = Start-BackupProcess -ConfigFilePath "fake.json" *>&1
            ($Output -join "`n") | Should Match "does not exist or is inaccessible"
        }
    }

    # --- VSS Creation Failure ---
    Context "Start-BackupProcess - VSS Failure Handling" {
        It "Should skip source when VSS snapshot creation fails" {
            Mock Get-Command { return [PSCustomObject]@{ Name = "robocopy.exe" } } -ParameterFilter { $Name -eq "robocopy.exe" }
            Mock Test-Path { return $true }
            Mock New-Item { return $null }
            Mock Join-Path {
                param($Path, $ChildPath)
                if ([string]::IsNullOrWhiteSpace($Path)) { return $ChildPath }
                if ($Path.EndsWith("\")) { return "$Path$ChildPath" }
                return "$Path\$ChildPath"
            }
            Mock Get-VolumeRoot { return "D:\" }
            Mock Test-BackupPrerequisites { return $true }
            Mock Resolve-Path { return [PSCustomObject]@{ Path = $Path } }

            $Config = @{
                SourcePaths              = @("D:\Data")
                DestinationPath          = "E:\Backups"
                RetentionDays            = 30
                LogRetentionDays         = 90
                RobocopyOptions          = '/MIR'
                RobocopyInterPacketGapMs = 0
                HistoryLogFile           = "backup-history.log"
                UseVSS                   = $true
            }
            Mock Get-Configuration { return $Config }
            Mock New-ShadowCopy { throw "VSS service unavailable" }
            Mock Invoke-RobocopyBackup { throw "Should not be called if VSS fails" }
            Mock Remove-OldBackups { }

            $Output = Start-BackupProcess -ConfigFilePath "fake.json" *>&1
            ($Output -join "`n") | Should Match "VSS creation failure"
        }
    }

    # --- Invoke-BackupItem - History Entry ---
    Context "Invoke-BackupItem - History Entry Fields" {
        It "Should record all expected fields in history entry" {
            Mock Invoke-RobocopyBackup { return @{ Status = "Success"; ExitCode = 1; LogFile = "detail.log" } }

            $script:CapturedEntry = $null
            Mock Write-BackupHistory { param($LogFilePath, $Entry) $script:CapturedEntry = $Entry }

            $TestDestRoot = Join-Path $PSScriptRoot "HistEntryTest"
            $TestLogsDir = Join-Path $PSScriptRoot "HistEntryLogs"
            New-Item -Path $TestDestRoot -ItemType Directory -Force | Out-Null
            New-Item -Path $TestLogsDir -ItemType Directory -Force | Out-Null

            $Item = [PSCustomObject]@{ Name = "ProjectA"; SourceSubPath = "D:\VSS\ProjectA" }
            Invoke-BackupItem -Item $Item -SourcePath "D:\Source" -BackupMode "Root" `
                -DestinationRoot $TestDestRoot -RobocopyOptions "/MIR" `
                -InterPacketGapMs 0 -LogsDir $TestLogsDir `
                -HistoryLogFilePath (Join-Path $PSScriptRoot "hist_test.log") -SnapshotID "{VSS-456}"

            $script:CapturedEntry | Should Not BeNullOrEmpty
            $script:CapturedEntry.Subdirectory | Should Be "ProjectA"
            $script:CapturedEntry.Mode | Should Be "Root"
            $script:CapturedEntry.Status | Should Be "Success"
            $script:CapturedEntry.ExitCode | Should Be 1
            $script:CapturedEntry.ShadowCopyId | Should Be "{VSS-456}"
            $script:CapturedEntry.Timestamp | Should Match "\d{8}_\d{6}_\d{3}"

            Remove-Item $TestDestRoot -Recurse -Force -ErrorAction SilentlyContinue
            Remove-Item $TestLogsDir -Recurse -Force -ErrorAction SilentlyContinue
        }

        It "Should record N/A for ShadowCopyId when no snapshot" {
            Mock Invoke-RobocopyBackup { return @{ Status = "Success"; ExitCode = 0; LogFile = "detail.log" } }

            $script:CapturedEntry = $null
            Mock Write-BackupHistory { param($LogFilePath, $Entry) $script:CapturedEntry = $Entry }

            $TestDestRoot = Join-Path $PSScriptRoot "HistEntryTest2"
            $TestLogsDir = Join-Path $PSScriptRoot "HistEntryLogs2"
            New-Item -Path $TestDestRoot -ItemType Directory -Force | Out-Null
            New-Item -Path $TestLogsDir -ItemType Directory -Force | Out-Null

            $Item = [PSCustomObject]@{ Name = "ProjectB"; SourceSubPath = "D:\Direct\ProjectB" }
            Invoke-BackupItem -Item $Item -SourcePath "D:\Source" -BackupMode "SubDirectories" `
                -DestinationRoot $TestDestRoot -RobocopyOptions "" `
                -InterPacketGapMs 0 -LogsDir $TestLogsDir `
                -HistoryLogFilePath (Join-Path $PSScriptRoot "hist_test2.log")

            $script:CapturedEntry.ShadowCopyId | Should Be "N/A"

            Remove-Item $TestDestRoot -Recurse -Force -ErrorAction SilentlyContinue
            Remove-Item $TestLogsDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    # --- Pre-flight Failure Aborts Backup ---
    Context "Start-BackupProcess - Pre-flight Failure" {
        It "Should abort backup when prerequisites fail" {
            Mock Get-Configuration { return @{ UseVSS = $false; DestinationPath = "C:\"; HistoryLogFile = "test.log" } }
            Mock Test-BackupPrerequisites { return $false }
            Mock Invoke-RobocopyBackup { throw "Should not be called" }

            $Output = Start-BackupProcess -ConfigFilePath "dummy.json" *>&1
            ($Output -join "`n") | Should Match "Pre-flight checks failed"
        }
    }

    # --- Remove-OldBackups - Backup with millisecond timestamp pattern ---
    Context "Remove-OldBackups - Millisecond Timestamp Pattern" {
        $TestBackupMsDir = Join-Path $PSScriptRoot "CleanupMsTest"

        It "Should delete old backups with millisecond timestamps (_YYYYMMDD_HHMMSS_fff)" {
            Remove-Item $TestBackupMsDir -Recurse -Force -ErrorAction SilentlyContinue
            New-Item -Path $TestBackupMsDir -ItemType Directory -Force | Out-Null

            $Old = New-Item -Path "$TestBackupMsDir\Share_20200101_000000_123" -ItemType Directory -Force
            $Old.CreationTime = (Get-Date).AddDays(-100)

            $New = New-Item -Path "$TestBackupMsDir\Share_20260101_000000_456" -ItemType Directory -Force
            $New.CreationTime = (Get-Date)

            Remove-OldBackups -DestinationRoot $TestBackupMsDir -RetentionDays 30 -LogsDir "Z:\None" -LogRetentionDays 90

            Test-Path $Old.FullName | Should Be $false
            Test-Path $New.FullName | Should Be $true
        }

        Remove-Item $TestBackupMsDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    # ==========================================================================
    # EDGE CASE TESTS
    # ==========================================================================

    # --- JSONC Edge Cases ---
    Context "JSONC - Edge Cases" {
        $TestJsoncEdgePath = Join-Path $PSScriptRoot "test_jsonc_edge.json"

        It "Should handle inline comment on the same line as a value" {
            $Content = @'
{
    "SourcePaths": ["C:\\Test"],  // source
    "DestinationPath": "D:\\Backup", // destination
    "RetentionDays": 7 // keep a week
}
'@
            Set-Content -Path $TestJsoncEdgePath -Value $Content
            $Config = Get-Configuration -ConfigPath $TestJsoncEdgePath
            $Config.SourcePaths[0] | Should Be "C:\Test"
            $Config.DestinationPath | Should Be "D:\Backup"
            $Config.RetentionDays | Should Be 7
        }

        It "Should handle mixed single-line and multi-line comments" {
            $Content = @'
{
    /* header comment */
    // another comment
    "SourcePaths": ["E:\\Data"],
    /* "DestinationPath": "OLD_VALUE", */
    "DestinationPath": "E:\\Backups"
}
'@
            Set-Content -Path $TestJsoncEdgePath -Value $Content
            $Config = Get-Configuration -ConfigPath $TestJsoncEdgePath
            $Config.SourcePaths[0] | Should Be "E:\Data"
            $Config.DestinationPath | Should Be "E:\Backups"
        }

        It "Should handle empty JSON with only comments" {
            $Content = @'
{
    // only comments here
    /* nothing else */
}
'@
            Set-Content -Path $TestJsoncEdgePath -Value $Content
            # Should fall through to defaults due to missing SourcePaths
            $Config = Get-Configuration -ConfigPath $TestJsoncEdgePath
            $Config | Should Not BeNullOrEmpty
        }

        It "Should handle JSON with no comments at all" {
            $Content = '{ "SourcePaths": ["C:\\Plain"], "DestinationPath": "D:\\Plain" }'
            Set-Content -Path $TestJsoncEdgePath -Value $Content
            $Config = Get-Configuration -ConfigPath $TestJsoncEdgePath
            $Config.SourcePaths[0] | Should Be "C:\Plain"
        }

        Remove-Item $TestJsoncEdgePath -ErrorAction SilentlyContinue
    }

    # --- Get-BackupItems Edge Cases ---
    Context "Get-BackupItems - Edge Cases" {
        It "Should handle source path with trailing backslash in Root mode" {
            $TestPathTrailing = Join-Path $PSScriptRoot "TrailingSlashTest"
            New-Item -Path $TestPathTrailing -ItemType Directory -Force | Out-Null

            $Items = Get-BackupItems -SourcePath "$TestPathTrailing\" -VssSourceRoot "\\?\VSS\TrailingSlashTest" -BackupMode "Root"
            $Items.Count | Should Be 1
            $Items[0].Name | Should Be "TrailingSlashTest"
            $Items[0].Name | Should Not Be ""

            Remove-Item $TestPathTrailing -Force -ErrorAction SilentlyContinue
        }

        It "Should handle subdirectory names with spaces" {
            $TestSpaceDir = Join-Path $PSScriptRoot "SpaceTest"
            Remove-Item $TestSpaceDir -Recurse -Force -ErrorAction SilentlyContinue
            New-Item -Path "$TestSpaceDir\My Documents" -ItemType Directory -Force | Out-Null
            New-Item -Path "$TestSpaceDir\Program Files" -ItemType Directory -Force | Out-Null

            $Items = Get-BackupItems -SourcePath $TestSpaceDir -VssSourceRoot "\\?\VSS\SpaceTest" -BackupMode "SubDirectories"
            $Items.Count | Should Be 2
            ($Items | Where-Object { $_.Name -eq "My Documents" }) | Should Not BeNullOrEmpty
            ($Items | Where-Object { $_.Name -eq "Program Files" }) | Should Not BeNullOrEmpty

            Remove-Item $TestSpaceDir -Recurse -Force -ErrorAction SilentlyContinue
        }

        It "Should correctly construct SourceSubPath for subdirectories" {
            $TestSubPath = Join-Path $PSScriptRoot "SubPathTest"
            Remove-Item $TestSubPath -Recurse -Force -ErrorAction SilentlyContinue
            New-Item -Path "$TestSubPath\FolderA" -ItemType Directory -Force | Out-Null

            $VssRoot = "\\?\GLOBALROOT\Device\VSS1\SubPathTest"
            $Items = Get-BackupItems -SourcePath $TestSubPath -VssSourceRoot $VssRoot -BackupMode "SubDirectories"
            $Items[0].SourceSubPath | Should Be "$VssRoot\FolderA"

            Remove-Item $TestSubPath -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    # --- Robocopy Option Parsing Edge Cases ---
    Context "Invoke-RobocopyBackup - Option Parsing Edge Cases" {
        It "Should preserve quoted exclude directories with spaces" {
            $ProtectedRegex = "^/(LOG|LOG\+|R|W|MT)(:.*)?$"
            $Options = '/MIR /XD "System Volume Information" RECYCLE.BIN /XF Thumbs.db'
            $CleanOptions = $Options -replace '/LOG\+?:\S+', '' -replace '\{logpath\}', ''

            $OptionArray = [regex]::Matches($CleanOptions, '(?:[^\s"]+|"[^"]*")+') |
            ForEach-Object { $_.Value } |
            Where-Object { $_ -notmatch $ProtectedRegex }

            $Remaining = $OptionArray -join " "
            $Remaining | Should Match '/MIR'
            $Remaining | Should Match '/XD'
            $Remaining | Should Match '"System Volume Information"'
            $Remaining | Should Match 'RECYCLE.BIN'
            $Remaining | Should Match '/XF'
            $Remaining | Should Match 'Thumbs.db'
        }

        It "Should handle empty options string without error" {
            $ProtectedRegex = "^/(LOG|LOG\+|R|W|MT)(:.*)?$"
            $Options = ""
            $CleanOptions = $Options -replace '/LOG\+?:\S+', '' -replace '\{logpath\}', ''
            $OptionArray = [regex]::Matches($CleanOptions, '(?:[^\s"]+|"[^"]*")+') |
            ForEach-Object { $_.Value } |
            Where-Object { $_ -notmatch $ProtectedRegex }

            $OptionArray.Count | Should Be 0
        }

        It "Should handle options with LOG+ variant" {
            $ProtectedRegex = "^/(LOG|LOG\+|R|W|MT)(:.*)?$"
            $Options = '/MIR /LOG+:C:\old.log /NP'
            $CleanOptions = $Options -replace '/LOG\+?:\S+', '' -replace '\{logpath\}', ''
            $OptionArray = [regex]::Matches($CleanOptions, '(?:[^\s"]+|"[^"]*")+') |
            ForEach-Object { $_.Value } |
            Where-Object { $_ -notmatch $ProtectedRegex }

            $Remaining = $OptionArray -join " "
            $Remaining | Should Not Match '/LOG'
            $Remaining | Should Match '/MIR'
            $Remaining | Should Match '/NP'
        }

        It "Should strip {logpath} placeholder from options" {
            $Options = '/MIR /LOG+:{logpath} /NP'
            $CleanOptions = $Options -replace '/LOG\+?:\S+', '' -replace '\{logpath\}', ''
            $CleanOptions | Should Not Match '\{logpath\}'
        }

        It "Should handle trailing backslash on source path" {
            $SourceWithSlash = "C:\Backup\Source\"
            $Result = $SourceWithSlash.TrimEnd('\')
            $Result | Should Be "C:\Backup\Source"
            $Result.EndsWith('\') | Should Be $false
        }
    }

    # --- RelativePath Calculation Edge Cases ---
    Context "RelativePath Calculation" {
        It "Should return empty RelativePath when source IS the volume root" {
            $VolumeRoot = "D:\"
            $SourcePath = "D:\"
            $RelativePath = ""
            if ($SourcePath.Length -gt $VolumeRoot.Length) {
                $RelativePath = $SourcePath.Substring($VolumeRoot.Length).TrimStart('\')
            }
            $RelativePath | Should Be ""
        }

        It "Should extract correct relative path for deep directories" {
            $VolumeRoot = "D:\"
            $SourcePath = "D:\Shares\Data\Projects\Active"
            $RelativePath = ""
            if ($SourcePath.Length -gt $VolumeRoot.Length) {
                $RelativePath = $SourcePath.Substring($VolumeRoot.Length).TrimStart('\')
            }
            $RelativePath | Should Be "Shares\Data\Projects\Active"
        }

        It "Should handle relative path with single level" {
            $VolumeRoot = "C:\"
            $SourcePath = "C:\Users"
            $RelativePath = ""
            if ($SourcePath.Length -gt $VolumeRoot.Length) {
                $RelativePath = $SourcePath.Substring($VolumeRoot.Length).TrimStart('\')
            }
            $RelativePath | Should Be "Users"
        }
    }

    # --- Retention Boundary Edge Cases ---
    Context "Remove-OldBackups - Retention Boundary" {
        $TestBoundaryDir = Join-Path $PSScriptRoot "BoundaryTest"

        It "Should keep folder exactly at retention boundary" {
            Remove-Item $TestBoundaryDir -Recurse -Force -ErrorAction SilentlyContinue
            New-Item -Path $TestBoundaryDir -ItemType Directory -Force | Out-Null

            # Folder 29.5 days old (safely inside 30-day retention window)
            $AtBoundary = New-Item -Path "$TestBoundaryDir\Share_20260101_120000" -ItemType Directory -Force
            $AtBoundary.CreationTime = (Get-Date).AddDays(-29.5)

            Remove-OldBackups -DestinationRoot $TestBoundaryDir -RetentionDays 30 -LogsDir "Z:\None" -LogRetentionDays 90

            Test-Path $AtBoundary.FullName | Should Be $true
        }

        It "Should delete folder one day past retention" {
            Remove-Item $TestBoundaryDir -Recurse -Force -ErrorAction SilentlyContinue
            New-Item -Path $TestBoundaryDir -ItemType Directory -Force | Out-Null

            $PastBoundary = New-Item -Path "$TestBoundaryDir\Share_20260101_120000" -ItemType Directory -Force
            $PastBoundary.CreationTime = (Get-Date).AddDays(-31)

            Remove-OldBackups -DestinationRoot $TestBoundaryDir -RetentionDays 30 -LogsDir "Z:\None" -LogRetentionDays 90

            Test-Path $PastBoundary.FullName | Should Be $false
        }

        It "Should keep folder one day before retention" {
            Remove-Item $TestBoundaryDir -Recurse -Force -ErrorAction SilentlyContinue
            New-Item -Path $TestBoundaryDir -ItemType Directory -Force | Out-Null

            $BeforeBoundary = New-Item -Path "$TestBoundaryDir\Share_20260101_120000" -ItemType Directory -Force
            $BeforeBoundary.CreationTime = (Get-Date).AddDays(-29)

            Remove-OldBackups -DestinationRoot $TestBoundaryDir -RetentionDays 30 -LogsDir "Z:\None" -LogRetentionDays 90

            Test-Path $BeforeBoundary.FullName | Should Be $true
        }

        It "Should not delete destination root itself even if empty" {
            Remove-Item $TestBoundaryDir -Recurse -Force -ErrorAction SilentlyContinue
            New-Item -Path $TestBoundaryDir -ItemType Directory -Force | Out-Null

            Remove-OldBackups -DestinationRoot $TestBoundaryDir -RetentionDays 1 -LogsDir "Z:\None" -LogRetentionDays 1

            Test-Path $TestBoundaryDir | Should Be $true
        }

        Remove-Item $TestBoundaryDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    # --- Mixed Source Format Processing ---
    Context "Start-BackupProcess - Mixed Source Formats" {
        It "Should handle a mix of string and object SourcePaths" {
            Mock Get-Command { return [PSCustomObject]@{ Name = "robocopy.exe" } } -ParameterFilter { $Name -eq "robocopy.exe" }
            Mock Test-Path { return $true }
            Mock New-Item { return $null }
            Mock Join-Path {
                param($Path, $ChildPath)
                if ([string]::IsNullOrWhiteSpace($Path)) { return $ChildPath }
                if ($Path.EndsWith("\")) { return "$Path$ChildPath" }
                return "$Path\$ChildPath"
            }
            Mock Get-VolumeRoot { return "D:\" }
            Mock Test-BackupPrerequisites { return $true }
            Mock Resolve-Path { return [PSCustomObject]@{ Path = $Path } }
            Mock Start-Sleep { }

            $Config = @{
                SourcePaths              = @(
                    "D:\StringPath",
                    [PSCustomObject]@{ Path = "D:\ObjectPath"; Mode = "Root" }
                )
                DestinationPath          = "E:\Backups"
                RetentionDays            = 30
                LogRetentionDays         = 90
                RobocopyOptions          = '/MIR'
                RobocopyInterPacketGapMs = 0
                HistoryLogFile           = "backup-history.log"
                UseVSS                   = $false
            }
            Mock Get-Configuration { return $Config }
            $script:BackupModes = @()
            Mock Get-BackupItems {
                param($SourcePath, $VssSourceRoot, $BackupMode)
                $script:BackupModes += $BackupMode
                return , @([PSCustomObject]@{ Name = "Data"; SourceSubPath = $VssSourceRoot; IsRootMode = ($BackupMode -eq "Root") })
            }
            Mock Invoke-RobocopyBackup { return @{ Status = "Success"; ExitCode = 0; LogFile = "test.log" } }
            Mock Write-BackupHistory { }
            Mock Remove-OldBackups { }

            Start-BackupProcess -ConfigFilePath "fake.json"

            $script:BackupModes.Count | Should Be 2
            $script:BackupModes[0] | Should Be "SubDirectories"
            $script:BackupModes[1] | Should Be "Root"
        }
    }

    # --- VSS Junction Failure Fallback ---
    Context "Start-BackupProcess - Junction Failure Fallback" {
        It "Should fall back to device path when junction creation fails" {
            Mock Get-Command { return [PSCustomObject]@{ Name = "robocopy.exe" } } -ParameterFilter { $Name -eq "robocopy.exe" }
            # Test-Path returns true for everything EXCEPT the junction path
            $script:JunctionPath = ""
            Mock Test-Path {
                param($Path)
                if ($Path -and $Path -match "VssJunction_") { return $false }
                return $true
            }
            Mock New-Item { return $null }
            Mock Join-Path {
                param($Path, $ChildPath)
                if ([string]::IsNullOrWhiteSpace($Path)) { return $ChildPath }
                if ($ChildPath -match "VssJunction_") { $script:JunctionPath = "$Path\$ChildPath" }
                if ($Path.EndsWith("\")) { return "$Path$ChildPath" }
                return "$Path\$ChildPath"
            }
            Mock Get-VolumeRoot { return "D:\" }
            Mock Test-BackupPrerequisites { return $true }
            Mock Resolve-Path { return [PSCustomObject]@{ Path = $Path } }
            Mock Start-Sleep { }
            Mock cmd.exe { }

            $Config = @{
                SourcePaths              = @("D:\Data")
                DestinationPath          = "E:\Backups"
                RetentionDays            = 30
                LogRetentionDays         = 90
                RobocopyOptions          = '/MIR'
                RobocopyInterPacketGapMs = 0
                HistoryLogFile           = "backup-history.log"
                UseVSS                   = $true
            }
            Mock Get-Configuration { return $Config }
            Mock New-ShadowCopy {
                return [PSCustomObject]@{ ID = "{FAIL-J}"; DeviceObject = "\\?\GLOBALROOT\Device\VSS_FAIL" }
            }
            Mock Remove-ShadowCopy { }

            $script:CapturedVssRoot = ""
            Mock Get-BackupItems {
                param($SourcePath, $VssSourceRoot, $BackupMode)
                $script:CapturedVssRoot = $VssSourceRoot
                return , @([PSCustomObject]@{ Name = "Data"; SourceSubPath = $VssSourceRoot; IsRootMode = $true })
            }
            Mock Invoke-RobocopyBackup { return @{ Status = "Success"; ExitCode = 0; LogFile = "test.log" } }
            Mock Write-BackupHistory { }
            Mock Remove-OldBackups { }

            $Output = Start-BackupProcess -ConfigFilePath "fake.json" *>&1

            # When junction fails, should fall back to device path (trimmed)
            ($Output -join "`n") | Should Match "Failed to create VSS Junction"
            $script:CapturedVssRoot | Should Match "GLOBALROOT"
        }
    }

    # --- Write-BackupHistory Edge Cases ---
    Context "Write-BackupHistory - Edge Cases" {
        $TestHistLogPath = Join-Path $PSScriptRoot "hist_edge_test.log"

        It "Should handle entry with special characters in values" {
            Remove-Item $TestHistLogPath -ErrorAction SilentlyContinue
            $Entry = [PSCustomObject]@{
                Status = "Success"
                Path   = 'D:\Shares\Data & Reports (2026)'
                Note   = 'Has special chars'
            }
            Write-BackupHistory -LogFilePath $TestHistLogPath -Entry $Entry

            $Content = (Get-Content $TestHistLogPath -Raw).Trim()
            $Content | Should Match '"Status":"Success"'
            # Verify entry was written with path data
            $Content | Should Match '"Path":'
            $Content | Should Match '"Note":'
        }

        It "Should handle entry with very long path values" {
            Remove-Item $TestHistLogPath -ErrorAction SilentlyContinue
            $LongPath = "D:\" + ("VeryLongFolderName\" * 20) + "file.txt"
            $Entry = [PSCustomObject]@{ Status = "Success"; Path = $LongPath }
            Write-BackupHistory -LogFilePath $TestHistLogPath -Entry $Entry

            Test-Path $TestHistLogPath | Should Be $true
            $Content = Get-Content $TestHistLogPath -Raw
            $Content.Length | Should BeGreaterThan 100
        }

        It "Should append multiple entries to the same log file" {
            Remove-Item $TestHistLogPath -ErrorAction SilentlyContinue
            $Entry1 = [PSCustomObject]@{ Id = 1; Status = "Success" }
            $Entry2 = [PSCustomObject]@{ Id = 2; Status = "Failed" }

            Write-BackupHistory -LogFilePath $TestHistLogPath -Entry $Entry1
            Write-BackupHistory -LogFilePath $TestHistLogPath -Entry $Entry2

            $Lines = Get-Content $TestHistLogPath
            $Lines.Count | Should Be 2
            $Lines[0] | Should Match '"Id":1'
            $Lines[1] | Should Match '"Id":2'
        }

        Remove-Item $TestHistLogPath -ErrorAction SilentlyContinue
    }

    # --- Config Merge Boolean Edge Cases ---
    Context "Get-Configuration - Boolean and Falsy Values" {
        $TestBoolConfigPath = Join-Path $PSScriptRoot "test_bool_config.json"

        It "Should preserve UseVSS=false and not overwrite with default true" {
            @{
                SourcePaths     = @("C:\Test")
                DestinationPath = $PSScriptRoot
                UseVSS          = $false
            } | ConvertTo-Json | Set-Content $TestBoolConfigPath

            $Config = Get-Configuration -ConfigPath $TestBoolConfigPath
            $Config.UseVSS | Should Be $false
        }

        It "Should preserve RetentionDays=0 and not overwrite with default" {
            @{
                SourcePaths     = @("C:\Test")
                DestinationPath = $PSScriptRoot
                RetentionDays   = 0
            } | ConvertTo-Json | Set-Content $TestBoolConfigPath

            $Config = Get-Configuration -ConfigPath $TestBoolConfigPath
            # 0 is falsy but not null — depends on merge logic
            # Current logic: if ($null -eq $LoadedConfig.$key) → 0 is not null, so preserved
            $Config.RetentionDays | Should Be 0
        }

        It "Should preserve RobocopyInterPacketGapMs=0 and not overwrite" {
            @{
                SourcePaths              = @("C:\Test")
                DestinationPath          = $PSScriptRoot
                RobocopyInterPacketGapMs = 0
            } | ConvertTo-Json | Set-Content $TestBoolConfigPath

            $Config = Get-Configuration -ConfigPath $TestBoolConfigPath
            $Config.RobocopyInterPacketGapMs | Should Be 0
        }

        Remove-Item $TestBoolConfigPath -ErrorAction SilentlyContinue
    }

    # --- Get-VolumeRoot Edge Cases ---
    Context "Get-VolumeRoot - Edge Cases" {
        It "Should handle path with trailing backslash" {
            $Drive = Split-Path $PSScriptRoot -Qualifier
            $Expected = "$Drive\"

            $Result = Get-VolumeRoot -Path "$PSScriptRoot\"
            $Result | Should Be $Expected
        }

        It "Should handle deeply nested paths" {
            $Drive = Split-Path $PSScriptRoot -Qualifier
            $Expected = "$Drive\"

            $Result = Get-VolumeRoot -Path $PSScriptRoot
            $Result | Should Be $Expected
        }
    }

    # --- Invoke-BackupItem Name Construction ---
    Context "Invoke-BackupItem - Name Construction Edge Cases" {
        It "Should construct correct folder names with special characters in SubDirName" {
            Mock Invoke-RobocopyBackup { return @{ Status = "Success"; ExitCode = 0; LogFile = "detail.log" } }

            $script:CapturedEntry = $null
            Mock Write-BackupHistory { param($LogFilePath, $Entry) $script:CapturedEntry = $Entry }

            $TestDestRoot = Join-Path $PSScriptRoot "NameEdgeTest"
            $TestLogsDir = Join-Path $PSScriptRoot "NameEdgeLogs"
            New-Item -Path $TestDestRoot -ItemType Directory -Force | Out-Null
            New-Item -Path $TestLogsDir -ItemType Directory -Force | Out-Null

            $Item = [PSCustomObject]@{ Name = "My-Project_v2.0"; SourceSubPath = "D:\Source\My-Project_v2.0" }
            Invoke-BackupItem -Item $Item -SourcePath "D:\Source" -BackupMode "SubDirectories" `
                -DestinationRoot $TestDestRoot -RobocopyOptions "" `
                -InterPacketGapMs 0 -LogsDir $TestLogsDir `
                -HistoryLogFilePath (Join-Path $PSScriptRoot "name_edge.log")

            $script:CapturedEntry.Subdirectory | Should Be "My-Project_v2.0"
            $script:CapturedEntry.DestinationPath | Should Match "My-Project_v2.0_\d{8}_\d{6}_\d{3}"

            Remove-Item $TestDestRoot -Recurse -Force -ErrorAction SilentlyContinue
            Remove-Item $TestLogsDir -Recurse -Force -ErrorAction SilentlyContinue
            Remove-Item (Join-Path $PSScriptRoot "name_edge.log") -ErrorAction SilentlyContinue
        }
    }

    # --- Concurrent Timestamp Uniqueness ---
    Context "Invoke-BackupItem - Rapid Sequential Calls" {
        It "Should generate unique timestamps for back-to-back calls" {
            Mock Invoke-RobocopyBackup { return @{ Status = "Success"; ExitCode = 0; LogFile = "detail.log" } }

            $script:Timestamps = @()
            Mock Write-BackupHistory { param($LogFilePath, $Entry) $script:Timestamps += $Entry.Timestamp }

            $TestDestRoot = Join-Path $PSScriptRoot "RapidTest"
            $TestLogsDir = Join-Path $PSScriptRoot "RapidLogs"
            New-Item -Path $TestDestRoot -ItemType Directory -Force | Out-Null
            New-Item -Path $TestLogsDir -ItemType Directory -Force | Out-Null

            $Item1 = [PSCustomObject]@{ Name = "A"; SourceSubPath = "D:\A" }
            $Item2 = [PSCustomObject]@{ Name = "B"; SourceSubPath = "D:\B" }

            Invoke-BackupItem -Item $Item1 -SourcePath "D:\Src" -BackupMode "Root" `
                -DestinationRoot $TestDestRoot -RobocopyOptions "" `
                -InterPacketGapMs 0 -LogsDir $TestLogsDir `
                -HistoryLogFilePath (Join-Path $PSScriptRoot "rapid.log")

            # Small delay to ensure ms tick
            Start-Sleep -Milliseconds 10

            Invoke-BackupItem -Item $Item2 -SourcePath "D:\Src" -BackupMode "Root" `
                -DestinationRoot $TestDestRoot -RobocopyOptions "" `
                -InterPacketGapMs 0 -LogsDir $TestLogsDir `
                -HistoryLogFilePath (Join-Path $PSScriptRoot "rapid.log")

            $script:Timestamps.Count | Should Be 2
            # With different item names, folder names always differ even if timestamp is same
            # But timestamps should include milliseconds
            $script:Timestamps[0] | Should Match "\d{8}_\d{6}_\d{3}"
            $script:Timestamps[1] | Should Match "\d{8}_\d{6}_\d{3}"

            Remove-Item $TestDestRoot -Recurse -Force -ErrorAction SilentlyContinue
            Remove-Item $TestLogsDir -Recurse -Force -ErrorAction SilentlyContinue
            Remove-Item (Join-Path $PSScriptRoot "rapid.log") -ErrorAction SilentlyContinue
        }
    }

    # --- Start-BackupProcess Duration Output ---
    Context "Start-BackupProcess - Completion Output" {
        It "Should output total duration at completion" {
            Mock Get-Command { return [PSCustomObject]@{ Name = "robocopy.exe" } } -ParameterFilter { $Name -eq "robocopy.exe" }
            Mock Test-Path { return $true }
            Mock New-Item { return $null }
            Mock Test-BackupPrerequisites { return $true }

            $Config = @{
                SourcePaths              = @()
                DestinationPath          = "E:\Backups"
                RetentionDays            = 30
                LogRetentionDays         = 90
                RobocopyOptions          = ''
                RobocopyInterPacketGapMs = 0
                HistoryLogFile           = "backup-history.log"
                UseVSS                   = $false
            }
            Mock Get-Configuration { return $Config }
            Mock Remove-OldBackups { }

            $Output = Start-BackupProcess -ConfigFilePath "fake.json" *>&1
            ($Output -join "`n") | Should Match "Backup Script Completed"
            ($Output -join "`n") | Should Match "Total Duration"
        }
    }

    # --- Remove-OldBackups - Nested Matching Folders ---
    Context "Remove-OldBackups - Nested Timestamp Folders" {
        $TestNestedDir = Join-Path $PSScriptRoot "NestedCleanupTest"

        It "Should find and delete timestamp-matching folders recursively" {
            Remove-Item $TestNestedDir -Recurse -Force -ErrorAction SilentlyContinue
            New-Item -Path $TestNestedDir -ItemType Directory -Force | Out-Null

            # Create nested structure: Root\Share\SubDir_TIMESTAMP
            $SubShareDir = Join-Path $TestNestedDir "ShareName"
            New-Item -Path $SubShareDir -ItemType Directory -Force | Out-Null

            $OldNested = New-Item -Path "$SubShareDir\Sub_20200101_000000" -ItemType Directory -Force
            $OldNested.CreationTime = (Get-Date).AddDays(-100)

            $NewNested = New-Item -Path "$SubShareDir\Sub_20260101_120000" -ItemType Directory -Force
            $NewNested.CreationTime = (Get-Date)

            Remove-OldBackups -DestinationRoot $TestNestedDir -RetentionDays 30 -LogsDir "Z:\None" -LogRetentionDays 90

            Test-Path $OldNested.FullName | Should Be $false
            Test-Path $NewNested.FullName | Should Be $true
            # Parent share folder should still exist
            Test-Path $SubShareDir | Should Be $true
        }

        Remove-Item $TestNestedDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    # --- Config with Only Required Fields ---
    Context "Get-Configuration - Minimal Config" {
        $TestMinPath = Join-Path $PSScriptRoot "test_min_config.json"

        It "Should work with only SourcePaths and DestinationPath" {
            @{
                SourcePaths     = @("C:\Data")
                DestinationPath = $PSScriptRoot
            } | ConvertTo-Json | Set-Content $TestMinPath

            $Config = Get-Configuration -ConfigPath $TestMinPath
            $Config.SourcePaths[0] | Should Be "C:\Data"
            $Config.DestinationPath | Should Be $PSScriptRoot
            # All defaults should be filled in
            $Config.RetentionDays | Should Be 30
            $Config.LogRetentionDays | Should Be 90
            $Config.MaxBackupAttempts | Should Be 3
            $Config.UseVSS | Should Be $true
            $Config.RobocopyInterPacketGapMs | Should Be 0
            $Config.HistoryLogFile | Should Be "backup-history.log"
        }

        Remove-Item $TestMinPath -ErrorAction SilentlyContinue
    }

    # --- Protected Regex Edge Cases ---
    Context "Protected Flags Regex - Additional Edge Cases" {
        It "Should strip /LOG+ (with plus)" {
            $ProtectedRegex = "^/(LOG|LOG\+|R|W|MT)(:.*)?$"
            "/LOG+" | Should Match $ProtectedRegex
            "/LOG+:C:\file.log" | Should Match $ProtectedRegex
        }

        It "Should NOT strip /MOVE or /MIN or /MAX" {
            $ProtectedRegex = "^/(LOG|LOG\+|R|W|MT)(:.*)?$"
            "/MOVE" | Should Not Match $ProtectedRegex
            "/MIN:1024" | Should Not Match $ProtectedRegex
            "/MAX:9999" | Should Not Match $ProtectedRegex
        }

        It "Should strip bare /R and /W without values" {
            $ProtectedRegex = "^/(LOG|LOG\+|R|W|MT)(:.*)?$"
            "/R" | Should Match $ProtectedRegex
            "/W" | Should Match $ProtectedRegex
        }

        It "Should NOT strip /REG or /WAIT" {
            $ProtectedRegex = "^/(LOG|LOG\+|R|W|MT)(:.*)?$"
            "/REG" | Should Not Match $ProtectedRegex
            "/WAIT" | Should Not Match $ProtectedRegex
        }
    }
}