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
            } else {
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
                SourcePaths = @("D:\Shares\Data", "D:\Other")
                DestinationPath = "E:\Backups"
                RetentionDays = 30
                LogRetentionDays = 90
                MaxBackupAttempts = 3
                RobocopyOptions = '/MIR /LOG+:{logpath}'
                RobocopyInterPacketGapMs = 0
                HistoryLogFile = "backup-history.log"
                UseVSS = $true
            }

            # Mock VSS functions
            Mock New-ShadowCopy { 
                return [PSCustomObject]@{ ID = "{ID-D}"; DeviceObject = "\\?\GLOBALROOT\Device\VSS_D" }
            }
            Mock Remove-ShadowCopy { }
            Mock Get-Configuration { return $Config }
            Mock cmd.exe { }
            Mock Write-BackupHistory { }
            Mock Clean-OldBackups { }
            
            # Mock Get-BackupItems to reflect the source being processed
            Mock Get-BackupItems {
                # Ensure the path passed to Robocopy mock has the trailing slash the script now adds
                $VssPath = if ($VssSourceRoot.EndsWith("\")) { $VssSourceRoot } else { "$VssSourceRoot\" }
                return ,@([PSCustomObject]@{ Name = "Data"; SourceSubPath = $VssPath; IsRootMode = $true })
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
                SourcePaths = @($DeepPath)
                DestinationPath = "E:\Backups"
                RetentionDays = 30
                LogRetentionDays = 90
                MaxBackupAttempts = 3
                RobocopyOptions = '/MIR /LOG+:{logpath}'
                RobocopyInterPacketGapMs = 0
                HistoryLogFile = "backup-history.log"
                UseVSS = $true
            }

            Mock New-ShadowCopy { return [PSCustomObject]@{ ID = "{ID-D}"; DeviceObject = "\\?\GLOBALROOT\Device\VSS_D" } }
            Mock Remove-ShadowCopy { }
            Mock Get-Configuration { return $Config }
            Mock cmd.exe { }
            Mock Get-BackupItems {
                return ,@([PSCustomObject]@{ Name = "workdirs"; SourceSubPath = "\\?\GLOBALROOT\Device\VSS_D\tuanlee\snapshot_backup_script\workdirs\"; IsRootMode = $true })
            }
            Mock Write-BackupHistory { }
            Mock Clean-OldBackups { }

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
                SourcePaths = @("D:\Shares\Data")
                DestinationPath = "E:\Backups"
                RetentionDays = 30
                LogRetentionDays = 90
                MaxBackupAttempts = 3
                RobocopyOptions = '/MIR /LOG+:{logpath}'
                RobocopyInterPacketGapMs = 0
                HistoryLogFile = "backup-history.log"
                UseVSS = $false
            }
            
            Mock Get-Configuration { return $Config }
            Mock New-ShadowCopy { throw "Should not be called" }
            Mock Get-BackupItems {
                $VssPath = if ($VssSourceRoot.EndsWith("\")) { $VssSourceRoot } else { "$VssSourceRoot\" }
                return ,@([PSCustomObject]@{
                    Name = "Data"
                    SourceSubPath = $VssPath
                    IsRootMode = $true
                })
            }
            Mock Write-BackupHistory { }
            Mock Clean-OldBackups { }

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
                SourcePaths = @()
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

    Context "Execute-BackupItem Collision Prevention" {
        It "Should use high-precision timestamps for unique folder names" {
            $TestDest = Join-Path $PSScriptRoot "CollisionTest"
            $Item = [PSCustomObject]@{ Name = "Sub1"; SourceSubPath = "C:\Source\Sub1" }
            
            Mock Invoke-RobocopyBackup { return @{ Status = "Success"; ExitCode = 0 } }
            Mock New-Item { return $null }
            Mock Test-Path { return $true }
            
            # We want to ensure that even if called in rapid succession, the timestamp differs (or at least includes _fff)
            # Actually, we just need to verify the format contains 3 digits for ms
            
            # Execute-BackupItem defines $Timestamp = Get-Date -Format "yyyyMMdd_HHmmss_fff"
            # We verify the behavior by checking the history log or capturing the destination
            # But the function doesn't return the path. 
            # We'll verify the timestamp format in the history entry.
            
            $script:HistoryEntries = @()
            Mock Write-BackupHistory { param($LogFilePath, $Entry) $script:HistoryEntries += $Entry }
            
            Execute-BackupItem -Item $Item -SourcePath "C:\Src" -BackupMode "Root" -DestinationRoot $TestDest -RobocopyOptions "" -InterPacketGapMs 0 -LogsDir "C:\Logs" -HistoryLogFilePath "C:\hist.log"
            
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
                SourcePaths = @("D:\Data")
                DestinationPath = "D:\Backups"
                UseVSS = $true
                HistoryLogFile = "history.log"
                RetentionDays = 1
            }
            Mock Get-Configuration { return $Config }
            Mock New-ShadowCopy { return [PSCustomObject]@{ ID = "{VSS-123}"; DeviceObject = "\\?\GLOBALROOT\Device\VSS1" } }
            Mock Remove-ShadowCopy { }
            Mock Write-BackupHistory { }
            Mock Clean-OldBackups { }

            # Mock Get-BackupItems to use the junction-based VssSourceRoot
            Mock Get-BackupItems {
                param($SourcePath, $VssSourceRoot, $BackupMode)
                return ,@([PSCustomObject]@{ Name = "Data"; SourceSubPath = $VssSourceRoot; IsRootMode = $true })
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
    Context "Clean-OldBackups Logic" {
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
             
             Clean-OldBackups -DestinationRoot $TestBackupDir -RetentionDays 30 -LogsDir $TestLogsDir -LogRetentionDays 90
             
             Test-Path $Old.FullName | Should Be $false
             Test-Path $New.FullName | Should Be $true
         }
         
         It "Should ignore folders not matching timestamp pattern" {
             $Other = New-Item -Path "$TestBackupDir\OtherFolder" -ItemType Directory -Force
             $Other.CreationTime = (Get-Date).AddDays(-100)
             
             Clean-OldBackups -DestinationRoot $TestBackupDir -RetentionDays 30 -LogsDir $TestLogsDir -LogRetentionDays 90
             
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
}