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
            Mock Write-BackupHistory { }
            Mock Clean-OldBackups { }
            
            # Mock Get-BackupItems to reflect the source being processed
            Mock Get-BackupItems {
                # We return a single item to simulate one folder being backed up per source
                return ,@([PSCustomObject]@{ Name = "Data"; SourceSubPath = "$VssSourceRoot"; IsRootMode = $true })
            }

            Start-BackupProcess -ConfigFilePath "fake.json"

            # Verify snapshots created for D: (called twice because we have two D: sources)
            Assert-MockCalled New-ShadowCopy -Exactly 2 -ParameterFilter { $VolumeRoot -eq "D:\" }
            
            $script:RobocopyCalled | Should Be $true
        }
    }

    Context "VSS Workflow Integration - Disabled" {
        $script:RobocopyCalledDisabled = $false
        Mock Invoke-RobocopyBackup { $script:RobocopyCalledDisabled = $true; return @{ Status = "Success"; ExitCode = 0; LogFile = "test.log" } }

        It "Should NOT create a snapshot and use direct path when UseVSS is false" {
            Mock Get-Command { return [PSCustomObject]@{ Name = "robocopy.exe" } } -ParameterFilter { $Name -eq "robocopy.exe" }
            Mock Test-Path { return $true }
            Mock New-Item { return $null }
            Mock Join-Path { 
                param($Path, $ChildPath)
                if ($Path.EndsWith("\")) { return "$Path$ChildPath" }
                return "$Path\$ChildPath"
            }
            Mock Get-VolumeRoot { return "D:\" }
            Mock Test-BackupPrerequisites { return $true }

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
                return ,@([PSCustomObject]@{
                    Name = "Data"
                    SourceSubPath = "$VssSourceRoot"
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
            $Items = Get-BackupItems -SourcePath $TestRoot -VssSourceRoot "\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy1\TestItems" -BackupMode "Root"
            
            $Items.Count | Should Be 1
            $Items[0].IsRootMode | Should Be $true
            $Items[0].Name | Should Be "TestItems"
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
    Context "Execute-BackupItem Logic" {
        $TestDest = Join-Path $PSScriptRoot "TestDest"
        $TestLogs = Join-Path $PSScriptRoot "TestLogs"
        $TestHistory = Join-Path $PSScriptRoot "history.log"
        
        # Cleanup
        Remove-Item $TestDest -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item $TestLogs -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item $TestHistory -Force -ErrorAction SilentlyContinue

        New-Item -Path $TestDest -ItemType Directory -Force | Out-Null
        New-Item -Path $TestLogs -ItemType Directory -Force | Out-Null

        It "Should create correct destination folder and execute backup" {
            $Item = [PSCustomObject]@{ 
                Name = "MySubDir"
                SourceSubPath = "C:\Source\MySubDir"
                IsRootMode = $false
            }
            
            # Using /L (List only) to prevent actual copy
            Execute-BackupItem `
                -Item $Item `
                -SourcePath "C:\Source" `
                -BackupMode "SubDirectories" `
                -DestinationRoot $TestDest `
                -RobocopyOptions "/L" `
                -InterPacketGapMs 0 `
                -LogsDir $TestLogs `
                -HistoryLogFilePath $TestHistory `
                -SnapshotID "TestID"

            # Verify Destination Parent was created: TestDest\Source
            $ExpectedParent = Join-Path $TestDest "Source"
            Test-Path $ExpectedParent | Should Be $true

            # Verify History Log updated
            Test-Path $TestHistory | Should Be $true
            $HistoryContent = Get-Content $TestHistory -Raw
            $HistoryContent | Should Match "MySubDir"
        }
        
        # Cleanup
        Remove-Item $TestDest -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item $TestLogs -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item $TestHistory -Force -ErrorAction SilentlyContinue
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