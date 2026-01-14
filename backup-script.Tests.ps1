$ScriptPath = Join-Path $PSScriptRoot "backup-script.ps1"
. $ScriptPath

Describe "Snapshot Backup Script - Extended Tests" {

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
            $HistoryContent = Get-Content $TestHistory
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
}