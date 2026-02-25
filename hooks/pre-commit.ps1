#Requires -Version 5.1
<#
.SYNOPSIS
    Pre-commit quality gate: lint, format, and test checks.
.DESCRIPTION
    Runs PSScriptAnalyzer (lint + formatting) and Pester tests on staged
    PowerShell files. Returns exit code 0 on success, 1 on failure.
    Called by .git/hooks/pre-commit.
#>

$ErrorActionPreference = "Continue"
$ScriptRoot = Split-Path -Parent (Split-Path -Parent $PSCommandPath)  # repo root
Set-Location $ScriptRoot

$Failed = $false

# ============================================================
# 1. PSScriptAnalyzer 鈥?Lint
# ============================================================
Write-Host "`n--- [1/3] PSScriptAnalyzer Lint ---" -ForegroundColor Cyan

if (-not (Get-Module PSScriptAnalyzer -ListAvailable)) {
    Write-Host "PSScriptAnalyzer not installed. Installing..." -ForegroundColor Yellow
    Install-Module PSScriptAnalyzer -Force -Scope CurrentUser -SkipPublisherCheck | Out-Null
}

Import-Module PSScriptAnalyzer -ErrorAction Stop

$ScriptFiles = @(
    "backup-script.ps1"
)

foreach ($File in $ScriptFiles) {
    $FilePath = Join-Path $ScriptRoot $File
    if (-not (Test-Path $FilePath)) { continue }

    Write-Host "  Analyzing: $File" -ForegroundColor Gray
    $Results = Invoke-ScriptAnalyzer -Path $FilePath -Severity @("Error", "Warning") -ExcludeRule @(
        "PSAvoidUsingWriteHost"           # Expected in a CLI script
        "PSUseShouldProcessForStateChangingFunctions"  # Not needed for backup script
        "PSUseSingularNouns"              # Intentional: Get-BackupItems, Remove-OldBackups, Test-BackupPrerequisites
        "PSAvoidUsingWMICmdlet"           # Required for VSS 鈥?PS 5.1 compatibility
        "PSUseBOMForUnicodeEncodedFile"   # UTF-8 without BOM is fine for modern systems
    )

    if ($Results.Count -gt 0) {
        $Failed = $true
        Write-Host "  FAIL: $($Results.Count) issue(s) found in $File" -ForegroundColor Red
        $Results | ForEach-Object {
            Write-Host "    [$($_.Severity)] Line $($_.Line): $($_.Message)" -ForegroundColor Yellow
            Write-Host "      Rule: $($_.RuleName)" -ForegroundColor DarkGray
        }
    }
    else {
        Write-Host "  PASS: No issues." -ForegroundColor Green
    }
}

# ============================================================
# 2. Formatting 鈥?Brace Style & Indentation
# ============================================================
Write-Host "`n--- [2/3] Formatting Checks ---" -ForegroundColor Cyan

foreach ($File in $ScriptFiles) {
    $FilePath = Join-Path $ScriptRoot $File
    if (-not (Test-Path $FilePath)) { continue }

    Write-Host "  Checking: $File" -ForegroundColor Gray

    $FormatRules = @(
        "PSUseConsistentIndentation"
        "PSUseConsistentWhitespace"
        "PSPlaceOpenBrace"
        "PSPlaceCloseBrace"
    )

    $FormatResults = Invoke-ScriptAnalyzer -Path $FilePath -IncludeRule $FormatRules -Settings @{
        IncludeDefaultRules = $false
        Rules               = @{
            PSUseConsistentIndentation = @{
                Enable          = $true
                IndentationSize = 4
                Kind            = "space"
            }
            PSUseConsistentWhitespace  = @{
                Enable                          = $true
                CheckOpenBrace                  = $true
                CheckOpenParen                  = $true
                CheckOperator                   = $false  # Too noisy for alignment
                CheckSeparator                  = $true
                CheckInnerBrace                 = $true
                CheckPipeForRedundantWhitespace = $true
            }
            PSPlaceOpenBrace           = @{
                Enable             = $true
                OnSameLine         = $true
                NewLineAfter       = $true
                IgnoreOneLineBlock = $true
            }
            PSPlaceCloseBrace          = @{
                Enable             = $true
                NewLineAfter       = $true
                IgnoreOneLineBlock = $true
                NoEmptyLineBefore  = $false
            }
        }
    }

    if ($FormatResults.Count -gt 0) {
        # Formatting issues are warnings, not blockers
        Write-Host "  WARN: $($FormatResults.Count) formatting suggestion(s)" -ForegroundColor Yellow
        $FormatResults | Select-Object -First 5 | ForEach-Object {
            Write-Host "    Line $($_.Line): $($_.Message)" -ForegroundColor DarkYellow
        }
        if ($FormatResults.Count -gt 5) {
            Write-Host "    ... and $($FormatResults.Count - 5) more" -ForegroundColor DarkGray
        }
    }
    else {
        Write-Host "  PASS: Formatting OK." -ForegroundColor Green
    }
}

# ============================================================
# 3. Pester Tests
# ============================================================
Write-Host "`n--- [3/3] Pester Tests ---" -ForegroundColor Cyan

$TestFile = Join-Path $ScriptRoot "backup-script.Tests.ps1"
if (Test-Path $TestFile) {
    $PesterResult = Invoke-Pester -Path $TestFile -PassThru -Quiet 6>$null 5>$null 4>$null 3>$null 2>$null

    if ($PesterResult.FailedCount -gt 0) {
        $Failed = $true
        Write-Host "  FAIL: $($PesterResult.FailedCount)/$($PesterResult.TotalCount) tests failed" -ForegroundColor Red
        $PesterResult.TestResult | Where-Object { $_.Result -eq 'Failed' } | ForEach-Object {
            Write-Host "    FAIL: $($_.Name)" -ForegroundColor Red
            Write-Host "      $($_.FailureMessage)" -ForegroundColor DarkYellow
        }
    }
    else {
        Write-Host "  PASS: $($PesterResult.PassedCount)/$($PesterResult.TotalCount) tests passed." -ForegroundColor Green
    }
}
else {
    Write-Host "  SKIP: No test file found." -ForegroundColor Yellow
}

# ============================================================
# Result
# ============================================================
Write-Host ""
if ($Failed) {
    Write-Host "=== QUALITY GATE FAILED ===" -ForegroundColor Red
    exit 1
}
else {
    Write-Host "=== ALL CHECKS PASSED ===" -ForegroundColor Green
    exit 0
}

