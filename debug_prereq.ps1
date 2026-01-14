Function Test-BackupPrerequisites {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [object]$Config
    )

    $Passed = $true
    Write-Host "Running Pre-flight Diagnostic Checks..." -ForegroundColor Cyan

    # 1. Check for Robocopy
    if (-not (Get-Command "robocopy.exe" -ErrorAction SilentlyContinue)) {
        Write-Error "CRITICAL: 'robocopy.exe' was not found."
        $Passed = $false
    } else {
        Write-Host "  [OK] Robocopy detected." -ForegroundColor Green
    }

    return $Passed
}

$Config = @{ UseVSS = $true; DestinationPath = "C:\" }
$Res = Test-BackupPrerequisites -Config $Config
Write-Host "Result Type: $($Res.GetType().Name)"
Write-Host "Result Value: $Res"

if (-not $Res) {
    Write-Host "Check Failed"
} else {
    Write-Host "Check Passed"
}

