try {
    $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    Write-Host "Success: $IsAdmin"
} catch {
    Write-Host "Error: $($_.Exception.Message)"
}
