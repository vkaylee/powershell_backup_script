$SourcePath = "D:\tuanlee\snapshot_backup_script\workdirs"
$ResolvedPath = (Resolve-Path $SourcePath).Path
$Drive = Split-Path $ResolvedPath -Qualifier
$VolumeRoot = "$Drive\"
$RelativePath = $ResolvedPath.Substring($VolumeRoot.Length).TrimStart('\')

# Create a temporary snapshot
Write-Host "Creating temporary snapshot..."
$ShadowCopyClass = [wmiclass]"root\cimv2:Win32_ShadowCopy"
$Result = $ShadowCopyClass.Create($VolumeRoot, "ClientAccessible")

if ($Result.ReturnValue -eq 0) {
    $ShadowID = $Result.ShadowID
    $Snapshot = Get-WmiObject Win32_ShadowCopy | Where-Object { $_.ID -eq $ShadowID }
    $DevicePath = $Snapshot.DeviceObject
    
    Write-Host "Snapshot Created: $ShadowID"
    Write-Host "Device Path: $DevicePath"
    
    # 1. Try mklink (requires admin)
    # We use a temporary directory name
    $LinkPath = "C:\vss_temp_link"
    if (Test-Path $LinkPath) { Remove-Item $LinkPath -Force }
    
    Write-Host "Creating Symbolic Link to $DevicePath\ at $LinkPath"
    # Note: DevicePath usually needs trailing \ for mklink /d
    $Target = "$DevicePath\"
    cmd /c mklink /d "$LinkPath" "$Target"
    
    $FullMappedPath = Join-Path $LinkPath $RelativePath
    Write-Host "Mapped VSS Path: $FullMappedPath"
    
    if (Test-Path "$FullMappedPath") {
        Write-Host "[OK] Path exists via Symlink." -ForegroundColor Green
        Get-ChildItem -Path "$FullMappedPath" | Select-Object Name | Format-Table
    } else {
        Write-Host "[FAIL] Path still not found via Symlink." -ForegroundColor Red
    }

    # Clean up
    if (Test-Path $LinkPath) { Remove-Item $LinkPath -Force }
    Write-Host "`nRemoving Snapshot..."
    $Snapshot.Delete()
} else {
    Write-Host "Failed to create snapshot. Code: $($Result.ReturnValue)"
}