# ==============================================================================
# DETECT: Physically check the file version of msvcp140.dll inside System32
# ==============================================================================
$DLLPath = "C:\Windows\System32\msvcp140.dll"
$NeedsRepair = $false
 
if (Test-Path $DLLPath) {
    # Extract the true file version metadata properties
    $FileVersionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($DLLPath)
    $RawVersionString = $FileVersionInfo.FileVersion
 
    # Use regex to isolate only the clean numeric dot notation (e.g., 14.00.24212.0)
    if ($RawVersionString -match '(\d+\.\d+\.\d+\.\d+)') {
        $CleanVersionString = $Matches[1]
        $ActualVersion = [version]$CleanVersionString
        # Modern 2015-2022 runtimes require a version greater than or equal to 14.30+
        # Windows Update / Printer drivers downgrade this file down to 14.13.x
        if ($ActualVersion -lt [version]"14.30.0.0") {
            Write-Warning "CRITICAL: msvcp140.dll version is downgraded to $CleanVersionString. Revit will crash."
            $NeedsRepair = $true
        } else {
            Write-Host "C++ Runtime file version is healthy ($CleanVersionString). Skipping repair to protect startup performance."
        }
    } else {
        # Fallback if the string pattern is unparseable directly by [version]
        Write-Warning "CRITICAL: msvcp140.dll has an unknown or unparseable format version string: '$RawVersionString'."
        $NeedsRepair = $true
    }
} else {
    Write-Warning "CRITICAL: msvcp140.dll is missing entirely from System32."
    $NeedsRepair = $true
}
 
# ==============================================================================
# REMEDIATE: Only runs if the physical file version is outdated or downgraded
# ==============================================================================
if ($NeedsRepair) {
    Write-Host "Initiating local repair loop..."
 
    $searchPatterns = @(
        "Microsoft Visual C++ 2015-2022 Redistributable (x64)*",
        "Microsoft Visual C++ v14 Redistributable (x64)*"
    )
 
    $registryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
 
    $apps = Get-ItemProperty $registryPaths | Where-Object {
        $name = $_.DisplayName
        foreach ($pattern in $searchPatterns) {
            if ($name -like $pattern) { return $true }
        }
        return $false
    }
 
    if ($apps) {
        foreach ($app in $apps) {
            $repairPath = $app.ModifyPath
            if ($repairPath) {
                # Clean up the arguments from the registry path string
                $exePath = ($repairPath -split " /modify")[0].Replace('"', '')
                if (Test-Path $exePath) {
                    Write-Host "Repairing: $($app.DisplayName)"
                    Start-Process -FilePath $exePath -ArgumentList "/repair /passive /norestart" -Wait -Verb RunAs
                    Write-Host "Repair complete: $($app.DisplayName)"
                }
                else {
                    Write-Error "Cached installer not found at '$exePath'."
                }
            }
        }
    }
    else {
        Write-Warning "Neither Visual C++ 2015-2022 nor v14 (x64) was found on this machine."
    }
}