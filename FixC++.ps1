# 1. Define the specific names to match both variations

$searchPatterns = @(
    "Microsoft Visual C++ 2015-2022 Redistributable (x64)*",
    "Microsoft Visual C++ v14 Redistributable (x64)*"
)
 
# 2. Search both 64-bit and 32-bit registry views

$registryPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)
 
# Find ALL apps that match any of our patterns

$apps = Get-ItemProperty $registryPaths | Where-Object {

    $name = $_.DisplayName

    foreach ($pattern in $searchPatterns) {if ($name -like $pattern) { return $true }

    }

    return $false

}
 
if ($apps) {

    foreach ($app in $apps) {

        $repairPath = $app.ModifyPath

        if ($repairPath) {

            # Strip the /modify flag and quotes to get the clean path

            $exePath = ($repairPath -split " /modify")[0].Replace('"', '')

            if (Test-Path $exePath) {

                Write-Host "Repairing: $($app.DisplayName)"

                Write-Host "Source: $exePath"
 
                # 3. Run the repair silently

                Start-Process -FilePath $exePath -ArgumentList "/repair /passive /norestart" -Wait -Verb RunAs

                Write-Host "Repair complete: $($app.DisplayName)"

                Write-Host ""

            }

            else { Write-Error "Cached installer not found at '$exePath' for $($app.DisplayName)."}

        }

        else { Write-Error "No ModifyPath found in the registry for $($app.DisplayName)."}

    }

}

else {Write-Warning "Neither Visual C++ 2015-2022 nor v14 (x64) was found on this machine."} 