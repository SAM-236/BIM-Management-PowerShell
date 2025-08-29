#If Access is running Close and remove the db then relaunch this will allow the software to rescan the PC after a few min to force it to find missing updates that are not displaying.

# Define the process name, file path, and executable path
$processName = "Adsk.AccessCore"
$accessCoreDBPath = "C:\ProgramData\Autodesk\ODIS\LocalCache.db"
$accessCoreExePath = "C:\Program Files\Autodesk\AdODIS\V1\Access\AdskAccessCore.exe"

# Function to check the process status
function Check-Process {
    param (
        [string]$name
    )
    Get-Process | Where-Object {$_.Name -eq $name} | Select-Object -ExpandProperty Id -ErrorAction SilentlyContinue
}

# Function to clear the specified file
function Clear-File {
    param (
        [string]$path
    )
    if (Test-Path $path) {
        Remove-Item $path -Force
    }
}

# Function to restart the process using a fixed path
function Restart-Process {
    Start-Process $accessCoreExePath
}

# Main script logic
Write-Host "Monitoring process: $processName"
do {
    Start-Sleep -Seconds 10 # Check every 10 seconds
    $processId = Check-Process -name $processName
    if (-not $processId) {
        Write-Host "Process $processName ended. Clearing file and restarting process..."
        Clear-File -path $accessCoreDBPath
        Restart-Process
        Write-Host "Process restarted. Monitoring again."
        break
    }
} while ($true) # Infinite loop, press Ctrl+C to exit