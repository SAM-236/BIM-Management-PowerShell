param(
    [string]$IconRoot = "$env:ProgramData\RevitIcons",
    [string]$LogDir = "$env:ProgramData\Logs",
    [string]$LogFileName = "RevitIconUpdate.log",
    [int[]]$Years = (2022..2027)
)

$iconRoot = $IconRoot
$years = $Years

# -----------------------------
# Logging Setup
# -----------------------------
$logFile = Join-Path $LogDir $LogFileName

try {
    if (!(Test-Path $LogDir)) {
        New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
    }
} catch {}

# Optional log rotation (5 MB)
try {
    if ((Test-Path $logFile) -and ((Get-Item $logFile).Length -gt 5MB)) {
        Rename-Item $logFile "$logFile.old" -Force
    }
} catch {}

function Write-Log {
    param ([string]$Message)

    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $entry = "$timestamp [$env:COMPUTERNAME][$env:USERNAME] $Message"

    try { Add-Content -Path $logFile -Value $entry } catch {}
    Write-Host $entry
}

Write-Log "==== Starting Revit Icon Update ===="
Write-Log "Icon Root: $iconRoot"
Write-Log "Log File: $logFile"

# Validate IconRoot
try {
    $iconRoot = (Resolve-Path $iconRoot).Path
} catch {
    Write-Log "ERROR: Invalid IconRoot → $IconRoot"
    exit 1
}

# -----------------------------
# Shortcut Paths
# -----------------------------
$paths = @(
    "$env:PUBLIC\Desktop",
    "$env:USERPROFILE\Desktop",
    "$env:APPDATA\Microsoft\Windows\Start Menu\Programs",
    "$env:ProgramData\Microsoft\Windows\Start Menu\Programs",
    "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
)

# Create shell object
try {
    $shell = New-Object -ComObject WScript.Shell
} catch {
    Write-Log "ERROR: Failed to create COM object"
    exit 1
}

# -----------------------------
# Shortcut Icon Updates
# -----------------------------
foreach ($path in $paths) {

    if (!(Test-Path $path)) {
        Write-Log "Skipping missing folder: $path"
        continue
    }

    Write-Log "Processing path: $path"

    $shortcuts = Get-ChildItem -Path $path -Filter "*.lnk" -Recurse -ErrorAction SilentlyContinue

    if (!$shortcuts) {
        Write-Log "No shortcuts found in: $path"
        continue
    }

    foreach ($lnk in $shortcuts) {

        if (!(Test-Path $lnk.FullName)) {
            Write-Log "Missing shortcut: $($lnk.FullName)"
            continue
        }

        try {
            $shortcut = $shell.CreateShortcut($lnk.FullName)
        }
        catch {
            Write-Log "Invalid shortcut: $($lnk.FullName)"
            continue
        }

        if (-not $shortcut) { continue }

        $matchedYear = $null

        foreach ($year in $years) {
            if (
                $lnk.Name -match "(?i)revit.*$year" -or
                ($shortcut.TargetPath -and $shortcut.TargetPath -match "(?i)revit.*$year")
            ) {
                $matchedYear = $year
                break
            }
        }

        if (-not $matchedYear) { continue }

        $iconPath = Join-Path $iconRoot "r$matchedYear.ico"

        if (!(Test-Path $iconPath)) {
            Write-Log "Missing icon file: $iconPath"
            continue
        }

        try {
            $newIcon = "$iconPath,0"

            if ($shortcut.IconLocation -ne $newIcon -or $shortcut.IconLocation -match "Windows\\Installer") {
                $shortcut.IconLocation = $newIcon
                $shortcut.Save()

                Write-Log "UPDATED → $($lnk.FullName) → r$matchedYear.ico"
            }
            else {
                Write-Log "UNCHANGED → $($lnk.FullName)"
            }
        }
        catch {
            Write-Log "FAILED → $($lnk.FullName)"
        }
    }
}

# -----------------------------
# MSI Cache Icon Override
# -----------------------------
Write-Log "Starting MSI Installer cache icon override..."

$installerRoot = "$env:WINDIR\Installer"

if (!(Test-Path $installerRoot)) {
    Write-Log "Installer cache not found"
}
else {

    foreach ($year in $years) {

        $iconSource = Join-Path $iconRoot "r$year.ico"

        if (!(Test-Path $iconSource)) {
            Write-Log "Skipping $year — missing icon"
            continue
        }

        $pattern = "*$($year.ToString().Substring(2,2))00-0510-0000-705C0D862004*"

        $folders = Get-ChildItem $installerRoot -Directory -ErrorAction SilentlyContinue |
                   Where-Object { $_.Name -like $pattern }

        if (!$folders) {
            Write-Log "No installer folder for Revit $year"
            continue
        }

        foreach ($folder in $folders) {

            $revIcon = Join-Path $folder.FullName "REVT.ico"

            if (!(Test-Path $revIcon)) {
                Write-Log "REVT.ico missing in $($folder.Name)"
                continue
            }

            try {
                Copy-Item -Path $iconSource -Destination $revIcon -Force

                Write-Log "UPDATED MSI ICON → $($folder.Name)\REVT.ico"
            }
            catch {
                Write-Log "FAILED MSI ICON → $($folder.Name)"
            }
        }
    }
}

# -----------------------------
# Refresh icon cache
# -----------------------------
try {
    $ie4uinit = "$env:WINDIR\System32\ie4uinit.exe"

    if (Test-Path $ie4uinit) {
        & $ie4uinit -show
        Write-Log "Icon cache refreshed"
    }
}
catch {
    Write-Log "Icon refresh failed"
}

Write-Log "==== Completed Revit Icon Update ===="