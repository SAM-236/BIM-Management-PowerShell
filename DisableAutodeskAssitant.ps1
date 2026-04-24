<#
Toggles Autodesk Assistant add-in for Revit 2027+:
- .addin  -> .disabled
- .disabled -> .addin

Dry run:
  .\Toggle-RevitAssistantAddin.ps1 -WhatIf -Verbose

Apply:
  .\Toggle-RevitAssistantAddin.ps1 -Verbose
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [int]$StartYear = 2027,
    [int]$EndYear   = 2035,
    [string]$BasePath = "C:\Program Files\Autodesk",
    [string]$RelativeAssistantPath = "AddIns\Assistant",
    [string]$Stem = "Autodesk.Assistant.Application"
)

$targets = @()

for ($year = $StartYear; $year -le $EndYear; $year++) {
    $revitRoot = Join-Path $BasePath ("Revit {0}" -f $year)
    $assistantFolder = Join-Path $revitRoot $RelativeAssistantPath

    if (Test-Path -LiteralPath $assistantFolder) {
        $targets += [pscustomobject]@{
            Year            = $year
            AssistantFolder = $assistantFolder
            AddinPath       = Join-Path $assistantFolder ($Stem + ".addin")
            DisabledPath    = Join-Path $assistantFolder ($Stem + ".disabled")
        }
    }
}

if (-not $targets) {
    Write-Host "No Revit $StartYear+ Assistant folders found under $BasePath." -ForegroundColor Yellow
    return
}

foreach ($t in $targets) {
    $addin    = $t.AddinPath
    $disabled = $t.DisabledPath

    # Prefer toggling based on what exists; avoid overwriting anything.
    if (Test-Path -LiteralPath $addin) {

        if (Test-Path -LiteralPath $disabled) {
            Write-Verbose "[$($t.Year)] Both exist ($addin and $disabled). Not changing anything."
            continue
        }

        if ($PSCmdlet.ShouldProcess($addin, "Rename to $([System.IO.Path]::GetFileName($disabled))")) {
            Rename-Item -LiteralPath $addin -NewName ([System.IO.Path]::GetFileName($disabled)) -ErrorAction Stop
            Write-Host "[$($t.Year)] Disabled: $addin -> $disabled" -ForegroundColor Green
        }

    } elseif (Test-Path -LiteralPath $disabled) {

        if ($PSCmdlet.ShouldProcess($disabled, "Rename to $([System.IO.Path]::GetFileName($addin))")) {
            Rename-Item -LiteralPath $disabled -NewName ([System.IO.Path]::GetFileName($addin)) -ErrorAction Stop
            Write-Host "[$($t.Year)] Enabled:  $disabled -> $addin" -ForegroundColor Cyan
        }

    } else {
        Write-Verbose "[$($t.Year)] Neither .addin nor .disabled found in $($t.AssistantFolder)."
    }
}
