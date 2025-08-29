#This tool is ment to clean up all collaberation cahce projects older than a sepcific number of days of access to the locla central file. 
#The File and its resources are removed so that Revit will repull all resources.
#log file created inorder to allow for review of disk space saved and guid of the rvt cloud file name to match to the acc url if needed. 

# Example call for Task Scheduler
Cleanup-CollaborationCache -DaysBack 120 -Log $true

function Cleanup-CollaborationCache {
    param (
        [int]$DaysBack = 120,
        [bool]$Log = $false
    )

    # Ensure the C:\Temp directory exists
    $LogDirectory = "C:\Temp"
    if (!(Test-Path -Path $LogDirectory)) {
        New-Item -ItemType Directory -Path $LogDirectory | Out-Null
    }

    $LogFile = Join-Path -Path $LogDirectory -ChildPath "$(Get-Date -Format "FileDateTime")_CleanupLog.html"
    $logContent = New-Object System.Collections.ArrayList

    function Add-Log {
        param ([string]$message)
        if ($Log) {
            $logContent.Add("<p>$message</p>") | Out-Null
        }
        Write-Output $message
    }

    # Start time logging
    $startTime = Get-Date
    Add-Log "Cleanup process started at $startTime"

    $runningRevitVersions = Get-RunningRevitVersions
    if ($runningRevitVersions.Count -eq 0) {
        Add-Log "No Revit instances are running. Proceeding with cleanup..."
    } else {
        Add-Log "The following Revit versions are running: $($runningRevitVersions -join ', ')"
    }

    $targetDate = (Get-Date).AddDays(-$DaysBack)
    $years = 2019..2025

    $totalFilesDeleted = 0
    $totalLayer2FoldersDeleted = 0
    $totalSizeDeletedBytes = 0

    $userDirs = Get-ChildItem -Path "C:\Users" -Directory

    foreach ($userDir in $userDirs) {
        if ($userDir.Name -notin @("Public", "Default", "Default User", "All Users")) {
            foreach ($year in $years) {
                if ($runningRevitVersions -contains $year) {
                    Add-Log "Skipping cleanup for Revit $year because it is currently running."
                    continue
                }

                $collabCacheBase = "$($userDir.FullName)\AppData\Local\Autodesk\Revit\Autodesk Revit $year\CollaborationCache"

                if (Test-Path $collabCacheBase) {
                    $guidFoldersLevel1 = Get-ChildItem -Path $collabCacheBase -Directory -ErrorAction SilentlyContinue

                    foreach ($guidFolder1 in $guidFoldersLevel1) {
                        $guidFoldersLevel2 = Get-ChildItem -Path $guidFolder1.FullName -Directory -ErrorAction SilentlyContinue

                        foreach ($guidFolder2 in $guidFoldersLevel2) {
                            $centralCachePath = Join-Path -Path $guidFolder2.FullName -ChildPath "CentralCache"
                            $shouldDelete = $true

                            if (Test-Path $centralCachePath) {
                                $filesInCentralCache = Get-ChildItem -Path $centralCachePath -File -Recurse -Force -ErrorAction SilentlyContinue
                                if ($filesInCentralCache | Where-Object { $_.LastWriteTime -ge $targetDate }) {
                                    $shouldDelete = $false
                                }
                            }

                            $filesInGuidFolder2 = Get-ChildItem -Path $guidFolder2.FullName -File -Recurse -Force -ErrorAction SilentlyContinue
                            if ($filesInGuidFolder2 | Where-Object { $_.LastWriteTime -ge $targetDate }) {
                                $shouldDelete = $false
                            }

                            if ($shouldDelete) {
                                $filesToDelete = Get-ChildItem -Path $guidFolder2.FullName -Recurse -File -Force -ErrorAction SilentlyContinue
                                $sizeToDelete = ($filesToDelete | Measure-Object -Property Length -Sum).Sum
                                $totalFilesDeleted += $filesToDelete.Count
                                $totalSizeDeletedBytes += $sizeToDelete

                                Add-Log "Deleting folder: $($guidFolder2.FullName)"
                                Remove-Item -Path $guidFolder2.FullName -Recurse -Force

                                if (!(Test-Path $guidFolder2.FullName)) {
                                    Add-Log "Deletion confirmed: $($guidFolder2.FullName) has been deleted."
                                    $totalLayer2FoldersDeleted++
                                } else {
                                    Add-Log "WARNING: Deletion failed for $($guidFolder2.FullName)."
                                }
                            } else {
                                Add-Log "Skipping folder: $($guidFolder2.FullName) (Files modified within time window)"
                            }
                        }
                    }
                }
            }
        }
    }

    $totalSizeDeletedGB = [math]::Round($totalSizeDeletedBytes / 1GB, 2)
    $endTime = Get-Date

    Add-Log "<h2>Summary of Deletions:</h2>"
    Add-Log "Total files deleted: $totalFilesDeleted"
    Add-Log "Total Layer 2 folders deleted: $totalLayer2FoldersDeleted"
    Add-Log "Total space freed: $totalSizeDeletedGB GB"
    Add-Log "Cleanup process completed at $endTime"

    if ($Log) {
        $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>Cleanup Report</title>
</head>
<body>
    $(($logContent -join "`n"))
</body>
</html>
"@

        $htmlContent | Out-File -FilePath $LogFile -Encoding utf8
        Write-Output "Log saved to $LogFile"
    }

    Write-Output "Cleanup process finished successfully. Report saved at $LogFile."
}



# Function to get running Revit versions
function Get-RunningRevitVersions {
    $runningRevitVersions = @()
    $revitProcesses = Get-Process -Name "Revit" -ErrorAction SilentlyContinue
    if ($revitProcesses) {
        foreach ($process in $revitProcesses) {
            try {
                $filePath = $process.Path
                $fileVersion = (Get-Item $filePath).VersionInfo.ProductVersion
                $match = [regex]::Match($fileVersion, "20\d{2}")
                if ($match.Success) {
                    $year = [int]$match.Value
                    $runningRevitVersions += $year
                    Write-Output "Revit $year is running (PID: $($process.Id))"
                }
            }
            catch {
                Write-Warning "Unable to determine version for process PID: $($process.Id)"
            }
        }
    }
    return $runningRevitVersions
}