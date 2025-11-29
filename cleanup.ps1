# cleanup.ps1
# PowerShell version of cleanup.vbs
# Cleans old version files, empty folders, and old log files

param(
    [string]$ConfigPath
)

# Script-scoped variables
$script:logContent = ""
$script:scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:logPath = Join-Path $script:scriptPath "log"
$script:statsFilesDeleted = 0
$script:statsFoldersDeleted = 0
$script:statsLogsDeleted = 0

# Ensure log directory exists
if (-not (Test-Path $script:logPath)) {
    New-Item -ItemType Directory -Path $script:logPath -Force | Out-Null
}

function Save-Log {
    $now = Get-Date
    $logFileName = $now.ToString("yyyyMMdd-HHmmss") + "-cleanup.log"
    $logFilePath = Join-Path $script:logPath $logFileName
    
    # Write log with UTF-8 encoding
    [System.IO.File]::WriteAllText($logFilePath, $script:logContent, [System.Text.Encoding]::UTF8)
}

function Write-FatalError {
    param([string]$message)
    
    $script:logContent += "`r`nFATAL ERROR $message"
    Save-Log
    exit 1
}

function Remove-OldVersions {
    param(
        [string]$verPath,
        [string]$originalFileName
    )
    
    # Keep only the 3 most recent versions of a file
    if (-not (Test-Path $verPath)) {
        return
    }
    
    # Collect all version files for this filename
    $versionFiles = Get-ChildItem -Path $verPath -File -ErrorAction SilentlyContinue | Where-Object {
        # Match files with YYYYMMDD_ prefix followed by the original filename
        if ($_.Name.EndsWith($originalFileName) -and 
            $_.Name.Length -ge (9 + $originalFileName.Length) -and
            $_.Name.Substring(8, 1) -eq "_") {
            
            $datePrefix = $_.Name.Substring(0, 8)
            return $datePrefix -match '^\d{8}$'
        }
        return $false
    }
    
    # Sort by date (newest first) and keep only 3 most recent
    $sortedFiles = $versionFiles | Sort-Object LastWriteTime -Descending
    
    if ($sortedFiles.Count -gt 3) {
        $filesToDelete = $sortedFiles | Select-Object -Skip 3
        
        foreach ($file in $filesToDelete) {
            try {
                Remove-Item $file.FullName -Force -ErrorAction Stop
                $script:logContent += "[DELETE] $($file.FullName)`r`n"
                $script:statsFilesDeleted++
            }
            catch {
                $script:logContent += "[ERROR] Failed to delete $($file.FullName): $($_.Exception.Message)`r`n"
            }
        }
    }
}

function Remove-EmptyFoldersRecursive {
    param([string]$folderPath)
    
    # Recursively remove empty folders (bottom-up approach)
    if (-not (Test-Path $folderPath)) {
        return
    }
    
    try {
        # First, process all subfolders recursively
        $subfolders = Get-ChildItem -Path $folderPath -Directory -ErrorAction SilentlyContinue
        
        foreach ($subfolder in $subfolders) {
            Remove-EmptyFoldersRecursive $subfolder.FullName
        }
        
        # After processing subfolders, check if this folder is now empty
        $items = Get-ChildItem -Path $folderPath -Force -ErrorAction SilentlyContinue
        
        if ($null -eq $items -or $items.Count -eq 0) {
            Remove-Item $folderPath -Force -ErrorAction Stop
            $script:logContent += "[RMDIR] $folderPath`r`n"
            $script:statsFoldersDeleted++
        }
    }
    catch {
        $script:logContent += "[ERROR] Failed to process folder $folderPath`: $($_.Exception.Message)`r`n"
    }
}

function Remove-OldLogFiles {
    # Remove log files older than 30 days
    if (-not (Test-Path $script:logPath)) {
        return
    }
    
    $cutoffDate = (Get-Date).AddDays(-30)
    
    try {
        $logFiles = Get-ChildItem -Path $script:logPath -Filter "*.log" -File -ErrorAction SilentlyContinue
        
        foreach ($logFile in $logFiles) {
            if ($logFile.LastWriteTime -lt $cutoffDate) {
                $daysOld = [math]::Floor((New-TimeSpan -Start $logFile.LastWriteTime -End (Get-Date)).TotalDays)
                
                try {
                    Remove-Item $logFile.FullName -Force -ErrorAction Stop
                    $script:logContent += "[DELETE_LOG] $($logFile.Name) (age: $daysOld days)`r`n"
                    $script:statsLogsDeleted++
                }
                catch {
                    $script:logContent += "[ERROR] Failed to delete log $($logFile.Name): $($_.Exception.Message)`r`n"
                }
            }
        }
    }
    catch {
        $script:logContent += "[ERROR] Failed to process log files: $($_.Exception.Message)`r`n"
    }
}

function Clear-VersionDirectory {
    param([string]$verPath)
    
    # Recursively clean all subdirectories in .ver folder
    if (-not (Test-Path $verPath)) {
        return
    }
    
    try {
        # Create hashtable to track unique base filenames
        $uniqueFiles = @{}
        
        # Collect all version files and extract original filenames
        $files = Get-ChildItem -Path $verPath -File -ErrorAction SilentlyContinue
        
        foreach ($file in $files) {
            $fileName = $file.Name
            
            # Check if it matches version pattern: YYYYMMDD_filename.ext
            if ($fileName.Length -ge 10 -and 
                $fileName.Substring(8, 1) -eq "_" -and 
                $fileName.Substring(0, 8) -match '^\d{8}$') {
                
                $originalName = $fileName.Substring(9) # Everything after YYYYMMDD_
                
                if (-not $uniqueFiles.ContainsKey($originalName)) {
                    $uniqueFiles[$originalName] = $true
                }
            }
        }
        
        # Clean versions for each unique filename
        foreach ($uniqueFile in $uniqueFiles.Keys) {
            Remove-OldVersions $verPath $uniqueFile
        }
        
        # Recursively process subdirectories
        $subdirs = Get-ChildItem -Path $verPath -Directory -ErrorAction SilentlyContinue
        
        foreach ($subdir in $subdirs) {
            Clear-VersionDirectory $subdir.FullName
        }
    }
    catch {
        $script:logContent += "[ERROR] Failed to process version directory $verPath`: $($_.Exception.Message)`r`n"
    }
}

# Main execution
$script:logContent += "Cleanup started at: $(Get-Date)`r`n`r`n"
$script:logContent += "Working directory: $script:scriptPath`r`n"

# Determine config path
if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Join-Path $script:scriptPath "backup.ini"
}

$script:logContent += "Config file: $ConfigPath`r`n"

if (-not (Test-Path $ConfigPath)) {
    Write-FatalError "Configuration file not found"
}

# Read target path from config
$targetPath = ""

try {
    $configContent = Get-Content $ConfigPath -Encoding UTF8 -ErrorAction Stop
    
    foreach ($line in $configContent) {
        $line = $line.Trim()
        
        # Skip empty lines, comments, and section headers
        if ($line.Length -eq 0 -or $line.StartsWith(";") -or $line.StartsWith("[")) {
            continue
        }
        
        # Parse key=value pairs
        if ($line -match '^([^=]+)=(.*)$') {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            
            if ($key -eq "targetPath" -and $value -ne "") {
                $targetPath = $value
                break
            }
        }
    }
}
catch {
    Write-FatalError "Failed to read configuration file: $($_.Exception.Message)"
}

if ($targetPath -eq "") {
    Write-FatalError "Target folder not defined in the configuration file"
}

$script:logContent += "Target path: $targetPath`r`n"

# Validate target path
$driveName = Split-Path -Qualifier $targetPath

if (-not (Test-Path $driveName)) {
    Write-FatalError "Target drive is not available: $driveName"
}

if (-not (Test-Path $targetPath)) {
    Write-FatalError "Target path does not exist: $targetPath"
}

# Check if .ver directory exists
$verRootPath = Join-Path $targetPath ".ver"

if (-not (Test-Path $verRootPath)) {
    $script:logContent += "No .ver directory found. Nothing to clean.`r`n"
    Save-Log
    exit 0
}

$script:logContent += "`r`n=-=-=-=-=-=-=-=-=-="
$script:logContent += "`r`n||    CLEANUP    ||"
$script:logContent += "`r`n=-=-=-=-=-=-=-=-=-=`r`n`r`n"

# Clean version directory
$script:logContent += "Cleaning version directory: $verRootPath`r`n"
Clear-VersionDirectory $verRootPath

$script:logContent += "`r`nRemoving empty folders...`r`n"
Remove-EmptyFoldersRecursive $verRootPath

$script:logContent += "`r`nCleaning old log files (older than 30 days)...`r`n"
Remove-OldLogFiles

# Statistics
$script:logContent += "`r`n===== STATISTICS =====`r`n"
$script:logContent += "Deleted old versions: $script:statsFilesDeleted`r`n"
$script:logContent += "Removed empty folders: $script:statsFoldersDeleted`r`n"
$script:logContent += "Deleted old log files: $script:statsLogsDeleted`r`n`r`n"

$script:logContent += "`r`nSUCCESS`r`n"
$script:logContent += "Cleanup completed at: $(Get-Date)`r`n"

Save-Log
