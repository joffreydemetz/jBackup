# jBackup - PowerShell Backup Script
# Automatic backup system with file versioning

param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "backup.ini")
)

$script:logContent = ""
$script:scriptPath = $PSScriptRoot
$script:invalidNamesPath = Join-Path $PSScriptRoot "invalidnames.txt"
$script:logPath = Join-Path $PSScriptRoot "log"
$script:targetPath = ""
$script:moveFiles = $false
$script:invalidFileNames = @()
$script:validSources = @()

# Statistics
$script:statsFolderNew = 0
$script:statsFolderSaved = 0
$script:statsFolderSkip = 0
$script:statsFileNew = 0
$script:statsFileUpdate = 0
$script:statsFileIgnore = 0
$script:statsFileDelete = 0
$script:statsFileInvalid = 0
$script:statsFileTemp = 0
$script:statsTotalBytes = 0

# Ensure log directory exists
if (-not (Test-Path $logPath)) {
    New-Item -ItemType Directory -Path $logPath -Force | Out-Null
}

function Save-Log {
    $now = Get-Date
    $logFile = $now.ToString("yyyyMMdd-HHmmss") + ".log"
    $logFilePath = Join-Path $logPath $logFile
    
    [System.IO.File]::WriteAllText($logFilePath, $script:logContent, [System.Text.Encoding]::UTF8)
}

function Write-FatalError {
    param([string]$Message)
    
    $script:logContent += "`r`n" + "FATAL ERROR " + $Message
    Save-Log
    exit 1
}

function Get-FolderSize {
    param([string]$FolderPath)
    
    $totalSize = 0
    
    if (-not (Test-Path $FolderPath)) {
        return 0
    }
    
    try {
        $files = Get-ChildItem -Path $FolderPath -File -Recurse -Force -ErrorAction SilentlyContinue
        foreach ($file in $files) {
            try {
                $totalSize += $file.Length
            }
            catch {
                # Skip files we can't access
            }
        }
    }
    catch {
        # Return 0 if we can't access the folder
    }
    
    return $totalSize
}

function Format-Bytes {
    param([long]$Bytes)
    
    if ($Bytes -ge 1GB) {
        return "{0:N2} GB" -f ($Bytes / 1GB)
    }
    elseif ($Bytes -ge 1MB) {
        return "{0:N2} MB" -f ($Bytes / 1MB)
    }
    elseif ($Bytes -ge 1KB) {
        return "{0:N2} KB" -f ($Bytes / 1KB)
    }
    else {
        return "$Bytes bytes"
    }
}

function Format-Duration {
    param([int]$Seconds)
    
    $hours = [Math]::Floor($Seconds / 3600)
    $minutes = [Math]::Floor(($Seconds % 3600) / 60)
    $secs = $Seconds % 60
    
    $result = ""
    if ($hours -gt 0) {
        $result += "${hours}h "
    }
    if ($minutes -gt 0 -or $hours -gt 0) {
        $result += "${minutes}m "
    }
    $result += "${secs}s"
    
    return $result
}

function ConvertTo-CamelCasePath {
    param([string]$FullPath)
    
    $driveLetter = ""
    if ($FullPath -match '^([A-Z]):') {
        $driveLetter = $matches[1].ToUpper()
        $pathWithoutDrive = $FullPath.Substring(2).TrimStart('\')
    }
    else {
        $pathWithoutDrive = $FullPath.TrimStart('\')
    }
    
    $parts = $pathWithoutDrive -split '\\'
    $result = $driveLetter
    
    foreach ($part in $parts) {
        if ($part) {
            # Clean special characters
            $cleanPart = $part -replace '[^a-zA-Z0-9]', '-'
            $cleanPart = $cleanPart.TrimEnd('-').ToLower()
            
            if ($cleanPart) {
                if ($result) {
                    $result += "_"
                }
                $result += $cleanPart
            }
        }
    }
    
    return $result
}

function Test-IgnoredFile {
    param([string]$FileName)
    
    $lowerName = $FileName.ToLower()
    $ext = [System.IO.Path]::GetExtension($FileName).TrimStart('.').ToLower()
    
    # Office temporary files
    if ($lowerName.StartsWith('.~')) { return $true }
    if ($ext.EndsWith('#')) { return $true }
    
    # Browser downloads
    if ($ext -eq 'crdownload' -or $ext -eq 'part') { return $true }
    
    # Common temporary extensions
    if ($ext -in @('tmp', 'temp', 'bak', 'cache')) { return $true }
    
    # System files
    if ($lowerName -in @('thumbs.db', 'desktop.ini')) { return $true }
    
    return $false
}

function Test-GenericFilename {
    param([string]$BaseName)
    
    # Remove trailing digits
    $cleanBaseName = $BaseName -replace '\d+$', ''
    $cleanBaseName = $cleanBaseName.Trim().ToLower()
    
    foreach ($invalidName in $script:invalidFileNames) {
        if ($cleanBaseName -eq $invalidName) {
            return $true
        }
    }
    
    return $false
}

function Test-SameFile {
    param(
        [string]$SourceFile,
        [string]$TargetFile
    )
    
    if ((Test-Path $SourceFile) -and (Test-Path $TargetFile)) {
        $srcInfo = Get-Item $SourceFile -Force
        $tgtInfo = Get-Item $TargetFile -Force
        
        if ($srcInfo.Length -eq $tgtInfo.Length -and 
            $srcInfo.LastWriteTime -eq $tgtInfo.LastWriteTime) {
            return $true
        }
    }
    
    return $false
}

function Test-SubfolderAlreadyProcessed {
    param([string]$SubfolderPath)
    
    $normalizedSubPath = $SubfolderPath.TrimEnd('\').ToLower()
    
    foreach ($source in $script:validSources) {
        $normalizedSourcePath = $source.SourcePath.TrimEnd('\').ToLower()
        if ($normalizedSubPath -eq $normalizedSourcePath) {
            return $true
        }
    }
    
    return $false
}

function Save-PreviousFileVersion {
    param(
        [string]$SourceFile,
        [string]$TargetFile
    )
    
    $targetInfo = Get-Item $TargetFile -Force
    $modDate = $targetInfo.LastWriteTime
    $prefixDate = $modDate.ToString("yyyyMMdd")
    
    $parentPath = Split-Path $TargetFile -Parent
    $fileName = Split-Path $TargetFile -Leaf
    
    # Calculate relative path from target root
    $relativePath = $parentPath.Substring($script:targetPath.Length).TrimStart('\')
    
    # Create .ver directory structure
    $verRootPath = Join-Path $script:targetPath ".ver"
    if (-not (Test-Path $verRootPath)) {
        New-Item -ItemType Directory -Path $verRootPath -Force | Out-Null
        $script:logContent += "[MKDIR] $verRootPath`r`n"
    }
    
    $verFilePath = if ($relativePath) {
        $path = Join-Path $verRootPath $relativePath
        if (-not (Test-Path $path)) {
            New-Item -ItemType Directory -Path $path -Force | Out-Null
            $script:logContent += "[MKDIR] $path`r`n"
        }
        $path
    }
    else {
        $verRootPath
    }
    
    $newName = "${prefixDate}_$fileName"
    $newPath = Join-Path $verFilePath $newName
    
    # Check if version already exists (replace for dev/testing)
    if (Test-Path $newPath) {
        Remove-Item $newPath -Force
        $script:logContent += "[DELETE] $newPath (replacing with newer version)`r`n"
    }
    
    Move-Item $TargetFile $newPath -Force
    $script:logContent += "[BACKUP] $newPath`r`n"
}

function Backup-FolderFiles {
    param(
        [string]$SourcePath,
        [string]$TargetPath
    )
    
    try {
        $files = Get-ChildItem -Path $SourcePath -File -Force -ErrorAction Stop
        
        foreach ($file in $files) {
            try {
                # Skip temporary files
                if (Test-IgnoredFile $file.Name) {
                    $script:logContent += "[TEMP] $($file.FullName)`r`n"
                    $script:statsFileTemp++
                    continue
                }
                
                # Check generic filename
                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
                if (Test-GenericFilename $baseName) {
                    $script:logContent += "[CHECK] $($file.FullName)`r`n"
                    $script:statsFileInvalid++
                    continue
                }
                
                $targetFile = Join-Path $TargetPath $file.Name
                
                # Check if files are identical
                if (Test-SameFile $file.FullName $targetFile) {
                    if ($script:moveFiles) {
                        $script:logContent += "[DELETE] $($file.FullName) -> $targetFile`r`n"
                        Remove-Item $file.FullName -Force
                        $script:statsFileDelete++
                    }
                    else {
                        $script:logContent += "[IGNORE] $($file.FullName) -> $targetFile`r`n"
                        $script:statsFileIgnore++
                    }
                }
                else {
                    $isUpdate = Test-Path $targetFile
                    
                    # Track bytes
                    $script:statsTotalBytes += $file.Length
                    
                    # Keep old version if updating
                    if ($isUpdate) {
                        Save-PreviousFileVersion $file.FullName $targetFile
                    }
                    
                    # Copy or move file
                    if ($script:moveFiles) {
                        Move-Item $file.FullName $targetFile -Force
                    }
                    else {
                        Copy-Item $file.FullName $targetFile -Force
                    }
                    
                    if ($isUpdate) {
                        $script:logContent += "[UPDATE] $($file.FullName) -> $targetFile`r`n"
                        $script:statsFileUpdate++
                    }
                    else {
                        $script:logContent += "[NEW] $($file.FullName) -> $targetFile`r`n"
                        $script:statsFileNew++
                    }
                }
            }
            catch {
                $script:logContent += "[ERROR] Failed to process file: $($file.FullName) - $($_.Exception.Message)`r`n"
            }
        }
    }
    catch {
        $script:logContent += "[ERROR] Cannot access folder: $SourcePath - $($_.Exception.Message)`r`n"
    }
}

function Backup-SubFolders {
    param(
        [string]$SourcePath,
        [string]$TargetPath
    )
    
    try {
        $folders = Get-ChildItem -Path $SourcePath -Directory -Force -ErrorAction Stop
        
        foreach ($folder in $folders) {
            # Check if subfolder is already processed separately
            if (Test-SubfolderAlreadyProcessed $folder.FullName) {
                $script:logContent += "[SKIP] $($folder.FullName) (already processed separately)`r`n"
                $script:statsFolderSkip++
                continue
            }
            
            $targetSubFolder = Join-Path $TargetPath $folder.Name
            
            if (-not (Test-Path $targetSubFolder)) {
                New-Item -ItemType Directory -Path $targetSubFolder -Force | Out-Null
                $script:logContent += "[MKDIR] $targetSubFolder`r`n"
                $script:statsFolderNew++
            }
            
            Backup-FolderFiles $folder.FullName $targetSubFolder
            Backup-SubFolders $folder.FullName $targetSubFolder
        }
    }
    catch {
        $script:logContent += "[ERROR] Cannot access subfolders: $SourcePath - $($_.Exception.Message)`r`n"
    }
}

# Main script execution
$script:logContent += "Started at: $(Get-Date)`r`n`r`n"
$script:logContent += "Working directory: $scriptPath`r`n"
$script:logContent += "Config file: $ConfigPath`r`n"

if (-not (Test-Path $ConfigPath)) {
    Write-FatalError "Configuration file not found"
}

if (-not (Test-Path $script:invalidNamesPath)) {
    Write-FatalError "Invalid names file not found: $($script:invalidNamesPath)"
}

$script:logContent += "Parsing config file ....`r`n"

# Parse configuration
$configContent = Get-Content $ConfigPath -Encoding UTF8
$sources = @()
$inMainSection = $false
$inSourcesSection = $false

foreach ($line in $configContent) {
    $line = $line.Trim()
    
    if ($line -eq "[Main]") {
        $inMainSection = $true
        $inSourcesSection = $false
        continue
    }
    
    if ($line -eq "[Sources]") {
        $inSourcesSection = $true
        $inMainSection = $false
        continue
    }
    
    if ($line -match '^\[.*\]$') {
        $inMainSection = $false
        $inSourcesSection = $false
        continue
    }
    
    if ($line -and -not $line.StartsWith(';') -and -not $line.StartsWith('#')) {
        if ($line -match '^(\w+)\s*=\s*(.+)$') {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            
            if ($inMainSection) {
                switch ($key) {
                    'targetPath' { $script:targetPath = $value }
                    'moveFiles' { $script:moveFiles = $value -eq '1' }
                }
            }
            elseif ($inSourcesSection -and $key -match '^sourceFolder\d+$') {
                if ($value) {
                    $parts = $value -split ',', 2
                    $sources += @{
                        SourcePath   = $parts[0].Trim()
                        TargetFolder = if ($parts.Count -gt 1) { $parts[1].Trim() } else { "" }
                    }
                }
            }
        }
    }
}

if (-not $script:targetPath) {
    Write-FatalError "Target folder not defined in the configuration file"
}

$moveFilesReadable = if ($script:moveFiles) { "Yes" } else { "No" }
$script:logContent += "Target path: $($script:targetPath)`r`n"
$script:logContent += "Move files: $moveFilesReadable`r`n`r`n"

# Check target path
$targetDrive = Split-Path $script:targetPath -Qualifier
if (-not (Test-Path $targetDrive)) {
    Write-FatalError "Target drive is not available: $targetDrive"
}

if (-not (Test-Path $script:targetPath)) {
    Write-FatalError "Target path does not exist: $($script:targetPath)"
}

# Check disk space
$drive = Get-PSDrive ($targetDrive.TrimEnd(':'))
$availableSpaceBefore = $drive.Free
$backupSizeBefore = Get-FolderSize $script:targetPath

$script:logContent += "Available disk space: $(Format-Bytes $availableSpaceBefore)`r`n"
$script:logContent += "Current backup size: $(Format-Bytes $backupSizeBefore)`r`n`r`n"

# Start timing
$startTime = Get-Date

# Validate sources
if ($sources.Count -eq 0) {
    Write-FatalError "No sources defined in the configuration file"
}

$script:logContent += "Check sources ....`r`n"

$sourceCountOK = 0
$sourceCountKO = 0

foreach ($i in 0..($sources.Count - 1)) {
    $source = $sources[$i]
    $script:logContent += "Source $($i+1): $($source.SourcePath)"
    
    if ($source.TargetFolder) {
        $script:logContent += " -> $($source.TargetFolder)\"
    }
    
    if (-not (Test-Path $source.SourcePath)) {
        $script:logContent += " .. [KO]`r`n"
        $sourceCountKO++
    }
    else {
        $script:logContent += " .. [OK]`r`n"
        $script:validSources += $source
        $sourceCountOK++
    }
}

if ($sourceCountOK -eq 0) {
    Write-FatalError "No valid source folders"
}

$script:logContent += "`r`nSources: $($sources.Count)`r`n"
$script:logContent += "  OK: $sourceCountOK`r`n"
$script:logContent += "  KO: $sourceCountKO`r`n`r`n"

# Load invalid filenames
$script:logContent += "Loading invalid filenames ....`r`n"

$invalidNamesContent = Get-Content $script:invalidNamesPath -Encoding UTF8
foreach ($line in $invalidNamesContent) {
    $line = $line.Trim()
    if ($line -and -not $line.StartsWith(';') -and -not $line.StartsWith('#')) {
        $script:invalidFileNames += $line.ToLower()
    }
}

$script:logContent += "  |-> $($script:invalidFileNames.Count) invalid filename patterns`r`n`r`n"

# Execute backup
$script:logContent += "`r`n=-=-=-=-=-=-=-=-=-=`r`n"
$script:logContent += "||    BACKUP     ||`r`n"
$script:logContent += "=-=-=-=-=-=-=-=-=-=`r`n`r`n"

foreach ($source in $script:validSources) {
    $script:logContent += " >> $($source.SourcePath)"
    
    if ($source.TargetFolder) {
        $script:logContent += " -> $($source.TargetFolder)`r`n"
        $targetFullPath = Join-Path $script:targetPath $source.TargetFolder
    }
    else {
        $script:logContent += "`r`n"
        $targetFolderName = ConvertTo-CamelCasePath $source.SourcePath
        $targetFullPath = Join-Path $script:targetPath $targetFolderName
    }
    
    if (-not (Test-Path $targetFullPath)) {
        New-Item -ItemType Directory -Path $targetFullPath -Force | Out-Null
        $script:logContent += "[MKDIR] $targetFullPath`r`n"
        $script:statsFolderNew++
    }
    
    Backup-FolderFiles $source.SourcePath $targetFullPath
    Backup-SubFolders $source.SourcePath $targetFullPath
    
    $script:statsFolderSaved++
}

# Calculate final statistics
$endTime = Get-Date
$duration = [int](($endTime - $startTime).TotalSeconds)

$drive = Get-PSDrive ($targetDrive.TrimEnd(':'))
$availableSpaceAfter = $drive.Free
$backupSizeAfter = Get-FolderSize $script:targetPath
$backupGrowth = $backupSizeAfter - $backupSizeBefore

$script:logContent += "`r`n===== STATISTICS =====`r`n"
$script:logContent += "Duration           : $(Format-Duration $duration)`r`n"
$script:logContent += "Processed folders  : $($script:statsFolderSaved)`r`n"
$script:logContent += "New folders        : $($script:statsFolderNew)`r`n"
$script:logContent += "Skipped folders    : $($script:statsFolderSkip)`r`n"
$script:logContent += "New files          : $($script:statsFileNew)`r`n"
$script:logContent += "Modified files     : $($script:statsFileUpdate)`r`n"
$script:logContent += "Ignored files      : $($script:statsFileIgnore)`r`n"
$script:logContent += "Invalid files      : $($script:statsFileInvalid)`r`n"
$script:logContent += "Temporary files    : $($script:statsFileTemp)`r`n"
$script:logContent += "Deleted files      : $($script:statsFileDelete)`r`n"
$script:logContent += "Data backed up     : $(Format-Bytes $script:statsTotalBytes)`r`n"
$script:logContent += "Total backup size  : $(Format-Bytes $backupSizeAfter)`r`n"
$script:logContent += "Backup growth      : $(Format-Bytes $backupGrowth)`r`n"
$script:logContent += "Available space    : $(Format-Bytes $availableSpaceAfter)`r`n`r`n"

if ($sourceCountKO -gt 0) {
    $script:logContent += "`r`nERRORS`r`n"
}
else {
    $script:logContent += "`r`nSUCCESS`r`n"
}

Save-Log
Write-Host "Backup completed. Log saved to: $(Join-Path $logPath (Get-Date -Format 'yyyyMMdd-HHmmss')).log" -ForegroundColor Green
