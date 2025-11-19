$scriptPath = Join-Path $PSScriptRoot "backup.vbs"
$configPath = Join-Path $PSScriptRoot "backup.ini"

$targetPath = Join-Path (Split-Path $PSScriptRoot -Parent) "Backup"
$taskPath = "\JDZ\"
$taskName = "jBackup"
$taskDesc = "Automatic daily backup."
$scheduleHour = "12:00"
$scheduleDelay = 'daily'
$everyDay = @("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")
$scheduleDays = $everyDay
$sources = @()

# Function to validate and prompt for target path
function Get-ValidTargetPath {
    param(
        [string]$InitialPath,
        [int]$MaxAttempts = 10
    )
    
    $attempts = 0
    $currentPath = $InitialPath
    
    while ($attempts -lt $MaxAttempts) {
        $attempts++
        
        # Prompt for path if empty or after failed validation
        if (-not $currentPath -or $attempts -gt 1) {
            $currentPath = Read-Host "Enter a valid target path for backups (attempt $attempts/$MaxAttempts)"
            
            if (-not $currentPath) {
                Write-Host "Target path is required" -ForegroundColor Red
                continue
            }
        }
        
        # Check if path exists
        if (-not (Test-Path $currentPath)) {
            $askCreate = Read-Host "Path does not exist. Do you want to create it? (Y/N)"
            if ($askCreate -eq 'Y' -or $askCreate -eq 'y') {
                try {
                    New-Item -ItemType Directory -Path $currentPath -Force | Out-Null
                    Write-Host "Path created: $currentPath" -ForegroundColor Green
                }
                catch {
                    Write-Host "Failed to create path: $currentPath" -ForegroundColor Red
                    Write-Host "  $_" -ForegroundColor Red
                    $currentPath = $null
                    continue
                }
            }
            else {
                Write-Host "Path not found" -ForegroundColor Red
                $currentPath = $null
                continue
            }
        }
        
        # Check if drive is accessible
        $driveLetter = Split-Path -Path $currentPath -Qualifier
        if (-not $driveLetter) {
            Write-Host "Cannot determine drive letter" -ForegroundColor Red
            $currentPath = $null
            continue
        }
        
        if (-not (Test-Path $driveLetter)) {
            Write-Host "Drive not accessible: $driveLetter" -ForegroundColor Red
            $currentPath = $null
            continue
        }
        
        # Check if path is writable
        $testFile = Join-Path $currentPath ".writetest_$(Get-Random).tmp"
        try {
            [System.IO.File]::WriteAllText($testFile, "test")
            Remove-Item $testFile -Force
            Write-Host "[OK] $currentPath" -ForegroundColor Green
            return $currentPath
        }
        catch {
            Write-Host "Path is not writable" -ForegroundColor Red
            Write-Host "  $_" -ForegroundColor Red
            $currentPath = $null
            continue
        }
    }
    
    # Max attempts reached
    Write-Host "FATAL ERROR: Maximum attempts reached. Could not configure a valid target path." -ForegroundColor Red
    exit 1
}

# Function to convert a source path to a clean folder name (lowercase with underscores)
function ConvertTo-FolderName {
    param(
        [string]$Path
    )
    
    # Get the absolute path
    $fullPath = [System.IO.Path]::GetFullPath($Path)
    
    # Extract drive letter and path
    $drive = [System.IO.Path]::GetPathRoot($fullPath)
    $pathWithoutDrive = $fullPath.Substring($drive.Length)
    
    # Get drive letter without colon and backslash
    $driveLetter = $drive.Substring(0, 1).ToLower()
    
    # Clean the path: replace backslashes with underscores, remove special characters
    $cleanPath = $pathWithoutDrive -replace '[\\/:*?"<>|]', '_'
    $cleanPath = $cleanPath -replace '[^a-zA-Z0-9_]', '_'
    $cleanPath = $cleanPath -replace '_+', '_'
    $cleanPath = $cleanPath.Trim('_').ToLower()
    
    # Combine drive letter with clean path
    if ($cleanPath) {
        return "${driveLetter}_$cleanPath"
    }
    else {
        return $driveLetter
    }
}

# Function to prompt for target folder and validate against default
function Get-TargetFolderInput {
    param(
        [string]$SourcePath,
        [string]$PromptMessage = "Enter target subfolder name (ex: 'Documents', 'Images', 'Documents/Important')"
    )
    
    # Validate the source path exists and is accessible
    if (-not (Test-Path $SourcePath)) {
        Write-Host "[KO] $SourcePath does not exist" -ForegroundColor Yellow
        return $null
    }
    
    try {
        $null = Get-ChildItem -Path $SourcePath -ErrorAction Stop
    }
    catch {
        Write-Host "[KO] $SourcePath is not accessible" -ForegroundColor Yellow
        return $null
    }
    
    # Generate default target folder name
    $defaultTargetFolder = ConvertTo-FolderName -Path $SourcePath
    
    # Prompt for target folder with default
    $targetFolder = Read-Host "$PromptMessage (default: $defaultTargetFolder, press Enter to use default)"
    $targetFolder = $targetFolder.Trim()
    
    # Check if user wants default
    if (-not $targetFolder -or $targetFolder -eq $defaultTargetFolder) {
        # Return empty string to indicate auto-generation in VBS
        return [PSCustomObject]@{
            SourcePath    = $SourcePath
            TargetFolder  = ""
            DisplayTarget = $defaultTargetFolder
        }
    }
    else {
        # Return custom value
        return [PSCustomObject]@{
            SourcePath    = $SourcePath
            TargetFolder  = $targetFolder
            DisplayTarget = $targetFolder
        }
    }
}

# Function to save configuration to INI file
function Save-BackupConfig {
    param(
        [string]$ConfigPath,
        [string]$TargetPath,
        [string]$ScheduleHour,
        [string]$ScheduleDelay,
        [array]$ScheduleDays,
        [array]$Sources  # Array of PSCustomObjects with SourcePath and TargetFolder properties
    )
    
    # Build sources section from PSCustomObject array
    $sourcesSection = ""
    for ($i = 0; $i -lt $Sources.Count; $i++) {
        $source = $Sources[$i]
        # Only include targetFolder if it's not empty (user explicitly set it)
        if ($source.TargetFolder) {
            $sourcesSection += "sourceFolder$($i+1)=$($source.SourcePath),$($source.TargetFolder)`r`n"
        }
        else {
            $sourcesSection += "sourceFolder$($i+1)=$($source.SourcePath)`r`n"
        }
    }
    
    # Build complete config
    $configContent = @"
[Main]
targetPath=$TargetPath
scheduleHour=$ScheduleHour
scheduleDelay=$ScheduleDelay
scheduleDays=$($ScheduleDays -join ',')

[Sources]
$sourcesSection
"@
    
    # Save to file
    $configContent | Out-File -FilePath $ConfigPath -Encoding UTF8
}

Write-Host ""
Write-Host "  ========================================  " -ForegroundColor White
Write-Host "           _                _               " -ForegroundColor DarkGreen
Write-Host "          | |              | |              " -ForegroundColor DarkGreen
Write-Host "       _  | |__   __ _  ___| | ___   _ ___  " -ForegroundColor DarkGreen
Write-Host "      |_| |  _ \ / _  |/ __| |/ / | | | O | " -ForegroundColor DarkRed
Write-Host "       _  | |_) | (_| | (__|   <| |_| | __| " -ForegroundColor DarkRed
Write-Host "      | | |____/ \__,_|\___|_|\_\\__,_|_|   " -ForegroundColor DarkRed
Write-Host "  _   | |                                   " -ForegroundColor DarkMagenta
Write-Host " | |__| |                                   " -ForegroundColor DarkMagenta
Write-Host "  \____/                                    " -ForegroundColor DarkMagenta
Write-Host "  ========================================  " -ForegroundColor White
Write-Host ""
Write-Host ""
Write-Host "This script configures an automatic backup system that:" -ForegroundColor Cyan
Write-Host "  - Backs up multiple source folders to a target directory" -ForegroundColor White
Write-Host "  - Creates a Windows scheduled task for automatic execution" -ForegroundColor White
Write-Host "  - Runs daily at a specified time AND 5 minutes after system startup" -ForegroundColor White
Write-Host "  - Waits for AC power to start, but continues if unplugged during backup" -ForegroundColor White
Write-Host "  - Preserves old file versions with YYYYMMDD_ prefix when files change" -ForegroundColor White
Write-Host "  - Skips identical files to save time and space" -ForegroundColor White
Write-Host "  - Logs all operations with detailed statistics" -ForegroundColor White
Write-Host ""
Write-Host "CONFIGURATION OPTIONS:" -ForegroundColor Yellow
Write-Host "  Target Path    : Root backup destination folder (default: $targetPath)" -ForegroundColor White
Write-Host "  Source Folders : Multiple folders to backup (with optional target subfolders)" -ForegroundColor White
Write-Host "  Execution Time : Hour of daily execution (default: $scheduleHour)" -ForegroundColor White
Write-Host "  Schedule       : Daily or Weekly execution (default: $scheduleDelay)" -ForegroundColor White
Write-Host "  Schedule days  : When weekly sync specify the day(s) (default: $($scheduleDays -join ', '))" -ForegroundColor White
Write-Host ""
Write-Host "ADVANCED FEATURES:" -ForegroundColor Yellow
Write-Host "  - Smart subfolder exclusion (skips subfolders already configured separately)" -ForegroundColor White
Write-Host "  - Target subfolder mapping (format: 'source,targetSubfolder')" -ForegroundColor White
Write-Host "  - Automatic target folder name generation from source path" -ForegroundColor White
Write-Host "  - Queue mode for catching up missed backups" -ForegroundColor White
Write-Host "  - Waits for AC power before starting (but won't interrupt if power is lost)" -ForegroundColor White
Write-Host ""
Write-Host "Press Enter to begin configuration..." -ForegroundColor Green
$null = Read-Host

$existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($existingTask) {
    Write-Host "Removing existing task ..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}

if (-not (Test-Path $scriptPath)) {
    Write-Host "FATAL ERROR: backup.vbs not found in current directory" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Setting up backup configuration ..." -ForegroundColor DarkBlue
Write-Host ""

# Read settings from config file if it exists
if (Test-Path $configPath) {
    $configContent = Get-Content $configPath -Raw

    if ($configContent) {
        $inMainSection = $false
        $inSourcesSection = $false
        foreach ($line in $configContent -split "`r`n") {
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
            
            if ($inMainSection -and $line -match '^(\w+)\s*=\s*(.+)$') {
                $key = $matches[1].Trim()
                $value = $matches[2].Trim()
                
                switch ($key) {
                    "targetPath" { $targetPath = $value }
                    "scheduleHour" { $scheduleHour = $value }
                    "scheduleDelay" { $scheduleDelay = $value }
                    "scheduleDays" { $scheduleDays = $value -split ',' | ForEach-Object { $_.Trim() } }
                }
            }

            if ( $inSourcesSection -and $line -match '^sourceFolder\d+\s*=\s*(.+)$') {
                $sourceEntry = $matches[1].Trim()
                # Parse source path and target subfolder (separated by comma)
                $parts = $sourceEntry -split ',', 2
                $sourcePath = $parts[0].Trim()
                $targetFolder = if ($parts.Count -gt 1 -and $parts[1].Trim()) { 
                    $parts[1].Trim() 
                }
                else { 
                    ""  # Empty if not specified - auto-naming will happen in VBS
                }
                
                # Add to sources array as PSCustomObject
                $sources += [PSCustomObject]@{
                    SourcePath   = $sourcePath
                    TargetFolder = $targetFolder
                }
            }
        }

        Write-Host "Loaded settings from config file" -ForegroundColor Green
    }
}

#####
# prompt for settings
#####

$inputTargetPath = Read-Host "Enter the Target Path for backups (current: $targetPath)"
if ( $inputTargetPath ) { 
    $targetPath = $inputTargetPath
}
$targetPath = Get-ValidTargetPath -InitialPath $targetPath

$inputScheduleHour = Read-Host "Enter the hour of execution for the scheduled task (0 - 23) (current: $scheduleHour)"
if ( $inputScheduleHour ) { 
    $scheduleHour = $inputScheduleHour
}

$inputScheduleDelay = Read-Host "Enter the delay for the scheduled task (daily, weekly) (current: $scheduleDelay)"
if ( $inputScheduleDelay ) { 
    $scheduleDelay = $inputScheduleDelay
}
if ( $inputScheduleDelay -eq 'weekly' ) {
    $inputScheduleDays = Read-Host "Enter the days of execution for the scheduled task (e.g. Monday, Wednesday, Friday) (current: everyday)"
    if ( $inputScheduleDays ) { 
        $scheduleDays = $inputScheduleDays -split ',' | ForEach-Object { $_.Trim() }
    }

    # everyday so 'daily'
    if ( $scheduleDays.Count -eq 7 ) {
        $scheduleDelay = 'daily'
        $scheduleDays = $everyDay
    }
}
else {
    $scheduleDelay = 'daily'
    $scheduleDays = $everyDay
}

# update config file
Write-Host "Saving config to $configPath" -ForegroundColor Yellow

Save-BackupConfig -ConfigPath $configPath `
    -TargetPath $targetPath `
    -ScheduleHour $scheduleHour `
    -ScheduleDelay $scheduleDelay `
    -ScheduleDays $scheduleDays `
    -Sources $sources

#####
# source folders setup
#####

Write-Host "Setting up source folders ..." -ForegroundColor Cyan

# Validate source folders from $sources array
if ($sources.Count -gt 0) {
    Write-Host "Checking source folders from config file" -ForegroundColor Yellow

    $validSources = @()
    $invalidSources = @()

    foreach ($source in $sources) {
        $sourcePath = $source.SourcePath
        $targetFolder = $source.TargetFolder
        
        # Check if path exists
        if (Test-Path $sourcePath) {
            # Check if files are accessible for copy (try to list files)
            try {
                $null = Get-ChildItem -Path $sourcePath -ErrorAction Stop
                $validSources += $source
                Write-Host "[OK] $sourcePath" -ForegroundColor Green
            }
            catch {
                Write-Host "[KO] $sourcePath not accessible" -ForegroundColor Yellow
                $invalidSources += $source
            }
        }
        else {
            Write-Host "[KO] $sourcePath not found" -ForegroundColor Yellow
            $invalidSources += $source
        }
    }

    # Process invalid sources
    foreach ($invalidSource in $invalidSources) {
        $displayPath = if ($invalidSource.TargetFolder) { 
            "$($invalidSource.SourcePath), $($invalidSource.TargetFolder)" 
        }
        else { 
            $invalidSource.SourcePath 
        }
        
        Write-Host "Invalid source folder: $displayPath" -ForegroundColor Yellow
        $inputSourceFolder = Read-Host "Enter a new source path to replace it, or press Enter to remove it"
        
        if ($inputSourceFolder) {
            $inputSourceFolder = $inputSourceFolder.Trim()
            
            # Get target folder input with validation (includes source path validation)
            $result = Get-TargetFolderInput -SourcePath $inputSourceFolder
            
            if ($result) {
                $validSources += [PSCustomObject]@{
                    SourcePath   = $result.SourcePath
                    TargetFolder = $result.TargetFolder
                }
                Write-Host "[OK] $($result.SourcePath) -> $($result.DisplayTarget)" -ForegroundColor Green
            }
        }
        else {
            Write-Host "Removing invalid source folder from configuration." -ForegroundColor Yellow
        }
    }

    # Update $sources with valid sources
    $sources = $validSources

    # If no valid source folders, ask for new ones
    if ($sources.Count -eq 0) {
        Write-Host "No valid source folders found. Please add at least one valid path." -ForegroundColor Yellow
    }
}

# add source folders (up to 50)

for ($i = $sources.Count + 1; $i -le 50; $i++) {
    if ($i -eq 1) {
        $sourcePath = Read-Host "Enter source folder $i to backup (required)"
        
        if (-not $sourcePath) {
            Write-Host "FATAL ERROR: At least one source folder is required" -ForegroundColor Red
            exit 1
        }
    }
    else {
        $sourcePath = Read-Host "Enter source folder $i to backup (optional, press Enter to skip)"
        
        if (-not $sourcePath) {
            break
        }
    }
    
    $sourcePath = $sourcePath.Trim()
    
    # Get target folder input with validation (includes source path validation)
    $result = Get-TargetFolderInput -SourcePath $sourcePath
    
    if ($result) {
        # Add to sources array as PSCustomObject
        $sources += [PSCustomObject]@{
            SourcePath   = $result.SourcePath
            TargetFolder = $result.TargetFolder
        }
        Write-Host "[OK] $($result.SourcePath) -> $($result.DisplayTarget)" -ForegroundColor Green
    }
}

# Final check
# Should not go there but just in case
if ($sources.Count -eq 0) {
    Write-Host "FATAL ERROR: No valid source folders were added" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Final configuration summary:" -ForegroundColor DarkBlue
Write-Host ""

Write-Host "Target Path: $targetPath" -ForegroundColor Cyan
Write-Host "Schedule Hour: $scheduleHour" -ForegroundColor Cyan
Write-Host "Schedule Delay: $scheduleDelay" -ForegroundColor Cyan
Write-Host "Schedule Days: $($scheduleDays -join ', ')" -ForegroundColor Cyan
Write-Host ""
Write-Host "Source Folders:" -ForegroundColor Cyan
foreach ($source in $sources) {
    $displayPath = if ($source.TargetFolder) { 
        "$($source.SourcePath) -> $($source.TargetFolder)" 
    }
    else { 
        $source.SourcePath 
    }
    Write-Host "  - $displayPath" -ForegroundColor White
}
Write-Host ""

# update config file
Write-Host "Saving config to $configPath" -ForegroundColor Yellow

Save-BackupConfig -ConfigPath $configPath `
    -TargetPath $targetPath `
    -ScheduleHour $scheduleHour `
    -ScheduleDelay $scheduleDelay `
    -ScheduleDays $scheduleDays `
    -Sources $sources

#####
# Setup the scheduled task
#####

Write-Host ""
Write-Host "Task configuration complete !" -ForegroundColor Green
Write-Host ""

# Prompt to execute the backup.vbs now
$response = Read-Host "Do you want to run the backup now? (Y/N)"
if ($response -eq 'Y' -or $response -eq 'y') {
    Write-Host "Executing backup script..." -ForegroundColor Green
    Start-Process -FilePath "wscript.exe" -ArgumentList "`"$scriptPath`"" -NoNewWindow -Wait
    Write-Host "Backup script execution completed." -ForegroundColor Green
}
Write-Host "Skipping backup script execution." -ForegroundColor Yellow

Write-Host ""

$scriptPath = (Resolve-Path $scriptPath).Path

Write-Host "Script ........................... $scriptPath" -ForegroundColor Cyan
Write-Host "Task Path ........................ $taskPath" -ForegroundColor Cyan
Write-Host "Task Name ........................ $taskName" -ForegroundColor Cyan
Write-Host "Schedule Hour .................... $scheduleHour" -ForegroundColor Cyan
Write-Host "Schedule Delay ................... $scheduleDelay" -ForegroundColor Cyan
Write-Host "Schedule Days .................... $($scheduleDays -join ', ')" -ForegroundColor Cyan
Write-Host "Start if on batteries ............ No" -ForegroundColor Cyan
Write-Host "Stop if going on batteries ....... No" -ForegroundColor Cyan
Write-Host "Start when available ............. Yes" -ForegroundColor Cyan
Write-Host "Run only if network available .... No" -ForegroundColor Cyan
Write-Host "Mandatory use login .............. No" -ForegroundColor Cyan
Write-Host "Run multiple instances ........... No" -ForegroundColor Cyan
Write-Host "Catch up missed backups .......... Yes" -ForegroundColor Cyan
Write-Host "Startup delay .................... 5 minutes" -ForegroundColor Cyan
Write-Host "Description ...................... Automatic daily backup." -ForegroundColor Cyan
Write-Host ""

Write-Host "Setting up action..." -ForegroundColor Green
    
$action = New-ScheduledTaskAction -Execute "wscript.exe" -Argument "`"$scriptPath`"" -WorkingDirectory $PSScriptRoot

Write-Host "Setting up triggers..." -ForegroundColor Green

if ( $scheduleDelay -eq 'weekly' ) {
    $trigger1 = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $scheduleDays -At "$scheduleHour"
}
else {
    $trigger1 = New-ScheduledTaskTrigger -Daily -At "$scheduleHour"
}

$trigger2 = New-ScheduledTaskTrigger -AtStartup
$trigger2.Delay = "PT5M"

$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries:$false `
    -DontStopIfGoingOnBatteries:$true `
    -StartWhenAvailable:$true `
    -RunOnlyIfNetworkAvailable:$false `
    -ExecutionTimeLimit (New-TimeSpan -Hours 4) `
    -MultipleInstances Queue

$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType S4U -RunLevel Highest

Write-Host ""
Write-Host "Registering task..." -ForegroundColor DarkBlue

Register-ScheduledTask `
    -TaskPath $taskPath `
    -TaskName $taskName `
    -Action $action `
    -Trigger $trigger1, $trigger2 `
    -Settings $settings `
    -Principal $principal `
    -Description $taskDesc

Write-Host ""
Write-Host "Task created successfully!" -ForegroundColor Green
Write-Host "  Get-ScheduledTask -TaskPath '$taskPath' -TaskName '$taskName'" -ForegroundColor White
Write-Host "  Start-ScheduledTask -TaskPath '$taskPath' -TaskName '$taskName'" -ForegroundColor White
Write-Host "  Unregister-ScheduledTask -TaskPath '$taskPath' -TaskName '$taskName' -Confirm:`$false" -ForegroundColor White
Write-Host ""

# Afficher les prochaines exécutions
Start-Sleep -Seconds 1  # Attendre que la tâche soit complètement enregistrée
try {
    $taskInfo = Get-ScheduledTaskInfo -TaskPath $taskPath -TaskName $taskName -ErrorAction Stop
    Write-Host "Next scheduled run: $($taskInfo.NextRunTime)"
}
catch {
    Write-Host "Task registered successfully but info not yet available." -ForegroundColor Yellow
    Write-Host "Run 'Get-ScheduledTaskInfo -TaskPath ""$taskPath"" -TaskName ""$taskName""' to see details later." -ForegroundColor Yellow
}

Write-Host "`nALL DONE" -ForegroundColor Green
exit 1
