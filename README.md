# jBackup - Automatic Backup System

A comprehensive VBScript-based backup solution with PowerShell configuration assistant for Windows systems.

## Features

### Core Functionality
- **Multi-folder backup**: Back up multiple source folders to a centralized target directory
- **File versioning**: Preserves old file versions with YYYYMMDD_ prefix in `.ver` directory (3 versions max)
- **Smart comparison**: Only copies files that have changed (size + date modified)
- **Subfolder exclusion**: Automatically skips subfolders that are configured as separate backup sources
- **Custom target mapping**: Specify custom target subfolder names or use auto-generated names

### Automation
- **Windows Task Scheduler integration**: Automatically creates scheduled tasks
- **Dual triggers**: Runs daily at specified time AND 5 minutes after system startup
- **Queue mode**: Catches up on missed backups when system was offline
- **Power management**: Waits for AC power to start, but continues if unplugged during backup
- **No login required**: Uses S4U authentication (runs without password)

### Logging & Monitoring
- **Detailed UTF-8 logs**: Categorized operations with timestamps
- **Statistics tracking**: Files copied, moved, errors, skipped folders
- **Log categories**: `[BACKUP]`, `[IGNORE]`, `[NEW]`, `[UPDATE]`, `[MOVE/NEW]`, `[MOVE/UPDATE]`, `[DELETE]`, `[CHECK]`, `[MKDIR]`, `[ERROR]`, `[INFO]`, `[SKIP]`
- **Log location**: `log\YYYYMMDD-HHMMSS.log`

### Generic Filename Detection
- **Automatic flagging**: Files with generic names (like `document.pdf`, `photo2.jpg`, `screenshot.png`) are flagged with `[CHECK]` in logs
- **Customizable list**: Generic filenames defined in `invalidnames.txt`
- **Easy maintenance**: Add or remove generic names without modifying the script
- **Common examples**: document, image, photo, file, screenshot, download, untitled, temp

## File Structure

```
jbackup/
├── backup.vbs          # Main backup script (VBScript)
├── cleanup.vbs         # Version cleanup script (VBScript)
├── setup.ps1           # Interactive configuration assistant (PowerShell)
├── backup.ini          # Configuration file (auto-generated)
├── invalidnames.txt    # Generic filename patterns (auto-generated)
├── log/                # Backup logs directory
│   └── YYYYMMDD-HHMMSS.log
│   └── YYYYMMDD-HHMMSS-cleanup.log
└── README.md           # This file
```

## Installation & Setup

### Prerequisites
- Windows operating system
- PowerShell 5.1 or later
- Administrator privileges (for Task Scheduler)

### Quick Start

1. **Place the files** in your desired location (e.g., `C:\jbackup`)

2. **Run the setup script**:
   ```powershell
   cd C:\jbackup
   .\setup.ps1
   ```

3. **Follow the interactive prompts** to configure:
   - Target backup path
   - Source folders to backup
   - Task scheduler settings
   - Execution schedule

### Configuration Options

#### Target Path
- Root destination folder for all backups
- Default: One level up from script directory (`C:\Backup`)
- Must be writable and accessible

#### Source Folders
- Multiple folders can be configured
- For each source, specify:
  - **Source path**: Folder to backup (relative to target path)
  - **Target subfolder**: Custom name or auto-generated (e.g., `c_users_john_documents`)

#### Schedule Settings
- **Schedule Hour**: Daily execution time (default: `12:00`)
- **Schedule Type**: Daily or Weekly
- **Schedule Days**: Days of week (if weekly)

## Configuration File Format

### backup.ini Structure

```ini
[Main]
targetPath=C:\Backup
moveFiles=false
scheduleHour=12:00
scheduleDelay=daily
scheduleDays=Monday,Tuesday,Wednesday,Thursday,Friday,Saturday,Sunday

[Sources]
sourceFolder1=C:\Users\John\Documents
sourceFolder2=C:\Users\John\Pictures,photos
sourceFolder3=C:\Projects\Code,projects
```

### invalidnames.txt Structure

```ini
; Generic filenames that will be flagged for review
; Add one filename per line (case-insensitive)
; Lines starting with ; or # are comments

document
image
photo
file
screenshot
download
untitled
temp
```

This file is automatically created by `setup.ps1` with a default list of 30+ common generic filenames. You can customize it by:
- Adding new generic patterns (one per line)
- Removing patterns you don't want to flag
- Using comments (`;` or `#`) for documentation

**How it works:**
- The backup script loads this file at startup
- Files with base names matching these patterns (after removing trailing digits) are flagged with `[CHECK]` in logs
- Example: `photo123.jpg`, `document_final.pdf`, `IMG_5678.jpg` will all be flagged
- These files are still backed up normally, but logged for review

### Source Folder Format
- **Without custom target**: `sourceFolder1=C:\path\to\source`
  - Auto-generates target folder name (e.g., `c_path_to_source`)
- **With custom target**: `sourceFolder2=C:\path\to\source,CustomName`
  - Uses specified target folder name

## Usage

### Manual Execution
Run the backup script manually:
```cmd
wscript.exe backup.vbs
```

Or double-click `backup.vbs` in Windows Explorer.

### Scheduled Execution
The setup script automatically creates two scheduled tasks:

**jBackup** - Main backup task that runs:
- **Daily** at the configured time (default: 12:00)
- **At startup** (5 minutes after system boots)

**jBackupCleanup** - Version cleanup task that runs:
- **Daily** at 2:00 AM
- Maintains the 3-version limit for all backed up files
- Removes old versions automatically

### Task Management

**View backup task details**:
```powershell
Get-ScheduledTask -TaskPath '\JDZ\' -TaskName 'jBackup'
```

**Start backup task manually**:
```powershell
Start-ScheduledTask -TaskPath '\JDZ\' -TaskName 'jBackup'
```

**View cleanup task details**:
```powershell
Get-ScheduledTask -TaskPath '\JDZ\' -TaskName 'jBackupCleanup'
```

**Start cleanup task manually**:
```powershell
Start-ScheduledTask -TaskPath '\JDZ\' -TaskName 'jBackupCleanup'
```

**Remove tasks**:
```powershell
Unregister-ScheduledTask -TaskPath '\JDZ\' -TaskName 'jBackup' -Confirm:$false
Unregister-ScheduledTask -TaskPath '\JDZ\' -TaskName 'jBackupCleanup' -Confirm:$false
```

## How It Works

### Backup Process

1. **Read configuration** from `backup.ini`
2. **Create log file** in `log\` directory
3. **For each source folder**:
   - Determine target subfolder (custom or auto-generated)
   - Create target directory if needed
   - **For each file**:
     - Check if file exists in target
     - Compare size and date modified
     - If changed: Move old file to `.ver` directory with YYYYMMDD_ prefix, copy new version
     - If new: Copy directly
   - **For each subfolder**:
     - Skip if subfolder is configured as separate source
     - Otherwise, recursively backup subfolder
4. **Log statistics** and completion

### Auto-Generated Folder Names

Source paths are converted to clean folder names:
- Drive letter extracted and lowercased
- Path segments joined with underscores
- Special characters removed
- Result: `c_users_john_documents`

Example:
```
C:\Users\John\Documents → c_users_john_documents
D:\My Files\Photos      → d_my_files_photos
```

### File Versioning

When a file changes, the old version is preserved in a `.ver` directory at the backup root with a date prefix (YYYYMMDD):
```
C:\Backup\documents\document.txt                 (current version)
C:\Backup\.ver\documents\20251119_document.txt    (version from 2025-11-19)
C:\Backup\.ver\documents\20251118_document.txt    (version from 2025-11-18)
```

**Version Management:**
- Versioned files are stored in `C:\Backup\.ver\` directory
- Directory structure mirrors the backup structure (e.g., `.ver\documents\`, `.ver\photos\`)
- Only the 3 most recent versions are kept per file
- Cleanup runs automatically daily at 2:00 AM via `jBackupCleanup` task
- Manual cleanup: run `cleanup.vbs` or start the cleanup task
- Log entries show `[DELETE]` when old versions are removed

### Generic Filename Handling

Files with generic names are flagged in logs but still backed up normally:
```
[CHECK] C:\source\document.pdf
[NEW] C:\Backup\target\document.pdf
```

This helps you identify files that might need better naming for organization. The detection works by:
1. Extracting the base filename (without extension)
2. Removing trailing digits (e.g., `photo123` → `photo`)
3. Checking against the `invalidnames.txt` list
4. Logging with `[CHECK]` if matched

Examples of files that will be flagged:
- `document.pdf`, `document2.docx`, `photo.jpg`, `photo123.png`, `IMG_5678.jpg`, `screenshot.png`, `screenshot_2024.png`, `new.zip`, `new file.xlsx`, `temp.txt`, `untitled.doc`

## Advanced Features

### Smart Subfolder Exclusion

If you configure both a parent and child folder:
```ini
sourceFolder1=C:\Users\John\Documents
sourceFolder2=C:\Users\John\Documents\Work,work_docs
```

The main Documents backup will automatically skip the `Work` subfolder, preventing duplication.

### Power Management

- **Won't start on battery**: Waits for AC power connection
- **Won't stop if unplugged**: Continues backup if power is lost during execution
- Prevents unnecessary battery drain on laptops

### Error Handling

- Invalid paths are detected during setup
- Inaccessible folders are logged but don't stop execution
- Failed file operations are logged with `[ERROR]` tag
- Configuration validation with retry logic

## Troubleshooting

### Setup Script Issues

**"FATAL ERROR: backup.vbs not found"**
- Ensure `backup.vbs` is in the same directory as `setup.ps1`

**"Drive not accessible"**
- Check if the target path drive exists and is mounted
- Ensure you have read/write permissions

**"Path is not writable"**
- Run PowerShell as Administrator
- Check folder permissions

### Backup Script Issues

**Files not being backed up**
- Check log files in `log\` directory
- Verify source paths in `backup.ini`
- Ensure sufficient disk space on target drive

**Old files not being versioned**
- Check if files actually changed (size or date)
- Identical files are skipped (not an error)

**Task not running automatically**
- Verify task exists: `Get-ScheduledTask -TaskName 'jBackup'`
- Check Task Scheduler History for errors
- Ensure user account has necessary permissions

## Technical Details

### System Requirements
- Windows 7 or later
- PowerShell 5.1+ (for setup)
- VBScript (built into Windows)
- Write access to target backup location

### Performance
- Only changed files are copied (smart comparison)
- Recursive subfolder processing
- No compression (direct file copies)
- Speed depends on file count and sizes

### Security
- Task runs with current user privileges
- S4U authentication (no password stored)
- Highest privilege level (optional admin tasks)
- No network access required

### Limitations
- Windows only (VBScript dependency)
- No compression or encryption
- No network backup validation
- Manual recovery required for file restoration

## License

This project is provided as-is for personal and commercial use.

## Support

For issues or questions:
1. Check the log files in `log\` directory
2. Verify configuration in `backup.ini`
3. Review Task Scheduler history
4. Ensure all prerequisites are met

## Version History

### 1.0.0
- Multi-folder backup support
- Smart subfolder exclusion
- Custom target folder mapping
- PowerShell configuration assistant
- UTF-8 logging with statistics
- Dual-trigger scheduling (daily + startup)
- Power management awareness
- Queue mode for missed backups

### 1.0.1
- Added JBCKPV_ prefix to versioned files (JBCKPV_[YYYYMMDD]_filename.ext) to avoid conflict with files already prefixed by a date
- Manage invalid filenames ignoring files like document.pdf or photo2.jpg (wait for file to be renamed before backing it up)
- Added option to move files instead of copying

### 1.0.2
- Changed versioning logic.. Versioned files now stored in `.ver` directory at backup root (keeps current backups clean)
- Simplified version prefix to date only (YYYYMMDD_filename.ext)
- Added cleanup.vbs script to remove old versions (3 version limit)
- Created jBackupCleanup scheduled task to run cleanup.vbs daily at 2:00 PM
- Added empty folder removal in cleanup task (removes empty directories after version cleanup)

### TODO 
- Calculate total backup size and duration
- Calculate disk space used by backups
- Calculate available disk space before backup starts

