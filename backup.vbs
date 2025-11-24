Set objShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

Dim logContent
Dim scriptPath, configPath, logPath

scriptPath = FSO.GetParentFolderName(WScript.ScriptFullName)
configPath = scriptPath & "\backup.ini"
invalidNamesPath = scriptPath & "\invalidnames.txt"
logPath = scriptPath & "\log"

' Ensure log directory exists
If Not FSO.FolderExists(logPath) Then
    FSO.CreateFolder(logPath)
End If

Sub SaveLog()
    nowTime = Now
    yearLog = Year(nowTime)
    monthLog = Right("0" & Month(nowTime), 2)
    dayLog = Right("0" & Day(nowTime), 2)
    hourLog = Right("0" & Hour(nowTime), 2)
    minuteLog = Right("0" & Minute(nowTime), 2)
    secondLog = Right("0" & Second(nowTime), 2)

    logFile = yearLog & monthLog & dayLog & "-" & hourLog & minuteLog & secondLog & ".log"
    logFilePath = logPath & "\" & logFile

    Set fileLogObj = FSO.CreateTextFile(logFilePath, True)
    fileLogObj.Write logContent
    fileLogObj.Close
End Sub

Sub FatalError(message)
    logContent = logContent & vbCrLf & "FATAL ERROR " & message
    SaveLog()
    WScript.Quit 1
End Sub

Function GetDriveSpace(drivePath)
    ' Returns available space on drive in bytes
    Dim drive
    Set drive = FSO.GetDrive(FSO.GetDriveName(drivePath))
    GetDriveSpace = drive.AvailableSpace
End Function

Function FormatBytes(bytes)
    ' Format bytes to human-readable format
    Dim result
    If bytes >= 1073741824 Then ' 1 GB
        result = FormatNumber(bytes / 1073741824, 2) & " GB"
    ElseIf bytes >= 1048576 Then ' 1 MB
        result = FormatNumber(bytes / 1048576, 2) & " MB"
    ElseIf bytes >= 1024 Then ' 1 KB
        result = FormatNumber(bytes / 1024, 2) & " KB"
    Else
        result = bytes & " bytes"
    End If
    FormatBytes = result
End Function

Function GetFolderSize(folderPath)
    ' Recursively calculate folder size in bytes
    Dim folder, file, subFolder, totalSize
    totalSize = 0
    
    If Not FSO.FolderExists(folderPath) Then
        GetFolderSize = 0
        Exit Function
    End If
    
    Set folder = FSO.GetFolder(folderPath)
    
    ' Add file sizes
    For Each file In folder.Files
        totalSize = totalSize + file.Size
    Next
    
    ' Add subfolder sizes recursively
    For Each subFolder In folder.SubFolders
        totalSize = totalSize + GetFolderSize(subFolder.Path)
    Next
    
    GetFolderSize = totalSize
End Function

Function FormatDuration(seconds)
    ' Format duration in seconds to human-readable format
    Dim hours, minutes, secs, result
    hours = Int(seconds / 3600)
    minutes = Int((seconds Mod 3600) / 60)
    secs = seconds Mod 60
    
    result = ""
    If hours > 0 Then
        result = hours & "h "
    End If
    If minutes > 0 Or hours > 0 Then
        result = result & minutes & "m "
    End If
    result = result & secs & "s"
    
    FormatDuration = result
End Function

Sub BackupSubFolders(sourceFolder, cheminDestParent)
    For Each subFolder In sourceFolder.Subfolders
        ' Check if this subfolder is already being processed as a separate source
        If IsSubfolderAlreadyProcessed(subFolder.Path) Then
            logContent = logContent & "[SKIP] " & subFolder.Path & " (already processed separately)" & vbCrLf
            statsFolderSkip = statsFolderSkip + 1
        Else
            targetSubFolderPath = cheminDestParent & "\" & subFolder.Name
            
            ' Create the subfolder in the target if it doesn't exist
            If Not FSO.FolderExists(targetSubFolderPath) Then
                FSO.CreateFolder(targetSubFolderPath)
                logContent = logContent & "[MKDIR] " & targetSubFolderPath & vbCrLf
                statsFolderNew = statsFolderNew + 1
            End If
            
            BackupFolderFiles subFolder.Files, targetSubFolderPath
            BackupSubFolders subFolder, targetSubFolderPath
        End If
    Next
End Sub

Sub BackupFolderFiles(files, currentPath)
    ' copy files from folder
    For Each file In files
        If IsGenericFilename(FSO.GetBaseName(file.Name)) Then
            logContent = logContent & "[CHECK] " & file.Path & vbCrLf
            statsFileInvalid = statsFileInvalid + 1
        Else
            targetFile = currentPath & "\" & file.Name

            ' Check if the files are identical (same size and modification date)
            If isSameFile(file.Path, targetFile) Then
                ' Delete source file if moving and files are identical
                If moveFiles Then
                    logContent = logContent & "[DELETE] " & file.Path & " -> " & targetFile & vbCrLf
                    FSO.DeleteFile file.Path
                    statsFileDelete = statsFileDelete + 1
                Else
                    logContent = logContent & "[IGNORE] " & file.Path & " -> " & targetFile & vbCrLf
                    statsFileIgnore = statsFileIgnore + 1
                End If
            Else
                srcFile = file.Path
                isAnUpdate = FSO.FileExists(targetFile)
                
                ' Track bytes backed up
                statsTotalBytes = statsTotalBytes + file.Size

                If isAnUpdate Then 
                    ' Keep old version
                    KeepCopyOfPreviousFile srcFile, targetFile
                End If

                ' Copy or Move file
                If moveFiles Then
                    FSO.MoveFile srcFile, targetFile
                Else
                    FSO.CopyFile srcFile, targetFile
                End If
                
                If isAnUpdate Then
                    logContent = logContent & "[UPDATE] " & srcFile & " -> " & targetFile & vbCrLf
                    statsFileUpdate = statsFileUpdate + 1
                Else 
                    logContent = logContent & "[NEW] " & srcFile & " -> " & targetFile & vbCrLf
                    statsFileNew = statsFileNew + 1
                End If
            End If
        End If
    Next
End Sub

Sub KeepCopyOfPreviousFile(srcFile, targetFile)    
    Dim srcFileObj, targetFileObj
    Set srcFileObj = FSO.GetFile(srcFile)
    Set targetFileObj = FSO.GetFile(targetFile)
    
    ' Different file, rename the old one with date prefix
    Dim modDate, yearVal, monthVal, dayVal, prefixDate
    modDate = targetFileObj.DateLastModified
    
    ' Format the date as YYYYMMDD
    yearVal = Year(modDate)
    monthVal = Right("0" & Month(modDate), 2)
    dayVal = Right("0" & Day(modDate), 2)
    prefixDate = yearVal & monthVal & dayVal
    
    ' Extract the path and file name
    parentPath = FSO.GetParentFolderName(targetFile)
    fileName = FSO.GetFileName(targetFile)
    baseName = FSO.GetBaseName(fileName)
    
    ' Calculate relative path from target root to maintain structure in .ver
    Dim relativePath
    relativePath = Mid(parentPath, Len(targetPath) + 2) ' +2 to skip leading backslash
    
    ' Create .ver directory structure
    Dim verRootPath, verFilePath
    verRootPath = targetPath & "\.ver"
    If Not FSO.FolderExists(verRootPath) Then
        FSO.CreateFolder(verRootPath)
        logContent = logContent & "[MKDIR] " & verRootPath & vbCrLf
    End If
    
    ' Create subdirectory in .ver if needed
    If relativePath <> "" Then
        verFilePath = verRootPath & "\" & relativePath
        If Not FSO.FolderExists(verFilePath) Then
            ' Create directory recursively
            CreateDirectoryRecursive verFilePath
        End If
    Else
        verFilePath = verRootPath
    End If
    
    ' Create the new name with date prefix only
    newName = prefixDate & "_" & fileName
    newPath = verFilePath & "\" & newName
    
    ' Move the old file to .ver directory
    ' Check if a version with today's date already exists (useful for dev/testing)
    If FSO.FileExists(newPath) Then
        ' Delete existing version from today and replace it
        FSO.DeleteFile newPath
        logContent = logContent & "[DELETE] " & newPath & " (replacing with newer version)" & vbCrLf
    End If
    
    targetFileObj.Move newPath
    logContent = logContent & "[BACKUP] " & newPath & vbCrLf
End Sub

Sub CreateDirectoryRecursive(dirPath)
    ' Create directory recursively if it doesn't exist
    If Not FSO.FolderExists(dirPath) Then
        Dim parentDir
        parentDir = FSO.GetParentFolderName(dirPath)
        If Not FSO.FolderExists(parentDir) Then
            CreateDirectoryRecursive parentDir
        End If
        FSO.CreateFolder(dirPath)
        logContent = logContent & "[MKDIR] " & dirPath & vbCrLf
    End If
End Sub

Function PathToCamelCase(fullPath)
    ' Convert path to lowercase with underscores (e.g., "C:\Users\PC\Documents" -> "c_users_pc_documents")
    Dim pathWithoutDrive, parts, i, result, part, cleanPart, j, char, driveLetter
     
    ' Extract and store drive letter
    driveLetter = ""
    If InStr(fullPath, ":") > 0 Then
        driveLetter = UCase(Left(fullPath, 1))
        pathWithoutDrive = Mid(fullPath, InStr(fullPath, ":") + 1)
    Else
        pathWithoutDrive = fullPath
    End If
    
    ' Remove leading backslash
    If Left(pathWithoutDrive, 1) = "\" Then
        pathWithoutDrive = Mid(pathWithoutDrive, 2)
    End If
    
    ' Split by backslash
    parts = Split(pathWithoutDrive, "\")
    result = ""
    
    ' Add drive letter prefix if present
    If driveLetter <> "" Then
        result = driveLetter
    End If
    
    For i = 0 To UBound(parts)
        part = parts(i)
        If Len(part) > 0 Then
            ' Clean special characters
            cleanPart = ""
            For j = 1 To Len(part)
                char = Mid(part, j, 1)
                ' Keep only alphanumeric characters, replace others with hyphen
                If (char >= "a" And char <= "z") Or (char >= "A" And char <= "Z") Or (char >= "0" And char <= "9") Then
                    cleanPart = cleanPart & char
                Else
                    ' Replace special char with hyphen (unless it's already a hyphen or space)
                    If char <> "-" And char <> " " Then
                        cleanPart = cleanPart & "-"
                    ElseIf char = " " Then
                        cleanPart = cleanPart & "-"
                    Else
                        cleanPart = cleanPart & char
                    End If
                End If
            Next
            
            ' Remove trailing hyphens
            Do While Right(cleanPart, 1) = "-"
                cleanPart = Left(cleanPart, Len(cleanPart) - 1)
            Loop
            
            If Len(cleanPart) > 0 Then
                ' Add underscore separator
                If result <> "" Then
                    result = result & "_"
                End If
                ' Convert to lowercase
                result = result & LCase(cleanPart)
            End If
        End If
    Next
    
    PathToCamelCase = result
End Function

Function IsGenericFilename(baseName)
    ' Check if a filename uses generic non-descriptive names
    ' Returns True if generic, False otherwise
    
    ' Remove trailing digits to get clean base name
    Dim cleanBaseName
    cleanBaseName = baseName
    Do While Len(cleanBaseName) > 0 And IsNumeric(Right(cleanBaseName, 1))
        cleanBaseName = Left(cleanBaseName, Len(cleanBaseName) - 1)
    Loop
    cleanBaseName = LCase(Trim(cleanBaseName))
    
    ' Check against array of generic names
    Dim i
    For i = 0 To UBound(invalidFileNames)
        If cleanBaseName = invalidFileNames(i) Then
            IsGenericFilename = True
            Exit Function
        End If
    Next
    
    IsGenericFilename = False
End Function

Function isSameFile(srcFile, targetFile)
    Dim srcFileObj, targetFileObj
    
    If FSO.FileExists(srcFile) And FSO.FileExists(targetFile) Then 
        Set srcFileObj = FSO.GetFile(srcFile)
        Set targetFileObj = FSO.GetFile(targetFile)
        
        ' Check if the files are identical (same size and modification date)
        If targetFileObj.Size = srcFileObj.Size And _
            DateDiff("s", targetFileObj.DateLastModified, srcFileObj.DateLastModified) = 0 Then
            isSameFile = True
        Else 
            isSameFile = False
        End If
    Else
        isSameFile = False
    End If
End Function

Function IsSubfolderAlreadyProcessed(subfolderPath)
    ' Check if the subfolder path is already in the validSources list
    Dim i, normalizedSubPath, normalizedSourcePath
    
    ' Normalize the subfolder path (remove trailing backslash and convert to lowercase)
    normalizedSubPath = subfolderPath
    If Right(normalizedSubPath, 1) = "\" Then
        normalizedSubPath = Left(normalizedSubPath, Len(normalizedSubPath) - 1)
    End If
    normalizedSubPath = LCase(normalizedSubPath)
    
    For i = 0 To UBound(validSources)
        ' Normalize source path
        normalizedSourcePath = validSources(i)
        If Right(normalizedSourcePath, 1) = "\" Then
            normalizedSourcePath = Left(normalizedSourcePath, Len(normalizedSourcePath) - 1)
        End If
        normalizedSourcePath = LCase(normalizedSourcePath)
        
        ' Compare paths
        If normalizedSubPath = normalizedSourcePath Then
            IsSubfolderAlreadyProcessed = True
            Exit Function
        End If
    Next
    
    IsSubfolderAlreadyProcessed = False
End Function

logContent = logContent & "Started at: " & Now & vbCrLf & vbCrLf
logContent = logContent & "Working directory: " & scriptPath & vbCrLf
logContent = logContent & "Config file: " & configPath & vbCrLf

If Not FSO.FileExists(configPath) Then
    FatalError "Configuration file not found"
End If

' Check if invalid names configuration file exists
If Not FSO.FileExists(invalidNamesPath) Then
    FatalError "Invalid names file not found: " & invalidNamesPath
End If

logContent = logContent & "Parsing config file ...." & vbCrLf

Dim targetPath, sourceCount, moveFiles
Dim invalidFileNames()
ReDim invalidFileNames(0)
Dim listeSources()
ReDim listeSources(0)

targetPath = ""
sourceCount = 0
moveFiles = False

Set configFileObj = FSO.OpenTextFile(configPath, 1)

Dim ligne, posEqual, cle, valeur
Do Until configFileObj.AtEndOfStream
    ligne = Trim(configFileObj.ReadLine)
    
    If Len(ligne) > 0 And Left(ligne, 1) <> ";" And Left(ligne, 1) <> "[" Then
        posEqual = InStr(ligne, "=")
        If posEqual > 0 Then
            cle = Trim(Left(ligne, posEqual - 1))
            valeur = Trim(Mid(ligne, posEqual + 1))
            
            If cle = "targetPath" Then
                If valeur <> "" Then
                    targetPath = valeur
                End If
            ElseIf cle = "moveFiles" Then
                If LCase(valeur) = "1" Then
                    moveFiles = True
                End If
            ElseIf Left(cle, 12) = "sourceFolder" Then
                ' Only add non-empty source folders
                If valeur <> "" Then

                    If sourceCount = 0 Then
                        listeSources(0) = valeur
                    Else
                        ReDim Preserve listeSources(sourceCount)
                        listeSources(sourceCount) = valeur
                    End If
                    sourceCount = sourceCount + 1
                End If
            End If
        End If
    End If
Loop

configFileObj.Close

' check targetPath
If targetPath = "" Then
    FatalError "Target folder not defined in the configuration file"
End If 

Dim moveFilesReadable
If moveFiles Then
    moveFilesReadable = "Yes"
Else
    moveFilesReadable = "No"
End If

logContent = logContent & "Target path: " & targetPath & vbCrLf
logContent = logContent & "Move files: " & moveFilesReadable & vbCrLf & vbCrLf

If Not FSO.DriveExists(FSO.GetDriveName(targetPath)) Then
    FatalError "Target drive is not available: " & FSO.GetDriveName(targetPath)
End If

If Not FSO.FolderExists(targetPath) Then
    FatalError "Target path does not exist: " & targetPath
End If

' Check available disk space
Dim availableSpaceBefore, backupSizeBefore
availableSpaceBefore = GetDriveSpace(targetPath)
backupSizeBefore = GetFolderSize(targetPath)

logContent = logContent & "Available disk space: " & FormatBytes(availableSpaceBefore) & vbCrLf
logContent = logContent & "Current backup size: " & FormatBytes(backupSizeBefore) & vbCrLf & vbCrLf

' Start timing
Dim startTime, endTime, duration
startTime = Timer

' check sources
If sourceCount = 0 Then
    FatalError "No sources defined in the configuration file" & vbCrLf
End If

logContent = logContent & "Check sources ...." & vbCrLf

Dim i, sourceCountOK, sourceCountKO
Dim validSources()

ReDim validSources(0)
sourceCountOK = 0
sourceCountKO = 0

' Also store target subfolders
Dim validTargetSubfolders()
ReDim validTargetSubfolders(0)

For i = 0 To sourceCount - 1
    ' Parse source entry (format: "path" or "path,targetSubfolder")
    Dim sourceEntry, commaPos, sourcePath, targetSubfolder
    sourceEntry = listeSources(i)
    commaPos = InStr(sourceEntry, ",")
    
    If commaPos > 0 Then
        sourcePath = Trim(Left(sourceEntry, commaPos - 1))
        targetSubfolder = Trim(Mid(sourceEntry, commaPos + 1))
    Else
        sourcePath = Trim(sourceEntry)
        targetSubfolder = ""
    End If
    
    logContent = logContent & "Source " & (i+1) & ": " & sourcePath
    If targetSubfolder <> "" Then
        logContent = logContent & " -> " & targetSubfolder & "/"
    End If
    
    If Not FSO.FolderExists(sourcePath) Then
        logContent = logContent & " .. [KO]" & vbCrLf
        sourceCountKO = sourceCountKO + 1
    Else
        logContent = logContent & " .. [OK]" & vbCrLf
        If sourceCountOK = 0 Then
            validSources(0) = sourcePath
            validTargetSubfolders(0) = targetSubfolder
        Else
            ReDim Preserve validSources(sourceCountOK)
            ReDim Preserve validTargetSubfolders(sourceCountOK)
            validSources(sourceCountOK) = sourcePath
            validTargetSubfolders(sourceCountOK) = targetSubfolder
        End If
        sourceCountOK = sourceCountOK + 1
    End If
Next

If sourceCountOK = 0 Then
    FatalError "No valid source folders"
End If

logContent = logContent & vbCrLf & "Sources: " & sourceCount & vbCrLf
logContent = logContent & "  OK: " & sourceCountOK & vbCrLf
logContent = logContent & "  KO: " & sourceCountKO & vbCrLf & vbCrLf

logContent = logContent & "Loading invalid filenames ...." & vbCrLf

' Load invalid filenames from file
Set invalidNamesFile = FSO.OpenTextFile(invalidNamesPath, 1)
Dim invalidCount
invalidCount = 0

Do Until invalidNamesFile.AtEndOfStream
    Dim nameLine
    nameLine = Trim(invalidNamesFile.ReadLine)
    
    ' Skip empty lines and comments
    If Len(nameLine) > 0 And Left(nameLine, 1) <> ";" And Left(nameLine, 1) <> "#" Then
        If invalidCount = 0 Then
            invalidFileNames(0) = LCase(nameLine)
        Else
            ReDim Preserve invalidFileNames(invalidCount)
            invalidFileNames(invalidCount) = LCase(nameLine)
        End If
        invalidCount = invalidCount + 1
    End If
Loop
invalidNamesFile.Close

logContent = logContent & "  |-> " & UBound(invalidFileNames) + 1 & " invalid filename patterns" & vbCrLf & vbCrLf

logContent = logContent & vbCrLf & "=-=-=-=-=-=-=-=-=-="
logContent = logContent & vbCrLf & "||    BACKUP     ||"
logContent = logContent & vbCrLf & "=-=-=-=-=-=-=-=-=-=" & vbCrLf & vbCrLf

Dim sourceFolderName, targetFolderName, targetFullPath
Dim statsFolderNew, statsFolderSaved, statsFolderSkip
Dim statsFileNew, statsFileUpdate, statsFileIgnore, statsFileDelete, statsFileInvalid
statsFolderNew = 0
statsFolderSaved = 0
statsFolderSkip = 0
statsFileNew = 0
statsFileUpdate = 0
statsFileIgnore = 0
statsFileDelete = 0
statsFileInvalid = 0
statsTotalBytes = 0

Dim sourceIndex
sourceIndex = 0

For Each sourceFolder In validSources
    logContent = logContent & " >> " & sourceFolder
    
    ' Get corresponding target subfolder
    Dim currentTargetSubfolder
    currentTargetSubfolder = validTargetSubfolders(sourceIndex)
    
    If currentTargetSubfolder <> "" Then
        logContent = logContent & " -> " & currentTargetSubfolder & vbCrLf
    Else
        logContent = logContent & vbCrLf
    End If

    ' extract folder name
    Set folderObj = FSO.GetFolder(sourceFolder)
    sourceFolderName = folderObj.Name
    
    ' set target path
    If currentTargetSubfolder <> "" Then
        ' Use specified target subfolder
        targetFullPath = targetPath & "\" & currentTargetSubfolder
    Else
        ' Use auto-generated name (fullpath without drive in lowercase with underscores)
        targetFolderName = PathToCamelCase(sourceFolder)
        targetFullPath = targetPath & "\" & targetFolderName
    End If
    
    ' create target path
    If Not FSO.FolderExists(targetFullPath) Then
        FSO.CreateFolder(targetFullPath)
        logContent = logContent & "[MKDIR] " & targetFullPath & vbCrLf
        statsFolderNew = statsFolderNew + 1
    End If
    
    sourceIndex = sourceIndex + 1

    BackupFolderFiles folderObj.Files, targetFullPath
    BackupSubFolders folderObj, targetFullPath
    
    statsFolderSaved = statsFolderSaved + 1
Next

' Calculate duration and final statistics
endTime = Timer
If endTime < startTime Then
    ' Handle midnight rollover
    duration = (86400 - startTime) + endTime
Else
    duration = endTime - startTime
End If

Dim availableSpaceAfter, backupSizeAfter, backupGrowth
availableSpaceAfter = GetDriveSpace(targetPath)
backupSizeAfter = GetFolderSize(targetPath)
backupGrowth = backupSizeAfter - backupSizeBefore

logContent = logContent & vbCrLf & "===== STATISTICS =====" & vbCrLf
logContent = logContent & "Duration           : " & FormatDuration(Int(duration)) & vbCrLf
logContent = logContent & "Processed folders  : " & statsFolderSaved & vbCrLf
logContent = logContent & "New folders        : " & statsFolderNew & vbCrLf
logContent = logContent & "Skipped folders    : " & statsFolderSkip & vbCrLf
logContent = logContent & "New files          : " & statsFileNew & vbCrLf
logContent = logContent & "Modified files     : " & statsFileUpdate & vbCrLf
logContent = logContent & "Ignored files      : " & statsFileIgnore & vbCrLf
logContent = logContent & "Invalid files      : " & statsFileInvalid & vbCrLf
logContent = logContent & "Deleted files      : " & statsFileDelete & vbCrLf
logContent = logContent & "Data backed up     : " & FormatBytes(statsTotalBytes) & vbCrLf
logContent = logContent & "Total backup size  : " & FormatBytes(backupSizeAfter) & vbCrLf
logContent = logContent & "Backup growth      : " & FormatBytes(backupGrowth) & vbCrLf
logContent = logContent & "Available space    : " & FormatBytes(availableSpaceAfter) & vbCrLf & vbCrLf

If sourceCountKO > 0 Then
    logContent = logContent & vbCrLf & "ERRORS" & vbCrLf 
Else
    logContent = logContent & vbCrLf & "SUCCESS" & vbCrLf
End If

SaveLog() 
