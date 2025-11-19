Set objShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

Dim logContent
Dim scriptPath, configPath, logPath

scriptPath = FSO.GetParentFolderName(WScript.ScriptFullName)
configPath = scriptPath & "\backup.ini"
logPath = scriptPath & "\log"

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

Sub CopySubFolder(sourceFolder, cheminDestParent)
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
            
            ' Copy files from the subfolder
            For Each file In subFolder.Files
                targetFile = targetSubFolderPath & "\" & file.Name
                
                ' Handle existing file (rename with date)
                newFileVersion = PrefixExistingFile(file.Path, targetFile)
                
                ' Copy the new file if necessary
                If newFileVersion Then
                    FSO.CopyFile file.Path, targetFile
                    logContent = logContent & "[MV] " & targetFile & vbCrLf
                    statsFileNew = statsFileNew + 1
                End If
            Next
            
            ' Recursion for nested subfolders
            CopySubFolder subFolder, targetSubFolderPath
        End If
    Next
End Sub

Sub FatalError(message)
    logContent = logContent & vbCrLf & "FATAL ERROR " & message
    SaveLog()
    WScript.Quit 1
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

Function PrefixExistingFile(srcFile, targetFile)
    ' Returns True if the existing file was renamed, False otherwise
    PrefixExistingFile = False
    
    If FSO.FileExists(targetFile) Then
        Set srcFileObj = FSO.GetFile(srcFile)
        Set targetFileObj = FSO.GetFile(targetFile)
        
        ' Check if the files are identical (same size and modification date)
        If targetFileObj.Size = srcFileObj.Size And _
            DateDiff("s", targetFileObj.DateLastModified, srcFileObj.DateLastModified) = 0 Then
            ' Files are identical, ignore
            logContent = logContent & "[IGNORE] " & targetFile & vbCrLf
            statsFileIgnore = statsFileIgnore + 1
            PrefixExistingFile = False
            Exit Function
        End If
        
        ' Different file, rename the old one with date prefix
        Dim modDate, year, month, day, prefixDate
        modDate = targetFileObj.DateLastModified
        
        ' Format the date as YYYYMMDD
        year = Year(modDate)
        month = Right("0" & Month(modDate), 2)
        day = Right("0" & Day(modDate), 2)
        prefixDate = year & month & day
        
        ' Extract the path and file name
        parentPath = FSO.GetParentFolderName(targetFile)
        fileName = FSO.GetFileName(targetFile)
        
        ' Create the new name with date prefix
        newName = prefixDate & "_" & fileName
        newPath = parentPath & "\" & newName
        
        ' Rename the old file with the date
        targetFileObj.Move newPath
        logContent = logContent & "[BACKUP] " & targetFile & vbCrLf
        statsFileUpdate = statsFileUpdate + 1
        PrefixExistingFile = True
    Else
        ' New file
        PrefixExistingFile = True
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

logContent = logContent & "Parsing config file ...." & vbCrLf

Dim targetPath, sourceCount
Dim listeSources()
ReDim listeSources(0)

targetPath = ""
sourceCount = 0

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

logContent = logContent & "Target path: " & targetPath & vbCrLf

If Not FSO.DriveExists(FSO.GetDriveName(targetPath)) Then
    FatalError "Target drive is not available: " & FSO.GetDriveName(targetPath)
End If

If Not FSO.FolderExists(targetPath) Then
    FatalError "Target path does not exist: " & targetPath
End If

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
logContent = logContent & "  KO: " & sourceCountKO & vbCrLf

logContent = logContent & vbCrLf & "=-=-=-=-=-=-=-=-=-="
logContent = logContent & vbCrLf & "||    BACKUP     ||"
logContent = logContent & vbCrLf & "=-=-=-=-=-=-=-=-=-=" & vbCrLf & vbCrLf

Dim sourceFolderName, targetFolderName, targetFullPath
Dim statsFolderNew, statsFolderSaved, statsFolderSkip
Dim statsFileNew, statsFileUpdate, statsFileIgnore
statsFolderNew = 0
statsFolderSaved = 0
statsFolderSkip = 0
statsFileNew = 0
statsFileUpdate = 0
statsFileIgnore = 0

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
    ' copy files from folder
    For Each file In folderObj.Files
        targetFile = targetFullPath & "\" & file.Name
        
        ' Check for new version of the file
        newFileVersion = PrefixExistingFile(file.Path, targetFile)
        
        ' Copy the new file if necessary
        If newFileVersion Then
            FSO.CopyFile file.Path, targetFile
            logContent = logContent & "[NEW] " & targetFile & vbCrLf
            statsFileNew = statsFileNew + 1
        End If
    Next
    
    ' Copy subfolders and their contents (recursive)
    CopySubFolder folderObj, targetFullPath
    
    statsFolderSaved = statsFolderSaved + 1
Next

logContent = logContent & vbCrLf & "===== STATISTICS =====" & vbCrLf
logContent = logContent & "Processed folders  : " & statsFolderSaved & vbCrLf
logContent = logContent & "New folders        : " & statsFolderNew & vbCrLf
logContent = logContent & "Skipped folders    : " & statsFolderSkip & vbCrLf
logContent = logContent & "New files          : " & statsFileNew & vbCrLf
logContent = logContent & "Modified files     : " & statsFileUpdate & vbCrLf
logContent = logContent & "Ignored files      : " & statsFileIgnore & vbCrLf & vbCrLf

If sourceCountKO > 0 Then
    logContent = logContent & vbCrLf & "ERRORS" & vbCrLf 
Else
    logContent = logContent & vbCrLf & "SUCCESS" & vbCrLf
End If

SaveLog() 
