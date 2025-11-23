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

    logFile = yearLog & monthLog & dayLog & "-" & hourLog & minuteLog & secondLog & "-cleanup.log"
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

Sub CleanOldVersions(verPath, originalFileName)
    ' Keep only the 3 most recent versions of a file
    Dim verFolder, file, versionFiles()
    Dim versionCount, i, j, tempFile
    
    If Not FSO.FolderExists(verPath) Then
        Exit Sub
    End If
    
    Set verFolder = FSO.GetFolder(verPath)
    versionCount = 0
    
    ' Collect all version files for this filename
    For Each file In verFolder.Files
        ' Match files with YYYYMMDD_ prefix followed by the original filename
        If Right(file.Name, Len(originalFileName)) = originalFileName And _
           Len(file.Name) >= (9 + Len(originalFileName)) And _
           Mid(file.Name, 9, 1) = "_" And _
           IsNumeric(Left(file.Name, 8)) Then
            ReDim Preserve versionFiles(versionCount)
            Set versionFiles(versionCount) = file
            versionCount = versionCount + 1
        End If
    Next
    
    ' Sort version files by date (bubble sort, newest first)
    If versionCount > 1 Then
        For i = 0 To versionCount - 2
            For j = i + 1 To versionCount - 1
                If versionFiles(i).DateLastModified < versionFiles(j).DateLastModified Then
                    Set tempFile = versionFiles(i)
                    Set versionFiles(i) = versionFiles(j)
                    Set versionFiles(j) = tempFile
                End If
            Next
        Next
    End If
    
    ' Delete versions older than the 3 most recent
    If versionCount > 3 Then
        For i = 3 To versionCount - 1
            FSO.DeleteFile versionFiles(i).Path
            logContent = logContent & "[DELETE] " & versionFiles(i).Path & vbCrLf
            statsFilesDeleted = statsFilesDeleted + 1
        Next
    End If
End Sub

Sub RemoveEmptyFolders(folderPath)
    ' Recursively remove empty folders
    Dim folder, subFolder, subfoldersArray, i, subfolderCount
    
    If Not FSO.FolderExists(folderPath) Then
        Exit Sub
    End If
    
    Set folder = FSO.GetFolder(folderPath)
    
    ' First, process all subfolders recursively (bottom-up approach)
    ' We need to store subfolders in an array because we can't delete while iterating
    subfolderCount = 0
    For Each subFolder In folder.SubFolders
        subfolderCount = subfolderCount + 1
    Next
    
    If subfolderCount > 0 Then
        ReDim subfoldersArray(subfolderCount - 1)
        i = 0
        For Each subFolder In folder.SubFolders
            subfoldersArray(i) = subFolder.Path
            i = i + 1
        Next
        
        ' Process each subfolder
        For i = 0 To subfolderCount - 1
            RemoveEmptyFolders subfoldersArray(i)
        Next
    End If
    
    ' After processing subfolders, check if this folder is now empty
    Set folder = FSO.GetFolder(folderPath)
    If folder.SubFolders.Count = 0 And folder.Files.Count = 0 Then
        FSO.DeleteFolder folderPath
        logContent = logContent & "[RMDIR] " & folderPath & vbCrLf
        statsFoldersDeleted = statsFoldersDeleted + 1
    End If
End Sub

Sub CleanVersionDirectory(verPath)
    ' Recursively clean all subdirectories in .ver folder
    Dim verFolder, subFolder, file, filesDict, uniqueFiles
    
    If Not FSO.FolderExists(verPath) Then
        Exit Sub
    End If
    
    Set verFolder = FSO.GetFolder(verPath)
    
    ' Create dictionary to track unique base filenames
    Set filesDict = CreateObject("Scripting.Dictionary")
    
    ' Collect all version files and extract original filenames
    For Each file In verFolder.Files
        ' Extract original filename (remove YYYYMMDD_ prefix)
        Dim fileName, originalName
        fileName = file.Name
        
        ' Check if it matches version pattern: YYYYMMDD_filename.ext
        If Len(fileName) >= 10 And Mid(fileName, 9, 1) = "_" And IsNumeric(Left(fileName, 8)) Then
            originalName = Mid(fileName, 10) ' Everything after YYYYMMDD_
            
            If Not filesDict.Exists(originalName) Then
                filesDict.Add originalName, True
            End If
        End If
    Next
    
    ' Clean versions for each unique filename
    For Each uniqueFile In filesDict.Keys
        CleanOldVersions verPath, uniqueFile
    Next
    
    ' Recursively process subdirectories
    For Each subFolder In verFolder.SubFolders
        CleanVersionDirectory subFolder.Path
    Next
End Sub

logContent = logContent & "Cleanup started at: " & Now & vbCrLf & vbCrLf
logContent = logContent & "Working directory: " & scriptPath & vbCrLf
logContent = logContent & "Config file: " & configPath & vbCrLf

If Not FSO.FileExists(configPath) Then
    FatalError "Configuration file not found"
End If

' Read target path from config
Dim targetPath
targetPath = ""

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
                    Exit Do
                End If
            End If
        End If
    End If
Loop

configFileObj.Close

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

' Check if .ver directory exists
Dim verRootPath
verRootPath = targetPath & "\.ver"

If Not FSO.FolderExists(verRootPath) Then
    logContent = logContent & "No .ver directory found. Nothing to clean." & vbCrLf
    SaveLog()
    WScript.Quit 0
End If

logContent = logContent & vbCrLf & "=-=-=-=-=-=-=-=-=-="
logContent = logContent & vbCrLf & "||    CLEANUP    ||"
logContent = logContent & vbCrLf & "=-=-=-=-=-=-=-=-=-=" & vbCrLf & vbCrLf

Dim statsFilesDeleted, statsFoldersDeleted
statsFilesDeleted = 0
statsFoldersDeleted = 0

' Clean version directory
logContent = logContent & "Cleaning version directory: " & verRootPath & vbCrLf
CleanVersionDirectory verRootPath

logContent = logContent & vbCrLf & "Removing empty folders..." & vbCrLf
RemoveEmptyFolders verRootPath

logContent = logContent & vbCrLf & "===== STATISTICS =====" & vbCrLf
logContent = logContent & "Deleted old versions: " & statsFilesDeleted & vbCrLf
logContent = logContent & "Removed empty folders: " & statsFoldersDeleted & vbCrLf & vbCrLf

logContent = logContent & vbCrLf & "SUCCESS" & vbCrLf
logContent = logContent & "Cleanup completed at: " & Now & vbCrLf

SaveLog()
