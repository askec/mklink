' ====================================================================================
' # mklink
' Stand alone and Total Commander interface for Windows mklink (NTFS symbolic links, hardlink and junction points)
' 
' ## Description
' 
' * VBScript to create symbolic, hardlinks or directory junctions via a rudimentary GUI that needs parameters.
' * Compatible with modern Windows versions, including Windows 11.
' * Supports multiple files/directory at once to one destination folder.
' * Make hard link or directory junction instead of symbolic links if file or directory is on the same volume.
' 
' Note: This script requires the 'mklink_gui.hta' file to be in the same directory.
' 
' ## Usage
' cscript /nologo mklink.vbs "C:\Path\To\DestinationFolder" "C:\Path\To\ListFile.txt"
' 
' ### Parameters
' `%1`: Destination Folder: The folder where the new links will be created.
' `%2`: Source List File: A text file containing one source file or folder path per line.
' 
' ### Exit Codes
' `0`: Success, no warnings.
' `1`: Script was cancelled by user.
' `2`: The base destination directory does not exist.
' `3`: Success, but with one or more warnings (e.g., source not found, destination exists).
' 
' ## Installation
' 
' 1. Just download all the files from src folder. 
' 1. Copy all files in a folder anywhere to your disk.
' 1. Configure a command or a single button in Total Commander :
' 
' Command: `cscript`
' Parameters: `/noLogo "<path_of_mklink.vbs>\mklink.vbs" "%T" "%L"`
' Icon file: `<path_of_mklink.vbs>\mklink.ico`
' Tooltip: `Make NTFS Link`
' 
' `<path_of_mklink.vbs>` is the path where you copied mklink.vbs.
'
' MIT License
' (c) 2025 askec - https://github.com/askec/mklink
'
' ====================================================================================

Option Explicit
On Error Resume Next

Dim objFSO, objShell
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

Call Main()

Sub Main()
    If WScript.Arguments.Count <> 2 Then
        MsgBox "This script requires exactly two parameters:" & vbCrLf & vbCrLf & _
               "1. The destination FOLDER path (in quotes)." & vbCrLf & _
               "2. The path to the source list file (in quotes).", _
               vbCritical, "Invalid Arguments"
        WScript.Quit(1)
    End If

    Dim strInitialDestPath, strSourceListFile
    strInitialDestPath = WScript.Arguments(0)
    strSourceListFile = WScript.Arguments(1)

    If Not objFSO.FileExists(strSourceListFile) Then
        MsgBox "The source list file does not exist:" & vbCrLf & strSourceListFile, vbCritical, "File Not Found"
        WScript.Quit(1)
    End If

    Dim strFinalDestFolder, bUseHardLinks
    Dim arrResults
    arrResults = LaunchHTA(strInitialDestPath)

    If IsArray(arrResults) And arrResults(0) = "CANCEL" Then
        WScript.Quit(1) ' Exit with a non-zero code to indicate cancellation
    End If

    strFinalDestFolder = arrResults(0)
    bUseHardLinks = CBool(arrResults(1))

    If Not objFSO.FolderExists(strFinalDestFolder) Then
         MsgBox "The destination location '" & strFinalDestFolder & "' does not exist. Cannot proceed.", vbCritical, "Location Not Found"
         WScript.Quit(2)
    End If

    Dim bHadWarnings
    bHadWarnings = ProcessSourceList(strSourceListFile, strFinalDestFolder, bUseHardLinks)

    If bHadWarnings Then
        WScript.Quit(3)
    Else
        WScript.Quit(0)
    End If
End Sub

Function ProcessSourceList(strListFile, strDestFolder, bHardLinks)
    ProcessSourceList = False
    Dim objFile, strLine, strSourceVolume, strDestVolume, strCommand, strFileName, strFullDestPath

    strDestVolume = GetVolume(strDestFolder)
    Set objFile = objFSO.OpenTextFile(strListFile, 1)

    Do While Not objFile.AtEndOfStream
        strLine = Trim(objFile.ReadLine)
        If strLine <> "" Then

            strFileName = objFSO.GetFileName(strLine)
            strFullDestPath = objFSO.BuildPath(strDestFolder, strFileName)

            If objFSO.FileExists(strFullDestPath) Or objFSO.FolderExists(strFullDestPath) Then
                MsgBox "Warning: Destination '" & strFullDestPath & "' already exists. Skipping.", vbExclamation, "Destination Exists"
                ProcessSourceList = True
            ElseIf objFSO.FileExists(strLine) Then
                strSourceVolume = GetVolume(strLine)
                If bHardLinks Then
                    If LCase(strSourceVolume) = LCase(strDestVolume) Then
                        strCommand = "cmd /c mklink /H " & Chr(34) & strFullDestPath & Chr(34) & " " & Chr(34) & strLine & Chr(34)
                        objShell.Run strCommand, 0, True
                    Else
                        MsgBox "Error: The source file is not on the same volume as the destination." & vbCrLf & _
                               "Source: " & strLine & " (" & strSourceVolume & ")" & vbCrLf & _
                               "Destination: " & strDestFolder & " (" & strDestVolume & ")", _
                               vbCritical, "Volume Mismatch"
                        ProcessSourceList = True
                    End If
                Else
                    strCommand = "cmd /c mklink " & Chr(34) & strFullDestPath & Chr(34) & " " & Chr(34) & strLine & Chr(34)
                    objShell.Run strCommand, 0, True
                End If

            ElseIf objFSO.FolderExists(strLine) Then
                strSourceVolume = GetVolume(strLine)
                If bHardLinks Then
                     If LCase(strSourceVolume) = LCase(strDestVolume) Then
                        strCommand = "cmd /c mklink /J " & Chr(34) & strFullDestPath & Chr(34) & " " & Chr(34) & strLine & Chr(34)
                        objShell.Run strCommand, 0, True
                    Else
                        MsgBox "Warning: The source directory is not on the same volume as the destination." & vbCrLf & _
                               "Source: " & strLine & " (" & strSourceVolume & ")" & vbCrLf & _
                               "Destination: " & strDestFolder & " (" & strDestVolume & ")", _
                               vbExclamation, "Volume Mismatch"
                        ProcessSourceList = True
                    End If
                Else
                    strCommand = "cmd /c mklink /D " & Chr(34) & strFullDestPath & Chr(34) & " " & Chr(34) & strLine & Chr(34)
                    objShell.Run strCommand, 0, True
                End If
            Else
                MsgBox "Warning: The source file or directory does not exist and was skipped:" & vbCrLf & strLine, vbExclamation, "Source Not Found"
                ProcessSourceList = True
            End If
        End If
    Loop
    objFile.Close
End Function

Function GetVolume(strPath)
    On Error Resume Next
    GetVolume = objFSO.GetDriveName(strPath)
    If Err.Number <> 0 Then GetVolume = ""
    On Error GoTo 0
End Function

Function LaunchHTA(strInitialDest)
    Dim strScriptPath, strHTAPath, strResultFile, strCommand

    strScriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
    strHTAPath = objFSO.BuildPath(strScriptPath, "mklink_gui.hta")

    If Not objFSO.FileExists(strHTAPath) Then
        MsgBox "Error: The required GUI file 'mklink_gui.hta' was not found." & vbCrLf & _
               "Please ensure it is in the same directory as the script.", vbCritical, "File Missing"
        LaunchHTA = Array("CANCEL", False)
        Exit Function
    End If

    strResultFile = objFSO.BuildPath(objFSO.GetSpecialFolder(2), objFSO.GetTempName & ".txt")

    strCommand = "mshta.exe " & Chr(34) & strHTAPath & Chr(34) & " " & Chr(34) & strInitialDest & Chr(34) & " " & Chr(34) & strResultFile & Chr(34)
    objShell.Run strCommand, 1, True

    If objFSO.FileExists(strResultFile) Then
        Dim objResultFile, strLine1, strLine2
        Set objResultFile = objFSO.OpenTextFile(strResultFile, 1)
        strLine1 = objResultFile.ReadLine
        If Not objResultFile.AtEndOfStream Then
            strLine2 = objResultFile.ReadLine
        Else
            strLine2 = "False" ' Default value if not present
        End If
        objResultFile.Close

        If UCase(strLine1) = "CANCEL" Then
            LaunchHTA = Array("CANCEL", False)
        Else
            LaunchHTA = Array(strLine1, CBool(strLine2))
        End If
        objFSO.DeleteFile strResultFile
    Else
        LaunchHTA = Array("CANCEL", False)
    End If
End Function
