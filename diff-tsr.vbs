' extensions: tsr
'
' TortoiseSVN Diff script for shared object repository files
'
' Author: Ethan Bell, 06 May 2015

Dim objArgs, oFile, oFileSystem, objRepo, oShell
Dim sBase, sBaseXML, sNew, sNewCopy, sNewXML, strWinMerge, tempFolder
Const TEMP = 2
Const READ_ONLY = 1

'Verify that enough arguments are present
Set objArgs = WScript.Arguments
if objArgs.Count < 2 then
    MsgBox "Usage: [CScript | WScript] compare.vbs base.tsr new.tsr", vbCritical, "Invalid arguments"
    WScript.Quit 1
end if

'Get name of base file (usually from repository), and new file (usually working copy)
sBase = objArgs(0)
sNew = objArgs(1)
Set objArgs = Nothing

'Gets the location of the temp folder, where the XML files and the copies of local .tsrs will go
tempFolder = WScript.CreateObject("Scripting.FileSystemObject").GetSpecialFolder(TEMP)

'Your working copy should be copied to the temp folder, so ExportToXML doesn't show it as modified
Set oFileSystem = CreateObject("Scripting.FileSystemObject")
If InStr(1, sNew, tempFolder, 1) =  0 Then 'Don't copy if the 'new' file is already in the Temp folder
    sNewCopy = tempFolder & "\tsr-diff-newCopy.tsr"
    Call oFileSystem.CopyFile(sNew, sNewCopy, True) 'Source, destination, overwrite
Else
    sNewCopy = sNew 'sNewCopy will be the location used for new file in the rest of the script
End If

'Either file will be read-only when pulled from repo. Open write permission for exportToXML
Set oFile = oFileSystem.GetFile(sBase)
'Remove read only setting if enabled
If oFile.Attributes AND READ_ONLY Then
    oFile.Attributes = oFile.Attributes XOR READ_ONLY
End If
Set oFile = Nothing
Set oFile = oFileSystem.GetFile(sNewCopy)
'Remove read only setting if enabled on the second file as well
If oFile.Attributes AND READ_ONLY Then
    oFile.Attributes = oFile.Attributes XOR READ_ONLY
End If
Set oFile = Nothing

'File name for XML output
sBaseXML = tempFolder & "\tsr-diff_baseOR.xml"
sNewXML = tempFolder & "\tsr-diff_newOR.xml"

'Delete existing XML files so the new ones can be created
Set oFileSystem = CreateObject("Scripting.FileSystemObject")
If oFileSystem.FileExists(sBaseXML) Then
    oFileSystem.DeleteFile(sBaseXML)
End If

If oFileSystem.FileExists(sNewXML) Then
    oFileSystem.DeleteFile(sNewXML)
End If
Set oFileSystem = Nothing

'UFT makes a utility available to work with object repository files in VBScript
Set objRepo = CreateObject("Mercury.ObjectRepositoryUtil")
Call objRepo.ExportToXML(sBase, sBaseXML)
Call objRepo.ExportToXML(sNewCopy, sNewXML) 
Set objRepo = Nothing

'Open WinMerge to compare the two XML files
Set oShell = WScript.CreateObject("WScript.Shell")
strWinMerge = Chr(34) & "C:\Program Files (x86)\WinMerge\WinMergeU.exe" & Chr(34) _
              & " -e -ub -dl Base -dr New " & sBaseXML & " " & sNewXML & " -wl"
oShell.Run(strWinMerge)
Set oShell = Nothing