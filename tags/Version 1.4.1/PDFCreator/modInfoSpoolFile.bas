Attribute VB_Name = "modInfoSpoolFile"
Option Explicit

Public colInfoSpoolFiles As New Collection

' SpoolFileName = Postscript data file
Public Function CreateInfoSpoolFile(SpoolFileName As String, Optional InfoSpoolFileName, Optional ClientComputer, Optional DocumentTitle, _
 Optional JobID, Optional PrinterName, Optional SessionID, Optional UserName, Optional WinStation)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020  Dim ini As clsINI, Path As String, File As String
50030
50040  Set ini = New clsINI
50050  ini.Section = "1"
50060
50070  If Not IsMissing(InfoSpoolFileName) Then
50080    ini.filename = CStr(InfoSpoolFileName)
50090   Else
50100    SplitPath SpoolFileName, , Path, , File
50110    ini.filename = CompletePath(Path) & File & ".inf"
50120  End If
50130
50140  If Not IsMissing(ClientComputer) Then
50150    ini.SaveKey CStr(ClientComputer), "ClientComputer"
50160   Else
50170    ini.SaveKey GetComputerName, "ClientComputer"
50180  End If
50190  If Not IsMissing(DocumentTitle) Then
50200    ini.SaveKey CStr(DocumentTitle), "DocumentTitle"
50210   Else
50220    ini.SaveKey GetPSTitle(SpoolFileName), "DocumentTitle"
50230  End If
50240  If Not IsMissing(JobID) Then
50250    ini.SaveKey CStr(JobID), "JobID"
50260   Else
50270    ini.SaveKey "0", "JobID"
50280  End If
50290  If Not IsMissing(PrinterName) Then
50300    ini.SaveKey CStr(PrinterName), "PrinterName"
50310   Else
50320    ini.SaveKey "", "PrinterName"
50330  End If
50340  If Not IsMissing(SessionID) Then
50350    ini.SaveKey CStr(SessionID), "SessionID"
50360   Else
50370    ini.SaveKey "", "SessionID"
50380  End If
50390  ini.SaveKey SpoolFileName, "SpoolFileName"
50400  If Not IsMissing(UserName) Then
50410    ini.SaveKey CStr(UserName), "UserName"
50420   Else
50430    ini.SaveKey GetUsername, "UserName"
50440  End If
50450  If Not IsMissing(WinStation) Then
50460    ini.SaveKey CStr(WinStation), "WinStation"
50470   Else
50480    ini.SaveKey Environ$("SESSIONNAME"), "WinStation"
50490  End If
50500  CreateInfoSpoolFile = ini.filename
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modInfoSpoolFile", "CreateInfoSpoolFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub KillInfoSpoolFiles(InfoSpoolFileName As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim isf As clsInfoSpoolFile, isfi As clsInfoSpoolFileInfo, i As Long, Path As String, File As String, PDFInfoFileName As String, StampFileName As String
50020  Set isf = New clsInfoSpoolFile
50030  isf.ReadInfoFile InfoSpoolFileName
50040  For i = 1 To isf.InfoFiles.Count
50050   Set isfi = isf.InfoFiles(i)
50060   KillFile isfi.SpoolFileName
50070  Next i
50080
50090  KillFile InfoSpoolFileName
50100
50110  SplitPath InfoSpoolFileName, , Path, , File
50120  PDFInfoFileName = CompletePath(Path) & File & ".mtd"
50130  KillFile PDFInfoFileName
50140  StampFileName = CompletePath(Path) & File & ".stm"
50150  KillFile StampFileName
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modInfoSpoolFile", "KillInfoSpoolFiles")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function GetInfoSpoolFileObject(filename As String) As clsInfoSpoolFile
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim isf As clsInfoSpoolFile
50020  If CollectionItemExists(filename, colInfoSpoolFiles) = False Then
50030    Set isf = New clsInfoSpoolFile
50040    isf.ReadInfoFile filename
50050    colInfoSpoolFiles.Add isf, filename
50060    Set GetInfoSpoolFileObject = isf
50070   Else
50080    Set GetInfoSpoolFileObject = colInfoSpoolFiles(filename)
50090  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modInfoSpoolFile", "GetInfoSpoolFileObject")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub RemoveInfoSpoolFileObject(filename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  colInfoSpoolFiles.Remove filename
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modInfoSpoolFile", "RemoveInfoSpoolFileObject")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

