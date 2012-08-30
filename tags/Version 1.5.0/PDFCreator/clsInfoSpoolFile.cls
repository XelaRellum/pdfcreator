VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInfoSpoolFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public FirstDocumentTitle As String
Public FirstPrinterName As String
Public FirstSpoolFileName As String
Public FirstClientComputer As String
Public FirstJobID As String
Public FirstSessionID As String
Public FirstUserName As String

Public sumFileSizes As Currency
Public InfoFiles As Collection

Private Sub Class_Initialize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Set InfoFiles = New Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsInfoSpoolFile", "Class_Initialize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Class_Terminate()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Set InfoFiles = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsInfoSpoolFile", "Class_Terminate")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function ReadInfoFile(ByVal InfoFileName As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ini As clsINI, sections As Collection, isfi As clsInfoSpoolFileInfo, Section As String, i As Long
50020  If Not FileExists(InfoFileName) Then
50030   ReadInfoFile = False
50040   Exit Function
50050  End If
50060  Set ini = New clsINI
50070  ini.filename = InfoFileName
50080  Set InfoFiles = New Collection
50090
50100  Set sections = ini.GetAllSectionsFromInifile()
50110  For i = 1 To sections.Count
50120   Set isfi = New clsInfoSpoolFileInfo
50130   isfi.SessionID = ini.GetKeyFromSection("SessionID", sections(i))
50140   isfi.WinStation = ini.GetKeyFromSection("WinStation", sections(i))
50150   isfi.UserName = ini.GetKeyFromSection("UserName", sections(i))
50160   isfi.ClientComputer = ini.GetKeyFromSection("ClientComputer", sections(i))
50170   isfi.SpoolFileName = ini.GetKeyFromSection("SpoolFileName", sections(i))
50180   isfi.PrinterName = ini.GetKeyFromSection("PrinterName", sections(i))
50190   isfi.JobID = ini.GetKeyFromSection("JobId", sections(i))
50200   isfi.DocumentTitle = ini.GetKeyFromSection("DocumentTitle", sections(i))
50210   If FileExists(isfi.SpoolFileName) Then
50220    sumFileSizes = sumFileSizes + GetFileLength(isfi.SpoolFileName)
50230    InfoFiles.Add isfi
50240    If InfoFiles.Count = 1 Then
50250     FirstDocumentTitle = isfi.DocumentTitle
50260     FirstPrinterName = isfi.PrinterName
50270     FirstSpoolFileName = isfi.SpoolFileName
50280     FirstClientComputer = isfi.ClientComputer
50290     FirstJobID = isfi.JobID
50300     FirstSessionID = isfi.SessionID
50310     FirstUserName = isfi.UserName
50320    End If
50330   End If
50340  Next i
50350  ReadInfoFile = True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsInfoSpoolFile", "ReadInfoFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
