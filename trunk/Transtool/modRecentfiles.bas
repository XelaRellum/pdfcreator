Attribute VB_Name = "modRecentfiles"
Option Explicit

Public Enum eRecentfileslistLocation
 Registry = 0
 ApplicationDatapath = 1
End Enum

Public Const MaxRecentfiles = 10

Private Const RecentFileRegRootkey = "Software\PDFCreator\TransTool\RecentFiles"
Private Const UserProjectRootkey = "Software\PDFCreator"
Private Const UserProjectTransToolRootkey = UserProjectRootkey & "\TransTool"

Public mRecentfileslistLocation As eRecentfileslistLocation

Public Sub SaveRecentfiles(RecentFiles As Collection)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If mRecentfileslistLocation = ApplicationDatapath Then
50020   SaveRecentfilesToUserAppdata RecentFiles
50030  End If
50040  If mRecentfileslistLocation = Registry Then
50050   SaveRecentfilesToRegistry RecentFiles
50060  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modRecentfiles", "SaveRecentfiles")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SaveRecentfilesToRegistry(RecentFiles As Collection)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry, i As Long
50020  Set reg = New clsRegistry
50030  reg.hkey = HKEY_CURRENT_USER
50040  reg.KeyRoot = RecentFileRegRootkey
50050  If reg.KeyExists = True Then
50060   reg.KeyRoot = UserProjectTransToolRootkey
50070   reg.DeleteKey "RecentFiles"
50080  End If
50090  reg.KeyRoot = RecentFileRegRootkey
50100  reg.CreateKey
50110  If RecentFiles.Count > 0 Then
50120   For i = 1 To RecentFilesCount
50130    If i <= RecentFiles.Count Then
50140     If Len(Trim$(RecentFiles(i))) > 0 Then
50150      reg.SetRegistryValue i, RecentFiles(i), REG_SZ
50160     End If
50170    End If
50180   Next i
50190  End If
50200  Set reg = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modRecentfiles", "SaveRecentfilesToRegistry")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SaveRecentfilesToUserAppdata(RecentFiles As Collection)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ini As clsINI, i As Long
50020  Set ini = New clsINI
50030  ini.Filename = CompletePath(GetMyAppData) & "PDFCreator\transtool.ini"
50040  If ini.CheckIniFile = False Then
50050   MakePath CompletePath(GetMyAppData) & "PDFCreator"
50060   ini.CreateIniFile
50070  End If
50080  ini.Section = "Recentfiles"
50090  ini.DeleteSectionFromInifile
50100  For i = 1 To RecentFiles.Count
50110   ini.SaveKey RecentFiles(i), CStr(i)
50120  Next i
50130  ini.FlushInifile
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modRecentfiles", "SaveRecentfilesToUserAppdata")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function GetRecentFiles() As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If mRecentfileslistLocation = ApplicationDatapath Then
50020   Set GetRecentFiles = GetRecentFilesFromUserAppdata
50030  End If
50040  If mRecentfileslistLocation = Registry Then
50050   Set GetRecentFiles = GetRecentFilesFromRegistry
50060  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modRecentfiles", "GetRecentFiles")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function GetRecentFilesFromRegistry() As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry, col As Collection, rfArr() As String, i As Long, _
  Index As Long
50030  ReDim rfArr(MaxRecentfiles)
50040  Set GetRecentFilesFromRegistry = New Collection
50050  Set reg = New clsRegistry
50060  reg.hkey = HKEY_CURRENT_USER
50070  reg.KeyRoot = RecentFileRegRootkey
50080  If reg.KeyExists = True Then
50090   Set col = reg.EnumRegistryValues(HKEY_CURRENT_USER, RecentFileRegRootkey)
50100   For i = 1 To col.Count
50110    If IsNumeric(col(i)(0)) = True Then
50120     Index = CLng(col(i)(0))
50130     If Index >= 1 And Index <= MaxRecentfiles Then
50140      rfArr(Index) = Trim$(col(i)(1))
50150     End If
50160    End If
50170   Next i
50180  End If
50190  Set reg = Nothing
50200  For i = 1 To MaxRecentfiles
50210   If Len(rfArr(i)) > 0 Then
50220    GetRecentFilesFromRegistry.Add rfArr(i)
50230   End If
50240  Next i
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modRecentfiles", "GetRecentFilesFromRegistry")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function GetRecentFilesFromUserAppdata() As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ini As clsINI, i As Long, keys As Collection, Value As String, _
  c As Long, tStr As String
50030  Set GetRecentFilesFromUserAppdata = New Collection
50040  Set ini = New clsINI
50050  ini.Filename = CompletePath(GetMyAppData) & "PDFCreator\transtool.ini"
50060  If ini.CheckIniFile = False Then
50070   Exit Function
50080  End If
50090  ini.Section = "Recentfiles"
50100  Set keys = ini.GetAllKeysFromSection(, , , True)
50110  If keys.Count = 0 Then
50120   Exit Function
50130  End If
50140  For i = 1 To MaxRecentfiles
50150   tStr = ini.GetKeyFromSection(CStr(i))
50160   If Len(tStr) > 0 Then
50170    GetRecentFilesFromUserAppdata.Add tStr
50180   End If
50190  Next i
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modRecentfiles", "GetRecentFilesFromUserAppdata")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function AddRecentfile(Filename As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim col As Collection, i As Long
50020  Set col = GetRecentFiles
50030  AddRecentfile = False
50040  If col.Count > 0 Then
50050    If UCase$(col(1)) <> UCase$(Filename) Then
50060     col.Add Filename, , 1
50070     For i = 2 To col.Count
50080      If UCase$(col(i)) = UCase$(col(1)) Then
50090       col.Remove i
50100       Exit For
50110      End If
50120     Next i
50130     SaveRecentfiles col
50140     AddRecentfile = True
50150    End If
50160   Else
50170    col.Add Filename
50180    SaveRecentfiles col
50190    AddRecentfile = True
50200  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modRecentfiles", "AddRecentfile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetRecentFile(RecentFilenumber As Long) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim col As Collection
50020  Set col = GetRecentFiles
50030  If RecentFilenumber <= col.Count Then
50040   GetRecentFile = col(RecentFilenumber)
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modRecentfiles", "GetRecentFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub RemoveRecentFile(RecentFilenumber As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim col As Collection
50020  Set col = GetRecentFiles
50030  If RecentFilenumber <= col.Count Then
50040   col.Remove RecentFilenumber
50050   SaveRecentfiles col
50060  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modRecentfiles", "RemoveRecentFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Property Let RecentfileslistLocation(SaveLocation As eRecentfileslistLocation)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  mRecentfileslistLocation = SaveLocation
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modRecentfiles", "RecentfileslistLocation [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get RecentfileslistLocation() As eRecentfileslistLocation
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  RecentfileslistLocation = mRecentfileslistLocation
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modRecentfiles", "RecentfileslistLocation [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

