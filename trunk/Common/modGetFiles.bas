Attribute VB_Name = "modGetFiles"
Option Explicit

Public Function FindFiles(ByVal Path As String, ByRef files As Collection, _
 Optional ByVal Pattern As String = "*.*", Optional ByVal Attributes As VbFileAttribute = vbNormal, _
 Optional ByVal Recursive As Boolean = True, Optional OnlyNotInUse As Boolean = False) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020  Const vbErr_PathNotFound = 76, INVALID_VALUE = -1
50030  Dim FileAttr As Long, filename As String, hFind As Long, WFD As WIN32_FIND_DATA
50040
50050  Path = CompletePath(Path)
50060
50070  If files Is Nothing Then
50080   Set files = New Collection
50090  End If
50100  Pattern = LCase$(Pattern)
50110
50120  hFind = FindFirstFileA(Path & "*", WFD)
50130  If hFind = INVALID_VALUE Then
50140   Exit Function
50150 '  Err.Raise vbErr_PathNotFound
50160  End If
50170
50180  Do
50190   filename = LeftB$(WFD.cFileName, InStrB(WFD.cFileName, vbNullChar))
50200   FileAttr = GetFileAttributesA(Path & filename)
50210   If FileAttr And vbDirectory Then
50220     If Recursive Then
50230      If FileAttr <> INVALID_VALUE And filename <> "." And filename <> ".." Then
50240       FindFiles = FindFiles + FindFiles(Path & filename, files, Pattern, Attributes)
50250      End If
50260     End If
50270    Else
50280     If (FileAttr And Attributes) = Attributes Then
50290      If LCase$(filename) Like Pattern Then
50300       FindFiles = FindFiles + 1
50310       If OnlyNotInUse = False Then
50320         AddFile files, Path, filename
50330        Else
50340         If FileInUse(Path & filename) = False Then
50350          AddFile files, Path, filename
50360         End If
50370       End If
50380      End If
50390     End If
50400   End If
50410  Loop While FindNextFileA(hFind, WFD)
50420  FindClose hFind
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGetFiles", "FindFiles")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub AddFile(ByRef files As Collection, Path As String, filename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim spoolFile As clsSpoolFile
50020  Set spoolFile = New clsSpoolFile
50030  spoolFile.Path = Path
50040  spoolFile.FullFileName = Path & filename
50050  spoolFile.FileLen = GetFileLength(Path & filename)
50060  spoolFile.FileDateTime = FileDateTime(Path & filename)
50070  spoolFile.FileDateTimeJobIdKey = GetFileDateTimeString(FileDateTime(Path & filename))
50080  files.Add spoolFile, spoolFile.FullFileName
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGetFiles", "AddFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function GetFileDateTimeString(value As Date) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  GetFileDateTimeString = Format$(value, "yyyymmddhhMMss")
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGetFiles", "GetFileDateTimeString")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
