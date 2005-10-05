Attribute VB_Name = "modGetFiles"
Option Explicit

Public Function FindFiles(ByVal Path As String, ByRef Files As Collection, _
 Optional ByVal Pattern As String = "*.*", Optional ByVal Attributes As VbFileAttribute = vbNormal, _
 Optional ByVal Recursive As Boolean = True, Optional OnlyNotInUse As Boolean = False) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020  Const vbErr_PathNotFound = 76, INVALID_VALUE = -1
50030  Dim FileAttr As Long, Filename As String, hFind As Long, WFD As WIN32_FIND_DATA
50040
50050  Path = CompletePath(Path)
50060
50070  If Files Is Nothing Then
50080   Set Files = New Collection
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
50190   Filename = LeftB$(WFD.cFileName, InStrB(WFD.cFileName, vbNullChar))
50200   FileAttr = GetFileAttributesA(Path & Filename)
50210   If FileAttr And vbDirectory Then
50220     If Recursive Then
50230      If FileAttr <> INVALID_VALUE And Filename <> "." And _
      Filename <> ".." Then
50250        FindFiles = FindFiles + FindFiles(Path & Filename, Files, Pattern, Attributes)
50260      End If
50270     End If
50280    Else
50290     If (FileAttr And Attributes) = Attributes Then
50300      If LCase$(Filename) Like Pattern Then
50310       FindFiles = FindFiles + 1
50320       If OnlyNotInUse = False Then
50330         Files.Add Path & "|" & Path & Filename & "|" & FileLen(Path & Filename) & "|" & FileDateTime(Path & Filename)
50340        Else
50350         If FileInUse(Path & Filename) = False Then
50360          Files.Add Path & "|" & Path & Filename & "|" & FileLen(Path & Filename) & "|" & FileDateTime(Path & Filename)
50370         End If
50380       End If
50390      End If
50400     End If
50410   End If
50420  Loop While FindNextFileA(hFind, WFD)
50430  FindClose hFind
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

