Attribute VB_Name = "modGetFiles"
Option Explicit

Public Function FindFiles(ByVal Path As String, ByRef Files As Collection, _
 Optional ByVal Pattern As String = "*.*", Optional ByVal Attributes As VbFileAttribute = vbNormal, _
 Optional ByVal Recursive As Boolean = True) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020  Const vbErr_PathNotFound = 76, INVALID_VALUE = -1
50030  Dim FileAttr As Long, Filename As String, hFind As Long, WFD As WIN32_FIND_DATA
50040
50050  If Right$(Path, 1) <> "\" Then
50060   Path = Path & "\"
50070  End If
50080
50090  If Files Is Nothing Then
50100   Set Files = New Collection
50110  End If
50120  Pattern = LCase$(Pattern)
50130
50140  hFind = FindFirstFileA(Path & "*", WFD)
50150  If hFind = INVALID_VALUE Then
50160   Exit Function
50170 '  Err.Raise vbErr_PathNotFound
50180  End If
50190
50200  Do
50210   Filename = LeftB$(WFD.cFileName, InStrB(WFD.cFileName, vbNullChar))
50220   FileAttr = GetFileAttributesA(Path & Filename)
50230   If FileAttr And vbDirectory Then
50240     If Recursive Then
50250      If FileAttr <> INVALID_VALUE And Filename <> "." And _
      Filename <> ".." Then
50270        FindFiles = FindFiles + FindFiles(Path & Filename, Files, Pattern, Attributes)
50280      End If
50290     End If
50300    Else
50310     If (FileAttr And Attributes) = Attributes Then
50320      If LCase$(Filename) Like Pattern Then
50330       FindFiles = FindFiles + 1
50340       Files.Add Path & "|" & Path & Filename & "|" & FileLen(Path & Filename) & "|" & FileDateTime(Path & Filename)
50350 '     tColl.Add Path & "|" & Path & tFilename & "|" & FileLen(Path & tFilename) & "|" & FileDateTime(Path & tFilename)
50360      End If
50370     End If
50380   End If
50390  Loop While FindNextFileA(hFind, WFD)
50400  FindClose hFind
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

