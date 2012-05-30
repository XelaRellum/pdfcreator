Attribute VB_Name = "modGetFiles"
Option Explicit

Public Function FindFiles(ByVal Path As String, ByRef files As Collection, _
 Optional ByVal Pattern As String = "*.*", Optional ByVal Attributes As VbFileAttribute = vbNormal, _
 Optional ByVal Recursive As Boolean = True, Optional OnlyNotInUse As Boolean = False) As Long

 Const vbErr_PathNotFound = 76, INVALID_VALUE = -1
 Dim FileAttr As Long, fileName As String, hFind As Long, WFD As WIN32_FIND_DATA

 Path = CompletePath(Path)

 If files Is Nothing Then
  Set files = New Collection
 End If
 Pattern = LCase$(Pattern)

 hFind = FindFirstFileA(Path & "*", WFD)
 If hFind = INVALID_VALUE Then
  Exit Function
'  Err.Raise vbErr_PathNotFound
 End If

 Do
  fileName = LeftB$(WFD.cFileName, InStrB(WFD.cFileName, vbNullChar))
  FileAttr = GetFileAttributesA(Path & fileName)
  If FileAttr And vbDirectory Then
    If Recursive Then
     If FileAttr <> INVALID_VALUE And fileName <> "." And fileName <> ".." Then
      FindFiles = FindFiles + FindFiles(Path & fileName, files, Pattern, Attributes)
     End If
    End If
   Else
    If (FileAttr And Attributes) = Attributes Then
     If LCase$(fileName) Like Pattern Then
      FindFiles = FindFiles + 1
      If OnlyNotInUse = False Then
        AddFile files, Path, fileName
       Else
        If FileInUse(Path & fileName) = False Then
         AddFile files, Path, fileName
        End If
      End If
     End If
    End If
  End If
 Loop While FindNextFileA(hFind, WFD)
 FindClose hFind
End Function

Private Sub AddFile(ByRef files As Collection, Path As String, fileName As String)
 Dim spoolFile As clsSpoolFile
 Set spoolFile = New clsSpoolFile
 spoolFile.Path = Path
 spoolFile.FullFileName = Path & fileName
 spoolFile.FileLen = FileLen(Path & fileName)
 spoolFile.FileDateTime = FileDateTime(Path & fileName)
 spoolFile.FileDateTimeKey = GetFileDateTimeString(FileDateTime(Path & fileName))
 files.Add spoolFile, spoolFile.FullFileName
End Sub

Private Function GetFileDateTimeString(value As Date) As String
 GetFileDateTimeString = Format$(value, "yyyymmddhhMMss")
End Function
