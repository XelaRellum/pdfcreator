VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrivesDirsFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "List drives, directories and files."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function GetLogicalDriveStrings Lib "Kernel32" Alias "GetLogicalDriveStringsA" _
 (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType Lib "Kernel32" Alias "GetDriveTypeA" _
 (ByVal nDrive As String) As Long

Private Declare Function FindFirstFile Lib "Kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "Kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "Kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "Kernel32" (ByVal hFindFile As Long) As Long

Const MAX_PATH = 260
Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private iCurrentDrive As String
Private iCurrentDirectory As String
Private iFilePattern As String

Public Property Let FilePattern(ByVal Pattern As String)
 iFilePattern = Pattern
End Property

Public Property Get FilePattern() As String
 FilePattern = iFilePattern
End Property

Public Property Let CurrentDirectory(ByVal Directory As String)
 iCurrentDirectory = Trim$(Directory)
 If Right$(iCurrentDirectory, 1) <> "\" Then
  iCurrentDirectory = iCurrentDirectory & "\"
 End If
End Property

Public Property Get CurrentDirectory() As String
 CurrentDirectory = iCurrentDirectory
End Property

Public Function GetFiles(Optional Directory As String, Optional IncludeHidden As Boolean = False) As Collection
 Dim tDir As String, attr As Long, tName As String, tCol As Collection, _
  attrD As Long, hFile As Long, WFD As WIN32_FIND_DATA, cont As Boolean

 On Error GoTo ErrorHandler
 Set tCol = New Collection
 Set GetFiles = tCol
 tDir = Trim$(Directory)
 If Len(tDir) = 0 Then
  tDir = iCurrentDirectory
 End If
 attr = vbNormal
 If IncludeHidden = True Then
  attr = attr + vbHidden
 End If
 cont = True
 hFile = FindFirstFile(tDir & FilePattern, WFD)
 If hFile <> INVALID_HANDLE_VALUE Then
  Do While cont = True
   tName = StripNulls(WFD.cFileName)
   If (tName <> ".") And (tName <> "..") Then
    If (GetFileAttributes(tDir & tName) And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY And (GetFileAttributes(tDir & tName) <> -1) Then
     If GetFileAttributes(tDir & tName) And FILE_ATTRIBUTE_ARCHIVE Then
      attrD = attrD + vbArchive
     End If
     If GetFileAttributes(tDir & tName) And FILE_ATTRIBUTE_READONLY Then
      attrD = attrD + vbReadOnly
     End If
     If GetFileAttributes(tDir & tName) And FILE_ATTRIBUTE_SYSTEM Then
      attrD = attrD + vbSystem
     End If
     If GetFileAttributes(tDir & tName) And FILE_ATTRIBUTE_HIDDEN And ((attr And vbHidden) = vbHidden) Then
      attrD = attrD + vbHidden
      tCol.Add tName & Chr$(0) & attrD
     End If
     If (GetFileAttributes(tDir & tName) And FILE_ATTRIBUTE_HIDDEN) <> FILE_ATTRIBUTE_HIDDEN Then
      tCol.Add tName & Chr$(0) & attrD
     End If
    End If
   End If
   DoEvents
   cont = FindNextFile(hFile, WFD)
  Loop
  cont = FindClose(hFile)
 End If
 Set GetFiles = tCol
 Exit Function
ErrorHandler:
 If Err.number <> 52 Then
  Debug.Print "GetFiles: " & tName & " - " & tDir & FilePattern
  MsgBox "GetFiles: " & Err.number & " " & Err.Description
 End If
End Function

Public Function GetDirectories(Optional Directory As String, Optional IncludeHidden As Boolean = False, Optional OnlySubDirExists As Boolean = False) As Collection
 On Error GoTo ErrorHandler
 Dim tDir As String, attr As Long, tName As String, tCol As Collection, _
  attrD As Long, hFile As Long, WFD As WIN32_FIND_DATA, cont As Boolean

 Set tCol = New Collection
 Set GetDirectories = tCol
 tDir = Trim$(Directory)
 If Len(tDir) = 0 Then
  tDir = iCurrentDirectory
 End If
 attr = vbDirectory
 If IncludeHidden = True Then
  attr = attr + vbHidden
 End If

 cont = True
 hFile = FindFirstFile(tDir & "*", WFD)
 If hFile <> INVALID_HANDLE_VALUE Then
  Do While cont = True
   tName = StripNulls(WFD.cFileName)
   If (tName <> ".") And (tName <> "..") Then
    If (GetFileAttributes(tDir & tName) And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY And (GetFileAttributes(tDir & tName) <> -1) Then
     attrD = vbDirectory
     If GetFileAttributes(tDir & tName) And FILE_ATTRIBUTE_READONLY Then
      attrD = attrD + vbReadOnly
     End If
     If GetFileAttributes(tDir & tName) And FILE_ATTRIBUTE_SYSTEM Then
      attrD = attrD + vbSystem
     End If
     If GetFileAttributes(tDir & tName) And FILE_ATTRIBUTE_HIDDEN And ((attr And vbHidden) = vbHidden) Then
      attrD = attrD + vbHidden
      tCol.Add tName & Chr$(0) & attrD
     End If
     If (GetFileAttributes(tDir & tName) And FILE_ATTRIBUTE_HIDDEN) <> FILE_ATTRIBUTE_HIDDEN Then
      tCol.Add tName & Chr$(0) & attrD
     End If
    End If
   End If
   cont = FindNextFile(hFile, WFD)
   If tCol.Count > 0 And OnlySubDirExists = True Then
    Set GetDirectories = tCol
    cont = FindClose(hFile)
    Exit Function
   End If
   DoEvents
  Loop
  cont = FindClose(hFile)
 End If

 Set GetDirectories = tCol
 Exit Function
ErrorHandler:
 If Err.number <> 52 Then
  MsgBox "GetDirectories: " & Err.number & " " & Err.Description
 End If
End Function

Public Function GetDrives(Optional NoRemovableVolume As Boolean = False) As Collection
 Dim L As Long, Buffer As String, res As Long, drives As String, _
  drv() As String, label As String, dTyp As Long, i As Long
 L = 64: Buffer = Space(L)
 res = GetLogicalDriveStrings(L, Buffer)
 drives = Left$(Buffer, res)
 Set GetDrives = New Collection
 If Len(drives) > 0 Then
  If InStr(drives, Chr$(0)) > 0 Then
    drv = Split(drives, Chr$(0))
    For i = LBound(drv) To UBound(drv)
     If Trim$(drv(i)) <> "" Then
      dTyp = GetDriveType(drv(i))
      If dTyp <> 2 And dTyp <> 5 Then
        label = Dir(drv(i), vbVolume)
       Else
        If NoRemovableVolume = False Then
         label = Dir(drv(i), vbVolume)
        End If
      End If

      GetDrives.Add drv(i) & Chr$(0) & label & Chr$(0) & dTyp
     End If
     Next i
    Else
     GetDrives.Add drives
   End If
  End If
End Function

Private Sub Class_Initialize()
 CurrentDirectory = "C:\"
 FilePattern = "*.*"
End Sub

Private Function StripNulls(OriginalStr As String) As String
 If (InStr(OriginalStr, Chr(0)) > 0) Then
  OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
 End If
 StripNulls = OriginalStr
End Function

