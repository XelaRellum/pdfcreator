Attribute VB_Name = "modGeneral"
Option Explicit

Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Const MAX_PATH = 260

Private Const SE_ERR_NOASSOC = 31
Private Const SE_ERR_NOTFOUND = 2

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetTempPathA Lib "kernel32" _
 (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" _
 Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectoryA Lib "kernel32" _
 (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Function GetTempFileNameA Lib "kernel32" _
 (ByVal lpszPath As String, ByVal lpPrefixString As String, _
  ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

Private Declare Function ShellExecute Lib "shell32.dll" _
  Alias "ShellExecuteA" (ByVal hWnd As Long, _
  ByVal lpOperation As String, ByVal lpFile As String, _
  ByVal lpParameters As String, _
  ByVal lpDirectory As String, _
  ByVal nShowCmd As Long) As Long


Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
      
Private Const MF_BYPOSITION = &H400
Private Const MF_REMOVE = &H1000

Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (pDicDesc As IconType, riid As CLSIdType, ByVal fown As Long, lpUnk As Object) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As ShellFileInfoType, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

Private Type IconType
 cbSize As Long
 picType As PictureTypeConstants
 hIcon As Long
End Type

Private Type CLSIdType
 ID(16) As Byte
End Type

Private Type ShellFileInfoType
 hIcon As Long
 iIcon As Long
 dwAttributes As Long
 szDisplayName As String * 260
 szTypeName As String * 80
End Type

Public Enum IconSize
 Large = &H100&
 Small = &H101&
End Enum

Private Const SHGFI_USEFILEATTRIBUTES = &H10

Private Declare Function FindFirstFile Lib "kernel32" _
        Alias "FindFirstFileA" (ByVal lpFileName As String, _
        lpFindFileData As WIN32_FIND_DATA) As Long
        
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

Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4

Public Function LoadIcon(ByVal lngSize As IconSize, ByVal strPath As String, Icon As IPictureDisp) As Long
 Dim iuUnkown As IUnknown, itIcon As IconType, CLSID As CLSIdType, _
  sfiShellInfo As ShellFileInfoType
 Call SHGetFileInfo(strPath, 0&, sfiShellInfo, Len(sfiShellInfo), SHGFI_USEFILEATTRIBUTES Or lngSize)
 With itIcon
  .cbSize = Len(itIcon)
  .picType = vbPicTypeIcon
  .hIcon = sfiShellInfo.hIcon
 End With
 CLSID.ID(8) = &HC0
 CLSID.ID(15) = &H46
 Call OleCreatePictureIndirect(itIcon, CLSID, 1&, iuUnkown)
 Set Icon = iuUnkown
 If Icon Is Nothing Then
   LoadIcon = 1
  Else
   LoadIcon = 0
 End If
End Function

Public Sub SplitPath(FullPath As String, Optional Drive As String, Optional Path As String, Optional Filename As String, Optional File As String, Optional Extension As String)
 Dim nPos As Integer
 nPos = InStrRev(FullPath, "\")
 If nPos > 0 Then
   If Left$(FullPath, 2) = "\\" Then
    If nPos = 2 Then
     Drive = FullPath: Path = "": Filename = "": File = ""
     Extension = ""
     Exit Sub
    End If
   End If
   Path = Left$(FullPath, nPos - 1)
   Filename = Mid$(FullPath, nPos + 1)
   nPos = InStrRev(Filename, ".")
   If nPos > 0 Then
     File = Left$(Filename, nPos - 1)
     Extension = Mid$(Filename, nPos + 1)
    Else
     File = Filename
     Extension = ""
   End If
  Else
   Path = FullPath: Filename = ""
   File = "": Extension = ""
 End If
 If Left$(Path, 2) = "\\" Then
   nPos = InStr(3, Path, "\")
   If nPos Then
     Drive = Left$(Path, nPos - 1)
    Else
     Drive = Path
   End If
  Else
   If Len(Path) = 2 Then
    If Right$(Path, 1) = ":" Then
     Path = Path & "\"
    End If
   End If
   If Mid$(Path, 2, 2) = ":\" Then
    Drive = Left$(Path, 2)
   End If
 End If
End Sub

Public Function GetWindowsDirectory() As String
 On Local Error Resume Next
 Dim res As Long, Windir As String
 Windir = Space$(MAX_PATH)
 res = GetWindowsDirectoryA(MAX_PATH, Windir)
 If res > 0 Then
   GetWindowsDirectory = Left$(Windir, res)
  Else
   GetWindowsDirectory = "C:\Windows"
 End If
End Function

Public Function GetTempPath() As String
 On Local Error Resume Next
 Dim res As Long, Tempdir As String, r As clsRegistry
 
 Set r = New clsRegistry
 r.hkey = HKEY_CURRENT_USER
 r.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
 Tempdir = Trim$(r.GetRegistryValue("Local Settings"))
 If Right$(Tempdir, 1) = "\" Then
   GetTempPath = Tempdir & "Temp\"
  Else
   GetTempPath = Tempdir & "\Temp\"
 End If
 Set r = Nothing
 If Tempdir = "" Then
  Tempdir = Space$(MAX_PATH)
  res = GetTempPathA(MAX_PATH, Tempdir)
  If res > 0 Then
    GetTempPath = Left$(Tempdir, res)
   Else
    GetTempPath = "C:\"
  End If
 End If
End Function

Public Function GetTempFile(Optional ByVal Path As String, Optional Prefix As String) As String
 On Local Error Resume Next
 Dim res As Long, Tempfile As String, tPath As String
 Tempfile = Space$(MAX_PATH)
 tPath = Trim$(Path)
 If Dir(tPath, vbDirectory) = "" Then
  tPath = GetTempPath
 End If
 
 res = GetTempFileNameA(tPath, Prefix, 0, Tempfile)
 If res <> 0 Then
   GetTempFile = Left$(Tempfile, InStr(Tempfile, Chr$(0)) - 1)
  Else
   GetTempFile = "~.tmp"
 End If
End Function

Public Sub OpenDocument(sFilename As String)
 Dim sDirectory As String, res As Long, handle As Long
  
 handle = GetDesktopWindow()
 res = ShellExecute(handle, "open", sFilename, _
    vbNullString, vbNullString, vbNormalFocus)

 If res = SE_ERR_NOTFOUND Then
  ElseIf res = SE_ERR_NOASSOC Then
    sDirectory = Space$(260)
    res = GetSystemDirectory(sDirectory, Len(sDirectory))
    sDirectory = Left$(sDirectory, res)
    Call ShellExecute(handle, vbNullString, _
      "RUNDLL32.EXE", "shell32.dll,OpenAs_RunDLL " & _
      sFilename, sDirectory, vbNormalFocus)
  End If
End Sub

Public Function GetFiles(ByVal Path As String, Optional Searchmask As String = "*.*") As Collection
 Dim tColl As Collection, tFilename As String
 Set tColl = New Collection
 Path = Trim$(Path)
 If Right$(Path, 1) <> "\" Then
  Path = Path & "\"
 End If
 tFilename = Dir(Path & Searchmask)
 Do While tFilename <> ""
  tColl.Add Path & "|" & Path & tFilename & "|" & FileLen(Path & tFilename) & "|" & FileDateTime(Path & tFilename)
  tFilename = Dir()
  DoEvents
 Loop
 Set GetFiles = tColl
 Set tColl = Nothing
End Function


Public Sub CombineFiles(ByVal Filename As String, Files As Collection, stb As StatusBar)
 On Local Error GoTo ErrorHandler
 Dim i As Long, fnSource As Long, fnDest As Long, sBuffer As String, _
  aLen As Double, tLen As Double
 
 Filename = Trim$(Filename)
 If Filename = "" Or Files.Count = 0 Or Right$(Filename, 1) = "\" Then
  Exit Sub
 End If
 If Dir(Filename) <> "" Then
  Exit Sub
 End If
 If Files.Count = 1 Then
  Exit Sub
 End If
 fnDest = FreeFile
 aLen = 0
 For i = 1 To Files.Count
  aLen = aLen + FileLen(Files.item(i))
 Next i
 Open Filename For Binary As #fnDest
 For i = 1 To Files.Count
  DoEvents
  If FileLen(Files.item(i)) > 0 Then
   fnSource = FreeFile
   Open Files.item(i) For Binary Access Read As #fnSource
   sBuffer = String(LOF(fnSource), Chr$(0))
   Get #fnSource, , sBuffer
   Put #fnDest, , sBuffer
   Close #fnSource
  End If
  tLen = tLen + FileLen(Files.item(i))
  stb.Panels("Percent").Text = Format$(tLen / aLen, "0.0%")
  Kill Files.item(i)
  DoEvents
 Next i
 Close #fnDest
 stb.Panels("Percent").Text = ""
 Exit Sub
ErrorHandler:
' MsgBox Err.Description
End Sub

Public Sub SetFont(Frm As Form, ByVal Fontname As String, ByVal Charset As Long, ByVal Fontsize As Long)
 Dim ctl As Control
 
 For Each ctl In Frm.Controls
  If TypeOf ctl Is Label Or _
     TypeOf ctl Is Form Or _
     TypeOf ctl Is ComboBox Or _
     TypeOf ctl Is CheckBox Or _
     TypeOf ctl Is CommandButton Or _
     TypeOf ctl Is ListView Or _
     TypeOf ctl Is StatusBar Or _
     TypeOf ctl Is TextBox Or _
     TypeOf ctl Is Frame Then
   With ctl
    .Font = Fontname
    If Not (TypeOf ctl Is StatusBar) And Not (TypeOf ctl Is ListView) Then
     .Fontsize = Fontsize
    End If
    .Font.Charset = Charset
   End With
  End If
 Next ctl
End Sub

Public Function ReplaceForbiddenChars(chkStr As String, Optional ReplaceChar As String = "_") As String
 Dim tStr As String, i As Long, Forbiddenchars As String
 Forbiddenchars = "\/:*?<>|" & Chr$(34)
 tStr = chkStr
 For i = 1 To Len(Forbiddenchars)
  tStr = Replace$(tStr, Mid$(Forbiddenchars, i, 1), ReplaceChar)
 Next i
 ReplaceForbiddenChars = tStr
End Function

Public Function IsForbiddenChars(chkStr As String) As Boolean
 Dim i As Long, Forbiddenchars As String
 IsForbiddenChars = False
 Forbiddenchars = "\/:*?<>|" & Chr$(34)
 For i = 1 To Len(Forbiddenchars)
  If InStr(chkStr, Mid$(Forbiddenchars, i, 1)) > 0 Then
   IsForbiddenChars = True
   Exit Function
  End If
 Next i
End Function

Public Sub RemoveX(Frm As Form)
 Dim hMenu As Long, nPosition As Long

 hMenu = GetSystemMenu(Frm.hWnd, 0)
 If hMenu <> 0 Then
  nPosition = GetMenuItemCount(hMenu)
  Call RemoveMenu(hMenu, nPosition - 1, MF_REMOVE Or MF_BYPOSITION)
  Call RemoveMenu(hMenu, nPosition - 2, MF_REMOVE Or MF_BYPOSITION)
  Call DrawMenuBar(Frm.hWnd)
 End If
End Sub

Public Function GetUsername() As String
 Const MAX_USERNAME_LENGTH = 128
 Dim tStr As String
 tStr = String(MAX_USERNAME_LENGTH, Chr$(0))
 GetUserNameA tStr, MAX_USERNAME_LENGTH
 GetUsername = Left$(tStr, InStr(tStr, Chr$(0)) - 1)
End Function

Public Function GetComputerName() As String
 Const MAX_COMPUTERNAME_LENGTH As Long = 31
 Dim tStr As String
 tStr = String(MAX_COMPUTERNAME_LENGTH + 1, Chr$(0))
 GetComputerNameA tStr, MAX_COMPUTERNAME_LENGTH + 1
 GetComputerName = Left$(tStr, MAX_COMPUTERNAME_LENGTH + 1)
End Function

Public Function IsFormLoaded(Form As Form) As Boolean
 Dim nForm As Form
 For Each nForm In Forms
  If nForm Is Form Then
   IsFormLoaded = True
   Exit For
  End If
 Next
End Function

Public Function GetDesktop() As String
 On Local Error Resume Next
 Dim TempDesktop As String, r As clsRegistry
 
 Set r = New clsRegistry
 r.hkey = HKEY_CURRENT_USER
 r.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
 TempDesktop = Trim$(r.GetRegistryValue("Desktop"))
 If Right$(TempDesktop, 1) <> "\" Then
  GetDesktop = TempDesktop & "\"
 End If
 Set r = Nothing
 If TempDesktop = "" Then
  GetDesktop = GetSpecialFolder(ssfDESKTOPDIRECTORY)
  If Trim$(GetDesktop) = "" Then
   GetDesktop = "C:\"
  End If
 End If
End Function

Public Function GetMyFiles() As String
 On Local Error Resume Next
 Dim TempMyfiles As String, r As clsRegistry
 
 Set r = New clsRegistry
 r.hkey = HKEY_CURRENT_USER
 r.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
 TempMyfiles = Trim$(r.GetRegistryValue("Personal"))
 If Right$(TempMyfiles, 1) <> "\" Then
  GetMyFiles = TempMyfiles & "\"
 End If
 Set r = Nothing
 If TempMyfiles = "" Then
  GetMyFiles = GetSpecialFolder(ssfPERSONAL)
  If Trim$(GetMyFiles) = "" Then
   GetMyFiles = "C:\"
  End If
 End If
End Function

Public Function GetFileAttributesStr(Filename As String) As String
 Dim hFind As Long, wFD As WIN32_FIND_DATA, Attr As Long, AA As String
 hFind = FindFirstFile(Filename, wFD)
 Attr = wFD.dwFileAttributes
 If Attr And FILE_ATTRIBUTE_ARCHIVE Then AA = AA & "A"
 If Attr And FILE_ATTRIBUTE_COMPRESSED Then AA = AA & "C"
 If Attr And FILE_ATTRIBUTE_DIRECTORY Then AA = AA & "D"
 If Attr And FILE_ATTRIBUTE_HIDDEN Then AA = AA & "H"
 If Attr And FILE_ATTRIBUTE_NORMAL Then AA = AA & "N"
 If Attr And FILE_ATTRIBUTE_READONLY Then AA = AA & "R"
 If Attr And FILE_ATTRIBUTE_SYSTEM Then AA = AA & "S"
 GetFileAttributesStr = AA
End Function
