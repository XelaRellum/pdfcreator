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

Private Const TOKEN_QUERY = (&H8)
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetUserProfileDirectory Lib "userenv.dll" Alias "GetUserProfileDirectoryA" (ByVal hToken As Long, ByVal lpProfileDir As String, lpcchSize As Long) As Boolean

Public Type PROFILEINFO
 dwSize As Long
 dwFlags As Long
 lpUserName As String
 lpProfilePath As String
 lpDefaultPath As String
 lpServerName As String
 lpPolicyPath As String
 hProfile As Long
End Type

Private Declare Function LoadUserProfile Lib "userenv.dll" (ByVal hToken As Long, ByRef lpProfileInfo As PROFILEINFO) As Long

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'** Begin Decalrations for RunProgramWait
Private Const SW_HIDE = 0
Private Const SW_SHOW = 5
Private Const STARTF_USESHOWWINDOW = &H1

Private Type STARTUPINFO
   cb As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Private Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessId As Long
   dwThreadID As Long
End Type

Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
   lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
   lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
   lpStartupInfo As STARTUPINFO, lpProcessInformation As _
   PROCESS_INFORMATION) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Const NORMAL_PRIORITY_CLASS = &H20&
Const STILL_ACTIVE = &H103
'** End Decalrations for RunProgramWait

Public Const SHCNE_ASSOCCHANGED = &H8000000
Public Const SHCNF_IDLIST = &H0&

Public Declare Sub SHChangeNotify Lib "shell32.dll" _
(ByVal wEventId As Long, ByVal uFlags As Long, dwItem1 As Any, dwItem2 As Any)


' API's for IsValidPath
Private Declare Function InStrNullChar Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function PathIsUNC Lib "shlwapi.dll" Alias "PathIsUNCA" (ByVal pszPath As String) As Long
Private Declare Function PathIsUNCServer Lib "shlwapi.dll" Alias "PathIsUNCServerA" (ByVal pszPath As String) As Long
Private Declare Function PathRemoveFileSpec Lib "shlwapi.dll" Alias "PathRemoveFileSpecA" (ByVal pszPath As String) As Long
Private Declare Function PathStripToRoot Lib "shlwapi.dll" Alias "PathStripToRootA" (ByVal pszPath As String) As Long
Private Declare Sub PathUnquoteSpaces Lib "shlwapi.dll" Alias "PathUnquoteSpacesA" (ByVal lpsz As String)

'API's for FormInTaskbar
'Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Function LoadIcon(ByVal lngSize As IconSize, ByVal strPath As String, Icon As IPictureDisp) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim iuUnkown As IUnknown, itIcon As IconType, CLSID As CLSIdType, _
  sfiShellInfo As ShellFileInfoType
50030  Call SHGetFileInfo(strPath, 0&, sfiShellInfo, Len(sfiShellInfo), SHGFI_USEFILEATTRIBUTES Or lngSize)
50040  With itIcon
50050   .cbSize = Len(itIcon)
50060   .picType = vbPicTypeIcon
50070   .hIcon = sfiShellInfo.hIcon
50080  End With
50090  CLSID.ID(8) = &HC0
50100  CLSID.ID(15) = &H46
50110  Call OleCreatePictureIndirect(itIcon, CLSID, 1&, iuUnkown)
50120  Set Icon = iuUnkown
50130  If Icon Is Nothing Then
50140    LoadIcon = 1
50150   Else
50160    LoadIcon = 0
50170  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "LoadIcon")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub SplitPath(FullPath As String, Optional Drive As String, Optional Path As String, Optional FileName As String, Optional File As String, Optional Extension As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim nPos As Integer
50020  nPos = InStrRev(FullPath, "\")
50030  If nPos > 0 Then
50040    If Left$(FullPath, 2) = "\\" Then
50050     If nPos = 2 Then
50060      Drive = FullPath: Path = vbNullString: FileName = vbNullString: File = vbNullString
50070      Extension = vbNullString
50080      Exit Sub
50090     End If
50100    End If
50110    Path = Left$(FullPath, nPos - 1)
50120    FileName = Mid$(FullPath, nPos + 1)
50130    nPos = InStrRev(FileName, ".")
50140    If nPos > 0 Then
50150      File = Left$(FileName, nPos - 1)
50160      Extension = Mid$(FileName, nPos + 1)
50170     Else
50180      File = FileName
50190      Extension = vbNullString
50200    End If
50210   Else
50220    nPos = InStrRev(FullPath, ":")
50230    If nPos > 0 Then
50240      Path = Mid(FullPath, 1, nPos - 1): FileName = Mid(FullPath, nPos + 1)
50250      nPos = InStrRev(FileName, ".")
50260      If nPos > 0 Then
50270        File = Left$(FileName, nPos - 1)
50280        Extension = Mid$(FileName, nPos + 1)
50290       Else
50300        File = FileName
50310        Extension = vbNullString
50320      End If
50330     Else
50340      Path = vbNullString: FileName = FullPath
50350      nPos = InStrRev(FileName, ".")
50360      If nPos > 0 Then
50370        File = Left$(FileName, nPos - 1)
50380        Extension = Mid$(FileName, nPos + 1)
50390       Else
50400        File = FileName
50410        Extension = vbNullString
50420      End If
50430    End If
50440  End If
50450  If Left$(Path, 2) = "\\" Then
50460    nPos = InStr(3, Path, "\")
50470    If nPos Then
50480      Drive = Left$(Path, nPos - 1)
50490     Else
50500      Drive = Path
50510    End If
50520   Else
50530    If Len(Path) = 2 Then
50540     If Right$(Path, 1) = ":" Then
50550      Path = Path & "\"
50560     End If
50570    End If
50580    If Mid$(Path, 2, 2) = ":\" Then
50590     Drive = Left$(Path, 2)
50600    End If
50610  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "SplitPath")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function GetWindowsDirectory() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim res As Long, Windir As String
50020  Windir = Space$(MAX_PATH)
50030  res = GetWindowsDirectoryA(MAX_PATH, Windir)
50040  If res > 0 Then
50050    GetWindowsDirectory = Left$(Windir, res)
50060   Else
50070    GetWindowsDirectory = "C:\Windows"
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetWindowsDirectory")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetTempPath() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim res As Long, Tempdir As String, r As clsRegistry
50020
50030  Set r = New clsRegistry
50040  r.hkey = HKEY_CURRENT_USER
50050  r.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
50060  Tempdir = Trim$(r.GetRegistryValue("Local Settings"))
50070  If Tempdir <> "" Then
50080    If Right$(Tempdir, 1) = "\" Then
50090      Tempdir = Tempdir & "Temp\"
50100     Else
50110      Tempdir = Tempdir & "\Temp\"
50120    End If
50130    If Len(Dir(Tempdir, vbDirectory)) = 0 Then
50140     MakePath Tempdir
50150    End If
50160    If Dir(Tempdir, vbDirectory) = "" Then
50170     'Can't create tempdir
50180     Tempdir = Space$(MAX_PATH)
50190     res = GetTempPathA(MAX_PATH, Tempdir)
50200     If res > 0 Then
50210       GetTempPath = Left$(Tempdir, res)
50220      Else
50230       GetTempPath = "C:\"
50240     End If
50250    End If
50260    GetTempPath = Tempdir
50270   Else
50280    Tempdir = Space$(MAX_PATH)
50290    res = GetTempPathA(MAX_PATH, Tempdir)
50300    If res > 0 Then
50310      GetTempPath = Left$(Tempdir, res)
50320     Else
50330      GetTempPath = "C:\"
50340    End If
50350  End If
50360  Set r = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetTempPath")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetTempFile(Optional ByVal Path As String, Optional Prefix As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim res As Long, Tempfile As String, tPath As String
50020  Tempfile = Space$(MAX_PATH)
50030  tPath = Trim$(Path)
50040  If Dir(tPath, vbDirectory) = "" Then
50050   tPath = GetTempPath
50060  End If
50070
50080  res = GetTempFileNameA(tPath, Prefix, 0, Tempfile)
50090  If res <> 0 Then
50100    GetTempFile = Left$(Tempfile, InStr(Tempfile, Chr$(0)) - 1)
50110   Else
50120    GetTempFile = "~.tmp"
50130  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetTempFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub OpenDocument(sFilename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim sDirectory As String, res As Long, handle As Long
50020
50030  handle = GetDesktopWindow()
50040  res = ShellExecute(handle, "open", sFilename, _
    vbNullString, vbNullString, vbNormalFocus)
50060
50070  If res = SE_ERR_NOTFOUND Then
50080   ElseIf res = SE_ERR_NOASSOC Then
50090     sDirectory = Space$(260)
50100     res = GetSystemDirectory(sDirectory, Len(sDirectory))
50110     sDirectory = Left$(sDirectory, res)
50120     Call ShellExecute(handle, vbNullString, _
      "RUNDLL32.EXE", "shell32.dll,OpenAs_RunDLL " & _
      sFilename, sDirectory, vbNormalFocus)
50150   End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "OpenDocument")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function GetFiles(ByVal Path As String, Optional Searchmask As String = "*.*") As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tColl As Collection, tFilename As String
50020  Set tColl = New Collection
50030  Path = Trim$(Path)
50040  If Right$(Path, 1) <> "\" Then
50050   Path = Path & "\"
50060  End If
50070  tFilename = Dir(Path & Searchmask)
50080  Do While tFilename <> ""
50090   tColl.Add Path & "|" & Path & tFilename & "|" & FileLen(Path & tFilename) & "|" & FileDateTime(Path & tFilename)
50100   tFilename = Dir()
50110   DoEvents
50120  Loop
50130  Set GetFiles = tColl
50140  Set tColl = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetFiles")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub CombineFiles(ByVal FileName As String, Files As Collection, stb As StatusBar)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, fnSource As Long, fnDest As Long, sBuffer As String, _
  aLen As Double, tLen As Double
50030
50040  FileName = Trim$(FileName)
50050  If FileName = vbNullString Or Files.Count = 0 Or Right$(FileName, 1) = "\" Then
50060   Exit Sub
50070  End If
50080  If Len(Dir(FileName)) > 0 Then
50090   Exit Sub
50100  End If
50110  If Files.Count = 1 Then
50120   Exit Sub
50130  End If
50140  fnDest = FreeFile
50150  aLen = 0
50160  For i = 1 To Files.Count
50170   aLen = aLen + FileLen(Files.item(i))
50180  Next i
50190  Open FileName For Binary As #fnDest
50200  For i = 1 To Files.Count
50210   DoEvents
50220   If FileLen(Files.item(i)) > 0 Then
50230    fnSource = FreeFile
50240    Open Files.item(i) For Binary Access Read As #fnSource
50250    sBuffer = String(LOF(fnSource), Chr$(0))
50260    Get #fnSource, , sBuffer
50270    Put #fnDest, , sBuffer
50280    Close #fnSource
50290   End If
50300   tLen = tLen + FileLen(Files.item(i))
50310   stb.Panels("Percent").Text = Format$(tLen / aLen, "0.0%")
50320   Kill Files.item(i)
50330   DoEvents
50340  Next i
50350  Close #fnDest
50360  stb.Panels("Percent").Text = vbNullString
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "CombineFiles")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SetFont(Frm As Form, ByVal Fontname As String, ByVal Charset As Long, ByVal Fontsize As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ctl As Control
50020
50030  For Each ctl In Frm.Controls
50040   If TypeOf ctl Is Label Or _
     TypeOf ctl Is Form Or _
     TypeOf ctl Is ComboBox Or _
     TypeOf ctl Is CheckBox Or _
     TypeOf ctl Is CommandButton Or _
     TypeOf ctl Is ListView Or _
     TypeOf ctl Is StatusBar Or _
     TypeOf ctl Is TextBox Or _
     TypeOf ctl Is Frame Then
50130    With ctl
50140     .Font = Fontname
50150     If Not (TypeOf ctl Is StatusBar) And Not (TypeOf ctl Is ListView) Then
50160      .Fontsize = Fontsize
50170     End If
50180     .Font.Charset = Charset
50190    End With
50200   End If
50210  Next ctl
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "SetFont")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function ReplaceForbiddenChars(chkStr As String, Optional ReplaceChar As String = "_") As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tStr As String, i As Long, Forbiddenchars As String
50020  Forbiddenchars = "\/:*?<>|"""
50030  tStr = chkStr
50040  For i = 1 To Len(Forbiddenchars)
50050   tStr = Replace$(tStr, Mid$(Forbiddenchars, i, 1), ReplaceChar)
50060  Next i
50070  ReplaceForbiddenChars = tStr
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "ReplaceForbiddenChars")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsForbiddenChars(chkStr As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, Forbiddenchars As String
50020  IsForbiddenChars = False
50030  Forbiddenchars = "\/:*?<>|"""
50040  For i = 1 To Len(Forbiddenchars)
50050   If InStr(chkStr, Mid$(Forbiddenchars, i, 1)) > 0 Then
50060    IsForbiddenChars = True
50070    Exit Function
50080   End If
50090  Next i
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "IsForbiddenChars")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub RemoveX(Frm As Form)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim hMenu As Long, nPosition As Long
50020
50030  hMenu = GetSystemMenu(Frm.hWnd, 0)
50040  If hMenu <> 0 Then
50050   nPosition = GetMenuItemCount(hMenu)
50060   Call RemoveMenu(hMenu, nPosition - 1, MF_REMOVE Or MF_BYPOSITION)
50070   Call RemoveMenu(hMenu, nPosition - 2, MF_REMOVE Or MF_BYPOSITION)
50080   Call DrawMenuBar(Frm.hWnd)
50090  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "RemoveX")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function GetUsername() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Const MAX_USERNAME_LENGTH = 128
50020  Dim tStr As String
50030  tStr = String(MAX_USERNAME_LENGTH, Chr$(0))
50040  GetUserNameA tStr, MAX_USERNAME_LENGTH
50050  GetUsername = Left$(tStr, InStr(tStr, Chr$(0)) - 1)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetUsername")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetComputerName() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Const MAX_COMPUTERNAME_LENGTH As Long = 31
50020  Dim tStr As String
50030  tStr = String(MAX_COMPUTERNAME_LENGTH + 1, Chr$(0))
50040  GetComputerNameA tStr, MAX_COMPUTERNAME_LENGTH + 1
50050  GetComputerName = Left$(tStr, MAX_COMPUTERNAME_LENGTH + 1)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetComputerName")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsFormLoaded(Form As Form) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim nForm As Form
50020  For Each nForm In Forms
50030   If nForm Is Form Then
50040    IsFormLoaded = True
50050    Exit For
50060   End If
50070  Next
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "IsFormLoaded")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetDesktop() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim TempDesktop As String, r As clsRegistry
50020
50030  Set r = New clsRegistry
50040  r.hkey = HKEY_CURRENT_USER
50050  r.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
50060  TempDesktop = Trim$(r.GetRegistryValue("Desktop"))
50070  If Right$(TempDesktop, 1) <> "\" Then
50080   GetDesktop = TempDesktop & "\"
50090  End If
50100  Set r = Nothing
50110  If TempDesktop = "" Then
50120   GetDesktop = GetSpecialFolder(ssfDESKTOPDIRECTORY)
50130   If Trim$(GetDesktop) = "" Then
50140    GetDesktop = "C:\"
50150   End If
50160  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetDesktop")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetMyFiles() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim TempMyfiles As String, r As clsRegistry
50020
50030  Set r = New clsRegistry
50040  r.hkey = HKEY_CURRENT_USER
50050  r.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
50060  TempMyfiles = Trim$(r.GetRegistryValue("Personal"))
50070  If Right$(TempMyfiles, 1) <> "\" Then
50080   GetMyFiles = TempMyfiles & "\"
50090  End If
50100  Set r = Nothing
50110  If TempMyfiles = "" Then
50120   GetMyFiles = GetSpecialFolder(ssfPERSONAL)
50130   If Trim$(GetMyFiles) = "" Then
50140    GetMyFiles = "C:\"
50150   End If
50160  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetMyFiles")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetMyAppData() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim TempMyAppData As String, r As clsRegistry
50020
50030  Set r = New clsRegistry
50040  r.hkey = HKEY_CURRENT_USER
50050  r.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
50060  TempMyAppData = Trim$(r.GetRegistryValue("AppData"))
50070  If Right$(TempMyAppData, 1) <> "\" Then
50080   GetMyAppData = TempMyAppData & "\"
50090  End If
50100  Set r = Nothing
50110  If TempMyAppData = "" Then
50120   GetMyAppData = GetSpecialFolder(ssfAPPDATA)
50130   If Trim$(GetMyAppData) = "" Then
50140    GetMyAppData = "C:\"
50150   End If
50160  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetMyAppData")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetFileAttributesStr(FileName As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim hFind As Long, wFD As WIN32_FIND_DATA, Attr As Long, AA As String
50020  hFind = FindFirstFile(FileName, wFD)
50030  Attr = wFD.dwFileAttributes
50040  If Attr And FILE_ATTRIBUTE_ARCHIVE Then AA = AA & "A"
50050  If Attr And FILE_ATTRIBUTE_COMPRESSED Then AA = AA & "C"
50060  If Attr And FILE_ATTRIBUTE_DIRECTORY Then AA = AA & "D"
50070  If Attr And FILE_ATTRIBUTE_HIDDEN Then AA = AA & "H"
50080  If Attr And FILE_ATTRIBUTE_NORMAL Then AA = AA & "N"
50090  If Attr And FILE_ATTRIBUTE_READONLY Then AA = AA & "R"
50100  If Attr And FILE_ATTRIBUTE_SYSTEM Then AA = AA & "S"
50110  GetFileAttributesStr = AA
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetFileAttributesStr")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function CheckPath(PathOrFile As String) As Boolean
 On Error GoTo ErrorHandler
 CheckPath = True
 Dir PathOrFile
 Exit Function
ErrorHandler:
 CheckPath = False
End Function

Public Function LoadDLL(DLLPath As String) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  LoadDLL = LoadLibrary(DLLPath)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "LoadDLL")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub UnLoadDLL(DllHandle As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If DllHandle <> 0 Then
50020   FreeLibrary DllHandle
50030  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "UnLoadDLL")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function MakePath(ByVal Verz As String) As Boolean
 Dim Success As Boolean, dummy As String, Entry As String

 Err = 0: Success = True: dummy = "": Entry = Verz
 On Error Resume Next
 While Len(Entry) > 0 And Success = True
  If Left$(Entry, 1) = "\" Then
    dummy = dummy + "\"
    Entry = Mid$(Entry, 2)
   ElseIf Mid$(Entry, 2, 2) = ":\" Then
    dummy = dummy + Left$(Entry, 3)
    Entry = Mid$(Entry, 4)
  End If
  While Left$(Entry, 1) <> "\" And Len(Entry) > 0
   dummy = dummy + Left$(Entry, 1)
   Entry = Mid$(Entry, 2)
  Wend
  Err = 0
  MkDir dummy
  If Err <> 75 And Err <> 0 Then Success = False
 Wend
 On Error GoTo 0

 MakePath = Success
End Function

Public Function GetDrives() As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim L As Long, Buffer As String, res As Long, drives As String, drv() As String, _
  i As Long
50030  L = 64: Buffer = Space(L)
50040  res = GetLogicalDriveStrings(L, Buffer)
50050  drives = Left$(Buffer, res)
50060  Set GetDrives = New Collection
50070  If Len(drives) > 0 Then
50080   If InStr(drives, Chr$(0)) > 0 Then
50090     drv = Split(drives, Chr$(0))
50100     For i = LBound(drv) To UBound(drv)
50110      If Trim$(drv(i)) <> "" Then
50120       GetDrives.Add drv(i)
50130      End If
50140     Next i
50150    Else
50160     GetDrives.Add drives
50170   End If
50180  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetDrives")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function CompletePath(Path As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Path = Trim$(Path)
50020  If Right$(Path, 1) = "\" Then
50030    CompletePath = Path
50040   Else
50050    CompletePath = Path & "\"
50060  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "CompletePath")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function RunProgramWait(strCmdLine As String, Optional showWindow As Boolean = True) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim lngRetCode As Long, proc As PROCESS_INFORMATION, start As STARTUPINFO, _
  secAttr As SECURITY_ATTRIBUTES, lngRet As Long, lngExit As Long
50030
50040  start.cb = Len(start)
50050  If showWindow = False Then
50060   start.dwFlags = STARTF_USESHOWWINDOW
50070   start.wShowWindow = SW_HIDE
50080  End If
50090
50100  lngRet = CreateProcessA(0&, strCmdLine, _
  0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
50120
50130  GetExitCodeProcess proc.hProcess, lngExit
50140
50150  Do While lngExit = STILL_ACTIVE
50160   DoEvents
50170   Sleep 100
50180   GetExitCodeProcess proc.hProcess, lngExit
50190  Loop
50200  lngRet = CloseHandle(proc.hProcess)
50210
50220  RunProgramWait = lngExit
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "RunProgramWait")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function EnterPasswords(ByRef UserPass As String, ByRef OwnerPass As String, f As Form) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If Options.PDFUserPass = True Or Options.PDFOwnerPass = True Then
50020    With f
50030     .Visible = False
50040     .fraUserPass.Enabled = Options.PDFUserPass
50050     .lblUserPass.Enabled = Options.PDFUserPass
50060     .lblUserPassRepeat.Enabled = Options.PDFUserPass
50070     .fraOwnerPass.Enabled = Options.PDFOwnerPass
50080     .lblOwnerPass.Enabled = Options.PDFUserPass
50090     .lblOwnerPassRepeat.Enabled = Options.PDFUserPass
50100     .iPasswords = Abs(Options.PDFUserPass) + Abs(Options.PDFOwnerPass * 2)
50110     If Options.PDFUserPass = False Then
50120      .txtOwnerPass.SetFocus
50130     End If
50140     .Show
50150     Do
50160      Sleep 100
50170      DoEvents
50180     Loop While .bFinished = False
50190    End With
50200    EnterPasswords = f.bSuccess
50210    UserPass = f.txtUserPass.Text
50220    OwnerPass = f.txtOwnerPass.Text
50230    Unload f
50240   Else
50250    EnterPasswords = False
50260    UserPass = ""
50270    OwnerPass = ""
50280  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "EnterPasswords")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsGoodDrive(tDrive As String) As Boolean
 On Error GoTo ErrorHandler
 Dim tStr As String
 IsGoodDrive = False
 tStr = Dir(Mid$(tDrive, 1, 2) & "\", vbDirectory)
 IsGoodDrive = True
 Exit Function
ErrorHandler:
 Err.Clear
End Function

Public Function IsValidPath(Path As String, Optional ByVal TestUNCPaths As Boolean = True) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, c As Long, nBytes() As Byte, nChar As String, nOK As Boolean, _
  nParentPath As String, nPos As Long, nRoot As String, nPath As String, _
  nPathParts() As String, nStart As Long
50040
50050  Const kCharAsterisk = 42, kCharBackSlash = 92, kCharColon = 58, _
  kCharGreaterThan = 62, kCharLowerThan = 60, kCharPipe = 124, _
  kCharQuestion = 63, kCharQuote = 34, kCharSlash = 47
50080
50090  nPath = Path
50100  PathUnquoteSpaces nPath
50110  nPath = Left$(nPath, InStrNullChar(nPath))
50120  If Not CBool(PathIsUNC(nPath) And Not TestUNCPaths) Then
50130   If CBool(PathFileExists(nPath)) Then
50140    IsValidPath = True
50150    Exit Function
50160   End If
50170  End If
50180  nRoot = nPath
50190  If PathStripToRoot(nRoot) Then
50200   nRoot = Left$(nRoot, InStrNullChar(nRoot))
50210   If PathIsUNC(nRoot) Then
50220     nPos = InStrRev(nRoot, "\")
50230     nRoot = Left$(nRoot, nPos - 1)
50240     If PathIsUNCServer(nRoot) Then
50250      nPath = Mid$(nPath, Len(nRoot) + 1)
50260     End If
50270    Else
50280     nPath = Mid$(nPath, Len(nRoot) + 1)
50290   End If
50300   If Len(nPath) Then
50310    nPathParts = Split(nPath, "\")
50320    If Len(nPathParts(0)) Then
50330      nStart = 0
50340     Else
50350      nStart = 1
50360    End If
50370    For i = nStart To UBound(nPathParts)
50380     nBytes = StrConv(nPathParts(i), vbFromUnicode)
50390     For c = 0 To UBound(nBytes)
50400      Select Case nBytes(c)
      Case Is < 32, Is > 255, kCharAsterisk, kCharBackSlash, kCharColon, kCharGreaterThan, kCharLowerThan, kCharPipe, kCharQuestion, kCharQuote, kCharSlash
50420        nOK = False
50430        Exit For
50440       Case Else
50450        nOK = True
50460      End Select
50470     Next c
50480     If Not nOK Then
50490      Exit For
50500     End If
50510    Next i
50520    IsValidPath = nOK
50530   End If
50540  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "IsValidPath")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub FormInTaskbar(Form As Form, ByVal ShowInTaskBar As Boolean, Optional ByVal NoFlicker As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim nStyle As Long, nVisible As Boolean
50020  Const GWL_EXSTYLE = (-20), WS_EX_APPWINDOW = &H40000
50030  With Form
50040   nVisible = .Visible
50050   If NoFlicker And nVisible Then
50060    LockWindowUpdate GetDesktopWindow()
50070   End If
50080   .Visible = False
50090   nStyle = GetWindowLong(.hWnd, GWL_EXSTYLE)
50100   Select Case ShowInTaskBar
   Case True
50120     nStyle = nStyle Or WS_EX_APPWINDOW
50130    Case False
50140     nStyle = nStyle And Not WS_EX_APPWINDOW
50150   End Select
50160   SetWindowLong .hWnd, GWL_EXSTYLE, nStyle
50170   .Refresh
50180   .Visible = nVisible
50190   If NoFlicker And nVisible Then
50200    LockWindowUpdate 0&
50210   End If
50220  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "FormInTaskbar")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
