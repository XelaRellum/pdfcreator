Attribute VB_Name = "modGeneral"
Option Explicit

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function GetDriveType Lib "Kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Const MAX_PATH = 260

Private Const SE_ERR_NOASSOC = 31
Private Const SE_ERR_NOTFOUND = 2

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetTempPathA Lib "Kernel32" _
 (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetSystemDirectory Lib "Kernel32" _
 Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectoryA Lib "Kernel32" _
 (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Function GetTempFileNameA Lib "Kernel32" _
 (ByVal lpszPath As String, ByVal lpPrefixString As String, _
  ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

Public Declare Function ShellExecute Lib "shell32.dll" _
  Alias "ShellExecuteA" (ByVal hwnd As Long, _
  ByVal lpOperation As String, ByVal lpFile As String, _
  ByVal lpParameters As String, _
  ByVal lpDirectory As String, _
  ByVal nShowCmd As Long) As Long


Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
      
Private Const MF_BYPOSITION = &H400
Private Const MF_REMOVE = &H1000

Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "Kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

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

Private Declare Function FindFirstFile Lib "Kernel32" _
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
Private Declare Function GetCurrentProcess Lib "Kernel32" () As Long
Private Declare Function OpenProcessToken Lib "Advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetUserProfileDirectory Lib "userenv.dll" Alias "GetUserProfileDirectoryA" (ByVal hToken As Long, ByVal lpProfileDir As String, lpcchSize As Long) As Boolean

Public Type PROFILEINFO
 dwSize As Long
 dwFlags As Long
 lpUsername As String
 lpProfilePath As String
 lpDefaultPath As String
 lpServerName As String
 lpPolicyPath As String
 hProfile As Long
End Type

Private Declare Function LoadUserProfile Lib "userenv.dll" (ByVal hToken As Long, ByRef lpProfileInfo As PROFILEINFO) As Long

Private Declare Function LoadLibrary Lib "Kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "Kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "Kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long

Private Declare Function GetLogicalDriveStrings Lib "Kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

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

Private Declare Function GetExitCodeProcess Lib "Kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CreateProcessA Lib "Kernel32" (ByVal _
   lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
   lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
   lpStartupInfo As STARTUPINFO, lpProcessInformation As _
   PROCESS_INFORMATION) As Long
Public Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long

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

'---------------------------------------------------
'API's for FormInTaskbar
'Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long

Public Const MF_BYCOMMAND = &H0&
Public Const MF_BITMAP = &H4&
'---------------------------------------------------
'Api's for the help control
Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
 (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, _
 ByVal dwData As Long) As Long
  

Private Declare Function HtmlHelpTopic Lib "hhctrl.ocx" Alias "HtmlHelpA" _
 (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, _
  ByVal dwData As String) As Long

Public Const HH_DISPLAY_TOPIC = &H0
Public Const HH_SET_WIN_TYPE = &H4
Public Const HH_GET_WIN_TYPE = &H5
Public Const HH_GET_WIN_HANDLE = &H6
Public Const HH_DISPLAY_TEXT_POPUP = &HE   ' Display string resource ID or text in a pop-up window.
Public Const HH_HELP_CONTEXT = &HF         ' Display mapped numeric value in  dwData.
Public Const HH_TP_HELP_CONTEXTMENU = &H10 ' Text pop-up help, similar to WinHelp's HELP_CONTEXTMENU.
Public Const HH_TP_HELP_WM_HELP = &H11     ' text pop-up help, similar to WinHelp's HELP_WM_HELP.
Public Const HH_CLOSE_ALL = &H12
'---------------------------------------------------
Private Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExA" (lpExecInfo As SHELLEXECUTEINFO) As Long
Private Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Type SHELLEXECUTEINFO
 cbSize As Long
 fMask As Long
 hwnd As Long
 lpVerb As String
 lpFile As String
 lpParameters As String
 lpDirectory As String
 nShow As Long
 hInstApp As Long
 lpIDList As Long
 lpClass As String
 hkeyClass As Long
 dwHotKey As Long
 hIcon As Long
 hProcess As Long
End Type

Private Const SEE_MASK_CLASSKEY = &H3
Private Const SEE_MASK_CLASSNAME = &H1
Private Const SEE_MASK_CONNECTNETDRV = &H80
Private Const SEE_MASK_DOENVSUBST = &H200
Private Const SEE_MASK_FLAG_DDEWAIT = &H100
Private Const SEE_MASK_FLAG_NO_UI = &H400
Private Const SEE_MASK_HOTKEY = &H20
Private Const SEE_MASK_ICON = &H10
Private Const SEE_MASK_IDLIST = &H4
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40

Private Const SW_MAXIMIZE = 3
Private Const SW_MINIMIZE = 6
Private Const SW_RESTORE = 9
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMINNOACTIVE = 7

Private Const SW_SHOWNA = 8
Private Const SW_SHOWNOACTIVATE = 4
Private Const SW_SHOWNORMAL = 1

Private Const SE_ERR_ACCESSDENIED = 5
Private Const SE_ERR_ASSOCINCOMPLETE = 27
Private Const SE_ERR_DDEBUSY = 30
Private Const SE_ERR_DDEFAIL = 29
Private Const SE_ERR_DDETIMEOUT = 28
Private Const SE_ERR_DLLNOTFOUND = 32
Private Const SE_ERR_FNF = 2
Private Const SE_ERR_OOM = 8
Private Const SE_ERR_PNF = 3
Private Const SE_ERR_SHARE = 26

Private Const HOTKEYF_ALT = &H4
Private Const HOTKEYF_CONTROL = &H2
Private Const HOTKEYF_EXT = &H8
Private Const HOTKEYF_SHIFT = &H1

Private Const INFINITE = &HFFFF

Private Const WAIT_ABANDONED = &H80
Private Const WAIT_FAILED = &HFFFFFFFF
Private Const WAIT_OBJECT_0 = &H0
Private Const WAIT_TIMEOUT = &H102

Public Enum ShellAction
 Aopen = 0 'open
 APrint = 1 'print
 AExplore = 2 'explore
End Enum
 
Public Enum eSaveOpenType
 saveFile = 0
 OpenFile = 1
End Enum

Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationAborted As Long
    hNameMaps As Long
    sProgress As String
End Type

Private Const FO_MOVE = &H1
Private Const FO_COPY = &H2&
Private Const FO_DELETE = &H3&
Private Const FO_RENAME = &H4&

Private Const FOF_MULTIDESTFILES = &H1&
Private Const FOF_CONFIRMMOUSE = &H2&
Private Const FOF_SILENT = &H4   ' 'don't create progress/report
Private Const FOF_NOCONFIRMATION = &H10&    ''Don't prompt the user.
Private Const FOF_ALLOWUNDO = &H40&
Private Const FOF_NOCONFIRMMKDIR = &H200&

Private Declare Function SHFileOperation _
    Lib "shell32.dll" Alias "SHFileOperationA" _
    (lpFileOp As Any) As Long

Private Declare Sub MoveMemory Lib "Kernel32" _
    Alias "RtlMoveMemory" _
    (pDest As Any, _
    pSource As Any, _
    ByVal dwLength As Long)

Private Declare Sub OemToChar Lib "user32" Alias "OemToCharA" (ByVal StrFrom As String, ByVal StrTo As String)
Private Declare Sub CharToOem Lib "user32" Alias "CharToOemA" (ByVal StrFrom As String, ByVal StrTo As String)

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

Public Sub SplitPath(FullPath As String, Optional Drive As String, Optional Path As String, Optional Filename As String, Optional File As String, Optional Extension As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim nPos As Integer
50020  nPos = InStrRev(FullPath, "\")
50030  If nPos > 0 Then
50040    If Left$(FullPath, 2) = "\\" Then
50050     If nPos = 2 Then
50060      Drive = FullPath: Path = vbNullString: Filename = vbNullString: File = vbNullString
50070      Extension = vbNullString
50080      Exit Sub
50090     End If
50100    End If
50110    Path = Left$(FullPath, nPos - 1)
50120    Filename = Mid$(FullPath, nPos + 1)
50130    nPos = InStrRev(Filename, ".")
50140    If nPos > 0 Then
50150      File = Left$(Filename, nPos - 1)
50160      Extension = Mid$(Filename, nPos + 1)
50170     Else
50180      File = Filename
50190      Extension = vbNullString
50200    End If
50210   Else
50220    nPos = InStrRev(FullPath, ":")
50230    If nPos > 0 Then
50240      Path = Mid(FullPath, 1, nPos - 1): Filename = Mid(FullPath, nPos + 1)
50250      nPos = InStrRev(Filename, ".")
50260      If nPos > 0 Then
50270        File = Left$(Filename, nPos - 1)
50280        Extension = Mid$(Filename, nPos + 1)
50290       Else
50300        File = Filename
50310        Extension = vbNullString
50320      End If
50330     Else
50340      Path = vbNullString: Filename = FullPath
50350      nPos = InStrRev(Filename, ".")
50360      If nPos > 0 Then
50370        File = Left$(Filename, nPos - 1)
50380        Extension = Mid$(Filename, nPos + 1)
50390       Else
50400        File = Filename
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
50010  Dim res As Long, TempDir As String
50020
50030  If IsWin95 = True Or IsWin98 = True Or IsWinME = True Then
50040    TempDir = GetTempPathApi
50050    If LenB(Environ$("Redmon_User")) > 0 Then
50060      TempDir = CompletePath(TempDir) & "PDFCreator\" & Environ$("Redmon_User")
50070     Else
50080      TempDir = CompletePath(TempDir) & "PDFCreator\" & GetUsername
50090    End If
50100   Else
50110    If IsWinNT4 = True Then
50120      TempDir = GetMyAppData
50130      If LenB(Environ$("Redmon_User")) > 0 Then
50140        TempDir = CompletePath(TempDir) & "PDFCreator\" & Environ$("Redmon_User")
50150       Else
50160        TempDir = CompletePath(TempDir) & "PDFCreator\" & GetUsername
50170      End If
50180     Else
50190      TempDir = GetMyLocalAppData
50200      TempDir = Left(TempDir, InStrRev(TempDir, "\")) & "Temp\PDFCreator\"
50210    End If
50220  End If
50230  If Trim$(TempDir) = "" Then
50240   TempDir = "C:\Temp\"
50250  End If
50260
50270  GetTempPath = CompletePath(TempDir)
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

Public Function GetTempPathApi() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim res As Long, TempDir As String
50020
50030  TempDir = Space$(MAX_PATH)
50040  res = GetTempPathA(MAX_PATH, TempDir)
50050  If res > 0 Then
50060    GetTempPathApi = Left$(TempDir, res)
50070   Else
50080    GetTempPathApi = "C:\"
50090  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetTempPathApi")
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
50040  If DirExists(tPath) = False Then
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

Public Sub CombineFiles(ByVal Filename As String, Files As Collection, Optional stb As StatusBar, _
 Optional BufferSize As Long = 65536#)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, j As Long, fnSource As Long, fnDest As Long, sBuffer As String, _
  aLen As Double, tLen As Double, bsize As Long, fpos As Long
50030
50040  bsize = BufferSize
50050  Filename = Trim$(Filename)
50060  If Filename = vbNullString Or Files.Count = 0 Or Right$(Filename, 1) = "\" Then
50070   Exit Sub
50080  End If
50090 ' If Len(Dir(FileName)) > 0 Then
50100 '  Exit Sub
50110 ' End If
50120  If Files.Count = 1 Then
50130   Exit Sub
50140  End If
50150  fnDest = FreeFile
50160  aLen = 0: tLen = 0
50170  For i = 1 To Files.Count
50180   aLen = aLen + FileLen(Files.item(i))
50190  Next i
50200  Open Filename For Binary As #fnDest
50210  For i = 1 To Files.Count
50220   DoEvents
50230   If FileLen(Files.item(i)) > 0 Then
50240    fnSource = FreeFile
50250    Open Files.item(i) For Binary Access Read As #fnSource
50260    If bsize > LOF(fnSource) Then
50270     bsize = LOF(fnSource)
50280    End If
50290    fpos = 1
50300    For j = 1 To LOF(fnSource) \ bsize
50310     fpos = (j - 1) * bsize + 1
50320     Seek #fnSource, fpos
50330     sBuffer = Input(bsize, fnSource)
50340     Put #fnDest, , sBuffer
50350     tLen = tLen + bsize
50360     If IsObject(stb) = True Then
50370 '     stb.Panels("Percent").Text = Format$(tLen / aLen, "0.0%")
50380     End If
50390     DoEvents
50400    Next j
50410    If LOF(fnSource) > (j - 1) * bsize Then
50420     fpos = (j - 1) * bsize + 1
50430     Seek #fnSource, fpos
50440     sBuffer = Input(LOF(fnSource) - (j - 1) * bsize, fnSource)
50450     Put #fnDest, , sBuffer
50460     tLen = tLen + (LOF(fnSource) - (j - 1) * bsize)
50470 '    stb.Panels("Percent").Text = Format$(tLen / aLen, "0.0%")
50480    End If
50490    Close #fnSource
50500   End If
50510   Kill Files.item(i)
50520   DoEvents
50530  Next i
50540  Close #fnDest
50550 ' stb.Panels("Percent").Text = vbNullString
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

Public Sub CombineFilesOld(ByVal Filename As String, Files As Collection, stb As StatusBar)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, fnSource As Long, fnDest As Long, sBuffer As String, _
  aLen As Double, tLen As Double
50030
50040  Filename = Trim$(Filename)
50050  If Filename = vbNullString Or Files.Count = 0 Or Right$(Filename, 1) = "\" Then
50060   Exit Sub
50070  End If
50080  If Len(Dir(Filename)) > 0 Then
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
50190  Open Filename For Binary As #fnDest
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
Select Case ErrPtnr.OnError("modGeneral", "CombineFilesOld")
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
50040   If TypeOf ctl Is label Or _
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
50030  Forbiddenchars = "\/:*?<>|%"""
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
50030  hMenu = GetSystemMenu(Frm.hwnd, 0)
50040  If hMenu <> 0 Then
50050   nPosition = GetMenuItemCount(hMenu)
50060   Call RemoveMenu(hMenu, nPosition - 1, MF_REMOVE Or MF_BYPOSITION)
50070   Call RemoveMenu(hMenu, nPosition - 2, MF_REMOVE Or MF_BYPOSITION)
50080   Call DrawMenuBar(Frm.hwnd)
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
50050  GetComputerName = Left$(tStr, InStr(tStr, Chr$(0)) - 1)
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
50010  Dim TempDesktop As String, reg As clsRegistry
50020  Set reg = New clsRegistry
50030  reg.hkey = HKEY_CURRENT_USER
50040  reg.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
50050  TempDesktop = reg.GetRegistryValue("Desktop")
50060  Set reg = Nothing
50070  If Trim$(TempDesktop) = "" Then
50080   TempDesktop = "C:\"
50090  End If
50100  GetDesktop = TempDesktop
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
50010  Dim TempMyFiles, reg As clsRegistry
50020  Set reg = New clsRegistry
50030  reg.hkey = HKEY_CURRENT_USER
50040  reg.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
50050  TempMyFiles = reg.GetRegistryValue("Personal")
50060  Set reg = Nothing
50070  If Trim$(TempMyFiles) = "" Then
50080   TempMyFiles = "C:\"
50090  End If
50100  GetMyFiles = TempMyFiles
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
50010  Dim TempMyAppData As String, reg As clsRegistry
50020  Set reg = New clsRegistry
50030  reg.hkey = HKEY_CURRENT_USER
50040  reg.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
50050  TempMyAppData = reg.GetRegistryValue("AppData")
50060  Set reg = Nothing
50070  If Trim$(TempMyAppData) = "" Then
50080   TempMyAppData = "C:\"
50090  End If
50100  GetMyAppData = TempMyAppData
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

Public Function GetMyLocalAppData() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim TempMyLocalAppData As String, reg As clsRegistry
50020  Set reg = New clsRegistry
50030  reg.hkey = HKEY_CURRENT_USER
50040  reg.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
50050  TempMyLocalAppData = reg.GetRegistryValue("Local AppData")
50060  Set reg = Nothing
50070  If Trim$(TempMyLocalAppData) = "" Then
50080   TempMyLocalAppData = "C:\"
50090  End If
50100  GetMyLocalAppData = TempMyLocalAppData
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetMyLocalAppData")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetFileAttributesStr(Filename As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim hFind As Long, WFD As WIN32_FIND_DATA, attr As Long, AA As String
50020  hFind = FindFirstFile(Filename, WFD)
50030  attr = WFD.dwFileAttributes
50040  If attr And FILE_ATTRIBUTE_ARCHIVE Then AA = AA & "A"
50050  If attr And FILE_ATTRIBUTE_COMPRESSED Then AA = AA & "C"
50060  If attr And FILE_ATTRIBUTE_DIRECTORY Then AA = AA & "D"
50070  If attr And FILE_ATTRIBUTE_HIDDEN Then AA = AA & "H"
50080  If attr And FILE_ATTRIBUTE_NORMAL Then AA = AA & "N"
50090  If attr And FILE_ATTRIBUTE_READONLY Then AA = AA & "R"
50100  If attr And FILE_ATTRIBUTE_SYSTEM Then AA = AA & "S"
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
 Dir PathOrFile, vbDirectory + vbHidden + vbSystem + vbReadOnly
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
50010  If Len(Path) = 0 Then
50020   Exit Function
50030  End If
50040  Path = Trim$(Path)
50050  If Right$(Path, 1) = "\" Then
50060    CompletePath = LTrim$(Path)
50070   Else
50080    CompletePath = LTrim$(Path) & "\"
50090  End If
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
50010  Dim lngRetCode As Long, proc As PROCESS_INFORMATION, Start As STARTUPINFO, _
  secAttr As SECURITY_ATTRIBUTES, lngRet As Long, lngExit As Long
50030
50040  Start.cb = Len(Start)
50050  If showWindow = False Then
50060   Start.dwFlags = STARTF_USESHOWWINDOW
50070   Start.wShowWindow = SW_HIDE
50080  End If
50090
50100  lngRet = CreateProcessA(0&, strCmdLine, _
  0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, Start, proc)
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
50010  If Options.PDFUserPass <> 0 Or Options.PDFOwnerPass <> 0 Then
50020    With f
50030     .Visible = False
50040     .fraUserPass.Enabled = Options.PDFUserPass
50050     .lblUserPass.Enabled = Options.PDFUserPass
50060     .lblUserPassRepeat.Enabled = Options.PDFUserPass
50070     .fraOwnerPass.Enabled = Options.PDFOwnerPass
50080     .lblOwnerPass.Enabled = Options.PDFOwnerPass
50090     .lblOwnerPassRepeat.Enabled = Options.PDFOwnerPass
50100     .iPasswords = Abs(Options.PDFUserPass) + Abs(Options.PDFOwnerPass * 2)
50110     .Show vbModal
50120     Do
50130      Sleep 100
50140      DoEvents
50150     Loop While .bFinished = False
50160    End With
50170    EnterPasswords = f.bSuccess
50180    UserPass = f.txtUserPass.Text
50190    OwnerPass = f.txtOwnerPass.Text
50200    Unload f
50210   Else
50220    EnterPasswords = False
50230    UserPass = ""
50240    OwnerPass = ""
50250  End If
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
 tStr = Dir(Mid$(tDrive, 1, 2) & "\", vbDirectory + vbHidden)
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
50401      Select Case nBytes(c)
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
50090   nStyle = GetWindowLong(.hwnd, GWL_EXSTYLE)
50101   Select Case ShowInTaskBar
         Case True
50120     nStyle = nStyle Or WS_EX_APPWINDOW
50130    Case False
50140     nStyle = nStyle And Not WS_EX_APPWINDOW
50150   End Select
50160   SetWindowLong .hwnd, GWL_EXSTYLE, nStyle
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

Public Function GetStringRessource(DLLExeFilename As String, ResID As Long) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim hLib As Long, sBuf As String, lLen As Long
50020
50030  hLib = LoadLibrary(DLLExeFilename)
50040  If hLib Then
50050   sBuf = Space$(255)
50060   lLen = LoadString(hLib, ResID, sBuf, Len(sBuf))
50070   If lLen Then
50080     sBuf = Left$(sBuf, lLen)
50090    Else
50100     sBuf = vbNullString
50110   End If
50120   FreeLibrary hLib
50130  End If
50140
50150  GetStringRessource = sBuf
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "GetStringRessource")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function


Private Sub ReplaceStrInFile(Filename As String, SearchStr As String, _
 ReplaceStr As String, Optional BufferSize As Long = 65536)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Content As String, iSearch As Long, pLen As Long, chunk As Long, fpos As Long, _
  fnI As Long, fnO As Long, LSearchStr As Long, LReplaceStr As Long, _
  Tempfile As String
50040  If Len(SearchStr) = 0 Or Len(ReplaceStr) = 0 Then
50050   Exit Sub
50060  End If
50070  fnI = FreeFile
50080  Open Filename For Binary Access Read As #fnI
50090  fnO = FreeFile
50100  Tempfile = GetTempFile(CompletePath(GetPDFCreatorTempfolder), "~RE")
50110  Open Tempfile For Output As #fnO
50120  LSearchStr = Len(SearchStr)
50130  LReplaceStr = Len(ReplaceStr)
50140  chunk = BufferSize
50150  fpos = 1: pLen = LOF(fnI)
50160  Do Until pLen = 0
50170   If LOF(fnI) - fpos + 1 < chunk Then
50180    chunk = LOF(fnI) - fpos + 1
50190   End If
50200   Content = Space$(chunk)
50210   Get #fnI, fpos, Content
50220   iSearch = InStrRev(Content, SearchStr)
50230   If iSearch = 0 Then
50240     pLen = chunk - LSearchStr + 1
50250     DoEvents
50260     Print #fnO, Left(Content, pLen);
50270     DoEvents
50280     fpos = fpos + pLen
50290    Else
50300     pLen = iSearch + LSearchStr - 1
50310     If pLen < chunk - LSearchStr + 1 Then
50320      pLen = chunk - LSearchStr + 1
50330     End If
50340     fpos = fpos + pLen
50350     Do Until iSearch = 0
50360      Content = Left(Content, iSearch - 1) & ReplaceStr & Mid(Content, iSearch + LSearchStr)
50370      iSearch = InStrRev(Content, SearchStr, iSearch)
50380      pLen = pLen + LReplaceStr - LSearchStr
50390     Loop
50400     DoEvents
50410     Print #fnO, Left(Content, pLen);
50420     DoEvents
50430   End If
50440  Loop
50450  Print #fnO, Right(Content, LSearchStr - 1);
50460  Close #fnO: Close #fnI
50470  Name Tempfile As Filename
50480  DoEvents
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "ReplaceStrInFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub HTMLHelp_ShowTopic(Optional sTopicFile As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim sHelpFile As String, hwndHelp As Long
50020
50030  sHelpFile = App.Path & "\PDFCreator.chm"
50040  If sTopicFile = "" Then
50050    hwndHelp = HtmlHelp(0, sHelpFile, HH_DISPLAY_TOPIC, ByVal 0&)
50060   Else
50070    hwndHelp = HtmlHelpTopic(0, sHelpFile, HH_DISPLAY_TOPIC, sTopicFile)
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "HTMLHelp_ShowTopic")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function DirExists(DirStr As String) As Boolean
 On Error GoTo ErrorHandler
 DirExists = GetAttr(DirStr) Or vbDirectory
 Exit Function
ErrorHandler:
 DirExists = False
End Function

Public Function FileExists(FileStr As String) As Boolean
 On Error GoTo ErrorHandler
 FileExists = GetAttr(FileStr)
 Exit Function
ErrorHandler:
 FileExists = False
End Function

Public Function UnLoadDLL(DLLHandle As Long) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim res As Long
50020  If DLLHandle <> 0 Then
50030   res = FreeLibrary(DLLHandle)
50040   If res <> 0 Then
50050     UnLoadDLL = True
50060    Else
50070     UnLoadDLL = False
50080   End If
50090  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "UnLoadDLL")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function UnloadDLLComplete(ByRef DLLHandle As Long) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim c As Long
50020  If DLLHandle = 0 Then
50030   Exit Function
50040  End If
50050  c = 0
50060  Do While UnLoadDLL(DLLHandle)
50070   c = c + 1
50080   DoEvents
50090  Loop
50100  UnloadDLLComplete = c: DLLHandle = 0
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "UnloadDLLComplete")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function ExecuteAndWait(hwnd As Long, Filename As String, Action As ShellAction) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim RetVal As Long, ShExInfo As SHELLEXECUTEINFO
50020
50030  With ShExInfo
50040   .cbSize = Len(ShExInfo)
50050 '  .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_FLAG_NO_UI Or SEE_MASK_CLASSNAME Or _
           SEE_MASK_FLAG_DDEWAIT
50070   .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI Or _
           SEE_MASK_FLAG_DDEWAIT
50090   .hwnd = hwnd
50101   Select Case Action
         Case ShellAction.APrint
50120     .lpVerb = "print"
50130    Case ShellAction.Aopen
50140     .lpVerb = "open"
50150    Case ShellAction.AExplore
50160     .lpVerb = "explore"
50170    End Select
50180
50190   .lpFile = Filename
50200   .lpParameters = vbNullChar
50210   .lpDirectory = vbNullChar
50220   .nShow = SW_SHOWMINIMIZED
50230  End With
50240
50250  ExecuteAndWait = ShellExecuteEx(ShExInfo)
50260
50270  Do
50280   DoEvents
50290  Loop Until WaitForSingleObject(ShExInfo.hProcess, 0) <> WAIT_TIMEOUT
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "ExecuteAndWait")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function FileISPrintable(Filename As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Ext As String, reg As clsRegistry
50020  FileISPrintable = False
50030  If Len(Filename) = 0 Then
50040   Exit Function
50050  End If
50060  SplitPath Filename, , , , , Ext
50070  If Len(Ext) = 0 Then
50080   Exit Function
50090  End If
50100  Set reg = New clsRegistry
50110  reg.hkey = HKEY_LOCAL_MACHINE
50120  reg.KeyRoot = "Software\CLASSES\." & Ext
50130  If reg.KeyExists = False Then
50140   Set reg = Nothing
50150   Exit Function
50160  End If
50170  reg.KeyRoot = "Software\CLASSES\" & reg.GetRegistryValue("") & "\shell\print"
50180  If reg.KeyExists = False Then
50190   Set reg = Nothing
50200   Exit Function
50210  End If
50220  FileISPrintable = True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "FileISPrintable")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub AddExplorerIntegration()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry, Keys As Collection, i As Long, sKey As String, Path As String
50020  Set reg = New clsRegistry
50030  reg.hkey = HKEY_CLASSES_ROOT
50040  Set Keys = reg.EnumRegistryKeys(HKEY_CLASSES_ROOT, "")
50050  For i = 1 To Keys.Count
50060   If Mid(Keys(i), 1, 1) = "." Then
50070    reg.KeyRoot = Keys(i)
50080    reg.Subkey = ""
50090    sKey = reg.GetRegistryValue("")
50100    If LenB(sKey) > 0 Then
50110     reg.KeyRoot = sKey
50120     If reg.KeyExists = True Then
50130      reg.Subkey = "shell\print\command"
50140      If reg.KeyExists = True Then
50150       If LenB(Trim$(reg.GetRegistryValue(""))) > 0 Then
50160        reg.KeyRoot = sKey & "\shell\" & Uninstall_GUID
50170        reg.Subkey = ""
50180        If reg.KeyExists = False Then
50190         Path = CompletePath(GetPDFCreatorApplicationPath)
50200         If Len(Path) > 1 Then
50210          reg.CreateKey
50220          reg.SetRegistryValue "", LanguageStrings.OptionsShellIntegrationCaption, REG_SZ
50230          reg.CreateKey "command"
50240          reg.KeyRoot = sKey & "\shell\" & Uninstall_GUID & "\command"
50250          reg.SetRegistryValue "", Path & "pdfcreator.exe -PF""%1"" -NS", REG_SZ
50260         End If
50270        End If
50280       End If
50290      End If
50300     End If
50310    End If
50320   End If
50330  Next i
50340  Set Keys = Nothing
50350  Set reg = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "AddExplorerIntegration")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub RemoveExplorerIntegration()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry, Keys As Collection, i As Long
50020  Set reg = New clsRegistry
50030  reg.hkey = HKEY_CLASSES_ROOT
50040  Set Keys = reg.EnumRegistryKeys(HKEY_CLASSES_ROOT, "")
50050  For i = 1 To Keys.Count
50060   reg.KeyRoot = Keys(i) & "\shell"
50070   reg.Subkey = Uninstall_GUID & "\command"
50080   If reg.KeyExists Then
50090    reg.DeleteKey reg.Subkey
50100   End If
50110   reg.Subkey = Uninstall_GUID
50120   If reg.KeyExists Then
50130    reg.DeleteKey reg.Subkey
50140   End If
50150  Next i
50160  Set Keys = Nothing
50170  Set reg = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "RemoveExplorerIntegration")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function RemoveCompletePath(Path As String, Optional ShowDialog As Boolean) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ret As Long, SHFILEOP As SHFILEOPSTRUCT, ShowD As Long, FOBuf() As Byte, _
  LenSHFileOp As Long
50030
50040  LenSHFileOp = LenB(SHFILEOP)
50050  ReDim FOBuf(1 To LenSHFileOp)
50060  If ShowDialog Then
50070    ShowD = 0
50080   Else
50090    ShowD = FOF_SILENT
50100  End If
50110
50120  With SHFILEOP
50130   .wFunc = FO_DELETE
50140   .pFrom = Path
50150   .pTo = vbNullString
50160   .fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR Or ShowD
50170  End With
50180
50190  ret = SHFileOperation(SHFILEOP)
50200  DoEvents
50210
50220  Call MoveMemory(FOBuf(1), SHFILEOP, LenSHFileOp)
50230  Call MoveMemory(FOBuf(21), FOBuf(19), 12)
50240  Call MoveMemory(SHFILEOP, FOBuf(1), LenSHFileOp)
50250
50260  If ret <> 0 Then
50270   If SHFILEOP.fAnyOperationAborted <> 0 Then
50280    ret = -1
50290   End If
50300  End If
50310  RemoveCompletePath = ret
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "RemoveCompletePath")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function ASCIItoANSI(ByVal AsciiString As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  OemToChar AsciiString, AsciiString
50020  ASCIItoANSI = AsciiString
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "ASCIItoANSI")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function ANSItoASCII(ByVal AnsiString As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  CharToOem AnsiString, AnsiString
50020  ANSItoASCII = AnsiString
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral", "ANSItoASCII")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function


