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
 lpUserName As String
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
     Drive = FullPath: Path = vbNullString: Filename = vbNullString: File = vbNullString
     Extension = vbNullString
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
     Extension = vbNullString
   End If
  Else
   nPos = InStrRev(FullPath, ":")
   If nPos > 0 Then
     Path = Mid(FullPath, 1, nPos - 1): Filename = Mid(FullPath, nPos + 1)
     nPos = InStrRev(Filename, ".")
     If nPos > 0 Then
       File = Left$(Filename, nPos - 1)
       Extension = Mid$(Filename, nPos + 1)
      Else
       File = Filename
       Extension = vbNullString
     End If
    Else
     Path = vbNullString: Filename = FullPath
     nPos = InStrRev(Filename, ".")
     If nPos > 0 Then
       File = Left$(Filename, nPos - 1)
       Extension = Mid$(Filename, nPos + 1)
      Else
       File = Filename
       Extension = vbNullString
     End If
   End If
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
 Dim res As Long, Tempdir As String

 If Major > viMajorWindows9x_NT4 Then
   Tempdir = GetMyLocalAppData
  Else
   Tempdir = GetMyAppData
 End If

 If Trim$(Tempdir) = "" Then
   Tempdir = Space$(MAX_PATH)
   res = GetTempPathA(MAX_PATH, Tempdir)
   If res > 0 Then
     Tempdir = CompletePath(Left$(Tempdir, res))
    Else
     Tempdir = "C:\Temp\"
   End If
  Else
   Tempdir = Left(Tempdir, InStrRev(Tempdir, "\")) & "Temp\"
 End If


 If LenB(Environ$("Redmon_User")) > 0 Then
   Tempdir = CompletePath(Tempdir) & "PDFCreator\" & Environ$("Redmon_User")
  Else
   Tempdir = CompletePath(Tempdir) & "PDFCreator\" & GetUsername
 End If

 GetTempPath = CompletePath(Tempdir)
End Function

Public Function GetTempPathApi() As String
 Dim res As Long, Tempdir As String

 Tempdir = Space$(MAX_PATH)
 res = GetTempPathA(MAX_PATH, Tempdir)
 If res > 0 Then
   GetTempPathApi = Left$(Tempdir, res)
  Else
   GetTempPathApi = "C:\"
 End If
End Function

Public Function GetTempFile(Optional ByVal Path As String, Optional Prefix As String) As String
 Dim res As Long, Tempfile As String, tPath As String
 Tempfile = Space$(MAX_PATH)
 tPath = Trim$(Path)
 If DirExists(tPath) = False Then
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

Public Sub CombineFiles(ByVal Filename As String, Files As Collection, Optional stb As StatusBar, _
 Optional BufferSize As Long = 65536#)
 Dim i As Long, j As Long, fnSource As Long, fnDest As Long, sBuffer As String, _
  aLen As Double, tLen As Double, bsize As Long, fpos As Long

 bsize = BufferSize
 Filename = Trim$(Filename)
 If Filename = vbNullString Or Files.Count = 0 Or Right$(Filename, 1) = "\" Then
  Exit Sub
 End If
' If Len(Dir(FileName)) > 0 Then
'  Exit Sub
' End If
 If Files.Count = 1 Then
  Exit Sub
 End If
 fnDest = FreeFile
 aLen = 0: tLen = 0
 For i = 1 To Files.Count
  aLen = aLen + FileLen(Files.item(i))
 Next i
 Open Filename For Binary As #fnDest
 For i = 1 To Files.Count
  DoEvents
  If FileLen(Files.item(i)) > 0 Then
   fnSource = FreeFile
   Open Files.item(i) For Binary Access Read As #fnSource
   If bsize > LOF(fnSource) Then
    bsize = LOF(fnSource)
   End If
   fpos = 1
   For j = 1 To LOF(fnSource) \ bsize
    fpos = (j - 1) * bsize + 1
    Seek #fnSource, fpos
    sBuffer = Input(bsize, fnSource)
    Put #fnDest, , sBuffer
    tLen = tLen + bsize
    If IsObject(stb) = True Then
'     stb.Panels("Percent").Text = Format$(tLen / aLen, "0.0%")
    End If
    DoEvents
   Next j
   If LOF(fnSource) > (j - 1) * bsize Then
    fpos = (j - 1) * bsize + 1
    Seek #fnSource, fpos
    sBuffer = Input(LOF(fnSource) - (j - 1) * bsize, fnSource)
    Put #fnDest, , sBuffer
    tLen = tLen + (LOF(fnSource) - (j - 1) * bsize)
'    stb.Panels("Percent").Text = Format$(tLen / aLen, "0.0%")
   End If
   Close #fnSource
  End If
  Kill Files.item(i)
  DoEvents
 Next i
 Close #fnDest
' stb.Panels("Percent").Text = vbNullString
End Sub

Public Sub CombineFilesOld(ByVal Filename As String, Files As Collection, stb As StatusBar)
 Dim i As Long, fnSource As Long, fnDest As Long, sBuffer As String, _
  aLen As Double, tLen As Double

 Filename = Trim$(Filename)
 If Filename = vbNullString Or Files.Count = 0 Or Right$(Filename, 1) = "\" Then
  Exit Sub
 End If
 If Len(Dir(Filename)) > 0 Then
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
 stb.Panels("Percent").Text = vbNullString
End Sub

Public Sub SetFont(Frm As Form, ByVal Fontname As String, ByVal Charset As Long, ByVal Fontsize As Long)
 Dim ctl As Control

 For Each ctl In Frm.Controls
  If TypeOf ctl Is label Or _
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
 Forbiddenchars = "\/:*?<>|"""
 tStr = chkStr
 For i = 1 To Len(Forbiddenchars)
  tStr = Replace$(tStr, Mid$(Forbiddenchars, i, 1), ReplaceChar)
 Next i
 ReplaceForbiddenChars = tStr
End Function

Public Function IsForbiddenChars(chkStr As String) As Boolean
 Dim i As Long, Forbiddenchars As String
 IsForbiddenChars = False
 Forbiddenchars = "\/:*?<>|%"""
 For i = 1 To Len(Forbiddenchars)
  If InStr(chkStr, Mid$(Forbiddenchars, i, 1)) > 0 Then
   IsForbiddenChars = True
   Exit Function
  End If
 Next i
End Function

Public Sub RemoveX(Frm As Form)
 Dim hMenu As Long, nPosition As Long

 hMenu = GetSystemMenu(Frm.hwnd, 0)
 If hMenu <> 0 Then
  nPosition = GetMenuItemCount(hMenu)
  Call RemoveMenu(hMenu, nPosition - 1, MF_REMOVE Or MF_BYPOSITION)
  Call RemoveMenu(hMenu, nPosition - 2, MF_REMOVE Or MF_BYPOSITION)
  Call DrawMenuBar(Frm.hwnd)
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
 GetComputerName = Left$(tStr, InStr(tStr, Chr$(0)) - 1)
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
 Dim TempDesktop As String, reg As clsRegistry
 Set reg = New clsRegistry
 reg.hkey = HKEY_CURRENT_USER
 reg.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
 TempDesktop = reg.GetRegistryValue("Desktop")
 Set reg = Nothing
 If Trim$(TempDesktop) = "" Then
  TempDesktop = "C:\"
 End If
 GetDesktop = TempDesktop
End Function

Public Function GetMyFiles() As String
 Dim TempMyFiles, reg As clsRegistry
 Set reg = New clsRegistry
 reg.hkey = HKEY_CURRENT_USER
 reg.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
 TempMyFiles = reg.GetRegistryValue("Personal")
 Set reg = Nothing
 If Trim$(TempMyFiles) = "" Then
  TempMyFiles = "C:\"
 End If
 GetMyFiles = TempMyFiles
End Function

Public Function GetMyAppData() As String
 Dim TempMyAppData As String, reg As clsRegistry
 Set reg = New clsRegistry
 reg.hkey = HKEY_CURRENT_USER
 reg.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
 TempMyAppData = reg.GetRegistryValue("AppData")
 Set reg = Nothing
 If Trim$(TempMyAppData) = "" Then
  TempMyAppData = "C:\"
 End If
 GetMyAppData = TempMyAppData
End Function

Public Function GetMyLocalAppData() As String
 Dim TempMyLocalAppData As String, reg As clsRegistry
 Set reg = New clsRegistry
 reg.hkey = HKEY_CURRENT_USER
 reg.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
 TempMyLocalAppData = reg.GetRegistryValue("Local AppData")
 Set reg = Nothing
 If Trim$(TempMyLocalAppData) = "" Then
  TempMyLocalAppData = "C:\"
 End If
 GetMyLocalAppData = TempMyLocalAppData
End Function

Public Function GetFileAttributesStr(Filename As String) As String
 Dim hFind As Long, WFD As WIN32_FIND_DATA, attr As Long, AA As String
 hFind = FindFirstFile(Filename, WFD)
 attr = WFD.dwFileAttributes
 If attr And FILE_ATTRIBUTE_ARCHIVE Then AA = AA & "A"
 If attr And FILE_ATTRIBUTE_COMPRESSED Then AA = AA & "C"
 If attr And FILE_ATTRIBUTE_DIRECTORY Then AA = AA & "D"
 If attr And FILE_ATTRIBUTE_HIDDEN Then AA = AA & "H"
 If attr And FILE_ATTRIBUTE_NORMAL Then AA = AA & "N"
 If attr And FILE_ATTRIBUTE_READONLY Then AA = AA & "R"
 If attr And FILE_ATTRIBUTE_SYSTEM Then AA = AA & "S"
 GetFileAttributesStr = AA
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
 LoadDLL = LoadLibrary(DLLPath)
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
 Dim L As Long, Buffer As String, res As Long, drives As String, drv() As String, _
  i As Long
 L = 64: Buffer = Space(L)
 res = GetLogicalDriveStrings(L, Buffer)
 drives = Left$(Buffer, res)
 Set GetDrives = New Collection
 If Len(drives) > 0 Then
  If InStr(drives, Chr$(0)) > 0 Then
    drv = Split(drives, Chr$(0))
    For i = LBound(drv) To UBound(drv)
     If Trim$(drv(i)) <> "" Then
      GetDrives.Add drv(i)
     End If
    Next i
   Else
    GetDrives.Add drives
  End If
 End If
End Function

Public Function CompletePath(Path As String) As String
 If Len(Path) = 0 Then
  Exit Function
 End If
 Path = Trim$(Path)
 If Right$(Path, 1) = "\" Then
   CompletePath = LTrim$(Path)
  Else
   CompletePath = LTrim$(Path) & "\"
 End If
End Function

Public Function RunProgramWait(strCmdLine As String, Optional showWindow As Boolean = True) As Long
 Dim lngRetCode As Long, proc As PROCESS_INFORMATION, Start As STARTUPINFO, _
  secAttr As SECURITY_ATTRIBUTES, lngRet As Long, lngExit As Long

 Start.cb = Len(Start)
 If showWindow = False Then
  Start.dwFlags = STARTF_USESHOWWINDOW
  Start.wShowWindow = SW_HIDE
 End If

 lngRet = CreateProcessA(0&, strCmdLine, _
  0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, Start, proc)

 GetExitCodeProcess proc.hProcess, lngExit

 Do While lngExit = STILL_ACTIVE
  DoEvents
  Sleep 100
  GetExitCodeProcess proc.hProcess, lngExit
 Loop
 lngRet = CloseHandle(proc.hProcess)

 RunProgramWait = lngExit
End Function

Public Function EnterPasswords(ByRef UserPass As String, ByRef OwnerPass As String, f As Form) As Boolean
 If Options.PDFUserPass <> 0 Or Options.PDFOwnerPass <> 0 Then
   With f
    .Visible = False
    .fraUserPass.Enabled = Options.PDFUserPass
    .lblUserPass.Enabled = Options.PDFUserPass
    .lblUserPassRepeat.Enabled = Options.PDFUserPass
    .fraOwnerPass.Enabled = Options.PDFOwnerPass
    .lblOwnerPass.Enabled = Options.PDFOwnerPass
    .lblOwnerPassRepeat.Enabled = Options.PDFOwnerPass
    .iPasswords = Abs(Options.PDFUserPass) + Abs(Options.PDFOwnerPass * 2)
    .Show vbModal
    Do
     Sleep 100
     DoEvents
    Loop While .bFinished = False
   End With
   EnterPasswords = f.bSuccess
   UserPass = f.txtUserPass.Text
   OwnerPass = f.txtOwnerPass.Text
   Unload f
  Else
   EnterPasswords = False
   UserPass = ""
   OwnerPass = ""
 End If
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
 Dim i As Long, c As Long, nBytes() As Byte, nChar As String, nOK As Boolean, _
  nParentPath As String, nPos As Long, nRoot As String, nPath As String, _
  nPathParts() As String, nStart As Long

 Const kCharAsterisk = 42, kCharBackSlash = 92, kCharColon = 58, _
  kCharGreaterThan = 62, kCharLowerThan = 60, kCharPipe = 124, _
  kCharQuestion = 63, kCharQuote = 34, kCharSlash = 47

 nPath = Path
 PathUnquoteSpaces nPath
 nPath = Left$(nPath, InStrNullChar(nPath))
 If Not CBool(PathIsUNC(nPath) And Not TestUNCPaths) Then
  If CBool(PathFileExists(nPath)) Then
   IsValidPath = True
   Exit Function
  End If
 End If
 nRoot = nPath
 If PathStripToRoot(nRoot) Then
  nRoot = Left$(nRoot, InStrNullChar(nRoot))
  If PathIsUNC(nRoot) Then
    nPos = InStrRev(nRoot, "\")
    nRoot = Left$(nRoot, nPos - 1)
    If PathIsUNCServer(nRoot) Then
     nPath = Mid$(nPath, Len(nRoot) + 1)
    End If
   Else
    nPath = Mid$(nPath, Len(nRoot) + 1)
  End If
  If Len(nPath) Then
   nPathParts = Split(nPath, "\")
   If Len(nPathParts(0)) Then
     nStart = 0
    Else
     nStart = 1
   End If
   For i = nStart To UBound(nPathParts)
    nBytes = StrConv(nPathParts(i), vbFromUnicode)
    For c = 0 To UBound(nBytes)
     Select Case nBytes(c)
      Case Is < 32, Is > 255, kCharAsterisk, kCharBackSlash, kCharColon, kCharGreaterThan, kCharLowerThan, kCharPipe, kCharQuestion, kCharQuote, kCharSlash
       nOK = False
       Exit For
      Case Else
       nOK = True
     End Select
    Next c
    If Not nOK Then
     Exit For
    End If
   Next i
   IsValidPath = nOK
  End If
 End If
End Function

Public Sub FormInTaskbar(Form As Form, ByVal ShowInTaskBar As Boolean, Optional ByVal NoFlicker As Boolean)
 Dim nStyle As Long, nVisible As Boolean
 Const GWL_EXSTYLE = (-20), WS_EX_APPWINDOW = &H40000
 With Form
  nVisible = .Visible
  If NoFlicker And nVisible Then
   LockWindowUpdate GetDesktopWindow()
  End If
  .Visible = False
  nStyle = GetWindowLong(.hwnd, GWL_EXSTYLE)
  Select Case ShowInTaskBar
   Case True
    nStyle = nStyle Or WS_EX_APPWINDOW
   Case False
    nStyle = nStyle And Not WS_EX_APPWINDOW
  End Select
  SetWindowLong .hwnd, GWL_EXSTYLE, nStyle
  .Refresh
  .Visible = nVisible
  If NoFlicker And nVisible Then
   LockWindowUpdate 0&
  End If
 End With
End Sub

Public Function GetStringRessource(DLLExeFilename As String, ResID As Long) As String
 Dim hLib As Long, sBuf As String, lLen As Long

 hLib = LoadLibrary(DLLExeFilename)
 If hLib Then
  sBuf = Space$(255)
  lLen = LoadString(hLib, ResID, sBuf, Len(sBuf))
  If lLen Then
    sBuf = Left$(sBuf, lLen)
   Else
    sBuf = vbNullString
  End If
  FreeLibrary hLib
 End If

 GetStringRessource = sBuf
End Function


Private Sub ReplaceStrInFile(Filename As String, SearchStr As String, _
 ReplaceStr As String, Optional BufferSize As Long = 65536)
 Dim Content As String, iSearch As Long, pLen As Long, chunk As Long, fpos As Long, _
  fnI As Long, fnO As Long, LSearchStr As Long, LReplaceStr As Long, _
  Tempfile As String
 If Len(SearchStr) = 0 Or Len(ReplaceStr) = 0 Then
  Exit Sub
 End If
 fnI = FreeFile
 Open Filename For Binary Access Read As #fnI
 fnO = FreeFile
 Tempfile = GetTempFile(CompletePath(GetPDFCreatorTempfolder), "~RE")
 Open Tempfile For Output As #fnO
 LSearchStr = Len(SearchStr)
 LReplaceStr = Len(ReplaceStr)
 chunk = BufferSize
 fpos = 1: pLen = LOF(fnI)
 Do Until pLen = 0
  If LOF(fnI) - fpos + 1 < chunk Then
   chunk = LOF(fnI) - fpos + 1
  End If
  Content = Space$(chunk)
  Get #fnI, fpos, Content
  iSearch = InStrRev(Content, SearchStr)
  If iSearch = 0 Then
    pLen = chunk - LSearchStr + 1
    DoEvents
    Print #fnO, Left(Content, pLen);
    DoEvents
    fpos = fpos + pLen
   Else
    pLen = iSearch + LSearchStr - 1
    If pLen < chunk - LSearchStr + 1 Then
     pLen = chunk - LSearchStr + 1
    End If
    fpos = fpos + pLen
    Do Until iSearch = 0
     Content = Left(Content, iSearch - 1) & ReplaceStr & Mid(Content, iSearch + LSearchStr)
     iSearch = InStrRev(Content, SearchStr, iSearch)
     pLen = pLen + LReplaceStr - LSearchStr
    Loop
    DoEvents
    Print #fnO, Left(Content, pLen);
    DoEvents
  End If
 Loop
 Print #fnO, Right(Content, LSearchStr - 1);
 Close #fnO: Close #fnI
 Name Tempfile As Filename
 DoEvents
End Sub

Public Sub HTMLHelp_ShowTopic(Optional sTopicFile As String)
 Dim sHelpFile As String, hwndHelp As Long

 sHelpFile = App.Path & "\PDFCreator.chm"
 If sTopicFile = "" Then
   hwndHelp = HtmlHelp(0, sHelpFile, HH_DISPLAY_TOPIC, ByVal 0&)
  Else
   hwndHelp = HtmlHelpTopic(0, sHelpFile, HH_DISPLAY_TOPIC, sTopicFile)
 End If
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
 Dim res As Long
 If DLLHandle <> 0 Then
  res = FreeLibrary(DLLHandle)
  If res <> 0 Then
    UnLoadDLL = True
   Else
    UnLoadDLL = False
  End If
 End If
End Function

Public Function UnloadDLLComplete(ByRef DLLHandle As Long) As Long
 Dim c As Long
 If DLLHandle = 0 Then
  Exit Function
 End If
 c = 0
 Do While UnLoadDLL(DLLHandle)
  c = c + 1
  DoEvents
 Loop
 UnloadDLLComplete = c: DLLHandle = 0
End Function

Public Function ExecuteAndWait(hwnd As Long, Filename As String, Action As ShellAction) As Long
 Dim RetVal As Long, ShExInfo As SHELLEXECUTEINFO

 With ShExInfo
  .cbSize = Len(ShExInfo)
'  .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_FLAG_NO_UI Or SEE_MASK_CLASSNAME Or _
           SEE_MASK_FLAG_DDEWAIT
  .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI Or _
           SEE_MASK_FLAG_DDEWAIT
  .hwnd = hwnd
  Select Case Action
   Case ShellAction.APrint
    .lpVerb = "print"
   Case ShellAction.Aopen
    .lpVerb = "open"
   Case ShellAction.AExplore
    .lpVerb = "explore"
   End Select

  .lpFile = Filename
  .lpParameters = vbNullChar
  .lpDirectory = vbNullChar
  .nShow = SW_SHOWMINIMIZED
 End With

 ExecuteAndWait = ShellExecuteEx(ShExInfo)

 Do
  DoEvents
 Loop Until WaitForSingleObject(ShExInfo.hProcess, 0) <> WAIT_TIMEOUT
End Function

Public Function FileISPrintable(Filename As String) As Boolean
 Dim Ext As String, reg As clsRegistry
 FileISPrintable = False
 If Len(Filename) = 0 Then
  Exit Function
 End If
 SplitPath Filename, , , , , Ext
 If Len(Ext) = 0 Then
  Exit Function
 End If
 Set reg = New clsRegistry
 reg.hkey = HKEY_LOCAL_MACHINE
 reg.KeyRoot = "Software\CLASSES\." & Ext
 If reg.KeyExists = False Then
  Set reg = Nothing
  Exit Function
 End If
 reg.KeyRoot = "Software\CLASSES\" & reg.GetRegistryValue("") & "\shell\print"
 If reg.KeyExists = False Then
  Set reg = Nothing
  Exit Function
 End If
 FileISPrintable = True
End Function

Public Sub AddExplorerIntegration()
 Dim reg As clsRegistry, Keys As Collection, i As Long, sKey As String, Path As String
 Set reg = New clsRegistry
 reg.hkey = HKEY_CLASSES_ROOT
 Set Keys = reg.EnumRegistryKeys(HKEY_CLASSES_ROOT, "")
 For i = 1 To Keys.Count
  If Mid(Keys(i), 1, 1) = "." Then
   reg.KeyRoot = Keys(i)
   reg.Subkey = ""
   sKey = reg.GetRegistryValue("")
   If LenB(sKey) > 0 Then
    reg.KeyRoot = sKey
    If reg.KeyExists = True Then
     reg.Subkey = "shell\print\command"
     If reg.KeyExists = True Then
      If LenB(Trim$(reg.GetRegistryValue(""))) > 0 Then
       reg.KeyRoot = sKey & "\shell\" & Uninstall_GUID
       reg.Subkey = ""
       If reg.KeyExists = False Then
        Path = CompletePath(GetPDFCreatorApplicationPath)
        If Len(Path) > 1 Then
         reg.CreateKey
         reg.SetRegistryValue "", LanguageStrings.OptionsShellIntegrationCaption, REG_SZ
         reg.CreateKey "command"
         reg.KeyRoot = sKey & "\shell\" & Uninstall_GUID & "\command"
         reg.SetRegistryValue "", Path & "pdfcreator.exe -PF""%1"" -NS", REG_SZ
        End If
       End If
      End If
     End If
    End If
   End If
  End If
 Next i
 Set Keys = Nothing
 Set reg = Nothing
End Sub

Public Sub RemoveExplorerIntegration()
 Dim reg As clsRegistry, Keys As Collection, i As Long
 Set reg = New clsRegistry
 reg.hkey = HKEY_CLASSES_ROOT
 Set Keys = reg.EnumRegistryKeys(HKEY_CLASSES_ROOT, "")
 For i = 1 To Keys.Count
  reg.KeyRoot = Keys(i) & "\shell"
  reg.Subkey = Uninstall_GUID & "\command"
  If reg.KeyExists Then
   reg.DeleteKey reg.Subkey
  End If
  reg.Subkey = Uninstall_GUID
  If reg.KeyExists Then
   reg.DeleteKey reg.Subkey
  End If
 Next i
 Set Keys = Nothing
 Set reg = Nothing
End Sub

Public Function RemoveCompletePath(Path As String, Optional ShowDialog As Boolean) As Long
 Dim ret As Long, SHFILEOP As SHFILEOPSTRUCT, ShowD As Long, FOBuf() As Byte, _
  LenSHFileOp As Long

 LenSHFileOp = LenB(SHFILEOP)
 ReDim FOBuf(1 To LenSHFileOp)
 If ShowDialog Then
   ShowD = 0
  Else
   ShowD = FOF_SILENT
 End If
 
 With SHFILEOP
  .wFunc = FO_DELETE
  .pFrom = Path
  .pTo = vbNullString
  .fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR Or ShowD
 End With

 ret = SHFileOperation(SHFILEOP)
 DoEvents

 Call MoveMemory(FOBuf(1), SHFILEOP, LenSHFileOp)
 Call MoveMemory(FOBuf(21), FOBuf(19), 12)
 Call MoveMemory(SHFILEOP, FOBuf(1), LenSHFileOp)

 If ret <> 0 Then
  If SHFILEOP.fAnyOperationAborted <> 0 Then
   ret = -1
  End If
 End If
 RemoveCompletePath = ret
End Function

