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

Public Sub SplitPath(FullPath As String, Optional Drive As String, Optional Path As String, Optional FileName As String, Optional File As String, Optional Extension As String)
 Dim nPos As Integer
 nPos = InStrRev(FullPath, "\")
 If nPos > 0 Then
   If Left$(FullPath, 2) = "\\" Then
    If nPos = 2 Then
     Drive = FullPath: Path = vbNullString: FileName = vbNullString: File = vbNullString
     Extension = vbNullString
     Exit Sub
    End If
   End If
   Path = Left$(FullPath, nPos - 1)
   FileName = Mid$(FullPath, nPos + 1)
   nPos = InStrRev(FileName, ".")
   If nPos > 0 Then
     File = Left$(FileName, nPos - 1)
     Extension = Mid$(FileName, nPos + 1)
    Else
     File = FileName
     Extension = vbNullString
   End If
  Else
   nPos = InStrRev(FullPath, ":")
   If nPos > 0 Then
     Path = Mid(FullPath, 1, nPos - 1): FileName = Mid(FullPath, nPos + 1)
     nPos = InStrRev(FileName, ".")
     If nPos > 0 Then
       File = Left$(FileName, nPos - 1)
       Extension = Mid$(FileName, nPos + 1)
      Else
       File = FileName
       Extension = vbNullString
     End If
    Else
     Path = vbNullString: FileName = FullPath
     nPos = InStrRev(FileName, ".")
     If nPos > 0 Then
       File = Left$(FileName, nPos - 1)
       Extension = Mid$(FileName, nPos + 1)
      Else
       File = FileName
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
 Dim res As Long, Tempdir As String, r As clsRegistry

 Set r = New clsRegistry
 r.hkey = HKEY_CURRENT_USER
 r.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
 Tempdir = Trim$(r.GetRegistryValue("Local Settings"))
 If Tempdir <> "" Then
   If Right$(Tempdir, 1) = "\" Then
     Tempdir = Tempdir & "Temp\"
    Else
     Tempdir = Tempdir & "\Temp\"
   End If
   If Len(Dir(Tempdir, vbDirectory)) = 0 Then
    MakePath Tempdir
   End If
   If Dir(Tempdir, vbDirectory) = "" Then
    'Can't create tempdir
    Tempdir = Space$(MAX_PATH)
    res = GetTempPathA(MAX_PATH, Tempdir)
    If res > 0 Then
      GetTempPath = Left$(Tempdir, res)
     Else
      GetTempPath = "C:\"
    End If
   End If
   GetTempPath = Tempdir
  Else
   Tempdir = Space$(MAX_PATH)
   res = GetTempPathA(MAX_PATH, Tempdir)
   If res > 0 Then
     GetTempPath = Left$(Tempdir, res)
    Else
     GetTempPath = "C:\"
   End If
 End If
 Set r = Nothing
End Function

Public Function GetTempFile(Optional ByVal Path As String, Optional Prefix As String) As String
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

Public Sub CombineFiles(ByVal FileName As String, Files As Collection, stb As StatusBar)
 Dim i As Long, fnSource As Long, fnDest As Long, sBuffer As String, _
  aLen As Double, tLen As Double

 FileName = Trim$(FileName)
 If FileName = vbNullString Or Files.Count = 0 Or Right$(FileName, 1) = "\" Then
  Exit Sub
 End If
 If Len(Dir(FileName)) > 0 Then
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
 Open FileName For Binary As #fnDest
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
 Forbiddenchars = "\/:*?<>|"""
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

Public Function GetMyAppData() As String
 Dim TempMyAppData As String, r As clsRegistry

 Set r = New clsRegistry
 r.hkey = HKEY_CURRENT_USER
 r.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
 TempMyAppData = Trim$(r.GetRegistryValue("AppData"))
 If Right$(TempMyAppData, 1) <> "\" Then
  GetMyAppData = TempMyAppData & "\"
 End If
 Set r = Nothing
 If TempMyAppData = "" Then
  GetMyAppData = GetSpecialFolder(ssfAPPDATA)
  If Trim$(GetMyAppData) = "" Then
   GetMyAppData = "C:\"
  End If
 End If
End Function

Public Function GetFileAttributesStr(FileName As String) As String
 Dim hFind As Long, wFD As WIN32_FIND_DATA, Attr As Long, AA As String
 hFind = FindFirstFile(FileName, wFD)
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

Public Function CheckPath(PathOrFile As String) As Boolean
 On Local Error GoTo ErrorHandler
 CheckPath = True
 Dir PathOrFile
 Exit Function
ErrorHandler:
 CheckPath = False
End Function

Public Function LoadDLL(DLLPath As String) As Long
 LoadDLL = LoadLibrary(DLLPath)
End Function

Public Sub UnLoadDLL(DllHandle As Long)
 If DllHandle <> 0 Then
  FreeLibrary DllHandle
 End If
End Sub

Public Function MakePath(ByVal Verz As String) As Boolean
 Dim Success As Boolean, dummy As String, Entry As String

 Err = 0: Success = True: dummy = "": Entry = Verz
 On Local Error Resume Next
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
 On Local Error GoTo 0

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
 Path = Trim$(Path)
 If Right$(Path, 1) = "\" Then
   CompletePath = Path
  Else
   CompletePath = Path & "\"
 End If
End Function

Public Function RunProgramWait(strCmdLine As String, Optional showWindow As Boolean = True) As Long
 Dim lngRetCode As Long, proc As PROCESS_INFORMATION, start As STARTUPINFO, _
  secAttr As SECURITY_ATTRIBUTES, lngRet As Long, lngExit As Long

 start.cb = Len(start)
 If showWindow = False Then
  start.dwFlags = STARTF_USESHOWWINDOW
  start.wShowWindow = SW_HIDE
 End If

 lngRet = CreateProcessA(0&, strCmdLine, _
  0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)

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
 If Options.PDFUserPass = True Or Options.PDFOwnerPass = True Then
   With f
    .Visible = False
    .fraUserPass.Enabled = Options.PDFUserPass
    .lblUserPass.Enabled = Options.PDFUserPass
    .lblUserPassRepeat.Enabled = Options.PDFUserPass
    .fraOwnerPass.Enabled = Options.PDFOwnerPass
    .lblOwnerPass.Enabled = Options.PDFUserPass
    .lblOwnerPassRepeat.Enabled = Options.PDFUserPass
    .iPasswords = Abs(Options.PDFUserPass) + Abs(Options.PDFOwnerPass * 2)
    If Options.PDFUserPass = False Then
     .txtOwnerPass.SetFocus
    End If
    .Show
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
 On Local Error GoTo ErrorHandler:
 Dim tStr As String
 IsGoodDrive = False
 tStr = Dir(Mid$(tDrive, 1, 2) & "\", vbDirectory)
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
  nStyle = GetWindowLong(.hWnd, GWL_EXSTYLE)
  Select Case ShowInTaskBar
   Case True
    nStyle = nStyle Or WS_EX_APPWINDOW
   Case False
    nStyle = nStyle And Not WS_EX_APPWINDOW
  End Select
  SetWindowLong .hWnd, GWL_EXSTYLE, nStyle
  .Refresh
  .Visible = nVisible
  If NoFlicker And nVisible Then
   LockWindowUpdate 0&
  End If
 End With
End Sub
