Attribute VB_Name = "modRunAsUser"
Option Explicit

Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function LoadUserProfile Lib "userenv.dll" _
 Alias "LoadUserProfileA" _
 (ByVal hToken As Long, _
  ByVal lpProfileInfo As Long _
 ) As Boolean
    
Private Declare Function UnloadUserProfile Lib "userenv.dll" (ByVal hToken As Long, ByVal hProfile As Long) As Long

Private Declare Function CreateProcessAsUser Lib "advapi32.dll" _
 Alias "CreateProcessAsUserA" _
 (ByVal hToken As Long, _
  ByVal lpApplicationName As String, _
  ByVal lpCommandLine As String, _
  ByVal lpProcessAttributes As Long, _
  ByVal lpThreadAttributes As Long, _
  ByVal bInheritHandles As Long, _
  ByVal dwCreationFlags As Long, _
  ByVal lpEnvironment As Long, _
  ByVal lpCurrentDirectory As String, _
  lpStartupInfo As STARTUPINFO, _
  lpProcessInformation As PROCESS_INFORMATION _
  ) As Long
        
Private Declare Function WTSEnumerateProcesses Lib "wtsapi32.dll" _
 Alias "WTSEnumerateProcessesA" _
 (ByVal hServer As Long, _
  ByVal Reserved As Long, _
  ByVal Version As Long, _
  ByRef ppProcessInfo As Long, _
  ByRef pCount As Long _
 ) As Long

 Private Declare Function LookupAccountSid Lib "advapi32.dll" _
  Alias "LookupAccountSidA" _
  (ByVal lpSystemName As String, _
   ByVal Sid As Long, _
   ByVal Name As String, _
   cbName As Long, _
   ByVal ReferencedDomainName As String, _
   cbReferencedDomainName As Long, _
   peUse As Long _
  ) As Long
      
Private Declare Sub WTSFreeMemory Lib "wtsapi32.dll" (ByVal pMemory As Long)

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" _
(Destination As Any, Source As Any, ByVal Length As Long)
        
Private Const PROCESS_QUERY_INFORMATION As Long = (&H400)
Private Const PROCESS_VM_READ As Long = (&H10)
Private Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Private Const TOKEN_ASSIGN_PRIMARY As Long = &H1
Private Const TOKEN_DUPLICATE As Long = &H2
Private Const TOKEN_IMPERSONATE As Long = &H4
Private Const TOKEN_QUERY As Long = &H8
Private Const TOKEN_QUERY_SOURCE As Long = &H10
Private Const TOKEN_ADJUST_GROUPS As Long = &H40
Private Const TOKEN_ADJUST_PRIVILEGES As Long = &H20
Private Const TOKEN_ADJUST_SESSIONID As Long = &H100
Private Const TOKEN_ADJUST_DEFAULT As Long = &H80
Private Const TOKEN_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or TOKEN_ASSIGN_PRIMARY Or TOKEN_DUPLICATE Or TOKEN_IMPERSONATE Or TOKEN_QUERY Or TOKEN_QUERY_SOURCE Or TOKEN_ADJUST_PRIVILEGES Or TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_SESSIONID Or TOKEN_ADJUST_DEFAULT)
Private Const TOKEN_ALL_ACCESS_NT4 As Long = (STANDARD_RIGHTS_REQUIRED Or TOKEN_ASSIGN_PRIMARY Or TOKEN_DUPLICATE Or TOKEN_IMPERSONATE Or TOKEN_QUERY Or TOKEN_QUERY_SOURCE Or TOKEN_ADJUST_PRIVILEGES Or TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_DEFAULT)
Private Const CREATE_DEFAULT_ERROR_MODE As Long = &H4000000
Private Const SW_SHOW As Long = 5

' Terminal Server (WTSAPI32) constants
Public Const WTS_CURRENT_SERVER = 0&
Public Const WTS_CURRENT_SERVER_HANDLE = 0&
Public Const WTS_CURRENT_SERVER_NAME = vbNullString

Private Type SECURITY_ATTRIBUTES
 nLength As Long
 lpSecurityDescriptor As Long
 bInheritHandle As Long
End Type

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

Private Const PI_NOUI = 1
Private Const PI_APPLYPOLICY = 2

Private Type PROFILEINFO
 dwSize As Long
 dwFlags As Long
 lpUsername As Long
 lpProfilePath As Long
 lpDefaultPath As Long
 lpServerName As Long
 lpPolicyPath As Long
 hProfile As Long
End Type

Private Type WTS_PROCESS_INFO
 SessionID As Long
 ProcessID As Long
 pProcessName As Long
 pUserSid As Long
End Type

Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const LANG_NEUTRAL = &H0
Private Const SUBLANG_DEFAULT = &H1
Private Declare Function GetLastError Lib "Kernel32" () As Long
Private Declare Function FormatMessage Lib "Kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
   
' Function FindProcess
' Description:
'   Looks for a user process within the given session on a terminal server
' Inputs:
'   UserName: name of the printing user
'   SessionID: identifier of the session from where printing was invoked
' Output:
'   If found, then the pid, otherwise 0
' Last modification:
'   09/07/2004 Gergely Matefi
Private Function FindProcess(UserName As String, SessionID As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim RetVal As Long, Count As Long, i As Integer, lpBuffer As Long, _
  p As Long, udtProcessInfo As WTS_PROCESS_INFO, isMissing As Boolean, _
  ProcessID As Long
50040
50050  Dim bSuccess As Long       ' Status variable
50060  Dim pOwner As Long         ' Pointer to the Owner's SID
50070  Dim Name As String         ' Name of the file owner
50080  Dim domain_name As String  ' Name of the first domain for the owner
50090  Dim name_len As Long       ' Required length for the owner name
50100  Dim domain_len As Long     ' Required length for the domain name
50110  Dim deUse As Long          ' Pointer to a SID_NAME_USE enumerated
50120  Dim ProcessOwner As String
50130
50140  ProcessID = 0
50150  IfLoggingWriteLogfile "Process enumeration..."
50160  RetVal = WTSEnumerateProcesses(WTS_CURRENT_SERVER_HANDLE, 0&, 1, lpBuffer, Count)
50170  If RetVal <> 0 Then ' WTSEnumerateProcesses was successful
50180    p = lpBuffer
50190    i = 1
50200    isMissing = True
50210    While i <= Count And isMissing = True
50220     CopyMemory udtProcessInfo, ByVal p, LenB(udtProcessInfo)
50230     'IfLoggingWriteLogfile "Process " & udtProcessInfo.SessionID & " " & udtProcessInfo.ProcessID
50240     If udtProcessInfo.SessionID = SessionID And udtProcessInfo.ProcessID > 0 Then
50250      ' Retrieve the name of the account and the name of the first
50260      ' domain on which this SID is found.  Passes in the Owner's SID
50270      ' obtained previously.  Call LookupAccountSid twice, the first time
50280      ' to obtain the required size of the owner and domain names.
50290      Name = ""
50300      domain_name = ""
50310      name_len = 0
50320      domain_len = 0
50330      bSuccess = LookupAccountSid(vbNullString, udtProcessInfo.pUserSid, Name, _
      name_len, domain_name, domain_len, deUse)
50350      If name_len > 0 And domain_len > 0 Then
50360        '  Allocate the required space in the name and domain_name string
50370        '  variables. Allocate 1 byte less to avoid the appended NULL character.
50380        Name = Space(name_len - 1)
50390        domain_name = Space(domain_len - 1)
50400        bSuccess = LookupAccountSid(vbNullString, udtProcessInfo.pUserSid, Name, _
        name_len, domain_name, domain_len, deUse)
50420        If bSuccess <> 0 Then
50430          If Name = UserName Then
50440           ProcessID = udtProcessInfo.ProcessID
50450           isMissing = False
50460           IfLoggingWriteLogfile "Process found=" & ProcessID & " " & domain_name & " " & Name
50470          End If
50480         Else
50490          IfLoggingWriteLogfile "Can't look up account!"
50500        End If
50510       Else
50520        IfLoggingWriteLogfile "Invalid name length for LookupAccountSid!"
50530      End If
50540     End If
50550     i = i + 1
50560     p = p + LenB(udtProcessInfo)
50570    Wend
50580    WTSFreeMemory lpBuffer   'Free your memory buffer
50590   Else
50600    ' Error occurred calling WTSEnumerateProcesses
50610    IfLoggingWriteLogfile "Error occurred calling WTSEnumerateProcesses.  " & Err.LastDllError
50620  End If
50630  FindProcess = ProcessID
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modRunAsUser", "FindProcess")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

' Function FindExplorerProcess
' Description:
'   Looks for the processid of the explorer on a workstation
' Inputs:
'   -
' Output:
'   If found, then the pid, otherwise 0
' Last modification:
'   09/07/2004 Gergely Matefi
Private Function FindExplorerProcess() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Explorer As Long, rc As Long, ProcessID As Long
50020  Explorer = FindWindow("progman", vbNullString) 'progman
50030  If Explorer = 0 Then
50040    ' Explorer is not running as a shell
50050    IfLoggingWriteLogfile "Cannot find progman. Maybe the user has logged out."
50060    FindExplorerProcess = 0
50070   Else
50080    ' Explorer is running as the shell
50090    ' Get the Process ID
50100    rc = GetWindowThreadProcessId(Explorer, ProcessID)
50110    ' Get the handle to the Process ID
50120    FindExplorerProcess = ProcessID
50130  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modRunAsUser", "FindExplorerProcess")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

' Function GetUserSessionToken
' Description:
'   Looks for a security token for user UserName in session SessionID
' Inputs:
'   UserName - name of the printing user
'   SessionID - ID of the session from where printing was invoked
' Output:
'   hToken - handle of the found token
'   Returns 0 if successful otherwise 1
' Last modification:
'   09/07/2004 Gergely Matefi
Public Function GetUserSessionToken(UserName As String, SessionID As Long, hToken As Long) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ProcessID As Long, rc As Long, hProcess As Long
50020
50030  ' Check if we are in a terminal box
50040  If IsTerminalServer = True Then
50050    IfLoggingWriteLogfile "Terminal token is asked for " & UserName & " " & SessionID
50060    ProcessID = FindProcess(UserName, SessionID)
50070   Else
50080    IfLoggingWriteLogfile "Console token is asked for " & UserName
50090    ProcessID = FindExplorerProcess()
50100  End If
50110
50120  If ProcessID > 0 Then
50130    IfLoggingWriteLogfile "Process found, pid= " & ProcessID
50140    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, False, ProcessID)
50150    If IsWinNT4 = True Then
50160      rc = OpenProcessToken(hProcess, TOKEN_ALL_ACCESS_NT4, hToken)
50170     Else
50180      rc = OpenProcessToken(hProcess, TOKEN_ALL_ACCESS, hToken)
50190    End If
50200    CloseHandle (hProcess)
50210    GetUserSessionToken = 0
50220   Else
50230    GetUserSessionToken = 1
50240  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modRunAsUser", "GetUserSessionToken")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

' Sub CloseToken
' Description:
'   Closes an open security token
' Inputs:
'   hToken - handle of the token
' Output:
'   -
' Last modification:
'   09/07/2004 Gergely Matefi
Public Sub CloseToken(hToken As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  CloseHandle hToken
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modRunAsUser", "CloseToken")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

' Function LoadProfile
' Description:
'   Load user profile into registry
' Inputs:
'   UserName - name of the user
'   hToken - security token handle
' Output:
'   hProfile - handle of the loaded profile
'   Returns 0 if successful otherwise 1
' Last modification:
'   09/07/2004 Gergely Matefi
Public Function LoadProfile(sUsername As String, hToken As Long, hProfile As Long) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim PI As PROFILEINFO, lpPI As Long, res As Long
50020  With PI
50030   .dwSize = Len(PI)
50040   .dwFlags = PI_NOUI ' Or PI_APPLYPOLICY
50050   .dwFlags = 0
50060   .lpUsername = StrPtr(sUsername)
50070   .lpProfilePath = 0
50080   .lpDefaultPath = 0
50090   .lpServerName = 0
50100   .lpPolicyPath = 0
50110
50120   lpPI = VarPtr(PI)
50130   res = LoadUserProfile(hToken, lpPI)
50140   If res <> 0 Then
50150     hProfile = PI.hProfile
50160     LoadProfile = 0
50170    Else
50180     LoadProfile = Err.LastDllError
50190     IfLoggingWriteLogfile "LoadProfile failed, error=" & Err.LastDllError
50200   End If
50210  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modRunAsUser", "LoadProfile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

' Sub UnloadProfile
' Description:
'   Unload user profile from registry
' Inputs:
'   hToken - security token handle
'   hProfile - handle of the loaded profile
' Output:
'   -
' Last modification:
'   09/07/2004 Gergely Matefi
Public Sub UnloadProfile(ByVal hToken As Long, ByVal hProfile As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim rc As Long
50020  rc = UnloadUserProfile(hToken, hProfile)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modRunAsUser", "UnloadProfile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

' Sub GetUserLocalDirs
' Description:
'   Gets the local application and local temp folders of the user
' Inputs:
'   hProfile - handle of the user profile
' Output:
'   LocalAppData - local application directory absolute path
'   LocalTemp - local temp directory absolute path
' Last modification:
'   09/07/2004 Gergely Matefi
Public Sub GetUserLocalDirs(hProfile As Long, AppData As String, LocalTemp As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tmp As String, reg As clsRegistry
50020  Set reg = New clsRegistry
50030  With reg
50040   .hkey = hProfile
50050   .KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
50060   AppData = .GetRegistryValue("AppData") & "\"
50070   LocalTemp = .GetRegistryValue("Local Settings") & "\Temp\"
50080  End With
50090  Set reg = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modRunAsUser", "GetUserLocalDirs")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

' Function RunAsUser
' Description:
'   Starts Application with cmd line Parameters in the current Directory
'   within SessionId session under UserName account
' Inputs:
'   hToken - handle of the security token
'   Application - application path
'   Parameters - application cmd line params
'   CurrentDirectory - directory where the application will be started
' Output:
'   Returns 0 if successful otherwise the error code
' Last modification:
'   09/07/2004 Gergely Matefi
Public Function RunAsUser(hToken As Long, ByVal Application As String, ByVal Parameters As String, _
                ByVal CurrentDirectory As String) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Result As Long, si As STARTUPINFO, PI As PROCESS_INFORMATION, _
  strDesktop As String, Buffer As String
50030
50040  Buffer = Space(255)
50050
50060  strDesktop = "WinSta0\Default"
50070  si.lpDesktop = strDesktop
50080  si.cb = Len(si)
50090  WriteToSpecialLogfile "Application=" & Application
50100  WriteToSpecialLogfile "Parameters=" & Parameters
50110  Result = CreateProcessAsUser(hToken, Application, Parameters, 0&, 0&, False, _
  CREATE_DEFAULT_ERROR_MODE, 0&, CurrentDirectory, si, PI)
50130
50140  If Result = 0 Then
50150   RunAsUser = Err.LastDllError
50160   FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, Err.LastDllError, LANG_NEUTRAL, Buffer, 200, ByVal 0&
50170   IfLoggingWriteLogfile "CreateProcessAsUser() failed with error " & Err.LastDllError
50180   WriteToSpecialLogfile "CreateProcessAsUser() failed with error " & Err.LastDllError & " = " & Buffer
50190   CloseHandle hToken
50200   Exit Function
50210  End If
50220
50230  CloseHandle PI.hThread
50240  CloseHandle PI.hProcess
50250  RunAsUser = 0
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modRunAsUser", "RunAsUser")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
