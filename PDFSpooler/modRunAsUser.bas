Attribute VB_Name = "modRunAsUser"
Option Explicit

Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function LoadUserProfile Lib "userenv.dll" Alias _
    "LoadUserProfileA" (ByVal hToken As Long, _
    ByVal lpProfileInfo As Long) As Boolean
    
'Public Declare Function UnloadUserProfile Lib "userenv.dll" (ByVal hToken As Long, ByVal hProfile As Long) As Long
'ToDo: Unload Registry Profile

Private Declare Function CreateProcessAsUser Lib "advapi32.dll" _
        Alias "CreateProcessAsUserA" _
        (ByVal hToken As Long, _
        ByVal lpApplicationName As Long, _
        ByVal lpCommandLine As String, _
        ByVal lpProcessAttributes As Long, _
        ByVal lpThreadAttributes As Long, _
        ByVal bInheritHandles As Long, _
        ByVal dwCreationFlags As Long, _
        ByVal lpEnvironment As Long, _
        ByVal lpCurrentDirectory As String, _
        lpStartupInfo As STARTUPINFO, _
        lpProcessInformation As PROCESS_INFORMATION) As Long
        
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
    lpUserName As Long
    lpProfilePath As Long
    lpDefaultPath As Long
    lpServerName As Long
    lpPolicyPath As Long
    hProfile As Long
End Type

' This function gets a User Token
'
Private Function GetUserToken(UserName As String) As Long
 Dim Explorer As Long, _
     pid As Long, rc As Long, hProcess As Long, hToken As Long

 Explorer = FindWindow("progman", vbNullString) 'progman

 If Explorer = 0 Then
   ' Explorer is not running as a shell
   MsgBox "The user has logged out."
  Else
   ' Explorer is running as the shell
   ' Get the Process ID

   rc = GetWindowThreadProcessId(Explorer, pid)
   ' Get the handle to the Process ID

   hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, False, pid)
   ' Get the Impersonation Token for the handle
   Call ReadVersionInfo
   If WinNT4 = True Then
     rc = OpenProcessToken(hProcess, TOKEN_ALL_ACCESS_NT4, hToken)
    Else
     rc = OpenProcessToken(hProcess, TOKEN_ALL_ACCESS, hToken)
   End If
   ' Close the process

   CloseHandle (hProcess)
   'MsgBox "PID: " & pid & " Token: " & hToken
   GetUserToken = hToken
 End If
End Function

Public Function RunAsUser(ByVal CommandLine As String, _
                ByVal CurrentDirectory As String, _
                UserName As String) As Long
Dim Result As Long
Dim hToken As Long
Dim hProfile As Long
Dim si As STARTUPINFO
Dim PI As PROCESS_INFORMATION

Dim strDesktop As String

    hToken = GetUserToken(UserName)
    If hToken = 0 Then
        RunAsUser = Err.LastDllError
        ' LogonUser will fail with 1314 error code, if the user account associated
        ' with the calling security context does not have
        ' "Act as part of the operating system" permission
        MsgBox "Getting User Token failed with error " & Err.LastDllError, vbExclamation
        Exit Function
    End If

    strDesktop = "WinSta0\Default"
    si.lpDesktop = strDesktop
    si.cb = Len(si)

    'LoadProfile UserName, hToken, hProfile

    Result = CreateProcessAsUser(hToken, 0&, CommandLine, 0&, 0&, False, _
                CREATE_DEFAULT_ERROR_MODE, 0&, CurrentDirectory, si, PI)
                
    If Result = 0 Then
        RunAsUser = Err.LastDllError
        ' CreateProcessAsUser will fail with 1314 error code, if the user
        ' account associated with the calling security context does not have
        ' the following two permissions
        ' "Replace a process level token"
        ' "Increase Quotoas"
        MsgBox "CreateProcessAsUser() failed with error " & Err.LastDllError, vbExclamation
        CloseHandle hToken
        Exit Function
    End If

    CloseHandle hToken
    CloseHandle PI.hThread
    CloseHandle PI.hProcess
    RunAsUser = 0
End Function

Public Function LoadprofileUser(UserName As String) As Long
Dim Result As Long, hToken As Long, hProfile As Long

 hToken = GetUserToken(UserName)
 If hToken = 0 Then
  LoadprofileUser = Err.LastDllError
  ' LogonUser will fail with 1314 error code, if the user account associated
  ' with the calling security context does not have
  ' "Act as part of the operating system" permission
  MsgBox "Getting User Token failed with error " & Err.LastDllError, vbExclamation
  Exit Function
 End If

 Call LoadProfile(UserName, hToken, hProfile)
 CloseHandle hToken
 LoadprofileUser = 0
End Function

Public Function LoadProfile(sUsername As String, hToken As Long, _
    hProfile As Long) As Long

    Dim PI As PROFILEINFO
    Dim lpPI As Long, res As Long

    PI.dwSize = Len(PI)
    PI.dwFlags = PI_NOUI ' Or PI_APPLYPOLICY
    PI.dwFlags = 0
    PI.lpUserName = StrPtr(sUsername)
    PI.lpProfilePath = 0
    PI.lpDefaultPath = 0
    PI.lpServerName = 0
    PI.lpPolicyPath = 0

    lpPI = VarPtr(PI)
    res = LoadUserProfile(hToken, lpPI)
    If res <> 0 Then
      hProfile = PI.hProfile
      LoadProfile = 0
     Else
      LoadProfile = Err.LastDllError
    End If
End Function
