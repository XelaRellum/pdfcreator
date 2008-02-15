Attribute VB_Name = "modRunAsUser"
Option Explicit

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
'   11/20/2004 Frank Heindörfer
Public Function GetUserSessionToken(UserName As String, SessionID As Long, hToken As Long, Optional Logging As Boolean = False) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ProcessID As Long, rc As Long, hProcess As Long
50020
50030  ProcessID = FindProcess(UserName, SessionID)
50040  If Logging = True Then
50050   WriteToSpecialLogfile "ProcessID (UserName:" & UserName & _
   ", SessionID:" & SessionID & ") = " & ProcessID
50070  End If
50080
50090  If ProcessID > 0 Then
50100    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, False, ProcessID)
50110    If IsWinNT4 = True Then
50120      rc = OpenProcessToken(hProcess, TOKEN_ALL_ACCESS_NT4, hToken)
50130     Else
50140      rc = OpenProcessToken(hProcess, TOKEN_ALL_ACCESS, hToken)
50150    End If
50160    CloseHandle (hProcess)
50170    GetUserSessionToken = 0
50180   Else
50190    GetUserSessionToken = 1
50200  End If
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
50060   .lpUserName = StrPtr(sUsername)
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
50190 '    IfLoggingWriteLogfile "LoadProfile failed, error=" & Err.LastDllError
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
'   08/31/2005 Frank Heindörfer: Support für Win9x/WinNT added
Public Sub GetUserLocalDirs(hProfile As Long, AppData As String, LocalTemp As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tStr As String, reg As clsRegistry
50020  Set reg = New clsRegistry
50030  With reg
50040   .hkey = hProfile
50050   .KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
50060   AppData = CompletePath(.GetRegistryValue("AppData"))
50070   tStr = CompletePath(.GetRegistryValue("Local Settings"))
50080  End With
50090  Set reg = Nothing
50100  If LenB(tStr) = 0 Then
50110    If IsWin9xMe = True Then
50120      tStr = CompletePath(GetTempPathApi)
50130     Else
50140      If IsWinNT4 = True Then
50150       tStr = CompletePath(GetTempPathApi)
50160       If LenB(Environ$("Redmon_User")) > 0 Then
50170         tStr = tStr & Environ$("Redmon_User")
50180        Else
50190         tStr = tStr & GetUsername
50200       End If
50210      End If
50220    End If
50230   Else
50240    tStr = CompletePath(tStr) & "Temp\"
50250  End If
50260  LocalTemp = tStr
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
  
50020 Dim env As Long
50021 Dim lpEnv As Long
50022 env = 0
50023 lpEnv = 0
50024 lpEnv = VarPtr(env)
50025 Result = CreateEnvironmentBlock(lpEnv, hToken, 0)
50026 If Result = 0 Then
50027   WriteToSpecialLogfile "CreateEnvironmentBlock " & "fehlgeschlagen: " & Err.LastDllError
50028 Else
50029   WriteToSpecialLogfile "CreateEnvironmentBlock=" & "erfolgreich (" & lpEnv & "->" & env & ")"
50023 End If
  
50030
50040  Buffer = Space(255)
50050
50060  strDesktop = "WinSta0\Default"
50070  si.lpDesktop = strDesktop
50080  si.cb = Len(si)
50090  WriteToSpecialLogfile "Application=" & Application
50100  WriteToSpecialLogfile "Parameters=" & Parameters
50110  Result = CreateProcessAsUser(hToken, Application, Parameters, 0&, 0&, False, _
  CREATE_DEFAULT_ERROR_MODE Or CREATE_UNICODE_ENVIRONMENT, lpEnv, CurrentDirectory, si, PI)
50130
50140  If Result = 0 Then
50150   RunAsUser = Err.LastDllError
50160   FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, Err.LastDllError, LANG_NEUTRAL, Buffer, 200, ByVal 0&
50170   WriteToSpecialLogfile "CreateProcessAsUser() failed with error " & Err.LastDllError & " = " & Buffer
50180   CloseHandle hToken
50190   Exit Function
50200  End If
50200
50201 WriteToSpecialLogfile "CreateProcessAsUser() successful = "
50202 Result = WaitForSingleObject(PI.hProcess, 60000)
50203 WriteToSpecialLogfile "WaitForSingleObject = " & Result
50204 DestroyEnvironmentBlock lpEnv
50205
50210
50220  CloseHandle PI.hThread
50230  CloseHandle PI.hProcess
50240  RunAsUser = 0
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
