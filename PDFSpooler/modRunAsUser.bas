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
'   01/13/2009 Frank Heindörfer
Public Function GetUserSessionToken(UserName As String, SessionID As Long, hToken As Long, Optional Logging As Boolean = False) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ProcessIDs As Collection, rc As Long, hProcess As Long, i As Long, process As clsProcess, AllActiveServices As Collection, _
  service As clsService
50030
50040  GetUserSessionToken = 1
50050  Set ProcessIDs = FindProcess(UserName, SessionID)
50060  If Logging = True Then
50070   WriteToSpecialLogfile "Count of ProcessIDs (UserName:" & UserName & ", SessionID:" & SessionID & ") = " & ProcessIDs.Count
50080  End If
50090
50100  Set AllActiveServices = EnumLocalServices(SERVICE_ACTIVE)
50110  If Logging = True Then
50120   WriteToSpecialLogfile "Count of local services: " & AllActiveServices.Count
50130   For i = 1 To AllActiveServices.Count
50140    Set service = AllActiveServices(i)
50150    WriteToSpecialLogfile "Services: " & service.ServiceName & _
    " [ControlsAccepted: " & service.ControlsAccepted & ", ServiceType: " & service.ServiceType & _
    ", CurrentState: " & service.CurrentState & ", ImagePath: " & service.ImagePath & "]"
50180   Next i
50190  End If
50200
50210  For i = ProcessIDs.Count To 1 Step -1
50220   Set process = ProcessIDs(i)
50230   If process.ID > 0 Then
50240    If IsService(process.Modulname, AllActiveServices) = False Then
50250      If LCase$(process.Modulname) <> "iexplore.exe" And LCase$(process.Modulname) <> "crome.exe" Then
50260        If Logging = True Then
50270         WriteToSpecialLogfile "Process (ProcessID = " & process.ID & ", Modulename = " & process.Modulname & ") seems not to be a service."
50280        End If
50290        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, False, process.ID)
50300        If Logging = True Then
50310         WriteToSpecialLogfile "hProcess (ProcessID = " & process.ID & ", Modulename = " & process.Modulname & "): " & hProcess
50320        End If
50330        If IsWinNT4 = True Then
50340          rc = OpenProcessToken(hProcess, TOKEN_ALL_ACCESS_NT4, hToken)
50350         Else
50360          rc = OpenProcessToken(hProcess, TOKEN_ALL_ACCESS, hToken)
50370        End If
50380        If Logging = True Then
50390         WriteToSpecialLogfile "rc (OpenProcessToken):" & rc
50400        End If
50410        CloseHandle (hProcess)
50420        If rc <> 0 Then
50430         GetUserSessionToken = 0
50440         Exit For
50450        End If
50460       Else
50470        If Logging = True Then
50480         WriteToSpecialLogfile "Ignore process: ProcessID = " & process.ID & ", Modulename = " & process.Modulname
50490        End If
50500      End If
50510     Else
50520      If Logging = True Then
50530       WriteToSpecialLogfile "Possible service found: ProcessID = " & process.ID & ", Modulename = " & process.Modulname
50540      End If
50550    End If
50560   End If
50570  Next i
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
50070   WriteToSpecialLogfile "GetUserLocalDirs ->  AppData:" & AppData
50080   If IsWinVistaPlus Then
50090    Dim LocalAppData As String
50100    LocalAppData = CompletePath(.GetRegistryValue("Local AppData"))
50110    LocalTemp = CompletePath(CompletePath(LocalAppData) & "Temp")
50120    WriteToSpecialLogfile "Vista: GetUserLocalDirs ->  LocalTemp:" & LocalTemp
50130    Exit Sub
50140   End If
50150   tStr = CompletePath(.GetRegistryValue("Local Settings"))
50160   WriteToSpecialLogfile "GetUserLocalDirs ->  Local Settings:" & tStr
50170  End With
50180
50190  Set reg = Nothing
50200  If LenB(tStr) = 0 Then
50210    If IsWin9xMe = True Then
50220      tStr = CompletePath(GetTempPathApi)
50230     Else
50240      If IsWinNT4 = True Then
50250       tStr = CompletePath(GetTempPathApi)
50260       If LenB(Environ$("Redmon_User")) > 0 Then
50270         tStr = tStr & Environ$("Redmon_User")
50280        Else
50290         tStr = tStr & GetUsername
50300       End If
50310      End If
50320    End If
50330   Else
50340    tStr = CompletePath(tStr) & "Temp\"
50350  End If
50360  LocalTemp = tStr
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
50040 Dim env As Long
50050 Dim lpEnv As Long
50060 env = 0
50070 lpEnv = 0
50080 lpEnv = VarPtr(env)
50090 Result = CreateEnvironmentBlock(lpEnv, hToken, 0)
50100 If Result = 0 Then
50110   WriteToSpecialLogfile "CreateEnvironmentBlock " & "fehlgeschlagen: " & Err.LastDllError
50120 Else
50130   WriteToSpecialLogfile "CreateEnvironmentBlock=" & "erfolgreich (" & lpEnv & "->" & env & ")"
50140 End If
50150
50160
50170  Buffer = Space(255)
50180
50190  strDesktop = "WinSta0\Default"
50200  si.lpDesktop = strDesktop
50210  si.cb = Len(si)
50220  WriteToSpecialLogfile "Application=" & Application
50230  WriteToSpecialLogfile "Parameters=" & Parameters
50240  Result = CreateProcessAsUser(hToken, Application, Parameters, 0&, 0&, False, _
  CREATE_DEFAULT_ERROR_MODE Or CREATE_UNICODE_ENVIRONMENT, lpEnv, CurrentDirectory, si, PI)
50260
50270  If Result = 0 Then
50280   RunAsUser = Err.LastDllError
50290   FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, Err.LastDllError, LANG_NEUTRAL, Buffer, 200, ByVal 0&
50300   WriteToSpecialLogfile "CreateProcessAsUser() failed with error " & Err.LastDllError & " = " & Buffer
50310   CloseHandle hToken
50320   Exit Function
50330  End If
50340
50350  WriteToSpecialLogfile "CreateProcessAsUser() = successful"
50360
50370  CloseHandle PI.hThread
50380  CloseHandle PI.hProcess
50390  RunAsUser = 0
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
