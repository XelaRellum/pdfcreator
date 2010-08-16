Attribute VB_Name = "modProcesses"
Option Explicit

Public Type tProcess
 ExeName As String
 ID As Long
 UserName As String
End Type

Public Function FindProcess(UserName As String, Optional SessionID As Long = 0) As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim res As Long, ProcessIDs As Collection
50020  If IsTerminalServer = True Then
50030    Set ProcessIDs = FindTSProcess(UserName, SessionID)
50040    WriteToSpecialLogfile "Terminal token is asked for (" & UserName & ", " & SessionID & "): Found " & ProcessIDs.Count & " processes"
50050    If ProcessIDs.Count = 0 Then
50060     WriteToSpecialLogfile "Cannot detect a terminal token. Looking for console token."
50070     Set ProcessIDs = FindNormalProcess(UserName)
50080     WriteToSpecialLogfile "Console token is asked for (" & UserName & "): Found " & ProcessIDs.Count & " processes"
50090    End If
50100   Else
50110    Set ProcessIDs = FindNormalProcess(UserName)
50120    WriteToSpecialLogfile "Console token is asked for (" & UserName & "): Found " & ProcessIDs.Count & " processes"
50130  End If
50140  Set FindProcess = ProcessIDs
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modProcesses", "FindProcess")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

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
'   01/14/2009 Frank Heindörfer
Private Function FindTSProcess(UserName As String, SessionID As Long) As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim RetVal As Long, Count As Long, i As Integer, lpBuffer As Long, _
  p As Long, udtProcessInfo As WTS_PROCESS_INFO, ProcessIDs As Collection, process As clsProcess
50030  Set ProcessIDs = New Collection
50040  WriteToSpecialLogfile "Process enumeration..."
50050  RetVal = WTSEnumerateProcesses(WTS_CURRENT_SERVER_HANDLE, 0&, 1, lpBuffer, Count)
50060  If RetVal <> 0 Then ' WTSEnumerateProcesses was successful
50070    p = lpBuffer
50080    i = 1
50090    While i <= Count
50100     MoveMemory udtProcessInfo, ByVal p, LenB(udtProcessInfo)
50110     WriteToSpecialLogfile "Process: SessionID=" & udtProcessInfo.SessionID & " ProcessID=" & udtProcessInfo.ProcessID
50120     If udtProcessInfo.SessionID = SessionID And udtProcessInfo.ProcessID > 0 Then
50130      ' Retrieve the name of the account and the name of the first
50140      ' domain on which this SID is found.  Passes in the Owner's SID
50150      ' obtained previously.  Call LookupAccountSid twice, the first time
50160      ' to obtain the required size of the owner and domain names.
50170      If UCase$(GetUsernameFromUserSID(udtProcessInfo.pUserSid)) = UCase$(UserName) Then
50180       Set process = New clsProcess
50190       process.ID = udtProcessInfo.ProcessID
50200       process.Modulname = GetStrFromPtrA(udtProcessInfo.pProcessName)
50210       ProcessIDs.Add process
50220       WriteToSpecialLogfile "Process found: ProcessID=" & udtProcessInfo.ProcessID & " Username=" & UserName & " Modulname=" & process.Modulname
50230      End If
50240     End If
50250     i = i + 1
50260     p = p + LenB(udtProcessInfo)
50270    Wend
50280    WTSFreeMemory lpBuffer   'Free your memory buffer
50290   Else
50300    ' Error occurred calling WTSEnumerateProcesses
50310    WriteToSpecialLogfile "Error occurred calling WTSEnumerateProcesses.  " & RaiseAPIError
50320  End If
50330  Set FindTSProcess = ProcessIDs
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modProcesses", "FindTSProcess")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function FindNormalProcess(UserName As String) As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ProcessIDs As Collection, nResult As Long, lCb As Long, lCbNeeded As Long, lCbNeeded2 As Long, _
  lProcID() As Long, lModules(1 To 200) As Long, hProcess As Long, sModuleName As String, n As Long, c As Long, _
  process As clsProcess
50040  Set ProcessIDs = New Collection
50050  c = 0
50060  If IsWinNT4 = True Then
50070    lCb = 8: lCbNeeded = 96
50080    Do While lCb <= lCbNeeded
50090     lCb = lCb * 2
50100     ReDim lProcID(lCb / 4) As Long
50110     EnumProcesses lProcID(1), lCb, lCbNeeded
50120    Loop
50130    For n = 1 To lCbNeeded / 4
50140     hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lProcID(n))
50150     If hProcess <> 0 Then
50160      nResult = EnumProcessModules(hProcess, lModules(1), 200, lCbNeeded2)
50170      If nResult <> 0 Then
50180       sModuleName = Space(MAX_PATH)
50190       nResult = GetModuleFileNameEx(hProcess, lModules(1), sModuleName, Len(sModuleName))
50200       sModuleName = LCase$(Left$(sModuleName, nResult))
50210       If UCase$(GetProcessUserName(lProcID(n))) = UCase$(UserName) Then
50220        Set process = New clsProcess
50230        process.ID = lProcID(n)
50240        process.Modulname = sModuleName
50250        ProcessIDs.Add process
50260       End If
50270      End If
50280     End If
50290     CloseHandle hProcess
50300    Next n
50310   Else
50320    Dim lSnapshot As Long, uProcess As PROCESSENTRY32, Exefile As String
50330    lSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
50340    If lSnapshot <> 0 Then
50350      uProcess.dwSize = Len(uProcess)
50360      nResult = ProcessFirst(lSnapshot, uProcess)
50370      Do Until nResult = 0
50380       Exefile = uProcess.szexeFile
50390       Exefile = Left$(Exefile, InStr(Exefile, Chr$(0)) - 1)
50400       If Right$(LCase(Exefile), 4) = ".exe" Then
50410        If UCase$(GetProcessUserName(uProcess.th32ProcessID)) = UCase$(UserName) Then
50420         Set process = New clsProcess
50430         process.ID = uProcess.th32ProcessID
50440         process.Modulname = Exefile
50450         ProcessIDs.Add process
50460        End If
50470       End If
50480       nResult = ProcessNext(lSnapshot, uProcess)
50490      Loop
50500     CloseHandle lSnapshot
50510    End If
50520  End If
50530  Set FindNormalProcess = ProcessIDs
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modProcesses", "FindNormalProcess")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetProcessUserName(ByVal ProcessID As Long) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim hProcessID As Long, hToken As Long, res As Long, cbBuff As Long, _
  tiLen As Long, TG As TOKEN_GROUPS, TU As TOKEN_USER, _
  SIA As SID_IDENTIFIER_AUTHORITY, lSid As Long, cnt As Long, _
  sAcctName1 As String, sAcctName2 As String, cbAcctName As Long, _
  sDomainName As String, cbDomainName As Long, peUse As Long
50060
50070  tiLen = 0
50080  hProcessID = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessID)
50090  If hProcessID <> 0 Then
50100   If OpenProcessToken(hProcessID, TokenRead, hToken) = 1 Then
50110    res = GetTokenInformation(hToken, TokenUser, TU, tiLen, cbBuff)
50120    If res = 0 And cbBuff > 0 Then
50130     tiLen = cbBuff
50140     res = GetTokenInformation(hToken, TokenUser, TU, tiLen, cbBuff)
50150     If res = 1 And tiLen > 0 Then
50160      SIA.value(5) = SECURITY_NT_AUTHORITY
50170      res = AllocateAndInitializeSid(SIA, 2, _
      SECURITY_BUILTIN_DOMAIN_RID, _
      DOMAIN_ALIAS_RID_USERS, 0, 0, 0, 0, 0, 0, lSid)
50200      If res = 1 Then
50210       sAcctName1 = Space$(255)
50220       sDomainName = Space$(255)
50230       cbAcctName = 255
50240       cbDomainName = 255
50250       res = LookupAccountSid(vbNullString, lSid, sAcctName1, cbAcctName, sDomainName, cbDomainName, peUse)
50260       If res = 1 Then
50270        sAcctName2 = Space$(255)
50280        sDomainName = Space$(255)
50290        cbAcctName = 255
50300        cbDomainName = 255
50310        res = LookupAccountSid(vbNullString, TU.User.Sid, sAcctName2, cbAcctName, sDomainName, cbDomainName, peUse)
50320        GetProcessUserName = Replace(Trim(sAcctName2), Chr(0), "")
50330       End If
50340       FreeSid ByVal lSid
50350      End If
50360      CloseHandle hToken
50370     End If
50380    End If
50390   End If
50400   CloseHandle hProcessID
50410  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modProcesses", "GetProcessUserName")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetProcessUserNameA(ByVal ProcessID As Long) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim hProcessID As Long, hToken As Long, res As Long, cbBuff As Long, _
  tiLen As Long, TG As TOKEN_GROUPS, TU As TOKEN_USER, _
  SIA As SID_IDENTIFIER_AUTHORITY, lSid As Long, cnt As Long, _
  sAcctName1 As String, sAcctName2 As String, cbAcctName As Long, _
  sDomainName As String, cbDomainName As Long, peUse As Long
50060
50070  tiLen = 0
50080  hProcessID = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessID)
50090  If hProcessID <> 0 Then
50100   If OpenProcessToken(hProcessID, TokenRead, hToken) = 1 Then
50110    res = GetTokenInformation(hToken, TokenUser, TU, tiLen, cbBuff)
50120    If res = 0 And cbBuff > 0 Then
50130     tiLen = cbBuff
50140     res = GetTokenInformation(hToken, TokenUser, TU, tiLen, cbBuff)
50150     If res = 1 And tiLen > 0 Then
50160      SIA.value(5) = SECURITY_NT_AUTHORITY
50170      res = AllocateAndInitializeSid(SIA, 2, _
      SECURITY_BUILTIN_DOMAIN_RID, _
      DOMAIN_ALIAS_RID_USERS, 0, 0, 0, 0, 0, 0, lSid)
50200      If res = 1 Then
50210       sAcctName1 = Space$(255): sDomainName = Space$(255)
50220       cbAcctName = 255: cbDomainName = 255
50230       res = LookupAccountSid(vbNullString, lSid, sAcctName1, cbAcctName, sDomainName, cbDomainName, peUse)
50240       If res = 1 Then
50250        sAcctName2 = Space$(255): sDomainName = Space$(255)
50260        cbAcctName = 255: cbDomainName = 255
50270        res = LookupAccountSid(vbNullString, TU.User.Sid, sAcctName2, cbAcctName, sDomainName, cbDomainName, peUse)
50280        GetProcessUserNameA = Replace(Trim(sAcctName2), Chr(0), "")
50290 '       GetProcessUserNameA = GetUsernameFromUserSID(TU.User.Sid)
50300       End If
50310       FreeSid ByVal lSid
50320      End If
50330      CloseHandle hToken
50340     End If
50350    End If
50360   End If
50370   CloseHandle hProcessID
50380  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modProcesses", "GetProcessUserNameA")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function GetUsernameFromUserSID(UserSID As Long) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim bSuccess As Long, pOwner As Long, Name As String, domain_name As String, _
  name_len As Long, domain_len As Long, deUse As Long, ProcessOwner As String
50030
50040  Name = "": domain_name = "": name_len = 0: domain_len = 0
50050  bSuccess = LookupAccountSid(vbNullString, UserSID, Name, _
  name_len, domain_name, domain_len, deUse)
50070  If name_len > 0 And domain_len > 0 Then
50080    Name = Space(name_len - 1)
50090    domain_name = Space(domain_len - 1)
50100    bSuccess = LookupAccountSid(vbNullString, UserSID, Name, _
    name_len, domain_name, domain_len, deUse)
50120    If bSuccess <> 0 Then
50130      GetUsernameFromUserSID = Name
50140     Else
50150      WriteToSpecialLogfile "Can't look up account!"
50160    End If
50170   Else
50180    WriteToSpecialLogfile "Invalid name length for LookupAccountSid!"
50190  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modProcesses", "GetUsernameFromUserSID")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function GetUserSIDFromProcessID(ProcessID) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim hProcessID As Long, hToken As Long, res As Long, cbBuff As Long, _
  tiLen As Long, TG As TOKEN_GROUPS, TU As TOKEN_USER, _
  SIA As SID_IDENTIFIER_AUTHORITY, lSid As Long, cnt As Long, _
  sAcctName1 As String, sAcctName2 As String, cbAcctName As Long, _
  sDomainName As String, cbDomainName As Long, peUse As Long
50060  hProcessID = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessID)
50070  If hProcessID <> 0 Then
50080   If OpenProcessToken(hProcessID, TokenRead, hToken) = 1 Then
50090    res = GetTokenInformation(hToken, TokenUser, TU, tiLen, cbBuff)
50100    If res = 0 And cbBuff > 0 Then
50110     tiLen = cbBuff
50120     res = GetTokenInformation(hToken, TokenUser, TU, tiLen, cbBuff)
50130     If res = 1 And tiLen > 0 Then
50140      SIA.value(5) = SECURITY_NT_AUTHORITY
50150      res = AllocateAndInitializeSid(SIA, 2, _
      SECURITY_BUILTIN_DOMAIN_RID, _
      DOMAIN_ALIAS_RID_USERS, 0, 0, 0, 0, 0, 0, lSid)
50180      If res = 1 Then
50190       GetUserSIDFromProcessID = lSid
50200 '      FreeSid ByVal lSid
50210      End If
50220      CloseHandle hToken
50230     End If
50240    End If
50250   End If
50260   CloseHandle hProcessID
50270  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modProcesses", "GetUserSIDFromProcessID")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

