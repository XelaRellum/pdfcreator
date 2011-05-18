Attribute VB_Name = "modServices"
Option Explicit

Public Const SC_MANAGER_ALL_ACCESS   As Long = &HF003F
Public Const SERVICE_START           As Long = &H10
Public Const SERVICE_STOP            As Long = &H20
Public Const SERVICE_CONTROL_STOP    As Long = &H1
' SERVICE_QUERY_STATUS  = $4;
' SERVICE_RUNNING       = $4;

Public Type SERVICE_STATUS
  dwServiceType As Long
  dwCurrentState As Long
  dwControlsAccepted As Long
  dwWin32ExitCode As Long
  dwServiceSpecificExitCode As Long
  dwCheckPoint As Long
  dwWaitHint As Long
End Type

Public Declare Function ControlService Lib "advapi32.dll" (ByVal hService As Long, ByVal dwControl As Long, lpServiceStatus As SERVICE_STATUS) As Long
Public Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Public Declare Function StartServiceA Lib "advapi32.dll" (ByVal hService As Long, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long

Private Function OpenServiceManager() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  OpenServiceManager = OpenSCManager("", "ServicesActive", SC_MANAGER_ALL_ACCESS)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modServices", "OpenServiceManager")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function StartService(ServiceName As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim hSCM As Long, hService As Long
50020  hSCM = OpenServiceManager
50030  StartService = False
50040  If hSCM <> 0 Then
50050   hService = OpenService(hSCM, ServiceName, SERVICE_START)
50060   If hService <> 0 Then
50070    StartService = StartServiceA(hService, 0, 0)
50080    Call CloseServiceHandle(hService)
50090   End If
50100   Call CloseServiceHandle(hSCM)
50110  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modServices", "StartService")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function StopService(ServiceName As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim hSCM As Long, hService As Long, status As SERVICE_STATUS
50020  hSCM = OpenServiceManager
50030  StopService = False
50040  If hSCM <> 0 Then
50050   hService = OpenService(hSCM, ServiceName, SERVICE_STOP)
50060   If hService <> 0 Then
50070    StopService = ControlService(hService, SERVICE_CONTROL_STOP, status)
50080    Call CloseServiceHandle(hService)
50090   End If
50100   Call CloseServiceHandle(hSCM)
50110  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modServices", "StopService")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

