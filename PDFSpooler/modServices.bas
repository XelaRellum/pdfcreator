Attribute VB_Name = "modServices"
Option Explicit

Public Function EnumLocalServices(SERVICE_TYPE As eServiceState) As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim hSCManager As Long, pntr() As ENUM_SERVICE_STATUS, cbBuffSize As Long
50020  Dim cbRequired As Long, dwReturned As Long, hEnumResume As Long, cbBuffer As Long, success As Long, i As Long
50030
50040  Dim sSvcName As String, sDispName As String, dwState As Long, dwType As Long, dwCtrls As Long
50050  Dim service As clsService, services As Collection
50060  Set services = New Collection
50070  hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_ENUMERATE_SERVICE)
50080  If hSCManager <> 0 Then
50090   success = EnumServicesStatus(hSCManager, SERVICE_WIN32, SERVICE_TYPE, ByVal &H0, &H0, cbRequired, dwReturned, hEnumResume)
50100   If success = 0 And Err.LastDllError = ERROR_MORE_DATA Then
50110    cbBuffer = (cbRequired \ SIZEOF_SERVICE_STATUS) + 1
50120    ReDim pntr(0 To cbBuffer)
50130    cbBuffSize = cbBuffer * SIZEOF_SERVICE_STATUS
50140    hEnumResume = 0
50150    If EnumServicesStatus(hSCManager, SERVICE_WIN32, SERVICE_TYPE, pntr(0), cbBuffSize, cbRequired, dwReturned, hEnumResume) Then
50160     For i = 0 To dwReturned - 1
50170      Set service = New clsService
50180      With service
50190       .DisplayName = GetStrFromPtrA(pntr(i).lpDisplayName)
50200       .ServiceName = GetStrFromPtrA(pntr(i).lpServiceName)
50210       .CurrentState = pntr(i).ServiceStatus.dwCurrentState
50220       .ServiceType = pntr(i).ServiceStatus.dwServiceType
50230       .ControlsAccepted = pntr(i).ServiceStatus.dwControlsAccepted
50240       .ImagePath = GetServiceImagePath(.ServiceName)
50250      End With
50260      services.Add service
50270     Next i
50280    End If
50290   End If
50300  End If
50310  Call CloseServiceHandle(hSCManager)
50320  Set EnumLocalServices = services
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modServices", "EnumSystemServices")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetServiceImagePath(ServiceName As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry
50020  Set reg = New clsRegistry
50030  With reg
50040   .hkey = HKEY_LOCAL_MACHINE
50050   .KeyRoot = "SYSTEM\CurrentControlSet\Services"
50060   .Subkey = ServiceName
50070   If .KeyExists = False Then
50080    Set reg = Nothing
50090    Exit Function
50100   End If
50110   GetServiceImagePath = .GetRegistryValue("ImagePath")
50120  End With
50130  Set reg = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modServices", "GetServiceImagePath")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function IsService(ExeFilename As String, AllActiveServices As Collection) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim service As clsService, i As Long, filename As String
50020  IsService = False
50030  For i = 1 To AllActiveServices.Count
50040   Set service = AllActiveServices(i)
50050   SplitPath service.ImagePath, , , filename
50060   If InStr(1, filename, ExeFilename, vbTextCompare) > 0 Then
50070    IsService = True
50080    Exit Function
50090   End If
50100  Next i
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modServices", "IsService")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
