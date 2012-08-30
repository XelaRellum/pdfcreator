Attribute VB_Name = "modIP"
Option Explicit

Public IPErrStr As String

Public Function GetHostNameFromIP(ByVal sAddress As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ptrHosent As Long, hAddress As Long, nbytes As Long
50020  IPErrStr = ""
50030  If SocketsInitialize() Then
50040    hAddress = inet_addr(sAddress)
50050    If hAddress <> SOCKET_ERROR Then
50060      ptrHosent = gethostbyaddr(hAddress, 4, AF_INET)
50070      If ptrHosent <> 0 Then
50080        MoveMemory ptrHosent, ByVal ptrHosent, 4
50090        nbytes = lstrlen(ByVal ptrHosent)
50100        If nbytes > 0 Then
50110         sAddress = Space$(nbytes)
50120         MoveMemory ByVal sAddress, ByVal ptrHosent, nbytes
50130         GetHostNameFromIP = sAddress
50140        End If
50150       Else
50160        IPErrStr = "Call to gethostbyaddr failed."
50170      End If
50180      SocketsCleanup
50190     Else
50200      IPErrStr = "String passed is an invalid IP."
50210    End If
50220   Else
50230    IPErrStr = "Sockets failed to initialize."
50240  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modIP", "GetHostNameFromIP")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function SocketsInitialize() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim WSAD As WSADATA
50020  SocketsInitialize = WSAStartup(WS_VERSION_REQD, WSAD) = IP_SUCCESS
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modIP", "SocketsInitialize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub SocketsCleanup()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If WSACleanup() <> 0 Then
50020   MsgBox "Windows Sockets error occurred in Cleanup.", vbExclamation
50030  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modIP", "SocketsCleanup")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function IsIPAddress(ByVal IPAddressStr As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tStrf() As String, IPAddress(3) As Long
50020  IsIPAddress = False
50030  IPAddressStr = Trim(IPAddressStr)
50040  If LenB(IPAddressStr) > 0 Then
50050   If InStr(1, IPAddressStr, ".") > 0 Then
50060    tStrf = Split(IPAddressStr, ".")
50070    If UBound(tStrf) = 3 Then
50080     If IsNumeric(tStrf(0)) And IsNumeric(tStrf(1)) And _
     IsNumeric(tStrf(2)) And IsNumeric(tStrf(3)) Then
50100      IPAddress(0) = CLng(tStrf(0))
50110      IPAddress(1) = CLng(tStrf(1))
50120      IPAddress(2) = CLng(tStrf(2))
50130      IPAddress(3) = CLng(tStrf(3))
50140      If (IPAddress(0) >= 0 And IPAddress(0) <= 255) And _
      (IPAddress(1) >= 0 And IPAddress(1) <= 255) And _
      (IPAddress(2) >= 0 And IPAddress(2) <= 255) And _
      (IPAddress(3) >= 0 And IPAddress(3) <= 255) Then
50180       If Not (IPAddress(0) = 0 And IPAddress(1) = 0 And _
       IPAddress(2) = 0 And IPAddress(3) = 0) Then
50200        IsIPAddress = True
50210       End If
50220      End If
50230     End If
50240    End If
50250   End If
50260  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modIP", "IsIPAddress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
