Attribute VB_Name = "modIsAdmin"
Option Explicit

Public Function IsAdmin() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim hSC As Long, lRet As Long, tOSV  As OSVERSIONINFO
50020  tOSV.OSVSize = Len(tOSV)
50030  lRet = GetVersionEx(tOSV)
50040  If tOSV.PlatformID = VER_PLATFORM_WIN32_NT Then
50050    hSC = OpenSCManager(vbNullString, vbNullString, GENERIC_READ Or GENERIC_WRITE Or GENERIC_EXECUTE)
50060    If (hSC <> 0) Then
50070     IsAdmin = True
50080     CloseServiceHandle (hSC)
50090    End If
50100   Else
50110    IsAdmin = True
50120  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modIsAdmin", "IsAdmin")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

