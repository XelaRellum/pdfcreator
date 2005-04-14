Attribute VB_Name = "modNetuser"
Option Explicit

' Returns an ANSI string from a pointer to a Unicode string.
Public Function GetStrFromPtrW(lpszW As Long) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Dim sRtn As String
50020     sRtn = String$(lstrlenW(ByVal lpszW) * 2, 0)   ' 2 bytes/char
50030     ' WideCharToMultiByte also returns Unicode string length
50040     Call WideCharToMultiByte(CP_ACP, 0, ByVal lpszW, -1, ByVal sRtn, Len(sRtn), 0, 0)
50050     GetStrFromPtrW = GetStrFromBufferA(sRtn)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modNetuser", "GetStrFromPtrW")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
' Returns the string before first null char encountered (if any) from an ANSII string.
Public Function GetStrFromBufferA(sz As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     If InStr(sz, vbNullChar) Then
50020         GetStrFromBufferA = Left$(sz, InStr(sz, vbNullChar) - 1)
50030     Else
50040         ' If sz had no null char, the Left$ function
50050         ' above would return a zero length string ("").
50060         GetStrFromBufferA = sz
50070     End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modNetuser", "GetStrFromBufferA")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function GetUserProfilepath(sUsername As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim lpBuf As Long, ui3 As USER_INFO_3, bServer() As Byte, bUsername() As Byte
50020  bServer = "" & vbNullChar
50030  bUsername = sUsername & vbNullChar
50040  If (NetUserGetInfo(bServer(0), bUsername(0), 3, lpBuf) = NERR_Success) Then
50050   Call MoveMemory(ui3, ByVal lpBuf, Len(ui3))
50060   GetUserProfilepath = GetStrFromPtrW(ui3.usri3_profile)
50070   Call NetApiBufferFree(ByVal lpBuf)
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modNetuser", "GetUserProfilepath")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

