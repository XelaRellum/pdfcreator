Attribute VB_Name = "modNetuser"
Option Explicit

Const NERR_Success = 0
Private Const NERR_BASE = 2100
Private Const NERR_InvalidComputer = (NERR_BASE + 251)
Private Const NERR_UseNotFound = (NERR_BASE + 150)
Const CP_ACP = 0
Private Type USER_INFO_3
    usri3_name As Long
    usri3_password As Long
    usri3_password_age As Long
    usri3_priv As Long
    usri3_home_dir As Long
    usri3_comment As Long
    usri3_flags As Long
    usri3_script_path As Long
    usri3_auth_flags As Long
    usri3_full_name As Long
    usri3_usr_comment As Long
    usri3_parms As Long
    usri3_workstations As Long
    usri3_last_logon As Long
    usri3_last_logoff As Long
    usri3_acct_expires As Long
    usri3_max_storage As Long
    usri3_units_per_week As Long
    usri3_logon_hours As Byte
    usri3_bad_pw_count As Long
    usri3_num_logons As Long
    usri3_logon_server As String
    usri3_country_code As Long
    usri3_code_page As Long
    usri3_user_id As Long
    usri3_primary_group_id As Long
    usri3_profile As Long
    usri3_home_dir_drive As Long
    usri3_password_expired As Long
End Type
Private Declare Function NetUserGetInfo Lib "netapi32" (lpServer As Any, UserName As Byte, ByVal Level As Long, lpBuffer As Long) As Long
Private Declare Function NetApiBufferFree Lib "netapi32" (ByVal Buffer As Long) As Long
Private Declare Sub MoveMemory Lib "Kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Function lstrlenW Lib "Kernel32" (lpString As Any) As Long
Private Declare Function WideCharToMultiByte Lib "Kernel32" (ByVal codepage As Long, ByVal dwFlags As Long, lpWideCharStr As Any, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
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

