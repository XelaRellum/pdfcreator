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
    Dim sRtn As String
    sRtn = String$(lstrlenW(ByVal lpszW) * 2, 0)   ' 2 bytes/char
    ' WideCharToMultiByte also returns Unicode string length
    Call WideCharToMultiByte(CP_ACP, 0, ByVal lpszW, -1, ByVal sRtn, Len(sRtn), 0, 0)
    GetStrFromPtrW = GetStrFromBufferA(sRtn)
End Function
' Returns the string before first null char encountered (if any) from an ANSII string.
Public Function GetStrFromBufferA(sz As String) As String
    If InStr(sz, vbNullChar) Then
        GetStrFromBufferA = Left$(sz, InStr(sz, vbNullChar) - 1)
    Else
        ' If sz had no null char, the Left$ function
        ' above would return a zero length string ("").
        GetStrFromBufferA = sz
    End If
End Function

Private Function GetUserProfilepath(sUsername As String)
 Dim lpBuf As Long, ui3 As USER_INFO_3, bServer() As Byte, bUsername() As Byte
 bServer = "" & vbNullChar
 bUsername = sUsername & vbNullChar
 If (NetUserGetInfo(bServer(0), bUsername(0), 3, lpBuf) = NERR_Success) Then
  Call MoveMemory(ui3, ByVal lpBuf, Len(ui3))
  GetUserProfilepath = GetStrFromPtrW(ui3.usri3_profile)
  Call NetApiBufferFree(ByVal lpBuf)
 End If
End Function

