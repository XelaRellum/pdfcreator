Attribute VB_Name = "modIsAdmin"
Option Explicit
Option Base 0     ' Important assumption for this code

' Returns True if the thread is running in the
' user context of the local Administrator account
' Example:
'   MsgBox "Current user is the Administrator: " & IsAdmin
Public Function IsAdmin() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim hProcessToken As Long, BufferSize As Long, psidAdmin As Long, _
  lResult As Long, x As Integer, tpTokens As TOKEN_GROUPS, _
  tpSidAuth As SID_IDENTIFIER_AUTHORITY, llRetVal As Long, _
  InfoBuffer() As Long, sids() As SID_AND_ATTRIBUTES, llCount As Long, _
  llIdx As Long, llMax As Long
50060
50070  IsAdmin = False
50080  tpSidAuth.Value(5) = SECURITY_NT_AUTHORITY
50090
50100  ' Obtain current process token
50110  If Not OpenThreadToken(GetCurrentThread(), TOKEN_QUERY, True, _
  hProcessToken) Then
50130   Call OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, hProcessToken)
50140  End If
50150  If hProcessToken Then
50160   ' Deternine the buffer size required
50170   llRetVal = GetTokenInformation(hProcessToken, ByVal TokenGroups, 0, 0, BufferSize)             ' Determine required buffer size
50180   If BufferSize Then
50190    ReDim InfoBuffer((BufferSize \ 4) - 1) As Long
50200    ReDim sids(0 To tpTokens.GroupCount) As SID_AND_ATTRIBUTES
50210    ' Retrieve your token information
50220    lResult = GetTokenInformation(hProcessToken, ByVal TokenGroups, InfoBuffer(0), BufferSize, BufferSize)
50230    If lResult <> 1 Then Exit Function
50240    ' Move it from memory into the token structure
50250    Call MoveMemory(tpTokens, InfoBuffer(0), LenB(tpTokens))
50260    ' Retreive the admins sid pointer
50270    lResult = AllocateAndInitializeSid(tpSidAuth, 2, _
    SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, _
    0, 0, 0, psidAdmin)
50300    If lResult <> 1 Then Exit Function
50310    If IsValidSid(psidAdmin) Then
50320     For x = 0 To tpTokens.GroupCount - 1
50330      ' Run through your token sid pointers
50340      If IsValidSid(tpTokens.Groups(x).Sid) Then
50350       ' Test for a match between the admin sid equalling
50360       ' your Sid 's
50370       If EqualSid(ByVal tpTokens.Groups(x).Sid, ByVal psidAdmin) Then
50380        IsAdmin = True
50390        Exit For
50400       End If
50410      End If
50420     Next x
50430    End If
50440    If psidAdmin Then Call FreeSid(psidAdmin)
50450   End If
50460   Call CloseHandle(hProcessToken)
50470  End If
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

'##################################
' Developed for you by Elvio Serrao (Elvio.Serrao@nrma.com.au)
' Fixed some minor bugs by Frank Heindörfer (thesmilyface AT users.sourceforge.net)
'##################################
