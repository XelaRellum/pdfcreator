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
50010     Dim hProcessToken       As Long
50020     Dim BufferSize          As Long
50030     Dim psidAdmin           As Long
50040     Dim lResult             As Long
50050     Dim X                   As Integer
50060     Dim tpTokens            As TOKEN_GROUPS
50070     Dim tpSidAuth           As SID_IDENTIFIER_AUTHORITY
50080     Dim llRetVal            As Long
50090     Dim InfoBuffer()        As Long
50100     Dim sids()              As SID_AND_ATTRIBUTES
50110     Dim llCount             As Long
50120     Dim llIdx               As Long
50130     Dim llMax               As Long
50140     IsAdmin = False
50150     tpSidAuth.Value(5) = SECURITY_NT_AUTHORITY
50160
50170      ' Obtain current process token
50180     If Not OpenThreadToken(GetCurrentThread(), TOKEN_QUERY, True, _
        hProcessToken) Then
50200         Call OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, hProcessToken)
50210     End If
50220     If hProcessToken Then
50230
50240          ' Deternine the buffer size required
50250         llRetVal = GetTokenInformation(hProcessToken, ByVal TokenGroups, 0, 0, _
            BufferSize) ' Determine required buffer size
50270         If BufferSize Then
50280
50290             ReDim InfoBuffer((BufferSize \ 4) - 1) As Long
50300             ReDim sids(0 To tpTokens.GroupCount) As SID_AND_ATTRIBUTES
50310              ' Retrieve your token information
50320             lResult = GetTokenInformation(hProcessToken, ByVal TokenGroups, _
                InfoBuffer(0), BufferSize, BufferSize)
50340
50350             If lResult <> 1 Then Exit Function
50360
50370              ' Move it from memory into the token structure
50380             Call MoveMemory(tpTokens, InfoBuffer(0), LenB(tpTokens))
50390
50400              ' Retreive the admins sid pointer
50410             lResult = AllocateAndInitializeSid(tpSidAuth, 2, _
                SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, _
                0, 0, 0, psidAdmin)
50440             If lResult <> 1 Then Exit Function
50450             If IsValidSid(psidAdmin) Then
50460                 For X = 0 To tpTokens.GroupCount
50470
50480                      ' Run through your token sid pointers
50490                     If IsValidSid(tpTokens.Groups(X).Sid) Then
50500
50510                          ' Test for a match between the admin sid equalling
50520                          ' your Sid 's
50530                         If EqualSid(ByVal tpTokens.Groups(X).Sid, _
                            ByVal psidAdmin) Then
50550                             IsAdmin = True
50560                             Exit For
50570                         End If
50580                     End If
50590                 Next
50600             End If
50610             If psidAdmin Then Call FreeSid(psidAdmin)
50620         End If
50630         Call CloseHandle(hProcessToken)
50640     End If
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
'Developed for you by Elvio Serrao (Elvio.Serrao@nrma.com.au)
'##################################
