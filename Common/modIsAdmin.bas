Attribute VB_Name = "modIsAdmin"
Option Explicit
Option Base 0     ' Important assumption for this code

 'Fixed at this size for comfort. Could be bigger or made dynamic.
Private Const ANYSIZE_ARRAY As Long = 1000

 ' Security APIs
Private Const TokenUser = 1
Private Const TokenGroups = 2
Private Const TokenPrivileges = 3
Private Const TokenOwner = 4
Private Const TokenPrimaryGroup = 5
Private Const TokenDefaultDacl = 6
Private Const TokenSource = 7
Private Const TokenType = 8
Private Const TokenImpersonationLevel = 9
Private Const TokenStatistics = 10

 ' Token Specific Access Rights
Private Const TOKEN_ASSIGN_PRIMARY = &H1
Private Const TOKEN_DUPLICATE = &H2
Private Const TOKEN_IMPERSONATE = &H4
Private Const TOKEN_QUERY = &H8
Private Const TOKEN_QUERY_SOURCE = &H10
Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_ADJUST_GROUPS = &H40
Private Const TOKEN_ADJUST_DEFAULT = &H80

 ' NT well-known SIDs
Private Const SECURITY_DIALUP_RID = &H1
Private Const SECURITY_NETWORK_RID = &H2
Private Const SECURITY_BATCH_RID = &H3
Private Const SECURITY_INTERACTIVE_RID = &H4
Private Const SECURITY_SERVICE_RID = &H6
Private Const SECURITY_ANONYMOUS_LOGON_RID = &H7
Private Const SECURITY_LOGON_IDS_RID = &H5
Private Const SECURITY_LOCAL_SYSTEM_RID = &H12
Private Const SECURITY_NT_NON_UNIQUE = &H15
Private Const SECURITY_BUILTIN_DOMAIN_RID = &H20

 ' Well-known domain relative sub-authority values (RIDs)
Private Const DOMAIN_ALIAS_RID_ADMINS = &H220
Private Const DOMAIN_ALIAS_RID_USERS = &H221
Private Const DOMAIN_ALIAS_RID_GUESTS = &H222
Private Const DOMAIN_ALIAS_RID_POWER_USERS = &H223
Private Const DOMAIN_ALIAS_RID_ACCOUNT_OPS = &H224
Private Const DOMAIN_ALIAS_RID_SYSTEM_OPS = &H225
Private Const DOMAIN_ALIAS_RID_PRINT_OPS = &H226
Private Const DOMAIN_ALIAS_RID_BACKUP_OPS = &H227
Private Const DOMAIN_ALIAS_RID_REPLICATOR = &H228

Private Const SECURITY_NT_AUTHORITY = &H5

Private Type SID_AND_ATTRIBUTES
    Sid As Long
    Attributes As Long
End Type

Private Type TOKEN_GROUPS
    GroupCount As Long
    Groups(ANYSIZE_ARRAY) As SID_AND_ATTRIBUTES
End Type

Private Type SID_IDENTIFIER_AUTHORITY
    Value(0 To 5) As Byte
End Type

Private Declare Function GetCurrentProcess Lib "Kernel32" () As Long

Private Declare Function GetCurrentThread Lib "Kernel32" () As Long

Private Declare Function OpenProcessToken Lib "Advapi32" (ByVal ProcessHandle _
    As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long

Private Declare Function OpenThreadToken Lib "Advapi32" (ByVal ThreadHandle As _
    Long, ByVal DesiredAccess As Long, ByVal OpenAsSelf As Long, _
    TokenHandle As Long) As Long

Private Declare Function GetTokenInformation Lib "Advapi32" (ByVal TokenHandle _
    As Long, TokenInformationClass As Integer, TokenInformation As Any, _
    ByVal TokenInformationLength As Long, ReturnLength As Long) As Long

Private Declare Function AllocateAndInitializeSid Lib "Advapi32" _
    (pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, ByVal nSubAuthorityCount _
    As Byte, ByVal nSubAuthority0 As Long, ByVal nSubAuthority1 As Long, _
    ByVal nSubAuthority2 As Long, ByVal nSubAuthority3 As Long, _
    ByVal nSubAuthority4 As Long, ByVal nSubAuthority5 As Long, _
    ByVal nSubAuthority6 As Long, ByVal nSubAuthority7 As Long, _
    lpPSid As Long) As Long

Private Declare Function RtlMoveMemory Lib "Kernel32" (Dest As Any, _
    Source As Any, ByVal lSize As Long) As Long

Private Declare Function IsValidSid Lib "Advapi32" (ByVal pSid As Long) As Long

Private Declare Function EqualSid Lib "Advapi32" (pSid1 As Any, _
    pSid2 As Any) As Long

Private Declare Sub FreeSid Lib "Advapi32" (pSid As Any)

Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As _
    Long


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
50380             Call RtlMoveMemory(tpTokens, InfoBuffer(0), LenB(tpTokens))
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
