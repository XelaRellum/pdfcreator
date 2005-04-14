VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsURL"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function InternetCrackUrl Lib "wininet.dll" Alias "InternetCrackUrlA" (ByVal lpszURL As String, ByVal dwUrlLength As Long, ByVal dwFlags As Long, lpUrlComponents As URL_COMPONENTS) As Long
Private Declare Function InternetCanonicalizeUrl Lib "wininet.dll" Alias "InternetCanonicalizeUrlA" (ByVal lpszURL As String, ByVal lpszBuffer As String, lpdwBufferLength As Long, ByVal dwFlags As Long) As Long

Private Type URL_COMPONENTS       'typedef struct {
   StructSize As Long            '    DWORD dwStructSize;
   Scheme As String              '    LPSTR lpszScheme;
   SchemeLength As Long          '    DWORD dwSchemeLength;
   nScheme As InetSchemes 'Long               '    INTERNET_SCHEME nScheme;
   HostName As String            '    LPSTR lpszHostName;
   HostNameLength As Long        '    DWORD dwHostNameLength;
   nPort As Long                 '    INTERNET_PORT nPort;
   UserName As String            '    LPSTR lpszUserName;
   UserNameLength As Long        '    DWORD dwUserNameLength;
   Password As String            '    LPSTR lpszPassword;
   PasswordLength As Long        '    DWORD dwPasswordLength;
   URLPath As String             '    LPSTR lpszUrlPath;
   UrlPathLength As Long         '    DWORD dwUrlPathLength;
   ExtraInfo As String           '    LPSTR lpszExtraInfo;
   ExtraInfoLength As Long       '    DWORD dwExtraInfoLength;
End Type                         '} URL_COMPONENTS;

' Flags for InternetCrackUrl() and InternetCreateUrl()
Private Const ICU_ESCAPE = &H80000000       '// (un)escape URL characters
Private Const ICU_USERNAME = &H40000000     '// use internal username &amp; password

' Flags for InternetCanonicalizeUrl() and InternetCombineUrl()
Private Const ICU_NO_ENCODE = &H20000000    '// Don't convert unsafe characters to escape sequence
Private Const ICU_DECODE = &H10000000       '// Convert %XX escape sequences to characters
Private Const ICU_NO_META = &H8000000       '// Don't convert .. etc. meta path sequences
Private Const ICU_ENCODE_SPACES_ONLY = &H4000000   '// Encode spaces only
Private Const ICU_BROWSER_MODE = &H2000000  '// Special encode/decode rules for browser

' Possible error codes
Private Const ERROR_INSUFFICIENT_BUFFER As Long = 122

Private Const INTERNET_MAX_HOST_NAME_LENGTH = 256
Private Const INTERNET_MAX_USER_NAME_LENGTH = 128
Private Const INTERNET_MAX_PASSWORD_LENGTH = 128
Private Const INTERNET_MAX_PORT_NUMBER_LENGTH = 5          ' INTERNET_PORT is unsigned short
Private Const INTERNET_MAX_PORT_NUMBER_VALUE = 65535        ' maximum unsigned short value
Private Const INTERNET_MAX_PATH_LENGTH = 2048
Private Const INTERNET_MAX_SCHEME_LENGTH = 32               ' longest protocol name length
Private Const INTERNET_MAX_URL_LENGTH = INTERNET_MAX_SCHEME_LENGTH + 3 + INTERNET_MAX_PATH_LENGTH

Private mvarlScheme As InetSchemes
Private mvarScheme As String
Private mvarHostName As String
Private mvarPort As Long
Private mvarUserName As String
Private mvarPassword As String
Private mvarURLPath As String
Private mvarURLExtraInfo As String

Public Property Get URLExtraInfo() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 URLExtraInfo = mvarURLExtraInfo
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsURL", "URLExtraInfo [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get URLPath() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 URLPath = mvarURLPath
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsURL", "URLPath [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get Password() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 Password = mvarPassword
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsURL", "Password [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get UserName() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 UserName = mvarUserName
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsURL", "UserName [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get Port() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 Port = mvarPort
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsURL", "Port [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get HostName() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 HostName = mvarHostName
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsURL", "HostName [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get Scheme() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 Scheme = mvarScheme
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsURL", "Scheme [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get lScheme() As InetSchemes
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 lScheme = mvarlScheme
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsURL", "lScheme [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Function GetURL(ByVal URL As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 Dim COMPONENTS As URL_COMPONENTS
50020 Dim Buffer As String
50030 Dim BufLen As Long
50040 Const BufSize = 1024
50050
50060 GetURL = False
50070
50080 Buffer = Space$(BufSize)
50090 BufLen = Len(Buffer)
50100 If InternetCanonicalizeUrl(URL, Buffer, BufLen, ICU_BROWSER_MODE) Then
50110     URL = Left(Buffer, BufLen)
50120     With COMPONENTS
50130         .StructSize = Len(COMPONENTS)
50140         .Scheme = Space$(INTERNET_MAX_SCHEME_LENGTH)
50150         .SchemeLength = INTERNET_MAX_SCHEME_LENGTH
50160         .nScheme = InternetSchemeUnknown
50170         .HostName = Space$(INTERNET_MAX_HOST_NAME_LENGTH)
50180         .HostNameLength = INTERNET_MAX_HOST_NAME_LENGTH
50190         .nPort = 0
50200         .UserName = Space$(INTERNET_MAX_USER_NAME_LENGTH)
50210         .UserNameLength = INTERNET_MAX_USER_NAME_LENGTH
50220         .Password = Space$(INTERNET_MAX_PASSWORD_LENGTH)
50230         .PasswordLength = INTERNET_MAX_PASSWORD_LENGTH
50240         .URLPath = Space$(INTERNET_MAX_PATH_LENGTH)
50250         .UrlPathLength = INTERNET_MAX_PATH_LENGTH
50260         .ExtraInfo = Space$(BufSize)
50270         .ExtraInfoLength = BufSize
50280     End With
50290
50300     GetURL = CBool(InternetCrackUrl(URL, Len(URL), ICU_ESCAPE, COMPONENTS))
50310     With COMPONENTS
50320         mvarScheme = Left$(.Scheme, .SchemeLength)
50330         mvarlScheme = .nScheme
50340         mvarHostName = Left$(.HostName, .HostNameLength)
50350         mvarUserName = Left$(.UserName, .UserNameLength)
50360         mvarPassword = Left$(.Password, .PasswordLength)
50370         mvarURLPath = Left$(.URLPath, .UrlPathLength)
50380         mvarURLExtraInfo = Left$(.ExtraInfo, .ExtraInfoLength)
50390         mvarPort = .nPort
50400     End With
50410
50420 End If
50430
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsURL", "GetURL")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
