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

' Enumerated internet schemes
Public Enum InetSchemes
   InternetSchemePartial = -2
   InternetSchemeUnknown = -1
   InternetSchemeDefault = 0
   InternetSchemeFtp
   InternetSchemeGopher
   InternetSchemeHttp
   InternetSchemeHttps
   InternetSchemeFile
   InternetSchemeNews
   InternetSchemeMailto
   InternetSchemeSocks
   InternetSchemeFirst = InternetSchemeFtp
   InternetSchemeLast = InternetSchemeSocks
End Enum

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
URLExtraInfo = mvarURLExtraInfo
End Property

Public Property Get URLPath() As String
URLPath = mvarURLPath
End Property

Public Property Get Password() As String
Password = mvarPassword
End Property

Public Property Get UserName() As String
UserName = mvarUserName
End Property

Public Property Get Port() As Long
Port = mvarPort
End Property

Public Property Get HostName() As String
HostName = mvarHostName
End Property

Public Property Get Scheme() As String
Scheme = mvarScheme
End Property

Public Property Get lScheme() As InetSchemes
lScheme = mvarlScheme
End Property

Public Function GetURL(ByVal URL As String) As Boolean
Dim COMPONENTS As URL_COMPONENTS
Dim Buffer As String
Dim BufLen As Long
Const BufSize = 1024

GetURL = False

Buffer = Space$(BufSize)
BufLen = Len(Buffer)
If InternetCanonicalizeUrl(URL, Buffer, BufLen, ICU_BROWSER_MODE) Then
    URL = Left(Buffer, BufLen)
    With COMPONENTS
        .StructSize = Len(COMPONENTS)
        .Scheme = Space$(INTERNET_MAX_SCHEME_LENGTH)
        .SchemeLength = INTERNET_MAX_SCHEME_LENGTH
        .nScheme = InternetSchemeUnknown
        .HostName = Space$(INTERNET_MAX_HOST_NAME_LENGTH)
        .HostNameLength = INTERNET_MAX_HOST_NAME_LENGTH
        .nPort = 0
        .UserName = Space$(INTERNET_MAX_USER_NAME_LENGTH)
        .UserNameLength = INTERNET_MAX_USER_NAME_LENGTH
        .Password = Space$(INTERNET_MAX_PASSWORD_LENGTH)
        .PasswordLength = INTERNET_MAX_PASSWORD_LENGTH
        .URLPath = Space$(INTERNET_MAX_PATH_LENGTH)
        .UrlPathLength = INTERNET_MAX_PATH_LENGTH
        .ExtraInfo = Space$(BufSize)
        .ExtraInfoLength = BufSize
    End With

    GetURL = CBool(InternetCrackUrl(URL, Len(URL), ICU_ESCAPE, COMPONENTS))
    With COMPONENTS
        mvarScheme = Left$(.Scheme, .SchemeLength)
        mvarlScheme = .nScheme
        mvarHostName = Left$(.HostName, .HostNameLength)
        mvarUserName = Left$(.UserName, .UserNameLength)
        mvarPassword = Left$(.Password, .PasswordLength)
        mvarURLPath = Left$(.URLPath, .UrlPathLength)
        mvarURLExtraInfo = Left$(.ExtraInfo, .ExtraInfoLength)
        mvarPort = .nPort
    End With

End If

End Function
