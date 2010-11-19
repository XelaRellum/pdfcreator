VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Event Update()

' User agent constant.
Private Const scUserAgent = "PDFCreator"
' INTERNET_OPEN_TYPEs
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0                      ' use registry configuration
' Type of service to access.
Private Const INTERNET_SERVICE_HTTP = 3
' Brings the data across the wire even if it locally cached.
Private Const INTERNET_FLAG_RELOAD = &H80000000
' Response Headers to query
Private Const HTTP_QUERY_CONTENT_LENGTH = 5
Private Const HTTP_QUERY_STATUS_CODE = 19
Private Const HTTP_QUERY_FILE_URL = 904
Private Const INTERNET_OPTION_CONNECT_TIMEOUT = 2

' Opens a HTTP session for a given site.
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
(ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, _
ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, _
ByVal lFlags As Long, ByVal lContext As Long) As Long
' Initializes an application's use of the Win32 Internet functions
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
(ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
' Opens an HTTP request handle.
Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" _
(ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, _
ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
' Sends the specified request to the HTTP server.
Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal _
hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, sOptional As _
Any, ByVal lOptionalLength As Long) As Integer
' Reads data from a handle opened by the HttpOpenRequest function.
Private Declare Function InternetReadFile Lib "wininet.dll" _
(ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, _
lNumberOfBytesRead As Long) As Integer
' Closes a single Internet handle or a subtree of Internet handles.
Private Declare Function InternetCloseHandle Lib "wininet.dll" _
(ByVal hInet As Long) As Integer
' Queries for information about an HTTP request.
Private Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" _
(ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, _
ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer

Private Declare Function GetInputState Lib "user32" () As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" _
 (ByVal hInternet As Long, ByVal dwOption As Long, lpBuffer As Long, ByVal dwBufferLength As Long) As Boolean


Private mvarStatus As Byte
Private mvarSize As Long
Private mvarReceivedBytes As Long
Private mvarSpeed As Double
Private mvarStopDownload As Boolean
Private mvarDownloadTime As Long
Private mvarErrornumber As Long
Private mvarErrorDescription As String

Public Function DownloadString(URL As String, Optional TimeOutInMs As Long = 10000) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim hInternetSession    As Long
50020  Dim hInternetConnect    As Long
50030  Dim hHttpOpenRequest    As Long
50040  Dim iRetVal             As Integer
50050  Dim bDoLoop             As Boolean
50060  Dim sReadBuffer         As String * 8192
50070  Dim lNumberOfBytesRead  As Long
50080  Dim sBuffer2            As String
50090  Dim StatusCode          As Long
50100  Dim ContentLength       As Long
50110  Dim cURL                As New clsURL
50120  Dim Start               As Long
50130  Dim Ende                As Long
50140  Dim Duration            As Long
50150  Dim DLStart             As Long
50160  Dim resStr              As String
50170
50180  Dim bRet               As Boolean
50190
50200  If TimeOutInMs < 1000 Then
50210   TimeOutInMs = 1000
50220  End If
50230
50240  StopDownload = False
50250  Size = 0
50260  Speed = 0
50270  ReceivedBytes = 0
50280  Downloadtime = 0
50290  ErrorNumber = 0
50300  ErrorDescription = ""
50310  resStr = ""
50320
50330  status = 1
50340
50350  If cURL.GetURL(URL) And cURL.lScheme = InternetSchemeHttp Then
50360     hInternetSession = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
50370     If CBool(hInternetSession) Then
50380         hInternetConnect = InternetConnect(hInternetSession, cURL.HostName, cURL.Port, cURL.UserName, cURL.Password, INTERNET_SERVICE_HTTP, 0, 0)
50390         If hInternetConnect > 0 Then
50400             hHttpOpenRequest = HttpOpenRequest(hInternetConnect, "GET", cURL.URLPath & cURL.URLExtraInfo, "HTTP/1.1", vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
50410             If CBool(hHttpOpenRequest) Then
50420                 bRet = InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_CONNECT_TIMEOUT, TimeOutInMs, 4)
50430                 iRetVal = HttpSendRequest(hHttpOpenRequest, vbNullString, 0, 0, 0)
50440                 If iRetVal Then
50450                     StatusCode = CLng(GetQueryInfo(ByVal hHttpOpenRequest, HTTP_QUERY_STATUS_CODE))
50460                     If StatusCode <> 200 Then
50470                         ErrorNumber = StatusCode
50480                         ErrorDescription = LanguageStrings.MessagesMsg26
50490                         If StatusCode = 404 Then
50500                             ErrorNumber = 404
50510                             ErrorDescription = LanguageStrings.MessagesMsg27
50520                         End If
50530                         RaiseEvent Update
50540                         Call CloseAll(hHttpOpenRequest, hInternetSession, hInternetConnect, cURL)
50550                         Exit Function
50560                     End If
'50570                     ContentLength = CLng(GetQueryInfo(ByVal hHttpOpenRequest, HTTP_QUERY_CONTENT_LENGTH))
'50580                     Size = ContentLength
'50590                     If Size = 0 Then
'50600                         ErrorNumber = 3
'50610                         ErrorDescription = LanguageStrings.MessagesMsg28
'50620                         RaiseEvent Update
'50630                         Call CloseAll(hHttpOpenRequest, hInternetSession, hInternetConnect, cURL)
'50640                         Exit Function
'50650                     End If
50660                     If StatusCode = 200 Then
50670                         bDoLoop = True
50680                         Duration = 1: DLStart = GetTickCount
50690                         Do
50700                             sReadBuffer = vbNullString
50710                             Start = GetTickCount
50720                             bDoLoop = InternetReadFile(hHttpOpenRequest, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
50730                             Ende = GetTickCount
50740                             resStr = resStr & Left$(sReadBuffer, lNumberOfBytesRead)
50750                             Duration = Duration + Ende - Start
50760                             ReceivedBytes = ReceivedBytes + lNumberOfBytesRead
50770                             Speed = ReceivedBytes / (Duration / 1000)
50780                             Downloadtime = Ende - DLStart
50790                             RaiseEvent Update
50800                             If GetInputState <> 0 Then DoEvents
50810                             If StopDownload = True Then
50820                                 Call CloseAll(hHttpOpenRequest, hInternetSession, hInternetConnect, cURL)
50830                                 Exit Function
50840                             End If
50850                         Loop Until lNumberOfBytesRead < Len(sReadBuffer) Or bDoLoop = 0
50860                     End If
50870                  Else
50880                     ErrorNumber = 1
50890                     ErrorDescription = LanguageStrings.MessagesMsg44
50900                     RaiseEvent Update
50910                 End If
50920             End If
50930         End If
50940     End If
50950    Else
50960     ErrorNumber = 2
50970     ErrorDescription = LanguageStrings.MessagesMsg30
50980     RaiseEvent Update
50990  End If
51000  Call CloseAll(hHttpOpenRequest, hInternetSession, hInternetConnect, cURL)
51010  DownloadString = resStr
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsDownload", "DownloadString")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub CloseAll(hHOR As Long, hIS As Long, hIC As Long, URL As clsURL)
 On Error Resume Next
 Call InternetCloseHandle(hHOR)
 Call InternetCloseHandle(hIS)
 Call InternetCloseHandle(hIC)
 Set URL = Nothing

 DoEvents
 status = 0
End Sub


Public Property Get status() As Byte
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  status = mvarStatus
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsDownload", "status [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Private Property Let status(s As Byte)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  mvarStatus = s
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsDownload", "status [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get Size() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Size = mvarSize
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsDownload", "Size [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Private Property Let Size(s As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  mvarSize = s
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsDownload", "Size [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get ReceivedBytes() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ReceivedBytes = mvarReceivedBytes
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsDownload", "ReceivedBytes [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Private Property Let ReceivedBytes(R As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  mvarReceivedBytes = R
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsDownload", "ReceivedBytes [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get Speed() As Double
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Speed = mvarSpeed
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsDownload", "Speed [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Private Property Let Speed(s As Double)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  mvarSpeed = s
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsDownload", "Speed [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Private Property Get StopDownload() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  StopDownload = mvarStopDownload
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsDownload", "StopDownload [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let StopDownload(s As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  mvarStopDownload = s
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsDownload", "StopDownload [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get Downloadtime() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Downloadtime = mvarDownloadTime
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsDownload", "Downloadtime [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Private Property Let Downloadtime(t As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  mvarDownloadTime = t
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsDownload", "Downloadtime [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get ErrorNumber() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ErrorNumber = mvarErrornumber
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsDownload", "ErrorNumber [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Private Property Let ErrorNumber(e As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  mvarErrornumber = e
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsDownload", "ErrorNumber [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get ErrorDescription() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ErrorDescription = mvarErrorDescription
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsDownload", "ErrorDescription [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Private Property Let ErrorDescription(e As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  mvarErrorDescription = e
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsDownload", "ErrorDescription [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Private Function GetQueryInfo(ByVal hHttpRequest As Long, ByVal iInfoLevel As Long) As Variant
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim sBuffer         As String * 1024
50020  Dim lBufferLength   As Long
50030  lBufferLength = Len(sBuffer)
50040  If CBool(HttpQueryInfo(hHttpRequest, iInfoLevel, ByVal sBuffer, lBufferLength, 0)) Then GetQueryInfo = Left(sBuffer, lBufferLength)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsDownload", "GetQueryInfo")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
