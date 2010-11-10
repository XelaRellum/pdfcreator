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

Public Function DownloadString(URL As String) As String
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
50180  Dim lTimeoutInMs       As Long
50190  Dim bRet               As Boolean
50200
50210  StopDownload = False
50220  Size = 0
50230  Speed = 0
50240  ReceivedBytes = 0
50250  Downloadtime = 0
50260  ErrorNumber = 0
50270  ErrorDescription = ""
50280  resStr = ""
50290
50300  status = 1
50310  lTimeoutInMs = 1000
50320
50330  If cURL.GetURL(URL) And cURL.lScheme = InternetSchemeHttp Then
50340     hInternetSession = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
50350     If CBool(hInternetSession) Then
50360         hInternetConnect = InternetConnect(hInternetSession, cURL.HostName, cURL.Port, cURL.UserName, cURL.Password, INTERNET_SERVICE_HTTP, 0, 0)
50370         If hInternetConnect > 0 Then
50380             hHttpOpenRequest = HttpOpenRequest(hInternetConnect, "GET", cURL.URLPath & cURL.URLExtraInfo, "HTTP/1.1", vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
50390             If CBool(hHttpOpenRequest) Then
50400                 bRet = InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_CONNECT_TIMEOUT, lTimeoutInMs, 4)
50410                 iRetVal = HttpSendRequest(hHttpOpenRequest, vbNullString, 0, 0, 0)
50420                 If iRetVal Then
50430                     StatusCode = CLng(GetQueryInfo(ByVal hHttpOpenRequest, HTTP_QUERY_STATUS_CODE))
50440                     If StatusCode <> 200 Then
50450                         ErrorNumber = StatusCode
50460                         ErrorDescription = LanguageStrings.MessagesMsg26
50470                         If StatusCode = 404 Then
50480                             ErrorNumber = 404
50490                             ErrorDescription = LanguageStrings.MessagesMsg27
50500                         End If
50510                         RaiseEvent Update
50520                         Call CloseAll(hHttpOpenRequest, hInternetSession, hInternetConnect, cURL)
50530                         Exit Function
50540                     End If
50550                     ContentLength = CLng(GetQueryInfo(ByVal hHttpOpenRequest, HTTP_QUERY_CONTENT_LENGTH))
50560                     Size = ContentLength
50570                     If Size = 0 Then
50580                         ErrorNumber = 3
50590                         ErrorDescription = LanguageStrings.MessagesMsg28
50600                         RaiseEvent Update
50610                         Call CloseAll(hHttpOpenRequest, hInternetSession, hInternetConnect, cURL)
50620                         Exit Function
50630                     End If
50640                     If StatusCode = 200 Then
50650                         bDoLoop = True
50660                         Duration = 1: DLStart = GetTickCount
50670                         Do
50680                             sReadBuffer = vbNullString
50690                             Start = GetTickCount
50700                             bDoLoop = InternetReadFile(hHttpOpenRequest, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
50710                             Ende = GetTickCount
50720                             resStr = resStr & Left$(sReadBuffer, lNumberOfBytesRead)
50730                             Duration = Duration + Ende - Start
50740                             ReceivedBytes = ReceivedBytes + lNumberOfBytesRead
50750                             Speed = ReceivedBytes / (Duration / 1000)
50760                             Downloadtime = Ende - DLStart
50770                             RaiseEvent Update
50780                             If GetInputState <> 0 Then DoEvents
50790                             If StopDownload = True Then
50800                                 Call CloseAll(hHttpOpenRequest, hInternetSession, hInternetConnect, cURL)
50810                                 Exit Function
50820                             End If
50830                         Loop Until lNumberOfBytesRead < Len(sReadBuffer) Or bDoLoop = 0
50840                     End If
50850                  Else
50860                     ErrorNumber = 1
50870                     ErrorDescription = LanguageStrings.MessagesMsg44
50880                     RaiseEvent Update
50890                 End If
50900             End If
50910         End If
50920     End If
50930    Else
50940     ErrorNumber = 2
50950     ErrorDescription = LanguageStrings.MessagesMsg30
50960     RaiseEvent Update
50970  End If
50980  Call CloseAll(hHttpOpenRequest, hInternetSession, hInternetConnect, cURL)
50990  DownloadString = resStr
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
