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
Private Const scUserAgent = "http sample"
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
Private Declare Function GetTickCount Lib "Kernel32" () As Long

Private mvarStatus As Byte
Private mvarSize As Long
Private mvarReceivedBytes As Long
Private mvarSpeed As Double
Private mvarStopDownload As Boolean
Private mvarDownloadTime As Long
Private mvarErrornumber As Long
Private mvarErrorDescription As String

Public Function DownloadString(URL As String) As String
 Dim hInternetSession    As Long
 Dim hInternetConnect    As Long
 Dim hHttpOpenRequest    As Long
 Dim iRetVal             As Integer
 Dim bDoLoop             As Boolean
 Dim sReadBuffer         As String * 8192
 Dim lNumberOfBytesRead  As Long
 Dim sBuffer2            As String
 Dim StatusCode          As Long
 Dim ContentLength       As Long
 Dim cURL                As New clsURL
 Dim Start               As Long
 Dim Ende                As Long
 Dim Duration            As Long
 Dim DLStart             As Long
 Dim resStr              As String

 StopDownload = False
 Size = 0
 Speed = 0
 ReceivedBytes = 0
 Downloadtime = 0
 ErrorNumber = 0
 ErrorDescription = ""
 resStr = ""

 Status = 1

 If cURL.GetURL(URL) And cURL.lScheme = InternetSchemeHttp Then
    hInternetSession = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    If CBool(hInternetSession) Then
        hInternetConnect = InternetConnect(hInternetSession, cURL.HostName, cURL.Port, cURL.UserName, cURL.Password, INTERNET_SERVICE_HTTP, 0, 0)
        If hInternetConnect > 0 Then
            hHttpOpenRequest = HttpOpenRequest(hInternetConnect, "GET", cURL.URLPath & cURL.URLExtraInfo, "HTTP/1.1", vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
            If CBool(hHttpOpenRequest) Then
                iRetVal = HttpSendRequest(hHttpOpenRequest, vbNullString, 0, 0, 0)
                If iRetVal Then
                    StatusCode = CLng(GetQueryInfo(ByVal hHttpOpenRequest, HTTP_QUERY_STATUS_CODE))
                    If StatusCode <> 200 Then
                        ErrorNumber = StatusCode
                        ErrorDescription = LanguageStrings.MessagesMsg26
                        If StatusCode = 404 Then
                            ErrorNumber = 404
                            ErrorDescription = LanguageStrings.MessagesMsg27
                        End If
                        RaiseEvent Update
                        Call CloseAll(hHttpOpenRequest, hInternetSession, hInternetConnect, cURL)
                        Exit Function
                    End If
                    ContentLength = CLng(GetQueryInfo(ByVal hHttpOpenRequest, HTTP_QUERY_CONTENT_LENGTH))
                    Size = ContentLength
                    If Size = 0 Then
                        ErrorNumber = 3
                        ErrorDescription = LanguageStrings.MessagesMsg28
                        RaiseEvent Update
                        Call CloseAll(hHttpOpenRequest, hInternetSession, hInternetConnect, cURL)
                        Exit Function
                    End If
                    If StatusCode = 200 Then
                        bDoLoop = True
                        Duration = 1: DLStart = GetTickCount
                        Do
                            sReadBuffer = vbNullString
                            Start = GetTickCount
                            bDoLoop = InternetReadFile(hHttpOpenRequest, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
                            Ende = GetTickCount
                            resStr = resStr & Left$(sReadBuffer, lNumberOfBytesRead)
                            Duration = Duration + Ende - Start
                            ReceivedBytes = ReceivedBytes + lNumberOfBytesRead
                            Speed = ReceivedBytes / (Duration / 1000)
                            Downloadtime = Ende - DLStart
                            RaiseEvent Update
                            If GetInputState <> 0 Then DoEvents
                            If StopDownload = True Then
                                Call CloseAll(hHttpOpenRequest, hInternetSession, hInternetConnect, cURL)
                                Exit Function
                            End If
                        Loop Until lNumberOfBytesRead < Len(sReadBuffer) Or bDoLoop = 0
                    End If
                 Else
                    ErrorNumber = 1
                    ErrorDescription = LanguageStrings.MessagesMsg29
                    RaiseEvent Update
                End If
            End If
        End If
    End If
   Else
    ErrorNumber = 2
    ErrorDescription = LanguageStrings.MessagesMsg30
    RaiseEvent Update
 End If
 Call CloseAll(hHttpOpenRequest, hInternetSession, hInternetConnect, cURL)
 DownloadString = resStr
End Function

Private Sub CloseAll(hHOR As Long, hIS As Long, hIC As Long, URL As clsURL)
 On Error Resume Next
 Call InternetCloseHandle(hHOR)
 Call InternetCloseHandle(hIS)
 Call InternetCloseHandle(hIC)
 Set URL = Nothing

 DoEvents
 Status = 0
End Sub


Public Property Get Status() As Byte
 Status = mvarStatus
End Property

Private Property Let Status(s As Byte)
 mvarStatus = s
End Property

Public Property Get Size() As Long
 Size = mvarSize
End Property

Private Property Let Size(s As Long)
 mvarSize = s
End Property

Public Property Get ReceivedBytes() As Long
 ReceivedBytes = mvarReceivedBytes
End Property

Private Property Let ReceivedBytes(r As Long)
 mvarReceivedBytes = r
End Property

Public Property Get Speed() As Double
 Speed = mvarSpeed
End Property

Private Property Let Speed(s As Double)
 mvarSpeed = s
End Property

Private Property Get StopDownload() As Boolean
 StopDownload = mvarStopDownload
End Property

Public Property Let StopDownload(s As Boolean)
 mvarStopDownload = s
End Property

Public Property Get Downloadtime() As Long
 Downloadtime = mvarDownloadTime
End Property

Private Property Let Downloadtime(t As Long)
 mvarDownloadTime = t
End Property

Public Property Get ErrorNumber() As Long
 ErrorNumber = mvarErrornumber
End Property

Private Property Let ErrorNumber(e As Long)
 mvarErrornumber = e
End Property

Public Property Get ErrorDescription() As String
 ErrorDescription = mvarErrorDescription
End Property

Private Property Let ErrorDescription(e As String)
 mvarErrorDescription = e
End Property

Private Function GetQueryInfo(ByVal hHttpRequest As Long, ByVal iInfoLevel As Long) As Variant
 Dim sBuffer         As String * 1024
 Dim lBufferLength   As Long
 lBufferLength = Len(sBuffer)
 If CBool(HttpQueryInfo(hHttpRequest, iInfoLevel, ByVal sBuffer, lBufferLength, 0)) Then GetQueryInfo = Left(sBuffer, lBufferLength)
End Function
