Attribute VB_Name = "modHTTP"
Option Explicit

Private Declare Function InternetOpen Lib "wininet" Alias _
        "InternetOpenA" (ByVal sAgent As String, ByVal _
        lAccessType As Long, ByVal sProxyName As String, ByVal _
        sProxyBypass As String, ByVal lFlags As Long) As Long
        
Private Declare Function InternetCloseHandle Lib "wininet" _
        (ByVal hInet As Long) As Integer
        
Private Declare Function InternetReadFile Lib "wininet" _
        (ByVal hFile As Long, ByVal sBuffer As String, ByVal _
        lNumBytesToRead As Long, lNumberOfBytesRead As Long) _
        As Integer
        
Private Declare Function InternetOpenUrl Lib "wininet" Alias _
        "InternetOpenUrlA" (ByVal hInternetSession As Long, _
        ByVal lpszURL As String, ByVal lpszHeaders As String, _
        ByVal dwHeadersLength As Long, ByVal dwFlags As Long, _
        ByVal dwContext As Long) As Long


Const INTERNET_OPEN_TYPE_DIRECT = 1
Const INTERNET_OPEN_TYPE_PROXY = 3
Const INTERNET_FLAG_RELOAD = &H80000000

Const UserAgent = "PDFCreator"

Public Function HTTPGetFile(strFile As String) As String
    Dim l&, Buffer$, hOpen&, hFile&, Result&
    l = 50000
    Buffer = Space(l)

    hOpen = InternetOpen(UserAgent, INTERNET_OPEN_TYPE_DIRECT, _
                         vbNullString, vbNullString, 0)
  
    hFile = InternetOpenUrl(hOpen, strFile, vbNullString, _
                            ByVal 0&, INTERNET_FLAG_RELOAD, _
                            ByVal 0&)
                            
    Call InternetReadFile(hFile, Buffer, l, Result&)
    Call InternetCloseHandle(hFile)
    Call InternetCloseHandle(hOpen)
    
    HTTPGetFile = Left$(Buffer, Result)
End Function

Public Function HTTPInstallLanguageFile(File As String, Version As String, DownloadURL As String, ProgramPath As String) As Boolean
  Dim strLangFile As String, strFile As String
  
  HTTPInstallLanguageFile = False
   
  If (File = vbNullString) Or (Version = vbNullString) Or (DownloadURL = vbNullString) Or (ProgramPath = vbNullString) Then Exit Function
  ProgramPath = CompletePath(ProgramPath)
  
  If Right$(DownloadURL, 1) <> "/" Then DownloadURL = DownloadURL & "/"

  strLangFile = HTTPGetFile(DownloadURL & Version & "/" & File)
  If InStr(1, strLangFile, "[Common]", vbTextCompare) = 0 Then
    MsgBox LanguageStrings.MessagesMsg37, vbCritical
    Exit Function
  End If
  
  If Not DirExists(ProgramPath) Then
    MsgBox LanguageStrings.MessagesMsg10, vbCritical
    Exit Function
  End If


  strFile = ProgramPath & File

  If FileExists(strFile) Then
    If MsgBox(LanguageStrings.MessagesMsg05, vbYesNo) = vbNo Then
      Exit Function
    End If
  End If
  
  Open strFile For Output As #1
    Print #1, strLangFile
  Close #1
  
  MsgBox LanguageStrings.MessagesMsg38, vbInformation

  HTTPInstallLanguageFile = True
End Function

