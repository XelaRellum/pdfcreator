VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPDFCreatorMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal numBytes As Long)
                    
Private iStandardEmailClient As String

Public Property Get StandardEmailClient() As String
 StandardEmailClient = iStandardEmailClient
End Property

Public Function Send(ByVal Filename As String, Optional Subject As String = "") As Long
 If UCase$(iStandardEmailClient) = "MICROSOFT OUTLOOK" Then
   Send = MSOutlookEmail(Filename, Subject)
  Else
   Send = VBMapiEmail(Filename, Subject)
 End If
End Function

Private Sub Class_Initialize()
 iStandardEmailClient = GetStandardEmailClient
End Sub

Private Function GetStandardEmailClient() As String
 Const HKEY_LOCAL_MACHINE = &H80000002, REG_OPTION_NON_VOLATILE = 0
 
 Const ERROR_MORE_DATA = 234, ERROR_KEY_NOT_FOUND = &H2, REG_SZ = &H1

 Const SYNCHRONIZE = &H100000
 Const READ_CONTROL = &H20000
 Const STANDARD_RIGHTS_READ = (READ_CONTROL)
 Const KEY_ENUMERATE_SUB_KEYS = &H8
 Const KEY_QUERY_VALUE = &H1
 Const KEY_NOTIFY = &H10
 Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
 
 Dim ValueName As String, handle As Long, resString As String, _
  resBinary() As Byte, Length As Long, RetVal As Long, valueType As Long
    
 If RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SOFTWARE\Clients\Mail", REG_OPTION_NON_VOLATILE, KEY_READ, handle) Then
  Exit Function
 End If
 
 Length = 1024
 ReDim resBinary(0 To Length - 1) As Byte
 RetVal = RegQueryValueEx(handle, "", 0, valueType, resBinary(0), Length)
 If RetVal = ERROR_MORE_DATA Then
  ReDim resBinary(0 To Length - 1) As Byte
  RetVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), Length)
 End If
 If RetVal = ERROR_KEY_NOT_FOUND Then
  RegCloseKey (handle)
  Exit Function
 End If

 Select Case valueType
  Case REG_SZ
   If Length <> 0 Then
    resString = Space$(Length - 1)
    CopyMemory ByVal resString, resBinary(0), Length - 1
    GetStandardEmailClient = resString
   End If
   Case Else
  End Select
  RegCloseKey (handle)
End Function

Private Function MSOutlookEmail(Filename As String, Optional Subject As String = "") As Long
 On Local Error GoTo ErrorHandler
 Dim olApp As Object, oItem As Object
 MSOutlookEmail = 0
 Set olApp = CreateObject("Outlook.Application")
 If olApp Is Nothing Then
  MSOutlookEmail = -1
  Exit Function
 End If
 Set oItem = olApp.CreateItem(0)
 With oItem
  .To = "": .CC = "": .BCC = "": .Subject = Subject
  .Body = "": .Attachments.Add Filename
  .Display
 End With
 Set oItem = Nothing: Set olApp = Nothing
 Exit Function
ErrorHandler:
 MSOutlookEmail = -1
End Function

Private Function VBMapiEmail(Filename As String, Optional Subject As String = "") As Long
 On Local Error GoTo ErrorHandler
 Dim MAPISession As Object, MAPIMessages As Object
 
 VBMapiEmail = 0
 
 Set MAPISession = CreateObject("MSMAPI.MAPISession")
 Set MAPIMessages = CreateObject("MSMAPI.MAPIMessages")
 
 With MAPISession
  .LogonUI = True
  .SignOn
 End With
 With MAPIMessages
  .SessionID = MAPISession.SessionID
  .Compose
  .AttachmentIndex = 0
  .AttachmentPathName = Filename
  .MsgSubject = Subject
  .Send True
 End With
 MAPISession.SignOff
 Exit Function
ErrorHandler:
 VBMapiEmail = -1
End Function

