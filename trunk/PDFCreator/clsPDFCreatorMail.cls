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

Private iStandardEmailClient As String

Public Property Get StandardEmailClient() As String
 StandardEmailClient = iStandardEmailClient
End Property

Public Function Send(ByVal FileName As String, Optional Subject As String = "") As Long
 If UCase$(iStandardEmailClient) = "MICROSOFT OUTLOOK" Then
   Send = MSOutlookEmail(FileName, Subject)
  Else
   Send = VBMapiEmail(FileName, Subject)
 End If
End Function

Private Sub Class_Initialize()
 iStandardEmailClient = GetStandardEmailClient
End Sub

Private Function GetStandardEmailClient() As String
 Dim reg As clsRegistry
 Set reg = New clsRegistry
 reg.hkey = HKEY_LOCAL_MACHINE
 reg.KeyRoot = "SOFTWARE\Clients\Mail"
 GetStandardEmailClient = reg.GetRegistryValue("")
 Set reg = Nothing
End Function

Private Function MSOutlookEmail(FileName As String, Optional Subject As String = "") As Long
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
  .Attachments.Add FileName
  .Display
 End With
 Set oItem = Nothing: Set olApp = Nothing
 Exit Function
ErrorHandler:
 MSOutlookEmail = -1
End Function

Private Function VBMapiEmail(FileName As String, Optional Subject As String = "") As Long
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
  .AttachmentPathName = FileName
  .MsgSubject = Subject
  .Send True
 End With
 MAPISession.SignOff
 Exit Function
ErrorHandler:
 VBMapiEmail = -1
End Function
