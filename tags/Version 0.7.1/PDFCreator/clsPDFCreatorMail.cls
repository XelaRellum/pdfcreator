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
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  StandardEmailClient = iStandardEmailClient
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreatorMail", "StandardEmailClient [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Function Send(ByVal FileName As String, Optional Subject As String = "") As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If UCase$(iStandardEmailClient) = "MICROSOFT OUTLOOK" Then
50020    Send = MSOutlookEmail(FileName, Subject)
50030   Else
50040    Send = VBMapiEmail(FileName, Subject)
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreatorMail", "Send")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub Class_Initialize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  iStandardEmailClient = GetStandardEmailClient
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreatorMail", "Class_Initialize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function GetStandardEmailClient() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry
50020  Set reg = New clsRegistry
50030  reg.hkey = HKEY_LOCAL_MACHINE
50040  reg.KeyRoot = "SOFTWARE\Clients\Mail"
50050  GetStandardEmailClient = reg.GetRegistryValue("")
50060  Set reg = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreatorMail", "GetStandardEmailClient")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function MSOutlookEmail(FileName As String, Optional Subject As String = "") As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50020  Dim olApp As Object, oItem As Object
50030  MSOutlookEmail = 0
50040  Set olApp = CreateObject("Outlook.Application")
50050  If olApp Is Nothing Then
50060   MSOutlookEmail = -1
50070   Exit Function
50080  End If
50090  Set oItem = olApp.CreateItem(0)
50100  With oItem
50110   .To = "": .CC = "": .BCC = "": .Subject = Subject
50120   .Attachments.Add FileName
50130   .Display
50140  End With
50150  Set oItem = Nothing: Set olApp = Nothing
50160  Exit Function
'ErrorHandler:
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
'Exit Function
ErrPtnr_OnError:
50180  MSOutlookEmail = -1
Select Case ErrPtnr.OnError("clsPDFCreatorMail", "MSOutlookEmail")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function VBMapiEmail(FileName As String, Optional Subject As String = "") As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50020  Dim MAPISession As Object, MAPIMessages As Object
50030
50040  VBMapiEmail = 0
50050
50060  Set MAPISession = CreateObject("MSMAPI.MAPISession")
50070  Set MAPIMessages = CreateObject("MSMAPI.MAPIMessages")
50080
50090  With MAPISession
50100   .LogonUI = True
50110   .SignOn
50120  End With
50130  With MAPIMessages
50140   .SessionID = MAPISession.SessionID
50150   .Compose
50160   .AttachmentIndex = 0
50170   .AttachmentPathName = FileName
50180   .MsgSubject = Subject
50190   .Send True
50200  End With
50210  MAPISession.SignOff
50220  Exit Function
'ErrorHandler:
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
'Exit Function
ErrPtnr_OnError:
50240  VBMapiEmail = -1
Select Case ErrPtnr.OnError("clsPDFCreatorMail", "VBMapiEmail")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
