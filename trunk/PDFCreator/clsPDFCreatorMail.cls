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

Public Function Send(ByVal Filename As String, Optional Subject As String = "") As Long
 Select Case UCase$(iStandardEmailClient)
  Case "MICROSOFT OUTLOOK"
   Send = MSOutlookEmail(Filename, Subject)
  Case "PEGASUS MAIL"
   Send = PegasusEmail(Filename, Subject)
  Case "THE BAT!"
   Send = TheBatEmail(Filename, Subject)
  Case "FOXMAIL"
   Send = FoxEmail(Filename, Subject)
  Case "SYLPHEED"
   Send = SylpheedClaws(Filename, Subject)
  Case Else
   Send = VBMapiEmail(Filename, Subject)
 End Select
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

Private Function MSOutlookEmail(Filename As String, Optional Subject As String = "") As Long
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
  .Attachments.Add Filename
  .Display
 End With
 Set oItem = Nothing: Set olApp = Nothing
End Function

Private Function VBMapiEmail(Filename As String, Optional Subject As String = "") As Long
 On Error GoTo ErrorHandler
 Dim MAPISess As MAPISession, MAPIMess As MAPIMessages

 VBMapiEmail = 0

 Set MAPISess = New MAPISession
 Set MAPIMess = New MAPIMessages

 With MAPISess
  .LogonUI = True
  .SignOn
 End With
 With MAPIMess
  .SessionID = MAPISess.SessionID
  .Compose
  If Len(Trim$(Filename)) > 0 Then
   .AttachmentIndex = 0
   .AttachmentPathName = Filename
   .AttachmentName = Filename
  End If
  .MsgSubject = Subject
  .Send True
 End With
 MAPISess.SignOff
 Set MAPIMess = Nothing
 Set MAPISess = Nothing
 Exit Function
ErrorHandler:
 MsgBox LanguageStrings.MessagesMsg20 & vbCrLf & vbCrLf & _
  "Line " & Erl & ": [" & Err.number & "] " & Err.Description
 Err.Clear
End Function

Private Function PegasusEmail(Filename As String, Optional Subject As String = "") As Long
 Dim cReg As clsRegistry, PegasusBaseDir As String, WsendToPath As String
 PegasusEmail = 0
 Set cReg = New clsRegistry

 cReg.hkey = HKEY_CLASSES_ROOT
 cReg.KeyRoot = "Software\Pegasus Mail\BaseDir"
 PegasusBaseDir = cReg.GetRegistryValue("")
 WsendToPath = CompletePath(PegasusBaseDir) & "wsendto.exe"
 If Dir(WsendToPath) <> "" Then
   Shell """" & WsendToPath & """ """ & Filename & """"
  Else
   MsgBox LanguageStrings.MessagesMsg20
   PegasusEmail = -1
 End If
End Function

Private Function TheBatEmail(Filename As String, Optional Subject As String = "") As Long
 Dim cReg As clsRegistry, ExePath As String
 TheBatEmail = 0
 Set cReg = New clsRegistry

 cReg.hkey = HKEY_CURRENT_USER
 cReg.KeyRoot = "Software\Rit\The Bat!"
 ExePath = cReg.GetRegistryValue("EXE path")
 If Dir(ExePath) <> "" Then
   Shell """" & ExePath & """ """ & Filename & """"
  Else
   MsgBox LanguageStrings.MessagesMsg20
   TheBatEmail = -1
 End If
End Function

Private Function FoxEmail(Filename As String, Optional Subject As String = "") As Long
 Dim cReg As clsRegistry, ExePath As String
 FoxEmail = 0
 Set cReg = New clsRegistry

 cReg.hkey = HKEY_LOCAL_MACHINE
 cReg.KeyRoot = "Software\Aerofox\Foxmail"
 ExePath = cReg.GetRegistryValue("Executable")
 If Dir(ExePath) <> "" Then
   Shell """" & ExePath & """ /min """ & Filename & """"
  Else
   MsgBox LanguageStrings.MessagesMsg20
   FoxEmail = -1
 End If
End Function

Private Function SylpheedClaws(Filename As String, Optional Subject As String = "") As Long
 Dim cReg As clsRegistry, ExePath As String
 SylpheedClaws = 0
 Set cReg = New clsRegistry

 cReg.hkey = HKEY_LOCAL_MACHINE
 cReg.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Uninstall\Sylpheed-Claws_is1"
 ExePath = cReg.GetRegistryValue("Inno Setup: App Path")
 ExePath = CompletePath(ExePath) & "bin\sylpheed.exe"
 If Dir(ExePath) <> "" Then
   Shell """" & ExePath & """ --attach """ & Filename & """"
  Else
   MsgBox LanguageStrings.MessagesMsg20
   SylpheedClaws = -1
 End If
End Function
