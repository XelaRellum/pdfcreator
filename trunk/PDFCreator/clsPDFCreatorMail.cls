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

Public Function Send(ByVal Filename As String, Optional Subject As String = "") As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case UCase$(iStandardEmailClient)
        Case "MICROSOFT OUTLOOK"
50030    Send = MSOutlookEmail(Filename, Subject)
50040   Case "PEGASUS MAIL"
50050    Send = PegasusEmail(Filename, Subject)
50060   Case "THE BAT!"
50070    Send = TheBatEmail(Filename, Subject)
50080   Case "FOXMAIL"
50090    Send = FoxEmail(Filename, Subject)
50100   Case "SYLPHEED"
50110    Send = SylpheedClaws(Filename, Subject)
50120   Case "OUTLOOK EXPRESS"
50130    Send = VBMapiEmail(Filename, Subject)
50140   Case Else
50150    Send = SendToMailRecipientViaSendmailDll(Filename)
50160  End Select
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

Private Function MSOutlookEmail(Filename As String, Optional Subject As String = "") As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim olApp As Object, oItem As Object
50020  MSOutlookEmail = 0
50030  Set olApp = CreateObject("Outlook.Application")
50040  If olApp Is Nothing Then
50050   MSOutlookEmail = -1
50060   Exit Function
50070  End If
50080  Set oItem = olApp.CreateItem(0)
50090  With oItem
50100   .To = "": .CC = "": .BCC = "": .Subject = Subject
50110   .Attachments.Add Filename
50120   .Display
50130  End With
50140  Set oItem = Nothing: Set olApp = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreatorMail", "MSOutlookEmail")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim cReg As clsRegistry, PegasusBaseDir As String, WsendToPath As String
50020  PegasusEmail = 0
50030  Set cReg = New clsRegistry
50040
50050  cReg.hkey = HKEY_CLASSES_ROOT
50060  cReg.KeyRoot = "Software\Pegasus Mail\BaseDir"
50070  PegasusBaseDir = cReg.GetRegistryValue("")
50080  WsendToPath = CompletePath(PegasusBaseDir) & "wsendto.exe"
50090  If Dir(WsendToPath) <> vbNullString Then
50100    Shell """" & WsendToPath & """ """ & Filename & """"
50110   Else
50120    MsgBox LanguageStrings.MessagesMsg20
50130    PegasusEmail = -1
50140  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreatorMail", "PegasusEmail")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function TheBatEmail(Filename As String, Optional Subject As String = "") As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim cReg As clsRegistry, ExePath As String
50020  TheBatEmail = 0
50030  Set cReg = New clsRegistry
50040
50050  cReg.hkey = HKEY_CURRENT_USER
50060  cReg.KeyRoot = "Software\Rit\The Bat!"
50070  ExePath = cReg.GetRegistryValue("EXE path")
50080  If Dir(ExePath) <> vbNullString Then
50090    Shell """" & ExePath & """ """ & Filename & """"
50100   Else
50110    MsgBox LanguageStrings.MessagesMsg20
50120    TheBatEmail = -1
50130  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreatorMail", "TheBatEmail")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function FoxEmail(Filename As String, Optional Subject As String = "") As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim cReg As clsRegistry, ExePath As String
50020  FoxEmail = 0
50030  Set cReg = New clsRegistry
50040
50050  cReg.hkey = HKEY_LOCAL_MACHINE
50060  cReg.KeyRoot = "Software\Aerofox\Foxmail"
50070  ExePath = cReg.GetRegistryValue("Executable")
50080  If Dir(ExePath) <> vbNullString Then
50090    Shell """" & ExePath & """ /min """ & Filename & """"
50100   Else
50110    MsgBox LanguageStrings.MessagesMsg20
50120    FoxEmail = -1
50130  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreatorMail", "FoxEmail")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function SylpheedClaws(Filename As String, Optional Subject As String = "") As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim cReg As clsRegistry, ExePath As String
50020  SylpheedClaws = 0
50030  Set cReg = New clsRegistry
50040
50050  cReg.hkey = HKEY_LOCAL_MACHINE
50060  cReg.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Uninstall\Sylpheed-Claws_is1"
50070  ExePath = cReg.GetRegistryValue("Inno Setup: App Path")
50080  ExePath = CompletePath(ExePath) & "bin\sylpheed.exe"
50090  If Dir(ExePath) <> vbNullString Then
50100    Shell """" & ExePath & """ --attach """ & Filename & """"
50110   Else
50120    MsgBox LanguageStrings.MessagesMsg20
50130    SylpheedClaws = -1
50140  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreatorMail", "SylpheedClaws")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function


Private Function SendToMailRecipientViaSendmailDll(ByVal Filename As String) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tIID_IDropTarget As UUID, tCLSID_SendMail As UUID, oSendMail As IDropTarget, _
  oDO As IDataObject, lRes As Long
50030
50040  SendToMailRecipientViaSendmailDll = -1
50050  CLSIDFromString "{00000122-0000-0000-C000-000000000046}", tIID_IDropTarget
50060  CLSIDFromString "{9E56BE60-C50F-11CF-9A2C-00A0C90A90CE}", tCLSID_SendMail
50070  lRes = CoCreateInstance(tCLSID_SendMail, Nothing, CLSCTX_INPROC_SERVER, _
  tIID_IDropTarget, oSendMail)
50090
50100  If lRes = S_OK Then
50110    Set oDO = GetFileDataObject(Filename)
50120    oSendMail.DragEnter oDO, vbKeyLButton, 0, 0, DROPEFFECT_COPY
50130    oSendMail.Drop oDO, vbKeyLButton, 0, 0, DROPEFFECT_COPY
50140    SendToMailRecipientViaSendmailDll = 0
50150   Else
50160    Err.Raise lRes
50170  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreatorMail", "SendToMailRecipientViaSendmailDll")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function GetFileDataObject(ByVal Filename As String) As IDataObject
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tIID_IDataObject As UUID, tIID_IShellFolder As UUID, _
  oDesktop As IShellFolder, oParent As IShellFolder, oUnk As IUnknown, _
  sFolder As String, lPidl As Long, lPtr As Long
50040
50050  CLSIDFromString "{0000010e-0000-0000-C000-000000000046}", tIID_IDataObject
50060  CLSIDFromString IIDSTR_IShellFolder, tIID_IShellFolder
50070
50080  sFolder = Left$(Filename, InStrRev(Filename, "\") - 1)
50090  Filename = Mid$(Filename, Len(sFolder) + 2)
50100  If Right$(sFolder, 1) = ":" Then
50110   sFolder = sFolder + "\"
50120  End If
50130  Set oDesktop = SHGetDesktopFolder
50140  oDesktop.ParseDisplayName 0, 0, StrPtr(sFolder), lPtr, lPidl, 0
50150  oDesktop.BindToObject lPidl, 0, tIID_IShellFolder, lPtr
50160  MoveMemory oParent, lPtr, 4&
50170  CoTaskMemFree lPidl
50180  oParent.ParseDisplayName 0, 0, StrPtr(Filename), 0, lPidl, 0
50190  lPtr = oParent.GetUIObjectOf(0, 1, lPidl, tIID_IDataObject, 0)
50200  MoveMemory oUnk, lPtr, 4&
50210  CoTaskMemFree lPidl
50220  Set GetFileDataObject = oUnk
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsPDFCreatorMail", "GetFileDataObject")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function


