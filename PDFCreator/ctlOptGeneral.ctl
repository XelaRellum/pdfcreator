VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl ctlOptGeneral 
   ClientHeight    =   5850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13110
   ScaleHeight     =   5850
   ScaleWidth      =   13110
   ToolboxBitmap   =   "ctlOptGeneral.ctx":0000
   Begin PDFCreator.dmFrame dmFraCheckUpdate 
      Height          =   1065
      Left            =   6720
      TabIndex        =   23
      Top             =   4560
      Visible         =   0   'False
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   1879
      Caption         =   "Check Update"
      Caption3D       =   2
      BarColorFrom    =   16744576
      BarColorTo      =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextShaddowColor=   12582912
      Begin VB.CommandButton cmdCheckNow 
         Caption         =   "Check now"
         Height          =   315
         Left            =   4200
         TabIndex        =   26
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox cmbUpdateInterval 
         Height          =   315
         ItemData        =   "ctlOptGeneral.ctx":0312
         Left            =   120
         List            =   "ctlOptGeneral.ctx":0314
         Style           =   2  'Dropdown-Liste
         TabIndex        =   24
         Top             =   600
         Width           =   3870
      End
      Begin VB.Label lblUpdateInterval 
         AutoSize        =   -1  'True
         Caption         =   "Update interval"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   1080
      End
   End
   Begin PDFCreator.dmFrame dmFraBrowserAddOn 
      Height          =   1185
      Left            =   6720
      TabIndex        =   21
      Top             =   3240
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   2090
      Caption         =   "Browser Add On"
      Caption3D       =   2
      BarColorFrom    =   16744576
      BarColorTo      =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextShaddowColor=   12582912
      Begin VB.CommandButton cmdInstallBrowserAddOn 
         Caption         =   "&Install Browser Add On"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   2580
      End
   End
   Begin PDFCreator.dmFrame dmFraShellIntegration 
      Height          =   1065
      Left            =   120
      TabIndex        =   18
      Top             =   4560
      Visible         =   0   'False
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   1879
      Caption         =   "Shell integration"
      Caption3D       =   2
      BarColorFrom    =   16744576
      BarColorTo      =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextShaddowColor=   12582912
      Enabled         =   0   'False
      Begin VB.CommandButton cmdShellintegration 
         Caption         =   "Integrate PDFCreator into shell"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   420
         Width           =   2910
      End
      Begin VB.CommandButton cmdShellintegration 
         Caption         =   "Integrate PDFCreator into shell"
         Height          =   495
         Index           =   1
         Left            =   3150
         TabIndex        =   20
         Top             =   420
         Width           =   2910
      End
   End
   Begin PDFCreator.dmFrame dmFraProgGeneral2 
      Height          =   2745
      Left            =   6720
      TabIndex        =   12
      Top             =   360
      Visible         =   0   'False
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   4842
      Caption         =   "General 2"
      Caption3D       =   2
      BarColorFrom    =   16744576
      BarColorTo      =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextShaddowColor=   12582912
      Begin VB.CommandButton cmdAsso 
         Caption         =   "&Associate PDFCreator with Postscript files"
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   2580
      End
      Begin VB.CheckBox chkShowAnimation 
         Appearance      =   0  '2D
         Caption         =   "Show animation"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   120
         TabIndex        =   17
         Top             =   2220
         Width           =   5775
      End
      Begin VB.ComboBox cmbOptionsDesign 
         Height          =   315
         ItemData        =   "ctlOptGeneral.ctx":0316
         Left            =   120
         List            =   "ctlOptGeneral.ctx":0318
         Style           =   2  'Dropdown-Liste
         TabIndex        =   16
         Top             =   1620
         Width           =   3870
      End
      Begin PDFCreator.Line3D Line3D1 
         Height          =   30
         Index           =   3
         Left            =   105
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1200
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   53
         LineType        =   1
         3DHighlight     =   -2147483628
         3DShadow        =   -2147483632
         DrawStyle       =   0
      End
      Begin VB.Label lblOptionsDesign 
         AutoSize        =   -1  'True
         Caption         =   "Frame color of the setting dialog"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1380
         Width           =   2250
      End
   End
   Begin PDFCreator.dmFrame dmFraProgGeneral1 
      Height          =   4110
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   7250
      Caption         =   "General 1"
      Caption3D       =   2
      BarColorFrom    =   16744576
      BarColorTo      =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextShaddowColor=   12582912
      Begin VB.CheckBox chkNoConfirmMessageSwitchingDefaultprinter 
         Appearance      =   0  '2D
         Caption         =   "No confirm message switching PDFCreator temporarly as default printer."
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   2640
         Width           =   5775
      End
      Begin VB.ComboBox cmbSendMailMethod 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown-Liste
         TabIndex        =   11
         Top             =   3675
         Width           =   2580
      End
      Begin VB.CheckBox chkNoProcessingAtStartup 
         Appearance      =   0  '2D
         Caption         =   "No processing at startup"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   5775
      End
      Begin VB.CommandButton cmdTestpage 
         Caption         =   "&Print testpage"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2580
      End
      Begin PDFCreator.Line3D Line3D1 
         Height          =   30
         Index           =   0
         Left            =   105
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1200
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   53
         LineType        =   1
         3DHighlight     =   -2147483628
         3DShadow        =   -2147483632
         DrawStyle       =   0
      End
      Begin MSComctlLib.Slider sldProcessPriority 
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   1
         Max             =   3
         SelStart        =   1
         Value           =   1
      End
      Begin PDFCreator.Line3D Line3D1 
         Height          =   30
         Index           =   1
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2160
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   53
         LineType        =   1
         3DHighlight     =   -2147483628
         3DShadow        =   -2147483632
         DrawStyle       =   0
      End
      Begin PDFCreator.Line3D Line3D1 
         Height          =   30
         Index           =   2
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3240
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   53
         LineType        =   1
         3DHighlight     =   -2147483628
         3DShadow        =   -2147483632
         DrawStyle       =   0
      End
      Begin VB.Label lblSendMailMethod 
         AutoSize        =   -1  'True
         Caption         =   "Methode to send an email"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   3465
         Width           =   1830
      End
      Begin VB.Label lblProcessPriority 
         AutoSize        =   -1  'True
         Caption         =   "Processpriority: Normal"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1605
      End
   End
   Begin MSComctlLib.TabStrip tbstrProgGeneral 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   9975
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ctlOptGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private ControlsEnabled As Boolean

Public Sub SetControlsEnabled(value As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ControlsEnabled = value
50020
50030  chkNoConfirmMessageSwitchingDefaultprinter.Enabled = value
50040  chkNoProcessingAtStartup.Enabled = value
50050  cmbOptionsDesign.Enabled = value
50060  sldProcessPriority.Enabled = value
50070  cmbSendMailMethod.Enabled = value
50080  chkShowAnimation.Enabled = value
50090  cmbUpdateInterval.Enabled = value
50100
50110  cmdTestpage.Enabled = value
50120  lblProcessPriority.Enabled = value
50130  lblSendMailMethod.Enabled = value
50140  cmdShellintegration(0).Enabled = value
50150  cmdShellintegration(1).Enabled = value
50160  cmdAsso.Enabled = value
50170  lblOptionsDesign.Enabled = value
50180  cmdInstallBrowserAddOn.Enabled = value
50190  lblUpdateInterval.Enabled = value
50200  cmdCheckNow.Enabled = value
50210
50220  If value = True Then
50230    SetProgramOptions
50240    tbstrProgGeneral_Click
50250   Else
50260    dmFraProgGeneral1.Enabled = False
50270    dmFraBrowserAddOn.Enabled = False
50280    dmFraProgGeneral2.Enabled = False
50290    dmFraShellIntegration.Enabled = False
50300    dmFraCheckUpdate.Enabled = False
50310  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGeneral", "SetControlsEnabled")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdCheckNow_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim upd As clsUpdate
50020  Set upd = New clsUpdate
50030  upd.CheckForUpdates True
50040  SetLastUpdateCeck Now
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGeneral", "cmdCheckNow_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub UserControl_Initialize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ctl As Control, i As Long
50020
50030  tbstrProgGeneral.Left = 0
50040  tbstrProgGeneral.Top = 0
50050  tbstrProgGeneral.Height = dmFraProgGeneral1.Height + dmFraBrowserAddOn.Height + 500
50060  UserControl.Height = tbstrProgGeneral.Height + 100
50070
50080  With tbstrProgGeneral
50090   .Top = 50
50100   .Left = 0
50110  End With
50120
50130  With tbstrProgGeneral.Tabs
50140   .Clear
50150   For i = 1 To 2
50160    .Add
50170   Next i
50180  End With
50190  tbstrProgGeneral.Visible = True
50200
50210  With cmbSendMailMethod
50220   .Clear
50230   For i = 1 To 2
50240    .AddItem ""
50250   Next i
50260  End With
50270
50280  With cmbOptionsDesign
50290   .Clear
50300   For i = 1 To 2
50310    .AddItem ""
50320   Next i
50330  End With
50340
50350  With cmbUpdateInterval
50360   .Clear
50370   For i = 1 To 4
50380    .AddItem ""
50390   Next i
50400  End With
50410
50420  With sldProcessPriority
50430   .TextPosition = sldBelowRight
50440   .TickFrequency = 1
50450   .TickStyle = sldTopLeft
50461   Select Case .value
         Case 0: 'Idle
50480     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityIdle
50490    Case 1: 'Normal
50500     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityNormal
50510    Case 2: 'High
50520     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityHigh
50530    Case 3: 'Realtime
50540     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityRealtime
50550   End Select
50560  End With
50570
50580  If IsWin9xMe = False Then
50590    lblProcessPriority.Enabled = True
50600    sldProcessPriority.Enabled = True
50610   Else
50620    lblProcessPriority.Enabled = False
50630    sldProcessPriority.Enabled = False
50640  End If
50650
50660  If FileExists(CompletePath(App.Path) & "BrowserAddOn\PDFCreator Browser Add On-14_0_170_setup.exe") And (Not BrowserAddOnIsInstalled) Then
50670    dmFraBrowserAddOn.Enabled = True
50680    cmdInstallBrowserAddOn.Enabled = True
50690   Else
50700    dmFraBrowserAddOn.Enabled = False
50710    cmdInstallBrowserAddOn.Enabled = False
50720  End If
50730
50740  SetFrames Options.OptionsDesign
50750  ControlsEnabled = True
50760  SetProgramOptions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGeneral", "UserControl_Initialize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SetFrames(OptionsDesign As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ctl As Control
50020  For Each ctl In UserControl.Controls
50030   If TypeOf ctl Is dmFrame Then
50040    SetFrame ctl, OptionsDesign
50050   End If
50060  Next ctl
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGeneral", "SetFrames")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub UserControl_Resize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  tbstrProgGeneral.Width = UserControl.Width
50020  With dmFraProgGeneral1
50030   .Top = tbstrProgGeneral.ClientTop + 50
50040   .Left = tbstrProgGeneral.Left + (tbstrProgGeneral.Width - .Width) / 2
50050  End With
50060  With dmFraBrowserAddOn
50070   .Top = dmFraProgGeneral1.Top + dmFraProgGeneral1.Height + 50
50080   .Left = dmFraProgGeneral1.Left
50090  End With
50100  With dmFraProgGeneral2
50110   .Top = tbstrProgGeneral.ClientTop + 50
50120   .Left = tbstrProgGeneral.Left + (tbstrProgGeneral.Width - .Width) / 2
50130  End With
50140  With dmFraShellIntegration
50150   .Top = dmFraProgGeneral2.Top + dmFraProgGeneral2.Height + 50
50160   .Left = tbstrProgGeneral.Left + (tbstrProgGeneral.Width - .Width) / 2
50170  End With
50180  With dmFraCheckUpdate
50190   .Top = dmFraShellIntegration.Top + dmFraShellIntegration.Height + 50
50200   .Left = tbstrProgGeneral.Left + (tbstrProgGeneral.Width - .Width) / 2
50210  End With
50220
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGeneral", "UserControl_Resize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SetLanguageStrings()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With LanguageStrings
50020   dmFraProgGeneral1.Visible = True
50030
50040   dmFraProgGeneral1.Caption = .OptionsProgramGeneralDescription1
50050   dmFraProgGeneral2.Caption = .OptionsProgramGeneralDescription2
50060   tbstrProgGeneral.Tabs(1).Caption = LanguageStrings.OptionsProgramGeneralDescription1
50070   tbstrProgGeneral.Tabs(2).Caption = LanguageStrings.OptionsProgramGeneralDescription2
50080
50090   dmFraShellIntegration.Caption = .OptionsShellIntegration
50100   cmdShellintegration(0).Caption = .OptionsShellIntegrationAdd
50110   cmdShellintegration(1).Caption = .OptionsShellIntegrationRemove
50120   lblSendMailMethod.Caption = .OptionsSendMailMethod
50130   cmbSendMailMethod.List(0) = .OptionsSendMailMethodAutomatic
50140   cmbSendMailMethod.List(1) = .OptionsSendMailMethodMapi
50150   cmbSendMailMethod.List(2) = .OptionsSendMailMethodSendmailDLL
50160   chkNoConfirmMessageSwitchingDefaultprinter.Caption = .OptionsProgramSwitchingDefaultprinter
50170   chkNoProcessingAtStartup.Caption = .OptionsProgramNoProcessingAtStartup
50180   lblOptionsDesign.Caption = .OptionsProgramOptionsDesign
50190   cmbOptionsDesign.List(0) = .OptionsProgramOptionsDesignGradient
50200   cmbOptionsDesign.List(1) = .OptionsProgramOptionsDesignSimple
50210   chkShowAnimation.Caption = .OptionsProgramShowAnimation
50220   cmdTestpage.Caption = .OptionsPrintTestpage
50230   lblProcessPriority.Caption = .OptionsProcesspriority
50240   cmdAsso.Caption = .OptionsAssociatePSFiles
50250   dmFraBrowserAddOn.Caption = .OptionsBrowserAddOn
50260   cmdInstallBrowserAddOn.Caption = .OptionsBrowserAddOnInstall
50270   dmFraCheckUpdate.Caption = .OptionsCheckUpdateDescription
50280   lblUpdateInterval.Caption = .OptionsCheckUpdateInterval
50290   cmbUpdateInterval.List(0) = .OptionsCheckUpdateInterval01
50300   cmbUpdateInterval.List(1) = .OptionsCheckUpdateInterval02
50310   cmbUpdateInterval.List(2) = .OptionsCheckUpdateInterval03
50320   cmbUpdateInterval.List(3) = .OptionsCheckUpdateInterval04
50330   cmdCheckNow.Caption = .OptionsCheckUpdateNow
50340  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGeneral", "SetLanguageStrings")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SetOptions()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With Options1
50020   chkNoConfirmMessageSwitchingDefaultprinter.value = .NoConfirmMessageSwitchingDefaultprinter
50030   chkNoProcessingAtStartup.value = .NoProcessingAtStartup
50040   cmbOptionsDesign.ListIndex = .OptionsDesign
50050   sldProcessPriority.value = .ProcessPriority
50060   cmbSendMailMethod.ListIndex = .SendMailMethod
50070   chkShowAnimation.value = .ShowAnimation
50080   cmbUpdateInterval.ListIndex = .UpdateInterval
50090  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGeneral", "SetOptions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SetProgramOptions()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If IsWin9xMe = False Then
50020   If IsAdmin = False Then
50030    cmdShellintegration(0).Enabled = False
50040    cmdShellintegration(1).Enabled = False
50050   End If
50060  End If
50070  If IsPsAssociate = False Then
50080    cmdAsso.Enabled = True
50090   Else
50100    cmdAsso.Enabled = False
50110  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGeneral", "SetProgramOptions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub GetOptions()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With Options1
50020   .NoConfirmMessageSwitchingDefaultprinter = Abs(chkNoConfirmMessageSwitchingDefaultprinter.value)
50030   .NoProcessingAtStartup = Abs(chkNoProcessingAtStartup.value)
50040   If LenB(CStr(cmbOptionsDesign.ListIndex)) > 0 Then
50050    .OptionsDesign = cmbOptionsDesign.ListIndex
50060   End If
50070   If LenB(CStr(cmbSendMailMethod.ListIndex)) > 0 Then
50080    .SendMailMethod = cmbSendMailMethod.ListIndex
50090   End If
50100   .ShowAnimation = Abs(chkShowAnimation.value)
50110   If LenB(CStr(sldProcessPriority.value)) > 0 Then
50120    .ProcessPriority = sldProcessPriority.value
50130   End If
50140   .UpdateInterval = cmbUpdateInterval.ListIndex
50150  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGeneral", "GetOptions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub tbstrProgGeneral_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case tbstrProgGeneral.SelectedItem.Index
        Case 1
50030    dmFraProgGeneral2.Enabled = False
50040    dmFraProgGeneral2.Visible = False
50050    dmFraShellIntegration.Enabled = False
50060    dmFraShellIntegration.Visible = False
50070    dmFraCheckUpdate.Enabled = False
50080    dmFraCheckUpdate.Visible = False
50090    dmFraProgGeneral1.Visible = True
50100    dmFraBrowserAddOn.Visible = True
50110    If ControlsEnabled Then
50120     dmFraProgGeneral1.Enabled = True
50130     dmFraBrowserAddOn.Enabled = True
50140    End If
50150   Case 2
50160    dmFraProgGeneral1.Enabled = False
50170    dmFraProgGeneral1.Visible = False
50180    dmFraBrowserAddOn.Enabled = False
50190    dmFraBrowserAddOn.Visible = False
50200    dmFraProgGeneral2.Visible = True
50210    dmFraShellIntegration.Visible = True
50220    dmFraCheckUpdate.Visible = True
50230    If ControlsEnabled Then
50240     dmFraProgGeneral2.Enabled = True
50250     dmFraShellIntegration.Enabled = True
50260     dmFraCheckUpdate.Enabled = True
50270    End If
50280  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGeneral", "tbstrProgGeneral_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdAsso_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  PsAssociate
50020  SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
50030  cmdAsso.Enabled = False
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGeneral", "cmdAsso_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdTestpage_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  PrintTestpage frmMain
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGeneral", "cmdTestpage_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdShellintegration_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  MousePointer = vbHourglass
50020  cmdShellintegration(0).Enabled = False
50030  cmdShellintegration(1).Enabled = False
50041  Select Case Index
        Case 0
50060    AddExplorerIntegration
50070   Case 1
50080    RemoveExplorerIntegration
50090  End Select
50100  MousePointer = vbNormal
50110  cmdShellintegration(0).Enabled = True
50120  cmdShellintegration(1).Enabled = True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGeneral", "cmdShellintegration_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdInstallBrowserAddOn_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Shell """" & CompletePath(App.Path) & "BrowserAddOn\PDFCreator Browser Add On-14_0_170_setup.exe"""
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGeneral", "cmdInstallBrowserAddOn_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub sldProcessPriority_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & sldProcessPriority.Text
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGeneral", "sldProcessPriority_Change")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub sldProcessPriority_Scroll()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With sldProcessPriority
50021   Select Case .value
         Case 0: 'Idle
50040     .Text = LanguageStrings.OptionsProcesspriorityIdle
50050    Case 1: 'Normal
50060     .Text = LanguageStrings.OptionsProcesspriorityNormal
50070    Case 2: 'High
50080     .Text = LanguageStrings.OptionsProcesspriorityHigh
50090    Case 3: 'Realtime
50100     .Text = LanguageStrings.OptionsProcesspriorityRealtime
50110   End Select
50120  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGeneral", "sldProcessPriority_Scroll")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbOptionsDesign_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  SetFrame UserControl.Parent.dmFraDescription, cmbOptionsDesign.ListIndex
50020  UserControl.Parent.SetFrames cmbOptionsDesign.ListIndex
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGeneral", "cmbOptionsDesign_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
