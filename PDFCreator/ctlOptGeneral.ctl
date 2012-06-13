VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
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
      TabIndex        =   21
      Top             =   3240
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
         TabIndex        =   24
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox cmbUpdateInterval 
         Height          =   315
         ItemData        =   "ctlOptGeneral.ctx":0312
         Left            =   120
         List            =   "ctlOptGeneral.ctx":0319
         Style           =   2  'Dropdown-Liste
         TabIndex        =   22
         Top             =   600
         Width           =   3870
      End
      Begin VB.Label lblEnableNotice 
         AutoSize        =   -1  'True
         Caption         =   "You can set these options in the default profile only."
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   3645
      End
      Begin VB.Label lblUpdateInterval 
         AutoSize        =   -1  'True
         Caption         =   "Update interval"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1080
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
         Caption         =   "Remove shell Integration"
         Height          =   495
         Index           =   1
         Left            =   3150
         TabIndex        =   20
         Top             =   420
         Width           =   2910
      End
      Begin VB.Label lblEnableNotice 
         AutoSize        =   -1  'True
         Caption         =   "You can set these options in the default profile only."
         Enabled         =   0   'False
         Height          =   195
         Index           =   3
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Visible         =   0   'False
         Width           =   3645
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
         ItemData        =   "ctlOptGeneral.ctx":0330
         Left            =   120
         List            =   "ctlOptGeneral.ctx":0337
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
      Begin VB.Label lblEnableNotice 
         AutoSize        =   -1  'True
         Caption         =   "You can set these options in the default profile only."
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Visible         =   0   'False
         Width           =   3645
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
      Begin VB.Label lblEnableNotice 
         AutoSize        =   -1  'True
         Caption         =   "You can set these options in the default profile only."
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Visible         =   0   'False
         Width           =   3645
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

Private mEnabled As Boolean
Private mControlsEnabled As Boolean

Public Property Let ControlsEnabled(value As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long
50020  mEnabled = value
50030
50040  chkNoConfirmMessageSwitchingDefaultprinter.Enabled = mEnabled
50050  chkNoConfirmMessageSwitchingDefaultprinter.Visible = mEnabled
50060  chkNoProcessingAtStartup.Enabled = mEnabled
50070  chkNoProcessingAtStartup.Visible = mEnabled
50080  cmbOptionsDesign.Enabled = mEnabled
50090  cmbOptionsDesign.Visible = mEnabled
50100  sldProcessPriority.Enabled = mEnabled
50110  sldProcessPriority.Visible = mEnabled
50120  cmbSendMailMethod.Enabled = mEnabled
50130  cmbSendMailMethod.Visible = mEnabled
50140  chkShowAnimation.Enabled = mEnabled
50150  chkShowAnimation.Visible = mEnabled
50160  cmbUpdateInterval.Enabled = mEnabled
50170  cmbUpdateInterval.Visible = mEnabled
50180
50190  cmdTestpage.Enabled = mEnabled
50200  cmdTestpage.Visible = mEnabled
50210  lblProcessPriority.Enabled = mEnabled
50220  lblProcessPriority.Visible = mEnabled
50230  lblSendMailMethod.Enabled = mEnabled
50240  lblSendMailMethod.Visible = mEnabled
50250  cmdShellintegration(0).Enabled = mEnabled
50260  cmdShellintegration(0).Visible = mEnabled
50270  cmdShellintegration(1).Enabled = mEnabled
50280  cmdShellintegration(1).Visible = mEnabled
50290  cmdAsso.Enabled = mEnabled
50300  cmdAsso.Visible = mEnabled
50310  lblOptionsDesign.Enabled = mEnabled
50320  lblOptionsDesign.Visible = mEnabled
50330  lblUpdateInterval.Enabled = mEnabled
50340  lblUpdateInterval.Visible = mEnabled
50350  cmdCheckNow.Enabled = mEnabled
50360  cmdCheckNow.Visible = mEnabled
50370
50380  If mEnabled = True Then
50390    SetProgramOptions
50400    tbstrProgGeneral_Click
50410   Else
50420    'dmFraProgGeneral1.Enabled = False
50430    'dmFraProgGeneral2.Enabled = False
50440    dmFraShellIntegration.Enabled = False
50450    dmFraCheckUpdate.Enabled = False
50460  End If
50470
50480  For i = lblEnableNotice.LBound To lblEnableNotice.UBound
50490   lblEnableNotice(i).Visible = Not mEnabled
50500  Next i
50510  If mControlsEnabled Then
50520    For i = lblEnableNotice.LBound To lblEnableNotice.UBound
50530     lblEnableNotice(i).Enabled = Not mEnabled
50540    Next i
50550   Else
50560    For i = lblEnableNotice.LBound To lblEnableNotice.UBound
50570     lblEnableNotice(i).Enabled = False
50580    Next i
50590  End If
50600
50610  For i = Line3D1.LBound To Line3D1.UBound
50620   Line3D1(i).Visible = mEnabled
50630  Next i
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGeneral", "ControlsEnabled [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get ControlEnabled() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ControlEnabled = mEnabled
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGeneral", "ControlEnabled [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Sub SetControlsEnabled(value As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  mControlsEnabled = value
50020  ControlsEnabled = value
50030  dmFraProgGeneral1.Enabled = value
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
50030  upd.CheckForUpdates True, True
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
50050  tbstrProgGeneral.Height = dmFraProgGeneral2.Height + dmFraShellIntegration.Height + dmFraCheckUpdate.Height + 600
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
50650  SetFrames Options.OptionsDesign
50660
50670  For i = lblEnableNotice.LBound To lblEnableNotice.UBound
50680   lblEnableNotice(i).Top = 480
50690   lblEnableNotice(i).Left = 120
50700  Next i
50710
50720  mControlsEnabled = True
50730
50740  SetProgramOptions
50750
50760  SetFont
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

Public Sub SetFont()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With Options
50020   SetFontControls UserControl.Controls, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50030  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptGeneral", "SetFont")
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
50060  With dmFraProgGeneral2
50070   .Top = tbstrProgGeneral.ClientTop + 50
50080   .Left = tbstrProgGeneral.Left + (tbstrProgGeneral.Width - .Width) / 2
50090  End With
50100  With dmFraShellIntegration
50110   .Top = dmFraProgGeneral2.Top + dmFraProgGeneral2.Height + 50
50120   .Left = tbstrProgGeneral.Left + (tbstrProgGeneral.Width - .Width) / 2
50130  End With
50140  With dmFraCheckUpdate
50150   .Top = dmFraShellIntegration.Top + dmFraShellIntegration.Height + 50
50160   .Left = tbstrProgGeneral.Left + (tbstrProgGeneral.Width - .Width) / 2
50170  End With
50180
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
50010  Dim i As Long
50020  With LanguageStrings
50030   dmFraProgGeneral1.Visible = True
50040
50050   dmFraProgGeneral1.Caption = .OptionsProgramGeneralDescription1
50060   dmFraProgGeneral2.Caption = .OptionsProgramGeneralDescription2
50070   tbstrProgGeneral.Tabs(1).Caption = LanguageStrings.OptionsProgramGeneralDescription1
50080   tbstrProgGeneral.Tabs(2).Caption = LanguageStrings.OptionsProgramGeneralDescription2
50090
50100   dmFraShellIntegration.Caption = .OptionsShellIntegration
50110   cmdShellintegration(0).Caption = .OptionsShellIntegrationAdd
50120   cmdShellintegration(1).Caption = .OptionsShellIntegrationRemove
50130   lblSendMailMethod.Caption = .OptionsSendMailMethod
50140   cmbSendMailMethod.List(0) = .OptionsSendMailMethodAutomatic
50150   cmbSendMailMethod.List(1) = .OptionsSendMailMethodMapi
50160   cmbSendMailMethod.List(2) = .OptionsSendMailMethodSendmailDLL
50170   chkNoConfirmMessageSwitchingDefaultprinter.Caption = .OptionsProgramSwitchingDefaultprinter
50180   chkNoProcessingAtStartup.Caption = .OptionsProgramNoProcessingAtStartup
50190   lblOptionsDesign.Caption = .OptionsProgramOptionsDesign
50200   cmbOptionsDesign.List(0) = .OptionsProgramOptionsDesignGradient
50210   cmbOptionsDesign.List(1) = .OptionsProgramOptionsDesignSimple
50220   chkShowAnimation.Caption = .OptionsProgramShowAnimation
50230   cmdTestpage.Caption = .OptionsPrintTestpage
50240   lblProcessPriority.Caption = .OptionsProcesspriority
50250   cmdAsso.Caption = .OptionsAssociatePSFiles
50260   dmFraCheckUpdate.Caption = .OptionsCheckUpdateDescription
50270   lblUpdateInterval.Caption = .OptionsCheckUpdateInterval
50280   cmbUpdateInterval.List(0) = .OptionsCheckUpdateInterval01
50290   cmbUpdateInterval.List(1) = .OptionsCheckUpdateInterval02
50300   cmbUpdateInterval.List(2) = .OptionsCheckUpdateInterval03
50310   cmbUpdateInterval.List(3) = .OptionsCheckUpdateInterval04
50320   cmdCheckNow.Caption = .OptionsCheckUpdateNow
50330   For i = lblEnableNotice.LBound To lblEnableNotice.UBound
50340    lblEnableNotice(i).Caption = .OptionsEnableNotice
50350   Next i
50360  End With
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
50100    If mControlsEnabled Then
50110     dmFraProgGeneral1.Enabled = True
50120    End If
50130   Case 2
50140    dmFraProgGeneral1.Enabled = False
50150    dmFraProgGeneral1.Visible = False
50160    dmFraProgGeneral2.Visible = True
50170    dmFraShellIntegration.Visible = True
50180    dmFraCheckUpdate.Visible = True
50190    If mControlsEnabled Then
50200     dmFraProgGeneral2.Enabled = True
50210     dmFraShellIntegration.Enabled = True
50220     dmFraCheckUpdate.Enabled = True
50230    End If
50240  End Select
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
