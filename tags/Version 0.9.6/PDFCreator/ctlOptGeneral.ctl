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
   Begin PDFCreator.dmFrame dmFraShellIntegration 
      Height          =   1065
      Left            =   120
      TabIndex        =   18
      Top             =   4560
      Visible         =   0   'False
      Width           =   6195
      _extentx        =   10927
      _extenty        =   1879
      caption         =   "Shell integration"
      caption3d       =   2
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "ctlOptGeneral.ctx":0312
      textshaddowcolor=   12582912
      enabled         =   0   'False
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
      _extentx        =   10927
      _extenty        =   4842
      caption         =   "General 2"
      caption3d       =   2
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "ctlOptGeneral.ctx":033E
      textshaddowcolor=   12582912
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
         ItemData        =   "ctlOptGeneral.ctx":036A
         Left            =   120
         List            =   "ctlOptGeneral.ctx":036C
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
         _extentx        =   10451
         _extenty        =   53
         linetype        =   1
         3dhighlight     =   -2147483628
         3dshadow        =   -2147483632
         drawstyle       =   0
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
      _extentx        =   10927
      _extenty        =   7250
      caption         =   "General 1"
      caption3d       =   2
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "ctlOptGeneral.ctx":036E
      textshaddowcolor=   12582912
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
         _extentx        =   10451
         _extenty        =   53
         linetype        =   1
         3dhighlight     =   -2147483628
         3dshadow        =   -2147483632
         drawstyle       =   0
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
         _extentx        =   10451
         _extenty        =   53
         linetype        =   1
         3dhighlight     =   -2147483628
         3dshadow        =   -2147483632
         drawstyle       =   0
      End
      Begin PDFCreator.Line3D Line3D1 
         Height          =   30
         Index           =   2
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3240
         Width           =   5925
         _extentx        =   10451
         _extenty        =   53
         linetype        =   1
         3dhighlight     =   -2147483628
         3dshadow        =   -2147483632
         drawstyle       =   0
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

Private Sub UserControl_Initialize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ctl As Control
50020  Dim i As Long
50030
50040  tbstrProgGeneral.Left = 0
50050  tbstrProgGeneral.Top = 0
50060  tbstrProgGeneral.Height = dmFraProgGeneral1.Height + 500
50070  UserControl.Height = tbstrProgGeneral.Height + 100
50080
50090  With tbstrProgGeneral
50100   .Top = 50
50110   .Left = 0
50120  End With
50130
50140  With tbstrProgGeneral.Tabs
50150   .Clear
50160   For i = 1 To 2
50170    .Add
50180   Next i
50190  End With
50200  tbstrProgGeneral.Visible = True
50210
50220  With cmbSendMailMethod
50230   .Clear
50240   For i = 1 To 2
50250    .AddItem ""
50260   Next i
50270  End With
50280
50290  With cmbOptionsDesign
50300   .Clear
50310   For i = 1 To 2
50320    .AddItem ""
50330   Next i
50340  End With
50350
50360  If IsWin9xMe = False Then
50370   If IsAdmin = False Then
50380    cmdShellintegration(0).Enabled = False
50390    cmdShellintegration(1).Enabled = False
50400   End If
50410  End If
50420  If IsPsAssociate = False Then
50430    cmdAsso.Enabled = True
50440   Else
50450    cmdAsso.Enabled = False
50460  End If
50470  With sldProcessPriority
50480   .TextPosition = sldBelowRight
50490   .TickFrequency = 1
50500   .TickStyle = sldTopLeft
50511   Select Case .value
         Case 0: 'Idle
50530     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityIdle
50540    Case 1: 'Normal
50550     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityNormal
50560    Case 2: 'High
50570     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityHigh
50580    Case 3: 'Realtime
50590     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityRealtime
50600   End Select
50610  End With
50620
50630  If IsWin9xMe = False Then
50640    lblProcessPriority.Enabled = True
50650    sldProcessPriority.Enabled = True
50660   Else
50670    lblProcessPriority.Enabled = False
50680    sldProcessPriority.Enabled = False
50690  End If
50700
50710  SetFrames Options.OptionsDesign
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
50060  With dmFraProgGeneral2
50070   .Top = tbstrProgGeneral.ClientTop + 50
50080   .Left = tbstrProgGeneral.Left + (tbstrProgGeneral.Width - .Width) / 2
50090  End With
50100  With dmFraShellIntegration
50110   .Top = dmFraProgGeneral2.Top + dmFraProgGeneral2.Height + 50
50120   .Left = tbstrProgGeneral.Left + (tbstrProgGeneral.Width - .Width) / 2
50130  End With
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
50250  End With
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
50010  With Options
50020   chkNoConfirmMessageSwitchingDefaultprinter = .NoConfirmMessageSwitchingDefaultprinter
50030   chkNoProcessingAtStartup = .NoProcessingAtStartup
50040   cmbOptionsDesign.ListIndex = .OptionsDesign
50050   sldProcessPriority.value = .ProcessPriority
50060   cmbSendMailMethod.ListIndex = .SendMailMethod
50070   chkShowAnimation.value = .ShowAnimation
50080  End With
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

Public Sub GetOptions()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With Options
50020   .NoConfirmMessageSwitchingDefaultprinter = Abs(chkNoConfirmMessageSwitchingDefaultprinter)
50030   .NoProcessingAtStartup = Abs(chkNoProcessingAtStartup)
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
50140  End With
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
50070    dmFraProgGeneral1.Enabled = True
50080    dmFraProgGeneral1.Visible = True
50090   Case 2
50100    dmFraProgGeneral1.Enabled = False
50110    dmFraProgGeneral1.Visible = False
50120    dmFraProgGeneral2.Enabled = True
50130    dmFraProgGeneral2.Visible = True
50140    dmFraShellIntegration.Enabled = True
50150    dmFraShellIntegration.Visible = True
50160  End Select
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
