VERSION 5.00
Begin VB.UserControl ctlOptPrint 
   ClientHeight    =   5700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6600
   ScaleHeight     =   5700
   ScaleWidth      =   6600
   ToolboxBitmap   =   "ctlOptPrint.ctx":0000
   Begin PDFCreator.dmFrame dmFraProgPrint 
      Height          =   5490
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   9684
      Caption         =   "Print after saving"
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
      Begin VB.CheckBox chkPrintAfterSavingMaxResolution 
         Appearance      =   0  '2D
         Caption         =   "Set maximum resolution"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   4080
         Width           =   6015
      End
      Begin VB.ComboBox cmbPrintAfterSavingMaxResolution 
         Height          =   315
         Left            =   420
         TabIndex        =   11
         Text            =   "cmbPrintAfterSavingMaxResolution"
         Top             =   4440
         Width           =   2130
      End
      Begin VB.ComboBox cmbPrintAfterSavingBitsPerPixel 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown-Liste
         TabIndex        =   9
         Top             =   2400
         Width           =   2130
      End
      Begin VB.CheckBox chkPrintAfterSaving 
         Appearance      =   0  '2D
         Caption         =   "Print after saving"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   420
         Width           =   6015
      End
      Begin VB.ComboBox cmbPrintAfterSavingPrinter 
         Enabled         =   0   'False
         Height          =   315
         Left            =   420
         TabIndex        =   3
         Top             =   1635
         Width           =   4470
      End
      Begin VB.ComboBox cmbPrintAfterSavingQueryUser 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown-Liste
         TabIndex        =   5
         Top             =   1035
         Width           =   4750
      End
      Begin VB.CheckBox chkPrintAfterSavingNoCancel 
         Appearance      =   0  '2D
         Caption         =   "No cancel dialog"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   5040
         Width           =   6015
      End
      Begin VB.CheckBox chkPrintAfterSavingDuplex 
         Appearance      =   0  '2D
         Caption         =   "Duplex"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   3045
         Width           =   6015
      End
      Begin VB.ComboBox cmbPrintAfterSavingTumble 
         Height          =   315
         ItemData        =   "ctlOptPrint.ctx":0312
         Left            =   420
         List            =   "ctlOptPrint.ctx":0314
         Style           =   2  'Dropdown-Liste
         TabIndex        =   8
         Top             =   3360
         Width           =   4470
      End
      Begin VB.Label lblPrintAfterSavingBitsPerPixel 
         AutoSize        =   -1  'True
         Caption         =   "Bits per pixel"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   885
      End
      Begin VB.Label lblPrintAfterSavingPrinter 
         AutoSize        =   -1  'True
         Caption         =   "Printer"
         Height          =   195
         Left            =   420
         TabIndex        =   2
         Top             =   1425
         Width           =   450
      End
      Begin VB.Label lblPrintAfterSavingQueryUser 
         AutoSize        =   -1  'True
         Caption         =   "Query user"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   825
         Width           =   765
      End
   End
End
Attribute VB_Name = "ctlOptPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub chkPrintAfterSavingMaxResolution_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkPrintAfterSavingMaxResolution.value = 1 Then
50020    ViewMaxResolution True
50030   Else
50040    ViewMaxResolution False
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptPrint", "chkPrintAfterSavingMaxResolution_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbPrintAfterSavingMaxResolution_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  KeyAscii = AllowedKeypressChars(KeyAscii)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptPrint", "cmbPrintAfterSavingMaxResolution_KeyPress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbPrintAfterSavingQueryUser_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If cmbPrintAfterSavingQueryUser.ListIndex = 0 Then
50020    cmbPrintAfterSavingPrinter.Enabled = True
50030    cmbPrintAfterSavingPrinter.BackColor = &H80000005
50040   Else
50050    cmbPrintAfterSavingPrinter.Enabled = False
50060    cmbPrintAfterSavingPrinter.BackColor = &H8000000F
50070  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptPrint", "cmbPrintAfterSavingQueryUser_Click")
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
50010  Dim ctl As Control
50020  Dim i As Long, p As Printer
50030  dmFraProgPrint.Left = 0
50040  dmFraProgPrint.Top = 0
50050  UserControl.Height = dmFraProgPrint.Height
50060  cmbPrintAfterSavingPrinter.Clear
50070  For Each p In Printers
50080   cmbPrintAfterSavingPrinter.AddItem p.DeviceName
50090  Next p
50100  With cmbPrintAfterSavingQueryUser
50110   .Clear
50120   For i = 1 To 4
50130    .AddItem ""
50140   Next i
50150  End With
50160  With cmbPrintAfterSavingTumble
50170   .Clear
50180   For i = 1 To 2
50190    .AddItem ""
50200   Next i
50210  End With
50220  If chkPrintAfterSaving.value = 1 Then
50230    ViewPrintAfterSaving True
50240   Else
50250    ViewPrintAfterSaving False
50260  End If
50270  With cmbPrintAfterSavingBitsPerPixel
50280   .Clear
50290   For i = 1 To 3
50300    .AddItem ""
50310   Next i
50320  End With
50330  With cmbPrintAfterSavingMaxResolution
50340   .Clear
50350   .AddItem "72"
50360   .AddItem "96"
50370   .AddItem "150"
50380   .AddItem "300"
50390   .AddItem "600"
50400   .AddItem "1200"
50410  End With
50420
50430  SetFrames Options.OptionsDesign
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptPrint", "UserControl_Initialize")
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
Select Case ErrPtnr.OnError("ctlOptPrint", "SetFrames")
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
50010  dmFraProgPrint.Width = UserControl.Width
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptPrint", "UserControl_Resize")
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
50020   dmFraProgPrint.Caption = .OptionsProgramPrintSymbol
50030   chkPrintAfterSaving.Caption = .OptionsPrintAfterSaving
50040   lblPrintAfterSavingPrinter.Caption = .OptionsPrintAfterSavingPrinter
50050
50060   lblPrintAfterSavingQueryUser.Caption = .OptionsPrintAfterSavingQueryUser
50070   cmbPrintAfterSavingQueryUser.List(0) = .OptionsPrintAfterSavingQueryUserOff
50080   cmbPrintAfterSavingQueryUser.List(1) = .OptionsPrintAfterSavingQueryUserStandardPrinterDialog
50090   cmbPrintAfterSavingQueryUser.List(2) = .OptionsPrintAfterSavingQueryUserPrinterSetupDialog
50100   cmbPrintAfterSavingQueryUser.List(3) = .OptionsPrintAfterSavingQueryUserDefaultPrinter
50110
50120   chkPrintAfterSavingNoCancel.Caption = .OptionsPrintAfterSavingNoCancel
50130   chkPrintAfterSavingDuplex.Caption = .OptionsPrintAfterSavingDuplex
50140   cmbPrintAfterSavingTumble.List(0) = .OptionsPrintAfterSavingDuplexTumbleOff
50150   cmbPrintAfterSavingTumble.List(1) = .OptionsPrintAfterSavingDuplexTumbleOn
50160
50170   lblPrintAfterSavingBitsPerPixel.Caption = .OptionsPrintAfterSavingBitsPerPixel
50180   cmbPrintAfterSavingBitsPerPixel.List(0) = .OptionsPrintAfterSavingBitsPerPixelMono
50190   cmbPrintAfterSavingBitsPerPixel.List(1) = .OptionsPrintAfterSavingBitsPerPixelCMYK
50200   cmbPrintAfterSavingBitsPerPixel.List(2) = .OptionsPrintAfterSavingBitsPerPixelTrueColor
50210
50220   chkPrintAfterSavingMaxResolution.Caption = .OptionsPrintAfterSavingMaxResolution
50230  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptPrint", "SetLanguageStrings")
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
50020   chkPrintAfterSavingDuplex.value = .PrintAfterSavingDuplex
50030   chkPrintAfterSavingNoCancel.value = .PrintAfterSavingNoCancel
50040   cmbPrintAfterSavingPrinter.Text = .PrintAfterSavingPrinter
50050   cmbPrintAfterSavingQueryUser.ListIndex = .PrintAfterSavingQueryUser
50060   cmbPrintAfterSavingTumble.ListIndex = .PrintAfterSavingTumble
50070   cmbPrintAfterSavingBitsPerPixel.ListIndex = .PrintAfterSavingBitsPerPixel
50080   chkPrintAfterSavingMaxResolution.value = .PrintAfterSavingMaxResolutionEnabled
50090   cmbPrintAfterSavingMaxResolution.Text = .PrintAfterSavingMaxResolution
50100   chkPrintAfterSaving.value = .PrintAfterSaving
50110   chkPrintAfterSaving_Click
50120  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptPrint", "SetOptions")
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
50010  Dim sMaxResolution As String, MaxResolution As Long
50020  With Options1
50030   .PrintAfterSaving = Abs(chkPrintAfterSaving.value)
50040   .PrintAfterSavingDuplex = Abs(chkPrintAfterSavingDuplex.value)
50050   .PrintAfterSavingNoCancel = Abs(chkPrintAfterSavingNoCancel.value)
50060   .PrintAfterSavingPrinter = cmbPrintAfterSavingPrinter.Text
50070   .PrintAfterSavingQueryUser = cmbPrintAfterSavingQueryUser.ListIndex
50080   .PrintAfterSavingTumble = cmbPrintAfterSavingTumble.ListIndex
50090   .PrintAfterSavingBitsPerPixel = cmbPrintAfterSavingBitsPerPixel.ListIndex
50100   .PrintAfterSavingMaxResolutionEnabled = Abs(chkPrintAfterSavingMaxResolution.value)
50110
50120   sMaxResolution = Trim$(CStr(cmbPrintAfterSavingMaxResolution.Text))
50130   If LenB(sMaxResolution) > 0 Then
50140    If IsNumeric(sMaxResolution) Then
50150     MaxResolution = CLng(sMaxResolution)
50160     If MaxResolution >= 72 Then
50170      .PrintAfterSavingMaxResolution = MaxResolution
50180     End If
50190    End If
50200   End If
50210  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptPrint", "GetOptions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkPrintAfterSaving_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkPrintAfterSaving.value = 1 Then
50020    ViewPrintAfterSaving True
50030   Else
50040    ViewPrintAfterSaving False
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptPrint", "chkPrintAfterSaving_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkPrintAfterSavingDuplex_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkPrintAfterSavingDuplex.value = 1 Then
50020    ViewPrintAfterTumble True
50030   Else
50040    ViewPrintAfterTumble False
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptPrint", "chkPrintAfterSavingDuplex_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ViewPrintAfterTumble(ViewIt As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  cmbPrintAfterSavingTumble.Enabled = ViewIt
50020
50030  If ViewIt Then
50040    cmbPrintAfterSavingTumble.BackColor = &H80000005
50050   Else
50060    cmbPrintAfterSavingTumble.BackColor = &H8000000F
50070  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptPrint", "ViewPrintAfterTumble")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ViewPrintAfterSaving(ViewIt As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  lblPrintAfterSavingQueryUser.Enabled = ViewIt
50020  cmbPrintAfterSavingQueryUser.Enabled = ViewIt
50030  chkPrintAfterSavingNoCancel.Enabled = ViewIt
50040  chkPrintAfterSavingDuplex.Enabled = ViewIt
50050  ViewPrintAfterSavingPrinter ViewIt
50060  chkPrintAfterSavingMaxResolution.Enabled = ViewIt
50070  ViewMaxResolution ViewIt
50080
50090  lblPrintAfterSavingBitsPerPixel.Enabled = ViewIt
50100  cmbPrintAfterSavingBitsPerPixel.Enabled = ViewIt
50110 ' cmbPrintAfterSavingPrinter.Enabled = ViewIt
50120
50130  If ViewIt Then
50140    cmbPrintAfterSavingQueryUser.BackColor = &H80000005
50150    cmbPrintAfterSavingBitsPerPixel.BackColor = &H80000005
50160   Else
50170    cmbPrintAfterSavingQueryUser.BackColor = &H8000000F
50180    cmbPrintAfterSavingBitsPerPixel.BackColor = &H8000000F
50190  End If
50200
50210  If chkPrintAfterSavingDuplex.value = 1 And ViewIt Then
50220    ViewPrintAfterTumple True
50230   Else
50240    ViewPrintAfterTumple False
50250  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptPrint", "ViewPrintAfterSaving")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ViewMaxResolution(ViewIt As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkPrintAfterSavingMaxResolution.value = 1 And ViewIt Then
50020    cmbPrintAfterSavingMaxResolution.Enabled = True
50030    cmbPrintAfterSavingMaxResolution.BackColor = &H80000005
50040   Else
50050    cmbPrintAfterSavingMaxResolution.Enabled = False
50060    cmbPrintAfterSavingMaxResolution.BackColor = &H8000000F
50070  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptPrint", "ViewMaxResolution")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ViewPrintAfterSavingPrinter(ViewIt As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  lblPrintAfterSavingPrinter.Enabled = ViewIt
50020  If ViewIt Then
50030    If cmbPrintAfterSavingQueryUser.ListIndex = 0 Then
50040      cmbPrintAfterSavingPrinter.Enabled = True
50050      cmbPrintAfterSavingPrinter.BackColor = &H80000005
50060     Else
50070      cmbPrintAfterSavingPrinter.Enabled = False
50080      cmbPrintAfterSavingPrinter.BackColor = &H8000000F
50090    End If
50100   Else
50110    cmbPrintAfterSavingPrinter.Enabled = ViewIt
50120    cmbPrintAfterSavingPrinter.BackColor = &H8000000F
50130  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptPrint", "ViewPrintAfterSavingPrinter")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ViewPrintAfterTumple(ViewIt As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  cmbPrintAfterSavingTumble.Enabled = ViewIt
50020
50030  If ViewIt Then
50040    cmbPrintAfterSavingTumble.BackColor = &H80000005
50050   Else
50060    cmbPrintAfterSavingTumble.BackColor = &H8000000F
50070  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptPrint", "ViewPrintAfterTumple")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
