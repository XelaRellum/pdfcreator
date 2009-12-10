VERSION 5.00
Begin VB.UserControl ctlOptPrint 
   ClientHeight    =   4425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   ScaleHeight     =   4425
   ScaleWidth      =   6765
   ToolboxBitmap   =   "ctlOptPrint.ctx":0000
   Begin PDFCreator.dmFrame dmFraProgPrint 
      Height          =   3930
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6932
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
         Height          =   315
         Left            =   105
         TabIndex        =   3
         Top             =   1155
         Width           =   4770
      End
      Begin VB.ComboBox cmbPrintAfterSavingQueryUser 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown-Liste
         TabIndex        =   5
         Top             =   1995
         Width           =   4770
      End
      Begin VB.CheckBox chkPrintAfterSavingNoCancel 
         Appearance      =   0  '2D
         Caption         =   "No cancel dialog"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2625
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
         Left            =   420
         Style           =   2  'Dropdown-Liste
         TabIndex        =   8
         Top             =   3360
         Width           =   4470
      End
      Begin VB.Label lblPrintAfterSavingPrinter 
         AutoSize        =   -1  'True
         Caption         =   "Printer"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   945
         Width           =   450
      End
      Begin VB.Label lblPrintAfterSavingQueryUser 
         AutoSize        =   -1  'True
         Caption         =   "Query user"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1785
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
50270
50280  SetFrames Options.OptionsDesign
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
50160  End With
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
50020   chkPrintAfterSaving.value = .PrintAfterSaving
50030   chkPrintAfterSavingDuplex.value = .PrintAfterSavingDuplex
50040   chkPrintAfterSavingNoCancel.value = .PrintAfterSavingNoCancel
50050   cmbPrintAfterSavingPrinter.Text = .PrintAfterSavingPrinter
50060   cmbPrintAfterSavingQueryUser.ListIndex = .PrintAfterSavingQueryUser
50070   cmbPrintAfterSavingTumble.ListIndex = .PrintAfterSavingTumble
50080  End With
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
50010  With Options1
50020   .PrintAfterSaving = Abs(chkPrintAfterSaving.value)
50030   .PrintAfterSavingDuplex = Abs(chkPrintAfterSavingDuplex.value)
50040   .PrintAfterSavingNoCancel = Abs(chkPrintAfterSavingNoCancel.value)
50050   .PrintAfterSavingPrinter = cmbPrintAfterSavingPrinter.Text
50060   If LenB(CStr(cmbPrintAfterSavingQueryUser.ListIndex)) > 0 Then
50070    .PrintAfterSavingQueryUser = cmbPrintAfterSavingQueryUser.ListIndex
50080   End If
50090   If LenB(CStr(cmbPrintAfterSavingTumble.ListIndex)) > 0 Then
50100    .PrintAfterSavingTumble = cmbPrintAfterSavingTumble.ListIndex
50110   End If
50120  End With
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

Private Sub ViewPrintAfterTumble(Viewit As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  cmbPrintAfterSavingTumble.Enabled = Viewit
50020
50030  If Viewit Then
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

Private Sub ViewPrintAfterSaving(Viewit As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  lblPrintAfterSavingPrinter.Enabled = Viewit
50020  cmbPrintAfterSavingPrinter.Enabled = Viewit
50030  lblPrintAfterSavingQueryUser.Enabled = Viewit
50040  cmbPrintAfterSavingQueryUser.Enabled = Viewit
50050  chkPrintAfterSavingNoCancel.Enabled = Viewit
50060  chkPrintAfterSavingDuplex.Enabled = Viewit
50070
50080  If Viewit Then
50090    cmbPrintAfterSavingPrinter.BackColor = &H80000005
50100    cmbPrintAfterSavingQueryUser.BackColor = &H80000005
50110   Else
50120    cmbPrintAfterSavingPrinter.BackColor = &H8000000F
50130    cmbPrintAfterSavingQueryUser.BackColor = &H8000000F
50140  End If
50150
50160  If chkPrintAfterSavingDuplex.value = 1 And Viewit Then
50170    ViewPrintAfterTumple True
50180   Else
50190    ViewPrintAfterTumple False
50200  End If
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

Private Sub ViewPrintAfterTumple(Viewit As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  cmbPrintAfterSavingTumble.Enabled = Viewit
50020
50030  If Viewit Then
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
