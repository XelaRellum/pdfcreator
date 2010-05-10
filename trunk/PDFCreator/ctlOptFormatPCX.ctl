VERSION 5.00
Begin VB.UserControl ctlOptFormatPCX 
   ClientHeight    =   1725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6645
   ScaleHeight     =   1725
   ScaleWidth      =   6645
   ToolboxBitmap   =   "ctlOptFormatPCX.ctx":0000
   Begin PDFCreator.dmFrame dmFraPCXGeneral 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   2566
      Caption         =   "PCX"
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
      Begin VB.TextBox txtBitmapResolution 
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Text            =   "72"
         Top             =   480
         Width           =   495
      End
      Begin VB.ComboBox cmbPCXColors 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown-Liste
         TabIndex        =   5
         Top             =   960
         Width           =   4215
      End
      Begin VB.Label lblBitmapResolution 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         Caption         =   "Resolution:"
         Height          =   195
         Left            =   1020
         TabIndex        =   1
         Top             =   525
         Width           =   795
      End
      Begin VB.Label lblBitmapDPI 
         AutoSize        =   -1  'True
         Caption         =   "dpi"
         Height          =   195
         Left            =   2520
         TabIndex        =   3
         Top             =   525
         Width           =   210
      End
      Begin VB.Label lblBitmapColors 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         Caption         =   "Colors:"
         Height          =   195
         Left            =   1335
         TabIndex        =   4
         Top             =   1020
         Width           =   480
      End
   End
End
Attribute VB_Name = "ctlOptFormatPCX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub SetControlsEnabled(value As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  cmbPCXColors.Enabled = value
50020  txtBitmapResolution.Enabled = value
50030  lblBitmapResolution.Enabled = value
50040  lblBitmapDPI.Enabled = value
50050  lblBitmapColors.Enabled = value
50060  dmFraPCXGeneral.Enabled = value
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPCX", "SetControlsEnabled")
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
50020  Dim i As Long
50030  dmFraPCXGeneral.Left = 0
50040  dmFraPCXGeneral.Top = 0
50050  UserControl.Height = dmFraPCXGeneral.Height
50060
50070  With cmbPCXColors
50080   .Clear
50090   For i = 1 To 6
50100    .AddItem ""
50110   Next i
50120   .ListIndex = 0
50130  End With
50140
50150  txtBitmapResolution.Text = 150
50160
50170  SetFrames Options.OptionsDesign
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPCX", "UserControl_Initialize")
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
Select Case ErrPtnr.OnError("ctlOptFormatPCX", "SetFrames")
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
50010  dmFraPCXGeneral.Width = UserControl.Width
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPCX", "UserControl_Resize")
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
50020   cmbPCXColors.List(0) = .OptionsPCXColorscount01
50030   cmbPCXColors.List(1) = .OptionsPCXColorscount02
50040   cmbPCXColors.List(2) = .OptionsPCXColorscount03
50050   cmbPCXColors.List(3) = .OptionsPCXColorscount04
50060   cmbPCXColors.List(4) = .OptionsPCXColorscount05
50070   cmbPCXColors.List(5) = .OptionsPCXColorscount06
50080
50090   dmFraPCXGeneral.Caption = .OptionsImageSettings
50100   lblBitmapResolution = .OptionsBitmapResolution
50110   lblBitmapColors = .OptionsPDFColors
50120  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPCX", "SetLanguageStrings")
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
50020   cmbPCXColors.ListIndex = .PCXColorscount
50030   txtBitmapResolution.Text = .PCXResolution
50040  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPCX", "SetOptions")
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
50020   If LenB(CStr(cmbPCXColors.ListIndex)) > 0 Then
50030    .PCXColorscount = cmbPCXColors.ListIndex
50040   End If
50050   If LenB(txtBitmapResolution.Text) > 0 Then
50060    .PCXResolution = txtBitmapResolution.Text
50070   End If
50080  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPCX", "GetOptions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

