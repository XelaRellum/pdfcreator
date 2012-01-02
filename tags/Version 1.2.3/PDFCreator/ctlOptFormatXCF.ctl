VERSION 5.00
Begin VB.UserControl ctlOptFormatXCF 
   ClientHeight    =   1725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6645
   ScaleHeight     =   1725
   ScaleWidth      =   6645
   ToolboxBitmap   =   "ctlOptFormatXCF.ctx":0000
   Begin PDFCreator.dmFrame dmFraXCFGeneral 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   2566
      Caption         =   "XCF"
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
      Begin VB.ComboBox cmbXCFColors 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown-Liste
         TabIndex        =   1
         Top             =   960
         Width           =   4215
      End
      Begin VB.Label lblBitmapResolution 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         Caption         =   "Resolution:"
         Height          =   195
         Left            =   1020
         TabIndex        =   5
         Top             =   525
         Width           =   795
      End
      Begin VB.Label lblBitmapDPI 
         AutoSize        =   -1  'True
         Caption         =   "dpi"
         Height          =   195
         Left            =   2520
         TabIndex        =   4
         Top             =   525
         Width           =   210
      End
      Begin VB.Label lblBitmapColors 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         Caption         =   "Colors:"
         Height          =   195
         Left            =   1335
         TabIndex        =   3
         Top             =   1020
         Width           =   480
      End
   End
End
Attribute VB_Name = "ctlOptFormatXCF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub UserControl_Initialize()
 Dim ctl As Control, i As Long
 dmFraXCFGeneral.Left = 0
 dmFraXCFGeneral.Top = 0
 UserControl.Height = dmFraXCFGeneral.Height

 With cmbXCFColors
  .Clear
  For i = 1 To 2
   .AddItem ""
  Next i
  .ListIndex = 0
 End With

 txtBitmapResolution.Text = 150
 
 SetFrames Options.OptionsDesign
End Sub

Public Sub SetFrames(OptionsDesign As Long)
 Dim ctl As Control
 For Each ctl In UserControl.Controls
  If TypeOf ctl Is dmFrame Then
   SetFrame ctl, OptionsDesign
  End If
 Next ctl
End Sub

Private Sub UserControl_Resize()
 dmFraXCFGeneral.Width = UserControl.Width
End Sub

Public Sub SetLanguageStrings()
 With LanguageStrings
  cmbXCFColors.List(0) = .OptionsXCFColorsCount01
  cmbXCFColors.List(1) = .OptionsXCFColorscount02

  dmFraXCFGeneral.Caption = .OptionsImageSettings
  lblBitmapResolution = .OptionsBitmapResolution
  lblBitmapColors = .OptionsPDFColors
 End With
End Sub

Public Sub SetOptions()
 With Options
  cmbXCFColors.ListIndex = .XCFColorsCount
  txtBitmapResolution.Text = .BitmapResolution
 End With
End Sub

Public Sub GetOptions()
 With Options
  If LenB(CStr(cmbXCFColors.ListIndex)) > 0 Then
   .XCFColorsCount = cmbXCFColors.ListIndex
  End If
  If LenB(txtBitmapResolution.Text) > 0 Then
   .BitmapResolution = txtBitmapResolution.Text
  End If
 End With
End Sub

