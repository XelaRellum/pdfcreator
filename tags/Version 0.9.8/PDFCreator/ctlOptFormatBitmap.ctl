VERSION 5.00
Begin VB.UserControl ctlOptFormatBitmap 
   ClientHeight    =   2250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   ScaleHeight     =   2250
   ScaleWidth      =   6765
   Begin PDFCreator.dmFrame dmFraBitmapGeneral 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3413
      Caption         =   "Bitmap"
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
      Begin VB.ComboBox cmbPNGColors 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown-Liste
         TabIndex        =   5
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox txtJPEGQuality 
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         Height          =   285
         Left            =   1920
         TabIndex        =   11
         Text            =   "75"
         Top             =   1440
         Width           =   495
      End
      Begin VB.ComboBox cmbJPEGColors 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown-Liste
         TabIndex        =   6
         Top             =   1080
         Width           =   2175
      End
      Begin VB.ComboBox cmbBMPColors 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown-Liste
         TabIndex        =   7
         Top             =   1200
         Width           =   2175
      End
      Begin VB.ComboBox cmbPCXColors 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown-Liste
         TabIndex        =   8
         Top             =   1320
         Width           =   2175
      End
      Begin VB.ComboBox cmbTIFFColors 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown-Liste
         TabIndex        =   9
         Top             =   1440
         Width           =   2175
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
      Begin VB.Label lblJPEGQuality 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         Caption         =   "Quality:"
         Height          =   195
         Left            =   1290
         TabIndex        =   10
         Top             =   1485
         Width           =   525
      End
      Begin VB.Label lblJPEQQualityProzent 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   2520
         TabIndex        =   12
         Top             =   1485
         Width           =   120
      End
   End
End
Attribute VB_Name = "ctlOptFormatBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub UserControl_Initialize()
 Dim ctl As Control
 Dim i As Long
 dmFraBitmapGeneral.Left = 0
 dmFraBitmapGeneral.Top = 0
 UserControl.Height = dmFraBitmapGeneral.Height
  
 cmbJPEGColors.Left = cmbPNGColors.Left
 cmbJPEGColors.Width = cmbPNGColors.Width
 cmbJPEGColors.Top = cmbPNGColors.Top
 cmbBMPColors.Left = cmbPNGColors.Left
 cmbBMPColors.Width = cmbPNGColors.Width
 cmbBMPColors.Top = cmbPNGColors.Top
 cmbPCXColors.Left = cmbPNGColors.Left
 cmbPCXColors.Width = cmbPNGColors.Width
 cmbPCXColors.Top = cmbPNGColors.Top
 cmbTIFFColors.Left = cmbPNGColors.Left
 cmbTIFFColors.Width = cmbPNGColors.Width
 cmbTIFFColors.Top = cmbPNGColors.Top
 
 With cmbPNGColors
  .Clear
  For i = 1 To 4
   .AddItem ""
  Next i
  .ListIndex = 0
 End With
 With cmbJPEGColors
  .Clear
  For i = 1 To 2
   .AddItem ""
  Next i
  .ListIndex = 0
 End With
 With cmbBMPColors
  .Clear
  For i = 1 To 7
   .AddItem ""
  Next i
  .ListIndex = 0
 End With
 With cmbPCXColors
  .Clear
  For i = 1 To 6
   .AddItem ""
  Next i
  .ListIndex = 0
 End With
 With cmbTIFFColors
  .Clear
  For i = 1 To 6
   .AddItem ""
  Next i
  .ListIndex = 0
 End With
 
 txtBitmapResolution.Text = 150
 
 For Each ctl In UserControl.Controls
  If TypeOf ctl Is dmFrame Then
   ctl.Font.Size = 10
   ctl.TextShaddowColor = &HC00000
   If ComputerScreenResolution <= 8 Or Options.OptionsDesign = 1 Then
     ctl.UseGradient = False: ctl.Caption3D = [Flat Caption]
     ctl.BarColorFrom = vbBlue
    Else
     ctl.UseGradient = True: ctl.Caption3D = [Raised Caption]
     ctl.BarColorFrom = &HFF8080
     ctl.BarColorTo = &H400000
   End If
  End If
 Next ctl
End Sub

Private Sub UserControl_Resize()
 dmFraBitmapGeneral.Width = UserControl.Width
End Sub

Public Sub SetLanguageStrings()
 With LanguageStrings
  cmbPNGColors.List(0) = .OptionsPNGColorscount01
  cmbPNGColors.List(1) = .OptionsPNGColorscount02
  cmbPNGColors.List(2) = .OptionsPNGColorscount03
  cmbPNGColors.List(3) = .OptionsPNGColorscount04
  cmbJPEGColors.List(0) = .OptionsJPEGColorscount01
  cmbJPEGColors.List(1) = .OptionsJPEGColorscount02
  cmbBMPColors.List(0) = .OptionsBMPColorscount01
  cmbBMPColors.List(1) = .OptionsBMPColorscount02
  cmbBMPColors.List(2) = .OptionsBMPColorscount03
  cmbBMPColors.List(3) = .OptionsBMPColorscount04
  cmbBMPColors.List(4) = .OptionsBMPColorscount05
  cmbBMPColors.List(5) = .OptionsBMPColorscount06
  cmbBMPColors.List(6) = .OptionsBMPColorscount07
  cmbPCXColors.List(0) = .OptionsPCXColorscount01
  cmbPCXColors.List(1) = .OptionsPCXColorscount02
  cmbPCXColors.List(2) = .OptionsPCXColorscount03
  cmbPCXColors.List(3) = .OptionsPCXColorscount04
  cmbPCXColors.List(4) = .OptionsPCXColorscount05
  cmbPCXColors.List(5) = .OptionsPCXColorscount06
  cmbTIFFColors.List(0) = .OptionsTIFFColorscount01
  cmbTIFFColors.List(1) = .OptionsTIFFColorscount02
  cmbTIFFColors.List(2) = .OptionsTIFFColorscount03
  cmbTIFFColors.List(3) = .OptionsTIFFColorscount04
  cmbTIFFColors.List(4) = .OptionsTIFFColorscount05
  cmbTIFFColors.List(5) = .OptionsTIFFColorscount06
  cmbTIFFColors.List(6) = .OptionsTIFFColorscount07
  cmbTIFFColors.List(7) = .OptionsTIFFColorscount08

  dmFraBitmapGeneral.Caption = .OptionsImageSettings
  lblBitmapResolution = .OptionsBitmapResolution
  lblJPEGQuality = .OptionsJPEGQuality
  lblBitmapColors = .OptionsPDFColors
 End With
End Sub

Public Sub SetOptions()
 With Options
  cmbBMPColors.ListIndex = .BMPColorscount
  cmbJPEGColors.ListIndex = .JPEGColorscount
  cmbPCXColors.ListIndex = .PCXColorscount
  cmbPNGColors.ListIndex = .PNGColorscount
  cmbTIFFColors.ListIndex = .TIFFColorscount
  txtJPEGQuality.Text = .JPEGQuality
  txtBitmapResolution.Text = .BitmapResolution
 End With
End Sub

Public Sub GetOptions()
 With Options
  If LenB(CStr(cmbBMPColors.ListIndex)) > 0 Then
   .BMPColorscount = cmbBMPColors.ListIndex
  End If
  If LenB(CStr(cmbJPEGColors.ListIndex)) > 0 Then
   .JPEGColorscount = cmbJPEGColors.ListIndex
  End If
  If LenB(CStr(cmbPCXColors.ListIndex)) > 0 Then
   .PCXColorscount = cmbPCXColors.ListIndex
  End If
  If LenB(CStr(cmbPNGColors.ListIndex)) > 0 Then
   .PNGColorscount = cmbPNGColors.ListIndex
  End If
  If LenB(CStr(cmbTIFFColors.ListIndex)) > 0 Then
   .TIFFColorscount = cmbTIFFColors.ListIndex
  End If
  If LenB(txtBitmapResolution.Text) > 0 Then
   .BitmapResolution = txtBitmapResolution.Text
  End If
  If LenB(txtJPEGQuality.Text) > 0 Then
   .JPEGQuality = txtJPEGQuality.Text
  End If
 End With
End Sub
