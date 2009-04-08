VERSION 5.00
Begin VB.UserControl ctlOptFormatPostscript 
   ClientHeight    =   1365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   ScaleHeight     =   1365
   ScaleWidth      =   6765
   Begin PDFCreator.dmFrame dmFraPSGeneral 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   1931
      Caption         =   "Postscript"
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
      Begin VB.ComboBox cmbPSLanguageLevel 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown-Liste
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
      Begin VB.ComboBox cmbEPSLanguageLevel 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown-Liste
         TabIndex        =   3
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblLangLevel 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         Caption         =   "Language Level:"
         Height          =   195
         Left            =   735
         TabIndex        =   1
         Top             =   510
         Width           =   1200
      End
   End
End
Attribute VB_Name = "ctlOptFormatPostscript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub UserControl_Initialize()
 Dim ctl As Control
 dmFraPSGeneral.Left = 0
 dmFraPSGeneral.Top = 0
 UserControl.Height = dmFraPSGeneral.Height
 
 With cmbPSLanguageLevel
  .AddItem "1"
  .AddItem "1.5"
  .AddItem "2"
  .AddItem "3"
 End With
 With cmbEPSLanguageLevel
  .AddItem "1"
  .AddItem "1.5"
  .AddItem "2"
  .AddItem "3"
 End With
  
 cmbEPSLanguageLevel.Top = cmbPSLanguageLevel.Top
 cmbEPSLanguageLevel.Left = cmbPSLanguageLevel.Left
 
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
 dmFraPSGeneral.Width = UserControl.Width
End Sub

Public Sub SetLanguageStrings()
 With LanguageStrings
  lblLangLevel.Caption = .OptionsPSLanguageLevel
 End With
End Sub

Public Sub SetOptions()
 With Options
  cmbEPSLanguageLevel.ListIndex = .EPSLanguageLevel
  cmbPSLanguageLevel.ListIndex = .PSLanguageLevel
 End With
End Sub

Public Sub GetOptions()
 With Options
  If LenB(CStr(cmbEPSLanguageLevel.ListIndex)) > 0 Then
   .EPSLanguageLevel = cmbEPSLanguageLevel.ListIndex
  End If
  If LenB(CStr(cmbPSLanguageLevel.ListIndex)) > 0 Then
   .PSLanguageLevel = cmbPSLanguageLevel.ListIndex
  End If
 End With
End Sub
