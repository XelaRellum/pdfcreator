VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Options"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9120
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame fraPDFCompress 
      Caption         =   "Compression"
      Height          =   3855
      Left            =   3720
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Frame fraPDFGrey 
         Caption         =   "Greyscale Images"
         Height          =   975
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   5535
         Begin VB.TextBox txtPDFGreyRes 
            Height          =   285
            Left            =   4440
            TabIndex        =   26
            Top             =   540
            Width           =   735
         End
         Begin VB.ComboBox cmbPDFGreyResample 
            Height          =   315
            ItemData        =   "frmOptions.frx":08CA
            Left            =   2280
            List            =   "frmOptions.frx":08CC
            Style           =   2  'Dropdown-Liste
            TabIndex        =   25
            Tag             =   "Bicubic|Subsample|Average"
            Top             =   540
            Width           =   1935
         End
         Begin VB.CheckBox chkPDFGreyResample 
            Caption         =   "Resample"
            Height          =   255
            Left            =   2280
            TabIndex        =   24
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox cmbPDFGreyComp 
            Height          =   315
            ItemData        =   "frmOptions.frx":08CE
            Left            =   120
            List            =   "frmOptions.frx":08D0
            Style           =   2  'Dropdown-Liste
            TabIndex        =   23
            Top             =   540
            Width           =   1935
         End
         Begin VB.CheckBox chkPDFGreyComp 
            Caption         =   "Compress"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblPDFGreyRes 
            Caption         =   "Resolution"
            Height          =   255
            Left            =   4440
            TabIndex        =   27
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame fraPDFMono 
         Caption         =   "Monochrome Images"
         Height          =   975
         Left            =   120
         TabIndex        =   14
         Top             =   2760
         Width           =   5535
         Begin VB.CheckBox chkPDFMonoComp 
            Caption         =   "Compress"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox cmbPDFMonoComp 
            Height          =   315
            ItemData        =   "frmOptions.frx":08D2
            Left            =   120
            List            =   "frmOptions.frx":08D4
            Style           =   2  'Dropdown-Liste
            TabIndex        =   18
            Top             =   540
            Width           =   1935
         End
         Begin VB.CheckBox chkPDFMonoResample 
            Caption         =   "Resample"
            Height          =   255
            Left            =   2280
            TabIndex        =   17
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox cmbPDFMonoResample 
            Height          =   315
            ItemData        =   "frmOptions.frx":08D6
            Left            =   2280
            List            =   "frmOptions.frx":08D8
            Style           =   2  'Dropdown-Liste
            TabIndex        =   16
            Tag             =   "Bicubic|Subsample|Average"
            Top             =   540
            Width           =   1935
         End
         Begin VB.TextBox txtPDFMonoRes 
            Height          =   285
            Left            =   4440
            TabIndex        =   15
            Top             =   540
            Width           =   735
         End
         Begin VB.Label lblPDFMonoRes 
            Caption         =   "Resolution"
            Height          =   255
            Left            =   4440
            TabIndex        =   20
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.CheckBox chkPDFTextComp 
         Caption         =   "Compress Text Objects"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   4335
      End
      Begin VB.Frame fraPDFColor 
         Caption         =   "Color Images"
         Height          =   975
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   5535
         Begin VB.TextBox txtPDFColorRes 
            Height          =   285
            Left            =   4440
            TabIndex        =   11
            Top             =   540
            Width           =   735
         End
         Begin VB.ComboBox cmbPDFColorResample 
            Height          =   315
            ItemData        =   "frmOptions.frx":08DA
            Left            =   2280
            List            =   "frmOptions.frx":08DC
            Style           =   2  'Dropdown-Liste
            TabIndex        =   10
            Tag             =   "Bicubic|Subsample|Average"
            Top             =   540
            Width           =   1935
         End
         Begin VB.CheckBox chkPDFColorResample 
            Caption         =   "Resample"
            Height          =   255
            Left            =   2280
            TabIndex        =   9
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox cmbPDFColorComp 
            Height          =   315
            ItemData        =   "frmOptions.frx":08DE
            Left            =   120
            List            =   "frmOptions.frx":08E0
            Style           =   2  'Dropdown-Liste
            TabIndex        =   8
            Top             =   540
            Width           =   1935
         End
         Begin VB.CheckBox chkPDFColorComp 
            Caption         =   "Compress"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblPDFColorRes 
            Caption         =   "Resolution"
            Height          =   255
            Left            =   4440
            TabIndex        =   12
            Top             =   240
            Width           =   975
         End
      End
   End
   Begin VB.Frame fraPDFFonts 
      Caption         =   "Font Options"
      Height          =   2895
      Left            =   3600
      TabIndex        =   51
      Top             =   2640
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CheckBox chkPDFEmbedAll 
         Caption         =   "Embed all Fonts"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   360
         Width           =   5175
      End
      Begin VB.CheckBox chkPDFSubSetFonts 
         Caption         =   "Subset Fonts, when percentage of used characters below:"
         Height          =   495
         Left            =   120
         TabIndex        =   53
         Top             =   780
         Width           =   5535
      End
      Begin VB.TextBox txtPDFSubSetPerc 
         Height          =   285
         Left            =   360
         TabIndex        =   52
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblPDFPerc 
         Caption         =   "%"
         Height          =   255
         Left            =   960
         TabIndex        =   55
         Top             =   1320
         Width           =   615
      End
   End
   Begin VB.Frame fraProgFont 
      Caption         =   "Programfont"
      Height          =   4695
      Left            =   3480
      TabIndex        =   28
      Top             =   2400
      Width           =   6255
      Begin VB.CommandButton cmdCancelTest 
         Caption         =   "CancelTest"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1680
         TabIndex        =   76
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test"
         Height          =   495
         Left            =   120
         TabIndex        =   33
         Top             =   4080
         Width           =   1335
      End
      Begin VB.TextBox txtProgramFontsize 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   5400
         TabIndex        =   32
         Text            =   "8"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtTest 
         Height          =   2655
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   31
         Top             =   1320
         Width           =   6015
      End
      Begin VB.ComboBox cmbCharset 
         Height          =   315
         Left            =   3000
         TabIndex        =   30
         Text            =   "cmbCharset"
         Top             =   600
         Width           =   2295
      End
      Begin VB.ComboBox cmbFonts 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown-Liste
         TabIndex        =   29
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label lblSize 
         Caption         =   "Size"
         Height          =   255
         Left            =   5400
         TabIndex        =   37
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblTesttext 
         Caption         =   "Here you can test the font."
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1080
         Width           =   5295
      End
      Begin VB.Label lblProgcharset 
         Caption         =   "Charset"
         Height          =   255
         Left            =   3000
         TabIndex        =   35
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblProgfont 
         Caption         =   "Programfont"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame fraPDFColors 
      Caption         =   "Color Options"
      Height          =   3495
      Left            =   3360
      TabIndex        =   56
      Top             =   2160
      Visible         =   0   'False
      Width           =   5775
      Begin VB.ComboBox cmbPDFColorModel 
         Height          =   315
         ItemData        =   "frmOptions.frx":08E2
         Left            =   120
         List            =   "frmOptions.frx":08E4
         Style           =   2  'Dropdown-Liste
         TabIndex        =   62
         Tag             =   "RGB|CMYK|GRAY"
         Top             =   360
         Width           =   2775
      End
      Begin VB.Frame fraPDFColorOptions 
         Caption         =   "Options"
         Height          =   1455
         Left            =   120
         TabIndex        =   58
         Top             =   1920
         Width           =   5535
         Begin VB.CheckBox chkPDFPreserveHalftone 
            Caption         =   "Preserve Halftone Information"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   1080
            Width           =   5175
         End
         Begin VB.CheckBox chkPDFPreserveTransfer 
            Caption         =   "Preserve Transfer Functions"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Tag             =   "Remove|Preserve"
            Top             =   720
            Width           =   5175
         End
         Begin VB.CheckBox chkPDFPreserveOverprint 
            Caption         =   "Preserve Overprint Settings"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   360
            Width           =   5175
         End
      End
      Begin VB.CheckBox chkPDFCMYKtoRGB 
         Caption         =   "Convert CMYK Images to RGB"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   840
         Width           =   3255
      End
   End
   Begin VB.Frame fraPDFSecurity 
      Caption         =   "Not implented yet"
      Height          =   3855
      Left            =   3240
      TabIndex        =   38
      Top             =   1920
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.Frame fraPDFGeneral 
      Caption         =   "General Options"
      Height          =   2895
      Left            =   3120
      TabIndex        =   39
      Top             =   1680
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CheckBox chkPDFASCII85 
         Caption         =   "Convert binary data to ASCII85"
         Height          =   255
         Left            =   2400
         TabIndex        =   44
         Top             =   2400
         Width           =   3255
      End
      Begin VB.ComboBox cmbPDFOverprint 
         Height          =   315
         ItemData        =   "frmOptions.frx":08E6
         Left            =   2400
         List            =   "frmOptions.frx":08E8
         Style           =   2  'Dropdown-Liste
         TabIndex        =   43
         Top             =   1860
         Width           =   2655
      End
      Begin VB.TextBox txtPDFRes 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   2400
         TabIndex        =   42
         Text            =   "600"
         Top             =   1320
         Width           =   615
      End
      Begin VB.ComboBox cmbPDFCompat 
         Height          =   315
         ItemData        =   "frmOptions.frx":08EA
         Left            =   2400
         List            =   "frmOptions.frx":08EC
         Style           =   2  'Dropdown-Liste
         TabIndex        =   41
         Top             =   360
         Width           =   2655
      End
      Begin VB.ComboBox cmbPDFRotate 
         Height          =   315
         ItemData        =   "frmOptions.frx":08EE
         Left            =   2400
         List            =   "frmOptions.frx":08F0
         Style           =   2  'Dropdown-Liste
         TabIndex        =   40
         Tag             =   "None|All|PageByPage"
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label lblPDFDPI 
         BackStyle       =   0  'Transparent
         Caption         =   "dpi"
         Height          =   255
         Left            =   3120
         TabIndex        =   49
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblPDFOverprint 
         Alignment       =   1  'Rechts
         Caption         =   "Overprint:"
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label lblPDFResolution 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Resolution:"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   1380
         Width           =   2175
      End
      Begin VB.Label lblPDFCompat 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Compatibility:"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   420
         Width           =   2175
      End
      Begin VB.Label lblPDFAutoRotate 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Auto-Rotate Pages:"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   900
         Width           =   2175
      End
   End
   Begin VB.Frame fraProgAutosave 
      Caption         =   "Autosave"
      Height          =   2895
      Left            =   3000
      TabIndex        =   82
      Top             =   1440
      Visible         =   0   'False
      Width           =   6255
      Begin VB.ComboBox cmbFilenameTokens 
         Height          =   315
         ItemData        =   "frmOptions.frx":08F2
         Left            =   3720
         List            =   "frmOptions.frx":08F4
         Style           =   2  'Dropdown-Liste
         TabIndex        =   93
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txtAutosaveFilename 
         Height          =   285
         Left            =   120
         TabIndex        =   91
         Text            =   "<DateTime>"
         Top             =   1680
         Width           =   3495
      End
      Begin VB.ComboBox cmbAutosaveFormat 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown-Liste
         TabIndex        =   89
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdGetAutosaveDirectory 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   5760
         TabIndex        =   86
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtAutosaveDirectory 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   85
         Top             =   2520
         Width           =   5535
      End
      Begin VB.CheckBox chkUseAutosaveDirectory 
         Caption         =   "For autosave use this directory"
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   2160
         Width           =   5895
      End
      Begin VB.CheckBox chkUseAutosave 
         Caption         =   "Use Autosave"
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   240
         Width           =   6015
      End
      Begin VB.Label lblFilenameTokens 
         Caption         =   "Add a Filename-Token"
         Height          =   255
         Left            =   3720
         TabIndex        =   92
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label lblAutosaveformat 
         Caption         =   "Autosaveformat"
         Height          =   255
         Left            =   120
         TabIndex        =   88
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblAutosaveFilename 
         Caption         =   "Filename"
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   1440
         Width           =   3495
      End
   End
   Begin VB.Frame fraProgGeneral 
      Caption         =   "General"
      Height          =   2415
      Left            =   2880
      TabIndex        =   78
      Top             =   1200
      Width           =   6255
      Begin VB.CommandButton cmdTestpage 
         Caption         =   "Print testpage"
         Height          =   375
         Left            =   120
         TabIndex        =   90
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CheckBox chkUseCreationDateNow 
         Caption         =   "Use the current Date/Time for 'Creation Date'"
         Height          =   255
         Left            =   120
         TabIndex        =   81
         Top             =   1320
         Width           =   5775
      End
      Begin VB.TextBox txtStandardAuthor 
         Height          =   285
         Left            =   120
         TabIndex        =   80
         Top             =   720
         Width           =   6015
      End
      Begin VB.CheckBox chkUseStandardAuthor 
         Caption         =   "Use Standardauthor"
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   7320
      TabIndex        =   1
      Top             =   6000
      Width           =   1575
   End
   Begin MSComctlLib.TreeView trv 
      Height          =   6375
      Left            =   120
      TabIndex        =   77
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   11245
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Frame fraBitmapGeneral 
      Caption         =   "Bitmap"
      Height          =   1935
      Left            =   2760
      TabIndex        =   63
      Top             =   960
      Width           =   6255
      Begin VB.ComboBox cmbTIFFColors 
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown-Liste
         TabIndex        =   75
         Top             =   1440
         Width           =   2175
      End
      Begin VB.ComboBox cmbPCXColors 
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown-Liste
         TabIndex        =   74
         Top             =   1320
         Width           =   2175
      End
      Begin VB.ComboBox cmbBMPColors 
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown-Liste
         TabIndex        =   73
         Top             =   1200
         Width           =   2175
      End
      Begin VB.ComboBox cmbJPEGColors 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown-Liste
         TabIndex        =   72
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtJPEGQuality 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   1920
         TabIndex        =   66
         Text            =   "75"
         Top             =   1440
         Width           =   495
      End
      Begin VB.ComboBox cmbPNGColors 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown-Liste
         TabIndex        =   65
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox txtBitmapResolution 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   1920
         TabIndex        =   64
         Text            =   "72"
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblJPEQQualityProzent 
         Caption         =   "%"
         Height          =   255
         Left            =   2520
         TabIndex        =   71
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label lblJPEGQuality 
         Alignment       =   1  'Rechts
         Caption         =   "Quality:"
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblBitmapColors 
         Alignment       =   1  'Rechts
         Caption         =   "Colors:"
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblBitmapDPI 
         Caption         =   "dpi"
         Height          =   255
         Left            =   2520
         TabIndex        =   68
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblBitmapResolution 
         Alignment       =   1  'Rechts
         Caption         =   "Resolution:"
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.PictureBox picOptions 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2640
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   6000
      Width           =   1575
   End
   Begin MSComctlLib.TabStrip tbstrPDFOptions 
      Height          =   4455
      Left            =   7080
      TabIndex        =   50
      Top             =   4080
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7858
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblOptions 
      Height          =   615
      Left            =   3360
      TabIndex        =   4
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkUseAutosave_Click()
 If chkUseAutosave.Value = 1 Then
   ViewAutosave True
  Else
   ViewAutosave False
 End If
End Sub

Private Sub chkUseAutosaveDirectory_Click()
 If chkUseAutosaveDirectory.Value = 1 Then
   ViewAutosaveDirectory True
  Else
   ViewAutosaveDirectory False
 End If
End Sub

Private Sub chkUseStandardAuthor_Click()
 If chkUseStandardAuthor.Value = 1 Then
   txtStandardAuthor.Enabled = True
   txtStandardAuthor.BackColor = &H80000005
  Else
   txtStandardAuthor.Enabled = False
   txtStandardAuthor.BackColor = &H8000000F
 End If
End Sub

Private Sub cmbCharset_Change()
 On Error GoTo ErrorHandler
 txtTest.Font.Charset = cmbCharset.Text
 Exit Sub
ErrorHandler:
 If Err.number = 380 Then
  cmbCharset.Text = 0
 End If
 Err.Clear
End Sub

Private Sub cmbCharset_Click()
 With cmbCharset
  .Text = .ItemData(.ListIndex)
 End With
 txtTest.Font.Charset = cmbCharset.Text
End Sub

Private Sub cmbCharset_KeyPress(KeyAscii As Integer)
 Dim allow As String, tStr As String
 allow = "0123456789" & Chr$(8) & Chr$(13)
 tStr = Chr$(KeyAscii)
 If InStr(1, allow, tStr) = 0 Then
   KeyAscii = 0
 End If
End Sub

Private Sub cmbFilenameTokens_Click()
 txtAutosaveFilename.Text = txtAutosaveFilename.Text & cmbFilenameTokens.Text
End Sub

Private Sub cmbFonts_Click()
 txtTest.Font.Name = cmbFonts.List(cmbFonts.ListIndex)
End Sub

Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdCancelTest_Click()
 With Options
  SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
  cmbCharset.Text = .ProgramFontCharset
  SetFont frmMain, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
 End With
End Sub

Private Sub cmdGetAutosaveDirectory_Click()
Dim strFolder As String

strFolder = BrowseForFolder(Me.hWnd, LanguageStrings.OptionsAutosaveDirectoryPrompt)
If Len(strFolder) = 0 Then Exit Sub

txtAutosaveDirectory.Text = strFolder
End Sub

Private Sub cmdReset_Click()
 Dim res As Long, Options As tOptions
 res = MsgBox(LanguageStrings.MessagesMsg03, vbYesNo)
 If res = vbYes Then
  Options = StandardOptions
  ShowOptions Me, Options
  With Options
   SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
   cmbCharset.Text = .ProgramFontCharset
   SetFont frmMain, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
  End With
 End If
End Sub

Private Sub cmdSave_Click()
 GetOptions Me, Options
 SaveOptions Options
 Unload Me
End Sub

Private Sub cmdTest_Click()
 Dim tCharset As Long
 tCharset = cmbCharset.Text
 SetFont Me, cmbFonts.List(cmbFonts.ListIndex), CLng(cmbCharset.Text), txtProgramFontsize.Text
 cmbCharset.Text = tCharset
 SetFont frmMain, cmbFonts.List(cmbFonts.ListIndex), CLng(cmbCharset.Text), txtProgramFontsize.Text
 cmdCancelTest.Enabled = True
End Sub

Private Sub cmdTestpage_Click()
 Dim TestPSPage As String, fn As Long, FileName As String
 TestPSPage = LoadResString(3000)
 fn = FreeFile
 FileName = GetTempFile(GetTempPath & "PDFCreator\", "~PD")
 Open FileName For Output As fn
 Print #fn, TestPSPage
 Close #fn
End Sub

Private Sub Form_Load()
 Const fraPDFTop = 1260, fraPDFLeft = 2860
 Dim pic As New StdPicture, i As Long, tStr As String, _
  cSystem As clsSystem, fi As Long, fc As Long, SMF As Collection
 
 Set cSystem = New clsSystem: Set SMF = cSystem.GetSystemFont(Me, Menu)
 
 Screen.MousePointer = vbHourglass
 
 fraBitmapGeneral.Visible = False
 fraBitmapGeneral.Top = fraPDFTop - 300
 fraBitmapGeneral.Left = fraPDFLeft - 200
 fraProgFont.Top = fraPDFTop - 300
 fraProgFont.Left = fraPDFLeft - 200
 fraProgGeneral.Top = fraPDFTop - 300
 fraProgGeneral.Left = fraPDFLeft - 200
 fraProgAutosave.Top = fraPDFTop - 300
 fraProgAutosave.Left = fraPDFLeft - 200
 fraPDFSecurity.Top = fraPDFTop
 fraPDFSecurity.Left = fraPDFLeft
 fraPDFFonts.Top = fraPDFTop
 fraPDFFonts.Left = fraPDFLeft
 fraPDFColors.Top = fraPDFTop
 fraPDFColors.Left = fraPDFLeft
 fraPDFGeneral.Top = fraPDFTop
 fraPDFGeneral.Left = fraPDFLeft
 fraPDFCompress.Top = fraPDFTop
 fraPDFCompress.Left = fraPDFLeft
 tbstrPDFOptions.Top = fraPDFTop - 400
 tbstrPDFOptions.Left = fraPDFLeft - 100
 
 txtTest.Text = ""
 For i = 33 To 255
  txtTest.Text = txtTest.Text & Chr$(i)
 Next i
 fi = -1
 With cmbFonts
  .Clear
  For i = 1 To Screen.FontCount
   tStr = Trim$(Screen.Fonts(i))
   If Len(tStr) > 0 Then
    cmbFonts.AddItem tStr
   End If
  Next i
  If .ListCount > 0 Then
    For i = 0 To cmbFonts.ListCount - 1
     If SMF.Count > 0 Then
      If UCase$(cmbFonts.List(i)) = UCase$(SMF(1)(0)) Then
       fi = i
      End If
     End If
    Next i
   Else
   .ListIndex = 0
  End If
 End With
 With cmbCharset
  .Clear
  .AddItem "0, Western": .ItemData(.NewIndex) = 0
  .AddItem "2, Symbol": .ItemData(.NewIndex) = 2
  .AddItem "77, Mac": .ItemData(.NewIndex) = 77
  .AddItem "161, Greek": .ItemData(.NewIndex) = 161
  .AddItem "162, Turkish": .ItemData(.NewIndex) = 162
  .AddItem "177, Hebrew": .ItemData(.NewIndex) = 177
  .AddItem "178, Arabic": .ItemData(.NewIndex) = 178
  .AddItem "186, Baltic": .ItemData(.NewIndex) = 186
  .AddItem "204, Cyrillic": .ItemData(.NewIndex) = 204
  .AddItem "238, Central European": .ItemData(.NewIndex) = 238
  .AddItem "255, DOS/OEM": .ItemData(.NewIndex) = 255
  .Text = 0
 End With
 If fi >= 0 Then
  cmbFonts.ListIndex = fi
  cmbCharset.Text = SMF(1)(2)
  txtProgramFontsize.Text = SMF(1)(1)
 End If
 txtTest.Font = cmbFonts.List(cmbFonts.ListIndex)
 txtTest.Font.Charset = cmbCharset.Text
 
 With LanguageStrings
  
  trv.Nodes.Clear
  trv.Indentation = 200
  trv.Nodes.Add , , "Program", .OptionsTreeProgram
  trv.Nodes.Add "Program", tvwChild, "ProgramGeneral", .OptionsProgramGeneralSymbol
  trv.Nodes.Add "Program", tvwChild, "ProgramFonts", .OptionsProgramFontSymbol
  trv.Nodes.Add "Program", tvwChild, "ProgramAutosave", .OptionsProgramAutosaveSymbol
  trv.Nodes.Add , , "Formats", .OptionsTreeFormats
  trv.Nodes.Add "Formats", tvwChild, "FormatsPDF", .OptionsPDFSymbol
  trv.Nodes.Add "Formats", tvwChild, "FormatsPNG", .OptionsPNGSymbol
  trv.Nodes.Add "Formats", tvwChild, "FormatsJPEG", .OptionsJPEGSymbol
  trv.Nodes.Add "Formats", tvwChild, "FormatsBMP", .OptionsBMPSymbol
  trv.Nodes.Add "Formats", tvwChild, "FormatsPCX", .OptionsPCXSymbol
  trv.Nodes.Add "Formats", tvwChild, "FormatsTIFF", .OptionsTIFFSymbol
  trv.Nodes("ProgramFonts").EnsureVisible
  trv.Nodes("FormatsPDF").EnsureVisible
  
  Set picOptions = LoadResPicture(2101, vbResIcon)
  fraProgFont.Visible = False
  fraProgGeneral.Visible = True
  
  lblOptions = .OptionsProgramGeneralDescription
  chkUseStandardAuthor.Caption = .OptionsUseStandardauthor
  chkUseCreationDateNow.Caption = .OptionsUseCreationDateNow
  chkUseAutosave.Caption = .OptionsUseAutosave
  cmdTestpage.Caption = .OptionsPrintTestpage
  lblAutosaveFilename.Caption = .OptionsAutosaveFilename
  lblFilenameTokens.Caption = .OptionsAutosaveFilenameTokens
  chkUseAutosaveDirectory.Caption = .OptionsUseAutosaveDirectory
  With cmbFilenameTokens
   .AddItem "<Author>"
   .AddItem "<Computername>"
   .AddItem "<DateTime>"
   .AddItem "<Title>"
   .AddItem "<Username>"
   .ListIndex = 0
  End With
  With cmbAutosaveFormat
   .AddItem "PDF"
   .AddItem "PNG"
   .AddItem "JPEG"
   .AddItem "BMP"
   .AddItem "PCX"
   .AddItem "TIFF"
  End With
  Me.Caption = .OptionsPDFOptions
  cmdCancel.Caption = .OptionsCancel
  cmdReset.Caption = .OptionsReset
  cmdSave.Caption = .OptionsSave
  tbstrPDFOptions.Tabs.Clear
  tbstrPDFOptions.Tabs.Add , "General", .OptionsPDFGeneral
  tbstrPDFOptions.Tabs.Add , "Compression", .OptionsPDFCompression
  tbstrPDFOptions.Tabs.Add , "Fonts", .OptionsPDFFonts
  tbstrPDFOptions.Tabs.Add , "Colors", .OptionsPDFColors
  tbstrPDFOptions.Tabs.Add , "Security", .OptionsPDFSecurity
  fraPDFGeneral.Caption = .OptionsPDFGeneralCaption
  lblPDFCompat.Caption = .OptionsPDFGeneralCompatibility
  lblPDFAutoRotate.Caption = .OptionsPDFGeneralAutorotate
  lblPDFResolution.Caption = .OptionsPDFGeneralResolution
  lblPDFOverprint.Caption = .OptionsPDFGeneralOverprint
  chkPDFASCII85.Caption = .OptionsPDFGeneralASCII85
  cmbPDFCompat.Clear
  cmbPDFCompat.AddItem .OptionsPDFGeneralCompatibility01
  cmbPDFCompat.AddItem .OptionsPDFGeneralCompatibility02
  cmbPDFCompat.AddItem .OptionsPDFGeneralCompatibility03
  cmbPDFRotate.Clear
  cmbPDFRotate.AddItem .OptionsPDFGeneralRotate01
  cmbPDFRotate.AddItem .OptionsPDFGeneralRotate02
  cmbPDFRotate.AddItem .OptionsPDFGeneralRotate03
  cmbPDFOverprint.Clear
  cmbPDFOverprint.AddItem .OptionsPDFGeneralOverprint01
  cmbPDFOverprint.AddItem .OptionsPDFGeneralOverprint02
  
  fraPDFCompress.Caption = .OptionsPDFCompressionCaption
  chkPDFTextComp.Caption = .OptionsPDFCompressionTextComp
  fraPDFColor.Caption = .OptionsPDFCompressionColor
  chkPDFColorComp.Caption = .OptionsPDFCompressionColorComp
  chkPDFColorResample.Caption = .OptionsPDFCompressionColorResample
  lblPDFColorRes.Caption = .OptionsPDFCompressionColorRes
  cmbPDFColorComp.Clear
  cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp01
  cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp02
  cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp03
  cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp04
  cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp05
  cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp06
  cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp07
  cmbPDFColorResample.Clear
  cmbPDFColorResample.AddItem .OptionsPDFCompressionColorResample01
  cmbPDFColorResample.AddItem .OptionsPDFCompressionColorResample02
  cmbPDFColorResample.AddItem .OptionsPDFCompressionColorResample03
  fraPDFGrey.Caption = .OptionsPDFCompressionGrey
  chkPDFGreyComp.Caption = .OptionsPDFCompressionGreyComp
  chkPDFGreyResample.Caption = .OptionsPDFCompressionGreyResample
  lblPDFGreyRes.Caption = .OptionsPDFCompressionGreyRes
  cmbPDFGreyComp.Clear
  cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp01
  cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp02
  cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp03
  cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp04
  cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp05
  cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp06
  cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp07
  cmbPDFGreyResample.Clear
  cmbPDFGreyResample.AddItem .OptionsPDFCompressionGreyResample01
  cmbPDFGreyResample.AddItem .OptionsPDFCompressionGreyResample02
  cmbPDFGreyResample.AddItem .OptionsPDFCompressionGreyResample03
  fraPDFMono.Caption = .OptionsPDFCompressionMono
  chkPDFMonoComp.Caption = .OptionsPDFCompressionMonoComp
  chkPDFMonoResample.Caption = .OptionsPDFCompressionMonoResample
  lblPDFMonoRes.Caption = .OptionsPDFCompressionMonoRes
  cmbPDFMonoComp.Clear
  cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp01
  cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp02
  cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp03
  cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp04
  cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp05
  cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp06
  cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp07
  cmbPDFMonoResample.Clear
  cmbPDFMonoResample.AddItem .OptionsPDFCompressionMonoResample01
  cmbPDFMonoResample.AddItem .OptionsPDFCompressionMonoResample02
  cmbPDFMonoResample.AddItem .OptionsPDFCompressionMonoResample03
  
  fraPDFFonts.Caption = .OptionsPDFFontsCaption
  chkPDFEmbedAll.Caption = .OptionsPDFFontsEmbedAll
  chkPDFSubSetFonts.Caption = .OptionsPDFFontsSubSetFonts
  
  fraPDFColors.Caption = .OptionsPDFColorsCaption
  chkPDFCMYKtoRGB.Caption = .OptionsPDFColorsCMYKtoRGB
  fraPDFColorOptions.Caption = .OptionsPDFColorsColorOptions
  chkPDFPreserveOverprint.Caption = .OptionsPDFColorsPreserveOverprint
  chkPDFPreserveTransfer.Caption = .OptionsPDFColorsPreserveTransfer
  chkPDFPreserveHalftone.Caption = .OptionsPDFColorsPreserveHalftone
  cmbPDFColorModel.Clear
  cmbPDFColorModel.AddItem .OptionsPDFColorsColorModel01
  cmbPDFColorModel.AddItem .OptionsPDFColorsColorModel02
  cmbPDFColorModel.AddItem .OptionsPDFColorsColorModel03
  
  fraPDFSecurity.Caption = .OptionsPDFSecurityCaption
 
  cmbPNGColors.AddItem .OptionsPNGColorscount01
  cmbPNGColors.AddItem .OptionsPNGColorscount02
  cmbPNGColors.AddItem .OptionsPNGColorscount03
  cmbPNGColors.AddItem .OptionsPNGColorscount04
  cmbJPEGColors.Left = cmbPNGColors.Left
  cmbJPEGColors.Width = cmbPNGColors.Width
  cmbJPEGColors.Top = cmbPNGColors.Top
  cmbJPEGColors.AddItem .OptionsJPEGColorscount01
  cmbJPEGColors.AddItem .OptionsJPEGColorscount02
  cmbBMPColors.Left = cmbPNGColors.Left
  cmbBMPColors.Width = cmbPNGColors.Width
  cmbBMPColors.Top = cmbPNGColors.Top
  cmbBMPColors.AddItem .OptionsBMPColorscount01
  cmbBMPColors.AddItem .OptionsBMPColorscount02
  cmbBMPColors.AddItem .OptionsBMPColorscount03
  cmbBMPColors.AddItem .OptionsBMPColorscount04
  cmbBMPColors.AddItem .OptionsBMPColorscount05
  cmbBMPColors.AddItem .OptionsBMPColorscount06
  cmbBMPColors.AddItem .OptionsBMPColorscount07
  cmbPCXColors.Left = cmbPNGColors.Left
  cmbPCXColors.Width = cmbPNGColors.Width
  cmbPCXColors.Top = cmbPNGColors.Top
  cmbPCXColors.AddItem .OptionsPCXColorscount01
  cmbPCXColors.AddItem .OptionsPCXColorscount02
  cmbPCXColors.AddItem .OptionsPCXColorscount03
  cmbPCXColors.AddItem .OptionsPCXColorscount04
  cmbPCXColors.AddItem .OptionsPCXColorscount05
  cmbPCXColors.AddItem .OptionsPCXColorscount06
  cmbTIFFColors.Left = cmbPNGColors.Left
  cmbTIFFColors.Width = cmbPNGColors.Width
  cmbTIFFColors.Top = cmbPNGColors.Top
  cmbTIFFColors.AddItem .OptionsTIFFColorscount01
  cmbTIFFColors.AddItem .OptionsTIFFColorscount02
  cmbTIFFColors.AddItem .OptionsTIFFColorscount03
  cmbTIFFColors.AddItem .OptionsTIFFColorscount04
  cmbTIFFColors.AddItem .OptionsTIFFColorscount05
  cmbTIFFColors.AddItem .OptionsTIFFColorscount06
  cmbTIFFColors.AddItem .OptionsTIFFColorscount07
  cmbTIFFColors.AddItem .OptionsTIFFColorscount08
 End With
 
 txtPDFRes.Text = 600
 cmbPDFCompat.ListIndex = 1
 cmbPDFRotate.ListIndex = 0
 cmbPDFOverprint.ListIndex = 0
 chkPDFASCII85.Value = 0
 
 chkPDFTextComp.Value = 1

 chkPDFColorComp.Value = 1
 chkPDFColorResample.Value = 0
 cmbPDFColorComp.ListIndex = 0
 cmbPDFColorResample.ListIndex = 0
 txtPDFColorRes.Text = 300

 chkPDFGreyComp.Value = 1
 chkPDFGreyResample.Value = 0
 cmbPDFGreyComp.ListIndex = 0
 cmbPDFGreyResample.ListIndex = 0
 txtPDFGreyRes.Text = 300

 chkPDFMonoComp.Value = 1
 chkPDFMonoResample.Value = 0
 cmbPDFMonoComp.ListIndex = 0
 cmbPDFMonoResample.ListIndex = 0
 txtPDFMonoRes.Text = 1200
 
 chkPDFEmbedAll.Value = 1
 chkPDFSubSetFonts.Value = 1
 txtPDFSubSetPerc.Text = 100

 cmbPDFColorModel.ListIndex = 1
 chkPDFCMYKtoRGB.Value = 1
 chkPDFPreserveOverprint.Value = 1
 chkPDFPreserveTransfer.Value = 1
 chkPDFPreserveHalftone.Value = 0
  
 cmbPNGColors.ListIndex = 0
 cmbJPEGColors.ListIndex = 0
 cmbBMPColors.ListIndex = 0
 cmbPCXColors.ListIndex = 0
 cmbTIFFColors.ListIndex = 0
 txtBitmapResolution.Text = 150
 
 cmbCharset.ListIndex = 0
 txtProgramFontsize.Text = 8
 
' chkUseStandardAuthor.Value = 1
 txtStandardAuthor.Text = ""
 
 ShowOptions Me, Options
 If chkUseStandardAuthor.Value = 1 Then
   txtStandardAuthor.Enabled = True
   txtStandardAuthor.BackColor = &H80000005
  Else
   txtStandardAuthor.Enabled = False
   txtStandardAuthor.BackColor = &H8000000F
 End If
 With Options
  SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
  cmbCharset.Text = .ProgramFontCharset
 End With
 If chkUseAutosave.Value = 1 Then
   ViewAutosave True
  Else
   ViewAutosave False
 End If

 Screen.MousePointer = vbNormal
End Sub

Private Sub tbstrPDFOptions_Click()
 fraPDFGeneral.Visible = False
 fraPDFCompress.Visible = False
 fraPDFFonts.Visible = False
 fraPDFColors.Visible = False
 fraPDFSecurity.Visible = False
 Select Case tbstrPDFOptions.SelectedItem.Index
  Case 1:
   fraPDFGeneral.Visible = True
  Case 2:
   fraPDFCompress.Visible = True
  Case 3:
   fraPDFFonts.Visible = True
  Case 4:
   fraPDFColors.Visible = True
  Case 5:
   fraPDFSecurity.Visible = True
 End Select
End Sub


Private Sub txtAutosaveFilename_Change()
 txtAutosaveFilename.ToolTipText = txtAutosaveFilename.Text
End Sub

Private Sub txtProgramFontSize_Change()
 Dim tL As Long
If Trim$(txtProgramFontsize.Text) = "" Then
  txtProgramFontsize.Text = 8
 End If
 tL = CLng(txtProgramFontsize.Text)
 If tL <= 0 Then
  tL = 1
 End If
 If tL > 72 Then
  tL = 72
 End If
 txtProgramFontsize.Text = tL
 txtTest.Font.Size = tL
End Sub

Private Sub txtProgramFontSize_KeyPress(KeyAscii As Integer)
 Dim allow As String, tStr As String
    
 allow = "0123456789" & Chr$(8) & Chr$(13)
    
 tStr = Chr$(KeyAscii)
 
 If InStr(1, allow, tStr) = 0 Then
   KeyAscii = 0
 End If
End Sub

Private Sub trv_Click()
lblJPEGQuality.Visible = False
cmbPNGColors.Visible = False
cmbJPEGColors.Visible = False
cmbBMPColors.Visible = False
cmbPCXColors.Visible = False
cmbTIFFColors.Visible = False
tbstrPDFOptions.Visible = False
fraProgFont.Visible = False
fraProgGeneral.Visible = False
fraProgAutosave.Visible = False
fraBitmapGeneral.Visible = False
fraPDFGeneral.Visible = False
fraPDFCompress.Visible = False
fraPDFFonts.Visible = False
fraPDFColors.Visible = False
fraPDFSecurity.Visible = False
txtJPEGQuality.Visible = False
lblJPEQQualityProzent.Visible = False

 Select Case trv.SelectedItem.Key
  Case "Program"
   Set picOptions = LoadResPicture(2101, vbResIcon)
   lblOptions = LanguageStrings.OptionsProgramGeneralDescription
   fraProgGeneral.Visible = True
  Case "ProgramGeneral"
   Set picOptions = LoadResPicture(2101, vbResIcon)
   lblOptions = LanguageStrings.OptionsProgramGeneralDescription
   fraProgGeneral.Visible = True
  Case "ProgramFonts"
   Set picOptions = LoadResPicture(2102, vbResIcon)
   lblOptions = LanguageStrings.OptionsProgramFontDescription
   fraProgFont.Visible = True
  Case "ProgramAutosave"
   Set picOptions = LoadResPicture(2103, vbResIcon)
   lblOptions = LanguageStrings.OptionsProgramAutosaveDescription
   fraProgAutosave.Visible = True
  Case "Formats"
   Set picOptions = LoadResPicture(2111, vbResIcon)
   lblOptions = LanguageStrings.OptionsPDFDescription
   tbstrPDFOptions.Visible = True
   fraPDFGeneral.Visible = True
  Case "FormatsPDF"
   Set picOptions = LoadResPicture(2111, vbResIcon)
   lblOptions = LanguageStrings.OptionsPDFDescription
   tbstrPDFOptions.Visible = True
   fraPDFGeneral.Visible = True
  Case "FormatsPNG"
   Set picOptions = LoadResPicture(2112, vbResIcon)
   lblOptions = LanguageStrings.OptionsPNGDescription
   fraBitmapGeneral.Visible = True
   cmbPNGColors.Visible = True
  Case "FormatsJPEG"
   Set picOptions = LoadResPicture(2113, vbResIcon)
   lblOptions = LanguageStrings.OptionsJPEGDescription
   fraBitmapGeneral.Visible = True
   lblJPEGQuality.Caption = LanguageStrings.OptionsJPEGQuality
   lblJPEGQuality.Visible = True
   txtJPEGQuality.Visible = True
   lblJPEQQualityProzent.Visible = True
   lblJPEQQualityProzent.Left = txtJPEGQuality.Left + txtJPEGQuality.Width + 100
   cmbJPEGColors.Visible = True
  Case "FormatsBMP"
   Set picOptions = LoadResPicture(2114, vbResIcon)
   lblOptions = LanguageStrings.OptionsBMPDescription
   fraBitmapGeneral.Visible = True
   cmbBMPColors.Visible = True
  Case "FormatsPCX"
   Set picOptions = LoadResPicture(2115, vbResIcon)
   lblOptions = LanguageStrings.OptionsPCXDescription
   fraBitmapGeneral.Visible = True
   cmbPCXColors.Visible = True
  Case "FormatsTIFF"
   Set picOptions = LoadResPicture(2116, vbResIcon)
   lblOptions = LanguageStrings.OptionsTIFFDescription
   fraBitmapGeneral.Visible = True
   cmbTIFFColors.Visible = True
 End Select
End Sub

Private Sub ViewAutosave(ViewIt As Boolean)
 lblAutosaveformat.Enabled = ViewIt
 cmbAutosaveFormat.Enabled = ViewIt
 lblAutosaveFilename.Enabled = ViewIt
 txtAutosaveFilename.Enabled = ViewIt
 lblFilenameTokens.Enabled = ViewIt
 cmbFilenameTokens.Enabled = ViewIt
 chkUseAutosaveDirectory.Enabled = ViewIt
 If ViewIt = True Then
   cmbAutosaveFormat.BackColor = &H80000005
   cmbFilenameTokens.BackColor = &H80000005
   txtAutosaveFilename.BackColor = &H80000005
  Else
   cmbAutosaveFormat.BackColor = &H8000000F
   cmbFilenameTokens.BackColor = &H8000000F
   txtAutosaveFilename.BackColor = &H8000000F
 End If
 If chkUseAutosaveDirectory.Value = 1 And ViewIt = True Then
   ViewAutosaveDirectory True
  Else
   ViewAutosaveDirectory False
 End If
End Sub

Private Sub ViewAutosaveDirectory(ViewIt As Boolean)
 If ViewIt = True Then
   txtAutosaveDirectory.Enabled = True
   txtAutosaveDirectory.BackColor = &HC0FFFF
   cmdGetAutosaveDirectory.Enabled = True
  Else
   txtAutosaveDirectory.Enabled = False
   txtAutosaveDirectory.BackColor = &H8000000F
   cmdGetAutosaveDirectory.Enabled = False
 End If
End Sub
