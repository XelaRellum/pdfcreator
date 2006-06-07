VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Options"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9165
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin PDFCreator.dmFrame dmFraProgGhostscript 
      Height          =   3150
      Left            =   2625
      TabIndex        =   13
      Top             =   945
      Visible         =   0   'False
      Width           =   6420
      _extentx        =   11324
      _extenty        =   5556
      caption         =   "Ghostscript"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":000C
      textshaddowcolor=   12582912
      Begin VB.CheckBox chkAddWindowsFontpath 
         Appearance      =   0  '2D
         Caption         =   "Add Windows fontpath"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   105
         TabIndex        =   177
         Top             =   2730
         Width           =   6105
      End
      Begin VB.TextBox txtAdditionalGhostscriptSearchpath 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   105
         TabIndex        =   175
         Top             =   2100
         Width           =   6105
      End
      Begin VB.TextBox txtAdditionalGhostscriptParameters 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   105
         TabIndex        =   174
         Top             =   1365
         Width           =   6105
      End
      Begin VB.ComboBox cmbGhostscript 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   105
         Style           =   2  'Dropdown-Liste
         TabIndex        =   22
         Top             =   630
         Width           =   4215
      End
      Begin VB.CommandButton cmdGetgsbinDirectory 
         Caption         =   "..."
         Height          =   255
         Left            =   5625
         TabIndex        =   21
         Top             =   3690
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtGSbin 
         Appearance      =   0  '2D
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   3690
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.TextBox txtGSlib 
         Appearance      =   0  '2D
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   4290
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.TextBox txtGSfonts 
         Appearance      =   0  '2D
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   4890
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.CommandButton cmdGetgslibDirectory 
         Caption         =   "..."
         Height          =   255
         Left            =   5625
         TabIndex        =   17
         Top             =   4290
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdGetgsfontsDirectory 
         Caption         =   "..."
         Height          =   255
         Left            =   5625
         TabIndex        =   16
         Top             =   4890
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtGSresource 
         Appearance      =   0  '2D
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   5490
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.CommandButton cmdGetgsresourceDirectory 
         Caption         =   "..."
         Height          =   255
         Left            =   5625
         TabIndex        =   14
         Top             =   5490
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblAdditionalGhostscriptSearchpath 
         AutoSize        =   -1  'True
         Caption         =   "Additional Ghostscript searchpath"
         Height          =   195
         Left            =   105
         TabIndex        =   176
         Top             =   1890
         Width           =   2370
      End
      Begin VB.Label lblAdditionalGhostscriptParameters 
         AutoSize        =   -1  'True
         Caption         =   "Additional Ghostscript parameters"
         Height          =   195
         Left            =   105
         TabIndex        =   173
         Top             =   1155
         Width           =   2355
      End
      Begin VB.Label lblGhostscriptversion 
         AutoSize        =   -1  'True
         Caption         =   "Ghostscriptversion"
         Height          =   195
         Left            =   105
         TabIndex        =   27
         Top             =   420
         Width           =   1305
      End
      Begin VB.Label lblGSbin 
         AutoSize        =   -1  'True
         Caption         =   "Ghostscript Binaries"
         Height          =   195
         Left            =   105
         TabIndex        =   26
         Top             =   3450
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label lblGSlib 
         AutoSize        =   -1  'True
         Caption         =   "Ghostscript Libraries"
         Height          =   195
         Left            =   105
         TabIndex        =   25
         Top             =   4050
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label lblGSfonts 
         AutoSize        =   -1  'True
         Caption         =   "Ghostscript Fonts"
         Height          =   195
         Left            =   105
         TabIndex        =   24
         Top             =   4650
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label lblGhostscriptResource 
         AutoSize        =   -1  'True
         Caption         =   "Ghostscript Resource"
         Height          =   195
         Left            =   105
         TabIndex        =   23
         Top             =   5250
         Visible         =   0   'False
         Width           =   1530
      End
   End
   Begin PDFCreator.dmFrame dmFraProgAutosave 
      Height          =   4485
      Left            =   2640
      TabIndex        =   31
      Top             =   1200
      Visible         =   0   'False
      Width           =   6375
      _extentx        =   11245
      _extenty        =   7911
      caption         =   "Autosave"
      caption3d       =   2
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":0038
      textshaddowcolor=   12582912
      Begin VB.CheckBox chkAutosaveStartStandardProgram 
         Appearance      =   0  '2D
         Caption         =   "After auto-saving open the document with the default program."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   105
         TabIndex        =   172
         Top             =   4095
         Width           =   5895
      End
      Begin VB.CommandButton cmdGetAutosaveDirectory 
         Caption         =   "..."
         Height          =   300
         Left            =   5760
         TabIndex        =   168
         Top             =   3120
         Width           =   375
      End
      Begin VB.ComboBox cmbAutosaveFormat 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown-Liste
         TabIndex        =   39
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox cmbAutoSaveFilenameTokens 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "frmOptions.frx":0064
         Left            =   3690
         List            =   "frmOptions.frx":0066
         Style           =   2  'Dropdown-Liste
         TabIndex        =   38
         Top             =   1785
         Width           =   2460
      End
      Begin VB.TextBox txtAutosaveFilename 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   120
         TabIndex        =   37
         Text            =   "<DateTime>"
         Top             =   1800
         Width           =   3495
      End
      Begin VB.TextBox txtAutosaveDirectory 
         Appearance      =   0  '2D
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   36
         Top             =   3120
         Width           =   5535
      End
      Begin VB.CheckBox chkUseAutosaveDirectory 
         Appearance      =   0  '2D
         Caption         =   "For autosave use this directory"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   2760
         Width           =   5895
      End
      Begin VB.CheckBox chkUseAutosave 
         Appearance      =   0  '2D
         Caption         =   "Use Autosave"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   6015
      End
      Begin VB.TextBox txtAutoSaveFilenamePreview 
         Appearance      =   0  '2D
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   2145
         Width           =   6015
      End
      Begin VB.TextBox txtAutoSaveDirectoryPreview 
         Appearance      =   0  '2D
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   3450
         Width           =   6015
      End
      Begin VB.Label lblAutosaveFilenameTokens 
         AutoSize        =   -1  'True
         Caption         =   "Add a Filename-Token"
         Height          =   195
         Left            =   3720
         TabIndex        =   42
         Top             =   1560
         Width           =   1605
      End
      Begin VB.Label lblAutosaveformat 
         Caption         =   "Autosaveformat"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label lblAutosaveFilename 
         AutoSize        =   -1  'True
         Caption         =   "Filename"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   1560
         Width           =   630
      End
   End
   Begin PDFCreator.dmFrame dmFraPDFGeneral 
      Height          =   2895
      Left            =   2730
      TabIndex        =   91
      Top             =   1785
      Visible         =   0   'False
      Width           =   6195
      _extentx        =   10927
      _extenty        =   5106
      caption         =   "General Options"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":0068
      Begin VB.CheckBox chkPDFASCII85 
         Appearance      =   0  '2D
         Caption         =   "Convert binary data to ASCII85"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   96
         Top             =   2520
         Width           =   3675
      End
      Begin VB.ComboBox cmbPDFOverprint 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "frmOptions.frx":0094
         Left            =   2400
         List            =   "frmOptions.frx":0096
         Style           =   2  'Dropdown-Liste
         TabIndex        =   95
         Top             =   1980
         Width           =   2655
      End
      Begin VB.TextBox txtPDFRes 
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         Height          =   285
         Left            =   2400
         TabIndex        =   94
         Text            =   "600"
         Top             =   1440
         Width           =   615
      End
      Begin VB.ComboBox cmbPDFCompat 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "frmOptions.frx":0098
         Left            =   2400
         List            =   "frmOptions.frx":009A
         Style           =   2  'Dropdown-Liste
         TabIndex        =   93
         Top             =   480
         Width           =   2655
      End
      Begin VB.ComboBox cmbPDFRotate 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "frmOptions.frx":009C
         Left            =   2400
         List            =   "frmOptions.frx":009E
         Style           =   2  'Dropdown-Liste
         TabIndex        =   92
         Tag             =   "None|All|PageByPage"
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label lblPDFDPI 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "dpi"
         Height          =   195
         Left            =   3120
         TabIndex        =   101
         Top             =   1485
         Width           =   210
      End
      Begin VB.Label lblPDFOverprint 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         Caption         =   "Overprint:"
         Height          =   195
         Left            =   1605
         TabIndex        =   100
         Top             =   2040
         Width           =   690
      End
      Begin VB.Label lblPDFResolution 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resolution:"
         Height          =   195
         Left            =   1500
         TabIndex        =   99
         Top             =   1485
         Width           =   795
      End
      Begin VB.Label lblPDFCompat 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compatibility:"
         Height          =   195
         Left            =   1380
         TabIndex        =   98
         Top             =   540
         Width           =   915
      End
      Begin VB.Label lblPDFAutoRotate 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Auto-Rotate Pages:"
         Height          =   195
         Left            =   900
         TabIndex        =   97
         Top             =   1020
         Width           =   1395
      End
   End
   Begin PDFCreator.dmFrame dmFraBitmapGeneral 
      Height          =   1935
      Left            =   2640
      TabIndex        =   74
      Top             =   1680
      Visible         =   0   'False
      Width           =   6375
      _extentx        =   11245
      _extenty        =   3413
      caption         =   "Bitmap"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":00A0
      Begin VB.ComboBox cmbTIFFColors 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown-Liste
         TabIndex        =   81
         Top             =   1440
         Width           =   2175
      End
      Begin VB.ComboBox cmbPCXColors 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown-Liste
         TabIndex        =   80
         Top             =   1320
         Width           =   2175
      End
      Begin VB.ComboBox cmbBMPColors 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown-Liste
         TabIndex        =   79
         Top             =   1200
         Width           =   2175
      End
      Begin VB.ComboBox cmbJPEGColors 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown-Liste
         TabIndex        =   78
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtJPEGQuality 
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         Height          =   285
         Left            =   1920
         TabIndex        =   77
         Text            =   "75"
         Top             =   1440
         Width           =   495
      End
      Begin VB.ComboBox cmbPNGColors 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown-Liste
         TabIndex        =   76
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox txtBitmapResolution 
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         Height          =   285
         Left            =   1920
         TabIndex        =   75
         Text            =   "72"
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblJPEQQualityProzent 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   2520
         TabIndex        =   86
         Top             =   1485
         Width           =   120
      End
      Begin VB.Label lblJPEGQuality 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         Caption         =   "Quality:"
         Height          =   195
         Left            =   1290
         TabIndex        =   85
         Top             =   1485
         Width           =   525
      End
      Begin VB.Label lblBitmapColors 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         Caption         =   "Colors:"
         Height          =   195
         Left            =   1335
         TabIndex        =   84
         Top             =   1020
         Width           =   480
      End
      Begin VB.Label lblBitmapDPI 
         AutoSize        =   -1  'True
         Caption         =   "dpi"
         Height          =   195
         Left            =   2520
         TabIndex        =   83
         Top             =   525
         Width           =   210
      End
      Begin VB.Label lblBitmapResolution 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         Caption         =   "Resolution:"
         Height          =   195
         Left            =   1020
         TabIndex        =   82
         Top             =   525
         Width           =   795
      End
   End
   Begin PDFCreator.dmFrame dmFraPSGeneral 
      Height          =   1095
      Left            =   2640
      TabIndex        =   87
      Top             =   1920
      Visible         =   0   'False
      Width           =   6375
      _extentx        =   11245
      _extenty        =   1931
      caption         =   "Postscript"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":00CC
      Begin VB.ComboBox cmbEPSLanguageLevel 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown-Liste
         TabIndex        =   89
         Top             =   600
         Width           =   735
      End
      Begin VB.ComboBox cmbPSLanguageLevel 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown-Liste
         TabIndex        =   88
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblLangLevel 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         Caption         =   "Language Level:"
         Height          =   195
         Left            =   735
         TabIndex        =   90
         Top             =   510
         Width           =   1200
      End
   End
   Begin PDFCreator.dmFrame dmFraPDFFonts 
      Height          =   1695
      Left            =   2760
      TabIndex        =   125
      Top             =   2400
      Visible         =   0   'False
      Width           =   6195
      _extentx        =   10927
      _extenty        =   2990
      caption         =   "Font options"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":00F8
      Begin VB.CheckBox chkPDFEmbedAll 
         Appearance      =   0  '2D
         Caption         =   "Embed all Fonts"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   128
         Top             =   360
         Width           =   5955
      End
      Begin VB.CheckBox chkPDFSubSetFonts 
         Appearance      =   0  '2D
         Caption         =   "Subset Fonts, when percentage of used characters below:"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   127
         Top             =   780
         Width           =   5955
      End
      Begin VB.TextBox txtPDFSubSetPerc 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   400
         TabIndex        =   126
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblPDFPerc 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   960
         TabIndex        =   129
         Top             =   1365
         Width           =   120
      End
   End
   Begin PDFCreator.dmFrame dmFraProgDirectories 
      Height          =   1410
      Left            =   2640
      TabIndex        =   43
      Top             =   1320
      Visible         =   0   'False
      Width           =   6495
      _extentx        =   11456
      _extenty        =   2487
      caption         =   "Directories"
      caption3d       =   2
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":0124
      textshaddowcolor=   12582912
      Begin VB.TextBox txtTemppathPreview 
         Appearance      =   0  '2D
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   171
         Top             =   945
         Width           =   5910
      End
      Begin VB.CommandButton cmdGetTemppath 
         Caption         =   "..."
         Height          =   300
         Left            =   5154
         TabIndex        =   166
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtTemppath 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   120
         TabIndex        =   44
         Top             =   600
         Width           =   4965
      End
      Begin VB.CommandButton cmdUsertempPath 
         Height          =   300
         Left            =   5640
         Picture         =   "frmOptions.frx":0150
         Style           =   1  'Grafisch
         TabIndex        =   167
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblPrintTempPath 
         AutoSize        =   -1  'True
         Caption         =   "Temppath"
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   360
         Width           =   720
      End
   End
   Begin PDFCreator.dmFrame dmfraProgSave 
      Height          =   1935
      Left            =   2640
      TabIndex        =   52
      Top             =   2160
      Visible         =   0   'False
      Width           =   6375
      _extentx        =   11245
      _extenty        =   3413
      caption         =   "Save"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":06DA
      Begin VB.CheckBox chkSpaces 
         Appearance      =   0  '2D
         Caption         =   "Remove leading and trailing spaces"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   1440
         Value           =   1  'Aktiviert
         Width           =   6015
      End
      Begin VB.TextBox txtSaveFilename 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   120
         TabIndex        =   55
         Text            =   "<Title>"
         Top             =   600
         Width           =   3495
      End
      Begin VB.ComboBox cmbSaveFilenameTokens 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "frmOptions.frx":0706
         Left            =   3720
         List            =   "frmOptions.frx":0708
         Style           =   2  'Dropdown-Liste
         TabIndex        =   54
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtSavePreview 
         Appearance      =   0  '2D
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   960
         Width           =   6015
      End
      Begin VB.Label lblSaveFilename 
         AutoSize        =   -1  'True
         Caption         =   "Filename"
         Height          =   195
         Left            =   120
         TabIndex        =   58
         Top             =   360
         Width           =   630
      End
      Begin VB.Label lblSaveFilenameTokens 
         AutoSize        =   -1  'True
         Caption         =   "Add a Filename-Token"
         Height          =   195
         Left            =   3720
         TabIndex        =   57
         Top             =   360
         Width           =   1605
      End
   End
   Begin PDFCreator.dmFrame dmFraProgDocument 
      Height          =   1935
      Left            =   2640
      TabIndex        =   46
      Top             =   1800
      Visible         =   0   'False
      Width           =   6375
      _extentx        =   11245
      _extenty        =   3413
      caption         =   "Document"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":070A
      Begin VB.ComboBox cmbAuthorTokens 
         Appearance      =   0  '2D
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmOptions.frx":0736
         Left            =   3720
         List            =   "frmOptions.frx":0738
         Style           =   2  'Dropdown-Liste
         TabIndex        =   51
         Top             =   840
         Width           =   2415
      End
      Begin VB.CheckBox chkUseStandardAuthor 
         Appearance      =   0  '2D
         Caption         =   "Use Standardauthor"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   360
         Width           =   6000
      End
      Begin VB.TextBox txtStandardAuthor 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   120
         TabIndex        =   49
         Top             =   840
         Width           =   3495
      End
      Begin VB.CheckBox chkUseCreationDateNow 
         Appearance      =   0  '2D
         Caption         =   "Use the current Date/Time for 'Creation Date'"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   1440
         Width           =   6000
      End
      Begin VB.Label lblAuthorTokens 
         AutoSize        =   -1  'True
         Caption         =   "Add a Author-Token"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3720
         TabIndex        =   47
         Top             =   600
         Width           =   1440
      End
   End
   Begin PDFCreator.dmFrame dmFraProgFont 
      Height          =   4695
      Left            =   2640
      TabIndex        =   65
      Top             =   1440
      Visible         =   0   'False
      Width           =   6375
      _extentx        =   11245
      _extenty        =   8281
      caption         =   "Programfont"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":073A
      Begin VB.CommandButton cmdCancelTest 
         Caption         =   "C&ancel test"
         Height          =   495
         Left            =   2310
         TabIndex        =   170
         Top             =   4095
         Width           =   1755
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "&Test"
         Height          =   495
         Left            =   120
         TabIndex        =   169
         Top             =   4095
         Width           =   1755
      End
      Begin VB.TextBox txtTest 
         Appearance      =   0  '2D
         Height          =   2655
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   69
         Top             =   1320
         Width           =   6135
      End
      Begin VB.ComboBox cmbCharset 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   3000
         TabIndex        =   68
         Text            =   "cmbCharset"
         Top             =   600
         Width           =   2295
      End
      Begin VB.ComboBox cmbFonts 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown-Liste
         TabIndex        =   67
         Top             =   600
         Width           =   2775
      End
      Begin VB.ComboBox cmbProgramFontsize 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   5400
         TabIndex        =   66
         Text            =   "8"
         Top             =   600
         Width           =   765
      End
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         Caption         =   "Size"
         Height          =   195
         Left            =   5400
         TabIndex        =   73
         Top             =   360
         Width           =   300
      End
      Begin VB.Label lblTesttext 
         AutoSize        =   -1  'True
         Caption         =   "Here you can test the font."
         Height          =   195
         Left            =   120
         TabIndex        =   72
         Top             =   1080
         Width           =   1890
      End
      Begin VB.Label lblProgcharset 
         AutoSize        =   -1  'True
         Caption         =   "Charset"
         Height          =   195
         Left            =   3000
         TabIndex        =   71
         Top             =   360
         Width           =   540
      End
      Begin VB.Label lblProgfont 
         AutoSize        =   -1  'True
         Caption         =   "Programfont"
         Height          =   195
         Left            =   120
         TabIndex        =   70
         Top             =   360
         Width           =   855
      End
   End
   Begin PDFCreator.dmFrame dmFraPDFSecurity 
      Height          =   5535
      Left            =   2730
      TabIndex        =   137
      Top             =   2205
      Visible         =   0   'False
      Width           =   6195
      _extentx        =   10927
      _extenty        =   9763
      caption         =   "Security"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":0766
      Begin PDFCreator.dmFrame dmFraPDFHighPermissions 
         Height          =   855
         Left            =   120
         TabIndex        =   152
         Top             =   4560
         Width           =   5955
         _extentx        =   10504
         _extenty        =   1508
         caption         =   "Enhanced permissions (128 Bit only)"
         barcolorfrom    =   16744576
         barcolorto      =   4194304
         font            =   "frmOptions.frx":0792
         Begin VB.CheckBox chkAllowAssembly 
            Appearance      =   0  '2D
            Caption         =   "Allow changes to the Assembly"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3100
            TabIndex        =   156
            Top             =   525
            Width           =   2760
         End
         Begin VB.CheckBox chkAllowScreenReaders 
            Appearance      =   0  '2D
            Caption         =   "Allow Screen Readers"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   155
            Top             =   525
            Width           =   2865
         End
         Begin VB.CheckBox chkAllowFillIn 
            Appearance      =   0  '2D
            Caption         =   "Allow filling in form fields"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3100
            TabIndex        =   154
            Top             =   300
            Width           =   2760
         End
         Begin VB.CheckBox chkAllowDegradedPrinting 
            Appearance      =   0  '2D
            Caption         =   "Allow printing in low resolution"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   153
            Top             =   300
            Width           =   2865
         End
      End
      Begin PDFCreator.dmFrame dmFraPDFPermissions 
         Height          =   855
         Left            =   120
         TabIndex        =   147
         Top             =   3600
         Width           =   5955
         _extentx        =   10504
         _extenty        =   1508
         caption         =   "Disallow user to"
         barcolorfrom    =   16744576
         barcolorto      =   4194304
         font            =   "frmOptions.frx":07BE
         Begin VB.CheckBox chkAllowModifyAnnotations 
            Appearance      =   0  '2D
            Caption         =   "modify comments"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3100
            TabIndex        =   151
            Top             =   525
            Width           =   2760
         End
         Begin VB.CheckBox chkAllowModifyContents 
            Appearance      =   0  '2D
            Caption         =   "modify the document"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3100
            TabIndex        =   150
            Top             =   300
            Width           =   2760
         End
         Begin VB.CheckBox chkAllowCopy 
            Appearance      =   0  '2D
            Caption         =   "copy text and images"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   149
            Top             =   525
            Width           =   2865
         End
         Begin VB.CheckBox chkAllowPrinting 
            Appearance      =   0  '2D
            Caption         =   "print the document"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   148
            Top             =   300
            Width           =   2865
         End
      End
      Begin PDFCreator.dmFrame dmFraSecurityPass 
         Height          =   855
         Left            =   120
         TabIndex        =   144
         Top             =   2640
         Width           =   5955
         _extentx        =   10504
         _extenty        =   1508
         caption         =   "Passwords"
         barcolorfrom    =   16744576
         barcolorto      =   4194304
         font            =   "frmOptions.frx":07EA
         Begin VB.CheckBox chkOwnerPass 
            Appearance      =   0  '2D
            Caption         =   "Password required to change Permissions and Passwords"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   146
            Top             =   525
            Width           =   5700
         End
         Begin VB.CheckBox chkUserPass 
            Appearance      =   0  '2D
            Caption         =   "Password required to open document"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   145
            Top             =   300
            Width           =   5700
         End
      End
      Begin PDFCreator.dmFrame dmFraPDFEncLevel 
         Height          =   855
         Left            =   120
         TabIndex        =   141
         Top             =   1680
         Width           =   5955
         _extentx        =   10504
         _extenty        =   1508
         caption         =   "Encryption level"
         barcolorfrom    =   16744576
         barcolorto      =   4194304
         font            =   "frmOptions.frx":0816
         Begin VB.OptionButton optEncHigh 
            Appearance      =   0  '2D
            Caption         =   "High (128 Bit - Adobe Acrobat 5.0 and above)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   143
            Top             =   525
            Width           =   5775
         End
         Begin VB.OptionButton optEncLow 
            Appearance      =   0  '2D
            Caption         =   "Low (40 Bit - Adobe Acrobat 3.0 and above)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   142
            Top             =   300
            Width           =   5775
         End
      End
      Begin PDFCreator.dmFrame dmFraPDFEncryptor 
         Height          =   855
         Left            =   120
         TabIndex        =   139
         Top             =   720
         Width           =   5955
         _extentx        =   10504
         _extenty        =   1508
         caption         =   "Encryptor"
         barcolorfrom    =   16744576
         barcolorto      =   4194304
         font            =   "frmOptions.frx":0842
         Begin VB.ComboBox cmbPDFEncryptor 
            Appearance      =   0  '2D
            Height          =   315
            ItemData        =   "frmOptions.frx":086E
            Left            =   120
            List            =   "frmOptions.frx":0870
            Style           =   2  'Dropdown-Liste
            TabIndex        =   140
            Top             =   360
            Width           =   5715
         End
      End
      Begin VB.CheckBox chkUseSecurity 
         Appearance      =   0  '2D
         Caption         =   "Use Security"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   138
         Top             =   360
         Width           =   5535
      End
   End
   Begin PDFCreator.dmFrame dmfraPDFCompress 
      Height          =   4335
      Left            =   2760
      TabIndex        =   102
      Top             =   1920
      Visible         =   0   'False
      Width           =   6195
      _extentx        =   10927
      _extenty        =   7646
      caption         =   "Compression"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":0872
      Begin PDFCreator.dmFrame dmFraPDFMono 
         Height          =   1095
         Left            =   120
         TabIndex        =   118
         Top             =   3120
         Width           =   5955
         _extentx        =   10504
         _extenty        =   1931
         caption         =   "Monochrome images"
         barcolorfrom    =   16744576
         barcolorto      =   4194304
         font            =   "frmOptions.frx":089E
         Begin VB.CheckBox chkPDFMonoComp 
            Appearance      =   0  '2D
            Caption         =   "Compress"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   123
            Top             =   360
            Width           =   2325
         End
         Begin VB.ComboBox cmbPDFMonoComp 
            Appearance      =   0  '2D
            Height          =   315
            ItemData        =   "frmOptions.frx":08CA
            Left            =   120
            List            =   "frmOptions.frx":08CC
            Style           =   2  'Dropdown-Liste
            TabIndex        =   122
            Top             =   660
            Width           =   2370
         End
         Begin VB.CheckBox chkPDFMonoResample 
            Appearance      =   0  '2D
            Caption         =   "Resample"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2520
            TabIndex        =   121
            Top             =   360
            Width           =   2295
         End
         Begin VB.ComboBox cmbPDFMonoResample 
            Appearance      =   0  '2D
            Height          =   315
            ItemData        =   "frmOptions.frx":08CE
            Left            =   2520
            List            =   "frmOptions.frx":08D0
            Style           =   2  'Dropdown-Liste
            TabIndex        =   120
            Tag             =   "Bicubic|Subsample|Average"
            Top             =   660
            Width           =   2370
         End
         Begin VB.TextBox txtPDFMonoRes 
            Appearance      =   0  '2D
            Height          =   285
            Left            =   4935
            TabIndex        =   119
            Top             =   660
            Width           =   735
         End
         Begin VB.Label lblPDFMonoRes 
            AutoSize        =   -1  'True
            Caption         =   "Resolution"
            Height          =   195
            Left            =   4935
            TabIndex        =   124
            Top             =   360
            Width           =   750
         End
      End
      Begin PDFCreator.dmFrame dmFraPDFGrey 
         Height          =   1095
         Left            =   120
         TabIndex        =   111
         Top             =   1920
         Width           =   5955
         _extentx        =   10504
         _extenty        =   1931
         caption         =   "Greyscale images"
         barcolorfrom    =   16744576
         barcolorto      =   4194304
         font            =   "frmOptions.frx":08D2
         Begin VB.TextBox txtPDFGreyRes 
            Appearance      =   0  '2D
            Height          =   285
            Left            =   4935
            TabIndex        =   116
            Top             =   660
            Width           =   735
         End
         Begin VB.ComboBox cmbPDFGreyResample 
            Appearance      =   0  '2D
            Height          =   315
            ItemData        =   "frmOptions.frx":08FE
            Left            =   2520
            List            =   "frmOptions.frx":0900
            Style           =   2  'Dropdown-Liste
            TabIndex        =   115
            Tag             =   "Bicubic|Subsample|Average"
            Top             =   660
            Width           =   2370
         End
         Begin VB.CheckBox chkPDFGreyResample 
            Appearance      =   0  '2D
            Caption         =   "Resample"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2520
            TabIndex        =   114
            Top             =   360
            Width           =   2295
         End
         Begin VB.ComboBox cmbPDFGreyComp 
            Appearance      =   0  '2D
            Height          =   315
            ItemData        =   "frmOptions.frx":0902
            Left            =   120
            List            =   "frmOptions.frx":0904
            Style           =   2  'Dropdown-Liste
            TabIndex        =   113
            Top             =   660
            Width           =   2370
         End
         Begin VB.CheckBox chkPDFGreyComp 
            Appearance      =   0  '2D
            Caption         =   "Compress"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   360
            Width           =   2325
         End
         Begin VB.Label lblPDFGreyRes 
            AutoSize        =   -1  'True
            Caption         =   "Resolution"
            Height          =   195
            Left            =   4935
            TabIndex        =   117
            Top             =   360
            Width           =   750
         End
      End
      Begin PDFCreator.dmFrame dmFraPDFColor 
         Height          =   1095
         Left            =   120
         TabIndex        =   104
         Top             =   720
         Width           =   5955
         _extentx        =   10504
         _extenty        =   1931
         caption         =   "Color images"
         barcolorfrom    =   16744576
         barcolorto      =   4194304
         font            =   "frmOptions.frx":0906
         Begin VB.TextBox txtPDFColorRes 
            Appearance      =   0  '2D
            Height          =   285
            Left            =   4935
            TabIndex        =   109
            Top             =   660
            Width           =   735
         End
         Begin VB.ComboBox cmbPDFColorResample 
            Appearance      =   0  '2D
            Height          =   315
            ItemData        =   "frmOptions.frx":0932
            Left            =   2520
            List            =   "frmOptions.frx":0934
            Style           =   2  'Dropdown-Liste
            TabIndex        =   108
            Tag             =   "Bicubic|Subsample|Average"
            Top             =   660
            Width           =   2370
         End
         Begin VB.CheckBox chkPDFColorResample 
            Appearance      =   0  '2D
            Caption         =   "Resample"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2520
            TabIndex        =   107
            Top             =   360
            Width           =   2295
         End
         Begin VB.ComboBox cmbPDFColorComp 
            Appearance      =   0  '2D
            Height          =   315
            ItemData        =   "frmOptions.frx":0936
            Left            =   120
            List            =   "frmOptions.frx":0938
            Style           =   2  'Dropdown-Liste
            TabIndex        =   106
            Top             =   660
            Width           =   2370
         End
         Begin VB.CheckBox chkPDFColorComp 
            Appearance      =   0  '2D
            Caption         =   "Compress"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   105
            Top             =   360
            Width           =   2325
         End
         Begin VB.Label lblPDFColorRes 
            AutoSize        =   -1  'True
            Caption         =   "Resolution"
            Height          =   195
            Left            =   4935
            TabIndex        =   110
            Top             =   360
            Width           =   750
         End
      End
      Begin VB.CheckBox chkPDFTextComp 
         Appearance      =   0  '2D
         Caption         =   "Compress Text Objects"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   103
         Top             =   360
         Width           =   5910
      End
   End
   Begin PDFCreator.dmFrame dmFraPDFColors 
      Height          =   1215
      Left            =   2760
      TabIndex        =   130
      Top             =   2760
      Visible         =   0   'False
      Width           =   6195
      _extentx        =   10927
      _extenty        =   2143
      caption         =   "Color options"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":093A
      Begin VB.ComboBox cmbPDFColorModel 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "frmOptions.frx":0966
         Left            =   120
         List            =   "frmOptions.frx":0968
         Style           =   2  'Dropdown-Liste
         TabIndex        =   132
         Tag             =   "RGB|CMYK|GRAY"
         Top             =   360
         Width           =   2775
      End
      Begin VB.CheckBox chkPDFCMYKtoRGB 
         Appearance      =   0  '2D
         Caption         =   "Convert CMYK Images to RGB"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   131
         Top             =   840
         Width           =   5880
      End
   End
   Begin PDFCreator.dmFrame dmFraPDFColorOptions 
      Height          =   1455
      Left            =   2760
      TabIndex        =   133
      Top             =   4080
      Visible         =   0   'False
      Width           =   6195
      _extentx        =   10927
      _extenty        =   2566
      caption         =   "Options"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":096A
      Begin VB.CheckBox chkPDFPreserveHalftone 
         Appearance      =   0  '2D
         Caption         =   "Preserve Halftone Information"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   136
         Top             =   1050
         Width           =   5910
      End
      Begin VB.CheckBox chkPDFPreserveTransfer 
         Appearance      =   0  '2D
         Caption         =   "Preserve Transfer Functions"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   135
         Tag             =   "Remove|Preserve"
         Top             =   720
         Width           =   5910
      End
      Begin VB.CheckBox chkPDFPreserveOverprint 
         Appearance      =   0  '2D
         Caption         =   "Preserve Overprint Settings"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   134
         Top             =   360
         Width           =   5910
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   7320
      Width           =   1815
   End
   Begin PDFCreator.isExplorerBar ieb 
      Align           =   3  'Links ausrichten
      Height          =   7890
      Left            =   0
      TabIndex        =   160
      Top             =   0
      Width           =   2535
      _extentx        =   4471
      _extenty        =   13917
      fontname        =   "MS Sans Serif"
      fontcharset     =   0
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   210
         Top             =   7245
      End
      Begin MSComctlLib.ImageList imlIeb 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   17
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":0996
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":0F30
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":14CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":1A64
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":1FFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":2398
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":2932
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":2ECC
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":3466
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":3A00
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":3F9A
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":4534
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":4ACE
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":5068
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":5602
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":5B9C
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":6476
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   495
      Left            =   4955
      TabIndex        =   2
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   7245
      TabIndex        =   0
      Top             =   7320
      Width           =   1815
   End
   Begin PDFCreator.dmFrame dmfraFilenameSubstitutions 
      Height          =   2535
      Left            =   2640
      TabIndex        =   59
      Top             =   4200
      Visible         =   0   'False
      Width           =   6375
      _extentx        =   11245
      _extenty        =   4471
      caption         =   "Filename substitutions"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":6D50
      Begin VB.CommandButton cmdFilenameSubst 
         Caption         =   "&Add"
         Height          =   375
         Index           =   0
         Left            =   4440
         TabIndex        =   163
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdFilenameSubstMove 
         Height          =   435
         Index           =   0
         Left            =   120
         Picture         =   "frmOptions.frx":6D7C
         Style           =   1  'Grafisch
         TabIndex        =   161
         Top             =   915
         Width           =   375
      End
      Begin VB.TextBox txtFilenameSubst 
         Appearance      =   0  '2D
         Height          =   285
         Index           =   1
         Left            =   2640
         TabIndex        =   62
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox chkFilenameSubst 
         Appearance      =   0  '2D
         Caption         =   "Substitutions only in <Title>"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   61
         Top             =   2160
         Value           =   1  'Aktiviert
         Width           =   3255
      End
      Begin VB.TextBox txtFilenameSubst 
         Appearance      =   0  '2D
         Height          =   285
         Index           =   0
         Left            =   600
         TabIndex        =   60
         Top             =   360
         Width           =   1695
      End
      Begin MSComctlLib.ListView lsvFilenameSubst 
         Height          =   1335
         Left            =   600
         TabIndex        =   63
         Top             =   720
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   2355
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton cmdFilenameSubstMove 
         Height          =   435
         Index           =   1
         Left            =   120
         Picture         =   "frmOptions.frx":7106
         Style           =   1  'Grafisch
         TabIndex        =   162
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton cmdFilenameSubst 
         Caption         =   "C&hange"
         Height          =   375
         Index           =   1
         Left            =   4440
         TabIndex        =   164
         Top             =   1155
         Width           =   1455
      End
      Begin VB.CommandButton cmdFilenameSubst 
         Caption         =   "&Delete"
         Height          =   375
         Index           =   2
         Left            =   4440
         TabIndex        =   165
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblEqual 
         Caption         =   "="
         Height          =   255
         Left            =   2400
         TabIndex        =   64
         Top             =   360
         Width           =   135
      End
   End
   Begin PDFCreator.dmFrame dmFraDescription 
      Height          =   1065
      Left            =   2640
      TabIndex        =   157
      Top             =   105
      Width           =   6420
      _extentx        =   11324
      _extenty        =   1879
      caption         =   ""
      barcolorfrom    =   8421631
      barcolorto      =   192
      font            =   "frmOptions.frx":7490
      Begin VB.PictureBox picOptions 
         Appearance      =   0  '2D
         BorderStyle     =   0  'Kein
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   105
         ScaleHeight     =   615
         ScaleWidth      =   615
         TabIndex        =   158
         Top             =   420
         Width           =   615
      End
      Begin VB.Label lblOptions 
         Height          =   615
         Left            =   735
         TabIndex        =   159
         Top             =   420
         Width           =   5655
      End
   End
   Begin MSComctlLib.TabStrip tbstrPDFOptions 
      Height          =   4935
      Left            =   2640
      TabIndex        =   9
      Top             =   1320
      Visible         =   0   'False
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   8705
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin PDFCreator.dmFrame dmFraProgGeneral 
      Height          =   4215
      Left            =   2640
      TabIndex        =   10
      Top             =   1050
      Visible         =   0   'False
      Width           =   6420
      _extentx        =   11324
      _extenty        =   7435
      caption         =   "General"
      caption3d       =   2
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":74BC
      textshaddowcolor=   12582912
      Begin VB.CommandButton cmdTestpage 
         Caption         =   "&Print testpage"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   3600
         Width           =   2580
      End
      Begin VB.CommandButton cmdAsso 
         Caption         =   "&Associate PDFCreator with Postscript files"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   420
         Width           =   2580
      End
      Begin PDFCreator.Line3D Line3D1 
         Height          =   30
         Index           =   0
         Left            =   120
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1080
         Width           =   6135
         _extentx        =   10821
         _extenty        =   53
         linetype        =   1
         3dhighlight     =   -2147483628
         3dshadow        =   -2147483632
         drawstyle       =   0
      End
      Begin VB.CheckBox chkNoConfirmMessageSwitchingDefaultprinter 
         Appearance      =   0  '2D
         Caption         =   "No confirm message switching PDFCreator temporarly as default printer."
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   2640
         Width           =   5775
      End
      Begin MSComctlLib.Slider sldProcessPriority 
         Height          =   495
         Left            =   120
         TabIndex        =   4
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
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2280
         Width           =   6135
         _extentx        =   10821
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
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   3360
         Width           =   6135
         _extentx        =   10821
         _extenty        =   53
         linetype        =   1
         3dhighlight     =   -2147483628
         3dshadow        =   -2147483632
         drawstyle       =   0
      End
      Begin VB.Label lblProcessPriority 
         AutoSize        =   -1  'True
         Caption         =   "Processpriority: Normal"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1605
      End
   End
   Begin PDFCreator.dmFrame dmFraShellIntegration 
      Height          =   1065
      Left            =   2640
      TabIndex        =   12
      Top             =   5565
      Width           =   6420
      _extentx        =   11324
      _extenty        =   1879
      caption         =   "Shell integration"
      caption3d       =   2
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":74E8
      textshaddowcolor=   12582912
      Begin VB.CommandButton cmdShellintegration 
         Caption         =   "Integrate PDFCreator into shell"
         Height          =   495
         Index           =   1
         Left            =   3240
         TabIndex        =   8
         Top             =   420
         Width           =   3015
      End
      Begin VB.CommandButton cmdShellintegration 
         Caption         =   "Integrate PDFCreator into shell"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   420
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SetPDFColorComprSettings()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkPDFColorComp.Value = 1 Then
50020    cmbPDFColorComp.Enabled = True
50030    If cmbPDFColorComp.ListIndex = 0 Then
50040      chkPDFColorResample.Enabled = False
50050      cmbPDFColorResample.Enabled = False
50060      lblPDFColorRes.Enabled = False
50070      txtPDFColorRes.Enabled = False
50080     Else
50090      chkPDFColorResample.Enabled = True
50100      If chkPDFColorResample.Value = 1 Then
50110        cmbPDFColorResample.Enabled = True
50120        lblPDFColorRes.Enabled = True
50130        txtPDFColorRes.Enabled = True
50140       Else
50150        cmbPDFColorResample.Enabled = False
50160        lblPDFColorRes.Enabled = False
50170        txtPDFColorRes.Enabled = False
50180      End If
50190    End If
50200   Else
50210    cmbPDFColorComp.Enabled = False
50220    chkPDFColorResample.Enabled = False
50230    cmbPDFColorResample.Enabled = False
50240    lblPDFColorRes.Enabled = False
50250    txtPDFColorRes.Enabled = False
50260  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "SetPDFColorComprSettings")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetPDFGreyComprSettings()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkPDFGreyComp.Value = 1 Then
50020    cmbPDFGreyComp.Enabled = True
50030    If cmbPDFGreyComp.ListIndex = 0 Then
50040      chkPDFGreyResample.Enabled = False
50050      cmbPDFGreyResample.Enabled = False
50060      lblPDFGreyRes.Enabled = False
50070      txtPDFGreyRes.Enabled = False
50080     Else
50090      chkPDFGreyResample.Enabled = True
50100      If chkPDFGreyResample.Value = 1 Then
50110        cmbPDFGreyResample.Enabled = True
50120        lblPDFGreyRes.Enabled = True
50130        txtPDFGreyRes.Enabled = True
50140       Else
50150        cmbPDFGreyResample.Enabled = False
50160        lblPDFGreyRes.Enabled = False
50170        txtPDFGreyRes.Enabled = False
50180      End If
50190    End If
50200   Else
50210    cmbPDFGreyComp.Enabled = False
50220    chkPDFGreyResample.Enabled = False
50230    cmbPDFGreyResample.Enabled = False
50240    lblPDFGreyRes.Enabled = False
50250    txtPDFGreyRes.Enabled = False
50260  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "SetPDFGreyComprSettings")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetPDFMonoComprSettings()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkPDFMonoComp.Value = 1 Then
50020    cmbPDFMonoComp.Enabled = True
50030    chkPDFMonoResample.Enabled = True
50040    If chkPDFMonoResample.Value = 1 Then
50050      cmbPDFMonoResample.Enabled = True
50060      lblPDFMonoRes.Enabled = True
50070      txtPDFMonoRes.Enabled = True
50080     Else
50090      cmbPDFMonoResample.Enabled = False
50100      lblPDFMonoRes.Enabled = False
50110      txtPDFMonoRes.Enabled = False
50120    End If
50130   Else
50140    cmbPDFMonoComp.Enabled = False
50150    chkPDFMonoResample.Enabled = False
50160    cmbPDFMonoResample.Enabled = False
50170    lblPDFMonoRes.Enabled = False
50180    txtPDFMonoRes.Enabled = False
50190  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "SetPDFMonoComprSettings")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub


Private Sub chkOwnerPass_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkUserPass.Value = 0 Then
50020   If chkOwnerPass.Value = 0 Then
50030    chkOwnerPass.Value = 1
50040   End If
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "chkOwnerPass_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkPDFColorComp_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  SetPDFColorComprSettings
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "chkPDFColorComp_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkPDFColorResample_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  SetPDFColorComprSettings
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "chkPDFColorResample_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkPDFGreyComp_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  SetPDFGreyComprSettings
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "chkPDFGreyComp_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkPDFGreyResample_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  SetPDFGreyComprSettings
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "chkPDFGreyResample_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkPDFMonoComp_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  SetPDFMonoComprSettings
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "chkPDFMonoComp_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkPDFMonoResample_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  SetPDFMonoComprSettings
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "chkPDFMonoResample_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub


Private Sub chkUseAutosave_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkUseAutosave.Value = 1 Then
50020    ViewAutosave True
50030   Else
50040    ViewAutosave False
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "chkUseAutosave_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkUseAutosaveDirectory_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkUseAutosaveDirectory.Value = 1 Then
50020    ViewAutosaveDirectory True
50030   Else
50040    ViewAutosaveDirectory False
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "chkUseAutosaveDirectory_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkUserPass_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkOwnerPass.Value = 0 Then
50020   If chkUserPass.Value = 0 Then
50030    chkUserPass.Value = 1
50040    chkOwnerPass.Value = 1
50050   End If
50060   SavePasswordsForThisSession = False
50070  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "chkUserPass_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkUseSecurity_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  UpdateSecurityFields
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "chkUseSecurity_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkUseStandardAuthor_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkUseStandardAuthor.Value = 1 Then
50020    txtStandardAuthor.Enabled = True
50030    txtStandardAuthor.BackColor = &H80000005
50040    cmbAuthorTokens.Enabled = True
50050    lblAuthorTokens.Enabled = True
50060   Else
50070    txtStandardAuthor.Enabled = False
50080    txtStandardAuthor.BackColor = &H8000000F
50090    cmbAuthorTokens.Enabled = False
50100    lblAuthorTokens.Enabled = False
50110  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "chkUseStandardAuthor_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbAuthorTokens_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txtStandardAuthor.Text = txtStandardAuthor.Text & cmbAuthorTokens.Text
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmbAuthorTokens_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbAutosaveFormat_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Ext As String
50020  txtAutosaveFilename.ToolTipText = txtAutosaveFilename.Text
50030  txtAutoSaveFilenamePreview.Text = GetSubstFilename("B:\dummy.dum", txtAutosaveFilename.Text, , True) & _
  GetAutosaveFormatExtension
50050  If IsValidPath("C:\" & txtAutoSaveFilenamePreview.Text) = False Then
50060    txtAutoSaveFilenamePreview.ForeColor = vbRed
50070   Else
50080    txtAutoSaveFilenamePreview.ForeColor = &H80000008
50090  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmbAutosaveFormat_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbCharset_Change()
 On Error GoTo ErrorHandler
 txtTest.Font.Charset = cmbCharset.Text
 Exit Sub
ErrorHandler:
 If Err.Number = 380 Then
  cmbCharset.Text = 0
 End If
 Err.Clear
End Sub

Private Sub cmbCharset_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With cmbCharset
50020   .Text = .ItemData(.ListIndex)
50030  End With
50040  txtTest.Font.Charset = cmbCharset.Text
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmbCharset_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbCharset_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim allow As String, tStr As String
50020  allow = "0123456789" & Chr$(8) & Chr$(13)
50030  tStr = Chr$(KeyAscii)
50040  If InStr(1, allow, tStr) = 0 Then
50050    KeyAscii = 0
50060  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmbCharset_KeyPress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbCharset_Validate(Cancel As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, tStr As String
50020  tStr = ""
50030  For i = 1 To Len(cmbCharset.Text)
50040   If InStr("0123456789", Mid(cmbCharset.Text, i, 1)) > 0 Then
50050     tStr = tStr & Mid(cmbCharset.Text, i, 1)
50060    Else
50070     Exit For
50080   End If
50090  Next i
50100  If Len(Trim$(tStr)) = 0 Then
50110    cmbCharset.Text = 0
50120   Else
50130    cmbCharset.Text = tStr
50140  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmbCharset_Validate")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbAutoSaveFilenameTokens_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txtAutosaveFilename.Text = txtAutosaveFilename.Text & cmbAutoSaveFilenameTokens.Text
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmbAutoSaveFilenameTokens_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbGhostscript_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry, gsv As String, tsf() As String, Path As String, tStr As String
50020
50030  gsv = cmbGhostscript.List(cmbGhostscript.ListIndex)
50040  Set reg = New clsRegistry: reg.hkey = HKEY_LOCAL_MACHINE
50050
50060  If InStr(gsv, ":") Then
50070    reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50080    txtGSbin.Text = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryBinaries"))
50090    txtGSfonts.Text = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
50100    txtGSlib.Text = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryLibraries"))
50110    txtGSresource.Text = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryResource"))
50120    Set reg = Nothing
50130    Exit Sub
50140   Else
50150    If InStr(UCase$(gsv), "AFPL") Then
50160     If InStr(gsv, " ") > 0 Then
50170      tsf = Split(gsv, " ")
50180      reg.KeyRoot = "SOFTWARE\AFPL Ghostscript\" & tsf(UBound(tsf))
50190      tStr = reg.GetRegistryValue("GS_DLL")
50200      SplitPath tStr, , Path
50210      txtGSbin.Text = CompletePath(Path)
50220      If InStrRev(Path, "\") > 0 Then
50230       txtGSlib.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50240       txtGSfonts.Text = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50250       If tsf(UBound(tsf)) <> "8.00" Then
50260        txtGSresource.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "Resource"
50270       End If
50280      End If
50290     End If
50300    End If
50310    If InStr(UCase$(gsv), "GNU") Then
50320     If InStr(gsv, " ") > 0 Then
50330      tsf = Split(gsv, " ")
50340      reg.KeyRoot = "SOFTWARE\GNU Ghostscript\" & tsf(UBound(tsf))
50350      tStr = reg.GetRegistryValue("GS_DLL")
50360      SplitPath tStr, , Path
50370      txtGSbin.Text = CompletePath(Path)
50380      If InStrRev(Path, "\") > 0 Then
50390       txtGSlib.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50400       txtGSfonts.Text = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50410       txtGSresource.Text = ""
50420      End If
50430     End If
50440    End If
50450    If InStr(UCase$(gsv), "GPL") Then
50460     If InStr(gsv, " ") > 0 Then
50470      tsf = Split(gsv, " ")
50480      reg.KeyRoot = "SOFTWARE\GPL Ghostscript\" & tsf(UBound(tsf))
50490      tStr = reg.GetRegistryValue("GS_DLL")
50500      SplitPath tStr, , Path
50510      txtGSbin.Text = CompletePath(Path)
50520      If InStrRev(Path, "\") > 0 Then
50530       txtGSlib.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50540       txtGSfonts.Text = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50550       txtGSresource.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "Resource"
50560      End If
50570     End If
50580    End If
50590  End If
50600  Set reg = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmbGhostscript_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbPDFColorComp_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  SetPDFColorComprSettings
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmbPDFColorComp_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbPDFGreyComp_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  SetPDFGreyComprSettings
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmbPDFGreyComp_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbPDFMonoComp_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  SetPDFMonoComprSettings
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmbPDFMonoComp_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbSaveFilenameTokens_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txtSaveFilename.Text = txtSaveFilename.Text & cmbSaveFilenameTokens.Text
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmbSaveFilenameTokens_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbFonts_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txtTest.Font.Name = cmbFonts.List(cmbFonts.ListIndex)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmbFonts_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbPDFCompat_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  UpdateSecurityFields
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmbPDFCompat_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdCancel_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Unload Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdCancel_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdCancelTest_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With Options
50020   SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50030   cmbCharset.Text = .ProgramFontCharset
50040   SetFont frmMain, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50050   ieb.Refresh
50060  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdCancelTest_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdFilenameSubst_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case Index
        Case 0: ' Add
50030    AddFilenameSubstitutions
50040   Case 1: ' Change
50050    ChangeFilenameSubstitutions
50060   Case 2: ' Delete
50070    DeleteFilenameSubstitutions
50080  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdFilenameSubst_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdFilenameSubstMove_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case Index
        Case 0: ' Up
50030    MoveUpFilenameSubstitutions
50040   Case 1: ' Down
50050    MoveDownFilenameSubstitutions
50060  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdFilenameSubstMove_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdGetAutosaveDirectory_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 Dim strFolder As String
50020 strFolder = BrowseForFolder(Me.hwnd, LanguageStrings.OptionsAutosaveDirectoryPrompt)
50030 If Len(strFolder) = 0 Then Exit Sub
50040 txtAutosaveDirectory.Text = CompletePath(strFolder)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdGetAutosaveDirectory_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdGetgsbinDirectory_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim strFolder As String, aw As Long
50020  strFolder = BrowseForFolder(Me.hwnd, LanguageStrings.OptionsGhostscriptBinariesDirectoryPrompt)
50030  If Len(strFolder) = 0 Then
50040   Exit Sub
50050  End If
50060  strFolder = CompletePath(strFolder)
50070  If FileExists(strFolder & GsDll) = False Then
50080   MsgBox LanguageStrings.MessagesMsg15
50090   Exit Sub
50100  End If
50110  If UCase$(CompletePath(Options.DirectoryGhostscriptBinaries)) <> UCase$(CompletePath(strFolder)) Then
50120   aw = MsgBox("The program must be restarted!", vbOKCancel)
50130   If aw = vbCancel Then
50140    Exit Sub
50150   End If
50160   txtGSbin.Text = strFolder
50170   GetOptions Me, Options
50180   SaveOptions Options
50190   Restart = True
50200   Unload Me
50210  End If
50220  With txtGSbin
50230   .Text = strFolder
50240   .ToolTipText = .Text
50250  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdGetgsbinDirectory_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdGetgsfontsDirectory_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim strFolder As String
50020  strFolder = BrowseForFolder(Me.hwnd, LanguageStrings.OptionsGhostscriptFontsDirectoryPrompt)
50030  If Len(strFolder) = 0 Then Exit Sub
50040  strFolder = CompletePath(strFolder)
50050  If LenB(Dir(strFolder & "*.afm", vbNormal)) = 0 And LenB(Dir(strFolder & "*.pfb", vbNormal)) = 0 Then
50060   MsgBox LanguageStrings.MessagesMsg16
50070   Exit Sub
50080  End If
50090  txtGSfonts.Text = strFolder
50100  With txtGSfonts
50110   .ToolTipText = .Text
50120  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdGetgsfontsDirectory_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdGetgslibDirectory_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim strFolder As String
50020  strFolder = BrowseForFolder(Me.hwnd, LanguageStrings.OptionsGhostscriptLibrariesDirectoryPrompt)
50030  If Len(strFolder) = 0 Then Exit Sub
50040  strFolder = CompletePath(strFolder)
50050  If LenB(Dir(strFolder & "*.*", vbNormal)) = 0 Then
50060   MsgBox LanguageStrings.MessagesMsg17
50070   Exit Sub
50080  End If
50090  With txtGSlib
50100   .Text = strFolder
50110   .ToolTipText = .Text
50120  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdGetgslibDirectory_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdGetgsresourceDirectory_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim strFolder As String
50020  strFolder = BrowseForFolder(Me.hwnd, LanguageStrings.OptionsGhostscriptResourceDirectoryPrompt)
50030  If Len(strFolder) = 0 Then Exit Sub
50040  strFolder = CompletePath(strFolder)
50050  With txtGSresource
50060   .Text = strFolder
50070   .ToolTipText = .Text
50080  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdGetgsresourceDirectory_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdGetTemppath_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim strFolder As String
50020  strFolder = BrowseForFolder(Me.hwnd, LanguageStrings.OptionsPrintertempDirectoryPrompt)
50030  If Len(strFolder) = 0 Then Exit Sub
50040  strFolder = CompletePath(strFolder)
50050  With txtTemppath
50060   .Text = strFolder
50070   .ToolTipText = .Text
50080  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdGetTemppath_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdReset_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim res As Long, Options As tOptions
50020  res = MsgBox(LanguageStrings.MessagesMsg03, vbYesNo)
50030  If res = vbYes Then
50040   Options = StandardOptions
50050   ShowOptions Me, Options
50060   With Options
50070    SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50080    cmbCharset.Text = .ProgramFontCharset
50090    SetFont frmMain, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50100    ieb.Refresh
50110   End With
50120  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdReset_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdSave_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tRestart As Boolean
50020  tRestart = False
50030  If UCase$(Options.DirectoryGhostscriptBinaries) <> UCase$(txtGSbin.Text) Then
50040   tRestart = True
50050  End If
50060  CorrectCmbCharset
50070  GetOptions Me, Options
50080  CorrectOptions
50090  SaveOptions Options
50100  If IsWin9xMe = False Then
50111   Select Case Options.ProcessPriority
               Case 0: 'Idle
50130     SetProcessPriority Idle
50140    Case 1: 'Normal
50150     SetProcessPriority Normal
50160    Case 2: 'High
50170     SetProcessPriority High
50180    Case 3: 'Realtime
50190     SetProcessPriority RealTime
50200   End Select
50210  End If
50220  If tRestart = True Then
50230   Restart = True
50240  End If
50250  Unload Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdSave_Click")
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
Select Case ErrPtnr.OnError("frmOptions", "cmdShellintegration_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdTest_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tCharset As Long, tStr As String, tFontSize As Long, tFontname As String, _
  tFontCharset As Long
50030  If InStr(1, cmbCharset.Text, ",", vbTextCompare) > 0 Then
50040    tStr = Trim$(Mid$(cmbCharset.Text, 1, InStr(1, cmbCharset.Text, ",", vbTextCompare) - 1))
50050   Else
50060    tStr = Trim$(cmbCharset.Text)
50070  End If
50080  If Len(tStr) = 0 Then
50090   cmbCharset.Text = 0
50100   Exit Sub
50110  End If
50120  If IsNumeric(tStr) = False Then
50130   cmbCharset.Text = 0
50140   Exit Sub
50150  End If
50160  tCharset = tStr
50170  With cmdTest.Font
50180   tFontname = .Name
50190   tFontSize = .Size
50200   tFontCharset = .Charset
50210  End With
50220  SetFont Me, cmbFonts.List(cmbFonts.ListIndex), CLng(tStr), cmbProgramFontsize.Text
50230  cmbCharset.Text = tCharset
50240  SetFont frmMain, cmbFonts.List(cmbFonts.ListIndex), CLng(tStr), cmbProgramFontsize.Text
50250  ieb.Refresh
50260  With cmdTest.Font
50270   .Name = tFontname
50280   .Size = tFontSize
50290   .Charset = tFontCharset
50300  End With
50310  With cmdCancelTest
50320   .Font.Name = tFontname
50330   .Font.Size = tFontSize
50340   .Font.Charset = tFontCharset
50350   .Enabled = True
50360  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdTest_Click")
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
Select Case ErrPtnr.OnError("frmOptions", "cmdTestpage_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdUsertempPath_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Temppath As String
50020  Temppath = "<Temp>PDFCreator\"
50030  If DirExists(ResolveEnvironment(GetSubstFilename2(Temppath))) = False Then
50040   MakePath ResolveEnvironment(GetSubstFilename2(Temppath))
50050  End If
50060  With txtTemppath
50070   .Text = Temppath
50080   .ToolTipText = ResolveEnvironment(GetSubstFilename2(Temppath))
50090  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdUsertempPath_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If KeyCode = vbKeyF1 Then
50020   KeyCode = 0
50030     Call HTMLHelp_ShowTopic("html\generalsettings.htm")
50040 '  MsgBox ieb.GetSelectedGroup & vbCrLf & ieb.GetSelectedItem
50051    Select Case ieb.GetSelectedGroup
          Case 1
50071      Select Case ieb.GetSelectedItem
            Case 1
50090        Call HTMLHelp_ShowTopic("html\generalsettings.htm")
50100       Case 2
50110        Call HTMLHelp_ShowTopic("html\ghostscript.htm")
50120       Case 3
50130        Call HTMLHelp_ShowTopic("html\docproperties.htm")
50140       Case 4
50150        Call HTMLHelp_ShowTopic("html\savesettings.htm")
50160       Case 5
50170        Call HTMLHelp_ShowTopic("html\autosave.htm")
50180       Case 6
50190        Call HTMLHelp_ShowTopic("html\directories.htm")
50200       Case 7
50210        Call HTMLHelp_ShowTopic("html\fontsetting.htm")
50220       Case Else
50230        Call HTMLHelp_ShowTopic("html\generalsettings.htm")
50240      End Select
50250     Case 2
50261      Select Case ieb.GetSelectedItem
            Case 1
50281        Select Case tbstrPDFOptions.SelectedItem.Index
              Case 1
50300          Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
50310         Case 2
50320          Call HTMLHelp_ShowTopic("html\pdfcompression.htm")
50330         Case 3
50340          Call HTMLHelp_ShowTopic("html\pdffonts.htm")
50350         Case 4
50360          Call HTMLHelp_ShowTopic("html\pdfcolors.htm")
50370         Case 5
50380          Call HTMLHelp_ShowTopic("html\pdfsecurity.htm")
50390         Case Else
50400          Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
50410        End Select
50420       Case 2
50430        Call HTMLHelp_ShowTopic("html\pngsettings.htm")
50440       Case 3
50450        Call HTMLHelp_ShowTopic("html\jpegsettings.htm")
50460       Case 4
50470        Call HTMLHelp_ShowTopic("html\bmpsettings.htm")
50480       Case 5
50490        Call HTMLHelp_ShowTopic("html\pcxsettings.htm")
50500       Case 6
50510        Call HTMLHelp_ShowTopic("html\tiffsettings.htm")
50520       Case 7
50530        Call HTMLHelp_ShowTopic("html\pssettings.htm")
50540       Case 8
50550        Call HTMLHelp_ShowTopic("html\epssettings.htm")
50560       Case Else
50570        Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
50580      End Select
50590    End Select
50600  End If
50610
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "Form_KeyDown")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Load()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Const fraPDFTop = 1360, fraPDFLeft = 2960
50020  Dim pic As New StdPicture, i As Long, tStr As String, gsvers As Collection, _
  fc As Long, reg As clsRegistry, tsf() As String, tStr2 As String, _
  ctl As Control
50050  Me.Icon = LoadResPicture(2120, vbResIcon)
50060  KeyPreview = True
50070
50080  With Screen
50090   .MousePointer = vbHourglass
50100   Move (.Width - Width) / 2, (.Height - Height) / 2
50110  End With
50120
50130  For Each ctl In Controls
50140   If TypeOf ctl Is dmFrame Then
50150    ctl.Font.Size = 10
50160    ctl.TextShaddowColor = &HC00000
50170    ctl.Caption3D = [Raised Caption]
50180    If ComputerScreenResolution <= 8 Or Options.OptionsDesign = 2 Then
50190     ctl.UseGradient = False: ctl.Caption3D = [Flat Caption]
50200     If UCase$(ctl.Name) = "DMFRADESCRIPTION" Then
50210       ctl.BarColorFrom = vbRed
50220      Else
50230       ctl.BarColorFrom = vbBlue
50240     End If
50250    End If
50260   End If
50270  Next ctl
50280
50290  With dmFraDescription
50300   .Caption = LanguageStrings.OptionsTreeProgram
50310   .Visible = True
50320  End With
50330  dmFraShellIntegration.Visible = True
50340  With dmFraProgGeneral
50350   .Visible = True
50360   .Top = dmFraDescription.Top + dmFraDescription.Height + 50
50370   .Left = dmFraDescription.Left
50380   dmFraShellIntegration.Top = .Top + .Height + 50
50390   dmFraShellIntegration.Left = .Left
50400   dmFraShellIntegration.Width = .Width
50410   dmFraProgGhostscript.Top = .Top
50420   dmFraProgGhostscript.Left = .Left
50430   dmFraProgGhostscript.Width = .Width
50440   dmFraProgAutosave.Top = .Top
50450   dmFraProgAutosave.Left = .Left
50460   dmFraProgAutosave.Width = .Width
50470   dmFraProgDirectories.Top = .Top
50480   dmFraProgDirectories.Left = .Left
50490   dmFraProgDirectories.Width = .Width
50500   dmFraProgDocument.Top = .Top
50510   dmFraProgDocument.Left = .Left
50520   dmFraProgDocument.Width = .Width
50530   dmfraProgSave.Top = .Top
50540   dmfraProgSave.Left = .Left
50550   dmfraProgSave.Width = .Width
50560   dmfraFilenameSubstitutions.Top = dmfraProgSave.Top + dmfraProgSave.Height + 50
50570   dmfraFilenameSubstitutions.Left = .Left
50580   dmfraFilenameSubstitutions.Width = .Width
50590   dmFraProgFont.Top = .Top
50600   dmFraProgFont.Left = .Left
50610   dmFraProgFont.Width = .Width
50620   dmFraBitmapGeneral.Top = .Top
50630   dmFraBitmapGeneral.Left = .Left
50640   dmFraBitmapGeneral.Width = .Width
50650   dmFraPSGeneral.Top = .Top
50660   dmFraPSGeneral.Left = .Left
50670   dmFraPSGeneral.Width = .Width
50680
50690   cmdCancel.Left = .Left
50700   cmdReset.Left = .Left + (.Width - cmdReset.Width) / 2
50710   cmdSave.Left = .Left + .Width - cmdSave.Width
50720  End With
50730
50740  With tbstrPDFOptions
50750   .Top = dmFraDescription.Top + dmFraDescription.Height + 50
50760   .Left = dmFraDescription.Left
50770   .Height = cmdCancel.Top - tbstrPDFOptions.Top - 50
50780   .Width = dmFraDescription.Width
50790  End With
50800
50810  With dmFraPDFGeneral
50820   .Top = tbstrPDFOptions.ClientTop + 100
50830   .Left = tbstrPDFOptions.Left + (tbstrPDFOptions.Width - .Width) / 2
50840   dmfraPDFCompress.Top = .Top
50850   dmfraPDFCompress.Left = .Left
50860   dmFraPDFFonts.Top = .Top
50870   dmFraPDFFonts.Left = .Left
50880   dmFraPDFColors.Top = .Top
50890   dmFraPDFColors.Left = .Left
50900   dmFraPDFColorOptions.Top = dmFraPDFColors.Top + dmFraPDFColors.Height + 50
50910   dmFraPDFColorOptions.Left = .Left
50920   dmFraPDFSecurity.Top = .Top
50930   dmFraPDFSecurity.Left = .Left
50940  End With
50950
50960  cmbEPSLanguageLevel.Top = cmbPSLanguageLevel.Top
50970  cmbEPSLanguageLevel.Left = cmbPSLanguageLevel.Left
50980
50990  ieb.DisableUpdates True
51000  ieb.ClearStructure
51010  ieb.SetImageList imlIeb
51020  With LanguageStrings
51030   ieb.AddGroup "Program", .OptionsTreeProgram, 0
51040   ieb.AddItem "Program", "General", .OptionsProgramGeneralSymbol, 1
51050   ieb.AddItem "Program", "Ghostscript", .OptionsProgramGhostscriptSymbol, 2
51060   ieb.AddItem "Program", "Document", .OptionsProgramDocumentSymbol, 3
51070   ieb.AddItem "Program", "Save", .OptionsProgramSaveSymbol, 4
51080   ieb.AddItem "Program", "AutoSave", .OptionsProgramAutosaveSymbol, 5
51090   ieb.AddItem "Program", "Directories", .OptionsProgramDirectoriesSymbol, 6
51100   ieb.AddItem "Program", "Fonts", .OptionsProgramFontSymbol, 7
51110   ieb.AddGroup "Formats", .OptionsTreeFormats, 0
51120   ieb.AddItem "Formats", "PDF", .OptionsPDFSymbol, 8
51130   ieb.AddItem "Formats", "PNG", .OptionsPNGSymbol, 9
51140   ieb.AddItem "Formats", "JPEG", .OptionsJPEGSymbol, 10
51150   ieb.AddItem "Formats", "BMP", .OptionsBMPSymbol, 11
51160   ieb.AddItem "Formats", "PCX", .OptionsPCXSymbol, 12
51170   ieb.AddItem "Formats", "TIFF", .OptionsTIFFSymbol, 13
51180   ieb.AddItem "Formats", "PS", .OptionsPSSymbol, 14
51190   ieb.AddItem "Formats", "EPS", .OptionsEPSSymbol, 15
51200   ieb.DisableUpdates False
51210
51220   Set picOptions = LoadResPicture(2101, vbResIcon)
51230   dmFraProgGeneral.Visible = True
51240
51250   dmFraProgGeneral.Caption = .OptionsProgramGeneralSymbol
51260   dmFraShellIntegration.Caption = .OptionsShellIntegration
51270   dmFraProgGhostscript.Caption = .OptionsProgramGhostscriptSymbol
51280   dmFraProgAutosave.Caption = .OptionsProgramAutosaveSymbol
51290   dmFraProgDirectories.Caption = .OptionsProgramDirectoriesSymbol
51300   dmFraProgDocument.Caption = .OptionsProgramDocumentSymbol
51310   dmFraProgFont.Caption = .OptionsProgramFontSymbol
51320   dmfraProgSave.Caption = .OptionsProgramSaveSymbol
51330
51340   cmdShellintegration(0).Caption = .OptionsShellIntegrationAdd
51350   cmdShellintegration(1).Caption = .OptionsShellIntegrationRemove
51360   If IsWin9xMe = False Then
51370    If IsAdmin = False Then
51380     cmdShellintegration(0).Enabled = False
51390     cmdShellintegration(1).Enabled = False
51400    End If
51410   End If
51420
51430   lblGhostscriptversion.Caption = .OptionsGhostscriptversion
51440   lblAdditionalGhostscriptParameters.Caption = .OptionsAdditionalGhostscriptParameters
51450   lblAdditionalGhostscriptSearchpath.Caption = .OptionsAdditionalGhostscriptSearchpath
51460   chkAddWindowsFontpath.Caption = .OptionsAddWindowsFontpath
51470
51480   lblSaveFilename.Caption = .OptionsSaveFilename
51490   lblSaveFilenameTokens.Caption = .OptionsSaveFilenameTokens
51500   dmfraFilenameSubstitutions.Caption = .OptionsSaveFilenameSubstitutions
51510   chkFilenameSubst.Caption = .OptionsSaveFilenameSubstitutionsTitle
51520   cmdFilenameSubst(0).Caption = .OptionsSaveFilenameAdd
51530   cmdFilenameSubst(1).Caption = .OptionsSaveFilenameChange
51540   cmdFilenameSubst(2).Caption = .OptionsSaveFilenameDelete
51550
51560   chkSpaces.Caption = .OptionsRemoveSpaces
51570   chkNoConfirmMessageSwitchingDefaultprinter.Caption = .OptionsProgramSwitchingDefaultprinter
51580   lblGSbin.Caption = .OptionsDirectoriesGSBin
51590   lblGSlib.Caption = .OptionsDirectoriesGSLibraries
51600   lblGSfonts.Caption = .OptionsDirectoriesGSFonts
51610   lblPrintTempPath.Caption = .OptionsDirectoriesTempPath
51620
51630   lblOptions = .OptionsProgramGeneralDescription
51640   lblAutosaveformat.Caption = .OptionsAutosaveFormat
51650   chkUseStandardAuthor.Caption = .OptionsUseStandardauthor
51660   chkUseCreationDateNow.Caption = .OptionsUseCreationDateNow
51670   chkUseAutosave.Caption = .OptionsUseAutosave
51680   cmdTestpage.Caption = .OptionsPrintTestpage
51690   lblAutosaveFilename.Caption = .OptionsAutosaveFilename
51700   lblAutosaveFilenameTokens.Caption = .OptionsAutosaveFilenameTokens
51710   chkUseAutosaveDirectory.Caption = .OptionsUseAutosaveDirectory
51720   lblAuthorTokens.Caption = .OptionsStandardAuthorToken
51730
51740   With cmbAutosaveFormat
51750    .AddItem "PDF"
51760    .AddItem "PNG"
51770    .AddItem "JPEG"
51780    .AddItem "BMP"
51790    .AddItem "PCX"
51800    .AddItem "TIFF"
51810    .AddItem "PS"
51820    .AddItem "EPS"
51830   End With
51840   With cmbSaveFilenameTokens
51850    .AddItem "<Author>"
51860    .AddItem "<Computername>"
51870    .AddItem "<DateTime>"
51880    .AddItem "<Title>"
51890    .AddItem "<Username>"
51900    .AddItem "<REDMON_DOCNAME>"
51910    .AddItem "<REDMON_DOCNAME_FILE>"
51920    .AddItem "<REDMON_DOCNAME_PATH>"
51930    .AddItem "<REDMON_JOB>"
51940    .AddItem "<REDMON_MACHINE>"
51950    .AddItem "<REDMON_PORT>"
51960    .AddItem "<REDMON_PRINTER>"
51970    .AddItem "<REDMON_SESSIONID>"
51980    .AddItem "<REDMON_USER>"
51990    .ListIndex = 0
52000   End With
52010   With cmbAuthorTokens
52020    .AddItem "<Computername>"
52030    .AddItem "<ClientComputer>"
52040    .AddItem "<DateTime>"
52050    .AddItem "<Title>"
52060    .AddItem "<Username>"
52070    .AddItem "<REDMON_DOCNAME>"
52080    .AddItem "<REDMON_DOCNAME_FILE>"
52090    .AddItem "<REDMON_DOCNAME_PATH>"
52100    .AddItem "<REDMON_JOB>"
52110    .AddItem "<REDMON_MACHINE>"
52120    .AddItem "<REDMON_PORT>"
52130    .AddItem "<REDMON_PRINTER>"
52140    .AddItem "<REDMON_SESSIONID>"
52150    .AddItem "<REDMON_USER>"
52160    .ListIndex = 0
52170   End With
52180   With cmbAutoSaveFilenameTokens
52190    .AddItem "<Author>"
52200    .AddItem "<Computername>"
52210    .AddItem "<ClientComputer>"
52220    .AddItem "<DateTime>"
52230    .AddItem "<Title>"
52240    .AddItem "<Username>"
52250    .AddItem "<REDMON_DOCNAME>"
52260    .AddItem "<REDMON_DOCNAME_FILE>"
52270    .AddItem "<REDMON_DOCNAME_PATH>"
52280    .AddItem "<REDMON_JOB>"
52290    .AddItem "<REDMON_MACHINE>"
52300    .AddItem "<REDMON_PORT>"
52310    .AddItem "<REDMON_PRINTER>"
52320    .AddItem "<REDMON_SESSIONID>"
52330    .AddItem "<REDMON_USER>"
52340    .ListIndex = 0
52350   End With
52360   Me.Caption = .DialogPrinterOptions
52370   cmdCancel.Caption = .OptionsCancel
52380   cmdReset.Caption = .OptionsReset
52390   cmdSave.Caption = .OptionsSave
52400   tbstrPDFOptions.Tabs.Clear
52410   tbstrPDFOptions.Tabs.Add , "General", .OptionsPDFGeneral
52420   tbstrPDFOptions.Tabs.Add , "Compression", .OptionsPDFCompression
52430   tbstrPDFOptions.Tabs.Add , "Fonts", .OptionsPDFFonts
52440   tbstrPDFOptions.Tabs.Add , "Colors", .OptionsPDFColors
52450   tbstrPDFOptions.Tabs.Add , "Security", .OptionsPDFSecurity
52460   dmFraPDFGeneral.Caption = .OptionsPDFGeneralCaption
52470   lblPDFCompat.Caption = .OptionsPDFGeneralCompatibility
52480   lblPDFAutoRotate.Caption = .OptionsPDFGeneralAutorotate
52490   lblPDFResolution.Caption = .OptionsPDFGeneralResolution
52500   lblPDFOverprint.Caption = .OptionsPDFGeneralOverprint
52510   lblProgfont.Caption = .OptionsProgramFont
52520   lblProgcharset.Caption = .OptionsProgramFontcharset
52530   lblSize.Caption = .OptionsProgramFontSize
52540   lblTesttext = .OptionsProgramFontTestdescription
52550   cmdTest.Caption = .OptionsProgramFontTest
52560   cmdCancelTest.Caption = .OptionsProgramFontCancelTest
52570   chkPDFASCII85.Caption = .OptionsPDFGeneralASCII85
52580   cmbPDFCompat.Clear
52590   cmbPDFCompat.AddItem .OptionsPDFGeneralCompatibility01
52600   cmbPDFCompat.AddItem .OptionsPDFGeneralCompatibility02
52610   cmbPDFCompat.AddItem .OptionsPDFGeneralCompatibility03
52620   cmbPDFRotate.Clear
52630   cmbPDFRotate.AddItem .OptionsPDFGeneralRotate01
52640   cmbPDFRotate.AddItem .OptionsPDFGeneralRotate02
52650   cmbPDFRotate.AddItem .OptionsPDFGeneralRotate03
52660   cmbPDFOverprint.Clear
52670   cmbPDFOverprint.AddItem .OptionsPDFGeneralOverprint01
52680   cmbPDFOverprint.AddItem .OptionsPDFGeneralOverprint02
52690
52700   dmfraPDFCompress.Caption = .OptionsPDFCompressionCaption
52710   chkPDFTextComp.Caption = .OptionsPDFCompressionTextComp
52720   dmFraPDFColor.Caption = .OptionsPDFCompressionColor
52730   chkPDFColorComp.Caption = .OptionsPDFCompressionColorComp
52740   chkPDFColorResample.Caption = .OptionsPDFCompressionColorResample
52750   lblPDFColorRes.Caption = .OptionsPDFCompressionColorRes
52760   cmbPDFColorComp.Clear
52770   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp01
52780   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp02
52790   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp03
52800   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp04
52810   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp05
52820   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp06
52830   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp07
52840 '  cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp08
52850   cmbPDFColorResample.Clear
52860   cmbPDFColorResample.AddItem .OptionsPDFCompressionColorResample01
52870   cmbPDFColorResample.AddItem .OptionsPDFCompressionColorResample02
52880 '  cmbPDFColorResample.AddItem .OptionsPDFCompressionColorResample03
52890   dmFraPDFGrey.Caption = .OptionsPDFCompressionGrey
52900   chkPDFGreyComp.Caption = .OptionsPDFCompressionGreyComp
52910   chkPDFGreyResample.Caption = .OptionsPDFCompressionGreyResample
52920   lblPDFGreyRes.Caption = .OptionsPDFCompressionGreyRes
52930   cmbPDFGreyComp.Clear
52940   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp01
52950   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp02
52960   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp03
52970   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp04
52980   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp05
52990   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp06
53000   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp07
53010 '  cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp08
53020   cmbPDFGreyResample.Clear
53030   cmbPDFGreyResample.AddItem .OptionsPDFCompressionGreyResample01
53040   cmbPDFGreyResample.AddItem .OptionsPDFCompressionGreyResample02
53050 '  cmbPDFGreyResample.AddItem .OptionsPDFCompressionGreyResample03
53060   dmFraPDFMono.Caption = .OptionsPDFCompressionMono
53070   chkPDFMonoComp.Caption = .OptionsPDFCompressionMonoComp
53080   chkPDFMonoResample.Caption = .OptionsPDFCompressionMonoResample
53090   lblPDFMonoRes.Caption = .OptionsPDFCompressionMonoRes
53100   cmbPDFMonoComp.Clear
53110   cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp01
53120   cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp02
53130   cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp03
53140 '  cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp04
53150   cmbPDFMonoResample.Clear
53160   cmbPDFMonoResample.AddItem .OptionsPDFCompressionMonoResample01
53170   cmbPDFMonoResample.AddItem .OptionsPDFCompressionMonoResample02
53180 '  cmbPDFMonoResample.AddItem .OptionsPDFCompressionMonoResample03
53190
53200   dmFraPDFFonts.Caption = .OptionsPDFFontsCaption
53210   chkPDFEmbedAll.Caption = .OptionsPDFFontsEmbedAll
53220   chkPDFSubSetFonts.Caption = .OptionsPDFFontsSubSetFonts
53230
53240   dmFraPDFColors.Caption = .OptionsPDFColorsCaption
53250   chkPDFCMYKtoRGB.Caption = .OptionsPDFColorsCMYKtoRGB
53260   dmFraPDFColorOptions.Caption = .OptionsPDFColorsColorOptions
53270   chkPDFPreserveOverprint.Caption = .OptionsPDFColorsPreserveOverprint
53280   chkPDFPreserveTransfer.Caption = .OptionsPDFColorsPreserveTransfer
53290   chkPDFPreserveHalftone.Caption = .OptionsPDFColorsPreserveHalftone
53300   cmbPDFColorModel.Clear
53310   cmbPDFColorModel.AddItem .OptionsPDFColorsColorModel01
53320   cmbPDFColorModel.AddItem .OptionsPDFColorsColorModel02
53330   cmbPDFColorModel.AddItem .OptionsPDFColorsColorModel03
53340
53350   dmFraPDFEncryptor.Caption = .OptionsPDFEncryptor
53360   dmFraPDFSecurity.Caption = .OptionsPDFSecurityCaption
53370   chkUseSecurity.Caption = .OptionsPDFUseSecurity
53380   dmFraPDFEncLevel.Caption = .OptionsPDFEncryptionLevel
53390   optEncHigh.Caption = .OptionsPDFEncryptionHigh
53400   optEncLow.Caption = .OptionsPDFEncryptionLow
53410   dmFraSecurityPass.Caption = .OptionsPDFPasswords
53420   chkUserPass.Caption = .OptionsPDFUserPass
53430   chkOwnerPass.Caption = .OptionsPDFOwnerPass
53440   dmFraPDFPermissions.Caption = .OptionsPDFDisallowUser
53450   dmFraPDFHighPermissions.Caption = .OptionsPDFEnhancedPermissions
53460   chkAllowPrinting.Caption = .OptionsPDFDisallowPrint
53470   chkAllowModifyContents.Caption = .OptionsPDFDisallowModify
53480   chkAllowCopy.Caption = .OptionsPDFDisallowCopy
53490   chkAllowModifyAnnotations.Caption = .OptionsPDFDisallowModifyComments
53500   chkAllowDegradedPrinting.Caption = .OptionsPDFAllowDegradedPrinting
53510   chkAllowFillIn.Caption = .OptionsPDFAllowFillIn
53520   chkAllowAssembly.Caption = .OptionsPDFAllowAssembly
53530   chkAllowScreenReaders.Caption = .OptionsPDFAllowScreenReaders
53540
53550   cmbPNGColors.AddItem .OptionsPNGColorscount01
53560   cmbPNGColors.AddItem .OptionsPNGColorscount02
53570   cmbPNGColors.AddItem .OptionsPNGColorscount03
53580   cmbPNGColors.AddItem .OptionsPNGColorscount04
53590   cmbJPEGColors.Left = cmbPNGColors.Left
53600   cmbJPEGColors.Width = cmbPNGColors.Width
53610   cmbJPEGColors.Top = cmbPNGColors.Top
53620   cmbJPEGColors.AddItem .OptionsJPEGColorscount01
53630   cmbJPEGColors.AddItem .OptionsJPEGColorscount02
53640   cmbBMPColors.Left = cmbPNGColors.Left
53650   cmbBMPColors.Width = cmbPNGColors.Width
53660   cmbBMPColors.Top = cmbPNGColors.Top
53670   cmbBMPColors.AddItem .OptionsBMPColorscount01
53680   cmbBMPColors.AddItem .OptionsBMPColorscount02
53690   cmbBMPColors.AddItem .OptionsBMPColorscount03
53700   cmbBMPColors.AddItem .OptionsBMPColorscount04
53710   cmbBMPColors.AddItem .OptionsBMPColorscount05
53720   cmbBMPColors.AddItem .OptionsBMPColorscount06
53730   cmbBMPColors.AddItem .OptionsBMPColorscount07
53740   cmbPCXColors.Left = cmbPNGColors.Left
53750   cmbPCXColors.Width = cmbPNGColors.Width
53760   cmbPCXColors.Top = cmbPNGColors.Top
53770   cmbPCXColors.AddItem .OptionsPCXColorscount01
53780   cmbPCXColors.AddItem .OptionsPCXColorscount02
53790   cmbPCXColors.AddItem .OptionsPCXColorscount03
53800   cmbPCXColors.AddItem .OptionsPCXColorscount04
53810   cmbPCXColors.AddItem .OptionsPCXColorscount05
53820   cmbPCXColors.AddItem .OptionsPCXColorscount06
53830   cmbTIFFColors.Left = cmbPNGColors.Left
53840   cmbTIFFColors.Width = cmbPNGColors.Width
53850   cmbTIFFColors.Top = cmbPNGColors.Top
53860   cmbTIFFColors.AddItem .OptionsTIFFColorscount01
53870   cmbTIFFColors.AddItem .OptionsTIFFColorscount02
53880   cmbTIFFColors.AddItem .OptionsTIFFColorscount03
53890   cmbTIFFColors.AddItem .OptionsTIFFColorscount04
53900   cmbTIFFColors.AddItem .OptionsTIFFColorscount05
53910   cmbTIFFColors.AddItem .OptionsTIFFColorscount06
53920   cmbTIFFColors.AddItem .OptionsTIFFColorscount07
53930   cmbTIFFColors.AddItem .OptionsTIFFColorscount08
53940
53950   dmFraBitmapGeneral.Caption = .OptionsImageSettings
53960   lblBitmapResolution = .OptionsBitmapResolution
53970   lblJPEGQuality = .OptionsJPEGQuality
53980   lblBitmapColors = .OptionsPDFColors
53990   lblProcessPriority.Caption = .OptionsProcesspriority
54000   lblLangLevel.Caption = .OptionsPSLanguageLevel
54010
54020   cmdAsso.Caption = .OptionsAssociatePSFiles
54030  End With
54040
54050  If IsPsAssociate = False Then
54060    cmdAsso.Enabled = True
54070   Else
54080    cmdAsso.Enabled = False
54090  End If
54100
54110  txtPDFRes.Text = 600
54120  cmbPDFCompat.ListIndex = 1
54130  cmbPDFRotate.ListIndex = 0
54140  cmbPDFOverprint.ListIndex = 0
54150  chkPDFASCII85.Value = 0
54160
54170  chkPDFTextComp.Value = 1
54180
54190  chkPDFColorComp.Value = 1
54200  chkPDFColorResample.Value = 0
54210  cmbPDFColorComp.ListIndex = 0
54220  cmbPDFColorResample.ListIndex = 0
54230  txtPDFColorRes.Text = 300
54240
54250  chkPDFGreyComp.Value = 1
54260  chkPDFGreyResample.Value = 0
54270  cmbPDFGreyComp.ListIndex = 0
54280  cmbPDFGreyResample.ListIndex = 0
54290  txtPDFGreyRes.Text = 300
54300
54310  chkPDFMonoComp.Value = 1
54320  chkPDFMonoResample.Value = 0
54330  cmbPDFMonoComp.ListIndex = 0
54340  cmbPDFMonoResample.ListIndex = 0
54350  txtPDFMonoRes.Text = 1200
54360
54370  chkPDFEmbedAll.Value = 1
54380  chkPDFSubSetFonts.Value = 1
54390  txtPDFSubSetPerc.Text = 100
54400
54410  cmbPDFColorModel.ListIndex = 1
54420  chkPDFCMYKtoRGB.Value = 1
54430  chkPDFPreserveOverprint.Value = 1
54440  chkPDFPreserveTransfer.Value = 1
54450  chkPDFPreserveHalftone.Value = 0
54460
54470  cmbPNGColors.ListIndex = 0
54480  cmbJPEGColors.ListIndex = 0
54490  cmbBMPColors.ListIndex = 0
54500  cmbPCXColors.ListIndex = 0
54510  cmbTIFFColors.ListIndex = 0
54520  txtBitmapResolution.Text = 150
54530
54540 ' chkUseStandardAuthor.Value = 1
54550  txtStandardAuthor.Text = vbNullString
54560
54570  With cmbPSLanguageLevel
54580   .AddItem "1"
54590   .AddItem "1.5"
54600   .AddItem "2"
54610   .AddItem "3"
54620  End With
54630  With cmbEPSLanguageLevel
54640   .AddItem "1"
54650   .AddItem "1.5"
54660   .AddItem "2"
54670   .AddItem "3"
54680  End With
54690
54700  With lsvFilenameSubst
54710   .Appearance = ccFlat
54720   .ColumnHeaders.Clear
54730   .ColumnHeaders.Add , "Str1", "", lsvFilenameSubst.Width / 2 - 140
54740   .ColumnHeaders.Add , "Str2", "", lsvFilenameSubst.Width / 2 - 140
54750   .HideColumnHeaders = True
54760   .GridLines = True
54770   .FullRowSelect = True
54780   .HideSelection = False
54790  End With
54800
54810  With cmbPDFEncryptor
54820   .Clear
54830   .AddItem "Ghostscript (>= 8.14)"
54840   .ItemData(.NewIndex) = 0
54850   .AddItem "PDFEnc"
54860   .ItemData(.NewIndex) = 1
54870
54880   SecurityIsPossible = True
54890
54900   If FileExists(GetPDFCreatorApplicationPath & "pdfenc.exe") = False Then
54910    .RemoveItem 1
54920    .ListIndex = 0
54930    Options.PDFEncryptor = .ItemData(.ListIndex)
54940   End If
54950   If GhostScriptSecurity = False Then
54960    .RemoveItem 0
54970   End If
54980   If .ListCount = 0 Then
54990     chkUseSecurity.Value = 0
55000     chkUseSecurity.Enabled = False
55010     SecurityIsPossible = False
55020    Else
55030     For i = 0 To .ListCount - 1
55040      If .ItemData(i) = Options.PDFEncryptor Then
55050       .ListIndex = i
55060       Exit For
55070      End If
55080     Next i
55090     If .ListIndex = -1 Then
55100      .ListIndex = 0
55110      Options.PDFEncryptor = .ItemData(.ListIndex)
55120     End If
55130   End If
55140  End With
55150
55160  If Options.PDFHighEncryption <> 0 Then
55170    optEncHigh.Value = True
55180   Else
55190    optEncLow.Value = True
55200  End If
55210
55220  cmdFilenameSubst(0).Top = lsvFilenameSubst.Top
55230  cmdFilenameSubst(1).Top = lsvFilenameSubst.Top + (lsvFilenameSubst.Height - cmdFilenameSubst(1).Height) / 2
55240  cmdFilenameSubst(2).Top = lsvFilenameSubst.Top + lsvFilenameSubst.Height - cmdFilenameSubst(2).Height
55250
55260  If chkUseStandardAuthor.Value = 1 Then
55270    txtStandardAuthor.Enabled = True
55280    txtStandardAuthor.BackColor = &H80000005
55290   Else
55300    txtStandardAuthor.Enabled = False
55310    txtStandardAuthor.BackColor = &H8000000F
55320  End If
55330  With Options
55340   SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
55350  End With
55360  ieb.Refresh
55370  If chkUseAutosave.Value = 1 Then
55380    ViewAutosave True
55390   Else
55400    ViewAutosave False
55410  End If
55420
55430  With txtGSbin
55440   .ToolTipText = .Text
55450  End With
55460  With txtGSlib
55470   .ToolTipText = .Text
55480  End With
55490  With txtGSfonts
55500   .ToolTipText = .Text
55510  End With
55520  With txtTemppath
55530   .ToolTipText = ResolveEnvironment(GetSubstFilename2(.Text))
55540  End With
55550
55560  With sldProcessPriority
55570   .TextPosition = sldBelowRight
55580   .TickFrequency = 1
55590   .TickStyle = sldTopLeft
55601   Select Case .Value
         Case 0: 'Idle
55620     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityIdle
55630    Case 1: 'Normal
55640     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityNormal
55650    Case 2: 'High
55660     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityHigh
55670    Case 3: 'Realtime
55680     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityRealtime
55690   End Select
55700  End With
55710
55720  If IsWin9xMe = False Then
55730    lblProcessPriority.Enabled = True
55740    sldProcessPriority.Enabled = True
55750   Else
55760    lblProcessPriority.Enabled = False
55770    sldProcessPriority.Enabled = False
55780  End If
55790  UpdateSecurityFields
55800
55810  tStr2 = CompletePath(UCase$(Trim$(Options.DirectoryGhostscriptBinaries)))
 cmbGhostscript.Clear: Set reg = New clsRegistry
55830  reg.hkey = HKEY_LOCAL_MACHINE
55840
55850  Set gsvers = GetAllGhostscriptversions
55860
55870  If gsvers.Count = 0 Then
55880    cmbGhostscript.Enabled = False
55890   Else
55900    For i = 1 To gsvers.Count
55910     cmbGhostscript.AddItem gsvers.Item(i)
55920    Next i
55930    cmbGhostscript.ListIndex = cmbGhostscript.ListCount - 1
55940    For i = 0 To cmbGhostscript.ListCount - 1
55950     tStr = ""
55960     If InStr(cmbGhostscript.List(i), ":") Then
55970       reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
55980       If tStr2 = CompletePath(UCase$(Trim$(reg.GetRegistryValue("GhostscriptDirectoryBinaries")))) Then
55990        cmbGhostscript.ListIndex = i
56000        Exit For
56010       End If
56020      Else
56030       If InStr(UCase$(cmbGhostscript.List(i)), "AFPL") Then
56040        reg.KeyRoot = "SOFTWARE\AFPL Ghostscript"
56050        If InStr(cmbGhostscript.List(i), " ") > 0 Then
56060         tsf = Split(cmbGhostscript.List(i), " ")
56070         reg.Subkey = tsf(UBound(tsf))
56080         tStr = reg.GetRegistryValue("GS_DLL")
56090         If tStr2 & "GSDLL32.DLL" = UCase$(tStr) Then
56100          cmbGhostscript.ListIndex = i
56110          Exit For
56120         End If
56130        End If
56140       End If
56150       If InStr(UCase$(cmbGhostscript.List(i)), "GNU") Then
56160        reg.KeyRoot = "SOFTWARE\GNU Ghostscript"
56170        If InStr(cmbGhostscript.List(i), " ") > 0 Then
56180         tsf = Split(cmbGhostscript.List(i), " ")
56190         reg.Subkey = tsf(UBound(tsf))
56200         tStr = reg.GetRegistryValue("GS_DLL")
56210         If tStr2 & "GSDLL32.DLL" = UCase$(tStr) Then
56220          cmbGhostscript.ListIndex = i
56230          Exit For
56240         End If
56250        End If
56260       End If
56270       If InStr(UCase$(cmbGhostscript.List(i)), "GPL") Then
56280        reg.KeyRoot = "SOFTWARE\GPL Ghostscript"
56290        If InStr(cmbGhostscript.List(i), " ") > 0 Then
56300         tsf = Split(cmbGhostscript.List(i), " ")
56310         reg.Subkey = tsf(UBound(tsf))
56320         tStr = reg.GetRegistryValue("GS_DLL")
56330         If tStr2 & "GSDLL32.DLL" = UCase$(tStr) Then
56340          cmbGhostscript.ListIndex = i
56350          Exit For
56360         End If
56370        End If
56380       End If
56390     End If
56400    Next i
56410  End If
56420  Set reg = Nothing
56430  With cmbGhostscript
56440   If .ListCount = 0 Then
56450    .Enabled = False
56460    .BackColor = &H8000000F
56470   End If
56480  End With
56490
56500  tbstrPDFOptions.ZOrder 1
56510
56520  If ShowOnlyOptions = True Then
56530   FormInTaskbar Me, True, True
56540   Caption = "PDFCreator - " & Caption
56550  End If
56560
56570  ShowAcceleratorsInForm Me, True
56580
56590  Timer1.Enabled = True
56600  Screen.MousePointer = vbNormal
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "Form_Load")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Resize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With cmbCharset
50020   .Top = cmbFonts.Top
50030   .Left = lblProgcharset.Left
50040   .Width = 2295
50050   .SelStart = 0
50060   .SelLength = 0
50070  End With
50080  With cmbProgramFontsize
50090   .Top = cmbFonts.Top
50100   .Left = lblSize.Left
50110   .Width = 765
50120   .SelStart = 0
50130   .SelLength = 0
50140  End With
50150  With cmbGhostscript
50160   .Top = lblGhostscriptversion.Top + lblGhostscriptversion.Height + 20
50170   .Left = lblGhostscriptversion.Left
50180   .Width = 4215
50190  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "Form_Resize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ieb_ItemClick(sGroup As String, sItemKey As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ctl As Control
50020  lblJPEGQuality.Visible = False
50030  cmbPNGColors.Visible = False
50040  cmbJPEGColors.Visible = False
50050  cmbBMPColors.Visible = False
50060  cmbPCXColors.Visible = False
50070  cmbTIFFColors.Visible = False
50080  tbstrPDFOptions.Visible = False
50090  For Each ctl In Controls
50100   If TypeOf ctl Is dmFrame Then
50110    ctl.Visible = False
50120    ctl.Enabled = False
50130   End If
50140  Next
50150  dmFraDescription.Visible = True
50160  dmFraDescription.Enabled = True
50170  tbstrPDFOptions.Enabled = False
50180  txtJPEGQuality.Visible = False
50190  lblJPEQQualityProzent.Visible = False
50200  dmFraPSGeneral.Visible = False
50210  cmbPSLanguageLevel.Visible = False
50220  cmbEPSLanguageLevel.Visible = False
50230
50241  Select Case UCase$(sGroup)
        Case "PROGRAM"
50261    Select Case UCase$(sItemKey)
          Case "GENERAL"
50280      Set picOptions = LoadResPicture(2101, vbResIcon)
50290      lblOptions = LanguageStrings.OptionsProgramGeneralDescription
50300      dmFraProgGeneral.Enabled = True
50310      dmFraShellIntegration.Enabled = True
50320      dmFraProgGeneral.Visible = True
50330      dmFraShellIntegration.Visible = True
50340      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50350     Case "GHOSTSCRIPT"
50360      Set picOptions = LoadResPicture(2119, vbResIcon)
50370      lblOptions = LanguageStrings.OptionsProgramGhostscriptDescription
50380      dmFraProgGhostscript.Enabled = True
50390      dmFraProgGhostscript.Visible = True
50400      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50410     Case "DOCUMENT"
50420      Set picOptions = LoadResPicture(2105, vbResIcon)
50430      lblOptions = LanguageStrings.OptionsProgramDocumentDescription
50440      dmFraProgDocument.Enabled = True
50450      dmFraProgDocument.Visible = True
50460      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50470     Case "SAVE"
50480      Set picOptions = LoadResPicture(2106, vbResIcon)
50490      lblOptions = LanguageStrings.OptionsProgramSaveDescription
50500      dmfraProgSave.Enabled = True
50510      dmfraProgSave.Visible = True
50520      dmfraFilenameSubstitutions.Visible = True
50530      dmfraFilenameSubstitutions.Enabled = True
50540      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50550     Case "AUTOSAVE"
50560      Set picOptions = LoadResPicture(2103, vbResIcon)
50570      lblOptions = LanguageStrings.OptionsProgramAutosaveDescription
50580      dmFraProgAutosave.Enabled = True
50590      dmFraProgAutosave.Visible = True
50600      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50610     Case "DIRECTORIES"
50620      Set picOptions = LoadResPicture(2104, vbResIcon)
50630      lblOptions = LanguageStrings.OptionsProgramDirectoriesDescription
50640      dmFraProgDirectories.Enabled = True
50650      dmFraProgDirectories.Visible = True
50660      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50670     Case "FONTS"
50680      Set picOptions = LoadResPicture(2102, vbResIcon)
50690      lblOptions = LanguageStrings.OptionsProgramFontDescription
50700      dmFraProgFont.Enabled = True
50710      dmFraProgFont.Visible = True
50720      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50730    End Select
50740   Case "FORMATS"
50751    Select Case UCase$(sItemKey)
          Case "PDF"
50770      Set picOptions = LoadResPicture(2111, vbResIcon)
50780      lblOptions = LanguageStrings.OptionsPDFDescription
50790      tbstrPDFOptions.Enabled = True
50800      tbstrPDFOptions.Visible = True
50810      dmFraPDFGeneral.Enabled = True
50820      dmFraPDFGeneral.Visible = True
50830      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
50840      dmFraPDFGeneral.Enabled = True
50850     Case "PNG"
50860      Set picOptions = LoadResPicture(2112, vbResIcon)
50870      lblOptions = LanguageStrings.OptionsPNGDescription
50880      dmFraBitmapGeneral.Enabled = True
50890      dmFraBitmapGeneral.Visible = True
50900      cmbPNGColors.Visible = True
50910      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
50920     Case "JPEG"
50930      Set picOptions = LoadResPicture(2113, vbResIcon)
50940      lblOptions = LanguageStrings.OptionsJPEGDescription
50950      dmFraBitmapGeneral.Enabled = True
50960      dmFraBitmapGeneral.Visible = True
50970      lblJPEGQuality.Caption = LanguageStrings.OptionsJPEGQuality
50980      lblJPEGQuality.Visible = True
50990      txtJPEGQuality.Visible = True
51000      lblJPEQQualityProzent.Visible = True
51010      lblJPEQQualityProzent.Left = txtJPEGQuality.Left + txtJPEGQuality.Width + 100
51020      cmbJPEGColors.Visible = True
51030      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51040     Case "BMP"
51050      Set picOptions = LoadResPicture(2114, vbResIcon)
51060      lblOptions = LanguageStrings.OptionsBMPDescription
51070      dmFraBitmapGeneral.Enabled = True
51080      dmFraBitmapGeneral.Visible = True
51090      cmbBMPColors.Visible = True
51100      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51110     Case "PCX"
51120      Set picOptions = LoadResPicture(2115, vbResIcon)
51130      lblOptions = LanguageStrings.OptionsPCXDescription
51140      dmFraBitmapGeneral.Enabled = True
51150      dmFraBitmapGeneral.Visible = True
51160      cmbPCXColors.Visible = True
51170      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51180     Case "TIFF"
51190      Set picOptions = LoadResPicture(2116, vbResIcon)
51200      lblOptions = LanguageStrings.OptionsTIFFDescription
51210      dmFraBitmapGeneral.Enabled = True
51220      dmFraBitmapGeneral.Visible = True
51230      cmbTIFFColors.Visible = True
51240      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51250     Case "PS"
51260      Set picOptions = LoadResPicture(2117, vbResIcon)
51270      lblOptions.Caption = LanguageStrings.OptionsPSDescription
51280      dmFraPSGeneral.Enabled = True
51290      dmFraPSGeneral.Visible = True
51300      cmbPSLanguageLevel.Visible = True
51310      dmFraPSGeneral.Caption = LanguageStrings.OptionsPSDescription
51320      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51330     Case "EPS"
51340      Set picOptions = LoadResPicture(2118, vbResIcon)
51350      lblOptions.Caption = LanguageStrings.OptionsEPSDescription
51360      dmFraPSGeneral.Enabled = True
51370      dmFraPSGeneral.Visible = True
51380      cmbEPSLanguageLevel.Visible = True
51390      dmFraPSGeneral.Caption = LanguageStrings.OptionsEPSDescription
51400      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51410    End Select
51420  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "ieb_ItemClick")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub lsvFilenameSubst_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Set_txtFilenameSubst
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "lsvFilenameSubst_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub optEncHigh_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  UpdateSecurityFields
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "optEncHigh_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub optEncLow_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  UpdateSecurityFields
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "optEncLow_Click")
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
Select Case ErrPtnr.OnError("frmOptions", "sldProcessPriority_Change")
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
50021   Select Case .Value
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
Select Case ErrPtnr.OnError("frmOptions", "sldProcessPriority_Scroll")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub tbstrPDFOptions_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  dmFraPDFGeneral.Visible = False
50020  dmfraPDFCompress.Visible = False
50030  dmFraPDFFonts.Visible = False
50040  dmFraPDFColors.Visible = False
50050  dmFraPDFColorOptions.Visible = False
50060  dmFraPDFSecurity.Visible = False
50070  dmFraPDFGeneral.Enabled = False
50080  dmfraPDFCompress.Enabled = False
50090  dmFraPDFFonts.Enabled = False
50100  dmFraPDFColors.Enabled = False
50110  dmFraPDFColorOptions.Enabled = False
50120  dmFraPDFSecurity.Enabled = False
50131  Select Case tbstrPDFOptions.SelectedItem.Index
        Case 1:
50150    dmFraPDFGeneral.Visible = True
50160    dmFraPDFGeneral.Enabled = True
50170   Case 2:
50180    dmfraPDFCompress.Visible = True
50190    dmfraPDFCompress.Enabled = True
50200    dmFraPDFColor.Visible = True
50210    dmFraPDFColor.Enabled = True
50220    dmFraPDFGrey.Visible = True
50230    dmFraPDFGrey.Enabled = True
50240    dmFraPDFMono.Visible = True
50250    dmFraPDFMono.Enabled = True
50260   Case 3:
50270    dmFraPDFFonts.Visible = True
50280    dmFraPDFFonts.Enabled = True
50290   Case 4:
50300    dmFraPDFColors.Visible = True
50310    dmFraPDFColorOptions.Visible = True
50320    dmFraPDFColors.Enabled = True
50330    dmFraPDFColorOptions.Enabled = True
50340   Case 5:
50350    dmFraPDFSecurity.Visible = True
50360    dmFraPDFSecurity.Enabled = True
50370    dmFraPDFEncryptor.Visible = True
50380    dmFraPDFEncryptor.Enabled = True
50390    dmFraPDFEncLevel.Visible = True
50400    dmFraPDFEncLevel.Enabled = True
50410    dmFraSecurityPass.Visible = True
50420    dmFraSecurityPass.Enabled = True
50430    dmFraPDFPermissions.Visible = True
50440    dmFraPDFPermissions.Enabled = True
50450    dmFraPDFHighPermissions.Visible = True
50460    dmFraPDFHighPermissions.Enabled = True
50470    If SecurityIsPossible = False Then
50480     MsgBox LanguageStrings.MessagesMsg19
50490    End If
50500  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "tbstrPDFOptions_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Timer1_Timer()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, fi As Long, tStr As String, SMF As Collection, _
  cSystem As clsSystem, ctl As Control
50030  Timer1.Enabled = False
50040  Set cSystem = New clsSystem
50050  Set SMF = cSystem.GetSystemFont(Me, Menu)
50060  txtTest.Text = vbNullString
50070  For i = 33 To 255
50080   txtTest.Text = txtTest.Text & Chr$(i)
50090  Next i
50100  fi = -1
50110  With cmbFonts
50120   .Clear
50130   For i = 1 To Screen.FontCount
50140    tStr = Trim$(Screen.Fonts(i))
50150    If Len(tStr) > 0 Then
50160     cmbFonts.AddItem tStr
50170    End If
50180   Next i
50190   If .ListCount > 0 Then
50200     For i = 0 To cmbFonts.ListCount - 1
50210      If SMF.Count > 0 Then
50220       If UCase$(cmbFonts.List(i)) = UCase$(SMF(1)(0)) Then
50230        fi = i
50240       End If
50250      End If
50260     Next i
50270    Else
50280    .ListIndex = 0
50290   End If
50300  End With
50310  With cmbCharset
50320   .Clear
50330   .AddItem "0, Western": .ItemData(.NewIndex) = 0
50340   .AddItem "2, Symbol": .ItemData(.NewIndex) = 2
50350   .AddItem "77, Mac": .ItemData(.NewIndex) = 77
50360   .AddItem "161, Greek": .ItemData(.NewIndex) = 161
50370   .AddItem "162, Turkish": .ItemData(.NewIndex) = 162
50380   .AddItem "177, Hebrew": .ItemData(.NewIndex) = 177
50390   .AddItem "178, Arabic": .ItemData(.NewIndex) = 178
50400   .AddItem "186, Baltic": .ItemData(.NewIndex) = 186
50410   .AddItem "204, Cyrillic": .ItemData(.NewIndex) = 204
50420   .AddItem "238, Central European": .ItemData(.NewIndex) = 238
50430   .AddItem "255, DOS/OEM": .ItemData(.NewIndex) = 255
50440   .Text = 0
50450  End With
50460  With cmbProgramFontsize
50470   .AddItem "8"
50480   .AddItem "9"
50490   .AddItem "10"
50500   .AddItem "11"
50510   .AddItem "12"
50520   .AddItem "14"
50530   .AddItem "16"
50540   .AddItem "18"
50550   .AddItem "20"
50560   .AddItem "22"
50570   .AddItem "24"
50580   .AddItem "26"
50590   .AddItem "28"
50600   .AddItem "36"
50610   .AddItem "48"
50620   .AddItem "72"
50630  End With
50640  cmbProgramFontsize.Text = 8
50650  cmbCharset.Text = cmbCharset.ItemData(0)
50660  cmbCharset.Text = Options.ProgramFontCharset
50670  For Each ctl In Controls
50680   If TypeOf ctl Is ComboBox Then
50690    ComboSetListWidth ctl
50700   End If
50710  Next ctl
50720
50730  SetOptimalComboboxHeigth cmbCharset, Me
50740  SetOptimalComboboxHeigth cmbProgramFontsize, Me
50750  SetOptimalComboboxHeigth cmbGhostscript, Me
50760
50770  Form_Resize
50780
50790  cmbProgramFontsize.Width = txtTest.Width - _
  (cmbProgramFontsize.Left - txtTest.Left)
50810
50820  If fi >= 0 Then
50830   cmbFonts.ListIndex = fi
50840   cmbCharset.Text = SMF(1)(2)
50850   cmbProgramFontsize.Text = SMF(1)(1)
50860   txtTest.Font = cmbFonts.List(cmbFonts.ListIndex)
50870   txtTest.Font.Charset = cmbCharset.Text
50880  End If
50890
50900  ShowOptions Me, Options
50910
50920  If Options.UseAutosaveDirectory = "1" Then
50930    ViewAutosaveDirectory True
50940   Else
50950    ViewAutosaveDirectory False
50960  End If
50970  If Options.UseAutosave = "1" Then
50980    ViewAutosave True
50990   Else
51000    ViewAutosave False
51010  End If
51020
51030  CheckCmdFilenameSubst
51040  CorrectCmbCharset
51050 ' Call ieb_ItemClick("PROGRAM", "GENERAL")
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "Timer1_Timer")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtAutosaveDirectory_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txtAutosaveDirectory.ToolTipText = txtAutosaveDirectory.Text
50020  With txtAutoSaveDirectoryPreview
50030   .Text = GetSubstFilename2(txtAutosaveDirectory.Text)
50040   .ToolTipText = .Text
50050   If IsValidPath(.Text) = False Then
50060     .ForeColor = vbRed
50070    Else
50080     .ForeColor = &H80000008
50090   End If
50100  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "txtAutosaveDirectory_Change")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtAutosaveFilename_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Ext As String
50020  txtAutosaveFilename.ToolTipText = txtAutosaveFilename.Text
50030  With txtAutoSaveFilenamePreview
50040   .Text = GetSubstFilename("B:\dummy.dum", txtAutosaveFilename.Text, , True) & GetAutosaveFormatExtension
50050   .ToolTipText = .Text
50060   If IsValidPath("C:\" & .Text) = False Then
50070     .ForeColor = vbRed
50080    Else
50090     .ForeColor = &H80000008
50100   End If
50110  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "txtAutosaveFilename_Change")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbProgramFontSize_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tL As Long
50020 If Trim$(cmbProgramFontsize.Text) = vbNullString Then
50030   cmbProgramFontsize.Text = 8
50040  End If
50050  tL = CLng(cmbProgramFontsize.Text)
50060  If tL <= 0 Then
50070   tL = 1
50080  End If
50090  If tL > 72 Then
50100   tL = 72
50110  End If
50120  cmbProgramFontsize.Text = tL
50130  txtTest.Font.Size = tL
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmbProgramFontSize_Change")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbProgramFontSize_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim allow As String, tStr As String
50020
50030  allow = "0123456789" & Chr$(8) & Chr$(13)
50040
50050  tStr = Chr$(KeyAscii)
50060
50070  If InStr(1, allow, tStr) = 0 Then
50080    KeyAscii = 0
50090  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmbProgramFontSize_KeyPress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbProgramFontsize_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tL As Long
50020 If Trim$(cmbProgramFontsize.Text) = vbNullString Then
50030   cmbProgramFontsize.Text = 8
50040  End If
50050  tL = CLng(cmbProgramFontsize.Text)
50060  If tL <= 0 Then
50070   tL = 1
50080  End If
50090  If tL > 72 Then
50100   tL = 72
50110  End If
50120  cmbProgramFontsize.Text = tL
50130  txtTest.Font.Size = tL
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmbProgramFontsize_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ViewAutosave(ViewIt As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  lblAutosaveformat.Enabled = ViewIt
50020  cmbAutosaveFormat.Enabled = ViewIt
50030  lblAutosaveFilename.Enabled = ViewIt
50040  txtAutosaveFilename.Enabled = ViewIt
50050  txtAutoSaveFilenamePreview.Enabled = ViewIt
50060  lblAutosaveFilenameTokens.Enabled = ViewIt
50070  cmbAutoSaveFilenameTokens.Enabled = ViewIt
50080  chkUseAutosaveDirectory.Enabled = ViewIt
50090  txtAutoSaveDirectoryPreview.Enabled = ViewIt
50100  If ViewIt = True Then
50110    cmbAutosaveFormat.BackColor = &H80000005
50120    cmbAutoSaveFilenameTokens.BackColor = &H80000005
50130    txtAutosaveFilename.BackColor = &H80000005
50140    txtAutosaveDirectory.BackColor = &H80000005
50150   Else
50160    cmbAutosaveFormat.BackColor = &H8000000F
50170    cmbAutoSaveFilenameTokens.BackColor = &H8000000F
50180    txtAutosaveFilename.BackColor = &H8000000F
50190    txtAutosaveDirectory.BackColor = &H8000000F
50200  End If
50210  If chkUseAutosaveDirectory.Value = 1 And ViewIt = True Then
50220    ViewAutosaveDirectory True
50230   Else
50240    ViewAutosaveDirectory False
50250  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "ViewAutosave")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ViewAutosaveDirectory(ViewIt As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txtAutosaveDirectory.Enabled = ViewIt
50020  txtAutoSaveDirectoryPreview.Enabled = ViewIt
50030  cmdGetAutosaveDirectory.Enabled = ViewIt
50040  If ViewIt = True Then
50050    txtAutosaveDirectory.BackColor = &H80000005
50060   Else
50070    txtAutosaveDirectory.BackColor = &H8000000F
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "ViewAutosaveDirectory")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub UpdateSecurityFields()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If cmbPDFCompat.ListIndex < 2 Then
50020   optEncLow.Value = True
50030  End If
50040  If chkUseSecurity.Value = False Then
50050    dmFraPDFEncryptor.Enabled = False
50060    cmbPDFEncryptor.Enabled = False
50070
50080    dmFraPDFEncLevel.Enabled = False
50090    optEncHigh.Enabled = False
50100    optEncLow.Enabled = False
50110
50120    dmFraSecurityPass.Enabled = False
50130    chkUserPass.Enabled = False
50140    chkOwnerPass.Enabled = False
50150
50160    dmFraPDFPermissions.Enabled = False
50170    chkAllowPrinting.Enabled = False
50180    chkAllowCopy.Enabled = False
50190    chkAllowModifyAnnotations.Enabled = False
50200    chkAllowModifyContents.Enabled = False
50210
50220    dmFraPDFHighPermissions.Enabled = False
50230    chkAllowDegradedPrinting.Enabled = False
50240    chkAllowFillIn.Enabled = False
50250    chkAllowScreenReaders.Enabled = False
50260    chkAllowAssembly.Enabled = False
50270   Else
50280    dmFraPDFEncryptor.Enabled = True
50290    cmbPDFEncryptor.Enabled = True
50300
50310    dmFraPDFEncLevel.Enabled = True
50320    If cmbPDFCompat.ListIndex >= 2 Then
50330      optEncHigh.Enabled = True
50340     Else
50350      optEncHigh.Enabled = False
50360    End If
50370    optEncLow.Enabled = True
50380
50390    dmFraSecurityPass.Enabled = True
50400    chkUserPass.Enabled = True
50410    chkOwnerPass.Enabled = True
50420
50430    dmFraPDFPermissions.Enabled = True
50440    chkAllowPrinting.Enabled = True
50450    chkAllowCopy.Enabled = True
50460    chkAllowModifyAnnotations.Enabled = True
50470    chkAllowModifyContents.Enabled = True
50480
50490    If optEncHigh.Value = True Then
50500      dmFraPDFHighPermissions.Enabled = True
50510      chkAllowDegradedPrinting.Enabled = True
50520      chkAllowFillIn.Enabled = True
50530      chkAllowScreenReaders.Enabled = True
50540      chkAllowAssembly.Enabled = True
50550     Else
50560      dmFraPDFHighPermissions.Enabled = False
50570      chkAllowDegradedPrinting.Enabled = False
50580      chkAllowFillIn.Enabled = False
50590      chkAllowScreenReaders.Enabled = False
50600      chkAllowAssembly.Enabled = False
50610    End If
50620  End If
50630  If chkOwnerPass.Value = 0 And chkUserPass.Value = 0 Then
50640   chkOwnerPass.Value = 1: Options.PDFOwnerPass = 1
50650  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "UpdateSecurityFields")
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
Select Case ErrPtnr.OnError("frmOptions", "cmdAsso_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub AddFilenameSubstitutions()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, res As Long
50020  res = CheckFilenameSubstitutions(0)
50031  Select Case res
              Case 0, 2:
50050    lsvFilenameSubst.ListItems.Add , , txtFilenameSubst(0).Text
50060    lsvFilenameSubst.ListItems(lsvFilenameSubst.ListItems.Count).SubItems(1) = txtFilenameSubst(1).Text
50070    lsvFilenameSubst.ListItems(lsvFilenameSubst.ListItems.Count).Selected = True
50080    lsvFilenameSubst.ListItems(lsvFilenameSubst.ListItems.Count).EnsureVisible
50090    Set_txtFilenameSubst
50100 '  Case 2:
50110 '   MsgBox LanguageStrings.MessagesMsg12 & _
    vbCrLf & vbTab & "\ / : * ? < > | """
50130   Case 3:
50140    MsgBox LanguageStrings.MessagesMsg11
50150  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "AddFilenameSubstitutions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ChangeFilenameSubstitutions()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, res As Long
50020  res = CheckFilenameSubstitutions(lsvFilenameSubst.SelectedItem.Index)
50031  Select Case res
              Case 0, 2:
50050    lsvFilenameSubst.SelectedItem.Text = txtFilenameSubst(0).Text
50060    lsvFilenameSubst.SelectedItem.SubItems(1) = txtFilenameSubst(1).Text
50070 '  Case 2:
50080 '   MsgBox LanguageStrings.MessagesMsg12 & _
    vbCrLf & vbTab & "\ / : * ? < > | """
50100   Case 3:
50110    MsgBox LanguageStrings.MessagesMsg11
50120  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "ChangeFilenameSubstitutions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub DeleteFilenameSubstitutions()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim oIndex As Long
50020  With lsvFilenameSubst
50030   oIndex = .SelectedItem.Index
50040   If .ListItems.Count > 0 Then
50050    .ListItems.Remove .SelectedItem.Index
50060    If oIndex > .ListItems.Count Then
50070     oIndex = .ListItems.Count
50080    End If
50090    If .ListItems.Count > 0 Then
50100     .ListItems(oIndex).Selected = True
50110     .ListItems(oIndex).EnsureVisible
50120    End If
50130    Set_txtFilenameSubst
50140   End If
50150  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "DeleteFilenameSubstitutions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub MoveUpFilenameSubstitutions()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tStrL As String, tStrR As String
50020  With lsvFilenameSubst
50030   tStrL = .ListItems(.SelectedItem.Index).Text
50040   tStrR = .ListItems(.SelectedItem.Index).SubItems(1)
50050   .ListItems(.SelectedItem.Index).Text = .ListItems(.SelectedItem.Index - 1).Text
50060   .ListItems(.SelectedItem.Index).SubItems(1) = .ListItems(.SelectedItem.Index - 1).SubItems(1)
50070   .ListItems(.SelectedItem.Index - 1).Text = tStrL
50080   .ListItems(.SelectedItem.Index - 1).SubItems(1) = tStrR
50090   .ListItems(.SelectedItem.Index - 1).Selected = True
50100   .ListItems(.SelectedItem.Index).EnsureVisible
50110  End With
50120  Set_txtFilenameSubst
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "MoveUpFilenameSubstitutions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub MoveDownFilenameSubstitutions()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tStrL As String, tStrR As String
50020  With lsvFilenameSubst
50030   tStrL = .ListItems(.SelectedItem.Index).Text
50040   tStrR = .ListItems(.SelectedItem.Index).SubItems(1)
50050   .ListItems(.SelectedItem.Index).Text = .ListItems(.SelectedItem.Index + 1).Text
50060   .ListItems(.SelectedItem.Index).SubItems(1) = .ListItems(.SelectedItem.Index + 1).SubItems(1)
50070   .ListItems(.SelectedItem.Index + 1).Text = tStrL
50080   .ListItems(.SelectedItem.Index + 1).SubItems(1) = tStrR
50090   .ListItems(.SelectedItem.Index + 1).Selected = True
50100   .ListItems(.SelectedItem.Index).EnsureVisible
50110  End With
50120  Set_txtFilenameSubst
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "MoveDownFilenameSubstitutions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function CheckFilenameSubstitutions(Index As Long) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long
50020  CheckFilenameSubstitutions = 0
50030  If Len(txtFilenameSubst(0).Text) = 0 Then
50040   CheckFilenameSubstitutions = 1
50050   Exit Function
50060  End If
50070  If IsForbiddenChars(txtFilenameSubst(0).Text) = True Then
50080   txtFilenameSubst(0).SetFocus
50090   CheckFilenameSubstitutions = 2
50100   Exit Function
50110  End If
50120  If IsForbiddenChars(txtFilenameSubst(1).Text) = True Then
50130   txtFilenameSubst(1).SetFocus
50140   CheckFilenameSubstitutions = 2
50150   Exit Function
50160  End If
50170  If Index = 0 Then
50180    For i = 1 To lsvFilenameSubst.ListItems.Count
50190     If UCase$(txtFilenameSubst(0).Text) = UCase$(lsvFilenameSubst.ListItems(i).Text) Then
50200      CheckFilenameSubstitutions = 3
50210      Exit Function
50220     End If
50230    Next i
50240   Else
50250    For i = 1 To lsvFilenameSubst.ListItems.Count
50260     If UCase$(txtFilenameSubst(0).Text) = UCase$(lsvFilenameSubst.ListItems(i).Text) And _
     Index <> lsvFilenameSubst.SelectedItem.Index Then
50280      CheckFilenameSubstitutions = 3
50290      Exit Function
50300     End If
50310    Next i
50320  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "CheckFilenameSubstitutions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub CheckCmdFilenameSubst()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If lsvFilenameSubst.ListItems.Count > 0 Then
50020    cmdFilenameSubst(1).Enabled = True
50030    cmdFilenameSubst(2).Enabled = True
50040   Else
50050    cmdFilenameSubst(1).Enabled = False
50060    cmdFilenameSubst(2).Enabled = False
50070  End If
50080  If lsvFilenameSubst.ListItems.Count > 1 Then
50090    cmdFilenameSubstMove(0).Enabled = True
50100    cmdFilenameSubstMove(1).Enabled = True
50110   Else
50120    cmdFilenameSubstMove(0).Enabled = False
50130    cmdFilenameSubstMove(1).Enabled = False
50140  End If
50150  If lsvFilenameSubst.ListItems.Count > 0 Then
50160   If lsvFilenameSubst.SelectedItem.Index = 1 Then
50170    cmdFilenameSubstMove(0).Enabled = False
50180   End If
50190   If lsvFilenameSubst.SelectedItem.Index = lsvFilenameSubst.ListItems.Count Then
50200    cmdFilenameSubstMove(1).Enabled = False
50210   End If
50220  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "CheckCmdFilenameSubst")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Set_txtFilenameSubst()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  CheckCmdFilenameSubst
50020  If lsvFilenameSubst.ListItems.Count > 0 Then
50030   txtFilenameSubst(0).Text = lsvFilenameSubst.SelectedItem.Text
50040   txtFilenameSubst(0).ToolTipText = txtFilenameSubst(0).Text
50050   txtFilenameSubst(1).Text = lsvFilenameSubst.SelectedItem.SubItems(1)
50060   txtFilenameSubst(1).ToolTipText = txtFilenameSubst(1).Text
50070  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "Set_txtFilenameSubst")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub


Private Sub txtSaveFilename_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txtSaveFilename.ToolTipText = txtSaveFilename.Text
50020  With txtSavePreview
50030   .Text = GetSubstFilename("C:\test.pdf", txtSaveFilename.Text, , True) & ".pdf"
50040   .ToolTipText = .Text
50050  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "txtSaveFilename_Change")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function GetAutosaveFormatExtension() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case cmbAutosaveFormat.ListIndex
        Case -1, 0
50030    GetAutosaveFormatExtension = ".pdf"
50040   Case 1
50050    GetAutosaveFormatExtension = ".png"
50060   Case 2
50070    GetAutosaveFormatExtension = ".jpg"
50080   Case 3
50090    GetAutosaveFormatExtension = ".bmp"
50100   Case 4
50110    GetAutosaveFormatExtension = ".pcx"
50120   Case 5
50130    GetAutosaveFormatExtension = ".tif"
50140   Case 6
50150    GetAutosaveFormatExtension = ".ps"
50160   Case 7
50170    GetAutosaveFormatExtension = ".eps"
50180  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "GetAutosaveFormatExtension")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub CorrectCmbCharset()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tStrf() As String
50020  If InStr(1, cmbCharset.Text, ",", vbTextCompare) > 0 Then
50030    tStrf = Split(cmbCharset.Text, ",")
50040    If Len(tStrf(0)) = 0 Then
50050      cmbCharset.Text = 0
50060     Else
50070      If IsNumeric(tStrf(0)) = False Then
50080        cmbCharset.Text = 0
50090       Else
50100        cmbCharset.Text = CLng(tStrf(0))
50110      End If
50120    End If
50130   Else
50140    If Len(cmbCharset.Text) = 0 Then
50150      cmbCharset.Text = 0
50160     Else
50170      If IsNumeric(cmbCharset.Text) = False Then
50180        cmbCharset.Text = 0
50190       Else
50200        cmbCharset.Text = CLng(cmbCharset.Text)
50210      End If
50220    End If
50230  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "CorrectCmbCharset")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtStandardAuthor_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txtStandardAuthor.ToolTipText = txtStandardAuthor.Text
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "txtStandardAuthor_Change")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtTemppath_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txtTemppath.ToolTipText = txtTemppath.Text
50020  With txtTemppathPreview
50030   .Text = ResolveEnvironment(GetSubstFilename2(txtTemppath.Text))
50040   .ToolTipText = .Text
50050  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "txtTemppath_Change")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
