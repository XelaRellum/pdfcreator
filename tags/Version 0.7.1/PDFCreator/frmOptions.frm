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
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame fraProgSave 
      Caption         =   "Save"
      Height          =   4095
      Left            =   2760
      TabIndex        =   132
      Top             =   960
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CheckBox chkSpaces 
         Caption         =   "Remove leading and trailing spaces"
         Height          =   255
         Left            =   120
         TabIndex        =   145
         Top             =   960
         Value           =   1  'Aktiviert
         Width           =   6015
      End
      Begin VB.Frame fraFilenameSubstitutions 
         Caption         =   "Filename substitutions"
         Height          =   2415
         Left            =   120
         TabIndex        =   137
         Top             =   1560
         Width           =   6015
         Begin MSComctlLib.ListView lsvFilenameSubst 
            Height          =   1335
            Left            =   600
            TabIndex        =   148
            Top             =   600
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
         Begin VB.TextBox txtFilenameSubst 
            Height          =   285
            Index           =   1
            Left            =   2640
            TabIndex        =   146
            Top             =   240
            Width           =   1695
         End
         Begin VB.CheckBox chkFilenameSubst 
            Caption         =   "Substitutions only in <Title>"
            Height          =   255
            Left            =   240
            TabIndex        =   144
            Top             =   2040
            Value           =   1  'Aktiviert
            Width           =   3255
         End
         Begin VB.CommandButton cmdFilenameSubst 
            Enabled         =   0   'False
            Height          =   375
            Index           =   4
            Left            =   120
            Picture         =   "frmOptions.frx":08CA
            Style           =   1  'Grafisch
            TabIndex        =   143
            Top             =   1320
            Width           =   375
         End
         Begin VB.CommandButton cmdFilenameSubst 
            Enabled         =   0   'False
            Height          =   375
            Index           =   3
            Left            =   120
            Picture         =   "frmOptions.frx":0A14
            Style           =   1  'Grafisch
            TabIndex        =   142
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox txtFilenameSubst 
            Height          =   285
            Index           =   0
            Left            =   600
            TabIndex        =   141
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmdFilenameSubst 
            Caption         =   "Add"
            Height          =   375
            Index           =   0
            Left            =   4440
            TabIndex        =   140
            Top             =   600
            Width           =   1455
         End
         Begin VB.CommandButton cmdFilenameSubst 
            Caption         =   "Change"
            Height          =   375
            Index           =   1
            Left            =   4440
            TabIndex        =   139
            Top             =   1080
            Width           =   1455
         End
         Begin VB.CommandButton cmdFilenameSubst 
            Caption         =   "Delete"
            Height          =   375
            Index           =   2
            Left            =   4440
            TabIndex        =   138
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label lblEqual 
            Caption         =   "="
            Height          =   255
            Left            =   2400
            TabIndex        =   147
            Top             =   240
            Width           =   135
         End
      End
      Begin VB.TextBox txtSaveFilename 
         Height          =   285
         Left            =   120
         TabIndex        =   134
         Text            =   "<Title>"
         Top             =   480
         Width           =   3495
      End
      Begin VB.ComboBox cmbSaveFilenameTokens 
         Height          =   315
         ItemData        =   "frmOptions.frx":0B5E
         Left            =   3720
         List            =   "frmOptions.frx":0B60
         Style           =   2  'Dropdown-Liste
         TabIndex        =   133
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblSaveFilename 
         Caption         =   "Filename"
         Height          =   255
         Left            =   120
         TabIndex        =   136
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label lblSaveFilenameTokens 
         Caption         =   "Add a Filename-Token"
         Height          =   255
         Left            =   3720
         TabIndex        =   135
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame fraProgAutosave 
      Caption         =   "Autosave"
      Height          =   2895
      Left            =   2640
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   6255
      Begin VB.ComboBox cmbAutosaveFormat 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown-Liste
         TabIndex        =   20
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox cmbAutoSaveFilenameTokens 
         Height          =   315
         ItemData        =   "frmOptions.frx":0B62
         Left            =   3720
         List            =   "frmOptions.frx":0B64
         Style           =   2  'Dropdown-Liste
         TabIndex        =   47
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txtAutosaveFilename 
         Height          =   285
         Left            =   120
         TabIndex        =   44
         Text            =   "<DateTime>"
         Top             =   1680
         Width           =   3495
      End
      Begin VB.CommandButton cmdGetAutosaveDirectory 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   5760
         TabIndex        =   74
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtAutosaveDirectory 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   2520
         Width           =   5535
      End
      Begin VB.CheckBox chkUseAutosaveDirectory 
         Caption         =   "For autosave use this directory"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   2160
         Width           =   5895
      End
      Begin VB.CheckBox chkUseAutosave 
         Caption         =   "Use Autosave"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   6015
      End
      Begin VB.Label lblAutosaveFilenameTokens 
         Caption         =   "Add a Filename-Token"
         Height          =   255
         Left            =   3720
         TabIndex        =   40
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label lblAutosaveformat 
         Caption         =   "Autosaveformat"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label lblAutosaveFilename 
         Caption         =   "Filename"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1440
         Width           =   3495
      End
   End
   Begin VB.Frame fraProgGeneral 
      Caption         =   "General"
      Height          =   2775
      Left            =   2880
      TabIndex        =   6
      Top             =   960
      Width           =   6255
      Begin VB.CommandButton cmdAsso 
         Caption         =   "Associate PDFCreator with Postscript files"
         Height          =   375
         Left            =   120
         TabIndex        =   127
         Top             =   360
         Width           =   3855
      End
      Begin MSComctlLib.Slider sldProcessPriority 
         Height          =   495
         Left            =   120
         TabIndex        =   85
         Top             =   1440
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   1
         Max             =   3
         SelStart        =   1
         Value           =   1
      End
      Begin VB.CommandButton cmdTestpage 
         Caption         =   "Print testpage"
         Height          =   375
         Left            =   120
         TabIndex        =   57
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label lblProcessPriority 
         Caption         =   "Processpriority: Normal"
         Height          =   255
         Left            =   120
         TabIndex        =   75
         Top             =   1080
         Width           =   2535
      End
   End
   Begin VB.Frame fraProgDocument 
      Caption         =   "Document"
      Height          =   1935
      Left            =   2640
      TabIndex        =   128
      Top             =   1560
      Width           =   6255
      Begin VB.CheckBox chkUseStandardAuthor 
         Caption         =   "Use Standardauthor"
         Height          =   255
         Left            =   120
         TabIndex        =   131
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox txtStandardAuthor 
         Height          =   285
         Left            =   120
         TabIndex        =   130
         Top             =   840
         Width           =   6015
      End
      Begin VB.CheckBox chkUseCreationDateNow 
         Caption         =   "Use the current Date/Time for 'Creation Date'"
         Height          =   255
         Left            =   120
         TabIndex        =   129
         Top             =   1440
         Width           =   5775
      End
   End
   Begin VB.Frame fraProgFont 
      Caption         =   "Programfont"
      Height          =   4935
      Left            =   2640
      TabIndex        =   29
      Top             =   960
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton cmdCancelTest 
         Caption         =   "CancelTest"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1680
         TabIndex        =   109
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test"
         Height          =   495
         Left            =   120
         TabIndex        =   108
         Top             =   4080
         Width           =   1335
      End
      Begin VB.TextBox txtProgramFontsize 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   5400
         TabIndex        =   54
         Text            =   "8"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtTest 
         Height          =   2655
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   69
         Top             =   1320
         Width           =   6015
      End
      Begin VB.ComboBox cmbCharset 
         Height          =   315
         Left            =   3000
         TabIndex        =   53
         Text            =   "cmbCharset"
         Top             =   600
         Width           =   2295
      End
      Begin VB.ComboBox cmbFonts 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown-Liste
         TabIndex        =   49
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label lblSize 
         Caption         =   "Size"
         Height          =   255
         Left            =   5400
         TabIndex        =   43
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblTesttext 
         Caption         =   "Here you can test the font."
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   1080
         Width           =   5295
      End
      Begin VB.Label lblProgcharset 
         Caption         =   "Charset"
         Height          =   255
         Left            =   3000
         TabIndex        =   42
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblProgfont 
         Caption         =   "Programfont"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame fraProgDirectories 
      Caption         =   "Directories"
      Height          =   3015
      Left            =   2640
      TabIndex        =   16
      Top             =   1920
      Visible         =   0   'False
      Width           =   6255
      Begin VB.TextBox txtTemppath 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   86
         Top             =   2520
         Width           =   5415
      End
      Begin VB.CommandButton cmdGetTemppath 
         Caption         =   "..."
         Height          =   255
         Left            =   5760
         TabIndex        =   88
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtGSbin 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   600
         Width           =   5415
      End
      Begin VB.CommandButton cmdGetgsbinDirectory 
         Caption         =   "..."
         Height          =   255
         Left            =   5760
         TabIndex        =   36
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdGetgsfontsDirectory 
         Caption         =   "..."
         Height          =   255
         Left            =   5760
         TabIndex        =   73
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton cmdGetgslibDirectory 
         Caption         =   "..."
         Height          =   255
         Left            =   5760
         TabIndex        =   59
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtGSfonts 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   1800
         Width           =   5415
      End
      Begin VB.TextBox txtGSlib 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   1200
         Width           =   5415
      End
      Begin VB.Label lblPrintTempPath 
         Caption         =   "Temppath"
         Height          =   255
         Left            =   240
         TabIndex        =   84
         Top             =   2280
         Width           =   2895
      End
      Begin VB.Label lblGSbin 
         Caption         =   "Ghostscript Binaries"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label lblGSfonts 
         Caption         =   "Ghostscript Fonts"
         Height          =   255
         Left            =   240
         TabIndex        =   62
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Label lblGSlib 
         Caption         =   "Ghostscript Libraries"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   960
         Width           =   2895
      End
   End
   Begin VB.Frame fraPDFCompress 
      Caption         =   "Compression"
      Height          =   3855
      Left            =   2640
      TabIndex        =   46
      Top             =   960
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Frame fraPDFGrey 
         Caption         =   "Greyscale Images"
         Height          =   975
         Left            =   120
         TabIndex        =   89
         Top             =   1680
         Width           =   5535
         Begin VB.TextBox txtPDFGreyRes 
            Height          =   285
            Left            =   4440
            TabIndex        =   96
            Top             =   540
            Width           =   735
         End
         Begin VB.ComboBox cmbPDFGreyResample 
            Height          =   315
            ItemData        =   "frmOptions.frx":0B66
            Left            =   2280
            List            =   "frmOptions.frx":0B68
            Style           =   2  'Dropdown-Liste
            TabIndex        =   95
            Tag             =   "Bicubic|Subsample|Average"
            Top             =   540
            Width           =   1935
         End
         Begin VB.CheckBox chkPDFGreyResample 
            Caption         =   "Resample"
            Height          =   255
            Left            =   2280
            TabIndex        =   92
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox cmbPDFGreyComp 
            Height          =   315
            ItemData        =   "frmOptions.frx":0B6A
            Left            =   120
            List            =   "frmOptions.frx":0B6C
            Style           =   2  'Dropdown-Liste
            TabIndex        =   94
            Top             =   540
            Width           =   1935
         End
         Begin VB.CheckBox chkPDFGreyComp 
            Caption         =   "Compress"
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblPDFGreyRes 
            Caption         =   "Resolution"
            Height          =   255
            Left            =   4440
            TabIndex        =   93
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame fraPDFMono 
         Caption         =   "Monochrome Images"
         Height          =   975
         Left            =   120
         TabIndex        =   98
         Top             =   2760
         Width           =   5535
         Begin VB.CheckBox chkPDFMonoComp 
            Caption         =   "Compress"
            Height          =   255
            Left            =   120
            TabIndex        =   99
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox cmbPDFMonoComp 
            Height          =   315
            ItemData        =   "frmOptions.frx":0B6E
            Left            =   120
            List            =   "frmOptions.frx":0B70
            Style           =   2  'Dropdown-Liste
            TabIndex        =   105
            Top             =   540
            Width           =   1935
         End
         Begin VB.CheckBox chkPDFMonoResample 
            Caption         =   "Resample"
            Height          =   255
            Left            =   2280
            TabIndex        =   100
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox cmbPDFMonoResample 
            Height          =   315
            ItemData        =   "frmOptions.frx":0B72
            Left            =   2280
            List            =   "frmOptions.frx":0B74
            Style           =   2  'Dropdown-Liste
            TabIndex        =   106
            Tag             =   "Bicubic|Subsample|Average"
            Top             =   540
            Width           =   1935
         End
         Begin VB.TextBox txtPDFMonoRes 
            Height          =   285
            Left            =   4440
            TabIndex        =   107
            Top             =   540
            Width           =   735
         End
         Begin VB.Label lblPDFMonoRes 
            Caption         =   "Resolution"
            Height          =   255
            Left            =   4440
            TabIndex        =   101
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.CheckBox chkPDFTextComp 
         Caption         =   "Compress Text Objects"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   240
         Width           =   4335
      End
      Begin VB.Frame fraPDFColor 
         Caption         =   "Color Images"
         Height          =   975
         Left            =   120
         TabIndex        =   64
         Top             =   600
         Width           =   5535
         Begin VB.TextBox txtPDFColorRes 
            Height          =   285
            Left            =   4440
            TabIndex        =   80
            Top             =   540
            Width           =   735
         End
         Begin VB.ComboBox cmbPDFColorResample 
            Height          =   315
            ItemData        =   "frmOptions.frx":0B76
            Left            =   2280
            List            =   "frmOptions.frx":0B78
            Style           =   2  'Dropdown-Liste
            TabIndex        =   79
            Tag             =   "Bicubic|Subsample|Average"
            Top             =   540
            Width           =   1935
         End
         Begin VB.CheckBox chkPDFColorResample 
            Caption         =   "Resample"
            Height          =   255
            Left            =   2280
            TabIndex        =   71
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox cmbPDFColorComp 
            Height          =   315
            ItemData        =   "frmOptions.frx":0B7A
            Left            =   120
            List            =   "frmOptions.frx":0B7C
            Style           =   2  'Dropdown-Liste
            TabIndex        =   78
            Top             =   540
            Width           =   1935
         End
         Begin VB.CheckBox chkPDFColorComp 
            Caption         =   "Compress"
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblPDFColorRes 
            Caption         =   "Resolution"
            Height          =   255
            Left            =   4440
            TabIndex        =   72
            Top             =   240
            Width           =   975
         End
      End
   End
   Begin VB.Frame fraPSGeneral 
      Caption         =   "Postscript"
      Height          =   1095
      Left            =   2640
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   6255
      Begin VB.ComboBox cmbEPSLanguageLevel 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown-Liste
         TabIndex        =   13
         Top             =   600
         Width           =   735
      End
      Begin VB.ComboBox cmbPSLanguageLevel 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown-Liste
         TabIndex        =   10
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblLangLevel 
         Alignment       =   1  'Rechts
         Caption         =   "Language Level:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame fraBitmapGeneral 
      Caption         =   "Bitmap"
      Height          =   1935
      Left            =   2760
      TabIndex        =   4
      Top             =   960
      Width           =   6255
      Begin VB.ComboBox cmbTIFFColors 
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown-Liste
         TabIndex        =   32
         Top             =   1440
         Width           =   2175
      End
      Begin VB.ComboBox cmbPCXColors 
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown-Liste
         TabIndex        =   27
         Top             =   1320
         Width           =   2175
      End
      Begin VB.ComboBox cmbBMPColors 
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown-Liste
         TabIndex        =   25
         Top             =   1200
         Width           =   2175
      End
      Begin VB.ComboBox cmbJPEGColors 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown-Liste
         TabIndex        =   22
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtJPEGQuality 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   1920
         TabIndex        =   30
         Text            =   "75"
         Top             =   1440
         Width           =   495
      End
      Begin VB.ComboBox cmbPNGColors 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown-Liste
         TabIndex        =   19
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox txtBitmapResolution 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   1920
         TabIndex        =   11
         Text            =   "72"
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblJPEQQualityProzent 
         Caption         =   "%"
         Height          =   255
         Left            =   2520
         TabIndex        =   31
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label lblJPEGQuality 
         Alignment       =   1  'Rechts
         Caption         =   "Quality:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblBitmapColors 
         Alignment       =   1  'Rechts
         Caption         =   "Colors:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblBitmapDPI 
         Caption         =   "dpi"
         Height          =   255
         Left            =   2520
         TabIndex        =   12
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblBitmapResolution 
         Alignment       =   1  'Rechts
         Caption         =   "Resolution:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame fraPDFFonts 
      Caption         =   "Font Options"
      Height          =   2895
      Left            =   3600
      TabIndex        =   39
      Top             =   2640
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CheckBox chkPDFEmbedAll 
         Caption         =   "Embed all Fonts"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   360
         Width           =   5175
      End
      Begin VB.CheckBox chkPDFSubSetFonts 
         Caption         =   "Subset Fonts, when percentage of used characters below:"
         Height          =   495
         Left            =   120
         TabIndex        =   61
         Top             =   780
         Width           =   5535
      End
      Begin VB.TextBox txtPDFSubSetPerc 
         Height          =   285
         Left            =   360
         TabIndex        =   76
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblPDFPerc 
         Caption         =   "%"
         Height          =   255
         Left            =   960
         TabIndex        =   77
         Top             =   1320
         Width           =   615
      End
   End
   Begin VB.Frame fraPDFColors 
      Caption         =   "Color Options"
      Height          =   3495
      Left            =   3360
      TabIndex        =   24
      Top             =   2160
      Visible         =   0   'False
      Width           =   5775
      Begin VB.ComboBox cmbPDFColorModel 
         Height          =   315
         ItemData        =   "frmOptions.frx":0B7E
         Left            =   120
         List            =   "frmOptions.frx":0B80
         Style           =   2  'Dropdown-Liste
         TabIndex        =   34
         Tag             =   "RGB|CMYK|GRAY"
         Top             =   360
         Width           =   2775
      End
      Begin VB.Frame fraPDFColorOptions 
         Caption         =   "Options"
         Height          =   1455
         Left            =   120
         TabIndex        =   81
         Top             =   1920
         Width           =   5535
         Begin VB.CheckBox chkPDFPreserveHalftone 
            Caption         =   "Preserve Halftone Information"
            Height          =   255
            Left            =   120
            TabIndex        =   97
            Top             =   1080
            Width           =   5175
         End
         Begin VB.CheckBox chkPDFPreserveTransfer 
            Caption         =   "Preserve Transfer Functions"
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Tag             =   "Remove|Preserve"
            Top             =   720
            Width           =   5175
         End
         Begin VB.CheckBox chkPDFPreserveOverprint 
            Caption         =   "Preserve Overprint Settings"
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   360
            Width           =   5175
         End
      End
      Begin VB.CheckBox chkPDFCMYKtoRGB 
         Caption         =   "Convert CMYK Images to RGB"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   840
         Width           =   3255
      End
   End
   Begin VB.Frame fraPDFGeneral 
      Caption         =   "General Options"
      Height          =   2895
      Left            =   3120
      TabIndex        =   14
      Top             =   1680
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CheckBox chkPDFASCII85 
         Caption         =   "Convert binary data to ASCII85"
         Height          =   255
         Left            =   2400
         TabIndex        =   82
         Top             =   2400
         Width           =   3255
      End
      Begin VB.ComboBox cmbPDFOverprint 
         Height          =   315
         ItemData        =   "frmOptions.frx":0B82
         Left            =   2400
         List            =   "frmOptions.frx":0B84
         Style           =   2  'Dropdown-Liste
         TabIndex        =   65
         Top             =   1860
         Width           =   2655
      End
      Begin VB.TextBox txtPDFRes 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   2400
         TabIndex        =   51
         Text            =   "600"
         Top             =   1320
         Width           =   615
      End
      Begin VB.ComboBox cmbPDFCompat 
         Height          =   315
         ItemData        =   "frmOptions.frx":0B86
         Left            =   2400
         List            =   "frmOptions.frx":0B88
         Style           =   2  'Dropdown-Liste
         TabIndex        =   21
         Top             =   360
         Width           =   2655
      End
      Begin VB.ComboBox cmbPDFRotate 
         Height          =   315
         ItemData        =   "frmOptions.frx":0B8A
         Left            =   2400
         List            =   "frmOptions.frx":0B8C
         Style           =   2  'Dropdown-Liste
         TabIndex        =   35
         Tag             =   "None|All|PageByPage"
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label lblPDFDPI 
         BackStyle       =   0  'Transparent
         Caption         =   "dpi"
         Height          =   255
         Left            =   3120
         TabIndex        =   52
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblPDFOverprint 
         Alignment       =   1  'Rechts
         Caption         =   "Overprint:"
         Height          =   375
         Left            =   120
         TabIndex        =   66
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label lblPDFResolution 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Resolution:"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   1380
         Width           =   2175
      End
      Begin VB.Label lblPDFCompat 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Compatibility:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   420
         Width           =   2175
      End
      Begin VB.Label lblPDFAutoRotate 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Auto-Rotate Pages:"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   900
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   7320
      TabIndex        =   104
      Top             =   6000
      Width           =   1575
   End
   Begin MSComctlLib.TreeView trv 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
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
   Begin VB.PictureBox picOptions 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2640
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   495
      Left            =   4920
      TabIndex        =   103
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2640
      TabIndex        =   102
      Top             =   6000
      Width           =   1575
   End
   Begin MSComctlLib.TabStrip tbstrPDFOptions 
      Height          =   4455
      Left            =   7080
      TabIndex        =   83
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
   Begin VB.Frame fraPDFSecurity 
      Caption         =   "Security"
      Height          =   4335
      Left            =   2760
      TabIndex        =   18
      Top             =   1200
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Frame fraPDFHighPermissions 
         Caption         =   "Enhanced Permissions (128 Bit only)"
         Height          =   855
         Left            =   120
         TabIndex        =   122
         Top             =   3360
         Width           =   5535
         Begin VB.CheckBox chkAllowAssembly 
            Caption         =   "Allow changes to the Assembly"
            Height          =   255
            Left            =   2760
            TabIndex        =   126
            Top             =   480
            Width           =   2535
         End
         Begin VB.CheckBox chkAllowScreenReaders 
            Caption         =   "Allow Screen Readers"
            Height          =   255
            Left            =   240
            TabIndex        =   125
            Top             =   480
            Width           =   2535
         End
         Begin VB.CheckBox chkAllowFillIn 
            Caption         =   "Allow filling in form fields"
            Height          =   255
            Left            =   2760
            TabIndex        =   124
            Top             =   240
            Width           =   2415
         End
         Begin VB.CheckBox chkAllowDegradedPrinting 
            Caption         =   "Allow printing in low resolution"
            Height          =   255
            Left            =   240
            TabIndex        =   123
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame fraPDFPermissions 
         Caption         =   "Disallow User to"
         Height          =   855
         Left            =   120
         TabIndex        =   117
         Top             =   2400
         Width           =   5535
         Begin VB.CheckBox chkAllowModifyAnnotations 
            Caption         =   "modify comments"
            Height          =   255
            Left            =   2760
            TabIndex        =   121
            Top             =   480
            Width           =   2295
         End
         Begin VB.CheckBox chkAllowModifyContents 
            Caption         =   "modify the document"
            Height          =   255
            Left            =   2760
            TabIndex        =   120
            Top             =   240
            Width           =   2535
         End
         Begin VB.CheckBox chkAllowCopy 
            Caption         =   "copy text and images"
            Height          =   255
            Left            =   240
            TabIndex        =   119
            Top             =   480
            Width           =   2535
         End
         Begin VB.CheckBox chkAllowPrinting 
            Caption         =   "print the document"
            Height          =   255
            Left            =   240
            TabIndex        =   118
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame fraSecurityPass 
         Caption         =   "Passwords"
         Height          =   855
         Left            =   120
         TabIndex        =   114
         Top             =   1440
         Width           =   5535
         Begin VB.CheckBox chkOwnerPass 
            Caption         =   "Password required to change Permissions and Passwords"
            Height          =   255
            Left            =   240
            TabIndex        =   116
            Top             =   480
            Width           =   5175
         End
         Begin VB.CheckBox chkUserPass 
            Caption         =   "Password required to open document"
            Height          =   255
            Left            =   240
            TabIndex        =   115
            Top             =   240
            Width           =   5175
         End
      End
      Begin VB.Frame fraPDFEncLevel 
         Caption         =   "Encryption Level"
         Height          =   855
         Left            =   120
         TabIndex        =   111
         Top             =   480
         Width           =   5535
         Begin VB.OptionButton optEncHigh 
            Caption         =   "High (128 Bit - Adobe Acrobat 5.0 and above)"
            Height          =   255
            Left            =   240
            TabIndex        =   113
            Top             =   480
            Width           =   4215
         End
         Begin VB.OptionButton optEncLow 
            Caption         =   "Low (40 Bit - Adobe Acrobat 3.0 and above)"
            Height          =   255
            Left            =   240
            TabIndex        =   112
            Top             =   240
            Width           =   4215
         End
      End
      Begin VB.CheckBox chkUseSecurity 
         Caption         =   "Use Security"
         Height          =   255
         Left            =   120
         TabIndex        =   110
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.Label lblOptions 
      Height          =   615
      Left            =   3360
      TabIndex        =   2
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
50040   Else
50050    txtStandardAuthor.Enabled = False
50060    txtStandardAuthor.BackColor = &H8000000F
50070  End If
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
50050  End With
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
50010  Select Case Index
  Case 0: ' Add
50030    AddFilenameSubstitutions
50040   Case 1: ' Change
50050    ChangeFilenameSubstitutions
50060   Case 2: ' Delete
50070    DeleteFilenameSubstitutions
50080   Case 3: ' Up
50090    MoveUpFilenameSubstitutions
50100   Case 4: ' Down
50110    MoveDownFilenameSubstitutions
50120  End Select
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

Private Sub cmdGetAutosaveDirectory_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 Dim strFolder As String
50020
50030 strFolder = BrowseForFolder(Me.hWnd, LanguageStrings.OptionsAutosaveDirectoryPrompt)
50040 If Len(strFolder) = 0 Then Exit Sub
50050 If Right$(strFolder, 1) <> "\" Then
50060  strFolder = strFolder & "\"
50070 End If
50080 txtAutosaveDirectory.Text = strFolder
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
50010  Dim strFolder As String
50020
50030  strFolder = BrowseForFolder(Me.hWnd, LanguageStrings.OptionsGhostscriptBinariesDirectoryPrompt)
50040  If Len(strFolder) = 0 Then Exit Sub
50050  strFolder = CompletePath(strFolder)
50060  txtGSbin.Text = strFolder
50070  With txtGSbin
50080   .ToolTipText = .Text
50090  End With
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
50020
50030  strFolder = BrowseForFolder(Me.hWnd, LanguageStrings.OptionsGhostscriptFontsDirectoryPrompt)
50040  If Len(strFolder) = 0 Then Exit Sub
50050  strFolder = CompletePath(strFolder)
50060  txtGSfonts.Text = strFolder
50070  With txtGSfonts
50080   .ToolTipText = .Text
50090  End With
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
50020
50030  strFolder = BrowseForFolder(Me.hWnd, LanguageStrings.OptionsGhostscriptLibrariesDirectoryPrompt)
50040  If Len(strFolder) = 0 Then Exit Sub
50050  strFolder = CompletePath(strFolder)
50060  txtGSlib.Text = strFolder
50070  With txtGSlib
50080   .ToolTipText = .Text
50090  End With
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

Private Sub cmdGetTemppath_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim strFolder As String
50020
50030  strFolder = BrowseForFolder(Me.hWnd, LanguageStrings.OptionsPrintertempDirectoryPrompt)
50040  If Len(strFolder) = 0 Then Exit Sub
50050  strFolder = CompletePath(strFolder)
50060  txtTemppath.Text = strFolder
50070  With txtTemppath
50080   .ToolTipText = .Text
50090  End With
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
50100   End With
50110  End If
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
50010  GetOptions Me, Options
50020  SaveOptions Options
50030  If IsWin9xMe = False Then
50040   Select Case Options.ProcessPriority
   Case 0: 'Idle
50060     SetProcessPriority Idle
50070    Case 1: 'Normal
50080     SetProcessPriority Normal
50090    Case 2: 'High
50100     SetProcessPriority High
50110    Case 3: 'Realtime
50120     SetProcessPriority RealTime
50130   End Select
50140  End If
50150
50160  GsDllLoaded = LoadDLL(Options.DirectoryGhostscriptBinaries & "\gsdll32.dll")
50170
50180  Unload Me
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

Private Sub cmdTest_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tCharset As Long
50020  tCharset = cmbCharset.Text
50030  SetFont Me, cmbFonts.List(cmbFonts.ListIndex), CLng(cmbCharset.Text), txtProgramFontsize.Text
50040  cmbCharset.Text = tCharset
50050  SetFont frmMain, cmbFonts.List(cmbFonts.ListIndex), CLng(cmbCharset.Text), txtProgramFontsize.Text
50060  cmdCancelTest.Enabled = True
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
50010  Dim TestPSPage As String, fn As Long, FileName As String
50020  TestPSPage = LoadResString(3000)
50030  TestPSPage = Replace(TestPSPage, "[TESTPAGE]", LanguageStrings.OptionsTestpage, , 1, vbTextCompare)
50040  TestPSPage = Replace(TestPSPage, "[DATE]", Now, , 1, vbTextCompare)
50050
50060  fn = FreeFile
50070  FileName = GetTempFile(GetPDFCreatorTempfolder, "~PD")
50080  Open FileName For Output As fn
50090  Print #fn, TestPSPage
50100  Close #fn
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
 fraProgSave.Top = fraPDFTop - 300
 fraProgSave.Left = fraPDFLeft - 200
 fraProgDirectories.Top = fraPDFTop - 300
 fraProgDirectories.Left = fraPDFLeft - 200
 fraProgDocument.Top = fraPDFTop - 300
 fraProgDocument.Left = fraPDFLeft - 200

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
 fraPSGeneral.Left = fraPDFLeft
 fraPSGeneral.Top = fraPDFTop - 300
 fraPSGeneral.Left = fraPDFLeft - 200
 tbstrPDFOptions.Top = fraPDFTop - 400
 tbstrPDFOptions.Left = fraPDFLeft - 100
 tbstrPDFOptions.Height = 4875

 cmbEPSLanguageLevel.Top = cmbPSLanguageLevel.Top
 cmbEPSLanguageLevel.Left = cmbPSLanguageLevel.Left

 txtTest.Text = vbNullString
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
  trv.Nodes.Add "Program", tvwChild, "ProgramDocument", .OptionsProgramDocumentSymbol
  trv.Nodes.Add "Program", tvwChild, "ProgramSave", .OptionsProgramSaveSymbol
  trv.Nodes.Add "Program", tvwChild, "ProgramAutosave", .OptionsProgramAutosaveSymbol
  trv.Nodes.Add "Program", tvwChild, "ProgramDirectories", .OptionsProgramDirectoriesSymbol
  trv.Nodes.Add "Program", tvwChild, "ProgramFonts", .OptionsProgramFontSymbol
  trv.Nodes.Add , , "Formats", .OptionsTreeFormats
  trv.Nodes.Add "Formats", tvwChild, "FormatsPDF", .OptionsPDFSymbol
  trv.Nodes.Add "Formats", tvwChild, "FormatsPNG", .OptionsPNGSymbol
  trv.Nodes.Add "Formats", tvwChild, "FormatsJPEG", .OptionsJPEGSymbol
  trv.Nodes.Add "Formats", tvwChild, "FormatsBMP", .OptionsBMPSymbol
  trv.Nodes.Add "Formats", tvwChild, "FormatsPCX", .OptionsPCXSymbol
  trv.Nodes.Add "Formats", tvwChild, "FormatsTIFF", .OptionsTIFFSymbol
  trv.Nodes.Add "Formats", tvwChild, "FormatsPS", .OptionsPSSymbol
  trv.Nodes.Add "Formats", tvwChild, "FormatsEPS", .OptionsEPSSymbol

  trv.Nodes("ProgramFonts").EnsureVisible
  trv.Nodes("FormatsPDF").EnsureVisible

  Set picOptions = LoadResPicture(2101, vbResIcon)
  fraProgFont.Visible = False
  fraProgGeneral.Visible = True

  fraProgGeneral.Caption = .OptionsProgramGeneralSymbol
  fraProgAutosave = .OptionsProgramAutosaveSymbol
  fraProgFont = .OptionsProgramFontSymbol
  fraProgDirectories = .OptionsProgramDirectoriesSymbol
  fraProgSave = .OptionsProgramSaveSymbol
  fraProgDocument = .OptionsProgramDocumentSymbol

  lblSaveFilename.Caption = .OptionsSaveFilename
  lblSaveFilenameTokens.Caption = .OptionsSaveFilenameTokens
  fraFilenameSubstitutions.Caption = .OptionsSaveFilenameSubstitutions
  chkFilenameSubst.Caption = .OptionsSaveFilenameSubstitutionsTitle
  cmdFilenameSubst(0).Caption = .OptionsSaveFilenameAdd
  cmdFilenameSubst(1).Caption = .OptionsSaveFilenameChange
  cmdFilenameSubst(2).Caption = .OptionsSaveFilenameDelete

  chkSpaces.Caption = .OptionsRemoveSpaces

  lblGSlib.Caption = .OptionsDirectoriesGSLibraries
  lblGSfonts.Caption = .OptionsDirectoriesGSFonts

  lblOptions = .OptionsProgramGeneralDescription
  lblAutosaveformat.Caption = .OptionsAutosaveFormat
  chkUseStandardAuthor.Caption = .OptionsUseStandardauthor
  chkUseCreationDateNow.Caption = .OptionsUseCreationDateNow
  chkUseAutosave.Caption = .OptionsUseAutosave
  cmdTestpage.Caption = .OptionsPrintTestpage
  lblAutosaveFilename.Caption = .OptionsAutosaveFilename
  lblAutosaveFilenameTokens.Caption = .OptionsAutosaveFilenameTokens
  chkUseAutosaveDirectory.Caption = .OptionsUseAutosaveDirectory
  With cmbAutoSaveFilenameTokens
   .AddItem "<Author>"
   .AddItem "<Computername>"
   .AddItem "<DateTime>"
   .AddItem "<Title>"
   .AddItem "<Username>"
   .ListIndex = 0
  End With
  With cmbSaveFilenameTokens
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
   .AddItem "PS"
   .AddItem "EPS"
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
'  tbstrPDFOptions.Tabs.Add , "Security", .OptionsPDFSecurity
  fraPDFGeneral.Caption = .OptionsPDFGeneralCaption
  lblPDFCompat.Caption = .OptionsPDFGeneralCompatibility
  lblPDFAutoRotate.Caption = .OptionsPDFGeneralAutorotate
  lblPDFResolution.Caption = .OptionsPDFGeneralResolution
  lblPDFOverprint.Caption = .OptionsPDFGeneralOverprint
  lblProgfont.Caption = .OptionsProgramFont
  lblProgcharset.Caption = .OptionsProgramFontcharset
  lblSize.Caption = .OptionsProgramFontSize
  lblTesttext = .OptionsProgramFontTestdescription
  cmdTest.Caption = .OptionsProgramFontTest
  cmdCancelTest.Caption = .OptionsProgramFontCancelTest
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
  chkUseSecurity.Caption = .OptionsPDFUseSecurity
  fraPDFEncLevel.Caption = .OptionsPDFEncryptionLevel
  optEncHigh.Caption = .OptionsPDFEncryptionHigh
  optEncLow.Caption = .OptionsPDFEncryptionLow
  fraSecurityPass.Caption = .OptionsPDFPasswords
  chkUserPass.Caption = .OptionsPDFUserPass
  chkOwnerPass.Caption = .OptionsPDFOwnerPass
  fraPDFPermissions.Caption = .OptionsPDFDisallowUser
  chkAllowPrinting.Caption = .OptionsPDFDisallowPrint
  chkAllowModifyContents.Caption = .OptionsPDFDisallowModify
  chkAllowCopy.Caption = .OptionsPDFDisallowCopy
  chkAllowModifyAnnotations.Caption = .OptionsPDFDisallowModifyComments
  chkAllowDegradedPrinting.Caption = .OptionsPDFAllowDegradedPrinting
  chkAllowFillIn.Caption = .OptionsPDFAllowFillIn
  chkAllowAssembly.Caption = .OptionsPDFAllowAssembly
  chkAllowScreenReaders.Caption = .OptionsPDFAllowScreenReaders

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

  fraBitmapGeneral.Caption = .OptionsImageSettings
  lblBitmapResolution = .OptionsBitmapResolution
  lblJPEGQuality = .OptionsJPEGQuality
  lblBitmapColors = .OptionsPDFColors
  lblProcessPriority.Caption = .OptionsProcesspriority
  lblLangLevel.Caption = .OptionsPSLanguageLevel

  cmdAsso.Caption = .OptionsAssociatePSFiles
 End With

 If IsPsAssociate = False Then
   cmdAsso.Enabled = True
  Else
   cmdAsso.Enabled = False
 End If

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
 txtStandardAuthor.Text = vbNullString

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

 With lsvFilenameSubst
  .Appearance = ccFlat
  .ColumnHeaders.Clear
  .ColumnHeaders.Add , "Str1", "", lsvFilenameSubst.Width / 2 - 140
  .ColumnHeaders.Add , "Str2", "", lsvFilenameSubst.Width / 2 - 140
  .HideColumnHeaders = True
  .GridLines = True
  .FullRowSelect = True
  .HideSelection = False
 End With
 
 ShowOptions Me, Options

 CheckCmdFilenameSubst

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

 With txtGSbin
  .ToolTipText = .Text
 End With
 With txtGSlib
  .ToolTipText = .Text
 End With
 With txtGSfonts
  .ToolTipText = .Text
 End With
 With txtTemppath
  .ToolTipText = .Text
 End With

 With sldProcessPriority
  .TextPosition = sldBelowRight
  .TickFrequency = 1
  .TickStyle = sldTopLeft
  Select Case .Value
   Case 0: 'Idle
    lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityIdle
   Case 1: 'Normal
    lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityNormal
   Case 2: 'High
    lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityHigh
   Case 3: 'Realtime
    lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityRealtime
  End Select
 End With
 If IsWin9xMe = False Then
   lblProcessPriority.Enabled = True
   sldProcessPriority.Enabled = True
  Else
   lblProcessPriority.Enabled = False
   sldProcessPriority.Enabled = False
 End If
 UpdateSecurityFields
 Screen.MousePointer = vbNormal
End Sub

Private Sub ShowProgOptions()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  lblJPEGQuality.Visible = False
50020  cmbPNGColors.Visible = False
50030  cmbJPEGColors.Visible = False
50040  cmbBMPColors.Visible = False
50050  cmbPCXColors.Visible = False
50060  cmbTIFFColors.Visible = False
50070  tbstrPDFOptions.Visible = False
50080  fraProgFont.Visible = False
50090  fraProgGeneral.Visible = False
50100  fraProgAutosave.Visible = False
50110  fraProgSave.Visible = False
50120  fraProgDirectories.Visible = False
50130  fraProgDocument.Visible = False
50140  fraBitmapGeneral.Visible = False
50150  fraPDFGeneral.Visible = False
50160  fraPDFCompress.Visible = False
50170  fraPDFFonts.Visible = False
50180  fraPDFColors.Visible = False
50190  fraPDFSecurity.Visible = False
50200  txtJPEGQuality.Visible = False
50210  lblJPEQQualityProzent.Visible = False
50220  fraPSGeneral.Visible = False
50230  cmbPSLanguageLevel.Visible = False
50240  cmbEPSLanguageLevel.Visible = False
50250
50260  Select Case trv.SelectedItem.Key
  Case "Program"
50280    Set picOptions = LoadResPicture(2101, vbResIcon)
50290    lblOptions = LanguageStrings.OptionsProgramGeneralDescription
50300    fraProgGeneral.Visible = True
50310   Case "ProgramGeneral"
50320    Set picOptions = LoadResPicture(2101, vbResIcon)
50330    lblOptions = LanguageStrings.OptionsProgramGeneralDescription
50340    fraProgGeneral.Visible = True
50350   Case "ProgramSave"
50360    Set picOptions = LoadResPicture(2106, vbResIcon)
50370    lblOptions = LanguageStrings.OptionsProgramSaveDescription
50380    fraProgSave.Visible = True
50390   Case "ProgramAutosave"
50400    Set picOptions = LoadResPicture(2103, vbResIcon)
50410    lblOptions = LanguageStrings.OptionsProgramAutosaveDescription
50420    fraProgAutosave.Visible = True
50430   Case "ProgramFonts"
50440    Set picOptions = LoadResPicture(2102, vbResIcon)
50450    lblOptions = LanguageStrings.OptionsProgramFontDescription
50460    fraProgFont.Visible = True
50470   Case "ProgramDirectories"
50480    Set picOptions = LoadResPicture(2104, vbResIcon)
50490    lblOptions = LanguageStrings.OptionsProgramDirectoriesDescription
50500    fraProgDirectories.Visible = True
50510   Case "ProgramDocument"
50520    Set picOptions = LoadResPicture(2105, vbResIcon)
50530    lblOptions = LanguageStrings.OptionsProgramDocumentDescription
50540    fraProgDocument.Visible = True
50550   Case "Formats"
50560    Set picOptions = LoadResPicture(2111, vbResIcon)
50570    lblOptions = LanguageStrings.OptionsPDFDescription
50580    tbstrPDFOptions.Visible = True
50590    fraPDFGeneral.Visible = True
50600   Case "FormatsPDF"
50610    Set picOptions = LoadResPicture(2111, vbResIcon)
50620    lblOptions = LanguageStrings.OptionsPDFDescription
50630    tbstrPDFOptions.Visible = True
50640    fraPDFGeneral.Visible = True
50650   Case "FormatsPNG"
50660    Set picOptions = LoadResPicture(2112, vbResIcon)
50670    lblOptions = LanguageStrings.OptionsPNGDescription
50680    fraBitmapGeneral.Visible = True
50690    cmbPNGColors.Visible = True
50700   Case "FormatsJPEG"
50710    Set picOptions = LoadResPicture(2113, vbResIcon)
50720    lblOptions = LanguageStrings.OptionsJPEGDescription
50730    fraBitmapGeneral.Visible = True
50740    lblJPEGQuality.Caption = LanguageStrings.OptionsJPEGQuality
50750    lblJPEGQuality.Visible = True
50760    txtJPEGQuality.Visible = True
50770    lblJPEQQualityProzent.Visible = True
50780    lblJPEQQualityProzent.Left = txtJPEGQuality.Left + txtJPEGQuality.Width + 100
50790    cmbJPEGColors.Visible = True
50800   Case "FormatsBMP"
50810    Set picOptions = LoadResPicture(2114, vbResIcon)
50820    lblOptions = LanguageStrings.OptionsBMPDescription
50830    fraBitmapGeneral.Visible = True
50840    cmbBMPColors.Visible = True
50850   Case "FormatsPCX"
50860    Set picOptions = LoadResPicture(2115, vbResIcon)
50870    lblOptions = LanguageStrings.OptionsPCXDescription
50880    fraBitmapGeneral.Visible = True
50890    cmbPCXColors.Visible = True
50900   Case "FormatsTIFF"
50910    Set picOptions = LoadResPicture(2116, vbResIcon)
50920    lblOptions = LanguageStrings.OptionsTIFFDescription
50930    fraBitmapGeneral.Visible = True
50940    cmbTIFFColors.Visible = True
50950   Case "FormatsPS"
50960    Set picOptions = LoadResPicture(2117, vbResIcon)
50970    lblOptions.Caption = LanguageStrings.OptionsPSDescription
50980    fraPSGeneral.Visible = True
50990    cmbPSLanguageLevel.Visible = True
51000    fraPSGeneral.Caption = LanguageStrings.OptionsPSDescription
51010   Case "FormatsEPS"
51020    Set picOptions = LoadResPicture(2118, vbResIcon)
51030    lblOptions.Caption = LanguageStrings.OptionsEPSDescription
51040    fraPSGeneral.Visible = True
51050    cmbEPSLanguageLevel.Visible = True
51060    fraPSGeneral.Caption = LanguageStrings.OptionsEPSDescription
51070  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "ShowProgOptions")
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
50020   Select Case .Value
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
50010  fraPDFGeneral.Visible = False
50020  fraPDFCompress.Visible = False
50030  fraPDFFonts.Visible = False
50040  fraPDFColors.Visible = False
50050  fraPDFSecurity.Visible = False
50060  Select Case tbstrPDFOptions.SelectedItem.Index
  Case 1:
50080    fraPDFGeneral.Visible = True
50090   Case 2:
50100    fraPDFCompress.Visible = True
50110   Case 3:
50120    fraPDFFonts.Visible = True
50130   Case 4:
50140    fraPDFColors.Visible = True
50150   Case 5:
50160    fraPDFSecurity.Visible = True
50170  End Select
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

Private Sub trv_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ShowProgOptions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "trv_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub trv_NodeClick(ByVal Node As MSComctlLib.Node)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ShowProgOptions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "trv_NodeClick")
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
50010  txtAutosaveFilename.ToolTipText = txtAutosaveFilename.Text
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

Private Sub txtProgramFontSize_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tL As Long
50020 If Trim$(txtProgramFontsize.Text) = "" Then
50030   txtProgramFontsize.Text = 8
50040  End If
50050  tL = CLng(txtProgramFontsize.Text)
50060  If tL <= 0 Then
50070   tL = 1
50080  End If
50090  If tL > 72 Then
50100   tL = 72
50110  End If
50120  txtProgramFontsize.Text = tL
50130  txtTest.Font.Size = tL
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "txtProgramFontSize_Change")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtProgramFontSize_KeyPress(KeyAscii As Integer)
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
Select Case ErrPtnr.OnError("frmOptions", "txtProgramFontSize_KeyPress")
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
50050  lblAutosaveFilenameTokens.Enabled = ViewIt
50060  cmbAutoSaveFilenameTokens.Enabled = ViewIt
50070  chkUseAutosaveDirectory.Enabled = ViewIt
50080  If ViewIt = True Then
50090    cmbAutosaveFormat.BackColor = &H80000005
50100    cmbAutoSaveFilenameTokens.BackColor = &H80000005
50110    txtAutosaveFilename.BackColor = &H80000005
50120   Else
50130    cmbAutosaveFormat.BackColor = &H8000000F
50140    cmbAutoSaveFilenameTokens.BackColor = &H8000000F
50150    txtAutosaveFilename.BackColor = &H8000000F
50160  End If
50170  If chkUseAutosaveDirectory.Value = 1 And ViewIt = True Then
50180    ViewAutosaveDirectory True
50190   Else
50200    ViewAutosaveDirectory False
50210  End If
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
50010  If ViewIt = True Then
50020    txtAutosaveDirectory.Enabled = True
50030    txtAutosaveDirectory.BackColor = &HC0FFFF
50040    cmdGetAutosaveDirectory.Enabled = True
50050   Else
50060    txtAutosaveDirectory.Enabled = False
50070    txtAutosaveDirectory.BackColor = &H8000000F
50080    cmdGetAutosaveDirectory.Enabled = False
50090  End If
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
50010  If chkUseSecurity.Value = False Then
50020    fraPDFEncLevel.Enabled = False
50030    optEncHigh.Enabled = False
50040    optEncLow.Enabled = False
50050
50060    fraSecurityPass.Enabled = False
50070    chkUserPass.Enabled = False
50080    chkOwnerPass.Enabled = False
50090
50100    fraPDFPermissions.Enabled = False
50110    chkAllowPrinting.Enabled = False
50120    chkAllowCopy.Enabled = False
50130    chkAllowModifyAnnotations.Enabled = False
50140    chkAllowModifyContents.Enabled = False
50150
50160    fraPDFHighPermissions.Enabled = False
50170    chkAllowDegradedPrinting.Enabled = False
50180    chkAllowFillIn.Enabled = False
50190    chkAllowScreenReaders.Enabled = False
50200    chkAllowAssembly.Enabled = False
50210   Else
50220    fraPDFEncLevel.Enabled = True
50230    If cmbPDFCompat.ListIndex >= 2 Then
50240      optEncHigh.Enabled = True
50250     Else
50260      optEncHigh.Enabled = False
50270    End If
50280    optEncLow.Enabled = True
50290
50300    fraSecurityPass.Enabled = True
50310    chkUserPass.Enabled = True
50320    chkOwnerPass.Enabled = True
50330
50340    fraPDFPermissions.Enabled = True
50350    chkAllowPrinting.Enabled = True
50360    chkAllowCopy.Enabled = True
50370    chkAllowModifyAnnotations.Enabled = True
50380    chkAllowModifyContents.Enabled = True
50390
50400    If optEncHigh.Value = True Then
50410      fraPDFHighPermissions.Enabled = True
50420      chkAllowDegradedPrinting.Enabled = True
50430      chkAllowFillIn.Enabled = True
50440      chkAllowScreenReaders.Enabled = True
50450      chkAllowAssembly.Enabled = True
50460     Else
50470      fraPDFHighPermissions.Enabled = False
50480      chkAllowDegradedPrinting.Enabled = False
50490      chkAllowFillIn.Enabled = False
50500      chkAllowScreenReaders.Enabled = False
50510      chkAllowAssembly.Enabled = False
50520    End If
50530  End If
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
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim i As Long, res As Long
50050  res = CheckFilenameSubstitutions(0)
50060  Select Case res
  Case 0:
50080    lsvFilenameSubst.ListItems.Add , , txtFilenameSubst(0).Text
50090    lsvFilenameSubst.ListItems(lsvFilenameSubst.ListItems.Count).SubItems(1) = txtFilenameSubst(1).Text
50100    lsvFilenameSubst.ListItems(lsvFilenameSubst.ListItems.Count).Selected = True
50110    lsvFilenameSubst.ListItems(lsvFilenameSubst.ListItems.Count).EnsureVisible
50120    Set_txtFilenameSubst
50130   Case 2:
50140    MsgBox LanguageStrings.MessagesMsg12 & _
    vbCrLf & vbTab & "\ / : * ? < > | """
50160   Case 3:
50170    MsgBox LanguageStrings.MessagesMsg11
50180  End Select
50190 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50200 Exit Sub
ErrPtnr_OnError:
50220 Select Case ErrPtnr.OnError("frmOptions", "AddFilenameSubstitutions")
Case 0: Resume
50240 Case 1: Resume Next
50250 Case 2: Exit Sub
50260 Case 3: End
50270 End Select
50280 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ChangeFilenameSubstitutions()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim i As Long, res As Long
50050  res = CheckFilenameSubstitutions(lsvFilenameSubst.SelectedItem.Index)
50060  Select Case res
  Case 0:
50080    lsvFilenameSubst.SelectedItem.Text = txtFilenameSubst(0).Text
50090    lsvFilenameSubst.SelectedItem.SubItems(1) = txtFilenameSubst(1).Text
50100   Case 2:
50110    MsgBox LanguageStrings.MessagesMsg12 & _
    vbCrLf & vbTab & "\ / : * ? < > | """
50130   Case 3:
50140    MsgBox LanguageStrings.MessagesMsg11
50150  End Select
50160 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50170 Exit Sub
ErrPtnr_OnError:
50190 Select Case ErrPtnr.OnError("frmOptions", "ChangeFilenameSubstitutions")
Case 0: Resume
50210 Case 1: Resume Next
50220 Case 2: Exit Sub
50230 Case 3: End
50240 End Select
50250 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub DeleteFilenameSubstitutions()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim oIndex As Long
50050  With lsvFilenameSubst
50060   oIndex = .SelectedItem.Index
50070   If .ListItems.Count > 0 Then
50080    .ListItems.Remove .SelectedItem.Index
50090    If oIndex > .ListItems.Count Then
50100     oIndex = .ListItems.Count
50110    End If
50120    If .ListItems.Count > 0 Then
50130     .ListItems(oIndex).Selected = True
50140     .ListItems(oIndex).EnsureVisible
50150    End If
50160    Set_txtFilenameSubst
50170   End If
50180  End With
50190 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50200 Exit Sub
ErrPtnr_OnError:
50220 Select Case ErrPtnr.OnError("frmOptions", "DeleteFilenameSubstitutions")
Case 0: Resume
50240 Case 1: Resume Next
50250 Case 2: Exit Sub
50260 Case 3: End
50270 End Select
50280 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub MoveUpFilenameSubstitutions()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim tStrL As String, tStrR As String
50050  With lsvFilenameSubst
50060   tStrL = .ListItems(.SelectedItem.Index).Text
50070   tStrR = .ListItems(.SelectedItem.Index).SubItems(1)
50080   .ListItems(.SelectedItem.Index).Text = .ListItems(.SelectedItem.Index - 1).Text
50090   .ListItems(.SelectedItem.Index).SubItems(1) = .ListItems(.SelectedItem.Index - 1).SubItems(1)
50100   .ListItems(.SelectedItem.Index - 1).Text = tStrL
50110   .ListItems(.SelectedItem.Index - 1).SubItems(1) = tStrR
50120   .ListItems(.SelectedItem.Index - 1).Selected = True
50130   .ListItems(.SelectedItem.Index).EnsureVisible
50140  End With
50150  Set_txtFilenameSubst
50160 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50170 Exit Sub
ErrPtnr_OnError:
50190 Select Case ErrPtnr.OnError("frmOptions", "MoveUpFilenameSubstitutions")
Case 0: Resume
50210 Case 1: Resume Next
50220 Case 2: Exit Sub
50230 Case 3: End
50240 End Select
50250 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub MoveDownFilenameSubstitutions()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim tStrL As String, tStrR As String
50050  With lsvFilenameSubst
50060   tStrL = .ListItems(.SelectedItem.Index).Text
50070   tStrR = .ListItems(.SelectedItem.Index).SubItems(1)
50080   .ListItems(.SelectedItem.Index).Text = .ListItems(.SelectedItem.Index + 1).Text
50090   .ListItems(.SelectedItem.Index).SubItems(1) = .ListItems(.SelectedItem.Index + 1).SubItems(1)
50100   .ListItems(.SelectedItem.Index + 1).Text = tStrL
50110   .ListItems(.SelectedItem.Index + 1).SubItems(1) = tStrR
50120   .ListItems(.SelectedItem.Index + 1).Selected = True
50130   .ListItems(.SelectedItem.Index).EnsureVisible
50140  End With
50150  Set_txtFilenameSubst
50160 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50170 Exit Sub
ErrPtnr_OnError:
50190 Select Case ErrPtnr.OnError("frmOptions", "MoveDownFilenameSubstitutions")
Case 0: Resume
50210 Case 1: Resume Next
50220 Case 2: Exit Sub
50230 Case 3: End
50240 End Select
50250 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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
50090    cmdFilenameSubst(3).Enabled = True
50100    cmdFilenameSubst(4).Enabled = True
50110   Else
50120    cmdFilenameSubst(3).Enabled = False
50130    cmdFilenameSubst(4).Enabled = False
50140  End If
50150  If lsvFilenameSubst.ListItems.Count > 0 Then
50160   If lsvFilenameSubst.SelectedItem.Index = 1 Then
50170    cmdFilenameSubst(3).Enabled = False
50180   End If
50190   If lsvFilenameSubst.SelectedItem.Index = lsvFilenameSubst.ListItems.Count Then
50200    cmdFilenameSubst(4).Enabled = False
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
