VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Options"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9315
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame fraProgAutosave 
      Caption         =   "Autosave"
      Height          =   3765
      Left            =   2760
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   6255
      Begin VB.TextBox txtAutoSaveDirectoryPreview 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   166
         Top             =   3330
         Width           =   6015
      End
      Begin VB.TextBox txtAutoSaveFilenamePreview 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   160
         Top             =   2025
         Width           =   6015
      End
      Begin VB.CheckBox chkUseAutosave 
         Caption         =   "Use Autosave"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   6015
      End
      Begin VB.CheckBox chkUseAutosaveDirectory 
         Caption         =   "For autosave use this directory"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2640
         Width           =   5895
      End
      Begin VB.TextBox txtAutosaveDirectory 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   3000
         Width           =   5535
      End
      Begin VB.CommandButton cmdGetAutosaveDirectory 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   285
         Left            =   5760
         TabIndex        =   11
         Top             =   3000
         Width           =   375
      End
      Begin VB.TextBox txtAutosaveFilename 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Text            =   "<DateTime>"
         Top             =   1680
         Width           =   3495
      End
      Begin VB.ComboBox cmbAutoSaveFilenameTokens 
         Height          =   315
         ItemData        =   "frmOptions.frx":548A
         Left            =   3690
         List            =   "frmOptions.frx":548C
         Style           =   2  'Dropdown-Liste
         TabIndex        =   9
         Top             =   1665
         Width           =   2460
      End
      Begin VB.ComboBox cmbAutosaveFormat 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown-Liste
         TabIndex        =   8
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblAutosaveFilename 
         AutoSize        =   -1  'True
         Caption         =   "Filename"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   630
      End
      Begin VB.Label lblAutosaveformat 
         Caption         =   "Autosaveformat"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label lblAutosaveFilenameTokens 
         AutoSize        =   -1  'True
         Caption         =   "Add a Filename-Token"
         Height          =   195
         Left            =   3720
         TabIndex        =   15
         Top             =   1440
         Width           =   1605
      End
   End
   Begin VB.Frame fraPDFFonts 
      Caption         =   "Font Options"
      Height          =   1740
      Left            =   3720
      TabIndex        =   83
      Top             =   2760
      Visible         =   0   'False
      Width           =   5775
      Begin VB.TextBox txtPDFSubSetPerc 
         Height          =   285
         Left            =   360
         TabIndex        =   86
         Top             =   1320
         Width           =   495
      End
      Begin VB.CheckBox chkPDFSubSetFonts 
         Caption         =   "Subset Fonts, when percentage of used characters below:"
         Height          =   495
         Left            =   120
         TabIndex        =   85
         Top             =   780
         Width           =   5535
      End
      Begin VB.CheckBox chkPDFEmbedAll 
         Caption         =   "Embed all Fonts"
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   360
         Width           =   5175
      End
      Begin VB.Label lblPDFPerc 
         Caption         =   "%"
         Height          =   255
         Left            =   960
         TabIndex        =   87
         Top             =   1320
         Width           =   615
      End
   End
   Begin VB.Frame fraProgSave 
      Caption         =   "Save"
      Height          =   4335
      Left            =   2880
      TabIndex        =   106
      Top             =   1560
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Frame fraFilenameSubstitutions 
         Caption         =   "Filename substitutions"
         Height          =   2415
         Left            =   120
         TabIndex        =   108
         Top             =   1800
         Width           =   6015
         Begin VB.CommandButton cmdFilenameSubst 
            Caption         =   "Delete"
            Height          =   375
            Index           =   2
            Left            =   4440
            TabIndex        =   117
            Top             =   1560
            Width           =   1455
         End
         Begin VB.CommandButton cmdFilenameSubst 
            Caption         =   "Change"
            Height          =   375
            Index           =   1
            Left            =   4440
            TabIndex        =   116
            Top             =   1080
            Width           =   1455
         End
         Begin VB.CommandButton cmdFilenameSubst 
            Caption         =   "Add"
            Height          =   375
            Index           =   0
            Left            =   4440
            TabIndex        =   115
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtFilenameSubst 
            Height          =   285
            Index           =   0
            Left            =   600
            TabIndex        =   114
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmdFilenameSubst 
            Enabled         =   0   'False
            Height          =   420
            Index           =   3
            Left            =   120
            Picture         =   "frmOptions.frx":548E
            Style           =   1  'Grafisch
            TabIndex        =   113
            Top             =   795
            Width           =   375
         End
         Begin VB.CommandButton cmdFilenameSubst 
            Enabled         =   0   'False
            Height          =   420
            Index           =   4
            Left            =   120
            Picture         =   "frmOptions.frx":5818
            Style           =   1  'Grafisch
            TabIndex        =   112
            Top             =   1320
            Width           =   375
         End
         Begin VB.CheckBox chkFilenameSubst 
            Caption         =   "Substitutions only in <Title>"
            Height          =   255
            Left            =   240
            TabIndex        =   111
            Top             =   2040
            Value           =   1  'Aktiviert
            Width           =   3255
         End
         Begin VB.TextBox txtFilenameSubst 
            Height          =   285
            Index           =   1
            Left            =   2640
            TabIndex        =   110
            Top             =   240
            Width           =   1695
         End
         Begin MSComctlLib.ListView lsvFilenameSubst 
            Height          =   1335
            Left            =   600
            TabIndex        =   109
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
         Begin VB.Label lblEqual 
            Caption         =   "="
            Height          =   255
            Left            =   2400
            TabIndex        =   118
            Top             =   240
            Width           =   135
         End
      End
      Begin VB.TextBox txtSavePreview 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   161
         Top             =   840
         Width           =   6015
      End
      Begin VB.ComboBox cmbSaveFilenameTokens 
         Height          =   315
         ItemData        =   "frmOptions.frx":5BA2
         Left            =   3720
         List            =   "frmOptions.frx":5BA4
         Style           =   2  'Dropdown-Liste
         TabIndex        =   120
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtSaveFilename 
         Height          =   285
         Left            =   120
         TabIndex        =   119
         Text            =   "<Title>"
         Top             =   480
         Width           =   3495
      End
      Begin VB.CheckBox chkSpaces 
         Caption         =   "Remove leading and trailing spaces"
         Height          =   255
         Left            =   120
         TabIndex        =   107
         Top             =   1320
         Value           =   1  'Aktiviert
         Width           =   6015
      End
      Begin VB.Label lblSaveFilenameTokens 
         AutoSize        =   -1  'True
         Caption         =   "Add a Filename-Token"
         Height          =   195
         Left            =   3720
         TabIndex        =   122
         Top             =   240
         Width           =   1605
      End
      Begin VB.Label lblSaveFilename 
         AutoSize        =   -1  'True
         Caption         =   "Filename"
         Height          =   195
         Left            =   120
         TabIndex        =   121
         Top             =   240
         Width           =   630
      End
   End
   Begin VB.Frame fraProgDocument 
      Caption         =   "Document"
      Height          =   1935
      Left            =   2760
      TabIndex        =   22
      Top             =   1680
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CheckBox chkUseCreationDateNow 
         Caption         =   "Use the current Date/Time for 'Creation Date'"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1440
         Width           =   6000
      End
      Begin VB.TextBox txtStandardAuthor 
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   3495
      End
      Begin VB.CheckBox chkUseStandardAuthor 
         Caption         =   "Use Standardauthor"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   6000
      End
      Begin VB.ComboBox cmbAuthorTokens 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmOptions.frx":5BA6
         Left            =   3720
         List            =   "frmOptions.frx":5BA8
         Style           =   2  'Dropdown-Liste
         TabIndex        =   23
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label lblAuthorTokens 
         AutoSize        =   -1  'True
         Caption         =   "Add a Author-Token"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3720
         TabIndex        =   27
         Top             =   600
         Width           =   1440
      End
   End
   Begin VB.Frame fraProgGeneral 
      Caption         =   "General"
      Height          =   4935
      Left            =   3000
      TabIndex        =   28
      Top             =   1080
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Frame fraShellintegration 
         Caption         =   "Shell integration"
         Height          =   1095
         Left            =   120
         TabIndex        =   157
         Top             =   2880
         Width           =   6015
         Begin VB.CommandButton cmdShellintegration 
            Caption         =   "Remove shell integration"
            Height          =   615
            Index           =   1
            Left            =   3240
            TabIndex        =   159
            Top             =   360
            Width           =   2655
         End
         Begin VB.CommandButton cmdShellintegration 
            Caption         =   "Integrate PDFCreator into shell"
            Height          =   615
            Index           =   0
            Left            =   120
            TabIndex        =   158
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.CheckBox chkNoConfirmMessageSwitchingDefaultprinter 
         Caption         =   "No confirm message switching PDFCreator temporarly as default printer."
         Height          =   495
         Left            =   120
         TabIndex        =   142
         Top             =   2160
         Width           =   5775
      End
      Begin VB.CommandButton cmdTestpage 
         Caption         =   "Print testpage"
         Height          =   495
         Left            =   120
         TabIndex        =   31
         Top             =   4320
         Width           =   1815
      End
      Begin VB.CommandButton cmdAsso 
         Caption         =   "Associate PDFCreator with Postscript files"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   3855
      End
      Begin MSComctlLib.Slider sldProcessPriority 
         Height          =   495
         Left            =   120
         TabIndex        =   30
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
      Begin VB.Label lblProcessPriority 
         AutoSize        =   -1  'True
         Caption         =   "Processpriority: Normal"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   1605
      End
   End
   Begin VB.Frame fraProgFont 
      Caption         =   "Programfont"
      Height          =   4935
      Left            =   2760
      TabIndex        =   33
      Top             =   1080
      Visible         =   0   'False
      Width           =   6255
      Begin VB.ComboBox cmbProgramFontsize 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   5400
         TabIndex        =   167
         Text            =   "8"
         Top             =   600
         Width           =   765
      End
      Begin VB.ComboBox cmbFonts 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown-Liste
         TabIndex        =   38
         Top             =   600
         Width           =   2775
      End
      Begin VB.ComboBox cmbCharset 
         Height          =   315
         Left            =   3000
         TabIndex        =   37
         Text            =   "cmbCharset"
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtTest 
         Height          =   2655
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   36
         Top             =   1320
         Width           =   6015
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test"
         Height          =   495
         Left            =   120
         TabIndex        =   35
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelTest 
         Caption         =   "CancelTest"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1680
         TabIndex        =   34
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label lblProgfont 
         Caption         =   "Programfont"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblProgcharset 
         Caption         =   "Charset"
         Height          =   255
         Left            =   3000
         TabIndex        =   41
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblTesttext 
         Caption         =   "Here you can test the font."
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1080
         Width           =   5295
      End
      Begin VB.Label lblSize 
         Caption         =   "Size"
         Height          =   255
         Left            =   5400
         TabIndex        =   39
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame fraProgGhostscript 
      Caption         =   "Ghostscript"
      Height          =   930
      Left            =   2760
      TabIndex        =   145
      Top             =   1080
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton cmdGetgsresourceDirectory 
         Caption         =   "..."
         Height          =   255
         Left            =   5760
         TabIndex        =   165
         Top             =   3000
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtGSresource 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   163
         Top             =   3000
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.CommandButton cmdGetgsfontsDirectory 
         Caption         =   "..."
         Height          =   255
         Left            =   5760
         TabIndex        =   154
         Top             =   2400
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdGetgslibDirectory 
         Caption         =   "..."
         Height          =   255
         Left            =   5760
         TabIndex        =   153
         Top             =   1800
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtGSfonts 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   152
         Top             =   2400
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.TextBox txtGSlib 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   151
         Top             =   1800
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.TextBox txtGSbin 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   149
         Top             =   1200
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.CommandButton cmdGetgsbinDirectory 
         Caption         =   "..."
         Height          =   255
         Left            =   5760
         TabIndex        =   148
         Top             =   1200
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox cmbGhostscript 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown-Liste
         TabIndex        =   147
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label lblGhostscriptResource 
         AutoSize        =   -1  'True
         Caption         =   "Ghostscript Resource"
         Height          =   195
         Left            =   240
         TabIndex        =   164
         Top             =   2760
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label lblGSfonts 
         AutoSize        =   -1  'True
         Caption         =   "Ghostscript Fonts"
         Height          =   195
         Left            =   240
         TabIndex        =   156
         Top             =   2160
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label lblGSlib 
         AutoSize        =   -1  'True
         Caption         =   "Ghostscript Libraries"
         Height          =   195
         Left            =   240
         TabIndex        =   155
         Top             =   1560
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label lblGSbin 
         AutoSize        =   -1  'True
         Caption         =   "Ghostscript Binaries"
         Height          =   195
         Left            =   240
         TabIndex        =   150
         Top             =   960
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label lblGhostscriptversion 
         Caption         =   "Ghostscriptversion"
         Height          =   255
         Left            =   240
         TabIndex        =   146
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.Frame fraPDFCompress 
      Caption         =   "Compression"
      Height          =   3855
      Left            =   2760
      TabIndex        =   43
      Top             =   1080
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Frame fraPDFColor 
         Caption         =   "Color Images"
         Height          =   975
         Left            =   120
         TabIndex        =   59
         Top             =   600
         Width           =   5535
         Begin VB.CheckBox chkPDFColorComp 
            Caption         =   "Compress"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox cmbPDFColorComp 
            Height          =   315
            ItemData        =   "frmOptions.frx":5BAA
            Left            =   120
            List            =   "frmOptions.frx":5BAC
            Style           =   2  'Dropdown-Liste
            TabIndex        =   63
            Top             =   540
            Width           =   2055
         End
         Begin VB.CheckBox chkPDFColorResample 
            Caption         =   "Resample"
            Height          =   255
            Left            =   2280
            TabIndex        =   62
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox cmbPDFColorResample 
            Height          =   315
            ItemData        =   "frmOptions.frx":5BAE
            Left            =   2280
            List            =   "frmOptions.frx":5BB0
            Style           =   2  'Dropdown-Liste
            TabIndex        =   61
            Tag             =   "Bicubic|Subsample|Average"
            Top             =   540
            Width           =   2055
         End
         Begin VB.TextBox txtPDFColorRes 
            Height          =   285
            Left            =   4440
            TabIndex        =   60
            Top             =   540
            Width           =   735
         End
         Begin VB.Label lblPDFColorRes 
            Caption         =   "Resolution"
            Height          =   255
            Left            =   4440
            TabIndex        =   65
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
      Begin VB.Frame fraPDFMono 
         Caption         =   "Monochrome Images"
         Height          =   975
         Left            =   120
         TabIndex        =   51
         Top             =   2760
         Width           =   5535
         Begin VB.TextBox txtPDFMonoRes 
            Height          =   285
            Left            =   4440
            TabIndex        =   56
            Top             =   540
            Width           =   735
         End
         Begin VB.ComboBox cmbPDFMonoResample 
            Height          =   315
            ItemData        =   "frmOptions.frx":5BB2
            Left            =   2280
            List            =   "frmOptions.frx":5BB4
            Style           =   2  'Dropdown-Liste
            TabIndex        =   55
            Tag             =   "Bicubic|Subsample|Average"
            Top             =   540
            Width           =   2055
         End
         Begin VB.CheckBox chkPDFMonoResample 
            Caption         =   "Resample"
            Height          =   255
            Left            =   2280
            TabIndex        =   54
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox cmbPDFMonoComp 
            Height          =   315
            ItemData        =   "frmOptions.frx":5BB6
            Left            =   120
            List            =   "frmOptions.frx":5BB8
            Style           =   2  'Dropdown-Liste
            TabIndex        =   53
            Top             =   540
            Width           =   2055
         End
         Begin VB.CheckBox chkPDFMonoComp 
            Caption         =   "Compress"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblPDFMonoRes 
            Caption         =   "Resolution"
            Height          =   255
            Left            =   4440
            TabIndex        =   57
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame fraPDFGrey 
         Caption         =   "Greyscale Images"
         Height          =   975
         Left            =   120
         TabIndex        =   44
         Top             =   1680
         Width           =   5535
         Begin VB.CheckBox chkPDFGreyComp 
            Caption         =   "Compress"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox cmbPDFGreyComp 
            Height          =   315
            ItemData        =   "frmOptions.frx":5BBA
            Left            =   120
            List            =   "frmOptions.frx":5BBC
            Style           =   2  'Dropdown-Liste
            TabIndex        =   48
            Top             =   540
            Width           =   2055
         End
         Begin VB.CheckBox chkPDFGreyResample 
            Caption         =   "Resample"
            Height          =   255
            Left            =   2280
            TabIndex        =   47
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox cmbPDFGreyResample 
            Height          =   315
            ItemData        =   "frmOptions.frx":5BBE
            Left            =   2280
            List            =   "frmOptions.frx":5BC0
            Style           =   2  'Dropdown-Liste
            TabIndex        =   46
            Tag             =   "Bicubic|Subsample|Average"
            Top             =   540
            Width           =   2055
         End
         Begin VB.TextBox txtPDFGreyRes 
            Height          =   285
            Left            =   4440
            TabIndex        =   45
            Top             =   540
            Width           =   735
         End
         Begin VB.Label lblPDFGreyRes 
            Caption         =   "Resolution"
            Height          =   255
            Left            =   4440
            TabIndex        =   50
            Top             =   240
            Width           =   975
         End
      End
   End
   Begin VB.Frame fraPDFSecurity 
      Caption         =   "Security"
      Height          =   5415
      Left            =   2880
      TabIndex        =   123
      Top             =   1320
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Frame fraPDFEncryptor 
         Caption         =   "Encryptor"
         Height          =   855
         Left            =   120
         TabIndex        =   143
         Top             =   600
         Width           =   5535
         Begin VB.ComboBox cmbPDFEncryptor 
            Height          =   315
            ItemData        =   "frmOptions.frx":5BC2
            Left            =   240
            List            =   "frmOptions.frx":5BC4
            Style           =   2  'Dropdown-Liste
            TabIndex        =   144
            Top             =   360
            Width           =   5175
         End
      End
      Begin VB.CheckBox chkUseSecurity 
         Caption         =   "Use Security"
         Height          =   255
         Left            =   120
         TabIndex        =   140
         Top             =   240
         Width           =   5535
      End
      Begin VB.Frame fraPDFEncLevel 
         Caption         =   "Encryption Level"
         Height          =   855
         Left            =   120
         TabIndex        =   137
         Top             =   1560
         Width           =   5535
         Begin VB.OptionButton optEncLow 
            Caption         =   "Low (40 Bit - Adobe Acrobat 3.0 and above)"
            Height          =   255
            Left            =   90
            TabIndex        =   139
            Top             =   225
            Width           =   5385
         End
         Begin VB.OptionButton optEncHigh 
            Caption         =   "High (128 Bit - Adobe Acrobat 5.0 and above)"
            Height          =   255
            Left            =   90
            TabIndex        =   138
            Top             =   450
            Width           =   5385
         End
      End
      Begin VB.Frame fraSecurityPass 
         Caption         =   "Passwords"
         Height          =   855
         Left            =   120
         TabIndex        =   134
         Top             =   2520
         Width           =   5535
         Begin VB.CheckBox chkUserPass 
            Caption         =   "Password required to open document"
            Height          =   255
            Left            =   90
            TabIndex        =   136
            Top             =   240
            Width           =   5385
         End
         Begin VB.CheckBox chkOwnerPass 
            Caption         =   "Password required to change Permissions and Passwords"
            Height          =   255
            Left            =   90
            TabIndex        =   135
            Top             =   480
            Width           =   5385
         End
      End
      Begin VB.Frame fraPDFPermissions 
         Caption         =   "Disallow User to"
         Height          =   855
         Left            =   120
         TabIndex        =   129
         Top             =   3480
         Width           =   5535
         Begin VB.CheckBox chkAllowPrinting 
            Caption         =   "print the document"
            Height          =   255
            Left            =   90
            TabIndex        =   133
            Top             =   240
            Width           =   2650
         End
         Begin VB.CheckBox chkAllowCopy 
            Caption         =   "copy text and images"
            Height          =   255
            Left            =   90
            TabIndex        =   132
            Top             =   480
            Width           =   2650
         End
         Begin VB.CheckBox chkAllowModifyContents 
            Caption         =   "modify the document"
            Height          =   255
            Left            =   2800
            TabIndex        =   131
            Top             =   240
            Width           =   2650
         End
         Begin VB.CheckBox chkAllowModifyAnnotations 
            Caption         =   "modify comments"
            Height          =   255
            Left            =   2800
            TabIndex        =   130
            Top             =   480
            Width           =   2650
         End
      End
      Begin VB.Frame fraPDFHighPermissions 
         Caption         =   "Enhanced Permissions (128 Bit only)"
         Height          =   855
         Left            =   120
         TabIndex        =   124
         Top             =   4440
         Width           =   5535
         Begin VB.CheckBox chkAllowDegradedPrinting 
            Caption         =   "Allow printing in low resolution"
            Height          =   255
            Left            =   90
            TabIndex        =   128
            Top             =   240
            Width           =   2650
         End
         Begin VB.CheckBox chkAllowFillIn 
            Caption         =   "Allow filling in form fields"
            Height          =   255
            Left            =   2800
            TabIndex        =   127
            Top             =   240
            Width           =   2650
         End
         Begin VB.CheckBox chkAllowScreenReaders 
            Caption         =   "Allow Screen Readers"
            Height          =   255
            Left            =   90
            TabIndex        =   126
            Top             =   480
            Width           =   2650
         End
         Begin VB.CheckBox chkAllowAssembly 
            Caption         =   "Allow changes to the Assembly"
            Height          =   255
            Left            =   2800
            TabIndex        =   125
            Top             =   480
            Width           =   2650
         End
      End
   End
   Begin VB.Frame fraProgDirectories 
      Caption         =   "Directories"
      Height          =   1095
      Left            =   2760
      TabIndex        =   18
      Top             =   3840
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton cmdUsertempPath 
         Height          =   285
         Left            =   5760
         Picture         =   "frmOptions.frx":5BC6
         Style           =   1  'Grafisch
         TabIndex        =   162
         Top             =   585
         Width           =   375
      End
      Begin VB.CommandButton cmdGetTemppath 
         Caption         =   "..."
         Height          =   285
         Left            =   5265
         TabIndex        =   20
         Top             =   585
         Width           =   375
      End
      Begin VB.TextBox txtTemppath 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   600
         Width           =   4965
      End
      Begin VB.Label lblPrintTempPath 
         AutoSize        =   -1  'True
         Caption         =   "Temppath"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Frame fraBitmapGeneral 
      Caption         =   "Bitmap"
      Height          =   1935
      Left            =   2880
      TabIndex        =   70
      Top             =   1080
      Width           =   6255
      Begin VB.TextBox txtBitmapResolution 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   1920
         TabIndex        =   77
         Text            =   "72"
         Top             =   480
         Width           =   495
      End
      Begin VB.ComboBox cmbPNGColors 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown-Liste
         TabIndex        =   76
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox txtJPEGQuality 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   1920
         TabIndex        =   75
         Text            =   "75"
         Top             =   1440
         Width           =   495
      End
      Begin VB.ComboBox cmbJPEGColors 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown-Liste
         TabIndex        =   74
         Top             =   1080
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
      Begin VB.ComboBox cmbPCXColors 
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown-Liste
         TabIndex        =   72
         Top             =   1320
         Width           =   2175
      End
      Begin VB.ComboBox cmbTIFFColors 
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown-Liste
         TabIndex        =   71
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblBitmapResolution 
         Alignment       =   1  'Rechts
         Caption         =   "Resolution:"
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblBitmapDPI 
         Caption         =   "dpi"
         Height          =   255
         Left            =   2520
         TabIndex        =   81
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblBitmapColors 
         Alignment       =   1  'Rechts
         Caption         =   "Colors:"
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblJPEGQuality 
         Alignment       =   1  'Rechts
         Caption         =   "Quality:"
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblJPEQQualityProzent 
         Caption         =   "%"
         Height          =   255
         Left            =   2520
         TabIndex        =   78
         Top             =   1440
         Width           =   255
      End
   End
   Begin VB.Frame fraPDFGeneral 
      Caption         =   "General Options"
      Height          =   2895
      Left            =   3240
      TabIndex        =   95
      Top             =   1800
      Visible         =   0   'False
      Width           =   5775
      Begin VB.ComboBox cmbPDFRotate 
         Height          =   315
         ItemData        =   "frmOptions.frx":5F50
         Left            =   2400
         List            =   "frmOptions.frx":5F52
         Style           =   2  'Dropdown-Liste
         TabIndex        =   100
         Tag             =   "None|All|PageByPage"
         Top             =   840
         Width           =   2655
      End
      Begin VB.ComboBox cmbPDFCompat 
         Height          =   315
         ItemData        =   "frmOptions.frx":5F54
         Left            =   2400
         List            =   "frmOptions.frx":5F56
         Style           =   2  'Dropdown-Liste
         TabIndex        =   99
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox txtPDFRes 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   2400
         TabIndex        =   98
         Text            =   "600"
         Top             =   1320
         Width           =   615
      End
      Begin VB.ComboBox cmbPDFOverprint 
         Height          =   315
         ItemData        =   "frmOptions.frx":5F58
         Left            =   2400
         List            =   "frmOptions.frx":5F5A
         Style           =   2  'Dropdown-Liste
         TabIndex        =   97
         Top             =   1860
         Width           =   2655
      End
      Begin VB.CheckBox chkPDFASCII85 
         Caption         =   "Convert binary data to ASCII85"
         Height          =   255
         Left            =   2400
         TabIndex        =   96
         Top             =   2400
         Width           =   3255
      End
      Begin VB.Label lblPDFAutoRotate 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Auto-Rotate Pages:"
         Height          =   255
         Left            =   120
         TabIndex        =   105
         Top             =   900
         Width           =   2175
      End
      Begin VB.Label lblPDFCompat 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Compatibility:"
         Height          =   255
         Left            =   120
         TabIndex        =   104
         Top             =   420
         Width           =   2175
      End
      Begin VB.Label lblPDFResolution 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Resolution:"
         Height          =   255
         Left            =   120
         TabIndex        =   103
         Top             =   1380
         Width           =   2175
      End
      Begin VB.Label lblPDFOverprint 
         Alignment       =   1  'Rechts
         Caption         =   "Overprint:"
         Height          =   375
         Left            =   120
         TabIndex        =   102
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label lblPDFDPI 
         BackStyle       =   0  'Transparent
         Caption         =   "dpi"
         Height          =   255
         Left            =   3120
         TabIndex        =   101
         Top             =   1320
         Width           =   735
      End
   End
   Begin VB.Frame fraPSGeneral 
      Caption         =   "Postscript"
      Height          =   1095
      Left            =   2760
      TabIndex        =   66
      Top             =   1080
      Visible         =   0   'False
      Width           =   6255
      Begin VB.ComboBox cmbPSLanguageLevel 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown-Liste
         TabIndex        =   68
         Top             =   480
         Width           =   735
      End
      Begin VB.ComboBox cmbEPSLanguageLevel 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown-Liste
         TabIndex        =   67
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblLangLevel 
         Alignment       =   1  'Rechts
         Caption         =   "Language Level:"
         Height          =   255
         Left            =   240
         TabIndex        =   69
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame fraPDFColors 
      Caption         =   "Color Options"
      Height          =   3495
      Left            =   3480
      TabIndex        =   88
      Top             =   2280
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CheckBox chkPDFCMYKtoRGB 
         Caption         =   "Convert CMYK Images to RGB"
         Height          =   255
         Left            =   120
         TabIndex        =   94
         Top             =   840
         Width           =   3255
      End
      Begin VB.Frame fraPDFColorOptions 
         Caption         =   "Options"
         Height          =   1455
         Left            =   120
         TabIndex        =   90
         Top             =   1920
         Width           =   5535
         Begin VB.CheckBox chkPDFPreserveOverprint 
            Caption         =   "Preserve Overprint Settings"
            Height          =   255
            Left            =   120
            TabIndex        =   93
            Top             =   360
            Width           =   5175
         End
         Begin VB.CheckBox chkPDFPreserveTransfer 
            Caption         =   "Preserve Transfer Functions"
            Height          =   255
            Left            =   120
            TabIndex        =   92
            Tag             =   "Remove|Preserve"
            Top             =   720
            Width           =   5175
         End
         Begin VB.CheckBox chkPDFPreserveHalftone 
            Caption         =   "Preserve Halftone Information"
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   1080
            Width           =   5175
         End
      End
      Begin VB.ComboBox cmbPDFColorModel 
         Height          =   315
         ItemData        =   "frmOptions.frx":5F5C
         Left            =   120
         List            =   "frmOptions.frx":5F5E
         Style           =   2  'Dropdown-Liste
         TabIndex        =   89
         Tag             =   "RGB|CMYK|GRAY"
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   7560
      TabIndex        =   5
      Top             =   7320
      Width           =   1575
   End
   Begin MSComctlLib.TreeView trv 
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   13573
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   7320
      Width           =   1575
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
   Begin VB.CommandButton cmdStyle 
      Enabled         =   0   'False
      Height          =   6255
      Left            =   2640
      TabIndex        =   6
      Top             =   960
      Width           =   6495
   End
   Begin MSComctlLib.TabStrip tbstrPDFOptions 
      Height          =   4935
      Left            =   3000
      TabIndex        =   141
      Top             =   1320
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   8705
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   2640
      X2              =   9120
      Y1              =   860
      Y2              =   860
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2640
      X2              =   9120
      Y1              =   840
      Y2              =   840
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

Private Sub SetPDFColorComprSettings()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If chkPDFColorComp.Value = 1 Then
50050    cmbPDFColorComp.Enabled = True
50060    If cmbPDFColorComp.ListIndex = 0 Then
50070      chkPDFColorResample.Enabled = False
50080      cmbPDFColorResample.Enabled = False
50090      lblPDFColorRes.Enabled = False
50100      txtPDFColorRes.Enabled = False
50110     Else
50120      chkPDFColorResample.Enabled = True
50130      If chkPDFColorResample.Value = 1 Then
50140        cmbPDFColorResample.Enabled = True
50150        lblPDFColorRes.Enabled = True
50160        txtPDFColorRes.Enabled = True
50170       Else
50180        cmbPDFColorResample.Enabled = False
50190        lblPDFColorRes.Enabled = False
50200        txtPDFColorRes.Enabled = False
50210      End If
50220    End If
50230   Else
50240    cmbPDFColorComp.Enabled = False
50250    chkPDFColorResample.Enabled = False
50260    cmbPDFColorResample.Enabled = False
50270    lblPDFColorRes.Enabled = False
50280    txtPDFColorRes.Enabled = False
50290  End If
50300 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50310 Exit Sub
ErrPtnr_OnError:
50331 Select Case ErrPtnr.OnError("frmOptions", "SetPDFColorComprSettings")
      Case 0: Resume
50350 Case 1: Resume Next
50360 Case 2: Exit Sub
50370 Case 3: End
50380 End Select
50390 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetPDFGreyComprSettings()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If chkPDFGreyComp.Value = 1 Then
50050    cmbPDFGreyComp.Enabled = True
50060    If cmbPDFGreyComp.ListIndex = 0 Then
50070      chkPDFGreyResample.Enabled = False
50080      cmbPDFGreyResample.Enabled = False
50090      lblPDFGreyRes.Enabled = False
50100      txtPDFGreyRes.Enabled = False
50110     Else
50120      chkPDFGreyResample.Enabled = True
50130      If chkPDFGreyResample.Value = 1 Then
50140        cmbPDFGreyResample.Enabled = True
50150        lblPDFGreyRes.Enabled = True
50160        txtPDFGreyRes.Enabled = True
50170       Else
50180        cmbPDFGreyResample.Enabled = False
50190        lblPDFGreyRes.Enabled = False
50200        txtPDFGreyRes.Enabled = False
50210      End If
50220    End If
50230   Else
50240    cmbPDFGreyComp.Enabled = False
50250    chkPDFGreyResample.Enabled = False
50260    cmbPDFGreyResample.Enabled = False
50270    lblPDFGreyRes.Enabled = False
50280    txtPDFGreyRes.Enabled = False
50290  End If
50300 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50310 Exit Sub
ErrPtnr_OnError:
50331 Select Case ErrPtnr.OnError("frmOptions", "SetPDFGreyComprSettings")
      Case 0: Resume
50350 Case 1: Resume Next
50360 Case 2: Exit Sub
50370 Case 3: End
50380 End Select
50390 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetPDFMonoComprSettings()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If chkPDFMonoComp.Value = 1 Then
50050    cmbPDFMonoComp.Enabled = True
50060    chkPDFMonoResample.Enabled = True
50070    If chkPDFMonoResample.Value = 1 Then
50080      cmbPDFMonoResample.Enabled = True
50090      lblPDFMonoRes.Enabled = True
50100      txtPDFMonoRes.Enabled = True
50110     Else
50120      cmbPDFMonoResample.Enabled = False
50130      lblPDFMonoRes.Enabled = False
50140      txtPDFMonoRes.Enabled = False
50150    End If
50160   Else
50170    cmbPDFMonoComp.Enabled = False
50180    chkPDFMonoResample.Enabled = False
50190    cmbPDFMonoResample.Enabled = False
50200    lblPDFMonoRes.Enabled = False
50210    txtPDFMonoRes.Enabled = False
50220  End If
50230 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50240 Exit Sub
ErrPtnr_OnError:
50261 Select Case ErrPtnr.OnError("frmOptions", "SetPDFMonoComprSettings")
      Case 0: Resume
50280 Case 1: Resume Next
50290 Case 2: Exit Sub
50300 Case 3: End
50310 End Select
50320 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkOwnerPass_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If chkUserPass.Value = 0 Then
50050   If chkOwnerPass.Value = 0 Then
50060    chkOwnerPass.Value = 1
50070   End If
50080  End If
50090 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50100 Exit Sub
ErrPtnr_OnError:
50121 Select Case ErrPtnr.OnError("frmOptions", "chkOwnerPass_Click")
      Case 0: Resume
50140 Case 1: Resume Next
50150 Case 2: Exit Sub
50160 Case 3: End
50170 End Select
50180 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkPDFColorComp_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  SetPDFColorComprSettings
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("frmOptions", "chkPDFColorComp_Click")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkPDFColorResample_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  SetPDFColorComprSettings
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("frmOptions", "chkPDFColorResample_Click")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkPDFGreyComp_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  SetPDFGreyComprSettings
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("frmOptions", "chkPDFGreyComp_Click")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkPDFGreyResample_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  SetPDFGreyComprSettings
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("frmOptions", "chkPDFGreyResample_Click")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkPDFMonoComp_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  SetPDFMonoComprSettings
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("frmOptions", "chkPDFMonoComp_Click")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkPDFMonoResample_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  SetPDFMonoComprSettings
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("frmOptions", "chkPDFMonoResample_Click")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkUseAutosave_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If chkUseAutosave.Value = 1 Then
50050    ViewAutosave True
50060   Else
50070    ViewAutosave False
50080  End If
50090 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50100 Exit Sub
ErrPtnr_OnError:
50121 Select Case ErrPtnr.OnError("frmOptions", "chkUseAutosave_Click")
      Case 0: Resume
50140 Case 1: Resume Next
50150 Case 2: Exit Sub
50160 Case 3: End
50170 End Select
50180 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkUseAutosaveDirectory_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If chkUseAutosaveDirectory.Value = 1 Then
50050    ViewAutosaveDirectory True
50060   Else
50070    ViewAutosaveDirectory False
50080  End If
50090 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50100 Exit Sub
ErrPtnr_OnError:
50121 Select Case ErrPtnr.OnError("frmOptions", "chkUseAutosaveDirectory_Click")
      Case 0: Resume
50140 Case 1: Resume Next
50150 Case 2: Exit Sub
50160 Case 3: End
50170 End Select
50180 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkUserPass_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If chkOwnerPass.Value = 0 Then
50050   If chkUserPass.Value = 0 Then
50060    chkUserPass.Value = 1
50070    chkOwnerPass.Value = 1
50080   End If
50090   SavePasswordsForThisSession = False
50100  End If
50110 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50120 Exit Sub
ErrPtnr_OnError:
50141 Select Case ErrPtnr.OnError("frmOptions", "chkUserPass_Click")
      Case 0: Resume
50160 Case 1: Resume Next
50170 Case 2: Exit Sub
50180 Case 3: End
50190 End Select
50200 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkUseSecurity_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  UpdateSecurityFields
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("frmOptions", "chkUseSecurity_Click")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkUseStandardAuthor_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If chkUseStandardAuthor.Value = 1 Then
50050    txtStandardAuthor.Enabled = True
50060    txtStandardAuthor.BackColor = &H80000005
50070    cmbAuthorTokens.Enabled = True
50080    lblAuthorTokens.Enabled = True
50090   Else
50100    txtStandardAuthor.Enabled = False
50110    txtStandardAuthor.BackColor = &H8000000F
50120    cmbAuthorTokens.Enabled = False
50130    lblAuthorTokens.Enabled = False
50140  End If
50150 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50160 Exit Sub
ErrPtnr_OnError:
50181 Select Case ErrPtnr.OnError("frmOptions", "chkUseStandardAuthor_Click")
      Case 0: Resume
50200 Case 1: Resume Next
50210 Case 2: Exit Sub
50220 Case 3: End
50230 End Select
50240 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbAuthorTokens_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  txtStandardAuthor.Text = txtStandardAuthor.Text & cmbAuthorTokens.Text
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("frmOptions", "cmbAuthorTokens_Click")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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
 If Err.number = 380 Then
  cmbCharset.Text = 0
 End If
 Err.Clear
End Sub

Private Sub cmbCharset_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  With cmbCharset
50050   .Text = .ItemData(.ListIndex)
50060  End With
50070  txtTest.Font.Charset = cmbCharset.Text
50080 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50090 Exit Sub
ErrPtnr_OnError:
50111 Select Case ErrPtnr.OnError("frmOptions", "cmbCharset_Click")
      Case 0: Resume
50130 Case 1: Resume Next
50140 Case 2: Exit Sub
50150 Case 3: End
50160 End Select
50170 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbCharset_KeyPress(KeyAscii As Integer)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim allow As String, tStr As String
50050  allow = "0123456789" & Chr$(8) & Chr$(13)
50060  tStr = Chr$(KeyAscii)
50070  If InStr(1, allow, tStr) = 0 Then
50080    KeyAscii = 0
50090  End If
50100 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50110 Exit Sub
ErrPtnr_OnError:
50131 Select Case ErrPtnr.OnError("frmOptions", "cmbCharset_KeyPress")
      Case 0: Resume
50150 Case 1: Resume Next
50160 Case 2: Exit Sub
50170 Case 3: End
50180 End Select
50190 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbCharset_Validate(Cancel As Boolean)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim i As Long, tStr As String
50050  tStr = ""
50060  For i = 1 To Len(cmbCharset.Text)
50070   If InStr("0123456789", Mid(cmbCharset.Text, i, 1)) > 0 Then
50080     tStr = tStr & Mid(cmbCharset.Text, i, 1)
50090    Else
50100     Exit For
50110   End If
50120  Next i
50130  If Len(Trim$(tStr)) = 0 Then
50140    cmbCharset.Text = 0
50150   Else
50160    cmbCharset.Text = tStr
50170  End If
50180 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50190 Exit Sub
ErrPtnr_OnError:
50211 Select Case ErrPtnr.OnError("frmOptions", "cmbCharset_Validate")
      Case 0: Resume
50230 Case 1: Resume Next
50240 Case 2: Exit Sub
50250 Case 3: End
50260 End Select
50270 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbAutoSaveFilenameTokens_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  txtAutosaveFilename.Text = txtAutosaveFilename.Text & cmbAutoSaveFilenameTokens.Text
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("frmOptions", "cmbAutoSaveFilenameTokens_Click")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbGhostscript_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim reg As clsRegistry, gsv As String, tsf() As String, Path As String, tStr As String
50050
50060  gsv = cmbGhostscript.List(cmbGhostscript.ListIndex)
50070  Set reg = New clsRegistry: reg.hkey = HKEY_LOCAL_MACHINE
50080
50090  If InStr(gsv, ":") Then
50100    reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50110    txtGSbin.Text = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryBinaries"))
50120    txtGSfonts.Text = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
50130    txtGSlib.Text = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryLibraries"))
50140    Set reg = Nothing
50150    Exit Sub
50160   Else
50170    If InStr(UCase$(gsv), "AFPL") Then
50180     If InStr(gsv, " ") > 0 Then
50190      tsf = Split(gsv, " ")
50200      reg.KeyRoot = "SOFTWARE\AFPL Ghostscript\" & tsf(UBound(tsf))
50210      tStr = reg.GetRegistryValue("GS_DLL")
50220      SplitPath tStr, , Path
50230      txtGSbin.Text = CompletePath(Path)
50240      If InStrRev(Path, "\") > 0 Then
50250       txtGSlib.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50260       txtGSfonts.Text = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50270      End If
50280     End If
50290    End If
50300    If InStr(UCase$(gsv), "GNU") Then
50310     If InStr(gsv, " ") > 0 Then
50320      tsf = Split(gsv, " ")
50330      reg.KeyRoot = "SOFTWARE\GNU Ghostscript\" & tsf(UBound(tsf))
50340      tStr = reg.GetRegistryValue("GS_DLL")
50350      SplitPath tStr, , Path
50360      txtGSbin.Text = CompletePath(Path)
50370      If InStrRev(Path, "\") > 0 Then
50380       txtGSlib.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50390       txtGSfonts.Text = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50400      End If
50410     End If
50420    End If
50430    If InStr(UCase$(gsv), "GPL") Then
50440     If InStr(gsv, " ") > 0 Then
50450      tsf = Split(gsv, " ")
50460      reg.KeyRoot = "SOFTWARE\GPL Ghostscript\" & tsf(UBound(tsf))
50470      tStr = reg.GetRegistryValue("GS_DLL")
50480      SplitPath tStr, , Path
50490      txtGSbin.Text = CompletePath(Path)
50500      If InStrRev(Path, "\") > 0 Then
50510       txtGSlib.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50520       txtGSfonts.Text = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50530      End If
50540     End If
50550    End If
50560  End If
50570
50580  Set reg = Nothing
50590 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50600 Exit Sub
ErrPtnr_OnError:
50621 Select Case ErrPtnr.OnError("frmOptions", "cmbGhostscript_Click")
      Case 0: Resume
50640 Case 1: Resume Next
50650 Case 2: Exit Sub
50660 Case 3: End
50670 End Select
50680 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbPDFColorComp_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  SetPDFColorComprSettings
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("frmOptions", "cmbPDFColorComp_Click")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbPDFGreyComp_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  SetPDFGreyComprSettings
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("frmOptions", "cmbPDFGreyComp_Click")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbPDFMonoComp_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  SetPDFMonoComprSettings
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("frmOptions", "cmbPDFMonoComp_Click")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbSaveFilenameTokens_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  txtSaveFilename.Text = txtSaveFilename.Text & cmbSaveFilenameTokens.Text
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("frmOptions", "cmbSaveFilenameTokens_Click")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbFonts_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  txtTest.Font.Name = cmbFonts.List(cmbFonts.ListIndex)
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("frmOptions", "cmbFonts_Click")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbPDFCompat_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  UpdateSecurityFields
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("frmOptions", "cmbPDFCompat_Click")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdCancel_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Unload Me
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("frmOptions", "cmdCancel_Click")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdCancelTest_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  With Options
50050   SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50060   cmbCharset.Text = .ProgramFontCharset
50070   SetFont frmMain, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50080  End With
50090 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50100 Exit Sub
ErrPtnr_OnError:
50121 Select Case ErrPtnr.OnError("frmOptions", "cmdCancelTest_Click")
      Case 0: Resume
50140 Case 1: Resume Next
50150 Case 2: Exit Sub
50160 Case 3: End
50170 End Select
50180 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdFilenameSubst_Click(Index As Integer)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50041  Select Case Index
              Case 0: ' Add
50060    AddFilenameSubstitutions
50070   Case 1: ' Change
50080    ChangeFilenameSubstitutions
50090   Case 2: ' Delete
50100    DeleteFilenameSubstitutions
50110   Case 3: ' Up
50120    MoveUpFilenameSubstitutions
50130   Case 4: ' Down
50140    MoveDownFilenameSubstitutions
50150  End Select
50160 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50170 Exit Sub
ErrPtnr_OnError:
50191 Select Case ErrPtnr.OnError("frmOptions", "cmdFilenameSubst_Click")
      Case 0: Resume
50210 Case 1: Resume Next
50220 Case 2: Exit Sub
50230 Case 3: End
50240 End Select
50250 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdGetAutosaveDirectory_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040 Dim strFolder As String
50050 strFolder = BrowseForFolder(Me.hwnd, LanguageStrings.OptionsAutosaveDirectoryPrompt)
50060 If Len(strFolder) = 0 Then Exit Sub
50070 txtAutosaveDirectory.Text = CompletePath(strFolder)
50080 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50090 Exit Sub
ErrPtnr_OnError:
50111 Select Case ErrPtnr.OnError("frmOptions", "cmdGetAutosaveDirectory_Click")
      Case 0: Resume
50130 Case 1: Resume Next
50140 Case 2: Exit Sub
50150 Case 3: End
50160 End Select
50170 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdGetgsbinDirectory_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim strFolder As String, aw As Long
50050  strFolder = BrowseForFolder(Me.hwnd, LanguageStrings.OptionsGhostscriptBinariesDirectoryPrompt)
50060  If Len(strFolder) = 0 Then
50070   Exit Sub
50080  End If
50090  strFolder = CompletePath(strFolder)
50100  If FileExists(strFolder & GsDll) = False Then
50110   MsgBox LanguageStrings.MessagesMsg15
50120   Exit Sub
50130  End If
50140  If UCase$(CompletePath(Options.DirectoryGhostscriptBinaries)) <> UCase$(CompletePath(strFolder)) Then
50150   aw = MsgBox("The program must be restarted!", vbOKCancel)
50160   If aw = vbCancel Then
50170    Exit Sub
50180   End If
50190   txtGSbin.Text = strFolder
50200   GetOptions Me, Options
50210   SaveOptions Options
50220   Restart = True
50230   Unload Me
50240  End If
50250  With txtGSbin
50260   .Text = strFolder
50270   .ToolTipText = .Text
50280  End With
50290 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50300 Exit Sub
ErrPtnr_OnError:
50321 Select Case ErrPtnr.OnError("frmOptions", "cmdGetgsbinDirectory_Click")
      Case 0: Resume
50340 Case 1: Resume Next
50350 Case 2: Exit Sub
50360 Case 3: End
50370 End Select
50380 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdGetgsfontsDirectory_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim strFolder As String
50050  strFolder = BrowseForFolder(Me.hwnd, LanguageStrings.OptionsGhostscriptFontsDirectoryPrompt)
50060  If Len(strFolder) = 0 Then Exit Sub
50070  strFolder = CompletePath(strFolder)
50080  If Len(Dir(strFolder & "*.afm", vbNormal)) = 0 And Len(Dir(strFolder & "*.pfb", vbNormal)) = 0 Then
50090   MsgBox LanguageStrings.MessagesMsg16
50100   Exit Sub
50110  End If
50120  txtGSfonts.Text = strFolder
50130  With txtGSfonts
50140   .ToolTipText = .Text
50150  End With
50160 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50170 Exit Sub
ErrPtnr_OnError:
50191 Select Case ErrPtnr.OnError("frmOptions", "cmdGetgsfontsDirectory_Click")
      Case 0: Resume
50210 Case 1: Resume Next
50220 Case 2: Exit Sub
50230 Case 3: End
50240 End Select
50250 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdGetgslibDirectory_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim strFolder As String
50050  strFolder = BrowseForFolder(Me.hwnd, LanguageStrings.OptionsGhostscriptLibrariesDirectoryPrompt)
50060  If Len(strFolder) = 0 Then Exit Sub
50070  strFolder = CompletePath(strFolder)
50080  If Len(Dir(strFolder & "*.*", vbNormal)) = 0 Then
50090   MsgBox LanguageStrings.MessagesMsg17
50100   Exit Sub
50110  End If
50120  With txtGSlib
50130   .Text = strFolder
50140   .ToolTipText = .Text
50150  End With
50160 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50170 Exit Sub
ErrPtnr_OnError:
50191 Select Case ErrPtnr.OnError("frmOptions", "cmdGetgslibDirectory_Click")
      Case 0: Resume
50210 Case 1: Resume Next
50220 Case 2: Exit Sub
50230 Case 3: End
50240 End Select
50250 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdGetgsresourceDirectory_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim strFolder As String
50050  strFolder = BrowseForFolder(Me.hwnd, LanguageStrings.OptionsGhostscriptResourceDirectoryPrompt)
50060  If Len(strFolder) = 0 Then Exit Sub
50070  strFolder = CompletePath(strFolder)
50080  With txtGSresource
50090   .Text = strFolder
50100   .ToolTipText = .Text
50110  End With
50120 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50130 Exit Sub
ErrPtnr_OnError:
50151 Select Case ErrPtnr.OnError("frmOptions", "cmdGetgsresourceDirectory_Click")
      Case 0: Resume
50170 Case 1: Resume Next
50180 Case 2: Exit Sub
50190 Case 3: End
50200 End Select
50210 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdGetTemppath_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim strFolder As String
50050  strFolder = BrowseForFolder(Me.hwnd, LanguageStrings.OptionsPrintertempDirectoryPrompt)
50060  If Len(strFolder) = 0 Then Exit Sub
50070  strFolder = CompletePath(strFolder)
50080  With txtTemppath
50090   .Text = strFolder
50100   .ToolTipText = .Text
50110  End With
50120 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50130 Exit Sub
ErrPtnr_OnError:
50151 Select Case ErrPtnr.OnError("frmOptions", "cmdGetTemppath_Click")
      Case 0: Resume
50170 Case 1: Resume Next
50180 Case 2: Exit Sub
50190 Case 3: End
50200 End Select
50210 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdReset_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim res As Long, Options As tOptions
50050  res = MsgBox(LanguageStrings.MessagesMsg03, vbYesNo)
50060  If res = vbYes Then
50070   Options = StandardOptions
50080   ShowOptions Me, Options
50090   With Options
50100    SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50110    cmbCharset.Text = .ProgramFontCharset
50120    SetFont frmMain, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50130   End With
50140  End If
50150 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50160 Exit Sub
ErrPtnr_OnError:
50181 Select Case ErrPtnr.OnError("frmOptions", "cmdReset_Click")
      Case 0: Resume
50200 Case 1: Resume Next
50210 Case 2: Exit Sub
50220 Case 3: End
50230 End Select
50240 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdSave_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim tRestart As Boolean
50050  tRestart = False
50060  If UCase$(Options.DirectoryGhostscriptBinaries) <> UCase$(txtGSbin.Text) Then
50070   tRestart = True
50080  End If
50090  CorrectCmbCharset
50100  GetOptions Me, Options
50110  SaveOptions Options
50120  If IsWin9xMe = False Then
50131   Select Case Options.ProcessPriority
               Case 0: 'Idle
50150     SetProcessPriority Idle
50160    Case 1: 'Normal
50170     SetProcessPriority Normal
50180    Case 2: 'High
50190     SetProcessPriority High
50200    Case 3: 'Realtime
50210     SetProcessPriority RealTime
50220   End Select
50230  End If
50240  If tRestart = True Then
50250   Restart = True
50260  End If
50270  Unload Me
50280 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50290 Exit Sub
ErrPtnr_OnError:
50311 Select Case ErrPtnr.OnError("frmOptions", "cmdSave_Click")
      Case 0: Resume
50330 Case 1: Resume Next
50340 Case 2: Exit Sub
50350 Case 3: End
50360 End Select
50370 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdShellintegration_Click(Index As Integer)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  MousePointer = vbHourglass
50050  cmdShellintegration(0).Enabled = False
50060  cmdShellintegration(1).Enabled = False
50071  Select Case Index
              Case 0
50090    AddExplorerIntegration
50100   Case 1
50110    RemoveExplorerIntegration
50120  End Select
50130  MousePointer = vbNormal
50140  cmdShellintegration(0).Enabled = True
50150  cmdShellintegration(1).Enabled = True
50160 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50170 Exit Sub
ErrPtnr_OnError:
50191 Select Case ErrPtnr.OnError("frmOptions", "cmdShellintegration_Click")
      Case 0: Resume
50210 Case 1: Resume Next
50220 Case 2: Exit Sub
50230 Case 3: End
50240 End Select
50250 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdTest_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim tCharset As Long, tStr As String, tFontSize As Long, tFontname As String, _
  tFontCharset As Long
50060  If InStr(1, cmbCharset.Text, ",", vbTextCompare) > 0 Then
50070    tStr = Trim$(Mid$(cmbCharset.Text, 1, InStr(1, cmbCharset.Text, ",", vbTextCompare) - 1))
50080   Else
50090    tStr = Trim$(cmbCharset.Text)
50100  End If
50110  If Len(tStr) = 0 Then
50120   cmbCharset.Text = 0
50130   Exit Sub
50140  End If
50150  If IsNumeric(tStr) = False Then
50160   cmbCharset.Text = 0
50170   Exit Sub
50180  End If
50190  tCharset = tStr
50200  With cmdTest
50210   tFontname = .Fontname
50220   tFontSize = .Fontsize
50230   tFontCharset = .Font.Charset
50240  End With
50250  SetFont Me, cmbFonts.List(cmbFonts.ListIndex), CLng(tStr), cmbProgramFontsize.Text
50260  cmbCharset.Text = tCharset
50270  SetFont frmMain, cmbFonts.List(cmbFonts.ListIndex), CLng(tStr), cmbProgramFontsize.Text
50280  With cmdTest
50290   .Fontname = tFontname
50300   .Fontsize = tFontSize
50310   .Font.Charset = tFontCharset
50320  End With
50330  With cmdCancelTest
50340   .Fontname = tFontname
50350   .Fontsize = tFontSize
50360   .Font.Charset = tFontCharset
50370   .Enabled = True
50380  End With
50390 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50400 Exit Sub
ErrPtnr_OnError:
50421 Select Case ErrPtnr.OnError("frmOptions", "cmdTest_Click")
      Case 0: Resume
50440 Case 1: Resume Next
50450 Case 2: Exit Sub
50460 Case 3: End
50470 End Select
50480 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdTestpage_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim TestPSPage As String, fn As Long, Filename As String, tStr As String
50050  TestPSPage = LoadResString(3000)
50060  TestPSPage = Replace(TestPSPage, "[TESTPAGE]", LanguageStrings.OptionsTestpage, , 1, vbTextCompare)
50070  TestPSPage = Replace(TestPSPage, "[DATE]", Now, , 1, vbTextCompare)
50080  TestPSPage = Replace(TestPSPage, "[PDFCREATORVERSION]", App.Title & " " & GetProgramReleaseStr, , 1, vbTextCompare)
50090
50100  fn = FreeFile
50110  tStr = CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory & "\" & GetUsername
50120  If DirExists(tStr) = False Then
50130   MakePath tStr
50140  End If
50150  Filename = GetTempFile(tStr, "~PD")
50160  Open Filename For Output As fn
50170  Print #fn, TestPSPage
50180  Close #fn
50190  frmMain.CheckPrintJobs
50200 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50210 Exit Sub
ErrPtnr_OnError:
50231 Select Case ErrPtnr.OnError("frmOptions", "cmdTestpage_Click")
      Case 0: Resume
50250 Case 1: Resume Next
50260 Case 2: Exit Sub
50270 Case 3: End
50280 End Select
50290 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdUsertempPath_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim Temppath As String
50050  Temppath = CompletePath(GetTempPath)
50060  If DirExists(Temppath) = False Then
50070   MakePath Temppath
50080  End If
50090  With txtTemppath
50100   .Text = Temppath
50110   .ToolTipText = .Text
50120  End With
50130 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50140 Exit Sub
ErrPtnr_OnError:
50161 Select Case ErrPtnr.OnError("frmOptions", "cmdUsertempPath_Click")
      Case 0: Resume
50180 Case 1: Resume Next
50190 Case 2: Exit Sub
50200 Case 3: End
50210 End Select
50220 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If KeyCode = vbKeyF1 Then
50050   KeyCode = 0
50061   Select Case trv.SelectedItem.Key
               Case "Program"
50080     Call HTMLHelp_ShowTopic("html\generalsettings.htm")
50090    Case "ProgramGeneral"
50100     Call HTMLHelp_ShowTopic("html\generalsettings.htm")
50110    Case "ProgramSave"
50120     Call HTMLHelp_ShowTopic("html\savesettings.htm")
50130    Case "ProgramAutosave"
50140     Call HTMLHelp_ShowTopic("html\autosave.htm")
50150    Case "ProgramFonts"
50160     Call HTMLHelp_ShowTopic("html\fontsetting.htm")
50170    Case "ProgramDirectories"
50180     Call HTMLHelp_ShowTopic("html\directories.htm")
50190    Case "ProgramDocument"
50200     Call HTMLHelp_ShowTopic("html\docproperties.htm")
50210    Case "Formats"
50220     Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
50230    Case "FormatsPDF"
50241     Select Case tbstrPDFOptions.SelectedItem.Index
                 Case 1:
50260       Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
50270      Case 2:
50280       Call HTMLHelp_ShowTopic("html\pdfcompression.htm")
50290      Case 3:
50300       Call HTMLHelp_ShowTopic("html\pdffonts.htm")
50310      Case 4:
50320       Call HTMLHelp_ShowTopic("html\pdfcolors.htm")
50330      Case 5:
50340       Call HTMLHelp_ShowTopic("html\pdfsecurity.htm")
50350     End Select
50360    Case "FormatsPNG"
50370     Call HTMLHelp_ShowTopic("html\pngsettings.htm")
50380    Case "FormatsJPEG"
50390     Call HTMLHelp_ShowTopic("html\jpegsettings.htm")
50400    Case "FormatsBMP"
50410     Call HTMLHelp_ShowTopic("html\bmpsettings.htm")
50420    Case "FormatsPCX"
50430     Call HTMLHelp_ShowTopic("html\pcxsettings.htm")
50440    Case "FormatsTIFF"
50450     Call HTMLHelp_ShowTopic("html\tiffsettings.htm")
50460    Case "FormatsPS"
50470     Call HTMLHelp_ShowTopic("html\pssettings.htm")
50480    Case "FormatsEPS"
50490     Call HTMLHelp_ShowTopic("html\epssettings.htm")
50500   End Select
50510  End If
50520 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50530 Exit Sub
ErrPtnr_OnError:
50551 Select Case ErrPtnr.OnError("frmOptions", "Form_KeyDown")
      Case 0: Resume
50570 Case 1: Resume Next
50580 Case 2: Exit Sub
50590 Case 3: End
50600 End Select
50610 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Load()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Const fraPDFTop = 1360, fraPDFLeft = 2960
50050  Dim pic As New StdPicture, i As Long, tStr As String, gsvers As Collection, _
  cSystem As clsSystem, fi As Long, fc As Long, SMF As Collection, _
  reg As clsRegistry, tsf() As String, tStr2 As String, ctl As Control
50080
50090  KeyPreview = True
50100
50110
50120  Set cSystem = New clsSystem
50130  Set SMF = cSystem.GetSystemFont(Me, Menu)
50140
50150  With Screen
50160   .MousePointer = vbHourglass
50170   Move (.Width - Width) / 2, (.Height - Height) / 2
50180  End With
50190
50200  fraBitmapGeneral.Visible = False
50210  fraBitmapGeneral.Top = fraPDFTop - 300
50220  fraBitmapGeneral.Left = fraPDFLeft - 200
50230  fraProgFont.Top = fraPDFTop - 300
50240  fraProgFont.Left = fraPDFLeft - 200
50250  fraProgGeneral.Top = fraPDFTop - 300
50260  fraProgGeneral.Left = fraPDFLeft - 200
50270  fraProgAutosave.Top = fraPDFTop - 300
50280  fraProgAutosave.Left = fraPDFLeft - 200
50290  fraProgSave.Top = fraPDFTop - 300
50300  fraProgSave.Left = fraPDFLeft - 200
50310  fraProgDirectories.Top = fraPDFTop - 300
50320  fraProgDirectories.Left = fraPDFLeft - 200
50330  fraProgDocument.Top = fraPDFTop - 300
50340  fraProgDocument.Left = fraPDFLeft - 200
50350  fraProgGhostscript.Top = fraPDFTop - 300
50360  fraProgGhostscript.Left = fraPDFLeft - 200
50370
50380  fraPDFSecurity.Top = fraPDFTop + 100
50390  fraPDFSecurity.Left = fraPDFLeft
50400  fraPDFFonts.Top = fraPDFTop + 100
50410  fraPDFFonts.Left = fraPDFLeft
50420  fraPDFColors.Top = fraPDFTop + 100
50430  fraPDFColors.Left = fraPDFLeft
50440  fraPDFGeneral.Top = fraPDFTop + 100
50450  fraPDFGeneral.Left = fraPDFLeft
50460  fraPDFCompress.Top = fraPDFTop + 100
50470  fraPDFCompress.Left = fraPDFLeft
50480  fraPSGeneral.Left = fraPDFLeft
50490  fraPSGeneral.Top = fraPDFTop - 300
50500  fraPSGeneral.Left = fraPDFLeft - 200
50510  tbstrPDFOptions.Top = fraPDFTop - 300
50520  tbstrPDFOptions.Left = fraPDFLeft - 200
50530  tbstrPDFOptions.Height = 5975
50540  tbstrPDFOptions.Width = 6215
50550
50560  cmbEPSLanguageLevel.Top = cmbPSLanguageLevel.Top
50570  cmbEPSLanguageLevel.Left = cmbPSLanguageLevel.Left
50580
50590  txtTest.Text = vbNullString
50600  For i = 33 To 255
50610   txtTest.Text = txtTest.Text & Chr$(i)
50620  Next i
50630  fi = -1
50640  With cmbFonts
50650   .Clear
50660   For i = 1 To Screen.FontCount
50670    tStr = Trim$(Screen.Fonts(i))
50680    If Len(tStr) > 0 Then
50690     cmbFonts.AddItem tStr
50700    End If
50710   Next i
50720   If .ListCount > 0 Then
50730     For i = 0 To cmbFonts.ListCount - 1
50740      If SMF.Count > 0 Then
50750       If UCase$(cmbFonts.List(i)) = UCase$(SMF(1)(0)) Then
50760        fi = i
50770       End If
50780      End If
50790     Next i
50800    Else
50810    .ListIndex = 0
50820   End If
50830  End With
50840  With cmbCharset
50850   .Clear
50860   .AddItem "0, Western": .ItemData(.NewIndex) = 0
50870   .AddItem "2, Symbol": .ItemData(.NewIndex) = 2
50880   .AddItem "77, Mac": .ItemData(.NewIndex) = 77
50890   .AddItem "161, Greek": .ItemData(.NewIndex) = 161
50900   .AddItem "162, Turkish": .ItemData(.NewIndex) = 162
50910   .AddItem "177, Hebrew": .ItemData(.NewIndex) = 177
50920   .AddItem "178, Arabic": .ItemData(.NewIndex) = 178
50930   .AddItem "186, Baltic": .ItemData(.NewIndex) = 186
50940   .AddItem "204, Cyrillic": .ItemData(.NewIndex) = 204
50950   .AddItem "238, Central European": .ItemData(.NewIndex) = 238
50960   .AddItem "255, DOS/OEM": .ItemData(.NewIndex) = 255
50970   .Text = 0
50980  End With
50990  If fi >= 0 Then
51000   cmbFonts.ListIndex = fi
51010   cmbCharset.Text = SMF(1)(2)
51020   cmbProgramFontsize.Text = SMF(1)(1)
51030   txtTest.Font = cmbFonts.List(cmbFonts.ListIndex)
51040   txtTest.Font.Charset = cmbCharset.Text
51050  End If
51060
51070
51080  trv.Nodes.Clear
51090  trv.Indentation = 200
51100  With LanguageStrings
51110   trv.Nodes.Add , , "Program", .OptionsTreeProgram
51120   trv.Nodes.Add "Program", tvwChild, "ProgramGeneral", .OptionsProgramGeneralSymbol
51130   trv.Nodes.Add "Program", tvwChild, "ProgramGhostscript", .OptionsProgramGhostscriptSymbol
51140   trv.Nodes.Add "Program", tvwChild, "ProgramDocument", .OptionsProgramDocumentSymbol
51150   trv.Nodes.Add "Program", tvwChild, "ProgramSave", .OptionsProgramSaveSymbol
51160   trv.Nodes.Add "Program", tvwChild, "ProgramAutosave", .OptionsProgramAutosaveSymbol
51170   trv.Nodes.Add "Program", tvwChild, "ProgramDirectories", .OptionsProgramDirectoriesSymbol
51180   trv.Nodes.Add "Program", tvwChild, "ProgramFonts", .OptionsProgramFontSymbol
51190   trv.Nodes.Add , , "Formats", .OptionsTreeFormats
51200   trv.Nodes.Add "Formats", tvwChild, "FormatsPDF", .OptionsPDFSymbol
51210   trv.Nodes.Add "Formats", tvwChild, "FormatsPNG", .OptionsPNGSymbol
51220   trv.Nodes.Add "Formats", tvwChild, "FormatsJPEG", .OptionsJPEGSymbol
51230   trv.Nodes.Add "Formats", tvwChild, "FormatsBMP", .OptionsBMPSymbol
51240   trv.Nodes.Add "Formats", tvwChild, "FormatsPCX", .OptionsPCXSymbol
51250   trv.Nodes.Add "Formats", tvwChild, "FormatsTIFF", .OptionsTIFFSymbol
51260   trv.Nodes.Add "Formats", tvwChild, "FormatsPS", .OptionsPSSymbol
51270   trv.Nodes.Add "Formats", tvwChild, "FormatsEPS", .OptionsEPSSymbol
51280
51290   trv.Nodes("ProgramFonts").EnsureVisible
51300   trv.Nodes("FormatsPDF").EnsureVisible
51310
51320   Set picOptions = LoadResPicture(2101, vbResIcon)
51330   fraProgFont.Visible = False
51340   fraProgGeneral.Visible = True
51350
51360   fraProgGeneral.Caption = .OptionsProgramGeneralSymbol
51370   fraProgGhostscript.Caption = .OptionsProgramGhostscriptSymbol
51380   fraProgAutosave.Caption = .OptionsProgramAutosaveSymbol
51390   fraProgFont.Caption = .OptionsProgramFontSymbol
51400   fraProgDirectories.Caption = .OptionsProgramDirectoriesSymbol
51410   fraProgSave.Caption = .OptionsProgramSaveSymbol
51420   fraProgDocument.Caption = .OptionsProgramDocumentSymbol
51430
51440   fraShellintegration.Caption = .OptionsShellIntegration
51450   cmdShellintegration(0).Caption = .OptionsShellIntegrationAdd
51460   cmdShellintegration(1).Caption = .OptionsShellIntegrationRemove
51470   If IsWin9xMe = False Then
51480    If IsAdmin = False Then
51490     cmdShellintegration(0).Enabled = False
51500     cmdShellintegration(1).Enabled = False
51510    End If
51520   End If
51530
51540   lblGhostscriptversion.Caption = .OptionsGhostscriptversion
51550
51560   lblSaveFilename.Caption = .OptionsSaveFilename
51570   lblSaveFilenameTokens.Caption = .OptionsSaveFilenameTokens
51580   fraFilenameSubstitutions.Caption = .OptionsSaveFilenameSubstitutions
51590   chkFilenameSubst.Caption = .OptionsSaveFilenameSubstitutionsTitle
51600   cmdFilenameSubst(0).Caption = .OptionsSaveFilenameAdd
51610   cmdFilenameSubst(1).Caption = .OptionsSaveFilenameChange
51620   cmdFilenameSubst(2).Caption = .OptionsSaveFilenameDelete
51630
51640   chkSpaces.Caption = .OptionsRemoveSpaces
51650   chkNoConfirmMessageSwitchingDefaultprinter.Caption = .OptionsProgramSwitchingDefaultprinter
51660   lblGSbin.Caption = .OptionsDirectoriesGSBin
51670   lblGSlib.Caption = .OptionsDirectoriesGSLibraries
51680   lblGSfonts.Caption = .OptionsDirectoriesGSFonts
51690   lblPrintTempPath.Caption = .OptionsDirectoriesTempPath
51700
51710   lblOptions = .OptionsProgramGeneralDescription
51720   lblAutosaveformat.Caption = .OptionsAutosaveFormat
51730   chkUseStandardAuthor.Caption = .OptionsUseStandardauthor
51740   chkUseCreationDateNow.Caption = .OptionsUseCreationDateNow
51750   chkUseAutosave.Caption = .OptionsUseAutosave
51760   cmdTestpage.Caption = .OptionsPrintTestpage
51770   lblAutosaveFilename.Caption = .OptionsAutosaveFilename
51780   lblAutosaveFilenameTokens.Caption = .OptionsAutosaveFilenameTokens
51790   chkUseAutosaveDirectory.Caption = .OptionsUseAutosaveDirectory
51800   lblAuthorTokens.Caption = .OptionsStandardAuthorToken
51810
51820   With cmbAutosaveFormat
51830    .AddItem "PDF"
51840    .AddItem "PNG"
51850    .AddItem "JPEG"
51860    .AddItem "BMP"
51870    .AddItem "PCX"
51880    .AddItem "TIFF"
51890    .AddItem "PS"
51900    .AddItem "EPS"
51910   End With
51920   With cmbSaveFilenameTokens
51930    .AddItem "<Author>"
51940    .AddItem "<Computername>"
51950    .AddItem "<DateTime>"
51960    .AddItem "<Title>"
51970    .AddItem "<Username>"
51980    .ListIndex = 0
51990   End With
52000   With cmbAuthorTokens
52010    .AddItem "<Computername>"
52020    .AddItem "<DateTime>"
52030    .AddItem "<Title>"
52040    .AddItem "<Username>"
52050    .ListIndex = 0
52060   End With
52070   With cmbAutoSaveFilenameTokens
52080    .AddItem "<Author>"
52090    .AddItem "<Computername>"
52100    .AddItem "<DateTime>"
52110    .AddItem "<Title>"
52120    .AddItem "<Username>"
52130    .ListIndex = 0
52140   End With
52150   Me.Caption = .DialogPrinterOptions
52160   cmdCancel.Caption = .OptionsCancel
52170   cmdReset.Caption = .OptionsReset
52180   cmdSave.Caption = .OptionsSave
52190   tbstrPDFOptions.Tabs.Clear
52200   tbstrPDFOptions.Tabs.Add , "General", .OptionsPDFGeneral
52210   tbstrPDFOptions.Tabs.Add , "Compression", .OptionsPDFCompression
52220   tbstrPDFOptions.Tabs.Add , "Fonts", .OptionsPDFFonts
52230   tbstrPDFOptions.Tabs.Add , "Colors", .OptionsPDFColors
52240   tbstrPDFOptions.Tabs.Add , "Security", .OptionsPDFSecurity
52250   fraPDFGeneral.Caption = .OptionsPDFGeneralCaption
52260   lblPDFCompat.Caption = .OptionsPDFGeneralCompatibility
52270   lblPDFAutoRotate.Caption = .OptionsPDFGeneralAutorotate
52280   lblPDFResolution.Caption = .OptionsPDFGeneralResolution
52290   lblPDFOverprint.Caption = .OptionsPDFGeneralOverprint
52300   lblProgfont.Caption = .OptionsProgramFont
52310   lblProgcharset.Caption = .OptionsProgramFontcharset
52320   lblSize.Caption = .OptionsProgramFontSize
52330   lblTesttext = .OptionsProgramFontTestdescription
52340   cmdTest.Caption = .OptionsProgramFontTest
52350   cmdCancelTest.Caption = .OptionsProgramFontCancelTest
52360   chkPDFASCII85.Caption = .OptionsPDFGeneralASCII85
52370   cmbPDFCompat.Clear
52380   cmbPDFCompat.AddItem .OptionsPDFGeneralCompatibility01
52390   cmbPDFCompat.AddItem .OptionsPDFGeneralCompatibility02
52400   cmbPDFCompat.AddItem .OptionsPDFGeneralCompatibility03
52410   cmbPDFRotate.Clear
52420   cmbPDFRotate.AddItem .OptionsPDFGeneralRotate01
52430   cmbPDFRotate.AddItem .OptionsPDFGeneralRotate02
52440   cmbPDFRotate.AddItem .OptionsPDFGeneralRotate03
52450   cmbPDFOverprint.Clear
52460   cmbPDFOverprint.AddItem .OptionsPDFGeneralOverprint01
52470   cmbPDFOverprint.AddItem .OptionsPDFGeneralOverprint02
52480
52490   fraPDFCompress.Caption = .OptionsPDFCompressionCaption
52500   chkPDFTextComp.Caption = .OptionsPDFCompressionTextComp
52510   fraPDFColor.Caption = .OptionsPDFCompressionColor
52520   chkPDFColorComp.Caption = .OptionsPDFCompressionColorComp
52530   chkPDFColorResample.Caption = .OptionsPDFCompressionColorResample
52540   lblPDFColorRes.Caption = .OptionsPDFCompressionColorRes
52550   cmbPDFColorComp.Clear
52560   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp01
52570   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp02
52580   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp03
52590   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp04
52600   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp05
52610   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp06
52620   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp07
52630   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp08
52640   cmbPDFColorResample.Clear
52650   cmbPDFColorResample.AddItem .OptionsPDFCompressionColorResample01
52660   cmbPDFColorResample.AddItem .OptionsPDFCompressionColorResample02
52670   cmbPDFColorResample.AddItem .OptionsPDFCompressionColorResample03
52680   fraPDFGrey.Caption = .OptionsPDFCompressionGrey
52690   chkPDFGreyComp.Caption = .OptionsPDFCompressionGreyComp
52700   chkPDFGreyResample.Caption = .OptionsPDFCompressionGreyResample
52710   lblPDFGreyRes.Caption = .OptionsPDFCompressionGreyRes
52720   cmbPDFGreyComp.Clear
52730   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp01
52740   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp02
52750   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp03
52760   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp04
52770   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp05
52780   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp06
52790   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp07
52800   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp08
52810   cmbPDFGreyResample.Clear
52820   cmbPDFGreyResample.AddItem .OptionsPDFCompressionGreyResample01
52830   cmbPDFGreyResample.AddItem .OptionsPDFCompressionGreyResample02
52840   cmbPDFGreyResample.AddItem .OptionsPDFCompressionGreyResample03
52850   fraPDFMono.Caption = .OptionsPDFCompressionMono
52860   chkPDFMonoComp.Caption = .OptionsPDFCompressionMonoComp
52870   chkPDFMonoResample.Caption = .OptionsPDFCompressionMonoResample
52880   lblPDFMonoRes.Caption = .OptionsPDFCompressionMonoRes
52890   cmbPDFMonoComp.Clear
52900   cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp01
52910   cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp02
52920   cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp03
52930   cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp04
52940   cmbPDFMonoResample.Clear
52950   cmbPDFMonoResample.AddItem .OptionsPDFCompressionMonoResample01
52960   cmbPDFMonoResample.AddItem .OptionsPDFCompressionMonoResample02
52970   cmbPDFMonoResample.AddItem .OptionsPDFCompressionMonoResample03
52980
52990   fraPDFFonts.Caption = .OptionsPDFFontsCaption
53000   chkPDFEmbedAll.Caption = .OptionsPDFFontsEmbedAll
53010   chkPDFSubSetFonts.Caption = .OptionsPDFFontsSubSetFonts
53020
53030   fraPDFColors.Caption = .OptionsPDFColorsCaption
53040   chkPDFCMYKtoRGB.Caption = .OptionsPDFColorsCMYKtoRGB
53050   fraPDFColorOptions.Caption = .OptionsPDFColorsColorOptions
53060   chkPDFPreserveOverprint.Caption = .OptionsPDFColorsPreserveOverprint
53070   chkPDFPreserveTransfer.Caption = .OptionsPDFColorsPreserveTransfer
53080   chkPDFPreserveHalftone.Caption = .OptionsPDFColorsPreserveHalftone
53090   cmbPDFColorModel.Clear
53100   cmbPDFColorModel.AddItem .OptionsPDFColorsColorModel01
53110   cmbPDFColorModel.AddItem .OptionsPDFColorsColorModel02
53120   cmbPDFColorModel.AddItem .OptionsPDFColorsColorModel03
53130
53140   fraPDFEncryptor.Caption = .OptionsPDFEncryptor
53150   fraPDFSecurity.Caption = .OptionsPDFSecurityCaption
53160   chkUseSecurity.Caption = .OptionsPDFUseSecurity
53170   fraPDFEncLevel.Caption = .OptionsPDFEncryptionLevel
53180   optEncHigh.Caption = .OptionsPDFEncryptionHigh
53190   optEncLow.Caption = .OptionsPDFEncryptionLow
53200   fraSecurityPass.Caption = .OptionsPDFPasswords
53210   chkUserPass.Caption = .OptionsPDFUserPass
53220   chkOwnerPass.Caption = .OptionsPDFOwnerPass
53230   fraPDFPermissions.Caption = .OptionsPDFDisallowUser
53240   fraPDFHighPermissions.Caption = .OptionsPDFEnhancedPermissions
53250   chkAllowPrinting.Caption = .OptionsPDFDisallowPrint
53260   chkAllowModifyContents.Caption = .OptionsPDFDisallowModify
53270   chkAllowCopy.Caption = .OptionsPDFDisallowCopy
53280   chkAllowModifyAnnotations.Caption = .OptionsPDFDisallowModifyComments
53290   chkAllowDegradedPrinting.Caption = .OptionsPDFAllowDegradedPrinting
53300   chkAllowFillIn.Caption = .OptionsPDFAllowFillIn
53310   chkAllowAssembly.Caption = .OptionsPDFAllowAssembly
53320   chkAllowScreenReaders.Caption = .OptionsPDFAllowScreenReaders
53330
53340   cmbPNGColors.AddItem .OptionsPNGColorscount01
53350   cmbPNGColors.AddItem .OptionsPNGColorscount02
53360   cmbPNGColors.AddItem .OptionsPNGColorscount03
53370   cmbPNGColors.AddItem .OptionsPNGColorscount04
53380   cmbJPEGColors.Left = cmbPNGColors.Left
53390   cmbJPEGColors.Width = cmbPNGColors.Width
53400   cmbJPEGColors.Top = cmbPNGColors.Top
53410   cmbJPEGColors.AddItem .OptionsJPEGColorscount01
53420   cmbJPEGColors.AddItem .OptionsJPEGColorscount02
53430   cmbBMPColors.Left = cmbPNGColors.Left
53440   cmbBMPColors.Width = cmbPNGColors.Width
53450   cmbBMPColors.Top = cmbPNGColors.Top
53460   cmbBMPColors.AddItem .OptionsBMPColorscount01
53470   cmbBMPColors.AddItem .OptionsBMPColorscount02
53480   cmbBMPColors.AddItem .OptionsBMPColorscount03
53490   cmbBMPColors.AddItem .OptionsBMPColorscount04
53500   cmbBMPColors.AddItem .OptionsBMPColorscount05
53510   cmbBMPColors.AddItem .OptionsBMPColorscount06
53520   cmbBMPColors.AddItem .OptionsBMPColorscount07
53530   cmbPCXColors.Left = cmbPNGColors.Left
53540   cmbPCXColors.Width = cmbPNGColors.Width
53550   cmbPCXColors.Top = cmbPNGColors.Top
53560   cmbPCXColors.AddItem .OptionsPCXColorscount01
53570   cmbPCXColors.AddItem .OptionsPCXColorscount02
53580   cmbPCXColors.AddItem .OptionsPCXColorscount03
53590   cmbPCXColors.AddItem .OptionsPCXColorscount04
53600   cmbPCXColors.AddItem .OptionsPCXColorscount05
53610   cmbPCXColors.AddItem .OptionsPCXColorscount06
53620   cmbTIFFColors.Left = cmbPNGColors.Left
53630   cmbTIFFColors.Width = cmbPNGColors.Width
53640   cmbTIFFColors.Top = cmbPNGColors.Top
53650   cmbTIFFColors.AddItem .OptionsTIFFColorscount01
53660   cmbTIFFColors.AddItem .OptionsTIFFColorscount02
53670   cmbTIFFColors.AddItem .OptionsTIFFColorscount03
53680   cmbTIFFColors.AddItem .OptionsTIFFColorscount04
53690   cmbTIFFColors.AddItem .OptionsTIFFColorscount05
53700   cmbTIFFColors.AddItem .OptionsTIFFColorscount06
53710   cmbTIFFColors.AddItem .OptionsTIFFColorscount07
53720   cmbTIFFColors.AddItem .OptionsTIFFColorscount08
53730
53740   fraBitmapGeneral.Caption = .OptionsImageSettings
53750   lblBitmapResolution = .OptionsBitmapResolution
53760   lblJPEGQuality = .OptionsJPEGQuality
53770   lblBitmapColors = .OptionsPDFColors
53780   lblProcessPriority.Caption = .OptionsProcesspriority
53790   lblLangLevel.Caption = .OptionsPSLanguageLevel
53800
53810   cmdAsso.Caption = .OptionsAssociatePSFiles
53820  End With
53830
53840  If IsPsAssociate = False Then
53850    cmdAsso.Enabled = True
53860   Else
53870    cmdAsso.Enabled = False
53880  End If
53890
53900  txtPDFRes.Text = 600
53910  cmbPDFCompat.ListIndex = 1
53920  cmbPDFRotate.ListIndex = 0
53930  cmbPDFOverprint.ListIndex = 0
53940  chkPDFASCII85.Value = 0
53950
53960  chkPDFTextComp.Value = 1
53970
53980  chkPDFColorComp.Value = 1
53990  chkPDFColorResample.Value = 0
54000  cmbPDFColorComp.ListIndex = 0
54010  cmbPDFColorResample.ListIndex = 0
54020  txtPDFColorRes.Text = 300
54030
54040  chkPDFGreyComp.Value = 1
54050  chkPDFGreyResample.Value = 0
54060  cmbPDFGreyComp.ListIndex = 0
54070  cmbPDFGreyResample.ListIndex = 0
54080  txtPDFGreyRes.Text = 300
54090
54100  chkPDFMonoComp.Value = 1
54110  chkPDFMonoResample.Value = 0
54120  cmbPDFMonoComp.ListIndex = 0
54130  cmbPDFMonoResample.ListIndex = 0
54140  txtPDFMonoRes.Text = 1200
54150
54160  chkPDFEmbedAll.Value = 1
54170  chkPDFSubSetFonts.Value = 1
54180  txtPDFSubSetPerc.Text = 100
54190
54200  cmbPDFColorModel.ListIndex = 1
54210  chkPDFCMYKtoRGB.Value = 1
54220  chkPDFPreserveOverprint.Value = 1
54230  chkPDFPreserveTransfer.Value = 1
54240  chkPDFPreserveHalftone.Value = 0
54250
54260  cmbPNGColors.ListIndex = 0
54270  cmbJPEGColors.ListIndex = 0
54280  cmbBMPColors.ListIndex = 0
54290  cmbPCXColors.ListIndex = 0
54300  cmbTIFFColors.ListIndex = 0
54310  txtBitmapResolution.Text = 150
54320
54330  cmbCharset.Text = cmbCharset.ItemData(0)
54340  cmbProgramFontsize.Text = 8
54350
54360 ' chkUseStandardAuthor.Value = 1
54370  txtStandardAuthor.Text = vbNullString
54380
54390  With cmbPSLanguageLevel
54400   .AddItem "1"
54410   .AddItem "1.5"
54420   .AddItem "2"
54430   .AddItem "3"
54440  End With
54450  With cmbEPSLanguageLevel
54460   .AddItem "1"
54470   .AddItem "1.5"
54480   .AddItem "2"
54490   .AddItem "3"
54500  End With
54510
54520  With lsvFilenameSubst
54530   .Appearance = ccFlat
54540   .ColumnHeaders.Clear
54550   .ColumnHeaders.Add , "Str1", "", lsvFilenameSubst.Width / 2 - 140
54560   .ColumnHeaders.Add , "Str2", "", lsvFilenameSubst.Width / 2 - 140
54570   .HideColumnHeaders = True
54580   .GridLines = True
54590   .FullRowSelect = True
54600   .HideSelection = False
54610  End With
54620
54630  With cmbPDFEncryptor
54640   .Clear
54650   .AddItem "Ghostscript (>= 8.14)"
54660   .ItemData(.NewIndex) = 0
54670   .AddItem "PDFEnc"
54680   .ItemData(.NewIndex) = 1
54690
54700   ShowOptions Me, Options
54710
54720   SecurityIsPossible = True
54730
54740   If FileExists(CompletePath(App.Path) & "pdfenc.exe") = False Then
54750    .RemoveItem 1
54760    .ListIndex = 0
54770    Options.PDFEncryptor = .ItemData(.ListIndex)
54780   End If
54790   If GhostScriptSecurity = False Then
54800    .RemoveItem 0
54810   End If
54820   If .ListCount = 0 Then
54830     chkUseSecurity.Value = 0
54840     chkUseSecurity.Enabled = False
54850     SecurityIsPossible = False
54860    Else
54870     For i = 0 To .ListCount - 1
54880      If .ItemData(i) = Options.PDFEncryptor Then
54890       .ListIndex = i
54900       Exit For
54910      End If
54920     Next i
54930     If .ListIndex = -1 Then
54940      .ListIndex = 0
54950      Options.PDFEncryptor = .ItemData(.ListIndex)
54960     End If
54970   End If
54980  End With
54990
55000
55010  If Options.PDFHighEncryption <> 0 Then
55020    optEncHigh.Value = True
55030   Else
55040    optEncLow.Value = True
55050  End If
55060
55070  CheckCmdFilenameSubst
55080
55090  If chkUseStandardAuthor.Value = 1 Then
55100    txtStandardAuthor.Enabled = True
55110    txtStandardAuthor.BackColor = &H80000005
55120   Else
55130    txtStandardAuthor.Enabled = False
55140    txtStandardAuthor.BackColor = &H8000000F
55150  End If
55160  With Options
55170   SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
55180   cmbCharset.Text = .ProgramFontCharset
55190  End With
55200  If chkUseAutosave.Value = 1 Then
55210    ViewAutosave True
55220   Else
55230    ViewAutosave False
55240  End If
55250
55260  With txtGSbin
55270   .ToolTipText = .Text
55280  End With
55290  With txtGSlib
55300   .ToolTipText = .Text
55310  End With
55320  With txtGSfonts
55330   .ToolTipText = .Text
55340  End With
55350  With txtTemppath
55360   .ToolTipText = .Text
55370  End With
55380
55390  With sldProcessPriority
55400   .TextPosition = sldBelowRight
55410   .TickFrequency = 1
55420   .TickStyle = sldTopLeft
55431   Select Case .Value
         Case 0: 'Idle
55450     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityIdle
55460    Case 1: 'Normal
55470     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityNormal
55480    Case 2: 'High
55490     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityHigh
55500    Case 3: 'Realtime
55510     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityRealtime
55520   End Select
55530  End With
55540
55550  If IsWin9xMe = False Then
55560    lblProcessPriority.Enabled = True
55570    sldProcessPriority.Enabled = True
55580   Else
55590    lblProcessPriority.Enabled = False
55600    sldProcessPriority.Enabled = False
55610  End If
55620  UpdateSecurityFields
55630
55640  tStr2 = CompletePath(UCase$(Trim$(Options.DirectoryGhostscriptBinaries)))
 cmbGhostscript.Clear: Set reg = New clsRegistry
55660  reg.hkey = HKEY_LOCAL_MACHINE
55670
55680  Set gsvers = GetAllGhostscriptversions
55690
55700  If gsvers.Count = 0 Then
55710    cmbGhostscript.Enabled = False
55720   Else
55730    For i = 1 To gsvers.Count
55740     cmbGhostscript.AddItem gsvers.item(i)
55750    Next i
55760    cmbGhostscript.ListIndex = cmbGhostscript.ListCount - 1
55770    For i = 0 To cmbGhostscript.ListCount - 1
55780     tStr = ""
55790     If InStr(cmbGhostscript.List(i), ":") Then
55800       reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
55810       If tStr2 = CompletePath(UCase$(Trim$(reg.GetRegistryValue("GhostscriptDirectoryBinaries")))) Then
55820        cmbGhostscript.ListIndex = i
55830        Exit For
55840       End If
55850      Else
55860       If InStr(UCase$(cmbGhostscript.List(i)), "AFPL") Then
55870        reg.KeyRoot = "SOFTWARE\AFPL Ghostscript"
55880        If InStr(cmbGhostscript.List(i), " ") > 0 Then
55890         tsf = Split(cmbGhostscript.List(i), " ")
55900         reg.Subkey = tsf(UBound(tsf))
55910         tStr = reg.GetRegistryValue("GS_DLL")
55920         If tStr2 & "GSDLL32.DLL" = UCase$(tStr) Then
55930          cmbGhostscript.ListIndex = i
55940          Exit For
55950         End If
55960        End If
55970       End If
55980       If InStr(UCase$(cmbGhostscript.List(i)), "GNU") Then
55990        reg.KeyRoot = "SOFTWARE\GNU Ghostscript"
56000        If InStr(cmbGhostscript.List(i), " ") > 0 Then
56010         tsf = Split(cmbGhostscript.List(i), " ")
56020         reg.Subkey = tsf(UBound(tsf))
56030         tStr = reg.GetRegistryValue("GS_DLL")
56040         If tStr2 & "GSDLL32.DLL" = UCase$(tStr) Then
56050          cmbGhostscript.ListIndex = i
56060          Exit For
56070         End If
56080        End If
56090       End If
56100       If InStr(UCase$(cmbGhostscript.List(i)), "GPL") Then
56110        reg.KeyRoot = "SOFTWARE\GPL Ghostscript"
56120        If InStr(cmbGhostscript.List(i), " ") > 0 Then
56130         tsf = Split(cmbGhostscript.List(i), " ")
56140         reg.Subkey = tsf(UBound(tsf))
56150         tStr = reg.GetRegistryValue("GS_DLL")
56160         If tStr2 & "GSDLL32.DLL" = UCase$(tStr) Then
56170          cmbGhostscript.ListIndex = i
56180          Exit For
56190         End If
56200        End If
56210       End If
56220     End If
56230    Next i
56240  End If
56250  Set reg = Nothing
56260  With cmbGhostscript
56270   If .ListCount = 0 Then
56280    .Enabled = False
56290    .BackColor = &H8000000F
56300   End If
56310  End With
56320
56330  With cmbProgramFontsize
56340   .AddItem "8"
56350   .AddItem "9"
56360   .AddItem "10"
56370   .AddItem "11"
56380   .AddItem "12"
56390   .AddItem "14"
56400   .AddItem "16"
56410   .AddItem "18"
56420   .AddItem "20"
56430   .AddItem "22"
56440   .AddItem "24"
56450   .AddItem "26"
56460   .AddItem "28"
56470   .AddItem "36"
56480   .AddItem "48"
56490   .AddItem "72"
56500  End With
56510
56520  For Each ctl In Controls
56530   If TypeOf ctl Is ComboBox Then
56540    ComboSetListWidth ctl
56550   End If
56560  Next ctl
56570
56580  SetOptimalComboboxHeigth cmbCharset, Me
56590  SetOptimalComboboxHeigth cmbProgramFontsize, Me
56600  SetOptimalComboboxHeigth cmbGhostscript, Me
56610  CorrectCmbCharset
56620  tbstrPDFOptions.ZOrder 1
56630  cmdStyle.ZOrder 1
56640  If ShowOnlyOptions = True Then
56650   FormInTaskbar Me, True, True
56660   Caption = "PDFCreator - " & Caption
56670  End If
56680  Screen.MousePointer = vbNormal
56690 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
56700 Exit Sub
ErrPtnr_OnError:
56721 Select Case ErrPtnr.OnError("frmOptions", "Form_Load")
      Case 0: Resume
56740 Case 1: Resume Next
56750 Case 2: Exit Sub
56760 Case 3: End
56770 End Select
56780 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Resize()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  With cmbCharset
50050   .Top = cmbFonts.Top
50060   .Left = lblProgcharset.Left
50070   .Width = 2295
50080   .SelStart = 0
50090   .SelLength = 0
50100  End With
50110  With cmbProgramFontsize
50120   .Top = cmbFonts.Top
50130   .Left = lblSize.Left
50140   .Width = 765
50150   .SelStart = 0
50160   .SelLength = 0
50170  End With
50180  With cmbGhostscript
50190   .Top = lblGhostscriptversion.Top + lblGhostscriptversion.Height + 20
50200   .Left = lblGhostscriptversion.Left
50210   .Width = 4215
50220  End With
50230 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50240 Exit Sub
ErrPtnr_OnError:
50261 Select Case ErrPtnr.OnError("frmOptions", "Form_Resize")
      Case 0: Resume
50280 Case 1: Resume Next
50290 Case 2: Exit Sub
50300 Case 3: End
50310 End Select
50320 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ShowProgOptions()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  lblJPEGQuality.Visible = False
50050  cmbPNGColors.Visible = False
50060  cmbJPEGColors.Visible = False
50070  cmbBMPColors.Visible = False
50080  cmbPCXColors.Visible = False
50090  cmbTIFFColors.Visible = False
50100  tbstrPDFOptions.Visible = False
50110  fraProgFont.Visible = False
50120  fraProgGeneral.Visible = False
50130  fraProgAutosave.Visible = False
50140  fraProgSave.Visible = False
50150  fraProgDirectories.Visible = False
50160  fraProgDocument.Visible = False
50170  fraBitmapGeneral.Visible = False
50180  fraPDFGeneral.Visible = False
50190  fraPDFCompress.Visible = False
50200  fraPDFFonts.Visible = False
50210  fraPDFColors.Visible = False
50220  fraPDFSecurity.Visible = False
50230  txtJPEGQuality.Visible = False
50240  lblJPEQQualityProzent.Visible = False
50250  fraPSGeneral.Visible = False
50260  cmbPSLanguageLevel.Visible = False
50270  cmbEPSLanguageLevel.Visible = False
50280  fraProgGhostscript.Visible = False
50290
50301  Select Case trv.SelectedItem.Key
              Case "Program"
50320    Set picOptions = LoadResPicture(2101, vbResIcon)
50330    lblOptions = LanguageStrings.OptionsProgramGeneralDescription
50340    fraProgGeneral.Visible = True
50350   Case "ProgramGeneral"
50360    Set picOptions = LoadResPicture(2101, vbResIcon)
50370    lblOptions = LanguageStrings.OptionsProgramGeneralDescription
50380    fraProgGeneral.Visible = True
50390   Case "ProgramGhostscript"
50400    Set picOptions = LoadResPicture(2119, vbResIcon)
50410    lblOptions = LanguageStrings.OptionsProgramGhostscriptDescription
50420    fraProgGhostscript.Visible = True
50430   Case "ProgramSave"
50440    Set picOptions = LoadResPicture(2106, vbResIcon)
50450    lblOptions = LanguageStrings.OptionsProgramSaveDescription
50460    fraProgSave.Visible = True
50470   Case "ProgramAutosave"
50480    Set picOptions = LoadResPicture(2103, vbResIcon)
50490    lblOptions = LanguageStrings.OptionsProgramAutosaveDescription
50500    fraProgAutosave.Visible = True
50510   Case "ProgramFonts"
50520    Set picOptions = LoadResPicture(2102, vbResIcon)
50530    lblOptions = LanguageStrings.OptionsProgramFontDescription
50540    fraProgFont.Visible = True
50550   Case "ProgramDirectories"
50560    Set picOptions = LoadResPicture(2104, vbResIcon)
50570    lblOptions = LanguageStrings.OptionsProgramDirectoriesDescription
50580    fraProgDirectories.Visible = True
50590   Case "ProgramDocument"
50600    Set picOptions = LoadResPicture(2105, vbResIcon)
50610    lblOptions = LanguageStrings.OptionsProgramDocumentDescription
50620    fraProgDocument.Visible = True
50630   Case "Formats"
50640    Set picOptions = LoadResPicture(2111, vbResIcon)
50650    lblOptions = LanguageStrings.OptionsPDFDescription
50660    tbstrPDFOptions.Visible = True
50670    fraPDFGeneral.Visible = True
50680   Case "FormatsPDF"
50690    Set picOptions = LoadResPicture(2111, vbResIcon)
50700    lblOptions = LanguageStrings.OptionsPDFDescription
50710    tbstrPDFOptions.Visible = True
50720    tbstrPDFOptions.Tabs(1).Selected = True
50730    fraPDFGeneral.Visible = True
50740   Case "FormatsPNG"
50750    Set picOptions = LoadResPicture(2112, vbResIcon)
50760    lblOptions = LanguageStrings.OptionsPNGDescription
50770    fraBitmapGeneral.Visible = True
50780    cmbPNGColors.Visible = True
50790   Case "FormatsJPEG"
50800    Set picOptions = LoadResPicture(2113, vbResIcon)
50810    lblOptions = LanguageStrings.OptionsJPEGDescription
50820    fraBitmapGeneral.Visible = True
50830    lblJPEGQuality.Caption = LanguageStrings.OptionsJPEGQuality
50840    lblJPEGQuality.Visible = True
50850    txtJPEGQuality.Visible = True
50860    lblJPEQQualityProzent.Visible = True
50870    lblJPEQQualityProzent.Left = txtJPEGQuality.Left + txtJPEGQuality.Width + 100
50880    cmbJPEGColors.Visible = True
50890   Case "FormatsBMP"
50900    Set picOptions = LoadResPicture(2114, vbResIcon)
50910    lblOptions = LanguageStrings.OptionsBMPDescription
50920    fraBitmapGeneral.Visible = True
50930    cmbBMPColors.Visible = True
50940   Case "FormatsPCX"
50950    Set picOptions = LoadResPicture(2115, vbResIcon)
50960    lblOptions = LanguageStrings.OptionsPCXDescription
50970    fraBitmapGeneral.Visible = True
50980    cmbPCXColors.Visible = True
50990   Case "FormatsTIFF"
51000    Set picOptions = LoadResPicture(2116, vbResIcon)
51010    lblOptions = LanguageStrings.OptionsTIFFDescription
51020    fraBitmapGeneral.Visible = True
51030    cmbTIFFColors.Visible = True
51040   Case "FormatsPS"
51050    Set picOptions = LoadResPicture(2117, vbResIcon)
51060    lblOptions.Caption = LanguageStrings.OptionsPSDescription
51070    fraPSGeneral.Visible = True
51080    cmbPSLanguageLevel.Visible = True
51090    fraPSGeneral.Caption = LanguageStrings.OptionsPSDescription
51100   Case "FormatsEPS"
51110    Set picOptions = LoadResPicture(2118, vbResIcon)
51120    lblOptions.Caption = LanguageStrings.OptionsEPSDescription
51130    fraPSGeneral.Visible = True
51140    cmbEPSLanguageLevel.Visible = True
51150    fraPSGeneral.Caption = LanguageStrings.OptionsEPSDescription
51160  End Select
51170 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
51180 Exit Sub
ErrPtnr_OnError:
51201 Select Case ErrPtnr.OnError("frmOptions", "ShowProgOptions")
      Case 0: Resume
51220 Case 1: Resume Next
51230 Case 2: Exit Sub
51240 Case 3: End
51250 End Select
51260 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub lsvFilenameSubst_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Set_txtFilenameSubst
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("frmOptions", "lsvFilenameSubst_Click")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub optEncHigh_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  UpdateSecurityFields
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("frmOptions", "optEncHigh_Click")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub optEncLow_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  UpdateSecurityFields
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("frmOptions", "optEncLow_Click")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub sldProcessPriority_Change()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & sldProcessPriority.Text
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("frmOptions", "sldProcessPriority_Change")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub sldProcessPriority_Scroll()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  With sldProcessPriority
50051   Select Case .Value
               Case 0: 'Idle
50070     .Text = LanguageStrings.OptionsProcesspriorityIdle
50080    Case 1: 'Normal
50090     .Text = LanguageStrings.OptionsProcesspriorityNormal
50100    Case 2: 'High
50110     .Text = LanguageStrings.OptionsProcesspriorityHigh
50120    Case 3: 'Realtime
50130     .Text = LanguageStrings.OptionsProcesspriorityRealtime
50140   End Select
50150  End With
50160 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50170 Exit Sub
ErrPtnr_OnError:
50191 Select Case ErrPtnr.OnError("frmOptions", "sldProcessPriority_Scroll")
      Case 0: Resume
50210 Case 1: Resume Next
50220 Case 2: Exit Sub
50230 Case 3: End
50240 End Select
50250 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub tbstrPDFOptions_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  fraPDFGeneral.Visible = False
50050  fraPDFCompress.Visible = False
50060  fraPDFFonts.Visible = False
50070  fraPDFColors.Visible = False
50080  fraPDFSecurity.Visible = False
50091  Select Case tbstrPDFOptions.SelectedItem.Index
              Case 1:
50110    fraPDFGeneral.Visible = True
50120   Case 2:
50130    fraPDFCompress.Visible = True
50140   Case 3:
50150    fraPDFFonts.Visible = True
50160   Case 4:
50170    fraPDFColors.Visible = True
50180   Case 5:
50190    fraPDFSecurity.Visible = True
50200    If SecurityIsPossible = False Then
50210     MsgBox LanguageStrings.MessagesMsg19
50220    End If
50230  End Select
50240 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50250 Exit Sub
ErrPtnr_OnError:
50271 Select Case ErrPtnr.OnError("frmOptions", "tbstrPDFOptions_Click")
      Case 0: Resume
50290 Case 1: Resume Next
50300 Case 2: Exit Sub
50310 Case 3: End
50320 End Select
50330 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub trv_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  ShowProgOptions
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("frmOptions", "trv_Click")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub trv_NodeClick(ByVal Node As MSComctlLib.Node)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  ShowProgOptions
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("frmOptions", "trv_NodeClick")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtAutosaveDirectory_Change()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  txtAutosaveDirectory.ToolTipText = txtAutosaveDirectory.Text
50050  txtAutoSaveDirectoryPreview.Text = GetSubstFilename("B:\dummy.dum", txtAutosaveDirectory.Text, , True)
50060  If IsValidPath(txtAutoSaveDirectoryPreview.Text) = False Then
50070    txtAutoSaveDirectoryPreview.ForeColor = vbRed
50080   Else
50090    txtAutoSaveDirectoryPreview.ForeColor = &H80000008
50100  End If
50110 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50120 Exit Sub
ErrPtnr_OnError:
50141 Select Case ErrPtnr.OnError("frmOptions", "txtAutosaveDirectory_Change")
      Case 0: Resume
50160 Case 1: Resume Next
50170 Case 2: Exit Sub
50180 Case 3: End
50190 End Select
50200 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtAutosaveFilename_Change()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim Ext As String
50050  txtAutosaveFilename.ToolTipText = txtAutosaveFilename.Text
50060  txtAutoSaveFilenamePreview.Text = GetSubstFilename("B:\dummy.dum", txtAutosaveFilename.Text, , True) & _
  GetAutosaveFormatExtension
50080  If IsValidPath("C:\" & txtAutoSaveFilenamePreview.Text) = False Then
50090    txtAutoSaveFilenamePreview.ForeColor = vbRed
50100   Else
50110    txtAutoSaveFilenamePreview.ForeColor = &H80000008
50120  End If
50130 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50140 Exit Sub
ErrPtnr_OnError:
50161 Select Case ErrPtnr.OnError("frmOptions", "txtAutosaveFilename_Change")
      Case 0: Resume
50180 Case 1: Resume Next
50190 Case 2: Exit Sub
50200 Case 3: End
50210 End Select
50220 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbProgramFontSize_Change()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim tL As Long
50050 If Trim$(cmbProgramFontsize.Text) = vbNullString Then
50060   cmbProgramFontsize.Text = 8
50070  End If
50080  tL = CLng(cmbProgramFontsize.Text)
50090  If tL <= 0 Then
50100   tL = 1
50110  End If
50120  If tL > 72 Then
50130   tL = 72
50140  End If
50150  cmbProgramFontsize.Text = tL
50160  txtTest.Font.Size = tL
50170 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50180 Exit Sub
ErrPtnr_OnError:
50201 Select Case ErrPtnr.OnError("frmOptions", "cmbProgramFontSize_Change")
      Case 0: Resume
50220 Case 1: Resume Next
50230 Case 2: Exit Sub
50240 Case 3: End
50250 End Select
50260 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbProgramFontSize_KeyPress(KeyAscii As Integer)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim allow As String, tStr As String
50050
50060  allow = "0123456789" & Chr$(8) & Chr$(13)
50070
50080  tStr = Chr$(KeyAscii)
50090
50100  If InStr(1, allow, tStr) = 0 Then
50110    KeyAscii = 0
50120  End If
50130 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50140 Exit Sub
ErrPtnr_OnError:
50161 Select Case ErrPtnr.OnError("frmOptions", "cmbProgramFontSize_KeyPress")
      Case 0: Resume
50180 Case 1: Resume Next
50190 Case 2: Exit Sub
50200 Case 3: End
50210 End Select
50220 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbProgramFontsize_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim tL As Long
50050 If Trim$(cmbProgramFontsize.Text) = vbNullString Then
50060   cmbProgramFontsize.Text = 8
50070  End If
50080  tL = CLng(cmbProgramFontsize.Text)
50090  If tL <= 0 Then
50100   tL = 1
50110  End If
50120  If tL > 72 Then
50130   tL = 72
50140  End If
50150  cmbProgramFontsize.Text = tL
50160  txtTest.Font.Size = tL
50170 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50180 Exit Sub
ErrPtnr_OnError:
50201 Select Case ErrPtnr.OnError("frmOptions", "cmbProgramFontsize_Click")
      Case 0: Resume
50220 Case 1: Resume Next
50230 Case 2: Exit Sub
50240 Case 3: End
50250 End Select
50260 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ViewAutosave(ViewIt As Boolean)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  lblAutosaveformat.Enabled = ViewIt
50050  cmbAutosaveFormat.Enabled = ViewIt
50060  lblAutosaveFilename.Enabled = ViewIt
50070  txtAutosaveFilename.Enabled = ViewIt
50080  txtAutoSaveFilenamePreview.Enabled = ViewIt
50090  lblAutosaveFilenameTokens.Enabled = ViewIt
50100  cmbAutoSaveFilenameTokens.Enabled = ViewIt
50110  chkUseAutosaveDirectory.Enabled = ViewIt
50120  txtAutoSaveDirectoryPreview.Enabled = ViewIt
50130  If ViewIt = True Then
50140    cmbAutosaveFormat.BackColor = &H80000005
50150    cmbAutoSaveFilenameTokens.BackColor = &H80000005
50160    txtAutosaveFilename.BackColor = &H80000005
50170    txtAutosaveDirectory.BackColor = &H80000005
50180   Else
50190    cmbAutosaveFormat.BackColor = &H8000000F
50200    cmbAutoSaveFilenameTokens.BackColor = &H8000000F
50210    txtAutosaveFilename.BackColor = &H8000000F
50220    txtAutosaveDirectory.BackColor = &H8000000F
50230  End If
50240  If chkUseAutosaveDirectory.Value = 1 And ViewIt = True Then
50250    ViewAutosaveDirectory True
50260   Else
50270    ViewAutosaveDirectory False
50280  End If
50290 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50300 Exit Sub
ErrPtnr_OnError:
50321 Select Case ErrPtnr.OnError("frmOptions", "ViewAutosave")
      Case 0: Resume
50340 Case 1: Resume Next
50350 Case 2: Exit Sub
50360 Case 3: End
50370 End Select
50380 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ViewAutosaveDirectory(ViewIt As Boolean)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If ViewIt = True Then
50050    txtAutosaveDirectory.Enabled = True
50060    txtAutosaveDirectory.BackColor = &H80000005
50070    cmdGetAutosaveDirectory.Enabled = True
50080   Else
50090    txtAutosaveDirectory.Enabled = False
50100    txtAutosaveDirectory.BackColor = &H8000000F
50110    cmdGetAutosaveDirectory.Enabled = False
50120  End If
50130 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50140 Exit Sub
ErrPtnr_OnError:
50161 Select Case ErrPtnr.OnError("frmOptions", "ViewAutosaveDirectory")
      Case 0: Resume
50180 Case 1: Resume Next
50190 Case 2: Exit Sub
50200 Case 3: End
50210 End Select
50220 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub UpdateSecurityFields()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If chkUseSecurity.Value = False Then
50050    fraPDFEncryptor.Enabled = False
50060    cmbPDFEncryptor.Enabled = False
50070
50080    fraPDFEncLevel.Enabled = False
50090    optEncHigh.Enabled = False
50100    optEncLow.Enabled = False
50110
50120    fraSecurityPass.Enabled = False
50130    chkUserPass.Enabled = False
50140    chkOwnerPass.Enabled = False
50150
50160    fraPDFPermissions.Enabled = False
50170    chkAllowPrinting.Enabled = False
50180    chkAllowCopy.Enabled = False
50190    chkAllowModifyAnnotations.Enabled = False
50200    chkAllowModifyContents.Enabled = False
50210
50220    fraPDFHighPermissions.Enabled = False
50230    chkAllowDegradedPrinting.Enabled = False
50240    chkAllowFillIn.Enabled = False
50250    chkAllowScreenReaders.Enabled = False
50260    chkAllowAssembly.Enabled = False
50270   Else
50280    fraPDFEncryptor.Enabled = True
50290    cmbPDFEncryptor.Enabled = True
50300
50310    fraPDFEncLevel.Enabled = True
50320    If cmbPDFCompat.ListIndex >= 2 Then
50330      optEncHigh.Enabled = True
50340     Else
50350      optEncHigh.Enabled = False
50360    End If
50370    optEncLow.Enabled = True
50380
50390    fraSecurityPass.Enabled = True
50400    chkUserPass.Enabled = True
50410    chkOwnerPass.Enabled = True
50420
50430    fraPDFPermissions.Enabled = True
50440    chkAllowPrinting.Enabled = True
50450    chkAllowCopy.Enabled = True
50460    chkAllowModifyAnnotations.Enabled = True
50470    chkAllowModifyContents.Enabled = True
50480
50490    If optEncHigh.Value = True Then
50500      fraPDFHighPermissions.Enabled = True
50510      chkAllowDegradedPrinting.Enabled = True
50520      chkAllowFillIn.Enabled = True
50530      chkAllowScreenReaders.Enabled = True
50540      chkAllowAssembly.Enabled = True
50550     Else
50560      fraPDFHighPermissions.Enabled = False
50570      chkAllowDegradedPrinting.Enabled = False
50580      chkAllowFillIn.Enabled = False
50590      chkAllowScreenReaders.Enabled = False
50600      chkAllowAssembly.Enabled = False
50610    End If
50620  End If
50630  If chkOwnerPass.Value = 0 And chkUserPass.Value = 0 Then
50640   chkOwnerPass.Value = 1: Options.PDFOwnerPass = 1
50650  End If
50660 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50670 Exit Sub
ErrPtnr_OnError:
50691 Select Case ErrPtnr.OnError("frmOptions", "UpdateSecurityFields")
      Case 0: Resume
50710 Case 1: Resume Next
50720 Case 2: Exit Sub
50730 Case 3: End
50740 End Select
50750 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdAsso_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  PsAssociate
50050  SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
50060  cmdAsso.Enabled = False
50070 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50080 Exit Sub
ErrPtnr_OnError:
50101 Select Case ErrPtnr.OnError("frmOptions", "cmdAsso_Click")
      Case 0: Resume
50120 Case 1: Resume Next
50130 Case 2: Exit Sub
50140 Case 3: End
50150 End Select
50160 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub AddFilenameSubstitutions()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim i As Long, res As Long
50050  res = CheckFilenameSubstitutions(0)
50061  Select Case res
              Case 0, 2:
50080    lsvFilenameSubst.ListItems.Add , , txtFilenameSubst(0).Text
50090    lsvFilenameSubst.ListItems(lsvFilenameSubst.ListItems.Count).SubItems(1) = txtFilenameSubst(1).Text
50100    lsvFilenameSubst.ListItems(lsvFilenameSubst.ListItems.Count).Selected = True
50110    lsvFilenameSubst.ListItems(lsvFilenameSubst.ListItems.Count).EnsureVisible
50120    Set_txtFilenameSubst
50130 '  Case 2:
50140 '   MsgBox LanguageStrings.MessagesMsg12 & _
    vbCrLf & vbTab & "\ / : * ? < > | """
50160   Case 3:
50170    MsgBox LanguageStrings.MessagesMsg11
50180  End Select
50190 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50200 Exit Sub
ErrPtnr_OnError:
50221 Select Case ErrPtnr.OnError("frmOptions", "AddFilenameSubstitutions")
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
50061  Select Case res
              Case 0, 2:
50080    lsvFilenameSubst.SelectedItem.Text = txtFilenameSubst(0).Text
50090    lsvFilenameSubst.SelectedItem.SubItems(1) = txtFilenameSubst(1).Text
50100 '  Case 2:
50110 '   MsgBox LanguageStrings.MessagesMsg12 & _
    vbCrLf & vbTab & "\ / : * ? < > | """
50130   Case 3:
50140    MsgBox LanguageStrings.MessagesMsg11
50150  End Select
50160 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50170 Exit Sub
ErrPtnr_OnError:
50191 Select Case ErrPtnr.OnError("frmOptions", "ChangeFilenameSubstitutions")
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
50221 Select Case ErrPtnr.OnError("frmOptions", "DeleteFilenameSubstitutions")
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
50191 Select Case ErrPtnr.OnError("frmOptions", "MoveUpFilenameSubstitutions")
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
50191 Select Case ErrPtnr.OnError("frmOptions", "MoveDownFilenameSubstitutions")
      Case 0: Resume
50210 Case 1: Resume Next
50220 Case 2: Exit Sub
50230 Case 3: End
50240 End Select
50250 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function CheckFilenameSubstitutions(Index As Long) As Long
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim i As Long
50050  CheckFilenameSubstitutions = 0
50060  If Len(txtFilenameSubst(0).Text) = 0 Then
50070   CheckFilenameSubstitutions = 1
50080   Exit Function
50090  End If
50100  If IsForbiddenChars(txtFilenameSubst(0).Text) = True Then
50110   txtFilenameSubst(0).SetFocus
50120   CheckFilenameSubstitutions = 2
50130   Exit Function
50140  End If
50150  If IsForbiddenChars(txtFilenameSubst(1).Text) = True Then
50160   txtFilenameSubst(1).SetFocus
50170   CheckFilenameSubstitutions = 2
50180   Exit Function
50190  End If
50200  If Index = 0 Then
50210    For i = 1 To lsvFilenameSubst.ListItems.Count
50220     If UCase$(txtFilenameSubst(0).Text) = UCase$(lsvFilenameSubst.ListItems(i).Text) Then
50230      CheckFilenameSubstitutions = 3
50240      Exit Function
50250     End If
50260    Next i
50270   Else
50280    For i = 1 To lsvFilenameSubst.ListItems.Count
50290     If UCase$(txtFilenameSubst(0).Text) = UCase$(lsvFilenameSubst.ListItems(i).Text) And _
     Index <> lsvFilenameSubst.SelectedItem.Index Then
50310      CheckFilenameSubstitutions = 3
50320      Exit Function
50330     End If
50340    Next i
50350  End If
50360 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50370 Exit Function
ErrPtnr_OnError:
50391 Select Case ErrPtnr.OnError("frmOptions", "CheckFilenameSubstitutions")
      Case 0: Resume
50410 Case 1: Resume Next
50420 Case 2: Exit Function
50430 Case 3: End
50440 End Select
50450 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub CheckCmdFilenameSubst()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If lsvFilenameSubst.ListItems.Count > 0 Then
50050    cmdFilenameSubst(1).Enabled = True
50060    cmdFilenameSubst(2).Enabled = True
50070   Else
50080    cmdFilenameSubst(1).Enabled = False
50090    cmdFilenameSubst(2).Enabled = False
50100  End If
50110  If lsvFilenameSubst.ListItems.Count > 1 Then
50120    cmdFilenameSubst(3).Enabled = True
50130    cmdFilenameSubst(4).Enabled = True
50140   Else
50150    cmdFilenameSubst(3).Enabled = False
50160    cmdFilenameSubst(4).Enabled = False
50170  End If
50180  If lsvFilenameSubst.ListItems.Count > 0 Then
50190   If lsvFilenameSubst.SelectedItem.Index = 1 Then
50200    cmdFilenameSubst(3).Enabled = False
50210   End If
50220   If lsvFilenameSubst.SelectedItem.Index = lsvFilenameSubst.ListItems.Count Then
50230    cmdFilenameSubst(4).Enabled = False
50240   End If
50250  End If
50260 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50270 Exit Sub
ErrPtnr_OnError:
50291 Select Case ErrPtnr.OnError("frmOptions", "CheckCmdFilenameSubst")
      Case 0: Resume
50310 Case 1: Resume Next
50320 Case 2: Exit Sub
50330 Case 3: End
50340 End Select
50350 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Set_txtFilenameSubst()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  CheckCmdFilenameSubst
50050  If lsvFilenameSubst.ListItems.Count > 0 Then
50060   txtFilenameSubst(0).Text = lsvFilenameSubst.SelectedItem.Text
50070   txtFilenameSubst(0).ToolTipText = txtFilenameSubst(0).Text
50080   txtFilenameSubst(1).Text = lsvFilenameSubst.SelectedItem.SubItems(1)
50090   txtFilenameSubst(1).ToolTipText = txtFilenameSubst(1).Text
50100  End If
50110 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50120 Exit Sub
ErrPtnr_OnError:
50141 Select Case ErrPtnr.OnError("frmOptions", "Set_txtFilenameSubst")
      Case 0: Resume
50160 Case 1: Resume Next
50170 Case 2: Exit Sub
50180 Case 3: End
50190 End Select
50200 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub


Private Sub txtSaveFilename_Change()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  txtSaveFilename.ToolTipText = txtSaveFilename.Text
50050  txtSavePreview.Text = GetSubstFilename("C:\test.pdf", txtSaveFilename.Text, , True) & ".pdf"
50060 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50070 Exit Sub
ErrPtnr_OnError:
50091 Select Case ErrPtnr.OnError("frmOptions", "txtSaveFilename_Change")
      Case 0: Resume
50110 Case 1: Resume Next
50120 Case 2: Exit Sub
50130 Case 3: End
50140 End Select
50150 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function GetAutosaveFormatExtension() As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50041  Select Case cmbAutosaveFormat.ListIndex
        Case -1, 0
50060    GetAutosaveFormatExtension = ".pdf"
50070   Case 1
50080    GetAutosaveFormatExtension = ".png"
50090   Case 2
50100    GetAutosaveFormatExtension = ".jpg"
50110   Case 3
50120    GetAutosaveFormatExtension = ".bmp"
50130   Case 4
50140    GetAutosaveFormatExtension = ".pcx"
50150   Case 5
50160    GetAutosaveFormatExtension = ".tif"
50170   Case 6
50180    GetAutosaveFormatExtension = ".ps"
50190   Case 7
50200    GetAutosaveFormatExtension = ".eps"
50210  End Select
50220 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50230 Exit Function
ErrPtnr_OnError:
50251 Select Case ErrPtnr.OnError("frmOptions", "GetAutosaveFormatExtension")
      Case 0: Resume
50270 Case 1: Resume Next
50280 Case 2: Exit Function
50290 Case 3: End
50300 End Select
50310 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub CorrectCmbCharset()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim tStrf() As String
50050  If InStr(1, cmbCharset.Text, ",", vbTextCompare) > 0 Then
50060    tStrf = Split(cmbCharset.Text, ",")
50070    If Len(tStrf(0)) = 0 Then
50080      cmbCharset.Text = 0
50090     Else
50100      If IsNumeric(tStrf(0)) = False Then
50110        cmbCharset.Text = 0
50120       Else
50130        cmbCharset.Text = CLng(tStrf(0))
50140      End If
50150    End If
50160   Else
50170    If Len(cmbCharset.Text) = 0 Then
50180      cmbCharset.Text = 0
50190     Else
50200      If IsNumeric(cmbCharset.Text) = False Then
50210        cmbCharset.Text = 0
50220       Else
50230        cmbCharset.Text = CLng(cmbCharset.Text)
50240      End If
50250    End If
50260  End If
50270 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50280 Exit Sub
ErrPtnr_OnError:
50301 Select Case ErrPtnr.OnError("frmOptions", "CorrectCmbCharset")
      Case 0: Resume
50320 Case 1: Resume Next
50330 Case 2: Exit Sub
50340 Case 3: End
50350 End Select
50360 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
