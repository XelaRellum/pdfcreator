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
   StartUpPosition =   3  'Windows-Standard
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
         Picture         =   "frmOptions.frx":058A
         Style           =   1  'Grafisch
         TabIndex        =   163
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
   Begin VB.Frame fraProgSave 
      Caption         =   "Save"
      Height          =   4335
      Left            =   2880
      TabIndex        =   107
      Top             =   1560
      Visible         =   0   'False
      Width           =   6255
      Begin VB.Frame fraFilenameSubstitutions 
         Caption         =   "Filename substitutions"
         Height          =   2415
         Left            =   120
         TabIndex        =   109
         Top             =   1800
         Width           =   6015
         Begin VB.CommandButton cmdFilenameSubst 
            Caption         =   "Delete"
            Height          =   375
            Index           =   2
            Left            =   4440
            TabIndex        =   118
            Top             =   1560
            Width           =   1455
         End
         Begin VB.CommandButton cmdFilenameSubst 
            Caption         =   "Change"
            Height          =   375
            Index           =   1
            Left            =   4440
            TabIndex        =   117
            Top             =   1080
            Width           =   1455
         End
         Begin VB.CommandButton cmdFilenameSubst 
            Caption         =   "Add"
            Height          =   375
            Index           =   0
            Left            =   4440
            TabIndex        =   116
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtFilenameSubst 
            Height          =   285
            Index           =   0
            Left            =   600
            TabIndex        =   115
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmdFilenameSubst 
            Enabled         =   0   'False
            Height          =   420
            Index           =   3
            Left            =   120
            Picture         =   "frmOptions.frx":0914
            Style           =   1  'Grafisch
            TabIndex        =   114
            Top             =   795
            Width           =   375
         End
         Begin VB.CommandButton cmdFilenameSubst 
            Enabled         =   0   'False
            Height          =   420
            Index           =   4
            Left            =   120
            Picture         =   "frmOptions.frx":0C9E
            Style           =   1  'Grafisch
            TabIndex        =   113
            Top             =   1320
            Width           =   375
         End
         Begin VB.CheckBox chkFilenameSubst 
            Caption         =   "Substitutions only in <Title>"
            Height          =   255
            Left            =   240
            TabIndex        =   112
            Top             =   2040
            Value           =   1  'Aktiviert
            Width           =   3255
         End
         Begin VB.TextBox txtFilenameSubst 
            Height          =   285
            Index           =   1
            Left            =   2640
            TabIndex        =   111
            Top             =   240
            Width           =   1695
         End
         Begin MSComctlLib.ListView lsvFilenameSubst 
            Height          =   1335
            Left            =   600
            TabIndex        =   110
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
            TabIndex        =   119
            Top             =   240
            Width           =   135
         End
      End
      Begin VB.TextBox txtSavePreview 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   162
         Top             =   840
         Width           =   6015
      End
      Begin VB.ComboBox cmbSaveFilenameTokens 
         Height          =   315
         ItemData        =   "frmOptions.frx":1028
         Left            =   3720
         List            =   "frmOptions.frx":102A
         Style           =   2  'Dropdown-Liste
         TabIndex        =   121
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtSaveFilename 
         Height          =   285
         Left            =   120
         TabIndex        =   120
         Text            =   "<Title>"
         Top             =   480
         Width           =   3495
      End
      Begin VB.CheckBox chkSpaces 
         Caption         =   "Remove leading and trailing spaces"
         Height          =   255
         Left            =   120
         TabIndex        =   108
         Top             =   1320
         Value           =   1  'Aktiviert
         Width           =   6015
      End
      Begin VB.Label lblSaveFilenameTokens 
         AutoSize        =   -1  'True
         Caption         =   "Add a Filename-Token"
         Height          =   195
         Left            =   3720
         TabIndex        =   123
         Top             =   240
         Width           =   1605
      End
      Begin VB.Label lblSaveFilename 
         AutoSize        =   -1  'True
         Caption         =   "Filename"
         Height          =   195
         Left            =   120
         TabIndex        =   122
         Top             =   240
         Width           =   630
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
      Begin VB.ComboBox cmbFonts 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown-Liste
         TabIndex        =   39
         Top             =   600
         Width           =   2775
      End
      Begin VB.ComboBox cmbCharset 
         Height          =   315
         Left            =   3000
         TabIndex        =   38
         Text            =   "cmbCharset"
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtTest 
         Height          =   2655
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   37
         Top             =   1320
         Width           =   6015
      End
      Begin VB.TextBox txtProgramFontsize 
         Alignment       =   1  'Rechts
         Height          =   315
         Left            =   5400
         TabIndex        =   36
         Text            =   "8"
         Top             =   600
         Width           =   615
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
         TabIndex        =   43
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblProgcharset 
         Caption         =   "Charset"
         Height          =   255
         Left            =   3000
         TabIndex        =   42
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblTesttext 
         Caption         =   "Here you can test the font."
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   1080
         Width           =   5295
      End
      Begin VB.Label lblSize 
         Caption         =   "Size"
         Height          =   255
         Left            =   5400
         TabIndex        =   40
         Top             =   360
         Width           =   615
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
         ItemData        =   "frmOptions.frx":102C
         Left            =   3720
         List            =   "frmOptions.frx":102E
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
   Begin VB.Frame fraBitmapGeneral 
      Caption         =   "Bitmap"
      Height          =   1935
      Left            =   2880
      TabIndex        =   71
      Top             =   1080
      Width           =   6255
      Begin VB.TextBox txtBitmapResolution 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   1920
         TabIndex        =   78
         Text            =   "72"
         Top             =   480
         Width           =   495
      End
      Begin VB.ComboBox cmbPNGColors 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown-Liste
         TabIndex        =   77
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox txtJPEGQuality 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   1920
         TabIndex        =   76
         Text            =   "75"
         Top             =   1440
         Width           =   495
      End
      Begin VB.ComboBox cmbJPEGColors 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown-Liste
         TabIndex        =   75
         Top             =   1080
         Width           =   2175
      End
      Begin VB.ComboBox cmbBMPColors 
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown-Liste
         TabIndex        =   74
         Top             =   1200
         Width           =   2175
      End
      Begin VB.ComboBox cmbPCXColors 
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown-Liste
         TabIndex        =   73
         Top             =   1320
         Width           =   2175
      End
      Begin VB.ComboBox cmbTIFFColors 
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown-Liste
         TabIndex        =   72
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblBitmapResolution 
         Alignment       =   1  'Rechts
         Caption         =   "Resolution:"
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblBitmapDPI 
         Caption         =   "dpi"
         Height          =   255
         Left            =   2520
         TabIndex        =   82
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblBitmapColors 
         Alignment       =   1  'Rechts
         Caption         =   "Colors:"
         Height          =   255
         Left            =   120
         TabIndex        =   81
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblJPEGQuality 
         Alignment       =   1  'Rechts
         Caption         =   "Quality:"
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblJPEQQualityProzent 
         Caption         =   "%"
         Height          =   255
         Left            =   2520
         TabIndex        =   79
         Top             =   1440
         Width           =   255
      End
   End
   Begin VB.Frame fraProgAutosave 
      Caption         =   "Autosave"
      Height          =   3495
      Left            =   2760
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   6255
      Begin VB.TextBox txtAutoSavePreview 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   161
         Top             =   2040
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
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
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
         ItemData        =   "frmOptions.frx":1030
         Left            =   3720
         List            =   "frmOptions.frx":1032
         Style           =   2  'Dropdown-Liste
         TabIndex        =   9
         Top             =   1680
         Width           =   2415
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
   Begin VB.Frame fraProgGhostscript 
      Caption         =   "Ghostscript"
      Height          =   975
      Left            =   2760
      TabIndex        =   146
      Top             =   1080
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton cmdGetgsfontsDirectory 
         Caption         =   "..."
         Height          =   255
         Left            =   5760
         TabIndex        =   155
         Top             =   2400
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdGetgslibDirectory 
         Caption         =   "..."
         Height          =   255
         Left            =   5760
         TabIndex        =   154
         Top             =   1800
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtGSfonts 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   153
         Top             =   2400
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.TextBox txtGSlib 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   152
         Top             =   1800
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.TextBox txtGSbin 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   150
         Top             =   1200
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.CommandButton cmdGetgsbinDirectory 
         Caption         =   "..."
         Height          =   255
         Left            =   5760
         TabIndex        =   149
         Top             =   1200
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox cmbGhostscript 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown-Liste
         TabIndex        =   148
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label lblGSfonts 
         Caption         =   "Ghostscript Fonts"
         Height          =   255
         Left            =   240
         TabIndex        =   157
         Top             =   2160
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblGSlib 
         Caption         =   "Ghostscript Libraries"
         Height          =   255
         Left            =   240
         TabIndex        =   156
         Top             =   1560
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblGSbin 
         Caption         =   "Ghostscript Binaries"
         Height          =   255
         Left            =   240
         TabIndex        =   151
         Top             =   960
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblGhostscriptversion 
         Caption         =   "Ghostscriptversion"
         Height          =   255
         Left            =   240
         TabIndex        =   147
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.Frame fraPDFSecurity 
      Caption         =   "Security"
      Height          =   5415
      Left            =   2880
      TabIndex        =   124
      Top             =   1320
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Frame fraPDFEncryptor 
         Caption         =   "Encryptor"
         Height          =   855
         Left            =   120
         TabIndex        =   144
         Top             =   600
         Width           =   5535
         Begin VB.ComboBox cmbPDFEncryptor 
            Height          =   315
            ItemData        =   "frmOptions.frx":1034
            Left            =   240
            List            =   "frmOptions.frx":1036
            Style           =   2  'Dropdown-Liste
            TabIndex        =   145
            Top             =   360
            Width           =   5175
         End
      End
      Begin VB.CheckBox chkUseSecurity 
         Caption         =   "Use Security"
         Height          =   255
         Left            =   120
         TabIndex        =   141
         Top             =   240
         Width           =   5535
      End
      Begin VB.Frame fraPDFEncLevel 
         Caption         =   "Encryption Level"
         Height          =   855
         Left            =   120
         TabIndex        =   138
         Top             =   1560
         Width           =   5535
         Begin VB.OptionButton optEncLow 
            Caption         =   "Low (40 Bit - Adobe Acrobat 3.0 and above)"
            Height          =   255
            Left            =   240
            TabIndex        =   140
            Top             =   240
            Width           =   4215
         End
         Begin VB.OptionButton optEncHigh 
            Caption         =   "High (128 Bit - Adobe Acrobat 5.0 and above)"
            Height          =   255
            Left            =   240
            TabIndex        =   139
            Top             =   480
            Width           =   4215
         End
      End
      Begin VB.Frame fraSecurityPass 
         Caption         =   "Passwords"
         Height          =   855
         Left            =   120
         TabIndex        =   135
         Top             =   2520
         Width           =   5535
         Begin VB.CheckBox chkUserPass 
            Caption         =   "Password required to open document"
            Height          =   255
            Left            =   240
            TabIndex        =   137
            Top             =   240
            Width           =   5175
         End
         Begin VB.CheckBox chkOwnerPass 
            Caption         =   "Password required to change Permissions and Passwords"
            Height          =   255
            Left            =   240
            TabIndex        =   136
            Top             =   480
            Width           =   5175
         End
      End
      Begin VB.Frame fraPDFPermissions 
         Caption         =   "Disallow User to"
         Height          =   855
         Left            =   120
         TabIndex        =   130
         Top             =   3480
         Width           =   5535
         Begin VB.CheckBox chkAllowPrinting 
            Caption         =   "print the document"
            Height          =   255
            Left            =   240
            TabIndex        =   134
            Top             =   240
            Width           =   2535
         End
         Begin VB.CheckBox chkAllowCopy 
            Caption         =   "copy text and images"
            Height          =   255
            Left            =   240
            TabIndex        =   133
            Top             =   480
            Width           =   2535
         End
         Begin VB.CheckBox chkAllowModifyContents 
            Caption         =   "modify the document"
            Height          =   255
            Left            =   2760
            TabIndex        =   132
            Top             =   240
            Width           =   2535
         End
         Begin VB.CheckBox chkAllowModifyAnnotations 
            Caption         =   "modify comments"
            Height          =   255
            Left            =   2760
            TabIndex        =   131
            Top             =   480
            Width           =   2295
         End
      End
      Begin VB.Frame fraPDFHighPermissions 
         Caption         =   "Enhanced Permissions (128 Bit only)"
         Height          =   855
         Left            =   120
         TabIndex        =   125
         Top             =   4440
         Width           =   5535
         Begin VB.CheckBox chkAllowDegradedPrinting 
            Caption         =   "Allow printing in low resolution"
            Height          =   255
            Left            =   240
            TabIndex        =   129
            Top             =   240
            Width           =   2535
         End
         Begin VB.CheckBox chkAllowFillIn 
            Caption         =   "Allow filling in form fields"
            Height          =   255
            Left            =   2760
            TabIndex        =   128
            Top             =   240
            Width           =   2415
         End
         Begin VB.CheckBox chkAllowScreenReaders 
            Caption         =   "Allow Screen Readers"
            Height          =   255
            Left            =   240
            TabIndex        =   127
            Top             =   480
            Width           =   2535
         End
         Begin VB.CheckBox chkAllowAssembly 
            Caption         =   "Allow changes to the Assembly"
            Height          =   255
            Left            =   2760
            TabIndex        =   126
            Top             =   480
            Width           =   2535
         End
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
         TabIndex        =   158
         Top             =   2880
         Width           =   6015
         Begin VB.CommandButton cmdShellintegration 
            Caption         =   "Remove shell integration"
            Height          =   615
            Index           =   1
            Left            =   3240
            TabIndex        =   160
            Top             =   360
            Width           =   2655
         End
         Begin VB.CommandButton cmdShellintegration 
            Caption         =   "Integrate PDFCreator into shell"
            Height          =   615
            Index           =   0
            Left            =   120
            TabIndex        =   159
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.CheckBox chkNoConfirmMessageSwitchingDefaultprinter 
         Caption         =   "No confirm message switching PDFCreator temporarly as default printer."
         Height          =   495
         Left            =   120
         TabIndex        =   143
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
         Caption         =   "Processpriority: Normal"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   2535
      End
   End
   Begin VB.Frame fraPDFCompress 
      Caption         =   "Compression"
      Height          =   3855
      Left            =   2760
      TabIndex        =   44
      Top             =   1080
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Frame fraPDFColor 
         Caption         =   "Color Images"
         Height          =   975
         Left            =   120
         TabIndex        =   60
         Top             =   600
         Width           =   5535
         Begin VB.CheckBox chkPDFColorComp 
            Caption         =   "Compress"
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox cmbPDFColorComp 
            Height          =   315
            ItemData        =   "frmOptions.frx":1038
            Left            =   120
            List            =   "frmOptions.frx":103A
            Style           =   2  'Dropdown-Liste
            TabIndex        =   64
            Top             =   540
            Width           =   2055
         End
         Begin VB.CheckBox chkPDFColorResample 
            Caption         =   "Resample"
            Height          =   255
            Left            =   2280
            TabIndex        =   63
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox cmbPDFColorResample 
            Height          =   315
            ItemData        =   "frmOptions.frx":103C
            Left            =   2280
            List            =   "frmOptions.frx":103E
            Style           =   2  'Dropdown-Liste
            TabIndex        =   62
            Tag             =   "Bicubic|Subsample|Average"
            Top             =   540
            Width           =   2055
         End
         Begin VB.TextBox txtPDFColorRes 
            Height          =   285
            Left            =   4440
            TabIndex        =   61
            Top             =   540
            Width           =   735
         End
         Begin VB.Label lblPDFColorRes 
            Caption         =   "Resolution"
            Height          =   255
            Left            =   4440
            TabIndex        =   66
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.CheckBox chkPDFTextComp 
         Caption         =   "Compress Text Objects"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   240
         Width           =   4335
      End
      Begin VB.Frame fraPDFMono 
         Caption         =   "Monochrome Images"
         Height          =   975
         Left            =   120
         TabIndex        =   52
         Top             =   2760
         Width           =   5535
         Begin VB.TextBox txtPDFMonoRes 
            Height          =   285
            Left            =   4440
            TabIndex        =   57
            Top             =   540
            Width           =   735
         End
         Begin VB.ComboBox cmbPDFMonoResample 
            Height          =   315
            ItemData        =   "frmOptions.frx":1040
            Left            =   2280
            List            =   "frmOptions.frx":1042
            Style           =   2  'Dropdown-Liste
            TabIndex        =   56
            Tag             =   "Bicubic|Subsample|Average"
            Top             =   540
            Width           =   2055
         End
         Begin VB.CheckBox chkPDFMonoResample 
            Caption         =   "Resample"
            Height          =   255
            Left            =   2280
            TabIndex        =   55
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox cmbPDFMonoComp 
            Height          =   315
            ItemData        =   "frmOptions.frx":1044
            Left            =   120
            List            =   "frmOptions.frx":1046
            Style           =   2  'Dropdown-Liste
            TabIndex        =   54
            Top             =   540
            Width           =   2055
         End
         Begin VB.CheckBox chkPDFMonoComp 
            Caption         =   "Compress"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblPDFMonoRes 
            Caption         =   "Resolution"
            Height          =   255
            Left            =   4440
            TabIndex        =   58
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame fraPDFGrey 
         Caption         =   "Greyscale Images"
         Height          =   975
         Left            =   120
         TabIndex        =   45
         Top             =   1680
         Width           =   5535
         Begin VB.CheckBox chkPDFGreyComp 
            Caption         =   "Compress"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox cmbPDFGreyComp 
            Height          =   315
            ItemData        =   "frmOptions.frx":1048
            Left            =   120
            List            =   "frmOptions.frx":104A
            Style           =   2  'Dropdown-Liste
            TabIndex        =   49
            Top             =   540
            Width           =   2055
         End
         Begin VB.CheckBox chkPDFGreyResample 
            Caption         =   "Resample"
            Height          =   255
            Left            =   2280
            TabIndex        =   48
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox cmbPDFGreyResample 
            Height          =   315
            ItemData        =   "frmOptions.frx":104C
            Left            =   2280
            List            =   "frmOptions.frx":104E
            Style           =   2  'Dropdown-Liste
            TabIndex        =   47
            Tag             =   "Bicubic|Subsample|Average"
            Top             =   540
            Width           =   2055
         End
         Begin VB.TextBox txtPDFGreyRes 
            Height          =   285
            Left            =   4440
            TabIndex        =   46
            Top             =   540
            Width           =   735
         End
         Begin VB.Label lblPDFGreyRes 
            Caption         =   "Resolution"
            Height          =   255
            Left            =   4440
            TabIndex        =   51
            Top             =   240
            Width           =   975
         End
      End
   End
   Begin VB.Frame fraPDFGeneral 
      Caption         =   "General Options"
      Height          =   2895
      Left            =   3240
      TabIndex        =   96
      Top             =   1800
      Visible         =   0   'False
      Width           =   5775
      Begin VB.ComboBox cmbPDFRotate 
         Height          =   315
         ItemData        =   "frmOptions.frx":1050
         Left            =   2400
         List            =   "frmOptions.frx":1052
         Style           =   2  'Dropdown-Liste
         TabIndex        =   101
         Tag             =   "None|All|PageByPage"
         Top             =   840
         Width           =   2655
      End
      Begin VB.ComboBox cmbPDFCompat 
         Height          =   315
         ItemData        =   "frmOptions.frx":1054
         Left            =   2400
         List            =   "frmOptions.frx":1056
         Style           =   2  'Dropdown-Liste
         TabIndex        =   100
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox txtPDFRes 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   2400
         TabIndex        =   99
         Text            =   "600"
         Top             =   1320
         Width           =   615
      End
      Begin VB.ComboBox cmbPDFOverprint 
         Height          =   315
         ItemData        =   "frmOptions.frx":1058
         Left            =   2400
         List            =   "frmOptions.frx":105A
         Style           =   2  'Dropdown-Liste
         TabIndex        =   98
         Top             =   1860
         Width           =   2655
      End
      Begin VB.CheckBox chkPDFASCII85 
         Caption         =   "Convert binary data to ASCII85"
         Height          =   255
         Left            =   2400
         TabIndex        =   97
         Top             =   2400
         Width           =   3255
      End
      Begin VB.Label lblPDFAutoRotate 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Auto-Rotate Pages:"
         Height          =   255
         Left            =   120
         TabIndex        =   106
         Top             =   900
         Width           =   2175
      End
      Begin VB.Label lblPDFCompat 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Compatibility:"
         Height          =   255
         Left            =   120
         TabIndex        =   105
         Top             =   420
         Width           =   2175
      End
      Begin VB.Label lblPDFResolution 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Resolution:"
         Height          =   255
         Left            =   120
         TabIndex        =   104
         Top             =   1380
         Width           =   2175
      End
      Begin VB.Label lblPDFOverprint 
         Alignment       =   1  'Rechts
         Caption         =   "Overprint:"
         Height          =   375
         Left            =   120
         TabIndex        =   103
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label lblPDFDPI 
         BackStyle       =   0  'Transparent
         Caption         =   "dpi"
         Height          =   255
         Left            =   3120
         TabIndex        =   102
         Top             =   1320
         Width           =   735
      End
   End
   Begin VB.Frame fraPDFFonts 
      Caption         =   "Font Options"
      Height          =   2895
      Left            =   3720
      TabIndex        =   84
      Top             =   2760
      Visible         =   0   'False
      Width           =   5775
      Begin VB.TextBox txtPDFSubSetPerc 
         Height          =   285
         Left            =   360
         TabIndex        =   87
         Top             =   1320
         Width           =   495
      End
      Begin VB.CheckBox chkPDFSubSetFonts 
         Caption         =   "Subset Fonts, when percentage of used characters below:"
         Height          =   495
         Left            =   120
         TabIndex        =   86
         Top             =   780
         Width           =   5535
      End
      Begin VB.CheckBox chkPDFEmbedAll 
         Caption         =   "Embed all Fonts"
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   360
         Width           =   5175
      End
      Begin VB.Label lblPDFPerc 
         Caption         =   "%"
         Height          =   255
         Left            =   960
         TabIndex        =   88
         Top             =   1320
         Width           =   615
      End
   End
   Begin VB.Frame fraPSGeneral 
      Caption         =   "Postscript"
      Height          =   1095
      Left            =   2760
      TabIndex        =   67
      Top             =   1080
      Visible         =   0   'False
      Width           =   6255
      Begin VB.ComboBox cmbPSLanguageLevel 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown-Liste
         TabIndex        =   69
         Top             =   480
         Width           =   735
      End
      Begin VB.ComboBox cmbEPSLanguageLevel 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown-Liste
         TabIndex        =   68
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblLangLevel 
         Alignment       =   1  'Rechts
         Caption         =   "Language Level:"
         Height          =   255
         Left            =   240
         TabIndex        =   70
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame fraPDFColors 
      Caption         =   "Color Options"
      Height          =   3495
      Left            =   3480
      TabIndex        =   89
      Top             =   2280
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CheckBox chkPDFCMYKtoRGB 
         Caption         =   "Convert CMYK Images to RGB"
         Height          =   255
         Left            =   120
         TabIndex        =   95
         Top             =   840
         Width           =   3255
      End
      Begin VB.Frame fraPDFColorOptions 
         Caption         =   "Options"
         Height          =   1455
         Left            =   120
         TabIndex        =   91
         Top             =   1920
         Width           =   5535
         Begin VB.CheckBox chkPDFPreserveOverprint 
            Caption         =   "Preserve Overprint Settings"
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   360
            Width           =   5175
         End
         Begin VB.CheckBox chkPDFPreserveTransfer 
            Caption         =   "Preserve Transfer Functions"
            Height          =   255
            Left            =   120
            TabIndex        =   93
            Tag             =   "Remove|Preserve"
            Top             =   720
            Width           =   5175
         End
         Begin VB.CheckBox chkPDFPreserveHalftone 
            Caption         =   "Preserve Halftone Information"
            Height          =   255
            Left            =   120
            TabIndex        =   92
            Top             =   1080
            Width           =   5175
         End
      End
      Begin VB.ComboBox cmbPDFColorModel 
         Height          =   315
         ItemData        =   "frmOptions.frx":105C
         Left            =   120
         List            =   "frmOptions.frx":105E
         Style           =   2  'Dropdown-Liste
         TabIndex        =   90
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
      TabIndex        =   142
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
50110    Set reg = Nothing
50120    Exit Sub
50130   Else
50140    If InStr(UCase$(gsv), "AFPL") Then
50150     If InStr(gsv, " ") > 0 Then
50160      tsf = Split(gsv, " ")
50170      reg.KeyRoot = "SOFTWARE\AFPL Ghostscript\" & tsf(UBound(tsf))
50180      tStr = reg.GetRegistryValue("GS_DLL")
50190      SplitPath tStr, , Path
50200      txtGSbin.Text = CompletePath(Path)
50210      txtGSlib.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50220      txtGSfonts.Text = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50230     End If
50240    End If
50250    If InStr(UCase$(gsv), "GNU") Then
50260     If InStr(gsv, " ") > 0 Then
50270      tsf = Split(gsv, " ")
50280      reg.KeyRoot = "SOFTWARE\GNU Ghostscript\" & tsf(UBound(tsf))
50290      tStr = reg.GetRegistryValue("GS_DLL")
50300      SplitPath tStr, , Path
50310      txtGSbin.Text = CompletePath(Path)
50320      txtGSlib.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50330      txtGSfonts.Text = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50340     End If
50350    End If
50360    If InStr(UCase$(gsv), "GPL") Then
50370     If InStr(gsv, " ") > 0 Then
50380      tsf = Split(gsv, " ")
50390      reg.KeyRoot = "SOFTWARE\GPL Ghostscript\" & tsf(UBound(tsf))
50400      tStr = reg.GetRegistryValue("GS_DLL")
50410      SplitPath tStr, , Path
50420      txtGSbin.Text = CompletePath(Path)
50430      txtGSlib.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50440      txtGSfonts.Text = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50450     End If
50460    End If
50470  End If
50480
50490  Set reg = Nothing
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
50011  Select Case Index
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
50030 strFolder = BrowseForFolder(Me.hwnd, LanguageStrings.OptionsAutosaveDirectoryPrompt)
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
50010  Dim strFolder As String, aw As Long
50020
50030  strFolder = BrowseForFolder(Me.hwnd, LanguageStrings.OptionsGhostscriptBinariesDirectoryPrompt)
50040  If Len(strFolder) = 0 Then Exit Sub
50050  strFolder = CompletePath(strFolder)
50060
50070  If LenB(Dir(strFolder & GsDll, vbNormal)) = 0 Then
50080   MsgBox LanguageStrings.MessagesMsg15
50090   Exit Sub
50100  End If
50110 ' UnloadDLLComplete GsDllLoaded
50120 ' GsDllLoaded = LoadDLL(strFolder & GsDll)
50130 ' If GsDllLoaded = 0 Then
50140 '   MsgBox LanguageStrings.MessagesMsg15
50150 '   Exit Sub
50160 '  Else
50170 '   UnLoadDLL GsDllLoaded
50180 ' End If
50190
50200  If UCase$(CompletePath(Options.DirectoryGhostscriptBinaries)) <> UCase$(CompletePath(strFolder)) Then
50210   aw = MsgBox("The program must be restarted!", vbOKCancel)
50220   If aw = vbCancel Then
50230    Exit Sub
50240   End If
50250   txtGSbin.Text = strFolder
50260   GetOptions Me, Options
50270   SaveOptions Options
50280   Restart = True
50290   Unload Me
50300  End If
50310 ' Options.DirectoryGhostscriptBinaries = strFolder
50320  txtGSbin.Text = strFolder
50330  With txtGSbin
50340   .ToolTipText = .Text
50350  End With
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
50030  strFolder = BrowseForFolder(Me.hwnd, LanguageStrings.OptionsGhostscriptFontsDirectoryPrompt)
50040  If Len(strFolder) = 0 Then Exit Sub
50050  strFolder = CompletePath(strFolder)
50060
50070  If Len(Dir(strFolder & "*.afm", vbNormal)) = 0 And Len(Dir(strFolder & "*.pfb", vbNormal)) = 0 Then
50080   MsgBox LanguageStrings.MessagesMsg16
50090   Exit Sub
50100  End If
50110
50120  txtGSfonts.Text = strFolder
50130  With txtGSfonts
50140   .ToolTipText = .Text
50150  End With
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
50030  strFolder = BrowseForFolder(Me.hwnd, LanguageStrings.OptionsGhostscriptLibrariesDirectoryPrompt)
50040  If Len(strFolder) = 0 Then Exit Sub
50050  strFolder = CompletePath(strFolder)
50060
50070  If Len(Dir(strFolder & "*.*", vbNormal)) = 0 Then
50080   MsgBox LanguageStrings.MessagesMsg17
50090   Exit Sub
50100  End If
50110
50120  txtGSlib.Text = strFolder
50130  With txtGSlib
50140   .ToolTipText = .Text
50150  End With
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
50030  strFolder = BrowseForFolder(Me.hwnd, LanguageStrings.OptionsPrintertempDirectoryPrompt)
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
50010  Dim tRestart As Boolean
50020  tRestart = False
50030  If UCase$(Options.DirectoryGhostscriptBinaries) <> UCase$(txtGSbin.Text) Then
50040   tRestart = True
50050  End If
50060  GetOptions Me, Options
50070  SaveOptions Options
50080  If IsWin9xMe = False Then
50091   Select Case Options.ProcessPriority
               Case 0: 'Idle
50110     SetProcessPriority Idle
50120    Case 1: 'Normal
50130     SetProcessPriority Normal
50140    Case 2: 'High
50150     SetProcessPriority High
50160    Case 3: 'Realtime
50170     SetProcessPriority RealTime
50180   End Select
50190  End If
50200  If tRestart = True Then
50210   Restart = True
50220  End If
50230  Unload Me
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
50010  Dim TestPSPage As String, fn As Long, Filename As String
50020  TestPSPage = LoadResString(3000)
50030  TestPSPage = Replace(TestPSPage, "[TESTPAGE]", LanguageStrings.OptionsTestpage, , 1, vbTextCompare)
50040  TestPSPage = Replace(TestPSPage, "[DATE]", Now, , 1, vbTextCompare)
50050  TestPSPage = Replace(TestPSPage, "[PDFCREATORVERSION]", App.Title & " " & GetProgramReleaseStr, , 1, vbTextCompare)
50060
50070  fn = FreeFile
50080  Filename = GetTempFile(GetPDFCreatorTempfolder, "~PD")
50090  Open Filename For Output As fn
50100  Print #fn, TestPSPage
50110  Close #fn
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
50020  Temppath = CompletePath(GetTempPath)
50030  If DirExists(Temppath) = False Then
50040   MakePath Temppath
50050  End If
50060  With txtTemppath
50070   .Text = Temppath
50080   .ToolTipText = .Text
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
50031   Select Case trv.SelectedItem.Key
               Case "Program"
50050     Call HTMLHelp_ShowTopic("html\generalsettings.htm")
50060    Case "ProgramGeneral"
50070     Call HTMLHelp_ShowTopic("html\generalsettings.htm")
50080    Case "ProgramSave"
50090     Call HTMLHelp_ShowTopic("html\savesettings.htm")
50100    Case "ProgramAutosave"
50110     Call HTMLHelp_ShowTopic("html\autosave.htm")
50120    Case "ProgramFonts"
50130     Call HTMLHelp_ShowTopic("html\fontsetting.htm")
50140    Case "ProgramDirectories"
50150     Call HTMLHelp_ShowTopic("html\directories.htm")
50160    Case "ProgramDocument"
50170     Call HTMLHelp_ShowTopic("html\docproperties.htm")
50180    Case "Formats"
50190     Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
50200    Case "FormatsPDF"
50211     Select Case tbstrPDFOptions.SelectedItem.Index
                 Case 1:
50230       Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
50240      Case 2:
50250       Call HTMLHelp_ShowTopic("html\pdfcompression.htm")
50260      Case 3:
50270       Call HTMLHelp_ShowTopic("html\pdffonts.htm")
50280      Case 4:
50290       Call HTMLHelp_ShowTopic("html\pdfcolors.htm")
50300      Case 5:
50310       Call HTMLHelp_ShowTopic("html\pdfsecurity.htm")
50320     End Select
50330    Case "FormatsPNG"
50340     Call HTMLHelp_ShowTopic("html\pngsettings.htm")
50350    Case "FormatsJPEG"
50360     Call HTMLHelp_ShowTopic("html\jpegsettings.htm")
50370    Case "FormatsBMP"
50380     Call HTMLHelp_ShowTopic("html\bmpsettings.htm")
50390    Case "FormatsPCX"
50400     Call HTMLHelp_ShowTopic("html\pcxsettings.htm")
50410    Case "FormatsTIFF"
50420     Call HTMLHelp_ShowTopic("html\tiffsettings.htm")
50430    Case "FormatsPS"
50440     Call HTMLHelp_ShowTopic("html\pssettings.htm")
50450    Case "FormatsEPS"
50460     Call HTMLHelp_ShowTopic("html\epssettings.htm")
50470   End Select
50480  End If
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
  cSystem As clsSystem, fi As Long, fc As Long, SMF As Collection, _
  reg As clsRegistry, tsf() As String, tStr2 As String
50050
50060  Me.KeyPreview = True
50070  Set cSystem = New clsSystem: Set SMF = cSystem.GetSystemFont(Me, Menu)
50080
50090  Screen.MousePointer = vbHourglass
50100
50110  fraBitmapGeneral.Visible = False
50120  fraBitmapGeneral.Top = fraPDFTop - 300
50130  fraBitmapGeneral.Left = fraPDFLeft - 200
50140  fraProgFont.Top = fraPDFTop - 300
50150  fraProgFont.Left = fraPDFLeft - 200
50160  fraProgGeneral.Top = fraPDFTop - 300
50170  fraProgGeneral.Left = fraPDFLeft - 200
50180  fraProgAutosave.Top = fraPDFTop - 300
50190  fraProgAutosave.Left = fraPDFLeft - 200
50200  fraProgSave.Top = fraPDFTop - 300
50210  fraProgSave.Left = fraPDFLeft - 200
50220  fraProgDirectories.Top = fraPDFTop - 300
50230  fraProgDirectories.Left = fraPDFLeft - 200
50240  fraProgDocument.Top = fraPDFTop - 300
50250  fraProgDocument.Left = fraPDFLeft - 200
50260  fraProgGhostscript.Top = fraPDFTop - 300
50270  fraProgGhostscript.Left = fraPDFLeft - 200
50280
50290  fraPDFSecurity.Top = fraPDFTop + 100
50300  fraPDFSecurity.Left = fraPDFLeft
50310  fraPDFFonts.Top = fraPDFTop + 100
50320  fraPDFFonts.Left = fraPDFLeft
50330  fraPDFColors.Top = fraPDFTop + 100
50340  fraPDFColors.Left = fraPDFLeft
50350  fraPDFGeneral.Top = fraPDFTop + 100
50360  fraPDFGeneral.Left = fraPDFLeft
50370  fraPDFCompress.Top = fraPDFTop + 100
50380  fraPDFCompress.Left = fraPDFLeft
50390  fraPSGeneral.Left = fraPDFLeft
50400  fraPSGeneral.Top = fraPDFTop - 300
50410  fraPSGeneral.Left = fraPDFLeft - 200
50420  tbstrPDFOptions.Top = fraPDFTop - 300
50430  tbstrPDFOptions.Left = fraPDFLeft - 200
50440  tbstrPDFOptions.Height = 5975
50450  tbstrPDFOptions.Width = 6215
50460
50470  cmbEPSLanguageLevel.Top = cmbPSLanguageLevel.Top
50480  cmbEPSLanguageLevel.Left = cmbPSLanguageLevel.Left
50490
50500  txtTest.Text = vbNullString
50510  For i = 33 To 255
50520   txtTest.Text = txtTest.Text & Chr$(i)
50530  Next i
50540  fi = -1
50550  With cmbFonts
50560   .Clear
50570   For i = 1 To Screen.FontCount
50580    tStr = Trim$(Screen.Fonts(i))
50590    If Len(tStr) > 0 Then
50600     cmbFonts.AddItem tStr
50610    End If
50620   Next i
50630   If .ListCount > 0 Then
50640     For i = 0 To cmbFonts.ListCount - 1
50650      If SMF.Count > 0 Then
50660       If UCase$(cmbFonts.List(i)) = UCase$(SMF(1)(0)) Then
50670        fi = i
50680       End If
50690      End If
50700     Next i
50710    Else
50720    .ListIndex = 0
50730   End If
50740  End With
50750  With cmbCharset
50760   .Clear
50770   .AddItem "0, Western": .ItemData(.NewIndex) = 0
50780   .AddItem "2, Symbol": .ItemData(.NewIndex) = 2
50790   .AddItem "77, Mac": .ItemData(.NewIndex) = 77
50800   .AddItem "161, Greek": .ItemData(.NewIndex) = 161
50810   .AddItem "162, Turkish": .ItemData(.NewIndex) = 162
50820   .AddItem "177, Hebrew": .ItemData(.NewIndex) = 177
50830   .AddItem "178, Arabic": .ItemData(.NewIndex) = 178
50840   .AddItem "186, Baltic": .ItemData(.NewIndex) = 186
50850   .AddItem "204, Cyrillic": .ItemData(.NewIndex) = 204
50860   .AddItem "238, Central European": .ItemData(.NewIndex) = 238
50870   .AddItem "255, DOS/OEM": .ItemData(.NewIndex) = 255
50880   .Text = 0
50890  End With
50900  If fi >= 0 Then
50910   cmbFonts.ListIndex = fi
50920   cmbCharset.Text = SMF(1)(2)
50930   txtProgramFontsize.Text = SMF(1)(1)
50940   txtTest.Font = cmbFonts.List(cmbFonts.ListIndex)
50950   txtTest.Font.Charset = cmbCharset.Text
50960  End If
50970
50980
50990  trv.Nodes.Clear
51000  trv.Indentation = 200
51010  With LanguageStrings
51020   trv.Nodes.Add , , "Program", .OptionsTreeProgram
51030   trv.Nodes.Add "Program", tvwChild, "ProgramGeneral", .OptionsProgramGeneralSymbol
51040   trv.Nodes.Add "Program", tvwChild, "ProgramGhostscript", .OptionsProgramGhostscriptSymbol
51050   trv.Nodes.Add "Program", tvwChild, "ProgramDocument", .OptionsProgramDocumentSymbol
51060   trv.Nodes.Add "Program", tvwChild, "ProgramSave", .OptionsProgramSaveSymbol
51070   trv.Nodes.Add "Program", tvwChild, "ProgramAutosave", .OptionsProgramAutosaveSymbol
51080   trv.Nodes.Add "Program", tvwChild, "ProgramDirectories", .OptionsProgramDirectoriesSymbol
51090   trv.Nodes.Add "Program", tvwChild, "ProgramFonts", .OptionsProgramFontSymbol
51100   trv.Nodes.Add , , "Formats", .OptionsTreeFormats
51110   trv.Nodes.Add "Formats", tvwChild, "FormatsPDF", .OptionsPDFSymbol
51120   trv.Nodes.Add "Formats", tvwChild, "FormatsPNG", .OptionsPNGSymbol
51130   trv.Nodes.Add "Formats", tvwChild, "FormatsJPEG", .OptionsJPEGSymbol
51140   trv.Nodes.Add "Formats", tvwChild, "FormatsBMP", .OptionsBMPSymbol
51150   trv.Nodes.Add "Formats", tvwChild, "FormatsPCX", .OptionsPCXSymbol
51160   trv.Nodes.Add "Formats", tvwChild, "FormatsTIFF", .OptionsTIFFSymbol
51170   trv.Nodes.Add "Formats", tvwChild, "FormatsPS", .OptionsPSSymbol
51180   trv.Nodes.Add "Formats", tvwChild, "FormatsEPS", .OptionsEPSSymbol
51190
51200   trv.Nodes("ProgramFonts").EnsureVisible
51210   trv.Nodes("FormatsPDF").EnsureVisible
51220
51230   Set picOptions = LoadResPicture(2101, vbResIcon)
51240   fraProgFont.Visible = False
51250   fraProgGeneral.Visible = True
51260
51270   fraProgGeneral.Caption = .OptionsProgramGeneralSymbol
51280   fraProgGhostscript.Caption = .OptionsProgramGhostscriptSymbol
51290   fraProgAutosave.Caption = .OptionsProgramAutosaveSymbol
51300   fraProgFont.Caption = .OptionsProgramFontSymbol
51310   fraProgDirectories.Caption = .OptionsProgramDirectoriesSymbol
51320   fraProgSave.Caption = .OptionsProgramSaveSymbol
51330   fraProgDocument.Caption = .OptionsProgramDocumentSymbol
51340
51350   fraShellintegration.Caption = .OptionsShellIntegration
51360   cmdShellintegration(0).Caption = .OptionsShellIntegrationAdd
51370   cmdShellintegration(1).Caption = .OptionsShellIntegrationRemove
51380   If IsWin9xMe = False Then
51390    If IsAdmin = False Then
51400     cmdShellintegration(0).Enabled = False
51410     cmdShellintegration(1).Enabled = False
51420    End If
51430   End If
51440
51450   lblGhostscriptversion.Caption = .OptionsGhostscriptversion
51460
51470   lblSaveFilename.Caption = .OptionsSaveFilename
51480   lblSaveFilenameTokens.Caption = .OptionsSaveFilenameTokens
51490   fraFilenameSubstitutions.Caption = .OptionsSaveFilenameSubstitutions
51500   chkFilenameSubst.Caption = .OptionsSaveFilenameSubstitutionsTitle
51510   cmdFilenameSubst(0).Caption = .OptionsSaveFilenameAdd
51520   cmdFilenameSubst(1).Caption = .OptionsSaveFilenameChange
51530   cmdFilenameSubst(2).Caption = .OptionsSaveFilenameDelete
51540
51550   chkSpaces.Caption = .OptionsRemoveSpaces
51560   chkNoConfirmMessageSwitchingDefaultprinter.Caption = .OptionsProgramSwitchingDefaultprinter
51570   lblGSbin.Caption = .OptionsDirectoriesGSBin
51580   lblGSlib.Caption = .OptionsDirectoriesGSLibraries
51590   lblGSfonts.Caption = .OptionsDirectoriesGSFonts
51600   lblPrintTempPath.Caption = .OptionsDirectoriesTempPath
51610
51620   lblOptions = .OptionsProgramGeneralDescription
51630   lblAutosaveformat.Caption = .OptionsAutosaveFormat
51640   chkUseStandardAuthor.Caption = .OptionsUseStandardauthor
51650   chkUseCreationDateNow.Caption = .OptionsUseCreationDateNow
51660   chkUseAutosave.Caption = .OptionsUseAutosave
51670   cmdTestpage.Caption = .OptionsPrintTestpage
51680   lblAutosaveFilename.Caption = .OptionsAutosaveFilename
51690   lblAutosaveFilenameTokens.Caption = .OptionsAutosaveFilenameTokens
51700   chkUseAutosaveDirectory.Caption = .OptionsUseAutosaveDirectory
51710   lblAuthorTokens.Caption = .OptionsStandardAuthorToken
51720
51730   With cmbAutoSaveFilenameTokens
51740    .AddItem "<Author>"
51750    .AddItem "<Computername>"
51760    .AddItem "<DateTime>"
51770    .AddItem "<Title>"
51780    .AddItem "<Username>"
51790    .ListIndex = 0
51800   End With
51810   With cmbSaveFilenameTokens
51820    .AddItem "<Author>"
51830    .AddItem "<Computername>"
51840    .AddItem "<DateTime>"
51850    .AddItem "<Title>"
51860    .AddItem "<Username>"
51870    .ListIndex = 0
51880   End With
51890   With cmbAuthorTokens
51900    .AddItem "<Computername>"
51910    .AddItem "<DateTime>"
51920    .AddItem "<Title>"
51930    .AddItem "<Username>"
51940    .ListIndex = 0
51950   End With
51960   With cmbAutosaveFormat
51970    .AddItem "PDF"
51980    .AddItem "PNG"
51990    .AddItem "JPEG"
52000    .AddItem "BMP"
52010    .AddItem "PCX"
52020    .AddItem "TIFF"
52030    .AddItem "PS"
52040    .AddItem "EPS"
52050   End With
52060   Me.Caption = .OptionsPDFOptions
52070   cmdCancel.Caption = .OptionsCancel
52080   cmdReset.Caption = .OptionsReset
52090   cmdSave.Caption = .OptionsSave
52100   tbstrPDFOptions.Tabs.Clear
52110   tbstrPDFOptions.Tabs.Add , "General", .OptionsPDFGeneral
52120   tbstrPDFOptions.Tabs.Add , "Compression", .OptionsPDFCompression
52130   tbstrPDFOptions.Tabs.Add , "Fonts", .OptionsPDFFonts
52140   tbstrPDFOptions.Tabs.Add , "Colors", .OptionsPDFColors
52150   tbstrPDFOptions.Tabs.Add , "Security", .OptionsPDFSecurity
52160   fraPDFGeneral.Caption = .OptionsPDFGeneralCaption
52170   lblPDFCompat.Caption = .OptionsPDFGeneralCompatibility
52180   lblPDFAutoRotate.Caption = .OptionsPDFGeneralAutorotate
52190   lblPDFResolution.Caption = .OptionsPDFGeneralResolution
52200   lblPDFOverprint.Caption = .OptionsPDFGeneralOverprint
52210   lblProgfont.Caption = .OptionsProgramFont
52220   lblProgcharset.Caption = .OptionsProgramFontcharset
52230   lblSize.Caption = .OptionsProgramFontSize
52240   lblTesttext = .OptionsProgramFontTestdescription
52250   cmdTest.Caption = .OptionsProgramFontTest
52260   cmdCancelTest.Caption = .OptionsProgramFontCancelTest
52270   chkPDFASCII85.Caption = .OptionsPDFGeneralASCII85
52280   cmbPDFCompat.Clear
52290   cmbPDFCompat.AddItem .OptionsPDFGeneralCompatibility01
52300   cmbPDFCompat.AddItem .OptionsPDFGeneralCompatibility02
52310   cmbPDFCompat.AddItem .OptionsPDFGeneralCompatibility03
52320   cmbPDFRotate.Clear
52330   cmbPDFRotate.AddItem .OptionsPDFGeneralRotate01
52340   cmbPDFRotate.AddItem .OptionsPDFGeneralRotate02
52350   cmbPDFRotate.AddItem .OptionsPDFGeneralRotate03
52360   cmbPDFOverprint.Clear
52370   cmbPDFOverprint.AddItem .OptionsPDFGeneralOverprint01
52380   cmbPDFOverprint.AddItem .OptionsPDFGeneralOverprint02
52390
52400   fraPDFCompress.Caption = .OptionsPDFCompressionCaption
52410   chkPDFTextComp.Caption = .OptionsPDFCompressionTextComp
52420   fraPDFColor.Caption = .OptionsPDFCompressionColor
52430   chkPDFColorComp.Caption = .OptionsPDFCompressionColorComp
52440   chkPDFColorResample.Caption = .OptionsPDFCompressionColorResample
52450   lblPDFColorRes.Caption = .OptionsPDFCompressionColorRes
52460   cmbPDFColorComp.Clear
52470   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp01
52480   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp02
52490   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp03
52500   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp04
52510   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp05
52520   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp06
52530   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp07
52540   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp08
52550   cmbPDFColorResample.Clear
52560   cmbPDFColorResample.AddItem .OptionsPDFCompressionColorResample01
52570   cmbPDFColorResample.AddItem .OptionsPDFCompressionColorResample02
52580   cmbPDFColorResample.AddItem .OptionsPDFCompressionColorResample03
52590   fraPDFGrey.Caption = .OptionsPDFCompressionGrey
52600   chkPDFGreyComp.Caption = .OptionsPDFCompressionGreyComp
52610   chkPDFGreyResample.Caption = .OptionsPDFCompressionGreyResample
52620   lblPDFGreyRes.Caption = .OptionsPDFCompressionGreyRes
52630   cmbPDFGreyComp.Clear
52640   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp01
52650   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp02
52660   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp03
52670   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp04
52680   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp05
52690   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp06
52700   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp07
52710   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp08
52720   cmbPDFGreyResample.Clear
52730   cmbPDFGreyResample.AddItem .OptionsPDFCompressionGreyResample01
52740   cmbPDFGreyResample.AddItem .OptionsPDFCompressionGreyResample02
52750   cmbPDFGreyResample.AddItem .OptionsPDFCompressionGreyResample03
52760   fraPDFMono.Caption = .OptionsPDFCompressionMono
52770   chkPDFMonoComp.Caption = .OptionsPDFCompressionMonoComp
52780   chkPDFMonoResample.Caption = .OptionsPDFCompressionMonoResample
52790   lblPDFMonoRes.Caption = .OptionsPDFCompressionMonoRes
52800   cmbPDFMonoComp.Clear
52810   cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp01
52820   cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp02
52830   cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp03
52840   cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp04
52850   cmbPDFMonoResample.Clear
52860   cmbPDFMonoResample.AddItem .OptionsPDFCompressionMonoResample01
52870   cmbPDFMonoResample.AddItem .OptionsPDFCompressionMonoResample02
52880   cmbPDFMonoResample.AddItem .OptionsPDFCompressionMonoResample03
52890
52900   fraPDFFonts.Caption = .OptionsPDFFontsCaption
52910   chkPDFEmbedAll.Caption = .OptionsPDFFontsEmbedAll
52920   chkPDFSubSetFonts.Caption = .OptionsPDFFontsSubSetFonts
52930
52940   fraPDFColors.Caption = .OptionsPDFColorsCaption
52950   chkPDFCMYKtoRGB.Caption = .OptionsPDFColorsCMYKtoRGB
52960   fraPDFColorOptions.Caption = .OptionsPDFColorsColorOptions
52970   chkPDFPreserveOverprint.Caption = .OptionsPDFColorsPreserveOverprint
52980   chkPDFPreserveTransfer.Caption = .OptionsPDFColorsPreserveTransfer
52990   chkPDFPreserveHalftone.Caption = .OptionsPDFColorsPreserveHalftone
53000   cmbPDFColorModel.Clear
53010   cmbPDFColorModel.AddItem .OptionsPDFColorsColorModel01
53020   cmbPDFColorModel.AddItem .OptionsPDFColorsColorModel02
53030   cmbPDFColorModel.AddItem .OptionsPDFColorsColorModel03
53040
53050   fraPDFEncryptor.Caption = .OptionsPDFEncryptor
53060   fraPDFSecurity.Caption = .OptionsPDFSecurityCaption
53070   chkUseSecurity.Caption = .OptionsPDFUseSecurity
53080   fraPDFEncLevel.Caption = .OptionsPDFEncryptionLevel
53090   optEncHigh.Caption = .OptionsPDFEncryptionHigh
53100   optEncLow.Caption = .OptionsPDFEncryptionLow
53110   fraSecurityPass.Caption = .OptionsPDFPasswords
53120   chkUserPass.Caption = .OptionsPDFUserPass
53130   chkOwnerPass.Caption = .OptionsPDFOwnerPass
53140   fraPDFPermissions.Caption = .OptionsPDFDisallowUser
53150   fraPDFHighPermissions.Caption = .OptionsPDFEnhancedPermissions
53160   chkAllowPrinting.Caption = .OptionsPDFDisallowPrint
53170   chkAllowModifyContents.Caption = .OptionsPDFDisallowModify
53180   chkAllowCopy.Caption = .OptionsPDFDisallowCopy
53190   chkAllowModifyAnnotations.Caption = .OptionsPDFDisallowModifyComments
53200   chkAllowDegradedPrinting.Caption = .OptionsPDFAllowDegradedPrinting
53210   chkAllowFillIn.Caption = .OptionsPDFAllowFillIn
53220   chkAllowAssembly.Caption = .OptionsPDFAllowAssembly
53230   chkAllowScreenReaders.Caption = .OptionsPDFAllowScreenReaders
53240
53250   cmbPNGColors.AddItem .OptionsPNGColorscount01
53260   cmbPNGColors.AddItem .OptionsPNGColorscount02
53270   cmbPNGColors.AddItem .OptionsPNGColorscount03
53280   cmbPNGColors.AddItem .OptionsPNGColorscount04
53290   cmbJPEGColors.Left = cmbPNGColors.Left
53300   cmbJPEGColors.Width = cmbPNGColors.Width
53310   cmbJPEGColors.Top = cmbPNGColors.Top
53320   cmbJPEGColors.AddItem .OptionsJPEGColorscount01
53330   cmbJPEGColors.AddItem .OptionsJPEGColorscount02
53340   cmbBMPColors.Left = cmbPNGColors.Left
53350   cmbBMPColors.Width = cmbPNGColors.Width
53360   cmbBMPColors.Top = cmbPNGColors.Top
53370   cmbBMPColors.AddItem .OptionsBMPColorscount01
53380   cmbBMPColors.AddItem .OptionsBMPColorscount02
53390   cmbBMPColors.AddItem .OptionsBMPColorscount03
53400   cmbBMPColors.AddItem .OptionsBMPColorscount04
53410   cmbBMPColors.AddItem .OptionsBMPColorscount05
53420   cmbBMPColors.AddItem .OptionsBMPColorscount06
53430   cmbBMPColors.AddItem .OptionsBMPColorscount07
53440   cmbPCXColors.Left = cmbPNGColors.Left
53450   cmbPCXColors.Width = cmbPNGColors.Width
53460   cmbPCXColors.Top = cmbPNGColors.Top
53470   cmbPCXColors.AddItem .OptionsPCXColorscount01
53480   cmbPCXColors.AddItem .OptionsPCXColorscount02
53490   cmbPCXColors.AddItem .OptionsPCXColorscount03
53500   cmbPCXColors.AddItem .OptionsPCXColorscount04
53510   cmbPCXColors.AddItem .OptionsPCXColorscount05
53520   cmbPCXColors.AddItem .OptionsPCXColorscount06
53530   cmbTIFFColors.Left = cmbPNGColors.Left
53540   cmbTIFFColors.Width = cmbPNGColors.Width
53550   cmbTIFFColors.Top = cmbPNGColors.Top
53560   cmbTIFFColors.AddItem .OptionsTIFFColorscount01
53570   cmbTIFFColors.AddItem .OptionsTIFFColorscount02
53580   cmbTIFFColors.AddItem .OptionsTIFFColorscount03
53590   cmbTIFFColors.AddItem .OptionsTIFFColorscount04
53600   cmbTIFFColors.AddItem .OptionsTIFFColorscount05
53610   cmbTIFFColors.AddItem .OptionsTIFFColorscount06
53620   cmbTIFFColors.AddItem .OptionsTIFFColorscount07
53630   cmbTIFFColors.AddItem .OptionsTIFFColorscount08
53640
53650   fraBitmapGeneral.Caption = .OptionsImageSettings
53660   lblBitmapResolution = .OptionsBitmapResolution
53670   lblJPEGQuality = .OptionsJPEGQuality
53680   lblBitmapColors = .OptionsPDFColors
53690   lblProcessPriority.Caption = .OptionsProcesspriority
53700   lblLangLevel.Caption = .OptionsPSLanguageLevel
53710
53720   cmdAsso.Caption = .OptionsAssociatePSFiles
53730  End With
53740
53750  If IsPsAssociate = False Then
53760    cmdAsso.Enabled = True
53770   Else
53780    cmdAsso.Enabled = False
53790  End If
53800
53810  txtPDFRes.Text = 600
53820  cmbPDFCompat.ListIndex = 1
53830  cmbPDFRotate.ListIndex = 0
53840  cmbPDFOverprint.ListIndex = 0
53850  chkPDFASCII85.Value = 0
53860
53870  chkPDFTextComp.Value = 1
53880
53890  chkPDFColorComp.Value = 1
53900  chkPDFColorResample.Value = 0
53910  cmbPDFColorComp.ListIndex = 0
53920  cmbPDFColorResample.ListIndex = 0
53930  txtPDFColorRes.Text = 300
53940
53950  chkPDFGreyComp.Value = 1
53960  chkPDFGreyResample.Value = 0
53970  cmbPDFGreyComp.ListIndex = 0
53980  cmbPDFGreyResample.ListIndex = 0
53990  txtPDFGreyRes.Text = 300
54000
54010  chkPDFMonoComp.Value = 1
54020  chkPDFMonoResample.Value = 0
54030  cmbPDFMonoComp.ListIndex = 0
54040  cmbPDFMonoResample.ListIndex = 0
54050  txtPDFMonoRes.Text = 1200
54060
54070  chkPDFEmbedAll.Value = 1
54080  chkPDFSubSetFonts.Value = 1
54090  txtPDFSubSetPerc.Text = 100
54100
54110  cmbPDFColorModel.ListIndex = 1
54120  chkPDFCMYKtoRGB.Value = 1
54130  chkPDFPreserveOverprint.Value = 1
54140  chkPDFPreserveTransfer.Value = 1
54150  chkPDFPreserveHalftone.Value = 0
54160
54170  cmbPNGColors.ListIndex = 0
54180  cmbJPEGColors.ListIndex = 0
54190  cmbBMPColors.ListIndex = 0
54200  cmbPCXColors.ListIndex = 0
54210  cmbTIFFColors.ListIndex = 0
54220  txtBitmapResolution.Text = 150
54230
54240  cmbCharset.ListIndex = 0
54250  txtProgramFontsize.Text = 8
54260
54270 ' chkUseStandardAuthor.Value = 1
54280  txtStandardAuthor.Text = vbNullString
54290
54300  With cmbPSLanguageLevel
54310   .AddItem "1"
54320   .AddItem "1.5"
54330   .AddItem "2"
54340   .AddItem "3"
54350  End With
54360  With cmbEPSLanguageLevel
54370   .AddItem "1"
54380   .AddItem "1.5"
54390   .AddItem "2"
54400   .AddItem "3"
54410  End With
54420
54430  With lsvFilenameSubst
54440   .Appearance = ccFlat
54450   .ColumnHeaders.Clear
54460   .ColumnHeaders.Add , "Str1", "", lsvFilenameSubst.Width / 2 - 140
54470   .ColumnHeaders.Add , "Str2", "", lsvFilenameSubst.Width / 2 - 140
54480   .HideColumnHeaders = True
54490   .GridLines = True
54500   .FullRowSelect = True
54510   .HideSelection = False
54520  End With
54530
54540  With cmbPDFEncryptor
54550   .Clear
54560   .AddItem "Ghostscript (>= 8.14)"
54570   .ItemData(.NewIndex) = 0
54580   .AddItem "PDFEnc"
54590   .ItemData(.NewIndex) = 1
54600
54610   ShowOptions Me, Options
54620
54630   SecurityIsPossible = True
54640
54650   If LenB(Dir(CompletePath(App.Path) & "pdfenc.exe")) = 0 Then
54660    .RemoveItem 1
54670    .ListIndex = 0
54680    Options.PDFEncryptor = .ItemData(.ListIndex)
54690   End If
54700   If GhostScriptSecurity = False Then
54710    .RemoveItem 0
54720   End If
54730   If .ListCount = 0 Then
54740     chkUseSecurity.Value = 0
54750     chkUseSecurity.Enabled = False
54760     SecurityIsPossible = False
54770    Else
54780     For i = 0 To .ListCount - 1
54790      If .ItemData(i) = Options.PDFEncryptor Then
54800       .ListIndex = i
54810       Exit For
54820      End If
54830     Next i
54840     If .ListIndex = -1 Then
54850      .ListIndex = 0
54860      Options.PDFEncryptor = .ItemData(.ListIndex)
54870     End If
54880   End If
54890  End With
54900
54910
54920  If Options.PDFHighEncryption <> 0 Then
54930    optEncHigh.Value = True
54940   Else
54950    optEncLow.Value = True
54960  End If
54970
54980  CheckCmdFilenameSubst
54990
55000  If chkUseStandardAuthor.Value = 1 Then
55010    txtStandardAuthor.Enabled = True
55020    txtStandardAuthor.BackColor = &H80000005
55030   Else
55040    txtStandardAuthor.Enabled = False
55050    txtStandardAuthor.BackColor = &H8000000F
55060  End If
55070  With Options
55080   SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
55090   cmbCharset.Text = .ProgramFontCharset
55100  End With
55110  If chkUseAutosave.Value = 1 Then
55120    ViewAutosave True
55130   Else
55140    ViewAutosave False
55150  End If
55160
55170  With txtGSbin
55180   .ToolTipText = .Text
55190  End With
55200  With txtGSlib
55210   .ToolTipText = .Text
55220  End With
55230  With txtGSfonts
55240   .ToolTipText = .Text
55250  End With
55260  With txtTemppath
55270   .ToolTipText = .Text
55280  End With
55290
55300  With sldProcessPriority
55310   .TextPosition = sldBelowRight
55320   .TickFrequency = 1
55330   .TickStyle = sldTopLeft
55341   Select Case .Value
         Case 0: 'Idle
55360     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityIdle
55370    Case 1: 'Normal
55380     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityNormal
55390    Case 2: 'High
55400     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityHigh
55410    Case 3: 'Realtime
55420     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityRealtime
55430   End Select
55440  End With
55450
55460  If IsWin9xMe = False Then
55470    lblProcessPriority.Enabled = True
55480    sldProcessPriority.Enabled = True
55490   Else
55500    lblProcessPriority.Enabled = False
55510    sldProcessPriority.Enabled = False
55520  End If
55530  UpdateSecurityFields
55540
55550  tStr2 = CompletePath(UCase$(Trim$(Options.DirectoryGhostscriptBinaries)))
 cmbGhostscript.Clear: Set reg = New clsRegistry
55570  reg.hkey = HKEY_LOCAL_MACHINE
55580
55590  Set gsvers = GetAllGhostscriptversions
55600
55610  If gsvers.Count = 0 Then
55620    cmbGhostscript.Enabled = False
55630   Else
55640    For i = 1 To gsvers.Count
55650     cmbGhostscript.AddItem gsvers.item(i)
55660    Next i
55670    cmbGhostscript.ListIndex = cmbGhostscript.ListCount - 1
55680    For i = 0 To cmbGhostscript.ListCount - 1
55690     tStr = ""
55700     If InStr(cmbGhostscript.List(i), ":") Then
55710       reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
55720       If tStr2 = CompletePath(UCase$(Trim$(reg.GetRegistryValue("GhostscriptDirectoryBinaries")))) Then
55730        cmbGhostscript.ListIndex = i
55740        Exit For
55750       End If
55760      Else
55770       If InStr(UCase$(cmbGhostscript.List(i)), "AFPL") Then
55780        reg.KeyRoot = "SOFTWARE\AFPL Ghostscript"
55790        If InStr(cmbGhostscript.List(i), " ") > 0 Then
55800         tsf = Split(cmbGhostscript.List(i), " ")
55810         reg.Subkey = tsf(UBound(tsf))
55820         tStr = reg.GetRegistryValue("GS_DLL")
55830         If tStr2 & "GSDLL32.DLL" = UCase$(tStr) Then
55840          cmbGhostscript.ListIndex = i
55850          Exit For
55860         End If
55870        End If
55880       End If
55890       If InStr(UCase$(cmbGhostscript.List(i)), "GNU") Then
55900        reg.KeyRoot = "SOFTWARE\GNU Ghostscript"
55910        If InStr(cmbGhostscript.List(i), " ") > 0 Then
55920         tsf = Split(cmbGhostscript.List(i), " ")
55930         reg.Subkey = tsf(UBound(tsf))
55940         tStr = reg.GetRegistryValue("GS_DLL")
55950         If tStr2 & "GSDLL32.DLL" = UCase$(tStr) Then
55960          cmbGhostscript.ListIndex = i
55970          Exit For
55980         End If
55990        End If
56000       End If
56010       If InStr(UCase$(cmbGhostscript.List(i)), "GPL") Then
56020        reg.KeyRoot = "SOFTWARE\GPL Ghostscript"
56030        If InStr(cmbGhostscript.List(i), " ") > 0 Then
56040         tsf = Split(cmbGhostscript.List(i), " ")
56050         reg.Subkey = tsf(UBound(tsf))
56060         tStr = reg.GetRegistryValue("GS_DLL")
56070         If tStr2 & "GSDLL32.DLL" = UCase$(tStr) Then
56080          cmbGhostscript.ListIndex = i
56090          Exit For
56100         End If
56110        End If
56120       End If
56130     End If
56140    Next i
56150  End If
56160  Set reg = Nothing
56170  With cmbGhostscript
56180   If .ListCount = 0 Then
56190    .Enabled = False
56200    .BackColor = &H8000000F
56210   End If
56220  End With
56230  tbstrPDFOptions.ZOrder 1
56240  cmdStyle.ZOrder 1
56250  Screen.MousePointer = vbNormal
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
50250  fraProgGhostscript.Visible = False
50260
50271  Select Case trv.SelectedItem.Key
              Case "Program"
50290    Set picOptions = LoadResPicture(2101, vbResIcon)
50300    lblOptions = LanguageStrings.OptionsProgramGeneralDescription
50310    fraProgGeneral.Visible = True
50320   Case "ProgramGeneral"
50330    Set picOptions = LoadResPicture(2101, vbResIcon)
50340    lblOptions = LanguageStrings.OptionsProgramGeneralDescription
50350    fraProgGeneral.Visible = True
50360   Case "ProgramGhostscript"
50370    Set picOptions = LoadResPicture(2119, vbResIcon)
50380    lblOptions = LanguageStrings.OptionsProgramGhostscriptDescription
50390    fraProgGhostscript.Visible = True
50400   Case "ProgramSave"
50410    Set picOptions = LoadResPicture(2106, vbResIcon)
50420    lblOptions = LanguageStrings.OptionsProgramSaveDescription
50430    fraProgSave.Visible = True
50440   Case "ProgramAutosave"
50450    Set picOptions = LoadResPicture(2103, vbResIcon)
50460    lblOptions = LanguageStrings.OptionsProgramAutosaveDescription
50470    fraProgAutosave.Visible = True
50480   Case "ProgramFonts"
50490    Set picOptions = LoadResPicture(2102, vbResIcon)
50500    lblOptions = LanguageStrings.OptionsProgramFontDescription
50510    fraProgFont.Visible = True
50520   Case "ProgramDirectories"
50530    Set picOptions = LoadResPicture(2104, vbResIcon)
50540    lblOptions = LanguageStrings.OptionsProgramDirectoriesDescription
50550    fraProgDirectories.Visible = True
50560   Case "ProgramDocument"
50570    Set picOptions = LoadResPicture(2105, vbResIcon)
50580    lblOptions = LanguageStrings.OptionsProgramDocumentDescription
50590    fraProgDocument.Visible = True
50600   Case "Formats"
50610    Set picOptions = LoadResPicture(2111, vbResIcon)
50620    lblOptions = LanguageStrings.OptionsPDFDescription
50630    tbstrPDFOptions.Visible = True
50640    fraPDFGeneral.Visible = True
50650   Case "FormatsPDF"
50660    Set picOptions = LoadResPicture(2111, vbResIcon)
50670    lblOptions = LanguageStrings.OptionsPDFDescription
50680    tbstrPDFOptions.Visible = True
50690    tbstrPDFOptions.Tabs(1).Selected = True
50700    fraPDFGeneral.Visible = True
50710   Case "FormatsPNG"
50720    Set picOptions = LoadResPicture(2112, vbResIcon)
50730    lblOptions = LanguageStrings.OptionsPNGDescription
50740    fraBitmapGeneral.Visible = True
50750    cmbPNGColors.Visible = True
50760   Case "FormatsJPEG"
50770    Set picOptions = LoadResPicture(2113, vbResIcon)
50780    lblOptions = LanguageStrings.OptionsJPEGDescription
50790    fraBitmapGeneral.Visible = True
50800    lblJPEGQuality.Caption = LanguageStrings.OptionsJPEGQuality
50810    lblJPEGQuality.Visible = True
50820    txtJPEGQuality.Visible = True
50830    lblJPEQQualityProzent.Visible = True
50840    lblJPEQQualityProzent.Left = txtJPEGQuality.Left + txtJPEGQuality.Width + 100
50850    cmbJPEGColors.Visible = True
50860   Case "FormatsBMP"
50870    Set picOptions = LoadResPicture(2114, vbResIcon)
50880    lblOptions = LanguageStrings.OptionsBMPDescription
50890    fraBitmapGeneral.Visible = True
50900    cmbBMPColors.Visible = True
50910   Case "FormatsPCX"
50920    Set picOptions = LoadResPicture(2115, vbResIcon)
50930    lblOptions = LanguageStrings.OptionsPCXDescription
50940    fraBitmapGeneral.Visible = True
50950    cmbPCXColors.Visible = True
50960   Case "FormatsTIFF"
50970    Set picOptions = LoadResPicture(2116, vbResIcon)
50980    lblOptions = LanguageStrings.OptionsTIFFDescription
50990    fraBitmapGeneral.Visible = True
51000    cmbTIFFColors.Visible = True
51010   Case "FormatsPS"
51020    Set picOptions = LoadResPicture(2117, vbResIcon)
51030    lblOptions.Caption = LanguageStrings.OptionsPSDescription
51040    fraPSGeneral.Visible = True
51050    cmbPSLanguageLevel.Visible = True
51060    fraPSGeneral.Caption = LanguageStrings.OptionsPSDescription
51070   Case "FormatsEPS"
51080    Set picOptions = LoadResPicture(2118, vbResIcon)
51090    lblOptions.Caption = LanguageStrings.OptionsEPSDescription
51100    fraPSGeneral.Visible = True
51110    cmbEPSLanguageLevel.Visible = True
51120    fraPSGeneral.Caption = LanguageStrings.OptionsEPSDescription
51130  End Select
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
50010  fraPDFGeneral.Visible = False
50020  fraPDFCompress.Visible = False
50030  fraPDFFonts.Visible = False
50040  fraPDFColors.Visible = False
50050  fraPDFSecurity.Visible = False
50061  Select Case tbstrPDFOptions.SelectedItem.Index
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
50170    If SecurityIsPossible = False Then
50180     MsgBox LanguageStrings.MessagesMsg19
50190    End If
50200  End Select
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
50020  txtAutoSavePreview.Text = GetSubstFilename("C:\test.pdf", txtAutosaveFilename.Text, , True) & ".pdf"
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
50050  txtAutoSavePreview.Enabled = ViewIt
50060  lblAutosaveFilenameTokens.Enabled = ViewIt
50070  cmbAutoSaveFilenameTokens.Enabled = ViewIt
50080  chkUseAutosaveDirectory.Enabled = ViewIt
50090  If ViewIt = True Then
50100    cmbAutosaveFormat.BackColor = &H80000005
50110    cmbAutoSaveFilenameTokens.BackColor = &H80000005
50120    txtAutosaveFilename.BackColor = &H80000005
50130   Else
50140    cmbAutosaveFormat.BackColor = &H8000000F
50150    cmbAutoSaveFilenameTokens.BackColor = &H8000000F
50160    txtAutosaveFilename.BackColor = &H8000000F
50170  End If
50180  If chkUseAutosaveDirectory.Value = 1 And ViewIt = True Then
50190    ViewAutosaveDirectory True
50200   Else
50210    ViewAutosaveDirectory False
50220  End If
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
50020    fraPDFEncryptor.Enabled = False
50030    cmbPDFEncryptor.Enabled = False
50040
50050    fraPDFEncLevel.Enabled = False
50060    optEncHigh.Enabled = False
50070    optEncLow.Enabled = False
50080
50090    fraSecurityPass.Enabled = False
50100    chkUserPass.Enabled = False
50110    chkOwnerPass.Enabled = False
50120
50130    fraPDFPermissions.Enabled = False
50140    chkAllowPrinting.Enabled = False
50150    chkAllowCopy.Enabled = False
50160    chkAllowModifyAnnotations.Enabled = False
50170    chkAllowModifyContents.Enabled = False
50180
50190    fraPDFHighPermissions.Enabled = False
50200    chkAllowDegradedPrinting.Enabled = False
50210    chkAllowFillIn.Enabled = False
50220    chkAllowScreenReaders.Enabled = False
50230    chkAllowAssembly.Enabled = False
50240   Else
50250    fraPDFEncryptor.Enabled = True
50260    cmbPDFEncryptor.Enabled = True
50270
50280    fraPDFEncLevel.Enabled = True
50290    If cmbPDFCompat.ListIndex >= 2 Then
50300      optEncHigh.Enabled = True
50310     Else
50320      optEncHigh.Enabled = False
50330    End If
50340    optEncLow.Enabled = True
50350
50360    fraSecurityPass.Enabled = True
50370    chkUserPass.Enabled = True
50380    chkOwnerPass.Enabled = True
50390
50400    fraPDFPermissions.Enabled = True
50410    chkAllowPrinting.Enabled = True
50420    chkAllowCopy.Enabled = True
50430    chkAllowModifyAnnotations.Enabled = True
50440    chkAllowModifyContents.Enabled = True
50450
50460    If optEncHigh.Value = True Then
50470      fraPDFHighPermissions.Enabled = True
50480      chkAllowDegradedPrinting.Enabled = True
50490      chkAllowFillIn.Enabled = True
50500      chkAllowScreenReaders.Enabled = True
50510      chkAllowAssembly.Enabled = True
50520     Else
50530      fraPDFHighPermissions.Enabled = False
50540      chkAllowDegradedPrinting.Enabled = False
50550      chkAllowFillIn.Enabled = False
50560      chkAllowScreenReaders.Enabled = False
50570      chkAllowAssembly.Enabled = False
50580    End If
50590  End If
50600  If chkOwnerPass.Value = 0 And chkUserPass.Value = 0 Then
50610   chkOwnerPass.Value = 1: Options.PDFOwnerPass = 1
50620  End If
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
              Case 0:
50050    lsvFilenameSubst.ListItems.Add , , txtFilenameSubst(0).Text
50060    lsvFilenameSubst.ListItems(lsvFilenameSubst.ListItems.Count).SubItems(1) = txtFilenameSubst(1).Text
50070    lsvFilenameSubst.ListItems(lsvFilenameSubst.ListItems.Count).Selected = True
50080    lsvFilenameSubst.ListItems(lsvFilenameSubst.ListItems.Count).EnsureVisible
50090    Set_txtFilenameSubst
50100   Case 2:
50110    MsgBox LanguageStrings.MessagesMsg12 & _
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
              Case 0:
50050    lsvFilenameSubst.SelectedItem.Text = txtFilenameSubst(0).Text
50060    lsvFilenameSubst.SelectedItem.SubItems(1) = txtFilenameSubst(1).Text
50070   Case 2:
50080    MsgBox LanguageStrings.MessagesMsg12 & _
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

Private Sub txtSaveFilename_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txtSaveFilename.ToolTipText = txtSaveFilename.Text
50020  txtSavePreview.Text = GetSubstFilename("C:\test.pdf", txtSaveFilename.Text, , True) & ".pdf"
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
