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
            Picture         =   "frmOptions.frx":000C
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
            Picture         =   "frmOptions.frx":0396
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
         ItemData        =   "frmOptions.frx":0720
         Left            =   3720
         List            =   "frmOptions.frx":0722
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
         ItemData        =   "frmOptions.frx":0724
         Left            =   3720
         List            =   "frmOptions.frx":0726
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
            ItemData        =   "frmOptions.frx":0728
            Left            =   120
            List            =   "frmOptions.frx":072A
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
            ItemData        =   "frmOptions.frx":072C
            Left            =   2280
            List            =   "frmOptions.frx":072E
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
            ItemData        =   "frmOptions.frx":0730
            Left            =   2280
            List            =   "frmOptions.frx":0732
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
            ItemData        =   "frmOptions.frx":0734
            Left            =   120
            List            =   "frmOptions.frx":0736
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
            ItemData        =   "frmOptions.frx":0738
            Left            =   120
            List            =   "frmOptions.frx":073A
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
            ItemData        =   "frmOptions.frx":073C
            Left            =   2280
            List            =   "frmOptions.frx":073E
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
            ItemData        =   "frmOptions.frx":0740
            Left            =   240
            List            =   "frmOptions.frx":0742
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
         ItemData        =   "frmOptions.frx":0744
         Left            =   3690
         List            =   "frmOptions.frx":0746
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
         Picture         =   "frmOptions.frx":0748
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
         ItemData        =   "frmOptions.frx":0AD2
         Left            =   2400
         List            =   "frmOptions.frx":0AD4
         Style           =   2  'Dropdown-Liste
         TabIndex        =   100
         Tag             =   "None|All|PageByPage"
         Top             =   840
         Width           =   2655
      End
      Begin VB.ComboBox cmbPDFCompat 
         Height          =   315
         ItemData        =   "frmOptions.frx":0AD6
         Left            =   2400
         List            =   "frmOptions.frx":0AD8
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
         ItemData        =   "frmOptions.frx":0ADA
         Left            =   2400
         List            =   "frmOptions.frx":0ADC
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
   Begin VB.Frame fraPDFFonts 
      Caption         =   "Font Options"
      Height          =   2895
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
         ItemData        =   "frmOptions.frx":0ADE
         Left            =   120
         List            =   "frmOptions.frx":0AE0
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
50210      If InStrRev(Path, "\") > 0 Then
50220       txtGSlib.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50230       txtGSfonts.Text = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50240      End If
50250     End If
50260    End If
50270    If InStr(UCase$(gsv), "GNU") Then
50280     If InStr(gsv, " ") > 0 Then
50290      tsf = Split(gsv, " ")
50300      reg.KeyRoot = "SOFTWARE\GNU Ghostscript\" & tsf(UBound(tsf))
50310      tStr = reg.GetRegistryValue("GS_DLL")
50320      SplitPath tStr, , Path
50330      txtGSbin.Text = CompletePath(Path)
50340      If InStrRev(Path, "\") > 0 Then
50350       txtGSlib.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50360       txtGSfonts.Text = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50370      End If
50380     End If
50390    End If
50400    If InStr(UCase$(gsv), "GPL") Then
50410     If InStr(gsv, " ") > 0 Then
50420      tsf = Split(gsv, " ")
50430      reg.KeyRoot = "SOFTWARE\GPL Ghostscript\" & tsf(UBound(tsf))
50440      tStr = reg.GetRegistryValue("GS_DLL")
50450      SplitPath tStr, , Path
50460      txtGSbin.Text = CompletePath(Path)
50470      If InStrRev(Path, "\") > 0 Then
50480       txtGSlib.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50490       txtGSfonts.Text = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50500      End If
50510     End If
50520    End If
50530  End If
50540
50550  Set reg = Nothing
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
50050  If Len(Dir(strFolder & "*.afm", vbNormal)) = 0 And Len(Dir(strFolder & "*.pfb", vbNormal)) = 0 Then
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
50050  If Len(Dir(strFolder & "*.*", vbNormal)) = 0 Then
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
50060  CorrectCmbCharset
50070  GetOptions Me, Options
50080  SaveOptions Options
50090  If IsWin9xMe = False Then
50101   Select Case Options.ProcessPriority
               Case 0: 'Idle
50120     SetProcessPriority Idle
50130    Case 1: 'Normal
50140     SetProcessPriority Normal
50150    Case 2: 'High
50160     SetProcessPriority High
50170    Case 3: 'Realtime
50180     SetProcessPriority RealTime
50190   End Select
50200  End If
50210  If tRestart = True Then
50220   Restart = True
50230  End If
50240  Unload Me
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
50170  With cmdTest
50180   tFontname = .Fontname
50190   tFontSize = .Fontsize
50200   tFontCharset = .Font.Charset
50210  End With
50220  SetFont Me, cmbFonts.List(cmbFonts.ListIndex), CLng(tStr), cmbProgramFontsize.Text
50230  cmbCharset.Text = tCharset
50240  SetFont frmMain, cmbFonts.List(cmbFonts.ListIndex), CLng(tStr), cmbProgramFontsize.Text
50250  With cmdTest
50260   .Fontname = tFontname
50270   .Fontsize = tFontSize
50280   .Font.Charset = tFontCharset
50290  End With
50300  With cmdCancelTest
50310   .Fontname = tFontname
50320   .Fontsize = tFontSize
50330   .Font.Charset = tFontCharset
50340   .Enabled = True
50350  End With
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
50010  Dim TestPSPage As String, fn As Long, Filename As String, tStr As String
50020  TestPSPage = LoadResString(3000)
50030  TestPSPage = Replace(TestPSPage, "[TESTPAGE]", LanguageStrings.OptionsTestpage, , 1, vbTextCompare)
50040  TestPSPage = Replace(TestPSPage, "[DATE]", Now, , 1, vbTextCompare)
50050  TestPSPage = Replace(TestPSPage, "[PDFCREATORVERSION]", App.Title & " " & GetProgramReleaseStr, , 1, vbTextCompare)
50060
50070  fn = FreeFile
50080  tStr = CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory & "\" & GetUsername
50090  If DirExists(tStr) = False Then
50100   MakePath tStr
50110  End If
50120  Filename = GetTempFile(tStr, "~PD")
50130  Open Filename For Output As fn
50140  Print #fn, TestPSPage
50150  Close #fn
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
  reg As clsRegistry, tsf() As String, tStr2 As String, ctl As Control
50050
50060  KeyPreview = True
50070  Icon = frmMain.Icon
50080
50090  Set cSystem = New clsSystem
50100  Set SMF = cSystem.GetSystemFont(Me, Menu)
50110
50120  With Screen
50130   .MousePointer = vbHourglass
50140   Move (.Width - Width) / 2, (.Height - Height) / 2
50150  End With
50160
50170  fraBitmapGeneral.Visible = False
50180  fraBitmapGeneral.Top = fraPDFTop - 300
50190  fraBitmapGeneral.Left = fraPDFLeft - 200
50200  fraProgFont.Top = fraPDFTop - 300
50210  fraProgFont.Left = fraPDFLeft - 200
50220  fraProgGeneral.Top = fraPDFTop - 300
50230  fraProgGeneral.Left = fraPDFLeft - 200
50240  fraProgAutosave.Top = fraPDFTop - 300
50250  fraProgAutosave.Left = fraPDFLeft - 200
50260  fraProgSave.Top = fraPDFTop - 300
50270  fraProgSave.Left = fraPDFLeft - 200
50280  fraProgDirectories.Top = fraPDFTop - 300
50290  fraProgDirectories.Left = fraPDFLeft - 200
50300  fraProgDocument.Top = fraPDFTop - 300
50310  fraProgDocument.Left = fraPDFLeft - 200
50320  fraProgGhostscript.Top = fraPDFTop - 300
50330  fraProgGhostscript.Left = fraPDFLeft - 200
50340
50350  fraPDFSecurity.Top = fraPDFTop + 100
50360  fraPDFSecurity.Left = fraPDFLeft
50370  fraPDFFonts.Top = fraPDFTop + 100
50380  fraPDFFonts.Left = fraPDFLeft
50390  fraPDFColors.Top = fraPDFTop + 100
50400  fraPDFColors.Left = fraPDFLeft
50410  fraPDFGeneral.Top = fraPDFTop + 100
50420  fraPDFGeneral.Left = fraPDFLeft
50430  fraPDFCompress.Top = fraPDFTop + 100
50440  fraPDFCompress.Left = fraPDFLeft
50450  fraPSGeneral.Left = fraPDFLeft
50460  fraPSGeneral.Top = fraPDFTop - 300
50470  fraPSGeneral.Left = fraPDFLeft - 200
50480  tbstrPDFOptions.Top = fraPDFTop - 300
50490  tbstrPDFOptions.Left = fraPDFLeft - 200
50500  tbstrPDFOptions.Height = 5975
50510  tbstrPDFOptions.Width = 6215
50520
50530  cmbEPSLanguageLevel.Top = cmbPSLanguageLevel.Top
50540  cmbEPSLanguageLevel.Left = cmbPSLanguageLevel.Left
50550
50560  txtTest.Text = vbNullString
50570  For i = 33 To 255
50580   txtTest.Text = txtTest.Text & Chr$(i)
50590  Next i
50600  fi = -1
50610  With cmbFonts
50620   .Clear
50630   For i = 1 To Screen.FontCount
50640    tStr = Trim$(Screen.Fonts(i))
50650    If Len(tStr) > 0 Then
50660     cmbFonts.AddItem tStr
50670    End If
50680   Next i
50690   If .ListCount > 0 Then
50700     For i = 0 To cmbFonts.ListCount - 1
50710      If SMF.Count > 0 Then
50720       If UCase$(cmbFonts.List(i)) = UCase$(SMF(1)(0)) Then
50730        fi = i
50740       End If
50750      End If
50760     Next i
50770    Else
50780    .ListIndex = 0
50790   End If
50800  End With
50810  With cmbCharset
50820   .Clear
50830   .AddItem "0, Western": .ItemData(.NewIndex) = 0
50840   .AddItem "2, Symbol": .ItemData(.NewIndex) = 2
50850   .AddItem "77, Mac": .ItemData(.NewIndex) = 77
50860   .AddItem "161, Greek": .ItemData(.NewIndex) = 161
50870   .AddItem "162, Turkish": .ItemData(.NewIndex) = 162
50880   .AddItem "177, Hebrew": .ItemData(.NewIndex) = 177
50890   .AddItem "178, Arabic": .ItemData(.NewIndex) = 178
50900   .AddItem "186, Baltic": .ItemData(.NewIndex) = 186
50910   .AddItem "204, Cyrillic": .ItemData(.NewIndex) = 204
50920   .AddItem "238, Central European": .ItemData(.NewIndex) = 238
50930   .AddItem "255, DOS/OEM": .ItemData(.NewIndex) = 255
50940   .Text = 0
50950  End With
50960  If fi >= 0 Then
50970   cmbFonts.ListIndex = fi
50980   cmbCharset.Text = SMF(1)(2)
50990   cmbProgramFontsize.Text = SMF(1)(1)
51000   txtTest.Font = cmbFonts.List(cmbFonts.ListIndex)
51010   txtTest.Font.Charset = cmbCharset.Text
51020  End If
51030
51040
51050  trv.Nodes.Clear
51060  trv.Indentation = 200
51070  With LanguageStrings
51080   trv.Nodes.Add , , "Program", .OptionsTreeProgram
51090   trv.Nodes.Add "Program", tvwChild, "ProgramGeneral", .OptionsProgramGeneralSymbol
51100   trv.Nodes.Add "Program", tvwChild, "ProgramGhostscript", .OptionsProgramGhostscriptSymbol
51110   trv.Nodes.Add "Program", tvwChild, "ProgramDocument", .OptionsProgramDocumentSymbol
51120   trv.Nodes.Add "Program", tvwChild, "ProgramSave", .OptionsProgramSaveSymbol
51130   trv.Nodes.Add "Program", tvwChild, "ProgramAutosave", .OptionsProgramAutosaveSymbol
51140   trv.Nodes.Add "Program", tvwChild, "ProgramDirectories", .OptionsProgramDirectoriesSymbol
51150   trv.Nodes.Add "Program", tvwChild, "ProgramFonts", .OptionsProgramFontSymbol
51160   trv.Nodes.Add , , "Formats", .OptionsTreeFormats
51170   trv.Nodes.Add "Formats", tvwChild, "FormatsPDF", .OptionsPDFSymbol
51180   trv.Nodes.Add "Formats", tvwChild, "FormatsPNG", .OptionsPNGSymbol
51190   trv.Nodes.Add "Formats", tvwChild, "FormatsJPEG", .OptionsJPEGSymbol
51200   trv.Nodes.Add "Formats", tvwChild, "FormatsBMP", .OptionsBMPSymbol
51210   trv.Nodes.Add "Formats", tvwChild, "FormatsPCX", .OptionsPCXSymbol
51220   trv.Nodes.Add "Formats", tvwChild, "FormatsTIFF", .OptionsTIFFSymbol
51230   trv.Nodes.Add "Formats", tvwChild, "FormatsPS", .OptionsPSSymbol
51240   trv.Nodes.Add "Formats", tvwChild, "FormatsEPS", .OptionsEPSSymbol
51250
51260   trv.Nodes("ProgramFonts").EnsureVisible
51270   trv.Nodes("FormatsPDF").EnsureVisible
51280
51290   Set picOptions = LoadResPicture(2101, vbResIcon)
51300   fraProgFont.Visible = False
51310   fraProgGeneral.Visible = True
51320
51330   fraProgGeneral.Caption = .OptionsProgramGeneralSymbol
51340   fraProgGhostscript.Caption = .OptionsProgramGhostscriptSymbol
51350   fraProgAutosave.Caption = .OptionsProgramAutosaveSymbol
51360   fraProgFont.Caption = .OptionsProgramFontSymbol
51370   fraProgDirectories.Caption = .OptionsProgramDirectoriesSymbol
51380   fraProgSave.Caption = .OptionsProgramSaveSymbol
51390   fraProgDocument.Caption = .OptionsProgramDocumentSymbol
51400
51410   fraShellintegration.Caption = .OptionsShellIntegration
51420   cmdShellintegration(0).Caption = .OptionsShellIntegrationAdd
51430   cmdShellintegration(1).Caption = .OptionsShellIntegrationRemove
51440   If IsWin9xMe = False Then
51450    If IsAdmin = False Then
51460     cmdShellintegration(0).Enabled = False
51470     cmdShellintegration(1).Enabled = False
51480    End If
51490   End If
51500
51510   lblGhostscriptversion.Caption = .OptionsGhostscriptversion
51520
51530   lblSaveFilename.Caption = .OptionsSaveFilename
51540   lblSaveFilenameTokens.Caption = .OptionsSaveFilenameTokens
51550   fraFilenameSubstitutions.Caption = .OptionsSaveFilenameSubstitutions
51560   chkFilenameSubst.Caption = .OptionsSaveFilenameSubstitutionsTitle
51570   cmdFilenameSubst(0).Caption = .OptionsSaveFilenameAdd
51580   cmdFilenameSubst(1).Caption = .OptionsSaveFilenameChange
51590   cmdFilenameSubst(2).Caption = .OptionsSaveFilenameDelete
51600
51610   chkSpaces.Caption = .OptionsRemoveSpaces
51620   chkNoConfirmMessageSwitchingDefaultprinter.Caption = .OptionsProgramSwitchingDefaultprinter
51630   lblGSbin.Caption = .OptionsDirectoriesGSBin
51640   lblGSlib.Caption = .OptionsDirectoriesGSLibraries
51650   lblGSfonts.Caption = .OptionsDirectoriesGSFonts
51660   lblPrintTempPath.Caption = .OptionsDirectoriesTempPath
51670
51680   lblOptions = .OptionsProgramGeneralDescription
51690   lblAutosaveformat.Caption = .OptionsAutosaveFormat
51700   chkUseStandardAuthor.Caption = .OptionsUseStandardauthor
51710   chkUseCreationDateNow.Caption = .OptionsUseCreationDateNow
51720   chkUseAutosave.Caption = .OptionsUseAutosave
51730   cmdTestpage.Caption = .OptionsPrintTestpage
51740   lblAutosaveFilename.Caption = .OptionsAutosaveFilename
51750   lblAutosaveFilenameTokens.Caption = .OptionsAutosaveFilenameTokens
51760   chkUseAutosaveDirectory.Caption = .OptionsUseAutosaveDirectory
51770   lblAuthorTokens.Caption = .OptionsStandardAuthorToken
51780
51790   With cmbAutosaveFormat
51800    .AddItem "PDF"
51810    .AddItem "PNG"
51820    .AddItem "JPEG"
51830    .AddItem "BMP"
51840    .AddItem "PCX"
51850    .AddItem "TIFF"
51860    .AddItem "PS"
51870    .AddItem "EPS"
51880   End With
51890   With cmbSaveFilenameTokens
51900    .AddItem "<Author>"
51910    .AddItem "<Computername>"
51920    .AddItem "<DateTime>"
51930    .AddItem "<Title>"
51940    .AddItem "<Username>"
51950    .ListIndex = 0
51960   End With
51970   With cmbAuthorTokens
51980    .AddItem "<Computername>"
51990    .AddItem "<DateTime>"
52000    .AddItem "<Title>"
52010    .AddItem "<Username>"
52020    .ListIndex = 0
52030   End With
52040   With cmbAutoSaveFilenameTokens
52050    .AddItem "<Author>"
52060    .AddItem "<Computername>"
52070    .AddItem "<DateTime>"
52080    .AddItem "<Title>"
52090    .AddItem "<Username>"
52100    .ListIndex = 0
52110   End With
52120   Me.Caption = .DialogPrinterOptions
52130   cmdCancel.Caption = .OptionsCancel
52140   cmdReset.Caption = .OptionsReset
52150   cmdSave.Caption = .OptionsSave
52160   tbstrPDFOptions.Tabs.Clear
52170   tbstrPDFOptions.Tabs.Add , "General", .OptionsPDFGeneral
52180   tbstrPDFOptions.Tabs.Add , "Compression", .OptionsPDFCompression
52190   tbstrPDFOptions.Tabs.Add , "Fonts", .OptionsPDFFonts
52200   tbstrPDFOptions.Tabs.Add , "Colors", .OptionsPDFColors
52210   tbstrPDFOptions.Tabs.Add , "Security", .OptionsPDFSecurity
52220   fraPDFGeneral.Caption = .OptionsPDFGeneralCaption
52230   lblPDFCompat.Caption = .OptionsPDFGeneralCompatibility
52240   lblPDFAutoRotate.Caption = .OptionsPDFGeneralAutorotate
52250   lblPDFResolution.Caption = .OptionsPDFGeneralResolution
52260   lblPDFOverprint.Caption = .OptionsPDFGeneralOverprint
52270   lblProgfont.Caption = .OptionsProgramFont
52280   lblProgcharset.Caption = .OptionsProgramFontcharset
52290   lblSize.Caption = .OptionsProgramFontSize
52300   lblTesttext = .OptionsProgramFontTestdescription
52310   cmdTest.Caption = .OptionsProgramFontTest
52320   cmdCancelTest.Caption = .OptionsProgramFontCancelTest
52330   chkPDFASCII85.Caption = .OptionsPDFGeneralASCII85
52340   cmbPDFCompat.Clear
52350   cmbPDFCompat.AddItem .OptionsPDFGeneralCompatibility01
52360   cmbPDFCompat.AddItem .OptionsPDFGeneralCompatibility02
52370   cmbPDFCompat.AddItem .OptionsPDFGeneralCompatibility03
52380   cmbPDFRotate.Clear
52390   cmbPDFRotate.AddItem .OptionsPDFGeneralRotate01
52400   cmbPDFRotate.AddItem .OptionsPDFGeneralRotate02
52410   cmbPDFRotate.AddItem .OptionsPDFGeneralRotate03
52420   cmbPDFOverprint.Clear
52430   cmbPDFOverprint.AddItem .OptionsPDFGeneralOverprint01
52440   cmbPDFOverprint.AddItem .OptionsPDFGeneralOverprint02
52450
52460   fraPDFCompress.Caption = .OptionsPDFCompressionCaption
52470   chkPDFTextComp.Caption = .OptionsPDFCompressionTextComp
52480   fraPDFColor.Caption = .OptionsPDFCompressionColor
52490   chkPDFColorComp.Caption = .OptionsPDFCompressionColorComp
52500   chkPDFColorResample.Caption = .OptionsPDFCompressionColorResample
52510   lblPDFColorRes.Caption = .OptionsPDFCompressionColorRes
52520   cmbPDFColorComp.Clear
52530   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp01
52540   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp02
52550   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp03
52560   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp04
52570   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp05
52580   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp06
52590   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp07
52600   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp08
52610   cmbPDFColorResample.Clear
52620   cmbPDFColorResample.AddItem .OptionsPDFCompressionColorResample01
52630   cmbPDFColorResample.AddItem .OptionsPDFCompressionColorResample02
52640   cmbPDFColorResample.AddItem .OptionsPDFCompressionColorResample03
52650   fraPDFGrey.Caption = .OptionsPDFCompressionGrey
52660   chkPDFGreyComp.Caption = .OptionsPDFCompressionGreyComp
52670   chkPDFGreyResample.Caption = .OptionsPDFCompressionGreyResample
52680   lblPDFGreyRes.Caption = .OptionsPDFCompressionGreyRes
52690   cmbPDFGreyComp.Clear
52700   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp01
52710   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp02
52720   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp03
52730   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp04
52740   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp05
52750   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp06
52760   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp07
52770   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp08
52780   cmbPDFGreyResample.Clear
52790   cmbPDFGreyResample.AddItem .OptionsPDFCompressionGreyResample01
52800   cmbPDFGreyResample.AddItem .OptionsPDFCompressionGreyResample02
52810   cmbPDFGreyResample.AddItem .OptionsPDFCompressionGreyResample03
52820   fraPDFMono.Caption = .OptionsPDFCompressionMono
52830   chkPDFMonoComp.Caption = .OptionsPDFCompressionMonoComp
52840   chkPDFMonoResample.Caption = .OptionsPDFCompressionMonoResample
52850   lblPDFMonoRes.Caption = .OptionsPDFCompressionMonoRes
52860   cmbPDFMonoComp.Clear
52870   cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp01
52880   cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp02
52890   cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp03
52900   cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp04
52910   cmbPDFMonoResample.Clear
52920   cmbPDFMonoResample.AddItem .OptionsPDFCompressionMonoResample01
52930   cmbPDFMonoResample.AddItem .OptionsPDFCompressionMonoResample02
52940   cmbPDFMonoResample.AddItem .OptionsPDFCompressionMonoResample03
52950
52960   fraPDFFonts.Caption = .OptionsPDFFontsCaption
52970   chkPDFEmbedAll.Caption = .OptionsPDFFontsEmbedAll
52980   chkPDFSubSetFonts.Caption = .OptionsPDFFontsSubSetFonts
52990
53000   fraPDFColors.Caption = .OptionsPDFColorsCaption
53010   chkPDFCMYKtoRGB.Caption = .OptionsPDFColorsCMYKtoRGB
53020   fraPDFColorOptions.Caption = .OptionsPDFColorsColorOptions
53030   chkPDFPreserveOverprint.Caption = .OptionsPDFColorsPreserveOverprint
53040   chkPDFPreserveTransfer.Caption = .OptionsPDFColorsPreserveTransfer
53050   chkPDFPreserveHalftone.Caption = .OptionsPDFColorsPreserveHalftone
53060   cmbPDFColorModel.Clear
53070   cmbPDFColorModel.AddItem .OptionsPDFColorsColorModel01
53080   cmbPDFColorModel.AddItem .OptionsPDFColorsColorModel02
53090   cmbPDFColorModel.AddItem .OptionsPDFColorsColorModel03
53100
53110   fraPDFEncryptor.Caption = .OptionsPDFEncryptor
53120   fraPDFSecurity.Caption = .OptionsPDFSecurityCaption
53130   chkUseSecurity.Caption = .OptionsPDFUseSecurity
53140   fraPDFEncLevel.Caption = .OptionsPDFEncryptionLevel
53150   optEncHigh.Caption = .OptionsPDFEncryptionHigh
53160   optEncLow.Caption = .OptionsPDFEncryptionLow
53170   fraSecurityPass.Caption = .OptionsPDFPasswords
53180   chkUserPass.Caption = .OptionsPDFUserPass
53190   chkOwnerPass.Caption = .OptionsPDFOwnerPass
53200   fraPDFPermissions.Caption = .OptionsPDFDisallowUser
53210   fraPDFHighPermissions.Caption = .OptionsPDFEnhancedPermissions
53220   chkAllowPrinting.Caption = .OptionsPDFDisallowPrint
53230   chkAllowModifyContents.Caption = .OptionsPDFDisallowModify
53240   chkAllowCopy.Caption = .OptionsPDFDisallowCopy
53250   chkAllowModifyAnnotations.Caption = .OptionsPDFDisallowModifyComments
53260   chkAllowDegradedPrinting.Caption = .OptionsPDFAllowDegradedPrinting
53270   chkAllowFillIn.Caption = .OptionsPDFAllowFillIn
53280   chkAllowAssembly.Caption = .OptionsPDFAllowAssembly
53290   chkAllowScreenReaders.Caption = .OptionsPDFAllowScreenReaders
53300
53310   cmbPNGColors.AddItem .OptionsPNGColorscount01
53320   cmbPNGColors.AddItem .OptionsPNGColorscount02
53330   cmbPNGColors.AddItem .OptionsPNGColorscount03
53340   cmbPNGColors.AddItem .OptionsPNGColorscount04
53350   cmbJPEGColors.Left = cmbPNGColors.Left
53360   cmbJPEGColors.Width = cmbPNGColors.Width
53370   cmbJPEGColors.Top = cmbPNGColors.Top
53380   cmbJPEGColors.AddItem .OptionsJPEGColorscount01
53390   cmbJPEGColors.AddItem .OptionsJPEGColorscount02
53400   cmbBMPColors.Left = cmbPNGColors.Left
53410   cmbBMPColors.Width = cmbPNGColors.Width
53420   cmbBMPColors.Top = cmbPNGColors.Top
53430   cmbBMPColors.AddItem .OptionsBMPColorscount01
53440   cmbBMPColors.AddItem .OptionsBMPColorscount02
53450   cmbBMPColors.AddItem .OptionsBMPColorscount03
53460   cmbBMPColors.AddItem .OptionsBMPColorscount04
53470   cmbBMPColors.AddItem .OptionsBMPColorscount05
53480   cmbBMPColors.AddItem .OptionsBMPColorscount06
53490   cmbBMPColors.AddItem .OptionsBMPColorscount07
53500   cmbPCXColors.Left = cmbPNGColors.Left
53510   cmbPCXColors.Width = cmbPNGColors.Width
53520   cmbPCXColors.Top = cmbPNGColors.Top
53530   cmbPCXColors.AddItem .OptionsPCXColorscount01
53540   cmbPCXColors.AddItem .OptionsPCXColorscount02
53550   cmbPCXColors.AddItem .OptionsPCXColorscount03
53560   cmbPCXColors.AddItem .OptionsPCXColorscount04
53570   cmbPCXColors.AddItem .OptionsPCXColorscount05
53580   cmbPCXColors.AddItem .OptionsPCXColorscount06
53590   cmbTIFFColors.Left = cmbPNGColors.Left
53600   cmbTIFFColors.Width = cmbPNGColors.Width
53610   cmbTIFFColors.Top = cmbPNGColors.Top
53620   cmbTIFFColors.AddItem .OptionsTIFFColorscount01
53630   cmbTIFFColors.AddItem .OptionsTIFFColorscount02
53640   cmbTIFFColors.AddItem .OptionsTIFFColorscount03
53650   cmbTIFFColors.AddItem .OptionsTIFFColorscount04
53660   cmbTIFFColors.AddItem .OptionsTIFFColorscount05
53670   cmbTIFFColors.AddItem .OptionsTIFFColorscount06
53680   cmbTIFFColors.AddItem .OptionsTIFFColorscount07
53690   cmbTIFFColors.AddItem .OptionsTIFFColorscount08
53700
53710   fraBitmapGeneral.Caption = .OptionsImageSettings
53720   lblBitmapResolution = .OptionsBitmapResolution
53730   lblJPEGQuality = .OptionsJPEGQuality
53740   lblBitmapColors = .OptionsPDFColors
53750   lblProcessPriority.Caption = .OptionsProcesspriority
53760   lblLangLevel.Caption = .OptionsPSLanguageLevel
53770
53780   cmdAsso.Caption = .OptionsAssociatePSFiles
53790  End With
53800
53810  If IsPsAssociate = False Then
53820    cmdAsso.Enabled = True
53830   Else
53840    cmdAsso.Enabled = False
53850  End If
53860
53870  txtPDFRes.Text = 600
53880  cmbPDFCompat.ListIndex = 1
53890  cmbPDFRotate.ListIndex = 0
53900  cmbPDFOverprint.ListIndex = 0
53910  chkPDFASCII85.Value = 0
53920
53930  chkPDFTextComp.Value = 1
53940
53950  chkPDFColorComp.Value = 1
53960  chkPDFColorResample.Value = 0
53970  cmbPDFColorComp.ListIndex = 0
53980  cmbPDFColorResample.ListIndex = 0
53990  txtPDFColorRes.Text = 300
54000
54010  chkPDFGreyComp.Value = 1
54020  chkPDFGreyResample.Value = 0
54030  cmbPDFGreyComp.ListIndex = 0
54040  cmbPDFGreyResample.ListIndex = 0
54050  txtPDFGreyRes.Text = 300
54060
54070  chkPDFMonoComp.Value = 1
54080  chkPDFMonoResample.Value = 0
54090  cmbPDFMonoComp.ListIndex = 0
54100  cmbPDFMonoResample.ListIndex = 0
54110  txtPDFMonoRes.Text = 1200
54120
54130  chkPDFEmbedAll.Value = 1
54140  chkPDFSubSetFonts.Value = 1
54150  txtPDFSubSetPerc.Text = 100
54160
54170  cmbPDFColorModel.ListIndex = 1
54180  chkPDFCMYKtoRGB.Value = 1
54190  chkPDFPreserveOverprint.Value = 1
54200  chkPDFPreserveTransfer.Value = 1
54210  chkPDFPreserveHalftone.Value = 0
54220
54230  cmbPNGColors.ListIndex = 0
54240  cmbJPEGColors.ListIndex = 0
54250  cmbBMPColors.ListIndex = 0
54260  cmbPCXColors.ListIndex = 0
54270  cmbTIFFColors.ListIndex = 0
54280  txtBitmapResolution.Text = 150
54290
54300  cmbCharset.Text = cmbCharset.ItemData(0)
54310  cmbProgramFontsize.Text = 8
54320
54330 ' chkUseStandardAuthor.Value = 1
54340  txtStandardAuthor.Text = vbNullString
54350
54360  With cmbPSLanguageLevel
54370   .AddItem "1"
54380   .AddItem "1.5"
54390   .AddItem "2"
54400   .AddItem "3"
54410  End With
54420  With cmbEPSLanguageLevel
54430   .AddItem "1"
54440   .AddItem "1.5"
54450   .AddItem "2"
54460   .AddItem "3"
54470  End With
54480
54490  With lsvFilenameSubst
54500   .Appearance = ccFlat
54510   .ColumnHeaders.Clear
54520   .ColumnHeaders.Add , "Str1", "", lsvFilenameSubst.Width / 2 - 140
54530   .ColumnHeaders.Add , "Str2", "", lsvFilenameSubst.Width / 2 - 140
54540   .HideColumnHeaders = True
54550   .GridLines = True
54560   .FullRowSelect = True
54570   .HideSelection = False
54580  End With
54590
54600  With cmbPDFEncryptor
54610   .Clear
54620   .AddItem "Ghostscript (>= 8.14)"
54630   .ItemData(.NewIndex) = 0
54640   .AddItem "PDFEnc"
54650   .ItemData(.NewIndex) = 1
54660
54670   ShowOptions Me, Options
54680
54690   SecurityIsPossible = True
54700
54710   If FileExists(CompletePath(App.Path) & "pdfenc.exe") = False Then
54720    .RemoveItem 1
54730    .ListIndex = 0
54740    Options.PDFEncryptor = .ItemData(.ListIndex)
54750   End If
54760   If GhostScriptSecurity = False Then
54770    .RemoveItem 0
54780   End If
54790   If .ListCount = 0 Then
54800     chkUseSecurity.Value = 0
54810     chkUseSecurity.Enabled = False
54820     SecurityIsPossible = False
54830    Else
54840     For i = 0 To .ListCount - 1
54850      If .ItemData(i) = Options.PDFEncryptor Then
54860       .ListIndex = i
54870       Exit For
54880      End If
54890     Next i
54900     If .ListIndex = -1 Then
54910      .ListIndex = 0
54920      Options.PDFEncryptor = .ItemData(.ListIndex)
54930     End If
54940   End If
54950  End With
54960
54970
54980  If Options.PDFHighEncryption <> 0 Then
54990    optEncHigh.Value = True
55000   Else
55010    optEncLow.Value = True
55020  End If
55030
55040  CheckCmdFilenameSubst
55050
55060  If chkUseStandardAuthor.Value = 1 Then
55070    txtStandardAuthor.Enabled = True
55080    txtStandardAuthor.BackColor = &H80000005
55090   Else
55100    txtStandardAuthor.Enabled = False
55110    txtStandardAuthor.BackColor = &H8000000F
55120  End If
55130  With Options
55140   SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
55150   cmbCharset.Text = .ProgramFontCharset
55160  End With
55170  If chkUseAutosave.Value = 1 Then
55180    ViewAutosave True
55190   Else
55200    ViewAutosave False
55210  End If
55220
55230  With txtGSbin
55240   .ToolTipText = .Text
55250  End With
55260  With txtGSlib
55270   .ToolTipText = .Text
55280  End With
55290  With txtGSfonts
55300   .ToolTipText = .Text
55310  End With
55320  With txtTemppath
55330   .ToolTipText = .Text
55340  End With
55350
55360  With sldProcessPriority
55370   .TextPosition = sldBelowRight
55380   .TickFrequency = 1
55390   .TickStyle = sldTopLeft
55401   Select Case .Value
         Case 0: 'Idle
55420     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityIdle
55430    Case 1: 'Normal
55440     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityNormal
55450    Case 2: 'High
55460     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityHigh
55470    Case 3: 'Realtime
55480     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityRealtime
55490   End Select
55500  End With
55510
55520  If IsWin9xMe = False Then
55530    lblProcessPriority.Enabled = True
55540    sldProcessPriority.Enabled = True
55550   Else
55560    lblProcessPriority.Enabled = False
55570    sldProcessPriority.Enabled = False
55580  End If
55590  UpdateSecurityFields
55600
55610  tStr2 = CompletePath(UCase$(Trim$(Options.DirectoryGhostscriptBinaries)))
 cmbGhostscript.Clear: Set reg = New clsRegistry
55630  reg.hkey = HKEY_LOCAL_MACHINE
55640
55650  Set gsvers = GetAllGhostscriptversions
55660
55670  If gsvers.Count = 0 Then
55680    cmbGhostscript.Enabled = False
55690   Else
55700    For i = 1 To gsvers.Count
55710     cmbGhostscript.AddItem gsvers.item(i)
55720    Next i
55730    cmbGhostscript.ListIndex = cmbGhostscript.ListCount - 1
55740    For i = 0 To cmbGhostscript.ListCount - 1
55750     tStr = ""
55760     If InStr(cmbGhostscript.List(i), ":") Then
55770       reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
55780       If tStr2 = CompletePath(UCase$(Trim$(reg.GetRegistryValue("GhostscriptDirectoryBinaries")))) Then
55790        cmbGhostscript.ListIndex = i
55800        Exit For
55810       End If
55820      Else
55830       If InStr(UCase$(cmbGhostscript.List(i)), "AFPL") Then
55840        reg.KeyRoot = "SOFTWARE\AFPL Ghostscript"
55850        If InStr(cmbGhostscript.List(i), " ") > 0 Then
55860         tsf = Split(cmbGhostscript.List(i), " ")
55870         reg.Subkey = tsf(UBound(tsf))
55880         tStr = reg.GetRegistryValue("GS_DLL")
55890         If tStr2 & "GSDLL32.DLL" = UCase$(tStr) Then
55900          cmbGhostscript.ListIndex = i
55910          Exit For
55920         End If
55930        End If
55940       End If
55950       If InStr(UCase$(cmbGhostscript.List(i)), "GNU") Then
55960        reg.KeyRoot = "SOFTWARE\GNU Ghostscript"
55970        If InStr(cmbGhostscript.List(i), " ") > 0 Then
55980         tsf = Split(cmbGhostscript.List(i), " ")
55990         reg.Subkey = tsf(UBound(tsf))
56000         tStr = reg.GetRegistryValue("GS_DLL")
56010         If tStr2 & "GSDLL32.DLL" = UCase$(tStr) Then
56020          cmbGhostscript.ListIndex = i
56030          Exit For
56040         End If
56050        End If
56060       End If
56070       If InStr(UCase$(cmbGhostscript.List(i)), "GPL") Then
56080        reg.KeyRoot = "SOFTWARE\GPL Ghostscript"
56090        If InStr(cmbGhostscript.List(i), " ") > 0 Then
56100         tsf = Split(cmbGhostscript.List(i), " ")
56110         reg.Subkey = tsf(UBound(tsf))
56120         tStr = reg.GetRegistryValue("GS_DLL")
56130         If tStr2 & "GSDLL32.DLL" = UCase$(tStr) Then
56140          cmbGhostscript.ListIndex = i
56150          Exit For
56160         End If
56170        End If
56180       End If
56190     End If
56200    Next i
56210  End If
56220  Set reg = Nothing
56230  With cmbGhostscript
56240   If .ListCount = 0 Then
56250    .Enabled = False
56260    .BackColor = &H8000000F
56270   End If
56280  End With
56290
56300  With cmbProgramFontsize
56310   .AddItem "8"
56320   .AddItem "9"
56330   .AddItem "10"
56340   .AddItem "11"
56350   .AddItem "12"
56360   .AddItem "14"
56370   .AddItem "16"
56380   .AddItem "18"
56390   .AddItem "20"
56400   .AddItem "22"
56410   .AddItem "24"
56420   .AddItem "26"
56430   .AddItem "28"
56440   .AddItem "36"
56450   .AddItem "48"
56460   .AddItem "72"
56470  End With
56480
56490  For Each ctl In Controls
56500   If TypeOf ctl Is ComboBox Then
56510    ComboSetListWidth ctl
56520   End If
56530  Next ctl
56540
56550  SetOptimalComboboxHeigth cmbCharset, Me
56560  SetOptimalComboboxHeigth cmbProgramFontsize, Me
56570  SetOptimalComboboxHeigth cmbGhostscript, Me
56580  CorrectCmbCharset
56590  tbstrPDFOptions.ZOrder 1
56600  cmdStyle.ZOrder 1
56610  If ShowOnlyOptions = True Then
56620   FormInTaskbar Me, True, True
56630   Caption = "PDFCreator - " & Caption
56640  End If
56650  Screen.MousePointer = vbNormal
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

Private Sub txtAutosaveDirectory_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txtAutosaveDirectory.ToolTipText = txtAutosaveFilename.Text
50020  txtAutoSaveDirectoryPreview.Text = GetSubstFilename("B:\dummy.dum", txtAutosaveDirectory.Text, , True)
50030  If IsValidPath(txtAutoSaveDirectoryPreview.Text) = False Then
50040    txtAutoSaveDirectoryPreview.ForeColor = vbRed
50050   Else
50060    txtAutoSaveDirectoryPreview.ForeColor = &H80000008
50070  End If
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
50010  Dim ext As String
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
50010  If ViewIt = True Then
50020    txtAutosaveDirectory.Enabled = True
50030    txtAutosaveDirectory.BackColor = &H80000005
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
50130    GetAutosaveFormatExtension = ".tiff"
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
