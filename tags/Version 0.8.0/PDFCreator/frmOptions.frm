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
   Begin VB.Frame fraProgGhostscript 
      Caption         =   "Ghostscript"
      Height          =   975
      Left            =   2760
      TabIndex        =   146
      Top             =   1080
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton cmdGSDirs 
         Caption         =   ">>"
         Height          =   375
         Left            =   4680
         TabIndex        =   160
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdGetgsfontsDirectory 
         Caption         =   "..."
         Height          =   255
         Left            =   5760
         TabIndex        =   157
         Top             =   2400
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdGetgslibDirectory 
         Caption         =   "..."
         Height          =   255
         Left            =   5760
         TabIndex        =   156
         Top             =   1800
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtGSfonts 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   155
         Top             =   2400
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.TextBox txtGSlib 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   154
         Top             =   1800
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.CommandButton cmdGS 
         Caption         =   "7"
         Height          =   255
         Index           =   0
         Left            =   5280
         TabIndex        =   153
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdGS 
         Caption         =   "8"
         Height          =   255
         Index           =   1
         Left            =   5280
         TabIndex        =   152
         Top             =   480
         Visible         =   0   'False
         Width           =   735
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
         TabIndex        =   159
         Top             =   2160
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblGSlib 
         Caption         =   "Ghostscript Libraries"
         Height          =   255
         Left            =   240
         TabIndex        =   158
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
   Begin VB.Frame fraProgSave 
      Caption         =   "Save"
      Height          =   4335
      Left            =   2880
      TabIndex        =   107
      Top             =   1560
      Visible         =   0   'False
      Width           =   6255
      Begin VB.TextBox txtSavePreview 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   165
         Top             =   840
         Width           =   6015
      End
      Begin VB.ComboBox cmbSaveFilenameTokens 
         Height          =   315
         ItemData        =   "frmOptions.frx":058A
         Left            =   3720
         List            =   "frmOptions.frx":058C
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
            Height          =   375
            Index           =   3
            Left            =   120
            Picture         =   "frmOptions.frx":058E
            Style           =   1  'Grafisch
            TabIndex        =   114
            Top             =   840
            Width           =   375
         End
         Begin VB.CommandButton cmdFilenameSubst 
            Enabled         =   0   'False
            Height          =   375
            Index           =   4
            Left            =   120
            Picture         =   "frmOptions.frx":06D8
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
         Caption         =   "Add a Filename-Token"
         Height          =   255
         Left            =   3720
         TabIndex        =   123
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblSaveFilename 
         Caption         =   "Filename"
         Height          =   255
         Left            =   120
         TabIndex        =   122
         Top             =   240
         Width           =   3495
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
         TabIndex        =   164
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
         Height          =   255
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
         ItemData        =   "frmOptions.frx":0822
         Left            =   3720
         List            =   "frmOptions.frx":0824
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
         Caption         =   "Filename"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   3495
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
         Caption         =   "Add a Filename-Token"
         Height          =   255
         Left            =   3720
         TabIndex        =   15
         Top             =   1440
         Width           =   2415
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
         TabIndex        =   161
         Top             =   2880
         Width           =   6015
         Begin VB.CommandButton cmdShellintegration 
            Caption         =   "Remove shell integration"
            Height          =   615
            Index           =   1
            Left            =   3240
            TabIndex        =   163
            Top             =   360
            Width           =   2655
         End
         Begin VB.CommandButton cmdShellintegration 
            Caption         =   "Integrate PDFCreator into shell"
            Height          =   615
            Index           =   0
            Left            =   120
            TabIndex        =   162
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
            ItemData        =   "frmOptions.frx":0826
            Left            =   120
            List            =   "frmOptions.frx":0828
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
            ItemData        =   "frmOptions.frx":082A
            Left            =   2280
            List            =   "frmOptions.frx":082C
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
            ItemData        =   "frmOptions.frx":082E
            Left            =   2280
            List            =   "frmOptions.frx":0830
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
            ItemData        =   "frmOptions.frx":0832
            Left            =   120
            List            =   "frmOptions.frx":0834
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
            ItemData        =   "frmOptions.frx":0836
            Left            =   120
            List            =   "frmOptions.frx":0838
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
            ItemData        =   "frmOptions.frx":083A
            Left            =   2280
            List            =   "frmOptions.frx":083C
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
   Begin VB.Frame fraProgDirectories 
      Caption         =   "Directories"
      Height          =   1095
      Left            =   2760
      TabIndex        =   18
      Top             =   3840
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton cmdGetTemppath 
         Caption         =   "..."
         Height          =   255
         Left            =   5760
         TabIndex        =   20
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtTemppath 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   600
         Width           =   5415
      End
      Begin VB.Label lblPrintTempPath 
         Caption         =   "Temppath"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   2895
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
            ItemData        =   "frmOptions.frx":083E
            Left            =   240
            List            =   "frmOptions.frx":0840
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
         ItemData        =   "frmOptions.frx":0842
         Left            =   2400
         List            =   "frmOptions.frx":0844
         Style           =   2  'Dropdown-Liste
         TabIndex        =   101
         Tag             =   "None|All|PageByPage"
         Top             =   840
         Width           =   2655
      End
      Begin VB.ComboBox cmbPDFCompat 
         Height          =   315
         ItemData        =   "frmOptions.frx":0846
         Left            =   2400
         List            =   "frmOptions.frx":0848
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
         ItemData        =   "frmOptions.frx":084A
         Left            =   2400
         List            =   "frmOptions.frx":084C
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
         ItemData        =   "frmOptions.frx":084E
         Left            =   120
         List            =   "frmOptions.frx":0850
         Style           =   2  'Dropdown-Liste
         TabIndex        =   90
         Tag             =   "RGB|CMYK|GRAY"
         Top             =   360
         Width           =   2775
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
         Height          =   285
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
         Width           =   5775
      End
      Begin VB.TextBox txtStandardAuthor 
         Height          =   285
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
         Width           =   2895
      End
      Begin VB.ComboBox cmbAuthorTokens 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmOptions.frx":0852
         Left            =   3720
         List            =   "frmOptions.frx":0854
         Style           =   2  'Dropdown-Liste
         TabIndex        =   23
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label lblAuthorTokens 
         Caption         =   "Add a Author-Token"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3720
         TabIndex        =   27
         Top             =   600
         Width           =   2415
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
 If chkPDFColorComp.Value = 1 Then
   cmbPDFColorComp.Enabled = True
   If cmbPDFColorComp.ListIndex = 0 Then
     chkPDFColorResample.Enabled = False
     cmbPDFColorResample.Enabled = False
     lblPDFColorRes.Enabled = False
     txtPDFColorRes.Enabled = False
    Else
     chkPDFColorResample.Enabled = True
     If chkPDFColorResample.Value = 1 Then
       cmbPDFColorResample.Enabled = True
       lblPDFColorRes.Enabled = True
       txtPDFColorRes.Enabled = True
      Else
       cmbPDFColorResample.Enabled = False
       lblPDFColorRes.Enabled = False
       txtPDFColorRes.Enabled = False
     End If
   End If
  Else
   cmbPDFColorComp.Enabled = False
   chkPDFColorResample.Enabled = False
   cmbPDFColorResample.Enabled = False
   lblPDFColorRes.Enabled = False
   txtPDFColorRes.Enabled = False
 End If
End Sub

Private Sub SetPDFGreyComprSettings()
 If chkPDFGreyComp.Value = 1 Then
   cmbPDFGreyComp.Enabled = True
   If cmbPDFGreyComp.ListIndex = 0 Then
     chkPDFGreyResample.Enabled = False
     cmbPDFGreyResample.Enabled = False
     lblPDFGreyRes.Enabled = False
     txtPDFGreyRes.Enabled = False
    Else
     chkPDFGreyResample.Enabled = True
     If chkPDFGreyResample.Value = 1 Then
       cmbPDFGreyResample.Enabled = True
       lblPDFGreyRes.Enabled = True
       txtPDFGreyRes.Enabled = True
      Else
       cmbPDFGreyResample.Enabled = False
       lblPDFGreyRes.Enabled = False
       txtPDFGreyRes.Enabled = False
     End If
   End If
  Else
   cmbPDFGreyComp.Enabled = False
   chkPDFGreyResample.Enabled = False
   cmbPDFGreyResample.Enabled = False
   lblPDFGreyRes.Enabled = False
   txtPDFGreyRes.Enabled = False
 End If
End Sub

Private Sub SetPDFMonoComprSettings()
 If chkPDFMonoComp.Value = 1 Then
   cmbPDFMonoComp.Enabled = True
   chkPDFMonoResample.Enabled = True
   If chkPDFMonoResample.Value = 1 Then
     cmbPDFMonoResample.Enabled = True
     lblPDFMonoRes.Enabled = True
     txtPDFMonoRes.Enabled = True
    Else
     cmbPDFMonoResample.Enabled = False
     lblPDFMonoRes.Enabled = False
     txtPDFMonoRes.Enabled = False
   End If
  Else
   cmbPDFMonoComp.Enabled = False
   chkPDFMonoResample.Enabled = False
   cmbPDFMonoResample.Enabled = False
   lblPDFMonoRes.Enabled = False
   txtPDFMonoRes.Enabled = False
 End If
End Sub

Private Sub chkOwnerPass_Click()
 If chkUserPass.Value = 0 Then
  If chkOwnerPass.Value = 0 Then
   chkOwnerPass.Value = 1
  End If
 End If
End Sub

Private Sub chkPDFColorComp_Click()
 SetPDFColorComprSettings
End Sub

Private Sub chkPDFColorResample_Click()
 SetPDFColorComprSettings
End Sub

Private Sub chkPDFGreyComp_Click()
 SetPDFGreyComprSettings
End Sub

Private Sub chkPDFGreyResample_Click()
 SetPDFGreyComprSettings
End Sub

Private Sub chkPDFMonoComp_Click()
 SetPDFMonoComprSettings
End Sub

Private Sub chkPDFMonoResample_Click()
 SetPDFMonoComprSettings
End Sub

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

Private Sub chkUserPass_Click()
 If chkOwnerPass.Value = 0 Then
  If chkUserPass.Value = 0 Then
   chkUserPass.Value = 1
   chkOwnerPass.Value = 1
  End If
  SavePasswordsForThisSession = False
 End If
End Sub

Private Sub chkUseSecurity_Click()
 UpdateSecurityFields
End Sub

Private Sub chkUseStandardAuthor_Click()
 If chkUseStandardAuthor.Value = 1 Then
   txtStandardAuthor.Enabled = True
   txtStandardAuthor.BackColor = &H80000005
   cmbAuthorTokens.Enabled = True
   lblAuthorTokens.Enabled = True
  Else
   txtStandardAuthor.Enabled = False
   txtStandardAuthor.BackColor = &H8000000F
   cmbAuthorTokens.Enabled = False
   lblAuthorTokens.Enabled = False
 End If
End Sub

Private Sub cmbAuthorTokens_Click()
 txtStandardAuthor.Text = txtStandardAuthor.Text & cmbAuthorTokens.Text
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

Private Sub cmbCharset_Validate(Cancel As Boolean)
 Dim i As Long, tStr As String
 tStr = ""
 For i = 1 To Len(cmbCharset.Text)
  If InStr("0123456789", Mid(cmbCharset.Text, i, 1)) > 0 Then
    tStr = tStr & Mid(cmbCharset.Text, i, 1)
   Else
    Exit For
  End If
 Next i
 If Len(Trim$(tStr)) = 0 Then
   cmbCharset.Text = 0
  Else
   cmbCharset.Text = tStr
 End If
End Sub

Private Sub cmbAutoSaveFilenameTokens_Click()
 txtAutosaveFilename.Text = txtAutosaveFilename.Text & cmbAutoSaveFilenameTokens.Text
End Sub

Private Sub cmbGhostscript_Click()
 Dim reg As clsRegistry, gsv As String, tsf() As String, Path As String, tStr As String

 gsv = cmbGhostscript.List(cmbGhostscript.ListIndex)
 Set reg = New clsRegistry: reg.hkey = HKEY_LOCAL_MACHINE

 If InStr(gsv, ":") Then
   reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
   txtGSbin.Text = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryBinaries"))
   txtGSfonts.Text = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
   txtGSlib.Text = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryLibraries"))
   Set reg = Nothing
   Exit Sub
  Else
   If InStr(UCase$(gsv), "AFPL") Then
    If InStr(gsv, " ") > 0 Then
     tsf = Split(gsv, " ")
     reg.KeyRoot = "SOFTWARE\AFPL Ghostscript\" & tsf(UBound(tsf))
     tStr = reg.GetRegistryValue("GS_DLL")
     SplitPath tStr, , Path
     txtGSbin.Text = CompletePath(Path)
     txtGSlib.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
     txtGSfonts.Text = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
    End If
   End If
   If InStr(UCase$(gsv), "GNU") Then
    If InStr(gsv, " ") > 0 Then
     tsf = Split(gsv, " ")
     reg.KeyRoot = "SOFTWARE\GNU Ghostscript\" & tsf(UBound(tsf))
     tStr = reg.GetRegistryValue("GS_DLL")
     SplitPath tStr, , Path
     txtGSbin.Text = CompletePath(Path)
     txtGSlib.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
     txtGSfonts.Text = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
    End If
   End If
 End If

 Set reg = Nothing
End Sub

Private Sub cmbPDFColorComp_Click()
 SetPDFColorComprSettings
End Sub

Private Sub cmbPDFGreyComp_Click()
 SetPDFGreyComprSettings
End Sub

Private Sub cmbPDFMonoComp_Click()
 SetPDFMonoComprSettings
End Sub

Private Sub cmbSaveFilenameTokens_Click()
 txtSaveFilename.Text = txtSaveFilename.Text & cmbSaveFilenameTokens.Text
End Sub

Private Sub cmbFonts_Click()
 txtTest.Font.Name = cmbFonts.List(cmbFonts.ListIndex)
End Sub

Private Sub cmbPDFCompat_Click()
 UpdateSecurityFields
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

Private Sub cmdFilenameSubst_Click(Index As Integer)
 Select Case Index
        Case 0: ' Add
   AddFilenameSubstitutions
  Case 1: ' Change
   ChangeFilenameSubstitutions
  Case 2: ' Delete
   DeleteFilenameSubstitutions
  Case 3: ' Up
   MoveUpFilenameSubstitutions
  Case 4: ' Down
   MoveDownFilenameSubstitutions
 End Select
End Sub

Private Sub cmdGetAutosaveDirectory_Click()
Dim strFolder As String

strFolder = BrowseForFolder(Me.hwnd, LanguageStrings.OptionsAutosaveDirectoryPrompt)
If Len(strFolder) = 0 Then Exit Sub
If Right$(strFolder, 1) <> "\" Then
 strFolder = strFolder & "\"
End If
txtAutosaveDirectory.Text = strFolder
End Sub

Private Sub cmdGetgsbinDirectory_Click()
 Dim strFolder As String, aw As Long

 strFolder = BrowseForFolder(Me.hwnd, LanguageStrings.OptionsGhostscriptBinariesDirectoryPrompt)
 If Len(strFolder) = 0 Then Exit Sub
 strFolder = CompletePath(strFolder)

 If LenB(Dir(strFolder & GsDll, vbNormal)) = 0 Then
  MsgBox LanguageStrings.MessagesMsg15
  Exit Sub
 End If
' UnloadDLLComplete GsDllLoaded
' GsDllLoaded = LoadDLL(strFolder & GsDll)
' If GsDllLoaded = 0 Then
'   MsgBox LanguageStrings.MessagesMsg15
'   Exit Sub
'  Else
'   UnLoadDLL GsDllLoaded
' End If

 If UCase$(CompletePath(Options.DirectoryGhostscriptBinaries)) <> UCase$(CompletePath(strFolder)) Then
  aw = MsgBox("The program must be restarted!", vbOKCancel)
  If aw = vbCancel Then
   Exit Sub
  End If
  txtGSbin.Text = strFolder
  GetOptions Me, Options
  SaveOptions Options
  Restart = True
  Unload Me
 End If
' Options.DirectoryGhostscriptBinaries = strFolder
 txtGSbin.Text = strFolder
 With txtGSbin
  .ToolTipText = .Text
 End With
End Sub

Private Sub cmdGetgsfontsDirectory_Click()
 Dim strFolder As String

 strFolder = BrowseForFolder(Me.hwnd, LanguageStrings.OptionsGhostscriptFontsDirectoryPrompt)
 If Len(strFolder) = 0 Then Exit Sub
 strFolder = CompletePath(strFolder)

 If Len(Dir(strFolder & "*.afm", vbNormal)) = 0 And Len(Dir(strFolder & "*.pfb", vbNormal)) = 0 Then
  MsgBox LanguageStrings.MessagesMsg16
  Exit Sub
 End If

 txtGSfonts.Text = strFolder
 With txtGSfonts
  .ToolTipText = .Text
 End With
End Sub

Private Sub cmdGetgslibDirectory_Click()
 Dim strFolder As String

 strFolder = BrowseForFolder(Me.hwnd, LanguageStrings.OptionsGhostscriptLibrariesDirectoryPrompt)
 If Len(strFolder) = 0 Then Exit Sub
 strFolder = CompletePath(strFolder)

 If Len(Dir(strFolder & "*.*", vbNormal)) = 0 Then
  MsgBox LanguageStrings.MessagesMsg17
  Exit Sub
 End If

 txtGSlib.Text = strFolder
 With txtGSlib
  .ToolTipText = .Text
 End With
End Sub

Private Sub cmdGetTemppath_Click()
 Dim strFolder As String

 strFolder = BrowseForFolder(Me.hwnd, LanguageStrings.OptionsPrintertempDirectoryPrompt)
 If Len(strFolder) = 0 Then Exit Sub
 strFolder = CompletePath(strFolder)
 txtTemppath.Text = strFolder
 With txtTemppath
  .ToolTipText = .Text
 End With
End Sub

Private Sub cmdGS_Click(Index As Integer)
 Select Case Index
        Case 0:
   txtGSbin.Text = "C:\gs7.06\gs7.06\bin\"
   txtGSlib.Text = "C:\gs7.06\gs7.06\lib\"
   txtGSfonts.Text = "C:\gs7.06\fonts\"
  Case 1:
   txtGSbin.Text = "C:\gs8.14\gs8.14\bin\"
   txtGSlib.Text = "C:\gs8.14\gs8.14\lib\"
   txtGSfonts.Text = "C:\gs8.14\fonts\"
 End Select
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
 Dim tRestart As Boolean
 tRestart = False
 If UCase$(Options.DirectoryGhostscriptBinaries) <> UCase$(txtGSbin.Text) Then
  tRestart = True
 End If
 GetOptions Me, Options
 SaveOptions Options
 If IsWin9xMe = False Then
  Select Case Options.ProcessPriority
         Case 0: 'Idle
    SetProcessPriority Idle
   Case 1: 'Normal
    SetProcessPriority Normal
   Case 2: 'High
    SetProcessPriority High
   Case 3: 'Realtime
    SetProcessPriority RealTime
  End Select
 End If
 If tRestart = True Then
  Restart = True
 End If
 Unload Me
End Sub

Private Sub cmdShellintegration_Click(Index As Integer)
 MousePointer = vbHourglass
 cmdShellintegration(0).Enabled = False
 cmdShellintegration(1).Enabled = False
 Select Case Index
        Case 0
   AddExplorerIntegration
  Case 1
   RemoveExplorerIntegration
 End Select
 MousePointer = vbNormal
 cmdShellintegration(0).Enabled = True
 cmdShellintegration(1).Enabled = True
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
 Dim TestPSPage As String, fn As Long, Filename As String
 TestPSPage = LoadResString(3000)
 TestPSPage = Replace(TestPSPage, "[TESTPAGE]", LanguageStrings.OptionsTestpage, , 1, vbTextCompare)
 TestPSPage = Replace(TestPSPage, "[DATE]", Now, , 1, vbTextCompare)
 TestPSPage = Replace(TestPSPage, "[PDFCREATORVERSION]", App.Title & " " & GetProgramReleaseStr, , 1, vbTextCompare)

 fn = FreeFile
 Filename = GetTempFile(GetPDFCreatorTempfolder, "~PD")
 Open Filename For Output As fn
 Print #fn, TestPSPage
 Close #fn
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF1 Then
  KeyCode = 0
  Select Case trv.SelectedItem.Key
         Case "Program"
    Call HTMLHelp_ShowTopic("html\generalsettings.htm")
   Case "ProgramGeneral"
    Call HTMLHelp_ShowTopic("html\generalsettings.htm")
   Case "ProgramSave"
    Call HTMLHelp_ShowTopic("html\savesettings.htm")
   Case "ProgramAutosave"
    Call HTMLHelp_ShowTopic("html\autosave.htm")
   Case "ProgramFonts"
    Call HTMLHelp_ShowTopic("html\fontsetting.htm")
   Case "ProgramDirectories"
    Call HTMLHelp_ShowTopic("html\directories.htm")
   Case "ProgramDocument"
    Call HTMLHelp_ShowTopic("html\docproperties.htm")
   Case "Formats"
    Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
   Case "FormatsPDF"
    Select Case tbstrPDFOptions.SelectedItem.Index
           Case 1:
      Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
     Case 2:
      Call HTMLHelp_ShowTopic("html\pdfcompression.htm")
     Case 3:
      Call HTMLHelp_ShowTopic("html\pdffonts.htm")
     Case 4:
      Call HTMLHelp_ShowTopic("html\pdfcolors.htm")
     Case 5:
      Call HTMLHelp_ShowTopic("html\pdfsecurity.htm")
    End Select
   Case "FormatsPNG"
    Call HTMLHelp_ShowTopic("html\pngsettings.htm")
   Case "FormatsJPEG"
    Call HTMLHelp_ShowTopic("html\jpegsettings.htm")
   Case "FormatsBMP"
    Call HTMLHelp_ShowTopic("html\bmpsettings.htm")
   Case "FormatsPCX"
    Call HTMLHelp_ShowTopic("html\pcxsettings.htm")
   Case "FormatsTIFF"
    Call HTMLHelp_ShowTopic("html\tiffsettings.htm")
   Case "FormatsPS"
    Call HTMLHelp_ShowTopic("html\pssettings.htm")
   Case "FormatsEPS"
    Call HTMLHelp_ShowTopic("html\epssettings.htm")
  End Select
 End If
End Sub

Private Sub Form_Load()
 Const fraPDFTop = 1360, fraPDFLeft = 2960
 Dim pic As New StdPicture, i As Long, tStr As String, gsvers As Collection, _
  cSystem As clsSystem, fi As Long, fc As Long, SMF As Collection, _
  reg As clsRegistry, tsf() As String, tStr2 As String

 Me.KeyPreview = True
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
 fraProgGhostscript.Top = fraPDFTop - 300
 fraProgGhostscript.Left = fraPDFLeft - 200

 fraPDFSecurity.Top = fraPDFTop + 100
 fraPDFSecurity.Left = fraPDFLeft
 fraPDFFonts.Top = fraPDFTop + 100
 fraPDFFonts.Left = fraPDFLeft
 fraPDFColors.Top = fraPDFTop + 100
 fraPDFColors.Left = fraPDFLeft
 fraPDFGeneral.Top = fraPDFTop + 100
 fraPDFGeneral.Left = fraPDFLeft
 fraPDFCompress.Top = fraPDFTop + 100
 fraPDFCompress.Left = fraPDFLeft
 fraPSGeneral.Left = fraPDFLeft
 fraPSGeneral.Top = fraPDFTop - 300
 fraPSGeneral.Left = fraPDFLeft - 200
 tbstrPDFOptions.Top = fraPDFTop - 300
 tbstrPDFOptions.Left = fraPDFLeft - 200
 tbstrPDFOptions.Height = 5975
 tbstrPDFOptions.Width = 6215

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
  txtTest.Font = cmbFonts.List(cmbFonts.ListIndex)
  txtTest.Font.Charset = cmbCharset.Text
 End If


 trv.Nodes.Clear
 trv.Indentation = 200
 With LanguageStrings
  trv.Nodes.Add , , "Program", .OptionsTreeProgram
  trv.Nodes.Add "Program", tvwChild, "ProgramGeneral", .OptionsProgramGeneralSymbol
  trv.Nodes.Add "Program", tvwChild, "ProgramGhostscript", .OptionsProgramGhostscriptSymbol
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
  fraProgGhostscript.Caption = .OptionsProgramGhostscriptSymbol
  fraProgAutosave.Caption = .OptionsProgramAutosaveSymbol
  fraProgFont.Caption = .OptionsProgramFontSymbol
  fraProgDirectories.Caption = .OptionsProgramDirectoriesSymbol
  fraProgSave.Caption = .OptionsProgramSaveSymbol
  fraProgDocument.Caption = .OptionsProgramDocumentSymbol

  fraShellintegration.Caption = .OptionsShellIntegration
  cmdShellintegration(0).Caption = .OptionsShellIntegrationAdd
  cmdShellintegration(1).Caption = .OptionsShellIntegrationRemove
  If WinVersion >= viWinNT351 Then
   If IsAdmin = False Then
    cmdShellintegration(0).Enabled = False
    cmdShellintegration(1).Enabled = False
   End If
  End If
  
  lblGhostscriptversion.Caption = .OptionsGhostscriptversion

  lblSaveFilename.Caption = .OptionsSaveFilename
  lblSaveFilenameTokens.Caption = .OptionsSaveFilenameTokens
  fraFilenameSubstitutions.Caption = .OptionsSaveFilenameSubstitutions
  chkFilenameSubst.Caption = .OptionsSaveFilenameSubstitutionsTitle
  cmdFilenameSubst(0).Caption = .OptionsSaveFilenameAdd
  cmdFilenameSubst(1).Caption = .OptionsSaveFilenameChange
  cmdFilenameSubst(2).Caption = .OptionsSaveFilenameDelete

  chkSpaces.Caption = .OptionsRemoveSpaces
  chkNoConfirmMessageSwitchingDefaultprinter.Caption = .OptionsProgramSwitchingDefaultprinter
  lblGSbin.Caption = .OptionsDirectoriesGSBin
  lblGSlib.Caption = .OptionsDirectoriesGSLibraries
  lblGSfonts.Caption = .OptionsDirectoriesGSFonts
  lblPrintTempPath.Caption = .OptionsDirectoriesTempPath

  lblOptions = .OptionsProgramGeneralDescription
  lblAutosaveformat.Caption = .OptionsAutosaveFormat
  chkUseStandardAuthor.Caption = .OptionsUseStandardauthor
  chkUseCreationDateNow.Caption = .OptionsUseCreationDateNow
  chkUseAutosave.Caption = .OptionsUseAutosave
  cmdTestpage.Caption = .OptionsPrintTestpage
  lblAutosaveFilename.Caption = .OptionsAutosaveFilename
  lblAutosaveFilenameTokens.Caption = .OptionsAutosaveFilenameTokens
  chkUseAutosaveDirectory.Caption = .OptionsUseAutosaveDirectory
  lblAuthorTokens.Caption = .OptionsStandardAuthorToken

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
  With cmbAuthorTokens
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
  tbstrPDFOptions.Tabs.Add , "Security", .OptionsPDFSecurity
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
  cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp08
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
  cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp08
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

  fraPDFEncryptor.Caption = .OptionsPDFEncryptor
  fraPDFSecurity.Caption = .OptionsPDFSecurityCaption
  chkUseSecurity.Caption = .OptionsPDFUseSecurity
  fraPDFEncLevel.Caption = .OptionsPDFEncryptionLevel
  optEncHigh.Caption = .OptionsPDFEncryptionHigh
  optEncLow.Caption = .OptionsPDFEncryptionLow
  fraSecurityPass.Caption = .OptionsPDFPasswords
  chkUserPass.Caption = .OptionsPDFUserPass
  chkOwnerPass.Caption = .OptionsPDFOwnerPass
  fraPDFPermissions.Caption = .OptionsPDFDisallowUser
  fraPDFHighPermissions.Caption = .OptionsPDFEnhancedPermissions
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

 With cmbPDFEncryptor
  .Clear
  .AddItem "Ghostscript (>= 8.14)"
  .ItemData(.NewIndex) = 0
  .AddItem "PDFEnc"
  .ItemData(.NewIndex) = 1

  ShowOptions Me, Options

  SecurityIsPossible = True

  If LenB(Dir(CompletePath(App.Path) & "pdfenc.exe")) = 0 Then
   .RemoveItem 1
   .ListIndex = 0
   Options.PDFEncryptor = .ItemData(.ListIndex)
  End If
  If GhostScriptSecurity = False Then
   .RemoveItem 0
  End If
  If .ListCount = 0 Then
    chkUseSecurity.Value = 0
    chkUseSecurity.Enabled = False
    SecurityIsPossible = False
   Else
    For i = 0 To .ListCount - 1
     If .ItemData(i) = Options.PDFEncryptor Then
      .ListIndex = i
      Exit For
     End If
    Next i
    If .ListIndex = -1 Then
     .ListIndex = 0
     Options.PDFEncryptor = .ItemData(.ListIndex)
    End If
  End If
 End With


 If Options.PDFHighEncryption <> 0 Then
   optEncHigh.Value = True
  Else
   optEncLow.Value = True
 End If

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

 tStr2 = CompletePath(UCase$(Trim$(Options.DirectoryGhostscriptBinaries)))
 cmbGhostscript.Clear: Set reg = New clsRegistry
 reg.hkey = HKEY_LOCAL_MACHINE

 Set gsvers = GetAllGhostscriptversions

 If gsvers.Count = 0 Then
   cmbGhostscript.Enabled = False
  Else
   For i = 1 To gsvers.Count
    cmbGhostscript.AddItem gsvers.item(i)
   Next i
   For i = 0 To cmbGhostscript.ListCount - 1
    tStr = ""
    If InStr(cmbGhostscript.List(i), ":") Then
      reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
      If tStr2 = CompletePath(UCase$(Trim$(reg.GetRegistryValue("GhostscriptDirectoryBinaries")))) Then
       cmbGhostscript.ListIndex = i
       Exit For
      End If
     Else
      If InStr(UCase$(cmbGhostscript.List(i)), "AFPL") Then
       reg.KeyRoot = "SOFTWARE\AFPL Ghostscript"
       If InStr(cmbGhostscript.List(i), " ") > 0 Then
        tsf = Split(cmbGhostscript.List(i), " ")
        reg.Subkey = tsf(UBound(tsf))
        tStr = reg.GetRegistryValue("GS_DLL")
        If tStr2 & "GSDLL32.DLL" = UCase$(tStr) Then
         cmbGhostscript.ListIndex = i
         Exit For
        End If
       End If
      End If
      If InStr(UCase$(cmbGhostscript.List(i)), "GNU") Then
       reg.KeyRoot = "SOFTWARE\GNU Ghostscript"
       If InStr(cmbGhostscript.List(i), " ") > 0 Then
        tsf = Split(cmbGhostscript.List(i), " ")
        reg.Subkey = tsf(UBound(tsf))
        tStr = reg.GetRegistryValue("GS_DLL")
        If tStr2 & "GSDLL32.DLL" = UCase$(tStr) Then
         cmbGhostscript.ListIndex = i
         Exit For
        End If
       End If
      End If
    End If
   Next i
 End If
 Set reg = Nothing
 With cmbGhostscript
  If .ListCount = 0 Then
   .Enabled = False
   .BackColor = &H8000000F
  End If
 End With
 Screen.MousePointer = vbNormal
End Sub

Private Sub ShowProgOptions()
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
 fraProgSave.Visible = False
 fraProgDirectories.Visible = False
 fraProgDocument.Visible = False
 fraBitmapGeneral.Visible = False
 fraPDFGeneral.Visible = False
 fraPDFCompress.Visible = False
 fraPDFFonts.Visible = False
 fraPDFColors.Visible = False
 fraPDFSecurity.Visible = False
 txtJPEGQuality.Visible = False
 lblJPEQQualityProzent.Visible = False
 fraPSGeneral.Visible = False
 cmbPSLanguageLevel.Visible = False
 cmbEPSLanguageLevel.Visible = False
 fraProgGhostscript.Visible = False

 Select Case trv.SelectedItem.Key
        Case "Program"
   Set picOptions = LoadResPicture(2101, vbResIcon)
   lblOptions = LanguageStrings.OptionsProgramGeneralDescription
   fraProgGeneral.Visible = True
  Case "ProgramGeneral"
   Set picOptions = LoadResPicture(2101, vbResIcon)
   lblOptions = LanguageStrings.OptionsProgramGeneralDescription
   fraProgGeneral.Visible = True
  Case "ProgramGhostscript"
   Set picOptions = LoadResPicture(2119, vbResIcon)
   lblOptions = LanguageStrings.OptionsProgramGhostscriptDescription
   fraProgGhostscript.Visible = True
  Case "ProgramSave"
   Set picOptions = LoadResPicture(2106, vbResIcon)
   lblOptions = LanguageStrings.OptionsProgramSaveDescription
   fraProgSave.Visible = True
  Case "ProgramAutosave"
   Set picOptions = LoadResPicture(2103, vbResIcon)
   lblOptions = LanguageStrings.OptionsProgramAutosaveDescription
   fraProgAutosave.Visible = True
  Case "ProgramFonts"
   Set picOptions = LoadResPicture(2102, vbResIcon)
   lblOptions = LanguageStrings.OptionsProgramFontDescription
   fraProgFont.Visible = True
  Case "ProgramDirectories"
   Set picOptions = LoadResPicture(2104, vbResIcon)
   lblOptions = LanguageStrings.OptionsProgramDirectoriesDescription
   fraProgDirectories.Visible = True
  Case "ProgramDocument"
   Set picOptions = LoadResPicture(2105, vbResIcon)
   lblOptions = LanguageStrings.OptionsProgramDocumentDescription
   fraProgDocument.Visible = True
  Case "Formats"
   Set picOptions = LoadResPicture(2111, vbResIcon)
   lblOptions = LanguageStrings.OptionsPDFDescription
   tbstrPDFOptions.Visible = True
   fraPDFGeneral.Visible = True
  Case "FormatsPDF"
   Set picOptions = LoadResPicture(2111, vbResIcon)
   lblOptions = LanguageStrings.OptionsPDFDescription
   tbstrPDFOptions.Visible = True
   tbstrPDFOptions.Tabs(1).Selected = True
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
  Case "FormatsPS"
   Set picOptions = LoadResPicture(2117, vbResIcon)
   lblOptions.Caption = LanguageStrings.OptionsPSDescription
   fraPSGeneral.Visible = True
   cmbPSLanguageLevel.Visible = True
   fraPSGeneral.Caption = LanguageStrings.OptionsPSDescription
  Case "FormatsEPS"
   Set picOptions = LoadResPicture(2118, vbResIcon)
   lblOptions.Caption = LanguageStrings.OptionsEPSDescription
   fraPSGeneral.Visible = True
   cmbEPSLanguageLevel.Visible = True
   fraPSGeneral.Caption = LanguageStrings.OptionsEPSDescription
 End Select
End Sub

Private Sub lsvFilenameSubst_Click()
 Set_txtFilenameSubst
End Sub

Private Sub optEncHigh_Click()
 UpdateSecurityFields
End Sub

Private Sub optEncLow_Click()
 UpdateSecurityFields
End Sub

Private Sub sldProcessPriority_Change()
 lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & sldProcessPriority.Text
End Sub

Private Sub sldProcessPriority_Scroll()
 With sldProcessPriority
  Select Case .Value
         Case 0: 'Idle
    .Text = LanguageStrings.OptionsProcesspriorityIdle
   Case 1: 'Normal
    .Text = LanguageStrings.OptionsProcesspriorityNormal
   Case 2: 'High
    .Text = LanguageStrings.OptionsProcesspriorityHigh
   Case 3: 'Realtime
    .Text = LanguageStrings.OptionsProcesspriorityRealtime
  End Select
 End With
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
   If SecurityIsPossible = False Then
    MsgBox LanguageStrings.MessagesMsg19
   End If
 End Select
End Sub

Private Sub trv_Click()
 ShowProgOptions
End Sub

Private Sub trv_NodeClick(ByVal Node As MSComctlLib.Node)
 ShowProgOptions
End Sub

Private Sub txtAutosaveFilename_Change()
 txtAutosaveFilename.ToolTipText = txtAutosaveFilename.Text
 txtAutoSavePreview.Text = GetSubstFilename("C:\test.pdf", txtAutosaveFilename.Text, , True) & ".pdf"
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

Private Sub ViewAutosave(ViewIt As Boolean)
 lblAutosaveformat.Enabled = ViewIt
 cmbAutosaveFormat.Enabled = ViewIt
 lblAutosaveFilename.Enabled = ViewIt
 txtAutosaveFilename.Enabled = ViewIt
 txtAutoSavePreview.Enabled = ViewIt
 lblAutosaveFilenameTokens.Enabled = ViewIt
 cmbAutoSaveFilenameTokens.Enabled = ViewIt
 chkUseAutosaveDirectory.Enabled = ViewIt
 If ViewIt = True Then
   cmbAutosaveFormat.BackColor = &H80000005
   cmbAutoSaveFilenameTokens.BackColor = &H80000005
   txtAutosaveFilename.BackColor = &H80000005
  Else
   cmbAutosaveFormat.BackColor = &H8000000F
   cmbAutoSaveFilenameTokens.BackColor = &H8000000F
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

Private Sub UpdateSecurityFields()
 If chkUseSecurity.Value = False Then
   fraPDFEncryptor.Enabled = False
   cmbPDFEncryptor.Enabled = False

   fraPDFEncLevel.Enabled = False
   optEncHigh.Enabled = False
   optEncLow.Enabled = False

   fraSecurityPass.Enabled = False
   chkUserPass.Enabled = False
   chkOwnerPass.Enabled = False

   fraPDFPermissions.Enabled = False
   chkAllowPrinting.Enabled = False
   chkAllowCopy.Enabled = False
   chkAllowModifyAnnotations.Enabled = False
   chkAllowModifyContents.Enabled = False

   fraPDFHighPermissions.Enabled = False
   chkAllowDegradedPrinting.Enabled = False
   chkAllowFillIn.Enabled = False
   chkAllowScreenReaders.Enabled = False
   chkAllowAssembly.Enabled = False
  Else
   fraPDFEncryptor.Enabled = True
   cmbPDFEncryptor.Enabled = True

   fraPDFEncLevel.Enabled = True
   If cmbPDFCompat.ListIndex >= 2 Then
     optEncHigh.Enabled = True
    Else
     optEncHigh.Enabled = False
   End If
   optEncLow.Enabled = True

   fraSecurityPass.Enabled = True
   chkUserPass.Enabled = True
   chkOwnerPass.Enabled = True

   fraPDFPermissions.Enabled = True
   chkAllowPrinting.Enabled = True
   chkAllowCopy.Enabled = True
   chkAllowModifyAnnotations.Enabled = True
   chkAllowModifyContents.Enabled = True

   If optEncHigh.Value = True Then
     fraPDFHighPermissions.Enabled = True
     chkAllowDegradedPrinting.Enabled = True
     chkAllowFillIn.Enabled = True
     chkAllowScreenReaders.Enabled = True
     chkAllowAssembly.Enabled = True
    Else
     fraPDFHighPermissions.Enabled = False
     chkAllowDegradedPrinting.Enabled = False
     chkAllowFillIn.Enabled = False
     chkAllowScreenReaders.Enabled = False
     chkAllowAssembly.Enabled = False
   End If
 End If
 If chkOwnerPass.Value = 0 And chkUserPass.Value = 0 Then
  chkOwnerPass.Value = 1: Options.PDFOwnerPass = 1
 End If
End Sub

Private Sub cmdAsso_Click()
 PsAssociate
 SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
 cmdAsso.Enabled = False
End Sub

Private Sub AddFilenameSubstitutions()
 Dim i As Long, res As Long
 res = CheckFilenameSubstitutions(0)
 Select Case res
        Case 0:
   lsvFilenameSubst.ListItems.Add , , txtFilenameSubst(0).Text
   lsvFilenameSubst.ListItems(lsvFilenameSubst.ListItems.Count).SubItems(1) = txtFilenameSubst(1).Text
   lsvFilenameSubst.ListItems(lsvFilenameSubst.ListItems.Count).Selected = True
   lsvFilenameSubst.ListItems(lsvFilenameSubst.ListItems.Count).EnsureVisible
   Set_txtFilenameSubst
  Case 2:
   MsgBox LanguageStrings.MessagesMsg12 & _
    vbCrLf & vbTab & "\ / : * ? < > | """
  Case 3:
   MsgBox LanguageStrings.MessagesMsg11
 End Select
End Sub

Private Sub ChangeFilenameSubstitutions()
 Dim i As Long, res As Long
 res = CheckFilenameSubstitutions(lsvFilenameSubst.SelectedItem.Index)
 Select Case res
        Case 0:
   lsvFilenameSubst.SelectedItem.Text = txtFilenameSubst(0).Text
   lsvFilenameSubst.SelectedItem.SubItems(1) = txtFilenameSubst(1).Text
  Case 2:
   MsgBox LanguageStrings.MessagesMsg12 & _
    vbCrLf & vbTab & "\ / : * ? < > | """
  Case 3:
   MsgBox LanguageStrings.MessagesMsg11
 End Select
End Sub

Private Sub DeleteFilenameSubstitutions()
 Dim oIndex As Long
 With lsvFilenameSubst
  oIndex = .SelectedItem.Index
  If .ListItems.Count > 0 Then
   .ListItems.Remove .SelectedItem.Index
   If oIndex > .ListItems.Count Then
    oIndex = .ListItems.Count
   End If
   If .ListItems.Count > 0 Then
    .ListItems(oIndex).Selected = True
    .ListItems(oIndex).EnsureVisible
   End If
   Set_txtFilenameSubst
  End If
 End With
End Sub

Private Sub MoveUpFilenameSubstitutions()
 Dim tStrL As String, tStrR As String
 With lsvFilenameSubst
  tStrL = .ListItems(.SelectedItem.Index).Text
  tStrR = .ListItems(.SelectedItem.Index).SubItems(1)
  .ListItems(.SelectedItem.Index).Text = .ListItems(.SelectedItem.Index - 1).Text
  .ListItems(.SelectedItem.Index).SubItems(1) = .ListItems(.SelectedItem.Index - 1).SubItems(1)
  .ListItems(.SelectedItem.Index - 1).Text = tStrL
  .ListItems(.SelectedItem.Index - 1).SubItems(1) = tStrR
  .ListItems(.SelectedItem.Index - 1).Selected = True
  .ListItems(.SelectedItem.Index).EnsureVisible
 End With
 Set_txtFilenameSubst
End Sub

Private Sub MoveDownFilenameSubstitutions()
 Dim tStrL As String, tStrR As String
 With lsvFilenameSubst
  tStrL = .ListItems(.SelectedItem.Index).Text
  tStrR = .ListItems(.SelectedItem.Index).SubItems(1)
  .ListItems(.SelectedItem.Index).Text = .ListItems(.SelectedItem.Index + 1).Text
  .ListItems(.SelectedItem.Index).SubItems(1) = .ListItems(.SelectedItem.Index + 1).SubItems(1)
  .ListItems(.SelectedItem.Index + 1).Text = tStrL
  .ListItems(.SelectedItem.Index + 1).SubItems(1) = tStrR
  .ListItems(.SelectedItem.Index + 1).Selected = True
  .ListItems(.SelectedItem.Index).EnsureVisible
 End With
 Set_txtFilenameSubst
End Sub

Private Function CheckFilenameSubstitutions(Index As Long) As Long
 Dim i As Long
 CheckFilenameSubstitutions = 0
 If Len(txtFilenameSubst(0).Text) = 0 Then
  CheckFilenameSubstitutions = 1
  Exit Function
 End If
 If IsForbiddenChars(txtFilenameSubst(0).Text) = True Then
  txtFilenameSubst(0).SetFocus
  CheckFilenameSubstitutions = 2
  Exit Function
 End If
 If IsForbiddenChars(txtFilenameSubst(1).Text) = True Then
  txtFilenameSubst(1).SetFocus
  CheckFilenameSubstitutions = 2
  Exit Function
 End If
 If Index = 0 Then
   For i = 1 To lsvFilenameSubst.ListItems.Count
    If UCase$(txtFilenameSubst(0).Text) = UCase$(lsvFilenameSubst.ListItems(i).Text) Then
     CheckFilenameSubstitutions = 3
     Exit Function
    End If
   Next i
  Else
   For i = 1 To lsvFilenameSubst.ListItems.Count
    If UCase$(txtFilenameSubst(0).Text) = UCase$(lsvFilenameSubst.ListItems(i).Text) And _
     Index <> lsvFilenameSubst.SelectedItem.Index Then
     CheckFilenameSubstitutions = 3
     Exit Function
    End If
   Next i
 End If
End Function

Private Sub CheckCmdFilenameSubst()
 If lsvFilenameSubst.ListItems.Count > 0 Then
   cmdFilenameSubst(1).Enabled = True
   cmdFilenameSubst(2).Enabled = True
  Else
   cmdFilenameSubst(1).Enabled = False
   cmdFilenameSubst(2).Enabled = False
 End If
 If lsvFilenameSubst.ListItems.Count > 1 Then
   cmdFilenameSubst(3).Enabled = True
   cmdFilenameSubst(4).Enabled = True
  Else
   cmdFilenameSubst(3).Enabled = False
   cmdFilenameSubst(4).Enabled = False
 End If
 If lsvFilenameSubst.ListItems.Count > 0 Then
  If lsvFilenameSubst.SelectedItem.Index = 1 Then
   cmdFilenameSubst(3).Enabled = False
  End If
  If lsvFilenameSubst.SelectedItem.Index = lsvFilenameSubst.ListItems.Count Then
   cmdFilenameSubst(4).Enabled = False
  End If
 End If
End Sub

Private Sub Set_txtFilenameSubst()
 CheckCmdFilenameSubst
 If lsvFilenameSubst.ListItems.Count > 0 Then
  txtFilenameSubst(0).Text = lsvFilenameSubst.SelectedItem.Text
  txtFilenameSubst(0).ToolTipText = txtFilenameSubst(0).Text
  txtFilenameSubst(1).Text = lsvFilenameSubst.SelectedItem.SubItems(1)
  txtFilenameSubst(1).ToolTipText = txtFilenameSubst(1).Text
 End If
End Sub

Private Sub txtSaveFilename_Change()
 txtSaveFilename.ToolTipText = txtSaveFilename.Text
 txtSavePreview.Text = GetSubstFilename("C:\test.pdf", txtSaveFilename.Text, , True) & ".pdf"
End Sub
