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
   Begin PDFCreator.dmFrame dmFraProgFont 
      Height          =   4695
      Left            =   2640
      TabIndex        =   65
      Top             =   1440
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   8281
      Caption         =   "Programfont"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin PDFCreator.dmFrame dmFraProgDirectories 
      Height          =   1095
      Left            =   2640
      TabIndex        =   43
      Top             =   1320
      Visible         =   0   'False
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1931
      Caption         =   "Directories"
      Caption3D       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextShaddowColor=   12582912
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
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   600
         Width           =   4965
      End
      Begin VB.CommandButton cmdUsertempPath 
         Height          =   300
         Left            =   5640
         Picture         =   "frmOptions.frx":548A
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
   Begin PDFCreator.dmFrame dmFraPDFGeneral 
      Height          =   2895
      Left            =   2730
      TabIndex        =   91
      Top             =   1785
      Visible         =   0   'False
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   5106
      Caption         =   "General Options"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
         ItemData        =   "frmOptions.frx":5A14
         Left            =   2400
         List            =   "frmOptions.frx":5A16
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
         ItemData        =   "frmOptions.frx":5A18
         Left            =   2400
         List            =   "frmOptions.frx":5A1A
         Style           =   2  'Dropdown-Liste
         TabIndex        =   93
         Top             =   480
         Width           =   2655
      End
      Begin VB.ComboBox cmbPDFRotate 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "frmOptions.frx":5A1C
         Left            =   2400
         List            =   "frmOptions.frx":5A1E
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
         Top             =   1440
         Width           =   210
      End
      Begin VB.Label lblPDFOverprint 
         Alignment       =   1  'Rechts
         Caption         =   "Overprint:"
         Height          =   375
         Left            =   120
         TabIndex        =   100
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label lblPDFResolution 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Resolution:"
         Height          =   255
         Left            =   120
         TabIndex        =   99
         Top             =   1500
         Width           =   2175
      End
      Begin VB.Label lblPDFCompat 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Compatibility:"
         Height          =   255
         Left            =   120
         TabIndex        =   98
         Top             =   540
         Width           =   2175
      End
      Begin VB.Label lblPDFAutoRotate 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Auto-Rotate Pages:"
         Height          =   255
         Left            =   120
         TabIndex        =   97
         Top             =   1020
         Width           =   2175
      End
   End
   Begin PDFCreator.dmFrame dmFraPDFSecurity 
      Height          =   5535
      Left            =   2730
      TabIndex        =   137
      Top             =   2205
      Visible         =   0   'False
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   9763
      Caption         =   "Security"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin PDFCreator.dmFrame dmFraPDFHighPermissions 
         Height          =   855
         Left            =   120
         TabIndex        =   152
         Top             =   4560
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   1508
         Caption         =   "Enhanced permissions (128 Bit only)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         _ExtentX        =   10504
         _ExtentY        =   1508
         Caption         =   "Disallow user to"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         _ExtentX        =   10504
         _ExtentY        =   1508
         Caption         =   "Passwords"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         _ExtentX        =   10504
         _ExtentY        =   1508
         Caption         =   "Encryption level"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         _ExtentX        =   10504
         _ExtentY        =   1508
         Caption         =   "Encryptor"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox cmbPDFEncryptor 
            Appearance      =   0  '2D
            Height          =   315
            ItemData        =   "frmOptions.frx":5A20
            Left            =   120
            List            =   "frmOptions.frx":5A22
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
      _ExtentX        =   10927
      _ExtentY        =   7646
      Caption         =   "Compression"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin PDFCreator.dmFrame dmFraPDFMono 
         Height          =   1095
         Left            =   120
         TabIndex        =   118
         Top             =   3120
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   1931
         Caption         =   "Monochrome images"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
            ItemData        =   "frmOptions.frx":5A24
            Left            =   120
            List            =   "frmOptions.frx":5A26
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
            ItemData        =   "frmOptions.frx":5A28
            Left            =   2520
            List            =   "frmOptions.frx":5A2A
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
         _ExtentX        =   10504
         _ExtentY        =   1931
         Caption         =   "Greyscale images"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
            ItemData        =   "frmOptions.frx":5A2C
            Left            =   2520
            List            =   "frmOptions.frx":5A2E
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
            ItemData        =   "frmOptions.frx":5A30
            Left            =   120
            List            =   "frmOptions.frx":5A32
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
         _ExtentX        =   10504
         _ExtentY        =   1931
         Caption         =   "Color images"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
            ItemData        =   "frmOptions.frx":5A34
            Left            =   2520
            List            =   "frmOptions.frx":5A36
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
            ItemData        =   "frmOptions.frx":5A38
            Left            =   120
            List            =   "frmOptions.frx":5A3A
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
      _ExtentX        =   10927
      _ExtentY        =   2143
      Caption         =   "Color options"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cmbPDFColorModel 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "frmOptions.frx":5A3C
         Left            =   120
         List            =   "frmOptions.frx":5A3E
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
      _ExtentX        =   10927
      _ExtentY        =   2566
      Caption         =   "Options"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin PDFCreator.dmFrame dmFraPDFFonts 
      Height          =   1695
      Left            =   2760
      TabIndex        =   125
      Top             =   2400
      Visible         =   0   'False
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   2990
      Caption         =   "Font options"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
         Left            =   360
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
         Top             =   1320
         Width           =   120
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
      _ExtentX        =   4471
      _ExtentY        =   13917
      FontName        =   "MS Sans Serif"
      FontCharset     =   0
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
               Picture         =   "frmOptions.frx":5A40
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":5FDA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":6574
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":6B0E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":70A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":7442
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":79DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":7F76
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":8510
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":8AAA
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":9044
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":95DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":9B78
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":A112
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":A6AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":AC46
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":B520
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
      _ExtentX        =   11245
      _ExtentY        =   4471
      Caption         =   "Filename substitutions"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
         Picture         =   "frmOptions.frx":BDFA
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
         Picture         =   "frmOptions.frx":C184
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
      _ExtentX        =   11324
      _ExtentY        =   1879
      Caption         =   ""
      BarColorFrom    =   8421631
      BarColorTo      =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin PDFCreator.dmFrame dmFraPSGeneral 
      Height          =   1095
      Left            =   2640
      TabIndex        =   87
      Top             =   1920
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   1931
      Caption         =   "Postscript"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
         Caption         =   "Language Level:"
         Height          =   255
         Left            =   240
         TabIndex        =   90
         Top             =   480
         Width           =   1695
      End
   End
   Begin PDFCreator.dmFrame dmFraBitmapGeneral 
      Height          =   1935
      Left            =   2640
      TabIndex        =   74
      Top             =   1680
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3413
      Caption         =   "Bitmap"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
         Caption         =   "%"
         Height          =   255
         Left            =   2520
         TabIndex        =   86
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label lblJPEGQuality 
         Alignment       =   1  'Rechts
         Caption         =   "Quality:"
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblBitmapColors 
         Alignment       =   1  'Rechts
         Caption         =   "Colors:"
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblBitmapDPI 
         Caption         =   "dpi"
         Height          =   255
         Left            =   2520
         TabIndex        =   83
         Top             =   480
         Width           =   495
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
   End
   Begin PDFCreator.dmFrame dmfraProgSave 
      Height          =   1935
      Left            =   2640
      TabIndex        =   52
      Top             =   2160
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3413
      Caption         =   "Save"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
         ItemData        =   "frmOptions.frx":C50E
         Left            =   3720
         List            =   "frmOptions.frx":C510
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
      _ExtentX        =   11245
      _ExtentY        =   3413
      Caption         =   "Document"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox cmbAuthorTokens 
         Appearance      =   0  '2D
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmOptions.frx":C512
         Left            =   3720
         List            =   "frmOptions.frx":C514
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
   Begin PDFCreator.dmFrame dmFraProgAutosave 
      Height          =   3855
      Left            =   2640
      TabIndex        =   31
      Top             =   1200
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6800
      Caption         =   "Autosave"
      Caption3D       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextShaddowColor=   12582912
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
         ItemData        =   "frmOptions.frx":C516
         Left            =   3690
         List            =   "frmOptions.frx":C518
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
   Begin PDFCreator.dmFrame dmFraProgGeneral 
      Height          =   4215
      Left            =   2640
      TabIndex        =   10
      Top             =   1050
      Visible         =   0   'False
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   7435
      Caption         =   "General"
      Caption3D       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextShaddowColor=   12582912
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
         _ExtentX        =   10821
         _ExtentY        =   53
         LineType        =   1
         3DHighlight     =   -2147483628
         3DShadow        =   -2147483632
         DrawStyle       =   0
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
         _ExtentX        =   10821
         _ExtentY        =   53
         LineType        =   1
         3DHighlight     =   -2147483628
         3DShadow        =   -2147483632
         DrawStyle       =   0
      End
      Begin PDFCreator.Line3D Line3D1 
         Height          =   30
         Index           =   2
         Left            =   120
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   3360
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   53
         LineType        =   1
         3DHighlight     =   -2147483628
         3DShadow        =   -2147483632
         DrawStyle       =   0
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
   Begin PDFCreator.dmFrame dmFraProgGhostscript 
      Height          =   1050
      Left            =   2625
      TabIndex        =   13
      Top             =   945
      Visible         =   0   'False
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1852
      Caption         =   "Ghostscript"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextShaddowColor=   12582912
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
         Top             =   1380
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtGSbin 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1380
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.TextBox txtGSlib 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1980
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.TextBox txtGSfonts 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2580
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.CommandButton cmdGetgslibDirectory 
         Caption         =   "..."
         Height          =   255
         Left            =   5625
         TabIndex        =   17
         Top             =   1980
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdGetgsfontsDirectory 
         Caption         =   "..."
         Height          =   255
         Left            =   5625
         TabIndex        =   16
         Top             =   2580
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtGSresource 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   3180
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.CommandButton cmdGetgsresourceDirectory 
         Caption         =   "..."
         Height          =   255
         Left            =   5625
         TabIndex        =   14
         Top             =   3180
         Visible         =   0   'False
         Width           =   375
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
         Top             =   1140
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label lblGSlib 
         AutoSize        =   -1  'True
         Caption         =   "Ghostscript Libraries"
         Height          =   195
         Left            =   105
         TabIndex        =   25
         Top             =   1740
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label lblGSfonts 
         AutoSize        =   -1  'True
         Caption         =   "Ghostscript Fonts"
         Height          =   195
         Left            =   105
         TabIndex        =   24
         Top             =   2340
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label lblGhostscriptResource 
         AutoSize        =   -1  'True
         Caption         =   "Ghostscript Resource"
         Height          =   195
         Left            =   105
         TabIndex        =   23
         Top             =   2940
         Visible         =   0   'False
         Width           =   1530
      End
   End
   Begin PDFCreator.dmFrame dmFraShellIntegration 
      Height          =   1065
      Left            =   2640
      TabIndex        =   12
      Top             =   5565
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1879
      Caption         =   "Shell integration"
      Caption3D       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextShaddowColor=   12582912
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
50161 Select Case ErrPtnr.OnError("frmOptions", "cmbAutosaveFormat_Click")
      Case 0: Resume
50180 Case 1: Resume Next
50190 Case 2: Exit Sub
50200 Case 3: End
50210 End Select
50220 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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
50040  Dim allow As String, tstr As String
50050  allow = "0123456789" & Chr$(8) & Chr$(13)
50060  tstr = Chr$(KeyAscii)
50070  If InStr(1, allow, tstr) = 0 Then
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
50040  Dim i As Long, tstr As String
50050  tstr = ""
50060  For i = 1 To Len(cmbCharset.Text)
50070   If InStr("0123456789", Mid(cmbCharset.Text, i, 1)) > 0 Then
50080     tstr = tstr & Mid(cmbCharset.Text, i, 1)
50090    Else
50100     Exit For
50110   End If
50120  Next i
50130  If Len(Trim$(tstr)) = 0 Then
50140    cmbCharset.Text = 0
50150   Else
50160    cmbCharset.Text = tstr
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
50040  Dim reg As clsRegistry, gsv As String, tsf() As String, Path As String, tstr As String
50050
50060  gsv = cmbGhostscript.List(cmbGhostscript.ListIndex)
50070  Set reg = New clsRegistry: reg.hkey = HKEY_LOCAL_MACHINE
50080
50090  If InStr(gsv, ":") Then
50100    reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50110    txtGSbin.Text = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryBinaries"))
50120    txtGSfonts.Text = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
50130    txtGSlib.Text = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryLibraries"))
50140    txtGSresource.Text = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryResource"))
50150    Set reg = Nothing
50160    Exit Sub
50170   Else
50180    If InStr(UCase$(gsv), "AFPL") Then
50190     If InStr(gsv, " ") > 0 Then
50200      tsf = Split(gsv, " ")
50210      reg.KeyRoot = "SOFTWARE\AFPL Ghostscript\" & tsf(UBound(tsf))
50220      tstr = reg.GetRegistryValue("GS_DLL")
50230      SplitPath tstr, , Path
50240      txtGSbin.Text = CompletePath(Path)
50250      If InStrRev(Path, "\") > 0 Then
50260       txtGSlib.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50270       txtGSfonts.Text = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50280       If tsf(UBound(tsf)) <> "8.00" Then
50290        txtGSresource.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "Resource"
50300       End If
50310      End If
50320     End If
50330    End If
50340    If InStr(UCase$(gsv), "GNU") Then
50350     If InStr(gsv, " ") > 0 Then
50360      tsf = Split(gsv, " ")
50370      reg.KeyRoot = "SOFTWARE\GNU Ghostscript\" & tsf(UBound(tsf))
50380      tstr = reg.GetRegistryValue("GS_DLL")
50390      SplitPath tstr, , Path
50400      txtGSbin.Text = CompletePath(Path)
50410      If InStrRev(Path, "\") > 0 Then
50420       txtGSlib.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50430       txtGSfonts.Text = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50440       txtGSresource.Text = ""
50450      End If
50460     End If
50470    End If
50480    If InStr(UCase$(gsv), "GPL") Then
50490     If InStr(gsv, " ") > 0 Then
50500      tsf = Split(gsv, " ")
50510      reg.KeyRoot = "SOFTWARE\GPL Ghostscript\" & tsf(UBound(tsf))
50520      tstr = reg.GetRegistryValue("GS_DLL")
50530      SplitPath tstr, , Path
50540      txtGSbin.Text = CompletePath(Path)
50550      If InStrRev(Path, "\") > 0 Then
50560       txtGSlib.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50570       txtGSfonts.Text = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50580       txtGSresource.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "Resource"
50590      End If
50600     End If
50610    End If
50620  End If
50630  Set reg = Nothing
50640 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50650 Exit Sub
ErrPtnr_OnError:
50671 Select Case ErrPtnr.OnError("frmOptions", "cmbGhostscript_Click")
      Case 0: Resume
50690 Case 1: Resume Next
50700 Case 2: Exit Sub
50710 Case 3: End
50720 End Select
50730 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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
50080   ieb.Refresh
50090  End With
50100 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50110 Exit Sub
ErrPtnr_OnError:
50131 Select Case ErrPtnr.OnError("frmOptions", "cmdCancelTest_Click")
      Case 0: Resume
50150 Case 1: Resume Next
50160 Case 2: Exit Sub
50170 Case 3: End
50180 End Select
50190 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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
50110  End Select
50120 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50130 Exit Sub
ErrPtnr_OnError:
50151 Select Case ErrPtnr.OnError("frmOptions", "cmdFilenameSubst_Click")
      Case 0: Resume
50170 Case 1: Resume Next
50180 Case 2: Exit Sub
50190 Case 3: End
50200 End Select
50210 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdFilenameSubstMove_Click(Index As Integer)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50041  Select Case Index
        Case 0: ' Up
50060    MoveUpFilenameSubstitutions
50070   Case 1: ' Down
50080    MoveDownFilenameSubstitutions
50090  End Select
50100 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50110 Exit Sub
ErrPtnr_OnError:
50131 Select Case ErrPtnr.OnError("frmOptions", "cmdFilenameSubstMove_Click")
      Case 0: Resume
50150 Case 1: Resume Next
50160 Case 2: Exit Sub
50170 Case 3: End
50180 End Select
50190 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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
50130    ieb.Refresh
50140   End With
50150  End If
50160 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50170 Exit Sub
ErrPtnr_OnError:
50191 Select Case ErrPtnr.OnError("frmOptions", "cmdReset_Click")
      Case 0: Resume
50210 Case 1: Resume Next
50220 Case 2: Exit Sub
50230 Case 3: End
50240 End Select
50250 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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
50040  Dim tCharset As Long, tstr As String, tFontSize As Long, tFontname As String, _
  tFontCharset As Long
50060  If InStr(1, cmbCharset.Text, ",", vbTextCompare) > 0 Then
50070    tstr = Trim$(Mid$(cmbCharset.Text, 1, InStr(1, cmbCharset.Text, ",", vbTextCompare) - 1))
50080   Else
50090    tstr = Trim$(cmbCharset.Text)
50100  End If
50110  If Len(tstr) = 0 Then
50120   cmbCharset.Text = 0
50130   Exit Sub
50140  End If
50150  If IsNumeric(tstr) = False Then
50160   cmbCharset.Text = 0
50170   Exit Sub
50180  End If
50190  tCharset = tstr
50200  With cmdTest.Font
50210   tFontname = .Name
50220   tFontSize = .Size
50230   tFontCharset = .Charset
50240  End With
50250  SetFont Me, cmbFonts.List(cmbFonts.ListIndex), CLng(tstr), cmbProgramFontsize.Text
50260  cmbCharset.Text = tCharset
50270  SetFont frmMain, cmbFonts.List(cmbFonts.ListIndex), CLng(tstr), cmbProgramFontsize.Text
50280  ieb.Refresh
50290  With cmdTest.Font
50300   .Name = tFontname
50310   .Size = tFontSize
50320   .Charset = tFontCharset
50330  End With
50340  With cmdCancelTest
50350   .Font.Name = tFontname
50360   .Font.Size = tFontSize
50370   .Font.Charset = tFontCharset
50380   .Enabled = True
50390  End With
50400 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50410 Exit Sub
ErrPtnr_OnError:
50431 Select Case ErrPtnr.OnError("frmOptions", "cmdTest_Click")
      Case 0: Resume
50450 Case 1: Resume Next
50460 Case 2: Exit Sub
50470 Case 3: End
50480 End Select
50490 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdTestpage_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim TestPSPage As String, fn As Long, Filename As String, tstr As String, _
  c As Collection
50060  frmMain.Timer1.Enabled = False
50070  TestPSPage = GetTestpageFromRessource
50080  TestPSPage = Replace(TestPSPage, "[INFOTITLE]", LanguageStrings.OptionsTestpage, , 1, vbTextCompare)
50090  TestPSPage = Replace(TestPSPage, "[INFORELEASE]", App.Title & " " & GetProgramReleaseStr, , 1, vbTextCompare)
50100  TestPSPage = Replace(TestPSPage, "[INFODATE]", Now, , 1, vbTextCompare)
50110  TestPSPage = Replace(TestPSPage, "[INFOAUTHORS]", "Philip Chinery, Frank Heind\224rfer", , 1, vbTextCompare)
50120  TestPSPage = Replace(TestPSPage, "[INFOHOMEPAGE]", Homepage, , 1, vbTextCompare)
50130  tstr = CompletePath(App.Path) & "PDFCreator.exe"
50140  If FileExists(tstr) = True Then
50150    Set c = GetFileVersion(tstr)
50160    tstr = "Version: " & c(2) & "; Size: " & Format(FileLen(tstr), "###,###,###,### Bytes")
50170   Else
50180    tstr = ""
50190  End If
50200  TestPSPage = Replace(TestPSPage, "[INFOPDFCREATOR]", tstr, , 1, vbTextCompare)
50210
50220  tstr = CompletePath(GetSystemDirectory()) & "PDFSpooler.exe"
50230  If FileExists(tstr) = True Then
50240    Set c = GetFileVersion(tstr)
50250    tstr = "Version: " & c(2) & "; Size: " & Format(FileLen(tstr), "###,###,###,### Bytes")
50260   Else
50270    tstr = ""
50280  End If
50290  TestPSPage = Replace(TestPSPage, "[INFOPDFSPOOLER]", tstr, , 1, vbTextCompare)
50300
50310  tstr = CompletePath(App.Path) & "Languages\Transtool.exe"
50320  If FileExists(tstr) = True Then
50330    Set c = GetFileVersion(tstr)
50340    tstr = "Version: " & c(2) & "; Size: " & Format(FileLen(tstr), "###,###,###,### Bytes")
50350   Else
50360    tstr = ""
50370  End If
50380  TestPSPage = Replace(TestPSPage, "[INFOTRANSTOOL]", tstr, , 1, vbTextCompare)
50390
50400  TestPSPage = Replace(TestPSPage, "[INFOCOMPUTER]", GetComputerName, , 1, vbTextCompare)
50410  tstr = GetWinVersionStr
50420  TestPSPage = Replace(TestPSPage, "[INFOWINDOWS]", _
  Mid(tstr, 1, IIf(InStr(1, tstr, "[") > 0, InStr(1, tstr, "[") - 1, Len(tstr))), 1, vbTextCompare)
50440
50450  fn = FreeFile
50460  tstr = CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory
50470  If DirExists(tstr) = False Then
50480   MakePath tstr
50490  End If
50500  Filename = GetTempFile(tstr, "~PS")
50510  Open Filename For Output As fn
50520  Print #fn, TestPSPage
50530  Close #fn
50540  frmMain.CheckPrintJobs
50550  frmMain.Timer1.Enabled = True
50560 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50570 Exit Sub
ErrPtnr_OnError:
50591 Select Case ErrPtnr.OnError("frmOptions", "cmdTestpage_Click")
      Case 0: Resume
50610 Case 1: Resume Next
50620 Case 2: Exit Sub
50630 Case 3: End
50640 End Select
50650 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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
50060     Call HTMLHelp_ShowTopic("html\generalsettings.htm")
50070 '  MsgBox ieb.GetSelectedGroup & vbCrLf & ieb.GetSelectedItem
50081    Select Case ieb.GetSelectedGroup
          Case 1
50101      Select Case ieb.GetSelectedItem
            Case 1
50120        Call HTMLHelp_ShowTopic("html\generalsettings.htm")
50130       Case 2
50140        Call HTMLHelp_ShowTopic("html\ghostscript.htm")
50150       Case 3
50160        Call HTMLHelp_ShowTopic("html\docproperties.htm")
50170       Case 4
50180        Call HTMLHelp_ShowTopic("html\savesettings.htm")
50190       Case 5
50200        Call HTMLHelp_ShowTopic("html\autosave.htm")
50210       Case 6
50220        Call HTMLHelp_ShowTopic("html\directories.htm")
50230       Case 7
50240        Call HTMLHelp_ShowTopic("html\fontsetting.htm")
50250       Case Else
50260        Call HTMLHelp_ShowTopic("html\generalsettings.htm")
50270      End Select
50280     Case 2
50291      Select Case ieb.GetSelectedItem
            Case 1
50311        Select Case tbstrPDFOptions.SelectedItem.Index
              Case 1
50330          Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
50340         Case 2
50350          Call HTMLHelp_ShowTopic("html\pdfcompression.htm")
50360         Case 3
50370          Call HTMLHelp_ShowTopic("html\pdffonts.htm")
50380         Case 4
50390          Call HTMLHelp_ShowTopic("html\pdfcolors.htm")
50400         Case 5
50410          Call HTMLHelp_ShowTopic("html\pdfsecurity.htm")
50420         Case Else
50430          Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
50440        End Select
50450       Case 2
50460        Call HTMLHelp_ShowTopic("html\pngsettings.htm")
50470       Case 3
50480        Call HTMLHelp_ShowTopic("html\jpegsettings.htm")
50490       Case 4
50500        Call HTMLHelp_ShowTopic("html\bmpsettings.htm")
50510       Case 5
50520        Call HTMLHelp_ShowTopic("html\pcxsettings.htm")
50530       Case 6
50540        Call HTMLHelp_ShowTopic("html\tiffsettings.htm")
50550       Case 7
50560        Call HTMLHelp_ShowTopic("html\pssettings.htm")
50570       Case 8
50580        Call HTMLHelp_ShowTopic("html\epssettings.htm")
50590       Case Else
50600        Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
50610      End Select
50620    End Select
50630  End If
50640
50650 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50660 Exit Sub
ErrPtnr_OnError:
50681 Select Case ErrPtnr.OnError("frmOptions", "Form_KeyDown")
      Case 0: Resume
50700 Case 1: Resume Next
50710 Case 2: Exit Sub
50720 Case 3: End
50730 End Select
50740 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Load()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Const fraPDFTop = 1360, fraPDFLeft = 2960
50050  Dim pic As New StdPicture, i As Long, tstr As String, gsvers As Collection, _
  fc As Long, reg As clsRegistry, tsf() As String, tStr2 As String, _
  ctl As Control
50080
50090  KeyPreview = True
50100
50110  With Screen
50120   .MousePointer = vbHourglass
50130   Move (.Width - Width) / 2, (.Height - Height) / 2
50140  End With
50150
50160  For Each ctl In Controls
50170   If TypeOf ctl Is dmFrame Then
50180    ctl.Font.Size = 10
50190    ctl.TextShaddowColor = &HC00000
50200    ctl.Caption3D = [Raised Caption]
50210    If ComputerScreenResolution <= 8 Or Options.OptionsDesign = 2 Then
50220     ctl.UseGradient = False: ctl.Caption3D = [Flat Caption]
50230     If UCase$(ctl.Name) = "DMFRADESCRIPTION" Then
50240       ctl.BarColorFrom = vbRed
50250      Else
50260       ctl.BarColorFrom = vbBlue
50270     End If
50280    End If
50290   End If
50300  Next ctl
50310
50320  With dmFraDescription
50330   .Caption = LanguageStrings.OptionsTreeProgram
50340   .Visible = True
50350  End With
50360  dmFraShellIntegration.Visible = True
50370  With dmFraProgGeneral
50380   .Visible = True
50390   .Top = dmFraDescription.Top + dmFraDescription.Height + 50
50400   .Left = dmFraDescription.Left
50410   dmFraShellIntegration.Top = .Top + .Height + 50
50420   dmFraShellIntegration.Left = .Left
50430   dmFraShellIntegration.Width = .Width
50440   dmFraProgGhostscript.Top = .Top
50450   dmFraProgGhostscript.Left = .Left
50460   dmFraProgGhostscript.Width = .Width
50470   dmFraProgAutosave.Top = .Top
50480   dmFraProgAutosave.Left = .Left
50490   dmFraProgAutosave.Width = .Width
50500   dmFraProgDirectories.Top = .Top
50510   dmFraProgDirectories.Left = .Left
50520   dmFraProgDirectories.Width = .Width
50530   dmFraProgDocument.Top = .Top
50540   dmFraProgDocument.Left = .Left
50550   dmFraProgDocument.Width = .Width
50560   dmfraProgSave.Top = .Top
50570   dmfraProgSave.Left = .Left
50580   dmfraProgSave.Width = .Width
50590   dmfraFilenameSubstitutions.Top = dmfraProgSave.Top + dmfraProgSave.Height + 50
50600   dmfraFilenameSubstitutions.Left = .Left
50610   dmfraFilenameSubstitutions.Width = .Width
50620   dmFraProgFont.Top = .Top
50630   dmFraProgFont.Left = .Left
50640   dmFraProgFont.Width = .Width
50650   dmFraBitmapGeneral.Top = .Top
50660   dmFraBitmapGeneral.Left = .Left
50670   dmFraBitmapGeneral.Width = .Width
50680   dmFraPSGeneral.Top = .Top
50690   dmFraPSGeneral.Left = .Left
50700   dmFraPSGeneral.Width = .Width
50710
50720   cmdCancel.Left = .Left
50730   cmdReset.Left = .Left + (.Width - cmdReset.Width) / 2
50740   cmdSave.Left = .Left + .Width - cmdSave.Width
50750  End With
50760
50770  With tbstrPDFOptions
50780   .Top = dmFraDescription.Top + dmFraDescription.Height + 50
50790   .Left = dmFraDescription.Left
50800   .Height = cmdCancel.Top - tbstrPDFOptions.Top - 50
50810   .Width = dmFraDescription.Width
50820  End With
50830
50840  With dmFraPDFGeneral
50850   .Top = tbstrPDFOptions.ClientTop + 100
50860   .Left = tbstrPDFOptions.Left + (tbstrPDFOptions.Width - .Width) / 2
50870   dmfraPDFCompress.Top = .Top
50880   dmfraPDFCompress.Left = .Left
50890   dmFraPDFFonts.Top = .Top
50900   dmFraPDFFonts.Left = .Left
50910   dmFraPDFColors.Top = .Top
50920   dmFraPDFColors.Left = .Left
50930   dmFraPDFColorOptions.Top = dmFraPDFColors.Top + dmFraPDFColors.Height + 50
50940   dmFraPDFColorOptions.Left = .Left
50950   dmFraPDFSecurity.Top = .Top
50960   dmFraPDFSecurity.Left = .Left
50970  End With
50980
50990  cmbEPSLanguageLevel.Top = cmbPSLanguageLevel.Top
51000  cmbEPSLanguageLevel.Left = cmbPSLanguageLevel.Left
51010
51020  ieb.DisableUpdates True
51030  ieb.ClearStructure
51040  ieb.SetImageList imlIeb
51050 ' trv.Nodes.Clear
51060 ' trv.Indentation = 200
51070  With LanguageStrings
51080   ieb.AddGroup "Program", .OptionsTreeProgram, 0
51090   ieb.AddItem "Program", "General", .OptionsProgramGeneralSymbol, 1
51100   ieb.AddItem "Program", "Ghostscript", .OptionsProgramGhostscriptSymbol, 2
51110   ieb.AddItem "Program", "Document", .OptionsProgramDocumentSymbol, 3
51120   ieb.AddItem "Program", "Save", .OptionsProgramSaveSymbol, 4
51130   ieb.AddItem "Program", "AutoSave", .OptionsProgramAutosaveSymbol, 5
51140   ieb.AddItem "Program", "Directories", .OptionsProgramDirectoriesSymbol, 6
51150   ieb.AddItem "Program", "Fonts", .OptionsProgramFontSymbol, 7
51160   ieb.AddGroup "Formats", .OptionsTreeFormats, 0
51170   ieb.AddItem "Formats", "PDF", .OptionsPDFSymbol, 8
51180   ieb.AddItem "Formats", "PNG", .OptionsPNGSymbol, 9
51190   ieb.AddItem "Formats", "JPEG", .OptionsJPEGSymbol, 10
51200   ieb.AddItem "Formats", "BMP", .OptionsBMPSymbol, 11
51210   ieb.AddItem "Formats", "PCX", .OptionsPCXSymbol, 12
51220   ieb.AddItem "Formats", "TIFF", .OptionsTIFFSymbol, 13
51230   ieb.AddItem "Formats", "PS", .OptionsPSSymbol, 14
51240   ieb.AddItem "Formats", "EPS", .OptionsEPSSymbol, 15
51250   ieb.DisableUpdates False
51260 '  trv.Nodes.Add , , "Program", .OptionsTreeProgram
51270 '  trv.Nodes.Add "Program", tvwChild, "ProgramGeneral", .OptionsProgramGeneralSymbol
51280 '  trv.Nodes.Add "Program", tvwChild, "ProgramGhostscript", .OptionsProgramGhostscriptSymbol
51290 '  trv.Nodes.Add "Program", tvwChild, "ProgramDocument", .OptionsProgramDocumentSymbol
51300 '  trv.Nodes.Add "Program", tvwChild, "ProgramSave", .OptionsProgramSaveSymbol
51310 '  trv.Nodes.Add "Program", tvwChild, "ProgramAutosave", .OptionsProgramAutosaveSymbol
51320 '  trv.Nodes.Add "Program", tvwChild, "ProgramDirectories", .OptionsProgramDirectoriesSymbol
51330 '  trv.Nodes.Add "Program", tvwChild, "ProgramFonts", .OptionsProgramFontSymbol
51340 '  trv.Nodes.Add , , "Formats", .OptionsTreeFormats
51350 '  trv.Nodes.Add "Formats", tvwChild, "FormatsPDF", .OptionsPDFSymbol
51360 '  trv.Nodes.Add "Formats", tvwChild, "FormatsPNG", .OptionsPNGSymbol
51370 '  trv.Nodes.Add "Formats", tvwChild, "FormatsJPEG", .OptionsJPEGSymbol
51380 '  trv.Nodes.Add "Formats", tvwChild, "FormatsBMP", .OptionsBMPSymbol
51390 '  trv.Nodes.Add "Formats", tvwChild, "FormatsPCX", .OptionsPCXSymbol
51400 '  trv.Nodes.Add "Formats", tvwChild, "FormatsTIFF", .OptionsTIFFSymbol
51410 '  trv.Nodes.Add "Formats", tvwChild, "FormatsPS", .OptionsPSSymbol
51420 '  trv.Nodes.Add "Formats", tvwChild, "FormatsEPS", .OptionsEPSSymbol
51430 '
51440 '  trv.Nodes("ProgramFonts").EnsureVisible
51450 '  trv.Nodes("FormatsPDF").EnsureVisible
51460
51470   Set picOptions = LoadResPicture(2101, vbResIcon)
51480   dmFraProgGeneral.Visible = True
51490
51500   dmFraProgGeneral.Caption = .OptionsProgramGeneralSymbol
51510   dmFraShellIntegration.Caption = .OptionsShellIntegration
51520   dmFraProgGhostscript.Caption = .OptionsProgramGhostscriptSymbol
51530   dmFraProgAutosave.Caption = .OptionsProgramAutosaveSymbol
51540   dmFraProgDirectories.Caption = .OptionsProgramDirectoriesSymbol
51550   dmFraProgDocument.Caption = .OptionsProgramDocumentSymbol
51560   dmFraProgFont.Caption = .OptionsProgramFontSymbol
51570   dmfraProgSave.Caption = .OptionsProgramSaveSymbol
51580
51590   cmdShellintegration(0).Caption = .OptionsShellIntegrationAdd
51600   cmdShellintegration(1).Caption = .OptionsShellIntegrationRemove
51610   If IsWin9xMe = False Then
51620    If IsAdmin = False Then
51630     cmdShellintegration(0).Enabled = False
51640     cmdShellintegration(1).Enabled = False
51650    End If
51660   End If
51670
51680   lblGhostscriptversion.Caption = .OptionsGhostscriptversion
51690
51700   lblSaveFilename.Caption = .OptionsSaveFilename
51710   lblSaveFilenameTokens.Caption = .OptionsSaveFilenameTokens
51720   dmfraFilenameSubstitutions.Caption = .OptionsSaveFilenameSubstitutions
51730   chkFilenameSubst.Caption = .OptionsSaveFilenameSubstitutionsTitle
51740   cmdFilenameSubst(0).Caption = .OptionsSaveFilenameAdd
51750   cmdFilenameSubst(1).Caption = .OptionsSaveFilenameChange
51760   cmdFilenameSubst(2).Caption = .OptionsSaveFilenameDelete
51770
51780   chkSpaces.Caption = .OptionsRemoveSpaces
51790   chkNoConfirmMessageSwitchingDefaultprinter.Caption = .OptionsProgramSwitchingDefaultprinter
51800   lblGSbin.Caption = .OptionsDirectoriesGSBin
51810   lblGSlib.Caption = .OptionsDirectoriesGSLibraries
51820   lblGSfonts.Caption = .OptionsDirectoriesGSFonts
51830   lblPrintTempPath.Caption = .OptionsDirectoriesTempPath
51840
51850   lblOptions = .OptionsProgramGeneralDescription
51860   lblAutosaveformat.Caption = .OptionsAutosaveFormat
51870   chkUseStandardAuthor.Caption = .OptionsUseStandardauthor
51880   chkUseCreationDateNow.Caption = .OptionsUseCreationDateNow
51890   chkUseAutosave.Caption = .OptionsUseAutosave
51900   cmdTestpage.Caption = .OptionsPrintTestpage
51910   lblAutosaveFilename.Caption = .OptionsAutosaveFilename
51920   lblAutosaveFilenameTokens.Caption = .OptionsAutosaveFilenameTokens
51930   chkUseAutosaveDirectory.Caption = .OptionsUseAutosaveDirectory
51940   lblAuthorTokens.Caption = .OptionsStandardAuthorToken
51950
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
52060   With cmbSaveFilenameTokens
52070    .AddItem "<Author>"
52080    .AddItem "<Computername>"
52090    .AddItem "<DateTime>"
52100    .AddItem "<Title>"
52110    .AddItem "<Username>"
52120    .AddItem "<REDMON_DOCNAME>"
52130    .AddItem "<REDMON_JOB>"
52140    .AddItem "<REDMON_MACHINE>"
52150    .AddItem "<REDMON_PORT>"
52160    .AddItem "<REDMON_PRINTER>"
52170    .AddItem "<REDMON_SESSIONID>"
52180    .AddItem "<REDMON_USER>"
52190    .ListIndex = 0
52200   End With
52210   With cmbAuthorTokens
52220    .AddItem "<Computername>"
52230    .AddItem "<ClientComputer>"
52240    .AddItem "<DateTime>"
52250    .AddItem "<Title>"
52260    .AddItem "<Username>"
52270    .AddItem "<REDMON_DOCNAME>"
52280    .AddItem "<REDMON_JOB>"
52290    .AddItem "<REDMON_MACHINE>"
52300    .AddItem "<REDMON_PORT>"
52310    .AddItem "<REDMON_PRINTER>"
52320    .AddItem "<REDMON_SESSIONID>"
52330    .AddItem "<REDMON_USER>"
52340    .ListIndex = 0
52350   End With
52360   With cmbAutoSaveFilenameTokens
52370    .AddItem "<Author>"
52380    .AddItem "<Computername>"
52390    .AddItem "<ClientComputer>"
52400    .AddItem "<DateTime>"
52410    .AddItem "<Title>"
52420    .AddItem "<Username>"
52430    .AddItem "<REDMON_DOCNAME>"
52440    .AddItem "<REDMON_JOB>"
52450    .AddItem "<REDMON_MACHINE>"
52460    .AddItem "<REDMON_PORT>"
52470    .AddItem "<REDMON_PRINTER>"
52480    .AddItem "<REDMON_SESSIONID>"
52490    .AddItem "<REDMON_USER>"
52500    .ListIndex = 0
52510   End With
52520   Me.Caption = .DialogPrinterOptions
52530   cmdCancel.Caption = .OptionsCancel
52540   cmdReset.Caption = .OptionsReset
52550   cmdSave.Caption = .OptionsSave
52560   tbstrPDFOptions.Tabs.Clear
52570   tbstrPDFOptions.Tabs.Add , "General", .OptionsPDFGeneral
52580   tbstrPDFOptions.Tabs.Add , "Compression", .OptionsPDFCompression
52590   tbstrPDFOptions.Tabs.Add , "Fonts", .OptionsPDFFonts
52600   tbstrPDFOptions.Tabs.Add , "Colors", .OptionsPDFColors
52610   tbstrPDFOptions.Tabs.Add , "Security", .OptionsPDFSecurity
52620   dmFraPDFGeneral.Caption = .OptionsPDFGeneralCaption
52630   lblPDFCompat.Caption = .OptionsPDFGeneralCompatibility
52640   lblPDFAutoRotate.Caption = .OptionsPDFGeneralAutorotate
52650   lblPDFResolution.Caption = .OptionsPDFGeneralResolution
52660   lblPDFOverprint.Caption = .OptionsPDFGeneralOverprint
52670   lblProgfont.Caption = .OptionsProgramFont
52680   lblProgcharset.Caption = .OptionsProgramFontcharset
52690   lblSize.Caption = .OptionsProgramFontSize
52700   lblTesttext = .OptionsProgramFontTestdescription
52710   cmdTest.Caption = .OptionsProgramFontTest
52720   cmdCancelTest.Caption = .OptionsProgramFontCancelTest
52730   chkPDFASCII85.Caption = .OptionsPDFGeneralASCII85
52740   cmbPDFCompat.Clear
52750   cmbPDFCompat.AddItem .OptionsPDFGeneralCompatibility01
52760   cmbPDFCompat.AddItem .OptionsPDFGeneralCompatibility02
52770   cmbPDFCompat.AddItem .OptionsPDFGeneralCompatibility03
52780   cmbPDFRotate.Clear
52790   cmbPDFRotate.AddItem .OptionsPDFGeneralRotate01
52800   cmbPDFRotate.AddItem .OptionsPDFGeneralRotate02
52810   cmbPDFRotate.AddItem .OptionsPDFGeneralRotate03
52820   cmbPDFOverprint.Clear
52830   cmbPDFOverprint.AddItem .OptionsPDFGeneralOverprint01
52840   cmbPDFOverprint.AddItem .OptionsPDFGeneralOverprint02
52850
52860   dmfraPDFCompress.Caption = .OptionsPDFCompressionCaption
52870   chkPDFTextComp.Caption = .OptionsPDFCompressionTextComp
52880   dmFraPDFColor.Caption = .OptionsPDFCompressionColor
52890   chkPDFColorComp.Caption = .OptionsPDFCompressionColorComp
52900   chkPDFColorResample.Caption = .OptionsPDFCompressionColorResample
52910   lblPDFColorRes.Caption = .OptionsPDFCompressionColorRes
52920   cmbPDFColorComp.Clear
52930   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp01
52940   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp02
52950   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp03
52960   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp04
52970   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp05
52980   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp06
52990   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp07
53000   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp08
53010   cmbPDFColorResample.Clear
53020   cmbPDFColorResample.AddItem .OptionsPDFCompressionColorResample01
53030   cmbPDFColorResample.AddItem .OptionsPDFCompressionColorResample02
53040   cmbPDFColorResample.AddItem .OptionsPDFCompressionColorResample03
53050   dmFraPDFGrey.Caption = .OptionsPDFCompressionGrey
53060   chkPDFGreyComp.Caption = .OptionsPDFCompressionGreyComp
53070   chkPDFGreyResample.Caption = .OptionsPDFCompressionGreyResample
53080   lblPDFGreyRes.Caption = .OptionsPDFCompressionGreyRes
53090   cmbPDFGreyComp.Clear
53100   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp01
53110   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp02
53120   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp03
53130   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp04
53140   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp05
53150   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp06
53160   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp07
53170   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp08
53180   cmbPDFGreyResample.Clear
53190   cmbPDFGreyResample.AddItem .OptionsPDFCompressionGreyResample01
53200   cmbPDFGreyResample.AddItem .OptionsPDFCompressionGreyResample02
53210   cmbPDFGreyResample.AddItem .OptionsPDFCompressionGreyResample03
53220   dmFraPDFMono.Caption = .OptionsPDFCompressionMono
53230   chkPDFMonoComp.Caption = .OptionsPDFCompressionMonoComp
53240   chkPDFMonoResample.Caption = .OptionsPDFCompressionMonoResample
53250   lblPDFMonoRes.Caption = .OptionsPDFCompressionMonoRes
53260   cmbPDFMonoComp.Clear
53270   cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp01
53280   cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp02
53290   cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp03
53300   cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp04
53310   cmbPDFMonoResample.Clear
53320   cmbPDFMonoResample.AddItem .OptionsPDFCompressionMonoResample01
53330   cmbPDFMonoResample.AddItem .OptionsPDFCompressionMonoResample02
53340   cmbPDFMonoResample.AddItem .OptionsPDFCompressionMonoResample03
53350
53360   dmFraPDFFonts.Caption = .OptionsPDFFontsCaption
53370   chkPDFEmbedAll.Caption = .OptionsPDFFontsEmbedAll
53380   chkPDFSubSetFonts.Caption = .OptionsPDFFontsSubSetFonts
53390
53400   dmFraPDFColors.Caption = .OptionsPDFColorsCaption
53410   chkPDFCMYKtoRGB.Caption = .OptionsPDFColorsCMYKtoRGB
53420   dmFraPDFColorOptions.Caption = .OptionsPDFColorsColorOptions
53430   chkPDFPreserveOverprint.Caption = .OptionsPDFColorsPreserveOverprint
53440   chkPDFPreserveTransfer.Caption = .OptionsPDFColorsPreserveTransfer
53450   chkPDFPreserveHalftone.Caption = .OptionsPDFColorsPreserveHalftone
53460   cmbPDFColorModel.Clear
53470   cmbPDFColorModel.AddItem .OptionsPDFColorsColorModel01
53480   cmbPDFColorModel.AddItem .OptionsPDFColorsColorModel02
53490   cmbPDFColorModel.AddItem .OptionsPDFColorsColorModel03
53500
53510   dmFraPDFEncryptor.Caption = .OptionsPDFEncryptor
53520   dmFraPDFSecurity.Caption = .OptionsPDFSecurityCaption
53530   chkUseSecurity.Caption = .OptionsPDFUseSecurity
53540   dmFraPDFEncLevel.Caption = .OptionsPDFEncryptionLevel
53550   optEncHigh.Caption = .OptionsPDFEncryptionHigh
53560   optEncLow.Caption = .OptionsPDFEncryptionLow
53570   dmFraSecurityPass.Caption = .OptionsPDFPasswords
53580   chkUserPass.Caption = .OptionsPDFUserPass
53590   chkOwnerPass.Caption = .OptionsPDFOwnerPass
53600   dmFraPDFPermissions.Caption = .OptionsPDFDisallowUser
53610   dmFraPDFHighPermissions.Caption = .OptionsPDFEnhancedPermissions
53620   chkAllowPrinting.Caption = .OptionsPDFDisallowPrint
53630   chkAllowModifyContents.Caption = .OptionsPDFDisallowModify
53640   chkAllowCopy.Caption = .OptionsPDFDisallowCopy
53650   chkAllowModifyAnnotations.Caption = .OptionsPDFDisallowModifyComments
53660   chkAllowDegradedPrinting.Caption = .OptionsPDFAllowDegradedPrinting
53670   chkAllowFillIn.Caption = .OptionsPDFAllowFillIn
53680   chkAllowAssembly.Caption = .OptionsPDFAllowAssembly
53690   chkAllowScreenReaders.Caption = .OptionsPDFAllowScreenReaders
53700
53710   cmbPNGColors.AddItem .OptionsPNGColorscount01
53720   cmbPNGColors.AddItem .OptionsPNGColorscount02
53730   cmbPNGColors.AddItem .OptionsPNGColorscount03
53740   cmbPNGColors.AddItem .OptionsPNGColorscount04
53750   cmbJPEGColors.Left = cmbPNGColors.Left
53760   cmbJPEGColors.Width = cmbPNGColors.Width
53770   cmbJPEGColors.Top = cmbPNGColors.Top
53780   cmbJPEGColors.AddItem .OptionsJPEGColorscount01
53790   cmbJPEGColors.AddItem .OptionsJPEGColorscount02
53800   cmbBMPColors.Left = cmbPNGColors.Left
53810   cmbBMPColors.Width = cmbPNGColors.Width
53820   cmbBMPColors.Top = cmbPNGColors.Top
53830   cmbBMPColors.AddItem .OptionsBMPColorscount01
53840   cmbBMPColors.AddItem .OptionsBMPColorscount02
53850   cmbBMPColors.AddItem .OptionsBMPColorscount03
53860   cmbBMPColors.AddItem .OptionsBMPColorscount04
53870   cmbBMPColors.AddItem .OptionsBMPColorscount05
53880   cmbBMPColors.AddItem .OptionsBMPColorscount06
53890   cmbBMPColors.AddItem .OptionsBMPColorscount07
53900   cmbPCXColors.Left = cmbPNGColors.Left
53910   cmbPCXColors.Width = cmbPNGColors.Width
53920   cmbPCXColors.Top = cmbPNGColors.Top
53930   cmbPCXColors.AddItem .OptionsPCXColorscount01
53940   cmbPCXColors.AddItem .OptionsPCXColorscount02
53950   cmbPCXColors.AddItem .OptionsPCXColorscount03
53960   cmbPCXColors.AddItem .OptionsPCXColorscount04
53970   cmbPCXColors.AddItem .OptionsPCXColorscount05
53980   cmbPCXColors.AddItem .OptionsPCXColorscount06
53990   cmbTIFFColors.Left = cmbPNGColors.Left
54000   cmbTIFFColors.Width = cmbPNGColors.Width
54010   cmbTIFFColors.Top = cmbPNGColors.Top
54020   cmbTIFFColors.AddItem .OptionsTIFFColorscount01
54030   cmbTIFFColors.AddItem .OptionsTIFFColorscount02
54040   cmbTIFFColors.AddItem .OptionsTIFFColorscount03
54050   cmbTIFFColors.AddItem .OptionsTIFFColorscount04
54060   cmbTIFFColors.AddItem .OptionsTIFFColorscount05
54070   cmbTIFFColors.AddItem .OptionsTIFFColorscount06
54080   cmbTIFFColors.AddItem .OptionsTIFFColorscount07
54090   cmbTIFFColors.AddItem .OptionsTIFFColorscount08
54100
54110   dmFraBitmapGeneral.Caption = .OptionsImageSettings
54120   lblBitmapResolution = .OptionsBitmapResolution
54130   lblJPEGQuality = .OptionsJPEGQuality
54140   lblBitmapColors = .OptionsPDFColors
54150   lblProcessPriority.Caption = .OptionsProcesspriority
54160   lblLangLevel.Caption = .OptionsPSLanguageLevel
54170
54180   cmdAsso.Caption = .OptionsAssociatePSFiles
54190  End With
54200
54210  If IsPsAssociate = False Then
54220    cmdAsso.Enabled = True
54230   Else
54240    cmdAsso.Enabled = False
54250  End If
54260
54270  txtPDFRes.Text = 600
54280  cmbPDFCompat.ListIndex = 1
54290  cmbPDFRotate.ListIndex = 0
54300  cmbPDFOverprint.ListIndex = 0
54310  chkPDFASCII85.Value = 0
54320
54330  chkPDFTextComp.Value = 1
54340
54350  chkPDFColorComp.Value = 1
54360  chkPDFColorResample.Value = 0
54370  cmbPDFColorComp.ListIndex = 0
54380  cmbPDFColorResample.ListIndex = 0
54390  txtPDFColorRes.Text = 300
54400
54410  chkPDFGreyComp.Value = 1
54420  chkPDFGreyResample.Value = 0
54430  cmbPDFGreyComp.ListIndex = 0
54440  cmbPDFGreyResample.ListIndex = 0
54450  txtPDFGreyRes.Text = 300
54460
54470  chkPDFMonoComp.Value = 1
54480  chkPDFMonoResample.Value = 0
54490  cmbPDFMonoComp.ListIndex = 0
54500  cmbPDFMonoResample.ListIndex = 0
54510  txtPDFMonoRes.Text = 1200
54520
54530  chkPDFEmbedAll.Value = 1
54540  chkPDFSubSetFonts.Value = 1
54550  txtPDFSubSetPerc.Text = 100
54560
54570  cmbPDFColorModel.ListIndex = 1
54580  chkPDFCMYKtoRGB.Value = 1
54590  chkPDFPreserveOverprint.Value = 1
54600  chkPDFPreserveTransfer.Value = 1
54610  chkPDFPreserveHalftone.Value = 0
54620
54630  cmbPNGColors.ListIndex = 0
54640  cmbJPEGColors.ListIndex = 0
54650  cmbBMPColors.ListIndex = 0
54660  cmbPCXColors.ListIndex = 0
54670  cmbTIFFColors.ListIndex = 0
54680  txtBitmapResolution.Text = 150
54690
54700 ' chkUseStandardAuthor.Value = 1
54710  txtStandardAuthor.Text = vbNullString
54720
54730  With cmbPSLanguageLevel
54740   .AddItem "1"
54750   .AddItem "1.5"
54760   .AddItem "2"
54770   .AddItem "3"
54780  End With
54790  With cmbEPSLanguageLevel
54800   .AddItem "1"
54810   .AddItem "1.5"
54820   .AddItem "2"
54830   .AddItem "3"
54840  End With
54850
54860  With lsvFilenameSubst
54870   .Appearance = ccFlat
54880   .ColumnHeaders.Clear
54890   .ColumnHeaders.Add , "Str1", "", lsvFilenameSubst.Width / 2 - 140
54900   .ColumnHeaders.Add , "Str2", "", lsvFilenameSubst.Width / 2 - 140
54910   .HideColumnHeaders = True
54920   .GridLines = True
54930   .FullRowSelect = True
54940   .HideSelection = False
54950  End With
54960
54970  With cmbPDFEncryptor
54980   .Clear
54990   .AddItem "Ghostscript (>= 8.14)"
55000   .ItemData(.NewIndex) = 0
55010   .AddItem "PDFEnc"
55020   .ItemData(.NewIndex) = 1
55030
55040 '  ShowOptions Me, Options
55050
55060   SecurityIsPossible = True
55070
55080   If FileExists(CompletePath(App.Path) & "pdfenc.exe") = False Then
55090    .RemoveItem 1
55100    .ListIndex = 0
55110    Options.PDFEncryptor = .ItemData(.ListIndex)
55120   End If
55130   If GhostScriptSecurity = False Then
55140    .RemoveItem 0
55150   End If
55160   If .ListCount = 0 Then
55170     chkUseSecurity.Value = 0
55180     chkUseSecurity.Enabled = False
55190     SecurityIsPossible = False
55200    Else
55210     For i = 0 To .ListCount - 1
55220      If .ItemData(i) = Options.PDFEncryptor Then
55230       .ListIndex = i
55240       Exit For
55250      End If
55260     Next i
55270     If .ListIndex = -1 Then
55280      .ListIndex = 0
55290      Options.PDFEncryptor = .ItemData(.ListIndex)
55300     End If
55310   End If
55320  End With
55330
55340  If Options.PDFHighEncryption <> 0 Then
55350    optEncHigh.Value = True
55360   Else
55370    optEncLow.Value = True
55380  End If
55390
55400  cmdFilenameSubst(0).Top = lsvFilenameSubst.Top
55410  cmdFilenameSubst(1).Top = lsvFilenameSubst.Top + (lsvFilenameSubst.Height - cmdFilenameSubst(1).Height) / 2
55420  cmdFilenameSubst(2).Top = lsvFilenameSubst.Top + lsvFilenameSubst.Height - cmdFilenameSubst(2).Height
55430
55440  CheckCmdFilenameSubst
55450
55460  If chkUseStandardAuthor.Value = 1 Then
55470    txtStandardAuthor.Enabled = True
55480    txtStandardAuthor.BackColor = &H80000005
55490   Else
55500    txtStandardAuthor.Enabled = False
55510    txtStandardAuthor.BackColor = &H8000000F
55520  End If
55530  With Options
55540   SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
55550  End With
55560  ieb.Refresh
55570  If chkUseAutosave.Value = 1 Then
55580    ViewAutosave True
55590   Else
55600    ViewAutosave False
55610  End If
55620
55630  With txtGSbin
55640   .ToolTipText = .Text
55650  End With
55660  With txtGSlib
55670   .ToolTipText = .Text
55680  End With
55690  With txtGSfonts
55700   .ToolTipText = .Text
55710  End With
55720  With txtTemppath
55730   .ToolTipText = .Text
55740  End With
55750
55760  With sldProcessPriority
55770   .TextPosition = sldBelowRight
55780   .TickFrequency = 1
55790   .TickStyle = sldTopLeft
55801   Select Case .Value
         Case 0: 'Idle
55820     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityIdle
55830    Case 1: 'Normal
55840     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityNormal
55850    Case 2: 'High
55860     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityHigh
55870    Case 3: 'Realtime
55880     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityRealtime
55890   End Select
55900  End With
55910
55920  If IsWin9xMe = False Then
55930    lblProcessPriority.Enabled = True
55940    sldProcessPriority.Enabled = True
55950   Else
55960    lblProcessPriority.Enabled = False
55970    sldProcessPriority.Enabled = False
55980  End If
55990  UpdateSecurityFields
56000
56010  tStr2 = CompletePath(UCase$(Trim$(Options.DirectoryGhostscriptBinaries)))
 cmbGhostscript.Clear: Set reg = New clsRegistry
56030  reg.hkey = HKEY_LOCAL_MACHINE
56040
56050  Set gsvers = GetAllGhostscriptversions
56060
56070  If gsvers.Count = 0 Then
56080    cmbGhostscript.Enabled = False
56090   Else
56100    For i = 1 To gsvers.Count
56110     cmbGhostscript.AddItem gsvers.Item(i)
56120    Next i
56130    cmbGhostscript.ListIndex = cmbGhostscript.ListCount - 1
56140    For i = 0 To cmbGhostscript.ListCount - 1
56150     tstr = ""
56160     If InStr(cmbGhostscript.List(i), ":") Then
56170       reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
56180       If tStr2 = CompletePath(UCase$(Trim$(reg.GetRegistryValue("GhostscriptDirectoryBinaries")))) Then
56190        cmbGhostscript.ListIndex = i
56200        Exit For
56210       End If
56220      Else
56230       If InStr(UCase$(cmbGhostscript.List(i)), "AFPL") Then
56240        reg.KeyRoot = "SOFTWARE\AFPL Ghostscript"
56250        If InStr(cmbGhostscript.List(i), " ") > 0 Then
56260         tsf = Split(cmbGhostscript.List(i), " ")
56270         reg.Subkey = tsf(UBound(tsf))
56280         tstr = reg.GetRegistryValue("GS_DLL")
56290         If tStr2 & "GSDLL32.DLL" = UCase$(tstr) Then
56300          cmbGhostscript.ListIndex = i
56310          Exit For
56320         End If
56330        End If
56340       End If
56350       If InStr(UCase$(cmbGhostscript.List(i)), "GNU") Then
56360        reg.KeyRoot = "SOFTWARE\GNU Ghostscript"
56370        If InStr(cmbGhostscript.List(i), " ") > 0 Then
56380         tsf = Split(cmbGhostscript.List(i), " ")
56390         reg.Subkey = tsf(UBound(tsf))
56400         tstr = reg.GetRegistryValue("GS_DLL")
56410         If tStr2 & "GSDLL32.DLL" = UCase$(tstr) Then
56420          cmbGhostscript.ListIndex = i
56430          Exit For
56440         End If
56450        End If
56460       End If
56470       If InStr(UCase$(cmbGhostscript.List(i)), "GPL") Then
56480        reg.KeyRoot = "SOFTWARE\GPL Ghostscript"
56490        If InStr(cmbGhostscript.List(i), " ") > 0 Then
56500         tsf = Split(cmbGhostscript.List(i), " ")
56510         reg.Subkey = tsf(UBound(tsf))
56520         tstr = reg.GetRegistryValue("GS_DLL")
56530         If tStr2 & "GSDLL32.DLL" = UCase$(tstr) Then
56540          cmbGhostscript.ListIndex = i
56550          Exit For
56560         End If
56570        End If
56580       End If
56590     End If
56600    Next i
56610  End If
56620  Set reg = Nothing
56630  With cmbGhostscript
56640   If .ListCount = 0 Then
56650    .Enabled = False
56660    .BackColor = &H8000000F
56670   End If
56680  End With
56690
56700  tbstrPDFOptions.ZOrder 1
56710  'cmdStyle.ZOrder 1
56720  If ShowOnlyOptions = True Then
56730   FormInTaskbar Me, True, True
56740   Caption = "PDFCreator - " & Caption
56750  End If
56760  Timer1.Enabled = True
56770  Screen.MousePointer = vbNormal
56780 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
56790 Exit Sub
ErrPtnr_OnError:
56811 Select Case ErrPtnr.OnError("frmOptions", "Form_Load")
      Case 0: Resume
56830 Case 1: Resume Next
56840 Case 2: Exit Sub
56850 Case 3: End
56860 End Select
56870 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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

Private Sub ieb_ItemClick(sGroup As String, sItemKey As String)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim ctl As Control
50050  lblJPEGQuality.Visible = False
50060  cmbPNGColors.Visible = False
50070  cmbJPEGColors.Visible = False
50080  cmbBMPColors.Visible = False
50090  cmbPCXColors.Visible = False
50100  cmbTIFFColors.Visible = False
50110  tbstrPDFOptions.Visible = False
50120  For Each ctl In Controls
50130   If TypeOf ctl Is dmFrame Then
50140    ctl.Visible = False
50150    ctl.Enabled = False
50160   End If
50170  Next
50180  dmFraDescription.Visible = True
50190  dmFraDescription.Enabled = True
50200  tbstrPDFOptions.Enabled = False
50210  txtJPEGQuality.Visible = False
50220  lblJPEQQualityProzent.Visible = False
50230  dmFraPSGeneral.Visible = False
50240  cmbPSLanguageLevel.Visible = False
50250  cmbEPSLanguageLevel.Visible = False
50260
50271  Select Case UCase$(sGroup)
        Case "PROGRAM"
50291    Select Case UCase$(sItemKey)
          Case "GENERAL"
50310      Set picOptions = LoadResPicture(2101, vbResIcon)
50320      lblOptions = LanguageStrings.OptionsProgramGeneralDescription
50330      dmFraProgGeneral.Enabled = True
50340      dmFraShellIntegration.Enabled = True
50350      dmFraProgGeneral.Visible = True
50360      dmFraShellIntegration.Visible = True
50370      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50380     Case "GHOSTSCRIPT"
50390      Set picOptions = LoadResPicture(2119, vbResIcon)
50400      lblOptions = LanguageStrings.OptionsProgramGhostscriptDescription
50410      dmFraProgGhostscript.Enabled = True
50420      dmFraProgGhostscript.Visible = True
50430      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50440     Case "DOCUMENT"
50450      Set picOptions = LoadResPicture(2105, vbResIcon)
50460      lblOptions = LanguageStrings.OptionsProgramDocumentDescription
50470      dmFraProgDocument.Enabled = True
50480      dmFraProgDocument.Visible = True
50490      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50500     Case "SAVE"
50510      Set picOptions = LoadResPicture(2106, vbResIcon)
50520      lblOptions = LanguageStrings.OptionsProgramSaveDescription
50530      dmfraProgSave.Enabled = True
50540      dmfraProgSave.Visible = True
50550      dmfraFilenameSubstitutions.Visible = True
50560      dmfraFilenameSubstitutions.Enabled = True
50570      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50580     Case "AUTOSAVE"
50590      Set picOptions = LoadResPicture(2103, vbResIcon)
50600      lblOptions = LanguageStrings.OptionsProgramAutosaveDescription
50610      dmFraProgAutosave.Enabled = True
50620      dmFraProgAutosave.Visible = True
50630      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50640     Case "DIRECTORIES"
50650      Set picOptions = LoadResPicture(2104, vbResIcon)
50660      lblOptions = LanguageStrings.OptionsProgramDirectoriesDescription
50670      dmFraProgDirectories.Enabled = True
50680      dmFraProgDirectories.Visible = True
50690      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50700     Case "FONTS"
50710      Set picOptions = LoadResPicture(2102, vbResIcon)
50720      lblOptions = LanguageStrings.OptionsProgramFontDescription
50730      dmFraProgFont.Enabled = True
50740      dmFraProgFont.Visible = True
50750      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50760    End Select
50770   Case "FORMATS"
50781    Select Case UCase$(sItemKey)
          Case "PDF"
50800      Set picOptions = LoadResPicture(2111, vbResIcon)
50810      lblOptions = LanguageStrings.OptionsPDFDescription
50820      tbstrPDFOptions.Enabled = True
50830      tbstrPDFOptions.Visible = True
50840      dmFraPDFGeneral.Enabled = True
50850      dmFraPDFGeneral.Visible = True
50860      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
50870      dmFraPDFGeneral.Enabled = True
50880     Case "PNG"
50890      Set picOptions = LoadResPicture(2112, vbResIcon)
50900      lblOptions = LanguageStrings.OptionsPNGDescription
50910      dmFraBitmapGeneral.Enabled = True
50920      dmFraBitmapGeneral.Visible = True
50930      cmbPNGColors.Visible = True
50940      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
50950     Case "JPEG"
50960      Set picOptions = LoadResPicture(2113, vbResIcon)
50970      lblOptions = LanguageStrings.OptionsJPEGDescription
50980      dmFraBitmapGeneral.Enabled = True
50990      dmFraBitmapGeneral.Visible = True
51000      lblJPEGQuality.Caption = LanguageStrings.OptionsJPEGQuality
51010      lblJPEGQuality.Visible = True
51020      txtJPEGQuality.Visible = True
51030      lblJPEQQualityProzent.Visible = True
51040      lblJPEQQualityProzent.Left = txtJPEGQuality.Left + txtJPEGQuality.Width + 100
51050      cmbJPEGColors.Visible = True
51060      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51070     Case "BMP"
51080      Set picOptions = LoadResPicture(2114, vbResIcon)
51090      lblOptions = LanguageStrings.OptionsBMPDescription
51100      dmFraBitmapGeneral.Enabled = True
51110      dmFraBitmapGeneral.Visible = True
51120      cmbBMPColors.Visible = True
51130      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51140     Case "PCX"
51150      Set picOptions = LoadResPicture(2115, vbResIcon)
51160      lblOptions = LanguageStrings.OptionsPCXDescription
51170      dmFraBitmapGeneral.Enabled = True
51180      dmFraBitmapGeneral.Visible = True
51190      cmbPCXColors.Visible = True
51200      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51210     Case "TIFF"
51220      Set picOptions = LoadResPicture(2116, vbResIcon)
51230      lblOptions = LanguageStrings.OptionsTIFFDescription
51240      dmFraBitmapGeneral.Enabled = True
51250      dmFraBitmapGeneral.Visible = True
51260      cmbTIFFColors.Visible = True
51270      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51280     Case "PS"
51290      Set picOptions = LoadResPicture(2117, vbResIcon)
51300      lblOptions.Caption = LanguageStrings.OptionsPSDescription
51310      dmFraPSGeneral.Enabled = True
51320      dmFraPSGeneral.Visible = True
51330      cmbPSLanguageLevel.Visible = True
51340      dmFraPSGeneral.Caption = LanguageStrings.OptionsPSDescription
51350      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51360     Case "EPS"
51370      Set picOptions = LoadResPicture(2118, vbResIcon)
51380      lblOptions.Caption = LanguageStrings.OptionsEPSDescription
51390      dmFraPSGeneral.Enabled = True
51400      dmFraPSGeneral.Visible = True
51410      cmbEPSLanguageLevel.Visible = True
51420      dmFraPSGeneral.Caption = LanguageStrings.OptionsEPSDescription
51430      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51440    End Select
51450  End Select
51460 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
51470 Exit Sub
ErrPtnr_OnError:
51491 Select Case ErrPtnr.OnError("frmOptions", "ieb_ItemClick")
      Case 0: Resume
51510 Case 1: Resume Next
51520 Case 2: Exit Sub
51530 Case 3: End
51540 End Select
51550 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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
50040  dmFraPDFGeneral.Visible = False
50050  dmfraPDFCompress.Visible = False
50060  dmFraPDFFonts.Visible = False
50070  dmFraPDFColors.Visible = False
50080  dmFraPDFColorOptions.Visible = False
50090  dmFraPDFSecurity.Visible = False
50100  dmFraPDFGeneral.Enabled = False
50110  dmfraPDFCompress.Enabled = False
50120  dmFraPDFFonts.Enabled = False
50130  dmFraPDFColors.Enabled = False
50140  dmFraPDFColorOptions.Enabled = False
50150  dmFraPDFSecurity.Enabled = False
50161  Select Case tbstrPDFOptions.SelectedItem.Index
        Case 1:
50180    dmFraPDFGeneral.Visible = True
50190    dmFraPDFGeneral.Enabled = True
50200   Case 2:
50210    dmfraPDFCompress.Visible = True
50220    dmfraPDFCompress.Enabled = True
50230    dmFraPDFColor.Visible = True
50240    dmFraPDFColor.Enabled = True
50250    dmFraPDFGrey.Visible = True
50260    dmFraPDFGrey.Enabled = True
50270    dmFraPDFMono.Visible = True
50280    dmFraPDFMono.Enabled = True
50290   Case 3:
50300    dmFraPDFFonts.Visible = True
50310    dmFraPDFFonts.Enabled = True
50320   Case 4:
50330    dmFraPDFColors.Visible = True
50340    dmFraPDFColorOptions.Visible = True
50350    dmFraPDFColors.Enabled = True
50360    dmFraPDFColorOptions.Enabled = True
50370   Case 5:
50380    dmFraPDFSecurity.Visible = True
50390    dmFraPDFSecurity.Enabled = True
50400    dmFraPDFEncryptor.Visible = True
50410    dmFraPDFEncryptor.Enabled = True
50420    dmFraPDFEncLevel.Visible = True
50430    dmFraPDFEncLevel.Enabled = True
50440    dmFraSecurityPass.Visible = True
50450    dmFraSecurityPass.Enabled = True
50460    dmFraPDFPermissions.Visible = True
50470    dmFraPDFPermissions.Enabled = True
50480    dmFraPDFHighPermissions.Visible = True
50490    dmFraPDFHighPermissions.Enabled = True
50500    If SecurityIsPossible = False Then
50510     MsgBox LanguageStrings.MessagesMsg19
50520    End If
50530  End Select
50540 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50550 Exit Sub
ErrPtnr_OnError:
50571 Select Case ErrPtnr.OnError("frmOptions", "tbstrPDFOptions_Click")
      Case 0: Resume
50590 Case 1: Resume Next
50600 Case 2: Exit Sub
50610 Case 3: End
50620 End Select
50630 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Timer1_Timer()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim i As Long, fi As Long, tstr As String, SMF As Collection, _
  cSystem As clsSystem, ctl As Control
50060  Timer1.Enabled = False
50070  Set cSystem = New clsSystem
50080  Set SMF = cSystem.GetSystemFont(Me, Menu)
50090  txtTest.Text = vbNullString
50100  For i = 33 To 255
50110   txtTest.Text = txtTest.Text & Chr$(i)
50120  Next i
50130  fi = -1
50140  With cmbFonts
50150   .Clear
50160   For i = 1 To Screen.FontCount
50170    tstr = Trim$(Screen.Fonts(i))
50180    If Len(tstr) > 0 Then
50190     cmbFonts.AddItem tstr
50200    End If
50210   Next i
50220   If .ListCount > 0 Then
50230     For i = 0 To cmbFonts.ListCount - 1
50240      If SMF.Count > 0 Then
50250       If UCase$(cmbFonts.List(i)) = UCase$(SMF(1)(0)) Then
50260        fi = i
50270       End If
50280      End If
50290     Next i
50300    Else
50310    .ListIndex = 0
50320   End If
50330  End With
50340  With cmbCharset
50350   .Clear
50360   .AddItem "0, Western": .ItemData(.NewIndex) = 0
50370   .AddItem "2, Symbol": .ItemData(.NewIndex) = 2
50380   .AddItem "77, Mac": .ItemData(.NewIndex) = 77
50390   .AddItem "161, Greek": .ItemData(.NewIndex) = 161
50400   .AddItem "162, Turkish": .ItemData(.NewIndex) = 162
50410   .AddItem "177, Hebrew": .ItemData(.NewIndex) = 177
50420   .AddItem "178, Arabic": .ItemData(.NewIndex) = 178
50430   .AddItem "186, Baltic": .ItemData(.NewIndex) = 186
50440   .AddItem "204, Cyrillic": .ItemData(.NewIndex) = 204
50450   .AddItem "238, Central European": .ItemData(.NewIndex) = 238
50460   .AddItem "255, DOS/OEM": .ItemData(.NewIndex) = 255
50470   .Text = 0
50480  End With
50490  With cmbProgramFontsize
50500   .AddItem "8"
50510   .AddItem "9"
50520   .AddItem "10"
50530   .AddItem "11"
50540   .AddItem "12"
50550   .AddItem "14"
50560   .AddItem "16"
50570   .AddItem "18"
50580   .AddItem "20"
50590   .AddItem "22"
50600   .AddItem "24"
50610   .AddItem "26"
50620   .AddItem "28"
50630   .AddItem "36"
50640   .AddItem "48"
50650   .AddItem "72"
50660  End With
50670  cmbProgramFontsize.Text = 8
50680  cmbCharset.Text = cmbCharset.ItemData(0)
50690  cmbCharset.Text = Options.ProgramFontCharset
50700  For Each ctl In Controls
50710   If TypeOf ctl Is ComboBox Then
50720    ComboSetListWidth ctl
50730   End If
50740  Next ctl
50750
50760  SetOptimalComboboxHeigth cmbCharset, Me
50770  SetOptimalComboboxHeigth cmbProgramFontsize, Me
50780  SetOptimalComboboxHeigth cmbGhostscript, Me
50790
50800  Form_Resize
50810
50820  cmbProgramFontsize.Width = txtTest.Width - _
  (cmbProgramFontsize.Left - txtTest.Left)
50840
50850  If fi >= 0 Then
50860   cmbFonts.ListIndex = fi
50870   cmbCharset.Text = SMF(1)(2)
50880   cmbProgramFontsize.Text = SMF(1)(1)
50890   txtTest.Font = cmbFonts.List(cmbFonts.ListIndex)
50900   txtTest.Font.Charset = cmbCharset.Text
50910  End If
50920
50930  ShowOptions Me, Options
50940
50950  CorrectCmbCharset
50960  Call ieb_ItemClick("PROGRAM", "GENERAL")
50970 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50980 Exit Sub
ErrPtnr_OnError:
51001 Select Case ErrPtnr.OnError("frmOptions", "Timer1_Timer")
      Case 0: Resume
51020 Case 1: Resume Next
51030 Case 2: Exit Sub
51040 Case 3: End
51050 End Select
51060 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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
50040  Dim allow As String, tstr As String
50050
50060  allow = "0123456789" & Chr$(8) & Chr$(13)
50070
50080  tstr = Chr$(KeyAscii)
50090
50100  If InStr(1, allow, tstr) = 0 Then
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
50040  If cmbPDFCompat.ListIndex < 2 Then
50050   optEncLow.Value = True
50060  End If
50070  If chkUseSecurity.Value = False Then
50080    dmFraPDFEncryptor.Enabled = False
50090    cmbPDFEncryptor.Enabled = False
50100
50110    dmFraPDFEncLevel.Enabled = False
50120    optEncHigh.Enabled = False
50130    optEncLow.Enabled = False
50140
50150    dmFraSecurityPass.Enabled = False
50160    chkUserPass.Enabled = False
50170    chkOwnerPass.Enabled = False
50180
50190    dmFraPDFPermissions.Enabled = False
50200    chkAllowPrinting.Enabled = False
50210    chkAllowCopy.Enabled = False
50220    chkAllowModifyAnnotations.Enabled = False
50230    chkAllowModifyContents.Enabled = False
50240
50250    dmFraPDFHighPermissions.Enabled = False
50260    chkAllowDegradedPrinting.Enabled = False
50270    chkAllowFillIn.Enabled = False
50280    chkAllowScreenReaders.Enabled = False
50290    chkAllowAssembly.Enabled = False
50300   Else
50310    dmFraPDFEncryptor.Enabled = True
50320    cmbPDFEncryptor.Enabled = True
50330
50340    dmFraPDFEncLevel.Enabled = True
50350    If cmbPDFCompat.ListIndex >= 2 Then
50360      optEncHigh.Enabled = True
50370     Else
50380      optEncHigh.Enabled = False
50390    End If
50400    optEncLow.Enabled = True
50410
50420    dmFraSecurityPass.Enabled = True
50430    chkUserPass.Enabled = True
50440    chkOwnerPass.Enabled = True
50450
50460    dmFraPDFPermissions.Enabled = True
50470    chkAllowPrinting.Enabled = True
50480    chkAllowCopy.Enabled = True
50490    chkAllowModifyAnnotations.Enabled = True
50500    chkAllowModifyContents.Enabled = True
50510
50520    If optEncHigh.Value = True Then
50530      dmFraPDFHighPermissions.Enabled = True
50540      chkAllowDegradedPrinting.Enabled = True
50550      chkAllowFillIn.Enabled = True
50560      chkAllowScreenReaders.Enabled = True
50570      chkAllowAssembly.Enabled = True
50580     Else
50590      dmFraPDFHighPermissions.Enabled = False
50600      chkAllowDegradedPrinting.Enabled = False
50610      chkAllowFillIn.Enabled = False
50620      chkAllowScreenReaders.Enabled = False
50630      chkAllowAssembly.Enabled = False
50640    End If
50650  End If
50660  If chkOwnerPass.Value = 0 And chkUserPass.Value = 0 Then
50670   chkOwnerPass.Value = 1: Options.PDFOwnerPass = 1
50680  End If
50690 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50700 Exit Sub
ErrPtnr_OnError:
50721 Select Case ErrPtnr.OnError("frmOptions", "UpdateSecurityFields")
      Case 0: Resume
50740 Case 1: Resume Next
50750 Case 2: Exit Sub
50760 Case 3: End
50770 End Select
50780 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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
50120    cmdFilenameSubstMove(0).Enabled = True
50130    cmdFilenameSubstMove(1).Enabled = True
50140   Else
50150    cmdFilenameSubstMove(0).Enabled = False
50160    cmdFilenameSubstMove(1).Enabled = False
50170  End If
50180  If lsvFilenameSubst.ListItems.Count > 0 Then
50190   If lsvFilenameSubst.SelectedItem.Index = 1 Then
50200    cmdFilenameSubstMove(0).Enabled = False
50210   End If
50220   If lsvFilenameSubst.SelectedItem.Index = lsvFilenameSubst.ListItems.Count Then
50230    cmdFilenameSubstMove(1).Enabled = False
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
