VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl ctlOptFormatPDF 
   AutoRedraw      =   -1  'True
   ClientHeight    =   10140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19335
   ScaleHeight     =   10140
   ScaleWidth      =   19335
   ToolboxBitmap   =   "ctlOptFormatPDF.ctx":0000
   Begin PDFCreator.dmFrame dmFraPDFSigning 
      Height          =   5535
      Left            =   12960
      TabIndex        =   68
      Top             =   120
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   9763
      Caption         =   "Signing"
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
      Begin VB.CheckBox chkSignPDF 
         Caption         =   "Sign pdf file"
         Height          =   255
         Left            =   120
         TabIndex        =   90
         Top             =   360
         Width           =   5895
      End
      Begin VB.CheckBox chkMultiSignature 
         Caption         =   "Multi signature allowed"
         Height          =   255
         Left            =   120
         TabIndex        =   89
         Top             =   5160
         Width           =   5895
      End
      Begin VB.TextBox txtSignatureLocation 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   120
         TabIndex        =   78
         Top             =   2760
         Width           =   5325
      End
      Begin VB.TextBox txtSignatureContact 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   3120
         TabIndex        =   76
         Top             =   2040
         Width           =   2805
      End
      Begin VB.TextBox txtSignatureReason 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   120
         TabIndex        =   74
         Top             =   2040
         Width           =   2805
      End
      Begin VB.TextBox txtPFXFilePreview 
         Appearance      =   0  '2D
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   72
         Top             =   1320
         Width           =   5910
      End
      Begin VB.CommandButton cmdGetPFXFile 
         Caption         =   "..."
         Height          =   300
         Left            =   5640
         TabIndex        =   70
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtPFXfile 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   120
         TabIndex        =   69
         Top             =   960
         Width           =   5325
      End
      Begin PDFCreator.dmFrame dmFraSignaturePosition 
         Height          =   1935
         Left            =   120
         TabIndex        =   79
         Top             =   3120
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   3413
         Caption         =   "Signature position"
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
         Begin VB.TextBox txtSignatureOnPage 
            Alignment       =   1  'Rechts
            Appearance      =   0  '2D
            Height          =   285
            Left            =   240
            TabIndex        =   94
            Text            =   "1"
            Top             =   960
            Width           =   1000
         End
         Begin VB.CheckBox chkSignatureVisible 
            Caption         =   "Signature visible in PDF"
            Height          =   255
            Left            =   240
            TabIndex        =   88
            Top             =   360
            Width           =   5415
         End
         Begin VB.TextBox txtRightY 
            Alignment       =   1  'Rechts
            Appearance      =   0  '2D
            Height          =   285
            Left            =   4560
            TabIndex        =   87
            Text            =   "200"
            Top             =   1560
            Width           =   1000
         End
         Begin VB.TextBox txtRightX 
            Alignment       =   1  'Rechts
            Appearance      =   0  '2D
            Height          =   285
            Left            =   3120
            TabIndex        =   85
            Text            =   "200"
            Top             =   1560
            Width           =   1000
         End
         Begin VB.TextBox txtLeftY 
            Alignment       =   1  'Rechts
            Appearance      =   0  '2D
            Height          =   285
            Left            =   1680
            TabIndex        =   83
            Text            =   "100"
            Top             =   1560
            Width           =   1000
         End
         Begin VB.TextBox txtLeftX 
            Alignment       =   1  'Rechts
            Appearance      =   0  '2D
            Height          =   285
            Left            =   240
            TabIndex        =   81
            Text            =   "100"
            Top             =   1560
            Width           =   1000
         End
         Begin VB.Label lblSignatureOnPage 
            AutoSize        =   -1  'True
            Caption         =   "Show signature on page"
            Height          =   195
            Left            =   240
            TabIndex        =   93
            Top             =   720
            Width           =   1725
         End
         Begin VB.Label lblRightY 
            AutoSize        =   -1  'True
            Caption         =   "Right Y"
            Height          =   195
            Left            =   4560
            TabIndex        =   86
            Top             =   1320
            Width           =   525
         End
         Begin VB.Label lblRightX 
            AutoSize        =   -1  'True
            Caption         =   "Right X"
            Height          =   195
            Left            =   3120
            TabIndex        =   84
            Top             =   1320
            Width           =   525
         End
         Begin VB.Label lblLeftY 
            AutoSize        =   -1  'True
            Caption         =   "Left Y"
            Height          =   195
            Left            =   1680
            TabIndex        =   82
            Top             =   1320
            Width           =   420
         End
         Begin VB.Label lblLeftX 
            AutoSize        =   -1  'True
            Caption         =   "Left X"
            Height          =   195
            Left            =   240
            TabIndex        =   80
            Top             =   1320
            Width           =   420
         End
      End
      Begin VB.Label lblSignatureLocation 
         AutoSize        =   -1  'True
         Caption         =   "Signature location"
         Height          =   195
         Left            =   120
         TabIndex        =   77
         Top             =   2520
         Width           =   1275
      End
      Begin VB.Label lblSignatureContact 
         AutoSize        =   -1  'True
         Caption         =   "Signature contact"
         Height          =   195
         Left            =   3120
         TabIndex        =   75
         Top             =   1800
         Width           =   1260
      End
      Begin VB.Label lblSignatureReason 
         AutoSize        =   -1  'True
         Caption         =   "Signature reason"
         Height          =   195
         Left            =   120
         TabIndex        =   73
         Top             =   1800
         Width           =   1200
      End
      Begin VB.Label lblPFXFile 
         AutoSize        =   -1  'True
         Caption         =   "PFX\P12 file"
         Height          =   195
         Left            =   120
         TabIndex        =   71
         Top             =   720
         Width           =   900
      End
   End
   Begin PDFCreator.dmFrame dmFraPDFColorOptions 
      Height          =   1455
      Left            =   120
      TabIndex        =   44
      Top             =   8640
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   2566
      Caption         =   "Options"
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
      Begin VB.CheckBox chkPDFPreserveOverprint 
         Appearance      =   0  '2D
         Caption         =   "Preserve Overprint Settings"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   360
         Width           =   5910
      End
      Begin VB.CheckBox chkPDFPreserveTransfer 
         Appearance      =   0  '2D
         Caption         =   "Preserve Transfer Functions"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Tag             =   "Remove|Preserve"
         Top             =   720
         Width           =   5910
      End
      Begin VB.CheckBox chkPDFPreserveHalftone 
         Appearance      =   0  '2D
         Caption         =   "Preserve Halftone Information"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   1050
         Width           =   5910
      End
   End
   Begin PDFCreator.dmFrame dmFraPDFColors 
      Height          =   1215
      Left            =   120
      TabIndex        =   41
      Top             =   7320
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   2143
      Caption         =   "Color options"
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
      Begin VB.CheckBox chkPDFCMYKtoRGB 
         Appearance      =   0  '2D
         Caption         =   "Convert CMYK Images to RGB"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   840
         Width           =   5880
      End
      Begin VB.ComboBox cmbPDFColorModel 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "ctlOptFormatPDF.ctx":0312
         Left            =   120
         List            =   "ctlOptFormatPDF.ctx":0314
         Style           =   2  'Dropdown-Liste
         TabIndex        =   42
         Tag             =   "RGB|CMYK|GRAY"
         Top             =   360
         Width           =   2775
      End
   End
   Begin PDFCreator.dmFrame dmFraPDFSecurity 
      Height          =   5535
      Left            =   6600
      TabIndex        =   48
      Top             =   4560
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   9763
      Caption         =   "Security"
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
      Begin VB.CheckBox chkUseSecurity 
         Appearance      =   0  '2D
         Caption         =   "Use Security"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   360
         Width           =   5535
      End
      Begin PDFCreator.dmFrame dmFraPDFHighPermissions 
         Height          =   855
         Left            =   120
         TabIndex        =   63
         Top             =   4560
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   1508
         Caption         =   "Enhanced permissions (128 Bit only)"
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
         Begin VB.CheckBox chkAllowDegradedPrinting 
            Appearance      =   0  '2D
            Caption         =   "Allow printing in low resolution"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   300
            Width           =   2865
         End
         Begin VB.CheckBox chkAllowFillIn 
            Appearance      =   0  '2D
            Caption         =   "Allow filling in form fields"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3100
            TabIndex        =   65
            Top             =   300
            Width           =   2760
         End
         Begin VB.CheckBox chkAllowScreenReaders 
            Appearance      =   0  '2D
            Caption         =   "Allow Screen Readers"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   525
            Width           =   2865
         End
         Begin VB.CheckBox chkAllowAssembly 
            Appearance      =   0  '2D
            Caption         =   "Allow changes to the Assembly"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3100
            TabIndex        =   67
            Top             =   525
            Width           =   2760
         End
      End
      Begin PDFCreator.dmFrame dmFraPDFPermissions 
         Height          =   855
         Left            =   120
         TabIndex        =   58
         Top             =   3600
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   1508
         Caption         =   "Disallow user to"
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
         Begin VB.CheckBox chkAllowPrinting 
            Appearance      =   0  '2D
            Caption         =   "print the document"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   300
            Width           =   2865
         End
         Begin VB.CheckBox chkAllowCopy 
            Appearance      =   0  '2D
            Caption         =   "copy text and images"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   525
            Width           =   2865
         End
         Begin VB.CheckBox chkAllowModifyContents 
            Appearance      =   0  '2D
            Caption         =   "modify the document"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3100
            TabIndex        =   60
            Top             =   300
            Width           =   2760
         End
         Begin VB.CheckBox chkAllowModifyAnnotations 
            Appearance      =   0  '2D
            Caption         =   "modify comments"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3100
            TabIndex        =   62
            Top             =   525
            Width           =   2760
         End
      End
      Begin PDFCreator.dmFrame dmFraSecurityPass 
         Height          =   855
         Left            =   120
         TabIndex        =   55
         Top             =   2640
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   1508
         Caption         =   "Passwords"
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
         Begin VB.CheckBox chkUserPass 
            Appearance      =   0  '2D
            Caption         =   "Password required to open document"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   300
            Width           =   5700
         End
         Begin VB.CheckBox chkOwnerPass 
            Appearance      =   0  '2D
            Caption         =   "Password required to change Permissions and Passwords"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   525
            Width           =   5700
         End
      End
      Begin PDFCreator.dmFrame dmFraPDFEncLevel 
         Height          =   855
         Left            =   120
         TabIndex        =   52
         Top             =   1680
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   1508
         Caption         =   "Encryption level"
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
         Begin VB.OptionButton optEncLow 
            Appearance      =   0  '2D
            Caption         =   "Low (40 Bit - Adobe Acrobat 3.0 and above)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   300
            Width           =   5775
         End
         Begin VB.OptionButton optEncHigh 
            Appearance      =   0  '2D
            Caption         =   "High (128 Bit - Adobe Acrobat 5.0 and above)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   525
            Width           =   5775
         End
      End
      Begin PDFCreator.dmFrame dmFraPDFEncryptor 
         Height          =   855
         Left            =   120
         TabIndex        =   50
         Top             =   720
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   1508
         Caption         =   "Encryptor"
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
         Begin VB.ComboBox cmbPDFEncryptor 
            Appearance      =   0  '2D
            Height          =   315
            ItemData        =   "ctlOptFormatPDF.ctx":0316
            Left            =   120
            List            =   "ctlOptFormatPDF.ctx":0318
            Style           =   2  'Dropdown-Liste
            TabIndex        =   51
            Top             =   360
            Width           =   5715
         End
      End
   End
   Begin PDFCreator.dmFrame dmFraPDFFonts 
      Height          =   1695
      Left            =   120
      TabIndex        =   36
      Top             =   5520
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   2990
      Caption         =   "Font options"
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
      Begin VB.TextBox txtPDFSubSetPerc 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   400
         TabIndex        =   39
         Top             =   1320
         Width           =   495
      End
      Begin VB.CheckBox chkPDFSubSetFonts 
         Appearance      =   0  '2D
         Caption         =   "Subset Fonts, when percentage of used characters below:"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   38
         Top             =   780
         Width           =   5955
      End
      Begin VB.CheckBox chkPDFEmbedAll 
         Appearance      =   0  '2D
         Caption         =   "Embed all Fonts"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   5955
      End
      Begin VB.Label lblPDFPerc 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   960
         TabIndex        =   40
         Top             =   1365
         Width           =   120
      End
   End
   Begin PDFCreator.dmFrame dmfraPDFCompress 
      Height          =   4335
      Left            =   6600
      TabIndex        =   13
      Top             =   120
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   7646
      Caption         =   "Compression"
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
      Begin VB.CheckBox chkPDFTextComp 
         Appearance      =   0  '2D
         Caption         =   "Compress Text Objects"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   5910
      End
      Begin PDFCreator.dmFrame dmFraPDFMono 
         Height          =   1095
         Left            =   120
         TabIndex        =   29
         Top             =   3120
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   1931
         Caption         =   "Monochrome images"
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
         Begin VB.TextBox txtPDFMonoRes 
            Appearance      =   0  '2D
            Height          =   285
            Left            =   4935
            TabIndex        =   35
            Top             =   660
            Width           =   735
         End
         Begin VB.ComboBox cmbPDFMonoResample 
            Appearance      =   0  '2D
            Height          =   315
            ItemData        =   "ctlOptFormatPDF.ctx":031A
            Left            =   2520
            List            =   "ctlOptFormatPDF.ctx":031C
            Style           =   2  'Dropdown-Liste
            TabIndex        =   34
            Tag             =   "Bicubic|Subsample|Average"
            Top             =   660
            Width           =   2370
         End
         Begin VB.CheckBox chkPDFMonoResample 
            Appearance      =   0  '2D
            Caption         =   "Resample"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2520
            TabIndex        =   31
            Top             =   360
            Width           =   2295
         End
         Begin VB.ComboBox cmbPDFMonoComp 
            Appearance      =   0  '2D
            Height          =   315
            ItemData        =   "ctlOptFormatPDF.ctx":031E
            Left            =   120
            List            =   "ctlOptFormatPDF.ctx":0320
            Style           =   2  'Dropdown-Liste
            TabIndex        =   33
            Top             =   660
            Width           =   2370
         End
         Begin VB.CheckBox chkPDFMonoComp 
            Appearance      =   0  '2D
            Caption         =   "Compress"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   2325
         End
         Begin VB.Label lblPDFMonoRes 
            AutoSize        =   -1  'True
            Caption         =   "Resolution"
            Height          =   195
            Left            =   4935
            TabIndex        =   32
            Top             =   360
            Width           =   750
         End
      End
      Begin PDFCreator.dmFrame dmFraPDFGrey 
         Height          =   1095
         Left            =   120
         TabIndex        =   22
         Top             =   1920
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   1931
         Caption         =   "Greyscale images"
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
         Begin VB.CheckBox chkPDFGreyComp 
            Appearance      =   0  '2D
            Caption         =   "Compress"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   2325
         End
         Begin VB.ComboBox cmbPDFGreyComp 
            Appearance      =   0  '2D
            Height          =   315
            ItemData        =   "ctlOptFormatPDF.ctx":0322
            Left            =   120
            List            =   "ctlOptFormatPDF.ctx":0324
            Style           =   2  'Dropdown-Liste
            TabIndex        =   26
            Top             =   660
            Width           =   2370
         End
         Begin VB.CheckBox chkPDFGreyResample 
            Appearance      =   0  '2D
            Caption         =   "Resample"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2520
            TabIndex        =   24
            Top             =   360
            Width           =   2295
         End
         Begin VB.ComboBox cmbPDFGreyResample 
            Appearance      =   0  '2D
            Height          =   315
            ItemData        =   "ctlOptFormatPDF.ctx":0326
            Left            =   2520
            List            =   "ctlOptFormatPDF.ctx":0328
            Style           =   2  'Dropdown-Liste
            TabIndex        =   27
            Tag             =   "Bicubic|Subsample|Average"
            Top             =   660
            Width           =   2370
         End
         Begin VB.TextBox txtPDFGreyRes 
            Appearance      =   0  '2D
            Height          =   285
            Left            =   4935
            TabIndex        =   28
            Top             =   660
            Width           =   735
         End
         Begin VB.Label lblPDFGreyRes 
            AutoSize        =   -1  'True
            Caption         =   "Resolution"
            Height          =   195
            Left            =   4935
            TabIndex        =   25
            Top             =   360
            Width           =   750
         End
      End
      Begin PDFCreator.dmFrame dmFraPDFColor 
         Height          =   1095
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   1931
         Caption         =   "Color images"
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
         Begin VB.CheckBox chkPDFColorComp 
            Appearance      =   0  '2D
            Caption         =   "Compress"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   2325
         End
         Begin VB.ComboBox cmbPDFColorComp 
            Appearance      =   0  '2D
            Height          =   315
            ItemData        =   "ctlOptFormatPDF.ctx":032A
            Left            =   120
            List            =   "ctlOptFormatPDF.ctx":032C
            Style           =   2  'Dropdown-Liste
            TabIndex        =   19
            Top             =   660
            Width           =   2370
         End
         Begin VB.CheckBox chkPDFColorResample 
            Appearance      =   0  '2D
            Caption         =   "Resample"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2520
            TabIndex        =   17
            Top             =   360
            Width           =   2295
         End
         Begin VB.ComboBox cmbPDFColorResample 
            Appearance      =   0  '2D
            Height          =   315
            ItemData        =   "ctlOptFormatPDF.ctx":032E
            Left            =   2520
            List            =   "ctlOptFormatPDF.ctx":0330
            Style           =   2  'Dropdown-Liste
            TabIndex        =   20
            Tag             =   "Bicubic|Subsample|Average"
            Top             =   660
            Width           =   2370
         End
         Begin VB.TextBox txtPDFColorRes 
            Appearance      =   0  '2D
            Height          =   285
            Left            =   4935
            TabIndex        =   21
            Top             =   660
            Width           =   735
         End
         Begin VB.Label lblPDFColorRes 
            AutoSize        =   -1  'True
            Caption         =   "Resolution"
            Height          =   195
            Left            =   4935
            TabIndex        =   18
            Top             =   360
            Width           =   750
         End
      End
   End
   Begin PDFCreator.dmFrame dmFraPDFGeneral 
      Height          =   4845
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   8546
      Caption         =   "General Options"
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
      Begin VB.ComboBox cmbPDFDefaultSettings 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "ctlOptFormatPDF.ctx":0332
         Left            =   120
         List            =   "ctlOptFormatPDF.ctx":0334
         Style           =   2  'Dropdown-Liste
         TabIndex        =   91
         Top             =   555
         Width           =   2655
      End
      Begin VB.ComboBox cmbPDFRotate 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "ctlOptFormatPDF.ctx":0336
         Left            =   120
         List            =   "ctlOptFormatPDF.ctx":0338
         Style           =   2  'Dropdown-Liste
         TabIndex        =   5
         Tag             =   "None|All|PageByPage"
         Top             =   1950
         Width           =   2655
      End
      Begin VB.ComboBox cmbPDFCompat 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "ctlOptFormatPDF.ctx":033A
         Left            =   120
         List            =   "ctlOptFormatPDF.ctx":033C
         Style           =   2  'Dropdown-Liste
         TabIndex        =   3
         Top             =   1215
         Width           =   2655
      End
      Begin VB.TextBox txtPDFRes 
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Text            =   "600"
         Top             =   2685
         Width           =   615
      End
      Begin VB.ComboBox cmbPDFOverprint 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "ctlOptFormatPDF.ctx":033E
         Left            =   120
         List            =   "ctlOptFormatPDF.ctx":0340
         Style           =   2  'Dropdown-Liste
         TabIndex        =   10
         Top             =   3420
         Width           =   2655
      End
      Begin VB.CheckBox chkPDFASCII85 
         Appearance      =   0  '2D
         Caption         =   "Convert binary data to ASCII85"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   3930
         Width           =   5880
      End
      Begin VB.CheckBox chkPDFOptimize 
         Appearance      =   0  '2D
         Caption         =   "Fast web view"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   4350
         Width           =   5880
      End
      Begin VB.Label lblPDFDefaultSettings 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default settings:"
         Height          =   195
         Left            =   120
         TabIndex        =   92
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label lblPDFAutoRotate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Auto-Rotate Pages:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1740
         Width           =   1395
      End
      Begin VB.Label lblPDFCompat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compatibility:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   1020
         Width           =   915
      End
      Begin VB.Label lblPDFResolution 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resolution:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   2475
         Width           =   795
      End
      Begin VB.Label lblPDFOverprint 
         AutoSize        =   -1  'True
         Caption         =   "Overprint:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   3210
         Width           =   690
      End
      Begin VB.Label lblPDFDPI 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "dpi"
         Height          =   195
         Left            =   795
         TabIndex        =   8
         Top             =   2730
         Width           =   210
      End
   End
   Begin MSComctlLib.TabStrip tbstrPDFOptions 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   9551
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ctlOptFormatPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub UserControl_Initialize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ctl As Control
50020  Dim i As Long
50030  tbstrPDFOptions.Left = 0
50040  tbstrPDFOptions.Top = 0
50050  tbstrPDFOptions.Height = dmFraPDFSecurity.Height + 420
50060  UserControl.Height = tbstrPDFOptions.Height + 500
50070
50080  With tbstrPDFOptions.Tabs
50090   .Clear
50100   .Add , "General"
50110   .Add , "Compression"
50120   .Add , "Fonts"
50130   .Add , "Colors"
50140   .Add , "Security"
50150   .Add , "Signing"
50160  End With
50170  With cmbPDFDefaultSettings
50180   .Clear
50190   For i = 1 To 5
50200    .AddItem ""
50210   Next i
50220  End With
50230  With cmbPDFCompat
50240   .Clear
50250   For i = 1 To 4
50260    .AddItem ""
50270   Next i
50280  End With
50290  With cmbPDFRotate
50300   .Clear
50310   For i = 1 To 3
50320    .AddItem ""
50330   Next i
50340  End With
50350  With cmbPDFOverprint
50360   .Clear
50370   For i = 1 To 2
50380    .AddItem ""
50390   Next i
50400  End With
50410  With cmbPDFColorComp
50420   .Clear
50430   For i = 1 To 7
50440    .AddItem ""
50450   Next i
50460  End With
50470  With cmbPDFColorResample
50480   .Clear
50490   For i = 1 To 2
50500    .AddItem ""
50510   Next i
50520  End With
50530  With cmbPDFGreyComp
50540   .Clear
50550   For i = 1 To 7
50560    .AddItem ""
50570   Next i
50580  End With
50590  With cmbPDFGreyResample
50600   .Clear
50610   For i = 1 To 2
50620    .AddItem ""
50630   Next i
50640  End With
50650  With cmbPDFMonoComp
50660   .Clear
50670   For i = 1 To 3
50680    .AddItem ""
50690   Next i
50700  End With
50710  With cmbPDFMonoResample
50720   .Clear
50730   For i = 1 To 2
50740    .AddItem ""
50750   Next i
50760  End With
50770  With cmbPDFColorModel
50780   .Clear
50790   For i = 1 To 3
50800    .AddItem ""
50810   Next i
50820  End With
50830  txtPDFRes.Text = 600
50840  cmbPDFCompat.ListIndex = 1
50850  cmbPDFRotate.ListIndex = 0
50860  cmbPDFOverprint.ListIndex = 0
50870  chkPDFASCII85.value = 0
50880
50890  chkPDFTextComp.value = 1
50900
50910  chkPDFColorComp.value = 1
50920  chkPDFColorResample.value = 0
50930  cmbPDFColorComp.ListIndex = 0
50940  cmbPDFColorResample.ListIndex = 0
50950  txtPDFColorRes.Text = 300
50960
50970  chkPDFGreyComp.value = 1
50980  chkPDFGreyResample.value = 0
50990  cmbPDFGreyComp.ListIndex = 0
51000  cmbPDFGreyResample.ListIndex = 0
51010  txtPDFGreyRes.Text = 300
51020
51030  chkPDFMonoComp.value = 1
51040  chkPDFMonoResample.value = 0
51050  cmbPDFMonoComp.ListIndex = 0
51060  cmbPDFMonoResample.ListIndex = 0
51070  txtPDFMonoRes.Text = 1200
51080
51090  chkPDFEmbedAll.value = 1
51100  chkPDFSubSetFonts.value = 1
51110  txtPDFSubSetPerc.Text = 100
51120
51130  cmbPDFColorModel.ListIndex = 1
51140  chkPDFCMYKtoRGB.value = 1
51150  chkPDFPreserveOverprint.value = 1
51160  chkPDFPreserveTransfer.value = 1
51170  chkPDFPreserveHalftone.value = 0
51180
51190  With cmbPDFEncryptor
51200   .Clear
51210   .AddItem "Ghostscript (>= 8.14)"
51220   .ItemData(.NewIndex) = 0
51230   .AddItem "PDFEnc"
51240   .ItemData(.NewIndex) = 1
51250
51260   SecurityIsPossible = True
51270
51280   If FileExists(GetPDFCreatorApplicationPath & "pdfenc.exe") = False Then
51290    .RemoveItem 1
51300    .ListIndex = 0
51310    Options.PDFEncryptor = .ItemData(.ListIndex)
51320   End If
51330   If GhostScriptSecurity = False Then
51340    .RemoveItem 0
51350   End If
51360   If .ListCount = 0 Then
51370     chkUseSecurity.value = 0
51380     chkUseSecurity.Enabled = False
51390     SecurityIsPossible = False
51400    Else
51410     For i = 0 To .ListCount - 1
51420      If .ItemData(i) = Options.PDFEncryptor Then
51430       .ListIndex = i
51440       Exit For
51450      End If
51460     Next i
51470     If .ListIndex = -1 Then
51480      .ListIndex = 0
51490      Options.PDFEncryptor = .ItemData(.ListIndex)
51500     End If
51510   End If
51520  End With
51530
51540  If Options.PDFHighEncryption <> 0 Then
51550    optEncHigh.value = True
51560   Else
51570    optEncLow.value = True
51580  End If
51590
51600   With tbstrPDFOptions
51610   .Top = 50
51620   .Left = 0
51630  End With
51640
51650  UpdateSecurityFields
51660
51670  tbstrPDFOptions.ZOrder 1
51680  tbstrPDFOptions_Click
51690
51700  SetFrames Options.OptionsDesign
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "UserControl_Initialize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SetFrames(OptionsDesign As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ctl As Control
50020  For Each ctl In UserControl.Controls
50030   If TypeOf ctl Is dmFrame Then
50040    SetFrame ctl, OptionsDesign
50050   End If
50060  Next ctl
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "SetFrames")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub UserControl_Resize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  tbstrPDFOptions.Width = UserControl.Width
50020  With dmFraPDFGeneral
50030   .Top = tbstrPDFOptions.ClientTop + 30
50040   .Left = tbstrPDFOptions.Left + (tbstrPDFOptions.Width - .Width) / 2
50050   dmfraPDFCompress.Top = .Top
50060   dmfraPDFCompress.Left = .Left
50070   dmFraPDFFonts.Top = .Top
50080   dmFraPDFFonts.Left = .Left
50090   dmFraPDFColors.Top = .Top
50100   dmFraPDFColors.Left = .Left
50110   dmFraPDFColorOptions.Top = dmFraPDFColors.Top + dmFraPDFColors.Height + 50
50120   dmFraPDFColorOptions.Left = .Left
50130   dmFraPDFSecurity.Top = .Top
50140   dmFraPDFSecurity.Left = .Left
50150   dmFraPDFSigning.Top = .Top
50160   dmFraPDFSigning.Left = .Left
50170  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "UserControl_Resize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SetLanguageStrings()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With LanguageStrings
50020   tbstrPDFOptions.Tabs(1).Caption = .OptionsPDFGeneral
50030   tbstrPDFOptions.Tabs(2).Caption = .OptionsPDFCompression
50040   tbstrPDFOptions.Tabs(3).Caption = .OptionsPDFFonts
50050   tbstrPDFOptions.Tabs(4).Caption = .OptionsPDFColors
50060   tbstrPDFOptions.Tabs(5).Caption = .OptionsPDFSecurity
50070   tbstrPDFOptions.Tabs(6).Caption = .OptionsPDFSigning
50080
50090   lblPDFDefaultSettings.Caption = .OptionsPDFGeneralDefaultSettings
50100   cmbPDFDefaultSettings.List(0) = .OptionsPDFGeneralDefaultSettingsDefault
50110   cmbPDFDefaultSettings.List(1) = .OptionsPDFGeneralDefaultSettingsScreen
50120   cmbPDFDefaultSettings.List(2) = .OptionsPDFGeneralDefaultSettingsEbook
50130   cmbPDFDefaultSettings.List(3) = .OptionsPDFGeneralDefaultSettingsPrinter
50140   cmbPDFDefaultSettings.List(4) = .OptionsPDFGeneralDefaultSettingsPrepress
50150
50160   cmbPDFCompat.List(0) = .OptionsPDFGeneralCompatibility01
50170   cmbPDFCompat.List(1) = .OptionsPDFGeneralCompatibility02
50180   cmbPDFCompat.List(2) = .OptionsPDFGeneralCompatibility03
50190   cmbPDFCompat.List(3) = .OptionsPDFGeneralCompatibility04
50200   cmbPDFRotate.List(0) = .OptionsPDFGeneralRotate01
50210   cmbPDFRotate.List(1) = .OptionsPDFGeneralRotate02
50220   cmbPDFRotate.List(2) = .OptionsPDFGeneralRotate03
50230   cmbPDFOverprint.List(0) = .OptionsPDFGeneralOverprint01
50240   cmbPDFOverprint.List(1) = .OptionsPDFGeneralOverprint02
50250
50260   cmbPDFColorComp.List(0) = .OptionsPDFCompressionColorComp01
50270   cmbPDFColorComp.List(1) = .OptionsPDFCompressionColorComp02
50280   cmbPDFColorComp.List(2) = .OptionsPDFCompressionColorComp03
50290   cmbPDFColorComp.List(3) = .OptionsPDFCompressionColorComp04
50300   cmbPDFColorComp.List(4) = .OptionsPDFCompressionColorComp05
50310   cmbPDFColorComp.List(5) = .OptionsPDFCompressionColorComp06
50320   cmbPDFColorComp.List(6) = .OptionsPDFCompressionColorComp07
50330
50340   cmbPDFColorResample.List(0) = .OptionsPDFCompressionColorResample01
50350   cmbPDFColorResample.List(1) = .OptionsPDFCompressionColorResample02
50360
50370   cmbPDFGreyComp.List(0) = .OptionsPDFCompressionGreyComp01
50380   cmbPDFGreyComp.List(1) = .OptionsPDFCompressionGreyComp02
50390   cmbPDFGreyComp.List(2) = .OptionsPDFCompressionGreyComp03
50400   cmbPDFGreyComp.List(3) = .OptionsPDFCompressionGreyComp04
50410   cmbPDFGreyComp.List(4) = .OptionsPDFCompressionGreyComp05
50420   cmbPDFGreyComp.List(5) = .OptionsPDFCompressionGreyComp06
50430   cmbPDFGreyComp.List(6) = .OptionsPDFCompressionGreyComp07
50440
50450   cmbPDFGreyResample.List(0) = .OptionsPDFCompressionGreyResample01
50460   cmbPDFGreyResample.List(1) = .OptionsPDFCompressionGreyResample02
50470
50480   cmbPDFMonoComp.List(0) = .OptionsPDFCompressionMonoComp01
50490   cmbPDFMonoComp.List(1) = .OptionsPDFCompressionMonoComp02
50500   cmbPDFMonoComp.List(2) = .OptionsPDFCompressionMonoComp03
50510
50520   cmbPDFMonoResample.List(0) = .OptionsPDFCompressionMonoResample01
50530   cmbPDFMonoResample.List(1) = .OptionsPDFCompressionMonoResample02
50540
50550   cmbPDFColorModel.List(0) = .OptionsPDFColorsColorModel01
50560   cmbPDFColorModel.List(1) = .OptionsPDFColorsColorModel02
50570   cmbPDFColorModel.List(2) = .OptionsPDFColorsColorModel03
50580
50590   dmFraPDFGeneral.Caption = .OptionsPDFGeneralCaption
50600   chkPDFOptimize.Caption = .OptionsPDFOptimize
50610   lblPDFCompat.Caption = .OptionsPDFGeneralCompatibility
50620   lblPDFAutoRotate.Caption = .OptionsPDFGeneralAutorotate
50630   lblPDFResolution.Caption = .OptionsPDFGeneralResolution
50640   lblPDFOverprint.Caption = .OptionsPDFGeneralOverprint
50650   chkPDFASCII85.Caption = .OptionsPDFGeneralASCII85
50660
50670   dmfraPDFCompress.Caption = .OptionsPDFCompressionCaption
50680   chkPDFTextComp.Caption = .OptionsPDFCompressionTextComp
50690   dmFraPDFColor.Caption = .OptionsPDFCompressionColor
50700   chkPDFColorComp.Caption = .OptionsPDFCompressionColorComp
50710   chkPDFColorResample.Caption = .OptionsPDFCompressionColorResample
50720   lblPDFColorRes.Caption = .OptionsPDFCompressionColorRes
50730   dmFraPDFGrey.Caption = .OptionsPDFCompressionGrey
50740   chkPDFGreyComp.Caption = .OptionsPDFCompressionGreyComp
50750   chkPDFGreyResample.Caption = .OptionsPDFCompressionGreyResample
50760   lblPDFGreyRes.Caption = .OptionsPDFCompressionGreyRes
50770   dmFraPDFMono.Caption = .OptionsPDFCompressionMono
50780   chkPDFMonoComp.Caption = .OptionsPDFCompressionMonoComp
50790   chkPDFMonoResample.Caption = .OptionsPDFCompressionMonoResample
50800   lblPDFMonoRes.Caption = .OptionsPDFCompressionMonoRes
50810
50820   dmFraPDFFonts.Caption = .OptionsPDFFontsCaption
50830   chkPDFEmbedAll.Caption = .OptionsPDFFontsEmbedAll
50840   chkPDFSubSetFonts.Caption = .OptionsPDFFontsSubSetFonts
50850
50860   dmFraPDFColors.Caption = .OptionsPDFColorsCaption
50870   chkPDFCMYKtoRGB.Caption = .OptionsPDFColorsCMYKtoRGB
50880   dmFraPDFColorOptions.Caption = .OptionsPDFColorsColorOptions
50890   chkPDFPreserveOverprint.Caption = .OptionsPDFColorsPreserveOverprint
50900   chkPDFPreserveTransfer.Caption = .OptionsPDFColorsPreserveTransfer
50910   chkPDFPreserveHalftone.Caption = .OptionsPDFColorsPreserveHalftone
50920
50930   dmFraPDFSigning.Caption = .OptionsPDFSigningCaption
50940   dmFraPDFEncryptor.Caption = .OptionsPDFEncryptor
50950   dmFraPDFSecurity.Caption = .OptionsPDFSecurityCaption
50960   chkUseSecurity.Caption = .OptionsPDFUseSecurity
50970   dmFraPDFEncLevel.Caption = .OptionsPDFEncryptionLevel
50980   optEncHigh.Caption = .OptionsPDFEncryptionHigh
50990   optEncLow.Caption = .OptionsPDFEncryptionLow
51000   dmFraSecurityPass.Caption = .OptionsPDFPasswords
51010   chkUserPass.Caption = .OptionsPDFUserPass
51020   chkOwnerPass.Caption = .OptionsPDFOwnerPass
51030   dmFraPDFPermissions.Caption = .OptionsPDFDisallowUser
51040   dmFraPDFHighPermissions.Caption = .OptionsPDFEnhancedPermissions
51050   chkAllowPrinting.Caption = .OptionsPDFDisallowPrint
51060   chkAllowModifyContents.Caption = .OptionsPDFDisallowModify
51070   chkAllowCopy.Caption = .OptionsPDFDisallowCopy
51080   chkAllowModifyAnnotations.Caption = .OptionsPDFDisallowModifyComments
51090   chkAllowDegradedPrinting.Caption = .OptionsPDFAllowDegradedPrinting
51100   chkAllowFillIn.Caption = .OptionsPDFAllowFillIn
51110   chkAllowAssembly.Caption = .OptionsPDFAllowAssembly
51120   chkAllowScreenReaders.Caption = .OptionsPDFAllowScreenReaders
51130
51140   chkSignPDF.Caption = .OptionsPDFSigningSignPdfFile
51150   lblPFXFile.Caption = .OptionsPDFSigningPfxFile
51160   lblSignatureReason.Caption = .OptionsPDFSigningSignatureReason
51170   lblSignatureContact.Caption = .OptionsPDFSigningSignatureContact
51180   lblSignatureLocation.Caption = .OptionsPDFSigningSignatureLocation
51190   dmFraSignaturePosition.Caption = .OptionsPDFSigningSignaturePosition
51200   chkSignatureVisible.Caption = .OptionsPDFSigningSignatureVisible
51210   lblSignatureOnPage.Caption = .OptionsPDFSigningSignatureOnPage
51220   lblLeftX.Caption = .OptionsPDFSigningSignaturePositionLeftX
51230   lblLeftY.Caption = .OptionsPDFSigningSignaturePositionLeftY
51240   lblRightX.Caption = .OptionsPDFSigningSignaturePositionRightX
51250   lblRightY.Caption = .OptionsPDFSigningSignaturePositionRightY
51260   chkMultiSignature.Caption = .OptionsPDFSigningSignatureMultiSignature
51270  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "SetLanguageStrings")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SetOptions()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With Options1
50020   chkAllowAssembly.value = .PDFAllowAssembly
50030   chkAllowDegradedPrinting.value = .PDFAllowDegradedPrinting
50040   chkAllowFillIn.value = .PDFAllowFillIn
50050   chkAllowScreenReaders.value = .PDFAllowScreenReaders
50060   chkPDFCMYKtoRGB.value = .PDFColorsCMYKToRGB
50070   cmbPDFColorModel.ListIndex = .PDFColorsColorModel
50080   chkPDFPreserveHalftone.value = .PDFColorsPreserveHalftone
50090   chkPDFPreserveOverprint.value = .PDFColorsPreserveOverprint
50100   chkPDFPreserveTransfer.value = .PDFColorsPreserveTransfer
50110   chkPDFColorComp.value = .PDFCompressionColorCompression
50120   cmbPDFColorComp.ListIndex = .PDFCompressionColorCompressionChoice
50130   chkPDFColorResample.value = .PDFCompressionColorResample
50140   cmbPDFColorResample.ListIndex = .PDFCompressionColorResampleChoice
50150   txtPDFColorRes.Text = .PDFCompressionColorResolution
50160   chkPDFGreyComp.value = .PDFCompressionGreyCompression
50170   cmbPDFGreyComp.ListIndex = .PDFCompressionGreyCompressionChoice
50180   chkPDFGreyResample.value = .PDFCompressionGreyResample
50190   cmbPDFGreyResample.ListIndex = .PDFCompressionGreyResampleChoice
50200   txtPDFGreyRes.Text = .PDFCompressionGreyResolution
50210   chkPDFMonoComp.value = .PDFCompressionMonoCompression
50220   cmbPDFMonoComp.ListIndex = .PDFCompressionMonoCompressionChoice
50230   chkPDFMonoResample.value = .PDFCompressionMonoResample
50240   cmbPDFMonoResample.ListIndex = .PDFCompressionMonoResampleChoice
50250   txtPDFMonoRes.Text = .PDFCompressionMonoResolution
50260   chkPDFTextComp.value = .PDFCompressionTextCompression
50270   chkAllowCopy.value = .PDFDisallowCopy
50280   chkAllowModifyAnnotations.value = .PDFDisallowModifyAnnotations
50290   chkAllowModifyContents.value = .PDFDisallowModifyContents
50300   chkAllowPrinting.value = .PDFDisallowPrinting
50310   cmbPDFEncryptor.ItemData(cmbPDFEncryptor.ListIndex) = .PDFEncryptor
50320   chkPDFEmbedAll.value = .PDFFontsEmbedAll
50330   chkPDFSubSetFonts.value = .PDFFontsSubSetFonts
50340   txtPDFSubSetPerc.Text = .PDFFontsSubSetFontsPercent
50350   chkPDFASCII85.value = .PDFGeneralASCII85
50360   cmbPDFRotate.ListIndex = .PDFGeneralAutorotate
50370   cmbPDFCompat.ListIndex = .PDFGeneralCompatibility
50380   cmbPDFDefaultSettings.ListIndex = .PDFGeneralDefault
50390   cmbPDFOverprint.ListIndex = .PDFGeneralOverprint
50400   txtPDFRes.Text = .PDFGeneralResolution
50410 '  optEncHigh.value = .PDFHighEncryption
50420 '  optEncLow.value = .PDFLowEncryption
50430   chkPDFOptimize.value = .PDFOptimize
50440   chkOwnerPass.value = .PDFOwnerPass
50450   chkUserPass.value = .PDFUserPass
50460   chkUseSecurity.value = .PDFUseSecurity
50470
50480   chkSignPDF.value = .PDFSigningSignPDF
50490   txtPFXfile.Text = .PDFSigningPFXFile
50500   txtSignatureReason.Text = .PDFSigningSignatureReason
50510   txtSignatureContact.Text = .PDFSigningSignatureContact
50520   txtSignatureLocation.Text = .PDFSigningSignatureLocation
50530
50540   chkSignatureVisible.value = .PDFSigningSignatureVisible
50550   txtSignatureOnPage.Text = .PDFSigningSignatureOnPage
50560   txtLeftX.Text = .PDFSigningSignatureLeftX
50570   txtLeftY.Text = .PDFSigningSignatureLeftY
50580   txtRightX.Text = .PDFSigningSignatureRightX
50590   txtRightY.Text = .PDFSigningSignatureRightY
50600   chkMultiSignature.value = .PDFSigningMultiSignature
50610  End With
50620  If chkSignPDF.value = 1 Then
50630    EnableControls True
50640   Else
50650    EnableControls False
50660  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "SetOptions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub GetOptions()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With Options1
50020   .PDFAllowAssembly = Abs(chkAllowAssembly.value)
50030   .PDFAllowDegradedPrinting = Abs(chkAllowDegradedPrinting.value)
50040   .PDFAllowFillIn = Abs(chkAllowFillIn.value)
50050   .PDFAllowScreenReaders = Abs(chkAllowScreenReaders.value)
50060   .PDFColorsCMYKToRGB = Abs(chkPDFCMYKtoRGB.value)
50070   If LenB(CStr(cmbPDFColorModel.ListIndex)) > 0 Then
50080    .PDFColorsColorModel = cmbPDFColorModel.ListIndex
50090   End If
50100   .PDFColorsPreserveHalftone = Abs(chkPDFPreserveHalftone.value)
50110   .PDFColorsPreserveOverprint = Abs(chkPDFPreserveOverprint.value)
50120   .PDFColorsPreserveTransfer = Abs(chkPDFPreserveTransfer.value)
50130   .PDFCompressionColorCompression = Abs(chkPDFColorComp.value)
50140   If LenB(CStr(cmbPDFColorComp.ListIndex)) > 0 Then
50150    .PDFCompressionColorCompressionChoice = cmbPDFColorComp.ListIndex
50160   End If
50170   .PDFCompressionColorResample = Abs(chkPDFColorResample.value)
50180   If LenB(CStr(cmbPDFColorResample.ListIndex)) > 0 Then
50190    .PDFCompressionColorResampleChoice = cmbPDFColorResample.ListIndex
50200   End If
50210   If LenB(txtPDFColorRes.Text) > 0 Then
50220    .PDFCompressionColorResolution = txtPDFColorRes.Text
50230   End If
50240   .PDFCompressionGreyCompression = Abs(chkPDFGreyComp.value)
50250   If LenB(CStr(cmbPDFGreyComp.ListIndex)) > 0 Then
50260    .PDFCompressionGreyCompressionChoice = cmbPDFGreyComp.ListIndex
50270   End If
50280   .PDFCompressionGreyResample = Abs(chkPDFGreyResample.value)
50290   If LenB(CStr(cmbPDFGreyResample.ListIndex)) > 0 Then
50300    .PDFCompressionGreyResampleChoice = cmbPDFGreyResample.ListIndex
50310   End If
50320   If LenB(txtPDFGreyRes.Text) > 0 Then
50330    .PDFCompressionGreyResolution = txtPDFGreyRes.Text
50340   End If
50350   .PDFCompressionMonoCompression = Abs(chkPDFMonoComp.value)
50360   If LenB(CStr(cmbPDFMonoComp.ListIndex)) > 0 Then
50370    .PDFCompressionMonoCompressionChoice = cmbPDFMonoComp.ListIndex
50380   End If
50390   .PDFCompressionMonoResample = Abs(chkPDFMonoResample.value)
50400   If LenB(CStr(cmbPDFMonoResample.ListIndex)) > 0 Then
50410    .PDFCompressionMonoResampleChoice = cmbPDFMonoResample.ListIndex
50420   End If
50430   If LenB(txtPDFMonoRes.Text) > 0 Then
50440    .PDFCompressionMonoResolution = txtPDFMonoRes.Text
50450   End If
50460   .PDFCompressionTextCompression = Abs(chkPDFTextComp.value)
50470   .PDFDisallowCopy = Abs(chkAllowCopy.value)
50480   .PDFDisallowModifyAnnotations = Abs(chkAllowModifyAnnotations.value)
50490   .PDFDisallowModifyContents = Abs(chkAllowModifyContents.value)
50500   .PDFDisallowPrinting = Abs(chkAllowPrinting.value)
50510   If cmbPDFEncryptor.ListIndex < 0 Then
50520     .PDFEncryptor = 0
50530    Else
50540     .PDFEncryptor = cmbPDFEncryptor.ItemData(cmbPDFEncryptor.ListIndex)
50550   End If
50560   .PDFFontsEmbedAll = Abs(chkPDFEmbedAll.value)
50570   .PDFFontsSubSetFonts = Abs(chkPDFSubSetFonts.value)
50580   If LenB(txtPDFSubSetPerc.Text) > 0 Then
50590    .PDFFontsSubSetFontsPercent = txtPDFSubSetPerc.Text
50600   End If
50610   .PDFGeneralASCII85 = Abs(chkPDFASCII85.value)
50620   If LenB(CStr(cmbPDFRotate.ListIndex)) > 0 Then
50630    .PDFGeneralAutorotate = cmbPDFRotate.ListIndex
50640   End If
50650   If LenB(CStr(cmbPDFCompat.ListIndex)) > 0 Then
50660    .PDFGeneralCompatibility = cmbPDFCompat.ListIndex
50670   End If
50680   If LenB(CStr(cmbPDFDefaultSettings.ListIndex)) > 0 Then
50690    .PDFGeneralDefault = cmbPDFDefaultSettings.ListIndex
50700   End If
50710   If LenB(CStr(cmbPDFOverprint.ListIndex)) > 0 Then
50720    .PDFGeneralOverprint = cmbPDFOverprint.ListIndex
50730   End If
50740   If LenB(txtPDFRes.Text) > 0 Then
50750    .PDFGeneralResolution = txtPDFRes.Text
50760   End If
50770   .PDFHighEncryption = Abs(optEncHigh.value)
50780   .PDFLowEncryption = Abs(optEncLow.value)
50790   .PDFOptimize = Abs(chkPDFOptimize.value)
50800   .PDFOwnerPass = Abs(chkOwnerPass.value)
50810   .PDFUserPass = Abs(chkUserPass.value)
50820   .PDFUseSecurity = Abs(chkUseSecurity.value)
50830
50840   .PDFSigningSignPDF = Abs(chkSignPDF.value)
50850   .PDFSigningPFXFile = txtPFXfile.Text
50860   .PDFSigningSignatureReason = txtSignatureReason.Text
50870   .PDFSigningSignatureContact = txtSignatureContact.Text
50880   .PDFSigningSignatureLocation = txtSignatureLocation.Text
50890
50900   .PDFSigningSignatureVisible = Abs(chkSignatureVisible.value)
50910   If LenB(txtSignatureOnPage.Text) > 0 Then
50920    .PDFSigningSignatureOnPage = txtSignatureOnPage.Text
50930   End If
50940   If LenB(txtLeftX.Text) > 0 Then
50950    .PDFSigningSignatureLeftX = txtLeftX.Text
50960   End If
50970   If LenB(txtLeftY.Text) > 0 Then
50980    .PDFSigningSignatureLeftY = txtLeftY.Text
50990   End If
51000   If LenB(txtRightX.Text) > 0 Then
51010    .PDFSigningSignatureRightX = txtRightX.Text
51020   End If
51030   If LenB(txtRightY.Text) > 0 Then
51040    .PDFSigningSignatureRightY = txtRightY.Text
51050   End If
51060   .PDFSigningMultiSignature = Abs(chkMultiSignature.value)
51070  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "GetOptions")
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
50070  dmFraPDFSigning.Visible = False
50080  dmFraPDFGeneral.Enabled = False
50090  dmfraPDFCompress.Enabled = False
50100  dmFraPDFFonts.Enabled = False
50110  dmFraPDFColors.Enabled = False
50120  dmFraPDFColorOptions.Enabled = False
50130  dmFraPDFSecurity.Enabled = False
50140  dmFraPDFSigning.Enabled = False
50151  Select Case tbstrPDFOptions.SelectedItem.Index
        Case 1:
50170    dmFraPDFGeneral.Visible = True
50180    dmFraPDFGeneral.Enabled = True
50190   Case 2:
50200    dmfraPDFCompress.Visible = True
50210    dmfraPDFCompress.Enabled = True
50220    dmFraPDFColor.Visible = True
50230    dmFraPDFColor.Enabled = True
50240    dmFraPDFGrey.Visible = True
50250    dmFraPDFGrey.Enabled = True
50260    dmFraPDFMono.Visible = True
50270    dmFraPDFMono.Enabled = True
50280   Case 3:
50290    dmFraPDFFonts.Visible = True
50300    dmFraPDFFonts.Enabled = True
50310   Case 4:
50320    dmFraPDFColors.Visible = True
50330    dmFraPDFColorOptions.Visible = True
50340    dmFraPDFColors.Enabled = True
50350    dmFraPDFColorOptions.Enabled = True
50360   Case 5:
50370    dmFraPDFSecurity.Visible = True
50380    dmFraPDFSecurity.Enabled = True
50390    dmFraPDFEncryptor.Visible = True
50400    dmFraPDFEncryptor.Enabled = True
50410    dmFraPDFEncLevel.Visible = True
50420    dmFraPDFEncLevel.Enabled = True
50430    dmFraSecurityPass.Visible = True
50440    dmFraSecurityPass.Enabled = True
50450    dmFraPDFPermissions.Visible = True
50460    dmFraPDFPermissions.Enabled = True
50470    dmFraPDFHighPermissions.Visible = True
50480    dmFraPDFHighPermissions.Enabled = True
50490    UpdateSecurityFields
50500    If cmbPDFCompat.ListIndex < 2 Then
50510      optEncLow.Enabled = True
50520      optEncHigh.Enabled = False
50530     Else
50540      optEncLow.Enabled = False
50550      optEncHigh.Enabled = True
50560    End If
50570    If SecurityIsPossible = False Then
50580     MsgBox LanguageStrings.MessagesMsg19
50590    End If
50600   Case 6:
50610    dmFraPDFSigning.Visible = True
50620    dmFraPDFSigning.Enabled = True
50630    If PDFSigningIsPossible = False Then
50640     chkSignPDF.Enabled = False
50650     EnableControls False
50660     MsgBox LanguageStrings.MessagesMsg39
50670    End If
50680  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "tbstrPDFOptions_Click")
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
50010  If chkUserPass.value = 0 Then
50020   If chkOwnerPass.value = 0 Then
50030    chkOwnerPass.value = 1
50040   End If
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "chkOwnerPass_Click")
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
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "chkPDFColorComp_Click")
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
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "chkPDFColorResample_Click")
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
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "chkPDFGreyComp_Click")
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
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "chkPDFGreyResample_Click")
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
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "chkPDFMonoComp_Click")
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
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "chkPDFMonoResample_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetPDFColorComprSettings()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkPDFColorComp.value = 1 Then
50020    cmbPDFColorComp.Enabled = True
50030    If cmbPDFColorComp.ListIndex = 0 Then
50040      chkPDFColorResample.Enabled = False
50050      cmbPDFColorResample.Enabled = False
50060      lblPDFColorRes.Enabled = False
50070      txtPDFColorRes.Enabled = False
50080     Else
50090      chkPDFColorResample.Enabled = True
50100      If chkPDFColorResample.value = 1 Then
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
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "SetPDFColorComprSettings")
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
50010  If chkPDFGreyComp.value = 1 Then
50020    cmbPDFGreyComp.Enabled = True
50030    If cmbPDFGreyComp.ListIndex = 0 Then
50040      chkPDFGreyResample.Enabled = False
50050      cmbPDFGreyResample.Enabled = False
50060      lblPDFGreyRes.Enabled = False
50070      txtPDFGreyRes.Enabled = False
50080     Else
50090      chkPDFGreyResample.Enabled = True
50100      If chkPDFGreyResample.value = 1 Then
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
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "SetPDFGreyComprSettings")
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
50010  If chkPDFMonoComp.value = 1 Then
50020    cmbPDFMonoComp.Enabled = True
50030    chkPDFMonoResample.Enabled = True
50040    If chkPDFMonoResample.value = 1 Then
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
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "SetPDFMonoComprSettings")
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
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "cmbPDFColorComp_Click")
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
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "cmbPDFGreyComp_Click")
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
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "cmbPDFMonoComp_Click")
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
50010  If chkUseSecurity.value = False Then
50020    dmFraPDFEncryptor.Enabled = False
50030    cmbPDFEncryptor.Enabled = False
50040
50050    dmFraPDFEncLevel.Enabled = False
50060    optEncHigh.Enabled = False
50070    optEncLow.Enabled = False
50080
50090    dmFraSecurityPass.Enabled = False
50100    chkUserPass.Enabled = False
50110    chkOwnerPass.Enabled = False
50120
50130    dmFraPDFPermissions.Enabled = False
50140    chkAllowPrinting.Enabled = False
50150    chkAllowCopy.Enabled = False
50160    chkAllowModifyAnnotations.Enabled = False
50170    chkAllowModifyContents.Enabled = False
50180
50190    dmFraPDFHighPermissions.Enabled = False
50200    chkAllowDegradedPrinting.Enabled = False
50210    chkAllowFillIn.Enabled = False
50220    chkAllowScreenReaders.Enabled = False
50230    chkAllowAssembly.Enabled = False
50240   Else
50250    dmFraPDFEncryptor.Enabled = True
50260    cmbPDFEncryptor.Enabled = True
50270    dmFraPDFEncLevel.Enabled = True
50280
50290    dmFraSecurityPass.Enabled = True
50300    chkUserPass.Enabled = True
50310    chkOwnerPass.Enabled = True
50320
50330    dmFraPDFPermissions.Enabled = True
50340    chkAllowPrinting.Enabled = True
50350    chkAllowCopy.Enabled = True
50360    chkAllowModifyAnnotations.Enabled = True
50370    chkAllowModifyContents.Enabled = True
50380
50390    If cmbPDFCompat.ListIndex < 2 Then
50400      optEncLow.Enabled = True
50410      optEncHigh.Enabled = False
50420      optEncLow.value = True
50430      chkAllowDegradedPrinting.Enabled = False
50440      chkAllowFillIn.Enabled = False
50450      chkAllowScreenReaders.Enabled = False
50460      chkAllowAssembly.Enabled = False
50470      dmFraPDFHighPermissions.Enabled = False
50480     Else
50490      optEncLow.Enabled = False
50500      optEncHigh.Enabled = True
50510      optEncHigh.value = True
50520      chkAllowDegradedPrinting.Enabled = True
50530      chkAllowFillIn.Enabled = True
50540      chkAllowScreenReaders.Enabled = True
50550      chkAllowAssembly.Enabled = True
50560      dmFraPDFHighPermissions.Enabled = True
50570    End If
50580  End If
50590
50600  If chkOwnerPass.value = 0 And chkUserPass.value = 0 Then
50610   chkOwnerPass.value = 1
50620   Options.PDFOwnerPass = 1
50630  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "UpdateSecurityFields")
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
50010  If chkOwnerPass.value = 0 Then
50020   If chkUserPass.value = 0 Then
50030    chkUserPass.value = 1
50040    chkOwnerPass.value = 1
50050   End If
50060   SavePasswordsForThisSession = False
50070  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "chkUserPass_Click")
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
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "chkUseSecurity_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Property Get PDFOptionsIndex()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  PDFOptionsIndex = tbstrPDFOptions.SelectedItem.Index
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "PDFOptionsIndex [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Private Sub chkSignPDF_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkSignPDF.value = 1 Then
50020    EnableControls True
50030   Else
50040    EnableControls False
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "chkSignPDF_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub EnableControls(value As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  lblPFXFile.Enabled = value
50020  txtPFXfile.Enabled = value
50030  cmdGetPFXFile.Enabled = value
50040  txtPFXFilePreview.Enabled = value
50050  lblSignatureReason.Enabled = value
50060  txtSignatureReason.Enabled = value
50070  lblSignatureContact.Enabled = value
50080  txtSignatureContact.Enabled = value
50090  lblSignatureLocation.Enabled = value
50100  dmFraSignaturePosition.Enabled = value
50110  chkSignatureVisible.Enabled = value
50120  If chkSignatureVisible.value = 1 Then
50130    EnableSignPositionControls True
50140   Else
50150    EnableSignPositionControls False
50160  End If
50170  chkMultiSignature.Enabled = value
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "EnableControls")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub EnableSignPositionControls(value As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  lblLeftX.Enabled = value
50020  txtLeftX.Enabled = value
50030  lblLeftY.Enabled = value
50040  txtLeftY.Enabled = value
50050  lblRightX.Enabled = value
50060  txtRightX.Enabled = value
50070  lblRightY.Enabled = value
50080  txtRightY.Enabled = value
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "EnableSignPositionControls")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkSignatureVisible_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkSignatureVisible.value = 1 Then
50020    EnableSignPositionControls True
50030   Else
50040    EnableSignPositionControls False
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "chkSignatureVisible_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdGetPFXFile_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim res As Long, files As Collection, certFilename As String
50020  With Options
50030   If LenB(.PDFSigningPFXFile) = 0 Then
50040     res = OpenFileDialog(files, "", "PFX\P12 files (*.pfx,*.p12)|*.pfx;*.p12|PFX files (*.pfx)|*pfx|P12 files (*.p12|*.p12", "*.pfx;*.p12", "C:\", "Choose a certificate", OFN_FILEMUSTEXIST Or OFN_EXPLORER Or OFN_LONGNAMES Or OFN_PATHMUSTEXIST, 0, 1)
50050     If res > 0 Then
50060      certFilename = files(1)
50070     End If
50080    Else
50090     certFilename = .PDFSigningPFXFile
50100   End If
50110   txtPFXfile.Text = certFilename
50120  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "cmdGetPFXFile_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtPFXfile_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txtPFXFilePreview.Text = txtPFXfile.Text
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "txtPFXfile_Change")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtSignatureOnPage_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  KeyAscii = AllowedKeypressChars(KeyAscii)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "txtSignatureOnPage_KeyPress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtLeftX_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  KeyAscii = AllowedKeypressChars(KeyAscii)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "txtLeftX_KeyPress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtLeftY_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  KeyAscii = AllowedKeypressChars(KeyAscii)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "txtLeftY_KeyPress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtRightX_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  KeyAscii = AllowedKeypressChars(KeyAscii)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "txtRightX_KeyPress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtRightY_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  KeyAscii = AllowedKeypressChars(KeyAscii)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "txtRightY_KeyPress")
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
50010  If cmbPDFCompat.ListIndex < 2 Then
50020    optEncLow.value = True
50030   Else
50040    optEncHigh.value = True
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "cmbPDFCompat_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
