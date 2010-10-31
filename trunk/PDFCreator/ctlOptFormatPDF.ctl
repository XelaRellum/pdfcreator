VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlOptFormatPDF 
   AutoRedraw      =   -1  'True
   ClientHeight    =   11385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19335
   ScaleHeight     =   11385
   ScaleWidth      =   19335
   ToolboxBitmap   =   "ctlOptFormatPDF.ctx":0000
   Begin PDFCreator.dmFrame dmFraPDFSigning 
      Height          =   5535
      Left            =   12960
      TabIndex        =   45
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
         TabIndex        =   67
         Top             =   360
         Width           =   5895
      End
      Begin VB.CheckBox chkMultiSignature 
         Caption         =   "Multi signature allowed"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   5160
         Width           =   5895
      End
      Begin VB.TextBox txtSignatureLocation 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   120
         TabIndex        =   55
         Top             =   2760
         Width           =   5325
      End
      Begin VB.TextBox txtSignatureContact 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   3120
         TabIndex        =   53
         Top             =   2040
         Width           =   2805
      End
      Begin VB.TextBox txtSignatureReason 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   120
         TabIndex        =   51
         Top             =   2040
         Width           =   2805
      End
      Begin VB.TextBox txtPFXFilePreview 
         Appearance      =   0  '2D
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   1320
         Width           =   5910
      End
      Begin VB.CommandButton cmdGetPFXFile 
         Caption         =   "..."
         Height          =   300
         Left            =   5640
         TabIndex        =   47
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtPFXfile 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   120
         TabIndex        =   46
         Top             =   960
         Width           =   5325
      End
      Begin PDFCreator.dmFrame dmFraSignaturePosition 
         Height          =   1935
         Left            =   120
         TabIndex        =   56
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
            TabIndex        =   71
            Text            =   "1"
            Top             =   960
            Width           =   1000
         End
         Begin VB.CheckBox chkSignatureVisible 
            Caption         =   "Signature visible in PDF"
            Height          =   255
            Left            =   240
            TabIndex        =   65
            Top             =   360
            Width           =   5415
         End
         Begin VB.TextBox txtRightY 
            Alignment       =   1  'Rechts
            Appearance      =   0  '2D
            Height          =   285
            Left            =   4560
            TabIndex        =   64
            Text            =   "200"
            Top             =   1560
            Width           =   1000
         End
         Begin VB.TextBox txtRightX 
            Alignment       =   1  'Rechts
            Appearance      =   0  '2D
            Height          =   285
            Left            =   3120
            TabIndex        =   62
            Text            =   "200"
            Top             =   1560
            Width           =   1000
         End
         Begin VB.TextBox txtLeftY 
            Alignment       =   1  'Rechts
            Appearance      =   0  '2D
            Height          =   285
            Left            =   1680
            TabIndex        =   60
            Text            =   "100"
            Top             =   1560
            Width           =   1000
         End
         Begin VB.TextBox txtLeftX 
            Alignment       =   1  'Rechts
            Appearance      =   0  '2D
            Height          =   285
            Left            =   240
            TabIndex        =   58
            Text            =   "100"
            Top             =   1560
            Width           =   1000
         End
         Begin VB.Label lblSignatureOnPage 
            AutoSize        =   -1  'True
            Caption         =   "Show signature on page"
            Height          =   195
            Left            =   240
            TabIndex        =   70
            Top             =   720
            Width           =   1725
         End
         Begin VB.Label lblRightY 
            AutoSize        =   -1  'True
            Caption         =   "Right Y"
            Height          =   195
            Left            =   4560
            TabIndex        =   63
            Top             =   1320
            Width           =   525
         End
         Begin VB.Label lblRightX 
            AutoSize        =   -1  'True
            Caption         =   "Right X"
            Height          =   195
            Left            =   3120
            TabIndex        =   61
            Top             =   1320
            Width           =   525
         End
         Begin VB.Label lblLeftY 
            AutoSize        =   -1  'True
            Caption         =   "Left Y"
            Height          =   195
            Left            =   1680
            TabIndex        =   59
            Top             =   1320
            Width           =   420
         End
         Begin VB.Label lblLeftX 
            AutoSize        =   -1  'True
            Caption         =   "Left X"
            Height          =   195
            Left            =   240
            TabIndex        =   57
            Top             =   1320
            Width           =   420
         End
      End
      Begin VB.Label lblSignatureLocation 
         AutoSize        =   -1  'True
         Caption         =   "Signature location"
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   2520
         Width           =   1275
      End
      Begin VB.Label lblSignatureContact 
         AutoSize        =   -1  'True
         Caption         =   "Signature contact"
         Height          =   195
         Left            =   3120
         TabIndex        =   52
         Top             =   1800
         Width           =   1260
      End
      Begin VB.Label lblSignatureReason 
         AutoSize        =   -1  'True
         Caption         =   "Signature reason"
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   1800
         Width           =   1200
      End
      Begin VB.Label lblPFXFile 
         AutoSize        =   -1  'True
         Caption         =   "PFX\P12 file"
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   720
         Width           =   900
      End
   End
   Begin PDFCreator.dmFrame dmFraPDFColorOptions 
      Height          =   1455
      Left            =   120
      TabIndex        =   23
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
         TabIndex        =   24
         Top             =   360
         Width           =   5910
      End
      Begin VB.CheckBox chkPDFPreserveTransfer 
         Appearance      =   0  '2D
         Caption         =   "Preserve Transfer Functions"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   25
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
         TabIndex        =   26
         Top             =   1050
         Width           =   5910
      End
   End
   Begin PDFCreator.dmFrame dmFraPDFColors 
      Height          =   1215
      Left            =   120
      TabIndex        =   20
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
         TabIndex        =   22
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
         TabIndex        =   21
         Tag             =   "RGB|CMYK|GRAY"
         Top             =   360
         Width           =   2775
      End
   End
   Begin PDFCreator.dmFrame dmFraPDFSecurity 
      Height          =   4815
      Left            =   6600
      TabIndex        =   27
      Top             =   5760
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   8493
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
         TabIndex        =   28
         Top             =   360
         Width           =   5535
      End
      Begin PDFCreator.dmFrame dmFraPDFHighPermissions 
         Height          =   855
         Left            =   120
         TabIndex        =   40
         Top             =   3840
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
            TabIndex        =   41
            Top             =   300
            Width           =   2865
         End
         Begin VB.CheckBox chkAllowFillIn 
            Appearance      =   0  '2D
            Caption         =   "Allow filling in form fields"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3100
            TabIndex        =   42
            Top             =   300
            Width           =   2760
         End
         Begin VB.CheckBox chkAllowScreenReaders 
            Appearance      =   0  '2D
            Caption         =   "Allow Screen Readers"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   525
            Width           =   2865
         End
         Begin VB.CheckBox chkAllowAssembly 
            Appearance      =   0  '2D
            Caption         =   "Allow changes to the Assembly"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3100
            TabIndex        =   44
            Top             =   525
            Width           =   2760
         End
      End
      Begin PDFCreator.dmFrame dmFraPDFPermissions 
         Height          =   855
         Left            =   120
         TabIndex        =   35
         Top             =   2880
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
            TabIndex        =   36
            Top             =   300
            Width           =   2865
         End
         Begin VB.CheckBox chkAllowCopy 
            Appearance      =   0  '2D
            Caption         =   "copy text and images"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   525
            Width           =   2865
         End
         Begin VB.CheckBox chkAllowModifyContents 
            Appearance      =   0  '2D
            Caption         =   "modify the document"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3100
            TabIndex        =   37
            Top             =   300
            Width           =   2760
         End
         Begin VB.CheckBox chkAllowModifyAnnotations 
            Appearance      =   0  '2D
            Caption         =   "modify comments"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3100
            TabIndex        =   39
            Top             =   525
            Width           =   2760
         End
      End
      Begin PDFCreator.dmFrame dmFraSecurityPass 
         Height          =   855
         Left            =   120
         TabIndex        =   32
         Top             =   1920
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
            TabIndex        =   33
            Top             =   300
            Width           =   5700
         End
         Begin VB.CheckBox chkOwnerPass 
            Appearance      =   0  '2D
            Caption         =   "Password required to change Permissions and Passwords"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   525
            Width           =   5700
         End
      End
      Begin PDFCreator.dmFrame dmFraPDFEncLevel 
         Height          =   1095
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   1931
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
         Begin VB.OptionButton optEncAes128 
            Appearance      =   0  '2D
            Caption         =   "Very high (AES 128 Bit - Adobe Acrobat 7.0 and above)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   98
            Top             =   750
            Width           =   5775
         End
         Begin VB.OptionButton optEncLow 
            Appearance      =   0  '2D
            Caption         =   "Low (40 Bit - Adobe Acrobat 3.0 and above)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   300
            Width           =   5775
         End
         Begin VB.OptionButton optEncHigh 
            Appearance      =   0  '2D
            Caption         =   "High (128 Bit - Adobe Acrobat 5.0 and above)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   525
            Width           =   5775
         End
      End
   End
   Begin PDFCreator.dmFrame dmFraPDFFonts 
      Height          =   1695
      Left            =   120
      TabIndex        =   15
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
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         Height          =   285
         Left            =   400
         TabIndex        =   18
         Top             =   1320
         Width           =   495
      End
      Begin VB.CheckBox chkPDFSubSetFonts 
         Appearance      =   0  '2D
         Caption         =   "Subset Fonts, when percentage of used characters below:"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   780
         Width           =   5955
      End
      Begin VB.CheckBox chkPDFEmbedAll 
         Appearance      =   0  '2D
         Caption         =   "Embed all Fonts"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   5955
      End
      Begin VB.Label lblPDFPerc 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   960
         TabIndex        =   19
         Top             =   1365
         Width           =   120
      End
   End
   Begin PDFCreator.dmFrame dmfraPDFCompress 
      Height          =   5535
      Left            =   6600
      TabIndex        =   13
      Top             =   120
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   9763
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
      Begin PDFCreator.dmFrame dmFraPDFMono 
         Height          =   1815
         Left            =   720
         TabIndex        =   88
         Top             =   3240
         Visible         =   0   'False
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   3201
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
            Alignment       =   1  'Rechts
            Appearance      =   0  '2D
            Height          =   285
            Left            =   2880
            TabIndex        =   93
            Top             =   1380
            Width           =   735
         End
         Begin VB.ComboBox cmbPDFMonoResample 
            Appearance      =   0  '2D
            Height          =   315
            ItemData        =   "ctlOptFormatPDF.ctx":0316
            Left            =   120
            List            =   "ctlOptFormatPDF.ctx":0318
            Style           =   2  'Dropdown-Liste
            TabIndex        =   92
            Tag             =   "Bicubic|Subsample|Average"
            Top             =   1380
            Width           =   2610
         End
         Begin VB.CheckBox chkPDFMonoResample 
            Appearance      =   0  '2D
            Caption         =   "Resample"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   1080
            Width           =   2610
         End
         Begin VB.ComboBox cmbPDFMonoComp 
            Appearance      =   0  '2D
            Height          =   315
            ItemData        =   "ctlOptFormatPDF.ctx":031A
            Left            =   120
            List            =   "ctlOptFormatPDF.ctx":031C
            Style           =   2  'Dropdown-Liste
            TabIndex        =   90
            Top             =   660
            Width           =   2610
         End
         Begin VB.CheckBox chkPDFMonoComp 
            Appearance      =   0  '2D
            Caption         =   "Compress"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   89
            Top             =   360
            Width           =   2610
         End
         Begin VB.Label lblPDFMonoRes 
            AutoSize        =   -1  'True
            Caption         =   "Resolution"
            Height          =   195
            Left            =   2880
            TabIndex        =   94
            Top             =   1080
            Width           =   750
         End
      End
      Begin PDFCreator.dmFrame dmFraPDFGrey 
         Height          =   1815
         Left            =   360
         TabIndex        =   81
         Top             =   2400
         Visible         =   0   'False
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   3201
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
         Begin VB.TextBox txtGreyCompressionFactor 
            Alignment       =   1  'Rechts
            Appearance      =   0  '2D
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   96
            Top             =   660
            Width           =   735
         End
         Begin VB.CheckBox chkPDFGreyComp 
            Appearance      =   0  '2D
            Caption         =   "Compress"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   86
            Top             =   360
            Width           =   2610
         End
         Begin VB.ComboBox cmbPDFGreyComp 
            Appearance      =   0  '2D
            Height          =   315
            ItemData        =   "ctlOptFormatPDF.ctx":031E
            Left            =   120
            List            =   "ctlOptFormatPDF.ctx":0320
            Style           =   2  'Dropdown-Liste
            TabIndex        =   85
            Top             =   660
            Width           =   2610
         End
         Begin VB.CheckBox chkPDFGreyResample 
            Appearance      =   0  '2D
            Caption         =   "Resample"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   84
            Top             =   1080
            Width           =   2610
         End
         Begin VB.ComboBox cmbPDFGreyResample 
            Appearance      =   0  '2D
            Height          =   315
            ItemData        =   "ctlOptFormatPDF.ctx":0322
            Left            =   120
            List            =   "ctlOptFormatPDF.ctx":0324
            Style           =   2  'Dropdown-Liste
            TabIndex        =   83
            Tag             =   "Bicubic|Subsample|Average"
            Top             =   1380
            Width           =   2610
         End
         Begin VB.TextBox txtPDFGreyRes 
            Alignment       =   1  'Rechts
            Appearance      =   0  '2D
            Height          =   285
            Left            =   2880
            TabIndex        =   82
            Top             =   1380
            Width           =   735
         End
         Begin VB.Label lblPDFGreyCompFac 
            AutoSize        =   -1  'True
            Caption         =   "Factor"
            Height          =   195
            Left            =   2880
            TabIndex        =   97
            Top             =   360
            Width           =   450
         End
         Begin VB.Label lblPDFGreyRes 
            AutoSize        =   -1  'True
            Caption         =   "Resolution"
            Height          =   195
            Left            =   2880
            TabIndex        =   87
            Top             =   1080
            Width           =   750
         End
      End
      Begin PDFCreator.dmFrame dmFraPDFColor 
         Height          =   1815
         Left            =   240
         TabIndex        =   72
         Top             =   1200
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   3201
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
         Begin VB.TextBox txtColorCompressionFactor 
            Alignment       =   1  'Rechts
            Appearance      =   0  '2D
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   79
            Top             =   660
            Width           =   735
         End
         Begin VB.CheckBox chkPDFColorResample 
            Appearance      =   0  '2D
            Caption         =   "Resample"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   75
            Top             =   1080
            Width           =   2610
         End
         Begin VB.CheckBox chkPDFColorComp 
            Appearance      =   0  '2D
            Caption         =   "Compress"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   360
            Width           =   2610
         End
         Begin VB.ComboBox cmbPDFColorComp 
            Appearance      =   0  '2D
            Height          =   315
            ItemData        =   "ctlOptFormatPDF.ctx":0326
            Left            =   120
            List            =   "ctlOptFormatPDF.ctx":0328
            Style           =   2  'Dropdown-Liste
            TabIndex        =   76
            Top             =   660
            Width           =   2610
         End
         Begin VB.ComboBox cmbPDFColorResample 
            Appearance      =   0  '2D
            Height          =   315
            ItemData        =   "ctlOptFormatPDF.ctx":032A
            Left            =   120
            List            =   "ctlOptFormatPDF.ctx":032C
            Style           =   2  'Dropdown-Liste
            TabIndex        =   74
            Tag             =   "Bicubic|Subsample|Average"
            Top             =   1380
            Width           =   2610
         End
         Begin VB.TextBox txtPDFColorRes 
            Alignment       =   1  'Rechts
            Appearance      =   0  '2D
            Height          =   285
            Left            =   2880
            TabIndex        =   73
            Top             =   1380
            Width           =   735
         End
         Begin VB.Label lblPDFColorCompFac 
            AutoSize        =   -1  'True
            Caption         =   "Factor"
            Height          =   195
            Left            =   2880
            TabIndex        =   80
            Top             =   360
            Width           =   450
         End
         Begin VB.Label lblPDFColorRes 
            AutoSize        =   -1  'True
            Caption         =   "Resolution"
            Height          =   195
            Left            =   2880
            TabIndex        =   78
            Top             =   1080
            Width           =   750
         End
      End
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
      Begin MSComctlLib.TabStrip tbstrPDFImageCompression 
         Height          =   2415
         Left            =   120
         TabIndex        =   95
         Top             =   720
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   4260
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Color images"
               Key             =   "ColorImages"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Greyscale images"
               Key             =   "GreyscaleImages"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Monochrome images"
               Key             =   "MonochromeImages"
               ImageVarType    =   2
            EndProperty
         EndProperty
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
         ItemData        =   "ctlOptFormatPDF.ctx":032E
         Left            =   120
         List            =   "ctlOptFormatPDF.ctx":0330
         Style           =   2  'Dropdown-Liste
         TabIndex        =   68
         Top             =   555
         Width           =   2655
      End
      Begin VB.ComboBox cmbPDFRotate 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "ctlOptFormatPDF.ctx":0332
         Left            =   120
         List            =   "ctlOptFormatPDF.ctx":0334
         Style           =   2  'Dropdown-Liste
         TabIndex        =   5
         Tag             =   "None|All|PageByPage"
         Top             =   1950
         Width           =   2655
      End
      Begin VB.ComboBox cmbPDFCompat 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "ctlOptFormatPDF.ctx":0336
         Left            =   120
         List            =   "ctlOptFormatPDF.ctx":0338
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
         ItemData        =   "ctlOptFormatPDF.ctx":033A
         Left            =   120
         List            =   "ctlOptFormatPDF.ctx":033C
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
         TabIndex        =   69
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
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   10610
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

Private ControlsEnabled As Boolean

Public Sub SetControlsEnabled(value As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ControlsEnabled = value
50020
50030  lblPDFDefaultSettings.Enabled = value
50040  cmbPDFDefaultSettings.Enabled = value
50050  lblPDFCompat.Enabled = value
50060  cmbPDFCompat.Enabled = value
50070  lblPDFAutoRotate.Enabled = value
50080  cmbPDFRotate.Enabled = value
50090  lblPDFResolution.Enabled = value
50100  txtPDFRes.Enabled = value
50110  lblPDFDPI.Enabled = value
50120  lblPDFOverprint.Enabled = value
50130  cmbPDFOverprint.Enabled = value
50140  chkPDFASCII85.Enabled = value
50150  chkPDFOptimize.Enabled = value
50160  dmFraPDFGeneral.Enabled = value
50170
50180  chkPDFTextComp.Enabled = value
50190  chkPDFColorComp.Enabled = value
50200  dmFraPDFColor.Enabled = value
50210  chkPDFGreyComp.Enabled = value
50220  dmFraPDFGrey.Enabled = value
50230  chkPDFMonoComp.Enabled = value
50240  dmFraPDFMono.Enabled = value
50250  dmfraPDFCompress.Enabled = value
50260
50270  chkPDFEmbedAll.Enabled = value
50280  chkPDFSubSetFonts.Enabled = value
50290  dmFraPDFFonts.Enabled = value
50300
50310  cmbPDFColorModel.Enabled = value
50320  chkPDFCMYKtoRGB.Enabled = value
50330  dmFraPDFColors.Enabled = value
50340  chkPDFPreserveOverprint.Enabled = value
50350  chkPDFPreserveTransfer.Enabled = value
50360  chkPDFPreserveHalftone.Enabled = value
50370  dmFraPDFColorOptions.Enabled = value
50380
50390  chkUseSecurity.Enabled = value
50400  dmFraPDFSecurity.Enabled = value
50410
50420  chkSignPDF.Enabled = value
50430  dmFraPDFSigning.Enabled = value
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "SetControlsEnabled")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkPDFSubSetFonts_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkPDFSubSetFonts.value = 1 And ControlsEnabled Then
50020    txtPDFSubSetPerc.Enabled = True
50030    lblPDFPerc.Enabled = True
50040   Else
50050    txtPDFSubSetPerc.Enabled = False
50060    lblPDFPerc.Enabled = False
50070  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "chkPDFSubSetFonts_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub optEncAes128_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  UpdateSecurityFields
50020  cmbPDFCompat.ListIndex = 3
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "optEncAes128_Click")
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
50020  cmbPDFCompat.ListIndex = 2
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "optEncHigh_Click")
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
50020  cmbPDFCompat.ListIndex = 1
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "optEncLow_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtColorCompressionFactor_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  KeyAscii = AllowedKeypressChars(KeyAscii, ".,")
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "txtColorCompressionFactor_KeyPress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub UserControl_Initialize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ctl As Control, i As Long
50020  ControlsEnabled = True
50030
50040  tbstrPDFOptions.Left = 0
50050  tbstrPDFOptions.Top = 0
50060  tbstrPDFOptions.Height = dmFraPDFSigning.Height + 420
50070  UserControl.Height = tbstrPDFOptions.Height + 500
50080
50090  With tbstrPDFOptions.Tabs
50100   .Clear
50110   .Add , "General"
50120   .Add , "Compression"
50130   .Add , "Fonts"
50140   .Add , "Colors"
50150   .Add , "Security"
50160   .Add , "Signing"
50170  End With
50180  With cmbPDFDefaultSettings
50190   .Clear
50200   For i = 1 To 5
50210    .AddItem ""
50220   Next i
50230  End With
50240  With cmbPDFCompat
50250   .Clear
50260   For i = 1 To 4
50270    .AddItem ""
50280   Next i
50290  End With
50300  With cmbPDFRotate
50310   .Clear
50320   For i = 1 To 3
50330    .AddItem ""
50340   Next i
50350  End With
50360  With cmbPDFOverprint
50370   .Clear
50380   For i = 1 To 2
50390    .AddItem ""
50400   Next i
50410  End With
50420  With cmbPDFColorComp
50430   .Clear
50440   For i = 1 To 8
50450    .AddItem ""
50460   Next i
50470  End With
50480  With cmbPDFColorResample
50490   .Clear
50500   For i = 1 To 2
50510    .AddItem ""
50520   Next i
50530  End With
50540  With cmbPDFGreyComp
50550   .Clear
50560   For i = 1 To 8
50570    .AddItem ""
50580   Next i
50590  End With
50600  With cmbPDFGreyResample
50610   .Clear
50620   For i = 1 To 2
50630    .AddItem ""
50640   Next i
50650  End With
50660  With cmbPDFMonoComp
50670   .Clear
50680   For i = 1 To 3
50690    .AddItem ""
50700   Next i
50710  End With
50720  With cmbPDFMonoResample
50730   .Clear
50740   For i = 1 To 2
50750    .AddItem ""
50760   Next i
50770  End With
50780  With cmbPDFColorModel
50790   .Clear
50800   For i = 1 To 3
50810    .AddItem ""
50820   Next i
50830  End With
50840  txtPDFRes.Text = 600
50850  cmbPDFCompat.ListIndex = 1
50860  cmbPDFRotate.ListIndex = 0
50870  cmbPDFOverprint.ListIndex = 0
50880  chkPDFASCII85.value = 0
50890
50900  chkPDFTextComp.value = 1
50910
50920  chkPDFColorComp.value = 1
50930  chkPDFColorResample.value = 0
50940  cmbPDFColorComp.ListIndex = 0
50950  cmbPDFColorResample.ListIndex = 0
50960  txtPDFColorRes.Text = 300
50970
50980  chkPDFGreyComp.value = 1
50990  chkPDFGreyResample.value = 0
51000  cmbPDFGreyComp.ListIndex = 0
51010  cmbPDFGreyResample.ListIndex = 0
51020  txtPDFGreyRes.Text = 300
51030
51040  chkPDFMonoComp.value = 1
51050  chkPDFMonoResample.value = 0
51060  cmbPDFMonoComp.ListIndex = 0
51070  cmbPDFMonoResample.ListIndex = 0
51080  txtPDFMonoRes.Text = 1200
51090
51100  chkPDFEmbedAll.value = 1
51110  chkPDFSubSetFonts.value = 1
51120  txtPDFSubSetPerc.Text = 100
51130
51140  cmbPDFColorModel.ListIndex = 1
51150  chkPDFCMYKtoRGB.value = 1
51160  chkPDFPreserveOverprint.value = 1
51170  chkPDFPreserveTransfer.value = 1
51180  chkPDFPreserveHalftone.value = 0
51190
51200   If GhostScriptSecurity = False Then
51210    chkUseSecurity.value = 0
51220    chkUseSecurity.Enabled = False
51230    SecurityIsPossible = False
51240   End If
51250
51260  If Options.PDFAes128Encryption <> 0 Then
51270    optEncAes128.value = True
51280   ElseIf Options.PDFHighEncryption <> 0 Then
51290    optEncHigh.value = True
51300   Else
51310    optEncLow.value = True
51320  End If
51330
51340  With tbstrPDFOptions
51350   .Top = 50
51360   .Left = 0
51370  End With
51380
51390  dmFraPDFGrey.Left = dmFraPDFColor.Left
51400  dmFraPDFGrey.Top = dmFraPDFColor.Top
51410  dmFraPDFMono.Left = dmFraPDFColor.Left
51420  dmFraPDFMono.Top = dmFraPDFColor.Top
51430  tbstrPDFImageCompression.Tabs(1).Selected = True
51440  dmFraPDFGrey.Visible = False
51450  dmFraPDFMono.Visible = False
51460
51470  UpdateSecurityFields
51480
51490  tbstrPDFOptions.ZOrder 1
51500  tbstrPDFOptions_Click
51510
51520  SetFrames Options.OptionsDesign
51530
51540  SetFont
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

Public Sub SetFont()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With Options
50020   SetFontControls UserControl.Controls, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50030  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "SetFont")
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
50320   cmbPDFColorComp.List(6) = .OptionsPDFCompressionColorComp09
50330   cmbPDFColorComp.List(7) = .OptionsPDFCompressionColorComp07
50340
50350   cmbPDFColorResample.List(0) = .OptionsPDFCompressionColorResample01
50360   cmbPDFColorResample.List(1) = .OptionsPDFCompressionColorResample02
50370
50380   cmbPDFGreyComp.List(0) = .OptionsPDFCompressionGreyComp01
50390   cmbPDFGreyComp.List(1) = .OptionsPDFCompressionGreyComp02
50400   cmbPDFGreyComp.List(2) = .OptionsPDFCompressionGreyComp03
50410   cmbPDFGreyComp.List(3) = .OptionsPDFCompressionGreyComp04
50420   cmbPDFGreyComp.List(4) = .OptionsPDFCompressionGreyComp05
50430   cmbPDFGreyComp.List(5) = .OptionsPDFCompressionGreyComp06
50440   cmbPDFGreyComp.List(6) = .OptionsPDFCompressionGreyComp09
50450   cmbPDFGreyComp.List(7) = .OptionsPDFCompressionGreyComp07
50460
50470   cmbPDFGreyResample.List(0) = .OptionsPDFCompressionGreyResample01
50480   cmbPDFGreyResample.List(1) = .OptionsPDFCompressionGreyResample02
50490
50500   cmbPDFMonoComp.List(0) = .OptionsPDFCompressionMonoComp01
50510   cmbPDFMonoComp.List(1) = .OptionsPDFCompressionMonoComp02
50520   cmbPDFMonoComp.List(2) = .OptionsPDFCompressionMonoComp03
50530
50540   cmbPDFMonoResample.List(0) = .OptionsPDFCompressionMonoResample01
50550   cmbPDFMonoResample.List(1) = .OptionsPDFCompressionMonoResample02
50560
50570   cmbPDFColorModel.List(0) = .OptionsPDFColorsColorModel01
50580   cmbPDFColorModel.List(1) = .OptionsPDFColorsColorModel02
50590   cmbPDFColorModel.List(2) = .OptionsPDFColorsColorModel03
50600
50610   dmFraPDFGeneral.Caption = .OptionsPDFGeneralCaption
50620   chkPDFOptimize.Caption = .OptionsPDFOptimize
50630   lblPDFCompat.Caption = .OptionsPDFGeneralCompatibility
50640   lblPDFAutoRotate.Caption = .OptionsPDFGeneralAutorotate
50650   lblPDFResolution.Caption = .OptionsPDFGeneralResolution
50660   lblPDFOverprint.Caption = .OptionsPDFGeneralOverprint
50670   chkPDFASCII85.Caption = .OptionsPDFGeneralASCII85
50680
50690   tbstrPDFImageCompression.Tabs(1).Caption = .OptionsPDFCompressionColor
50700   tbstrPDFImageCompression.Tabs(2).Caption = .OptionsPDFCompressionGrey
50710   tbstrPDFImageCompression.Tabs(3).Caption = .OptionsPDFCompressionMono
50720
50730   dmfraPDFCompress.Caption = .OptionsPDFCompressionCaption
50740   chkPDFTextComp.Caption = .OptionsPDFCompressionTextComp
50750   dmFraPDFColor.Caption = .OptionsPDFCompressionColor
50760   chkPDFColorComp.Caption = .OptionsPDFCompressionColorComp
50770         lblPDFColorCompFac.Caption = .OptionsPDFCompressionColorCompFac
50780   chkPDFColorResample.Caption = .OptionsPDFCompressionColorResample
50790   lblPDFColorRes.Caption = .OptionsPDFCompressionColorRes
50800   dmFraPDFGrey.Caption = .OptionsPDFCompressionGrey
50810   chkPDFGreyComp.Caption = .OptionsPDFCompressionGreyComp
50820         lblPDFGreyCompFac.Caption = .OptionsPDFCompressionGreyCompFac
50830   chkPDFGreyResample.Caption = .OptionsPDFCompressionGreyResample
50840   lblPDFGreyRes.Caption = .OptionsPDFCompressionGreyRes
50850   dmFraPDFMono.Caption = .OptionsPDFCompressionMono
50860   chkPDFMonoComp.Caption = .OptionsPDFCompressionMonoComp
50870   chkPDFMonoResample.Caption = .OptionsPDFCompressionMonoResample
50880   lblPDFMonoRes.Caption = .OptionsPDFCompressionMonoRes
50890
50900   dmFraPDFFonts.Caption = .OptionsPDFFontsCaption
50910   chkPDFEmbedAll.Caption = .OptionsPDFFontsEmbedAll
50920   chkPDFSubSetFonts.Caption = .OptionsPDFFontsSubSetFonts
50930
50940   dmFraPDFColors.Caption = .OptionsPDFColorsCaption
50950   chkPDFCMYKtoRGB.Caption = .OptionsPDFColorsCMYKtoRGB
50960   dmFraPDFColorOptions.Caption = .OptionsPDFColorsColorOptions
50970   chkPDFPreserveOverprint.Caption = .OptionsPDFColorsPreserveOverprint
50980   chkPDFPreserveTransfer.Caption = .OptionsPDFColorsPreserveTransfer
50990   chkPDFPreserveHalftone.Caption = .OptionsPDFColorsPreserveHalftone
51000
51010   dmFraPDFSigning.Caption = .OptionsPDFSigningCaption
51020   dmFraPDFSecurity.Caption = .OptionsPDFSecurityCaption
51030   chkUseSecurity.Caption = .OptionsPDFUseSecurity
51040   dmFraPDFEncLevel.Caption = .OptionsPDFEncryptionLevel
51050   optEncHigh.Caption = .OptionsPDFEncryptionHigh
51060   optEncLow.Caption = .OptionsPDFEncryptionLow
51070   dmFraSecurityPass.Caption = .OptionsPDFPasswords
51080   chkUserPass.Caption = .OptionsPDFUserPass
51090   chkOwnerPass.Caption = .OptionsPDFOwnerPass
51100   dmFraPDFPermissions.Caption = .OptionsPDFDisallowUser
51110   dmFraPDFHighPermissions.Caption = .OptionsPDFEnhancedPermissions
51120   chkAllowPrinting.Caption = .OptionsPDFDisallowPrint
51130   chkAllowModifyContents.Caption = .OptionsPDFDisallowModify
51140   chkAllowCopy.Caption = .OptionsPDFDisallowCopy
51150   chkAllowModifyAnnotations.Caption = .OptionsPDFDisallowModifyComments
51160   chkAllowDegradedPrinting.Caption = .OptionsPDFAllowDegradedPrinting
51170   chkAllowFillIn.Caption = .OptionsPDFAllowFillIn
51180   chkAllowAssembly.Caption = .OptionsPDFAllowAssembly
51190   chkAllowScreenReaders.Caption = .OptionsPDFAllowScreenReaders
51200
51210   chkSignPDF.Caption = .OptionsPDFSigningSignPdfFile
51220   lblPFXFile.Caption = .OptionsPDFSigningCertificateFile
51230   lblSignatureReason.Caption = .OptionsPDFSigningSignatureReason
51240   lblSignatureContact.Caption = .OptionsPDFSigningSignatureContact
51250   lblSignatureLocation.Caption = .OptionsPDFSigningSignatureLocation
51260   dmFraSignaturePosition.Caption = .OptionsPDFSigningSignaturePosition
51270   chkSignatureVisible.Caption = .OptionsPDFSigningSignatureVisible
51280   lblSignatureOnPage.Caption = .OptionsPDFSigningSignatureOnPage
51290   lblLeftX.Caption = .OptionsPDFSigningSignaturePositionLeftX
51300   lblLeftY.Caption = .OptionsPDFSigningSignaturePositionLeftY
51310   lblRightX.Caption = .OptionsPDFSigningSignaturePositionRightX
51320   lblRightY.Caption = .OptionsPDFSigningSignaturePositionRightY
51330   chkMultiSignature.Caption = .OptionsPDFSigningSignatureMultiSignature
51340  End With
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
50310   chkPDFEmbedAll.value = .PDFFontsEmbedAll
50320   chkPDFSubSetFonts.value = .PDFFontsSubSetFonts
50330   chkPDFSubSetFonts_Click
50340   txtPDFSubSetPerc.Text = .PDFFontsSubSetFontsPercent
50350   chkPDFASCII85.value = .PDFGeneralASCII85
50360   cmbPDFRotate.ListIndex = .PDFGeneralAutorotate
50370   cmbPDFCompat.ListIndex = .PDFGeneralCompatibility
50380   cmbPDFDefaultSettings.ListIndex = .PDFGeneralDefault
50390   cmbPDFOverprint.ListIndex = .PDFGeneralOverprint
50400   txtPDFRes.Text = .PDFGeneralResolution
50410
50420   chkPDFOptimize.value = .PDFOptimize
50430   chkOwnerPass.value = .PDFOwnerPass
50440   chkUserPass.value = .PDFUserPass
50450   chkUseSecurity.value = .PDFUseSecurity
50460
50470   chkSignPDF.value = .PDFSigningSignPDF
50480   txtPFXfile.Text = .PDFSigningPFXFile
50490   txtSignatureReason.Text = .PDFSigningSignatureReason
50500   txtSignatureContact.Text = .PDFSigningSignatureContact
50510   txtSignatureLocation.Text = .PDFSigningSignatureLocation
50520
50530   chkSignatureVisible.value = .PDFSigningSignatureVisible
50540   txtSignatureOnPage.Text = .PDFSigningSignatureOnPage
50550   txtLeftX.Text = .PDFSigningSignatureLeftX
50560   txtLeftY.Text = .PDFSigningSignatureLeftY
50570   txtRightX.Text = .PDFSigningSignatureRightX
50580   txtRightY.Text = .PDFSigningSignatureRightY
50590   chkMultiSignature.value = .PDFSigningMultiSignature
50600
50610   If pdfforgeDllInstalled Then
50620    optEncAes128.Enabled = True
50630    If .PDFAes128Encryption = 1 Then
50640     optEncAes128.value = True
50650    ElseIf .PDFHighEncryption = 1 Then
50660     optEncHigh.value = True
50670    Else
50680     optEncLow.value = True
50690    End If
50700   Else
50710    optEncAes128.Enabled = False
50720    If .PDFAes128Encryption = 1 Or .PDFHighEncryption = 1 Then
50730     optEncHigh.value = True
50740    Else
50750     optEncLow.value = True
50760    End If
50770   End If
50780  End With
50790
50800  If chkSignPDF.value = 1 Then
50810    EnableControls True
50820   Else
50830    EnableControls False
50840  End If
50850  UpdateSecurityFields
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
50160    If cmbPDFColorComp.ListIndex = 6 Then
50170     If IsNumeric(txtColorCompressionFactor.Text) Then
50180      .PDFCompressionColorCompressionJPEGManualFactor = CStr(CDbl(txtColorCompressionFactor.Text))
50190     End If
50200    End If
50210   End If
50220   .PDFCompressionColorResample = Abs(chkPDFColorResample.value)
50230   If LenB(CStr(cmbPDFColorResample.ListIndex)) > 0 Then
50240    .PDFCompressionColorResampleChoice = cmbPDFColorResample.ListIndex
50250   End If
50260   If LenB(txtPDFColorRes.Text) > 0 Then
50270    .PDFCompressionColorResolution = txtPDFColorRes.Text
50280   End If
50290   .PDFCompressionGreyCompression = Abs(chkPDFGreyComp.value)
50300   If LenB(CStr(cmbPDFGreyComp.ListIndex)) > 0 Then
50310    .PDFCompressionGreyCompressionChoice = cmbPDFGreyComp.ListIndex
50320    If cmbPDFGreyComp.ListIndex = 6 Then
50330     If IsNumeric(txtGreyCompressionFactor.Text) Then
50340      .PDFCompressionGreyCompressionJPEGManualFactor = CStr(CDbl(txtGreyCompressionFactor.Text))
50350     End If
50360    End If
50370   End If
50380   .PDFCompressionGreyResample = Abs(chkPDFGreyResample.value)
50390   If LenB(CStr(cmbPDFGreyResample.ListIndex)) > 0 Then
50400    .PDFCompressionGreyResampleChoice = cmbPDFGreyResample.ListIndex
50410   End If
50420   If LenB(txtPDFGreyRes.Text) > 0 Then
50430    .PDFCompressionGreyResolution = txtPDFGreyRes.Text
50440   End If
50450   .PDFCompressionMonoCompression = Abs(chkPDFMonoComp.value)
50460   If LenB(CStr(cmbPDFMonoComp.ListIndex)) > 0 Then
50470    .PDFCompressionMonoCompressionChoice = cmbPDFMonoComp.ListIndex
50480   End If
50490   .PDFCompressionMonoResample = Abs(chkPDFMonoResample.value)
50500   If LenB(CStr(cmbPDFMonoResample.ListIndex)) > 0 Then
50510    .PDFCompressionMonoResampleChoice = cmbPDFMonoResample.ListIndex
50520   End If
50530   If LenB(txtPDFMonoRes.Text) > 0 Then
50540    .PDFCompressionMonoResolution = txtPDFMonoRes.Text
50550   End If
50560   .PDFCompressionTextCompression = Abs(chkPDFTextComp.value)
50570   .PDFDisallowCopy = Abs(chkAllowCopy.value)
50580   .PDFDisallowModifyAnnotations = Abs(chkAllowModifyAnnotations.value)
50590   .PDFDisallowModifyContents = Abs(chkAllowModifyContents.value)
50600   .PDFDisallowPrinting = Abs(chkAllowPrinting.value)
50610   .PDFFontsEmbedAll = Abs(chkPDFEmbedAll.value)
50620   .PDFFontsSubSetFonts = Abs(chkPDFSubSetFonts.value)
50630   If LenB(txtPDFSubSetPerc.Text) > 0 Then
50640    .PDFFontsSubSetFontsPercent = txtPDFSubSetPerc.Text
50650   End If
50660   .PDFGeneralASCII85 = Abs(chkPDFASCII85.value)
50670   If LenB(CStr(cmbPDFRotate.ListIndex)) > 0 Then
50680    .PDFGeneralAutorotate = cmbPDFRotate.ListIndex
50690   End If
50700   If LenB(CStr(cmbPDFCompat.ListIndex)) > 0 Then
50710    .PDFGeneralCompatibility = cmbPDFCompat.ListIndex
50720   End If
50730   If LenB(CStr(cmbPDFDefaultSettings.ListIndex)) > 0 Then
50740    .PDFGeneralDefault = cmbPDFDefaultSettings.ListIndex
50750   End If
50760   If LenB(CStr(cmbPDFOverprint.ListIndex)) > 0 Then
50770    .PDFGeneralOverprint = cmbPDFOverprint.ListIndex
50780   End If
50790   If LenB(txtPDFRes.Text) > 0 Then
50800    .PDFGeneralResolution = txtPDFRes.Text
50810   End If
50820   .PDFHighEncryption = Abs(optEncHigh.value)
50830   .PDFLowEncryption = Abs(optEncLow.value)
50840   .PDFAes128Encryption = Abs(optEncAes128.value)
50850
50860   .PDFOptimize = Abs(chkPDFOptimize.value)
50870   .PDFOwnerPass = Abs(chkOwnerPass.value)
50880   .PDFUserPass = Abs(chkUserPass.value)
50890   .PDFUseSecurity = Abs(chkUseSecurity.value)
50900
50910   .PDFSigningSignPDF = Abs(chkSignPDF.value)
50920   .PDFSigningPFXFile = txtPFXfile.Text
50930   .PDFSigningSignatureReason = txtSignatureReason.Text
50940   .PDFSigningSignatureContact = txtSignatureContact.Text
50950   .PDFSigningSignatureLocation = txtSignatureLocation.Text
50960
50970   .PDFSigningSignatureVisible = Abs(chkSignatureVisible.value)
50980   If LenB(txtSignatureOnPage.Text) > 0 Then
50990    .PDFSigningSignatureOnPage = txtSignatureOnPage.Text
51000   End If
51010   If LenB(txtLeftX.Text) > 0 Then
51020    .PDFSigningSignatureLeftX = txtLeftX.Text
51030   End If
51040   If LenB(txtLeftY.Text) > 0 Then
51050    .PDFSigningSignatureLeftY = txtLeftY.Text
51060   End If
51070   If LenB(txtRightX.Text) > 0 Then
51080    .PDFSigningSignatureRightX = txtRightX.Text
51090   End If
51100   If LenB(txtRightY.Text) > 0 Then
51110    .PDFSigningSignatureRightY = txtRightY.Text
51120   End If
51130   .PDFSigningMultiSignature = Abs(chkMultiSignature.value)
51140  End With
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

Private Sub SelectPDFImagesCompressionControl()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  dmFraPDFColor.Visible = False
50020  dmFraPDFGrey.Visible = False
50030  dmFraPDFMono.Visible = False
50041  Select Case tbstrPDFImageCompression.SelectedItem.Index
        Case 1:
50060    dmFraPDFColor.Visible = True
50070   Case 2:
50080    dmFraPDFGrey.Visible = True
50090   Case 3:
50100    dmFraPDFMono.Visible = True
50110  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "SelectPDFImagesCompressionControl")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub tbstrPDFImageCompression_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  SelectPDFImagesCompressionControl
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("ctlOptFormatPDF", "tbstrPDFImageCompression_Click")
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
50180    If ControlsEnabled Then
50190     dmFraPDFGeneral.Enabled = True
50200    End If
50210   Case 2:
50220    dmfraPDFCompress.Visible = True
50230    If ControlsEnabled Then
50240     dmfraPDFCompress.Enabled = True
50250    End If
50260    SelectPDFImagesCompressionControl
50270   Case 3:
50280    dmFraPDFFonts.Visible = True
50290    If ControlsEnabled Then
50300     dmFraPDFFonts.Enabled = True
50310    End If
50320   Case 4:
50330    dmFraPDFColors.Visible = True
50340    dmFraPDFColorOptions.Visible = True
50350    If ControlsEnabled Then
50360     dmFraPDFColors.Enabled = True
50370     dmFraPDFColorOptions.Enabled = True
50380    End If
50390   Case 5:
50400    dmFraPDFSecurity.Visible = True
50410    dmFraPDFEncLevel.Visible = True
50420    dmFraSecurityPass.Visible = True
50430    dmFraPDFPermissions.Visible = True
50440    dmFraPDFHighPermissions.Visible = True
50450    If ControlsEnabled Then
50460     dmFraPDFSecurity.Enabled = True
50470     dmFraPDFEncLevel.Enabled = True
50480     dmFraSecurityPass.Enabled = True
50490     dmFraPDFPermissions.Enabled = True
50500     dmFraPDFHighPermissions.Enabled = True
50510    End If
50520    UpdateSecurityFields
50530    If cmbPDFCompat.ListIndex < 2 Then
50540      optEncLow.Enabled = True
50550      optEncHigh.Enabled = False
50560      optEncAes128.Enabled = False
50570     Else
50580      optEncLow.Enabled = False
50590      optEncHigh.Enabled = True
50600      optEncAes128.Enabled = True
50610    End If
50620    UpdateSecurityFields
50630    If SecurityIsPossible = False Then
50640     MsgBox LanguageStrings.MessagesMsg19
50650    End If
50660   Case 6:
50670    dmFraPDFSigning.Visible = True
50680    dmFraPDFSigning.Enabled = True
50690    If DotNet20Installed = False Or pdfforgeDllInstalled = False Then
50700     chkSignPDF.Enabled = False
50710     EnableControls False
50720     MsgBox LanguageStrings.MessagesMsg39
50730    End If
50740  End Select
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
50010  If chkPDFColorComp.value = 1 And ControlsEnabled Then
50020    cmbPDFColorComp.Enabled = True
50030    If cmbPDFColorComp.ListIndex = 0 Then
50040      chkPDFColorResample.Enabled = False
50050      cmbPDFColorResample.Enabled = False
50060      lblPDFColorRes.Enabled = False
50070      txtPDFColorRes.Enabled = False
50080      txtColorCompressionFactor.Locked = True
50090      txtColorCompressionFactor.BackColor = &H8000000F
50100      txtColorCompressionFactor.Text = ""
50110     Else
50120      chkPDFColorResample.Enabled = True
50130      If chkPDFColorResample.value = 1 Then
50140        cmbPDFColorResample.Enabled = True
50150        lblPDFColorRes.Enabled = True
50160        txtPDFColorRes.Enabled = True
50170       Else
50180        cmbPDFColorResample.Enabled = False
50190        lblPDFColorRes.Enabled = False
50200        txtPDFColorRes.Enabled = False
50210      End If
50221      Select Case cmbPDFColorComp.ListIndex
            Case 1:
50240        txtColorCompressionFactor.Locked = True
50250        txtColorCompressionFactor.BackColor = &H8000000F
50260        txtColorCompressionFactor.Text = Options.PDFCompressionColorCompressionJPEGMaximumFactor
50270       Case 2:
50280        txtColorCompressionFactor.Locked = True
50290        txtColorCompressionFactor.BackColor = &H8000000F
50300        txtColorCompressionFactor.Text = Options.PDFCompressionColorCompressionJPEGHighFactor
50310       Case 3:
50320        txtColorCompressionFactor.Locked = True
50330        txtColorCompressionFactor.BackColor = &H8000000F
50340        txtColorCompressionFactor.Text = Options.PDFCompressionColorCompressionJPEGMediumFactor
50350       Case 4:
50360        txtColorCompressionFactor.Locked = True
50370        txtColorCompressionFactor.BackColor = &H8000000F
50380        txtColorCompressionFactor.Text = Options.PDFCompressionColorCompressionJPEGLowFactor
50390       Case 5:
50400        txtColorCompressionFactor.Locked = True
50410        txtColorCompressionFactor.BackColor = &H8000000F
50420        txtColorCompressionFactor.Text = Options.PDFCompressionColorCompressionJPEGMinimumFactor
50430       Case 6:
50440        txtColorCompressionFactor.Locked = False
50450        txtColorCompressionFactor.BackColor = &H80000005
50460        txtColorCompressionFactor.Text = Options.PDFCompressionColorCompressionJPEGManualFactor
50470       Case Else:
50480        txtColorCompressionFactor.Locked = True
50490        txtColorCompressionFactor.BackColor = &H8000000F
50500        txtColorCompressionFactor.Text = ""
50510      End Select
50520    End If
50530   Else
50540    cmbPDFColorComp.Enabled = False
50550    chkPDFColorResample.Enabled = False
50560    cmbPDFColorResample.Enabled = False
50570    lblPDFColorRes.Enabled = False
50580    txtPDFColorRes.Enabled = False
50590  End If
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
50010  If chkPDFGreyComp.value = 1 And ControlsEnabled Then
50020    cmbPDFGreyComp.Enabled = True
50030    If cmbPDFGreyComp.ListIndex = 0 Then
50040      chkPDFGreyResample.Enabled = False
50050      cmbPDFGreyResample.Enabled = False
50060      lblPDFGreyRes.Enabled = False
50070      txtPDFGreyRes.Enabled = False
50080      txtGreyCompressionFactor.Locked = True
50090      txtGreyCompressionFactor.BackColor = &H8000000F
50100      txtGreyCompressionFactor.Text = ""
50110     Else
50120      chkPDFGreyResample.Enabled = True
50130      If chkPDFGreyResample.value = 1 Then
50140        cmbPDFGreyResample.Enabled = True
50150        lblPDFGreyRes.Enabled = True
50160        txtPDFGreyRes.Enabled = True
50170       Else
50180        cmbPDFGreyResample.Enabled = False
50190        lblPDFGreyRes.Enabled = False
50200        txtPDFGreyRes.Enabled = False
50210      End If
50221      Select Case cmbPDFGreyComp.ListIndex
            Case 1:
50240        txtGreyCompressionFactor.Locked = True
50250        txtGreyCompressionFactor.BackColor = &H8000000F
50260        txtGreyCompressionFactor.Text = Options.PDFCompressionGreyCompressionJPEGMaximumFactor
50270       Case 2:
50280        txtGreyCompressionFactor.Locked = True
50290        txtGreyCompressionFactor.BackColor = &H8000000F
50300        txtGreyCompressionFactor.Text = Options.PDFCompressionGreyCompressionJPEGHighFactor
50310       Case 3:
50320        txtGreyCompressionFactor.Locked = True
50330        txtGreyCompressionFactor.BackColor = &H8000000F
50340        txtGreyCompressionFactor.Text = Options.PDFCompressionGreyCompressionJPEGMediumFactor
50350       Case 4:
50360        txtGreyCompressionFactor.Locked = True
50370        txtGreyCompressionFactor.BackColor = &H8000000F
50380        txtGreyCompressionFactor.Text = Options.PDFCompressionGreyCompressionJPEGLowFactor
50390       Case 5:
50400        txtGreyCompressionFactor.Locked = True
50410        txtGreyCompressionFactor.BackColor = &H8000000F
50420        txtGreyCompressionFactor.Text = Options.PDFCompressionGreyCompressionJPEGMinimumFactor
50430       Case 6:
50440        txtGreyCompressionFactor.Locked = False
50450        txtGreyCompressionFactor.BackColor = &H80000005
50460        txtGreyCompressionFactor.Text = Options.PDFCompressionGreyCompressionJPEGManualFactor
50470       Case Else:
50480        txtGreyCompressionFactor.Locked = True
50490        txtGreyCompressionFactor.BackColor = &H8000000F
50500        txtGreyCompressionFactor.Text = ""
50510      End Select
50520    End If
50530   Else
50540    cmbPDFGreyComp.Enabled = False
50550    chkPDFGreyResample.Enabled = False
50560    cmbPDFGreyResample.Enabled = False
50570    lblPDFGreyRes.Enabled = False
50580    txtPDFGreyRes.Enabled = False
50590  End If
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
50010  If chkPDFMonoComp.value = 1 And ControlsEnabled Then
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
50010  If chkUseSecurity.value = False Or ControlsEnabled = False Then
50020    dmFraPDFEncLevel.Enabled = False
50030    optEncAes128.Enabled = False
50040    optEncHigh.Enabled = False
50050    optEncLow.Enabled = False
50060
50070    dmFraSecurityPass.Enabled = False
50080    chkUserPass.Enabled = False
50090    chkOwnerPass.Enabled = False
50100
50110    dmFraPDFPermissions.Enabled = False
50120    chkAllowPrinting.Enabled = False
50130    chkAllowCopy.Enabled = False
50140    chkAllowModifyAnnotations.Enabled = False
50150    chkAllowModifyContents.Enabled = False
50160
50170    dmFraPDFHighPermissions.Enabled = False
50180    chkAllowDegradedPrinting.Enabled = False
50190    chkAllowFillIn.Enabled = False
50200    chkAllowScreenReaders.Enabled = False
50210    chkAllowAssembly.Enabled = False
50220   Else
50230    dmFraPDFEncLevel.Enabled = True
50240
50250    dmFraSecurityPass.Enabled = True
50260    chkUserPass.Enabled = True
50270    chkOwnerPass.Enabled = True
50280
50290    dmFraPDFPermissions.Enabled = True
50300    chkAllowPrinting.Enabled = True
50310    chkAllowCopy.Enabled = True
50320    chkAllowModifyAnnotations.Enabled = True
50330    chkAllowModifyContents.Enabled = True
50340    optEncLow.Enabled = True
50350    optEncHigh.Enabled = True
50360    If pdfforgeDllInstalled Then
50370      optEncAes128.Enabled = True
50380     Else
50390      optEncAes128.Enabled = False
50400    End If
50410
50420    If optEncLow.value = True Then
50430      chkAllowDegradedPrinting.Enabled = False
50440      chkAllowFillIn.Enabled = False
50450      chkAllowScreenReaders.Enabled = False
50460      chkAllowAssembly.Enabled = False
50470      dmFraPDFHighPermissions.Enabled = False
50480     Else
50490      chkAllowDegradedPrinting.Enabled = True
50500      chkAllowFillIn.Enabled = True
50510      chkAllowScreenReaders.Enabled = True
50520      chkAllowAssembly.Enabled = True
50530      dmFraPDFHighPermissions.Enabled = True
50540    End If
50550  End If
50560
50570  If chkOwnerPass.value = 0 And chkUserPass.value = 0 Then
50580   chkOwnerPass.value = 1
50590   Options.PDFOwnerPass = 1
50600  End If
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
50010  If chkSignPDF.value = 1 And ControlsEnabled Then
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
50100  txtSignatureLocation.Enabled = value
50110  dmFraSignaturePosition.Enabled = value
50120  chkSignatureVisible.Enabled = value
50130  If chkSignatureVisible.value = 1 Then
50140    EnableSignPositionControls True
50150   Else
50160    EnableSignPositionControls False
50170  End If
50180  chkMultiSignature.Enabled = value
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
50010  lblSignatureOnPage.Enabled = value
50020  txtSignatureOnPage.Enabled = value
50030  lblLeftX.Enabled = value
50040  txtLeftX.Enabled = value
50050  lblLeftY.Enabled = value
50060  txtLeftY.Enabled = value
50070  lblRightX.Enabled = value
50080  txtRightX.Enabled = value
50090  lblRightY.Enabled = value
50100  txtRightY.Enabled = value
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
50010  Dim res As Long, files As Collection
50020
50030  res = OpenFileDialog(files, "", LanguageStrings.OptionsPDFSigningPfxP12Files + " (*.pfx,*.p12)|*.pfx;*.p12|" + LanguageStrings.OptionsPDFSigningPfxFiles + " (*.pfx)|*pfx|" + LanguageStrings.OptionsPDFSigningP12Files + " (*.p12|*.p12", "*.pfx;*.p12", "C:\", LanguageStrings.OptionsPDFSigningChooseCertifcateFile, OFN_FILEMUSTEXIST Or OFN_EXPLORER Or OFN_LONGNAMES Or OFN_PATHMUSTEXIST, 0, 1)
50040  If res > 0 Then
50050   txtPFXfile.Text = files(1)
50060  End If
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
50020   optEncLow.value = True
50030  End If
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
