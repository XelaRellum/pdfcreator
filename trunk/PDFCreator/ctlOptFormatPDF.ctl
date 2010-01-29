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
      _extentx        =   10927
      _extenty        =   9763
      caption         =   "Signing"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "ctlOptFormatPDF.ctx":0312
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
         _extentx        =   10504
         _extenty        =   3413
         caption         =   "Signature position"
         barcolorfrom    =   16744576
         barcolorto      =   4194304
         font            =   "ctlOptFormatPDF.ctx":033E
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
      _extentx        =   10927
      _extenty        =   2566
      caption         =   "Options"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "ctlOptFormatPDF.ctx":036A
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
      _extentx        =   10927
      _extenty        =   2143
      caption         =   "Color options"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "ctlOptFormatPDF.ctx":0396
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
         ItemData        =   "ctlOptFormatPDF.ctx":03C2
         Left            =   120
         List            =   "ctlOptFormatPDF.ctx":03C4
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
      _extentx        =   10927
      _extenty        =   9763
      caption         =   "Security"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "ctlOptFormatPDF.ctx":03C6
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
         _extentx        =   10504
         _extenty        =   1508
         caption         =   "Enhanced permissions (128 Bit only)"
         barcolorfrom    =   16744576
         barcolorto      =   4194304
         font            =   "ctlOptFormatPDF.ctx":03F2
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
         _extentx        =   10504
         _extenty        =   1508
         caption         =   "Disallow user to"
         barcolorfrom    =   16744576
         barcolorto      =   4194304
         font            =   "ctlOptFormatPDF.ctx":041E
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
         _extentx        =   10504
         _extenty        =   1508
         caption         =   "Passwords"
         barcolorfrom    =   16744576
         barcolorto      =   4194304
         font            =   "ctlOptFormatPDF.ctx":044A
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
         _extentx        =   10504
         _extenty        =   1508
         caption         =   "Encryption level"
         barcolorfrom    =   16744576
         barcolorto      =   4194304
         font            =   "ctlOptFormatPDF.ctx":0476
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
         _extentx        =   10504
         _extenty        =   1508
         caption         =   "Encryptor"
         barcolorfrom    =   16744576
         barcolorto      =   4194304
         font            =   "ctlOptFormatPDF.ctx":04A2
         Begin VB.ComboBox cmbPDFEncryptor 
            Appearance      =   0  '2D
            Height          =   315
            ItemData        =   "ctlOptFormatPDF.ctx":04CE
            Left            =   120
            List            =   "ctlOptFormatPDF.ctx":04D0
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
      _extentx        =   10927
      _extenty        =   2990
      caption         =   "Font options"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "ctlOptFormatPDF.ctx":04D2
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
      _extentx        =   10927
      _extenty        =   7646
      caption         =   "Compression"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "ctlOptFormatPDF.ctx":04FE
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
         _extentx        =   10504
         _extenty        =   1931
         caption         =   "Monochrome images"
         barcolorfrom    =   16744576
         barcolorto      =   4194304
         font            =   "ctlOptFormatPDF.ctx":052A
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
            ItemData        =   "ctlOptFormatPDF.ctx":0556
            Left            =   2520
            List            =   "ctlOptFormatPDF.ctx":0558
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
            ItemData        =   "ctlOptFormatPDF.ctx":055A
            Left            =   120
            List            =   "ctlOptFormatPDF.ctx":055C
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
         _extentx        =   10504
         _extenty        =   1931
         caption         =   "Greyscale images"
         barcolorfrom    =   16744576
         barcolorto      =   4194304
         font            =   "ctlOptFormatPDF.ctx":055E
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
            ItemData        =   "ctlOptFormatPDF.ctx":058A
            Left            =   120
            List            =   "ctlOptFormatPDF.ctx":058C
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
            ItemData        =   "ctlOptFormatPDF.ctx":058E
            Left            =   2520
            List            =   "ctlOptFormatPDF.ctx":0590
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
         _extentx        =   10504
         _extenty        =   1931
         caption         =   "Color images"
         barcolorfrom    =   16744576
         barcolorto      =   4194304
         font            =   "ctlOptFormatPDF.ctx":0592
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
            ItemData        =   "ctlOptFormatPDF.ctx":05BE
            Left            =   120
            List            =   "ctlOptFormatPDF.ctx":05C0
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
            ItemData        =   "ctlOptFormatPDF.ctx":05C2
            Left            =   2520
            List            =   "ctlOptFormatPDF.ctx":05C4
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
      _extentx        =   10927
      _extenty        =   8546
      caption         =   "General Options"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "ctlOptFormatPDF.ctx":05C6
      Begin VB.ComboBox cmbPDFDefaultSettings 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "ctlOptFormatPDF.ctx":05F2
         Left            =   120
         List            =   "ctlOptFormatPDF.ctx":05F4
         Style           =   2  'Dropdown-Liste
         TabIndex        =   91
         Top             =   555
         Width           =   2655
      End
      Begin VB.ComboBox cmbPDFRotate 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "ctlOptFormatPDF.ctx":05F6
         Left            =   120
         List            =   "ctlOptFormatPDF.ctx":05F8
         Style           =   2  'Dropdown-Liste
         TabIndex        =   5
         Tag             =   "None|All|PageByPage"
         Top             =   1950
         Width           =   2655
      End
      Begin VB.ComboBox cmbPDFCompat 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "ctlOptFormatPDF.ctx":05FA
         Left            =   120
         List            =   "ctlOptFormatPDF.ctx":05FC
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
         ItemData        =   "ctlOptFormatPDF.ctx":05FE
         Left            =   120
         List            =   "ctlOptFormatPDF.ctx":0600
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
 Dim ctl As Control
 Dim i As Long
 tbstrPDFOptions.Left = 0
 tbstrPDFOptions.Top = 0
 tbstrPDFOptions.Height = dmFraPDFSecurity.Height + 420
 UserControl.Height = tbstrPDFOptions.Height + 500

 With tbstrPDFOptions.Tabs
  .Clear
  .Add , "General"
  .Add , "Compression"
  .Add , "Fonts"
  .Add , "Colors"
  .Add , "Security"
  .Add , "Signing"
 End With
 With cmbPDFDefaultSettings
  .Clear
  For i = 1 To 5
   .AddItem ""
  Next i
 End With
 With cmbPDFCompat
  .Clear
  For i = 1 To 4
   .AddItem ""
  Next i
 End With
 With cmbPDFRotate
  .Clear
  For i = 1 To 3
   .AddItem ""
  Next i
 End With
 With cmbPDFOverprint
  .Clear
  For i = 1 To 2
   .AddItem ""
  Next i
 End With
 With cmbPDFColorComp
  .Clear
  For i = 1 To 7
   .AddItem ""
  Next i
 End With
 With cmbPDFColorResample
  .Clear
  For i = 1 To 2
   .AddItem ""
  Next i
 End With
 With cmbPDFGreyComp
  .Clear
  For i = 1 To 7
   .AddItem ""
  Next i
 End With
 With cmbPDFGreyResample
  .Clear
  For i = 1 To 2
   .AddItem ""
  Next i
 End With
 With cmbPDFMonoComp
  .Clear
  For i = 1 To 3
   .AddItem ""
  Next i
 End With
 With cmbPDFMonoResample
  .Clear
  For i = 1 To 2
   .AddItem ""
  Next i
 End With
 With cmbPDFColorModel
  .Clear
  For i = 1 To 3
   .AddItem ""
  Next i
 End With
 txtPDFRes.Text = 600
 cmbPDFCompat.ListIndex = 1
 cmbPDFRotate.ListIndex = 0
 cmbPDFOverprint.ListIndex = 0
 chkPDFASCII85.value = 0

 chkPDFTextComp.value = 1

 chkPDFColorComp.value = 1
 chkPDFColorResample.value = 0
 cmbPDFColorComp.ListIndex = 0
 cmbPDFColorResample.ListIndex = 0
 txtPDFColorRes.Text = 300

 chkPDFGreyComp.value = 1
 chkPDFGreyResample.value = 0
 cmbPDFGreyComp.ListIndex = 0
 cmbPDFGreyResample.ListIndex = 0
 txtPDFGreyRes.Text = 300

 chkPDFMonoComp.value = 1
 chkPDFMonoResample.value = 0
 cmbPDFMonoComp.ListIndex = 0
 cmbPDFMonoResample.ListIndex = 0
 txtPDFMonoRes.Text = 1200

 chkPDFEmbedAll.value = 1
 chkPDFSubSetFonts.value = 1
 txtPDFSubSetPerc.Text = 100

 cmbPDFColorModel.ListIndex = 1
 chkPDFCMYKtoRGB.value = 1
 chkPDFPreserveOverprint.value = 1
 chkPDFPreserveTransfer.value = 1
 chkPDFPreserveHalftone.value = 0

 With cmbPDFEncryptor
  .Clear
  .AddItem "Ghostscript (>= 8.14)"
  .ItemData(.NewIndex) = 0
  .AddItem "PDFEnc"
  .ItemData(.NewIndex) = 1

  SecurityIsPossible = True

  If FileExists(GetPDFCreatorApplicationPath & "pdfenc.exe") = False Then
   .RemoveItem 1
   .ListIndex = 0
   Options.PDFEncryptor = .ItemData(.ListIndex)
  End If
  If GhostScriptSecurity = False Then
   .RemoveItem 0
  End If
  If .ListCount = 0 Then
    chkUseSecurity.value = 0
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
   optEncHigh.value = True
  Else
   optEncLow.value = True
 End If

  With tbstrPDFOptions
  .Top = 50
  .Left = 0
 End With

 UpdateSecurityFields

 tbstrPDFOptions.ZOrder 1
 tbstrPDFOptions_Click

 SetFrames Options.OptionsDesign
End Sub

Public Sub SetFrames(OptionsDesign As Long)
 Dim ctl As Control
 For Each ctl In UserControl.Controls
  If TypeOf ctl Is dmFrame Then
   SetFrame ctl, OptionsDesign
  End If
 Next ctl
End Sub

Private Sub UserControl_Resize()
 tbstrPDFOptions.Width = UserControl.Width
 With dmFraPDFGeneral
  .Top = tbstrPDFOptions.ClientTop + 30
  .Left = tbstrPDFOptions.Left + (tbstrPDFOptions.Width - .Width) / 2
  dmfraPDFCompress.Top = .Top
  dmfraPDFCompress.Left = .Left
  dmFraPDFFonts.Top = .Top
  dmFraPDFFonts.Left = .Left
  dmFraPDFColors.Top = .Top
  dmFraPDFColors.Left = .Left
  dmFraPDFColorOptions.Top = dmFraPDFColors.Top + dmFraPDFColors.Height + 50
  dmFraPDFColorOptions.Left = .Left
  dmFraPDFSecurity.Top = .Top
  dmFraPDFSecurity.Left = .Left
  dmFraPDFSigning.Top = .Top
  dmFraPDFSigning.Left = .Left
 End With
End Sub

Public Sub SetLanguageStrings()
 With LanguageStrings
  tbstrPDFOptions.Tabs(1).Caption = .OptionsPDFGeneral
  tbstrPDFOptions.Tabs(2).Caption = .OptionsPDFCompression
  tbstrPDFOptions.Tabs(3).Caption = .OptionsPDFFonts
  tbstrPDFOptions.Tabs(4).Caption = .OptionsPDFColors
  tbstrPDFOptions.Tabs(5).Caption = .OptionsPDFSecurity
  tbstrPDFOptions.Tabs(6).Caption = .OptionsPDFSigning

  lblPDFDefaultSettings.Caption = .OptionsPDFGeneralDefaultSettings
  cmbPDFDefaultSettings.List(0) = .OptionsPDFGeneralDefaultSettingsDefault
  cmbPDFDefaultSettings.List(1) = .OptionsPDFGeneralDefaultSettingsScreen
  cmbPDFDefaultSettings.List(2) = .OptionsPDFGeneralDefaultSettingsEbook
  cmbPDFDefaultSettings.List(3) = .OptionsPDFGeneralDefaultSettingsPrinter
  cmbPDFDefaultSettings.List(4) = .OptionsPDFGeneralDefaultSettingsPrepress

  cmbPDFCompat.List(0) = .OptionsPDFGeneralCompatibility01
  cmbPDFCompat.List(1) = .OptionsPDFGeneralCompatibility02
  cmbPDFCompat.List(2) = .OptionsPDFGeneralCompatibility03
  cmbPDFCompat.List(3) = .OptionsPDFGeneralCompatibility04
  cmbPDFRotate.List(0) = .OptionsPDFGeneralRotate01
  cmbPDFRotate.List(1) = .OptionsPDFGeneralRotate02
  cmbPDFRotate.List(2) = .OptionsPDFGeneralRotate03
  cmbPDFOverprint.List(0) = .OptionsPDFGeneralOverprint01
  cmbPDFOverprint.List(1) = .OptionsPDFGeneralOverprint02

  cmbPDFColorComp.List(0) = .OptionsPDFCompressionColorComp01
  cmbPDFColorComp.List(1) = .OptionsPDFCompressionColorComp02
  cmbPDFColorComp.List(2) = .OptionsPDFCompressionColorComp03
  cmbPDFColorComp.List(3) = .OptionsPDFCompressionColorComp04
  cmbPDFColorComp.List(4) = .OptionsPDFCompressionColorComp05
  cmbPDFColorComp.List(5) = .OptionsPDFCompressionColorComp06
  cmbPDFColorComp.List(6) = .OptionsPDFCompressionColorComp07

  cmbPDFColorResample.List(0) = .OptionsPDFCompressionColorResample01
  cmbPDFColorResample.List(1) = .OptionsPDFCompressionColorResample02

  cmbPDFGreyComp.List(0) = .OptionsPDFCompressionGreyComp01
  cmbPDFGreyComp.List(1) = .OptionsPDFCompressionGreyComp02
  cmbPDFGreyComp.List(2) = .OptionsPDFCompressionGreyComp03
  cmbPDFGreyComp.List(3) = .OptionsPDFCompressionGreyComp04
  cmbPDFGreyComp.List(4) = .OptionsPDFCompressionGreyComp05
  cmbPDFGreyComp.List(5) = .OptionsPDFCompressionGreyComp06
  cmbPDFGreyComp.List(6) = .OptionsPDFCompressionGreyComp07

  cmbPDFGreyResample.List(0) = .OptionsPDFCompressionGreyResample01
  cmbPDFGreyResample.List(1) = .OptionsPDFCompressionGreyResample02

  cmbPDFMonoComp.List(0) = .OptionsPDFCompressionMonoComp01
  cmbPDFMonoComp.List(1) = .OptionsPDFCompressionMonoComp02
  cmbPDFMonoComp.List(2) = .OptionsPDFCompressionMonoComp03

  cmbPDFMonoResample.List(0) = .OptionsPDFCompressionMonoResample01
  cmbPDFMonoResample.List(1) = .OptionsPDFCompressionMonoResample02

  cmbPDFColorModel.List(0) = .OptionsPDFColorsColorModel01
  cmbPDFColorModel.List(1) = .OptionsPDFColorsColorModel02
  cmbPDFColorModel.List(2) = .OptionsPDFColorsColorModel03

  dmFraPDFGeneral.Caption = .OptionsPDFGeneralCaption
  chkPDFOptimize.Caption = .OptionsPDFOptimize
  lblPDFCompat.Caption = .OptionsPDFGeneralCompatibility
  lblPDFAutoRotate.Caption = .OptionsPDFGeneralAutorotate
  lblPDFResolution.Caption = .OptionsPDFGeneralResolution
  lblPDFOverprint.Caption = .OptionsPDFGeneralOverprint
  chkPDFASCII85.Caption = .OptionsPDFGeneralASCII85

  dmfraPDFCompress.Caption = .OptionsPDFCompressionCaption
  chkPDFTextComp.Caption = .OptionsPDFCompressionTextComp
  dmFraPDFColor.Caption = .OptionsPDFCompressionColor
  chkPDFColorComp.Caption = .OptionsPDFCompressionColorComp
  chkPDFColorResample.Caption = .OptionsPDFCompressionColorResample
  lblPDFColorRes.Caption = .OptionsPDFCompressionColorRes
  dmFraPDFGrey.Caption = .OptionsPDFCompressionGrey
  chkPDFGreyComp.Caption = .OptionsPDFCompressionGreyComp
  chkPDFGreyResample.Caption = .OptionsPDFCompressionGreyResample
  lblPDFGreyRes.Caption = .OptionsPDFCompressionGreyRes
  dmFraPDFMono.Caption = .OptionsPDFCompressionMono
  chkPDFMonoComp.Caption = .OptionsPDFCompressionMonoComp
  chkPDFMonoResample.Caption = .OptionsPDFCompressionMonoResample
  lblPDFMonoRes.Caption = .OptionsPDFCompressionMonoRes

  dmFraPDFFonts.Caption = .OptionsPDFFontsCaption
  chkPDFEmbedAll.Caption = .OptionsPDFFontsEmbedAll
  chkPDFSubSetFonts.Caption = .OptionsPDFFontsSubSetFonts

  dmFraPDFColors.Caption = .OptionsPDFColorsCaption
  chkPDFCMYKtoRGB.Caption = .OptionsPDFColorsCMYKtoRGB
  dmFraPDFColorOptions.Caption = .OptionsPDFColorsColorOptions
  chkPDFPreserveOverprint.Caption = .OptionsPDFColorsPreserveOverprint
  chkPDFPreserveTransfer.Caption = .OptionsPDFColorsPreserveTransfer
  chkPDFPreserveHalftone.Caption = .OptionsPDFColorsPreserveHalftone

  dmFraPDFSigning.Caption = .OptionsPDFSigningCaption
  dmFraPDFEncryptor.Caption = .OptionsPDFEncryptor
  dmFraPDFSecurity.Caption = .OptionsPDFSecurityCaption
  chkUseSecurity.Caption = .OptionsPDFUseSecurity
  dmFraPDFEncLevel.Caption = .OptionsPDFEncryptionLevel
  optEncHigh.Caption = .OptionsPDFEncryptionHigh
  optEncLow.Caption = .OptionsPDFEncryptionLow
  dmFraSecurityPass.Caption = .OptionsPDFPasswords
  chkUserPass.Caption = .OptionsPDFUserPass
  chkOwnerPass.Caption = .OptionsPDFOwnerPass
  dmFraPDFPermissions.Caption = .OptionsPDFDisallowUser
  dmFraPDFHighPermissions.Caption = .OptionsPDFEnhancedPermissions
  chkAllowPrinting.Caption = .OptionsPDFDisallowPrint
  chkAllowModifyContents.Caption = .OptionsPDFDisallowModify
  chkAllowCopy.Caption = .OptionsPDFDisallowCopy
  chkAllowModifyAnnotations.Caption = .OptionsPDFDisallowModifyComments
  chkAllowDegradedPrinting.Caption = .OptionsPDFAllowDegradedPrinting
  chkAllowFillIn.Caption = .OptionsPDFAllowFillIn
  chkAllowAssembly.Caption = .OptionsPDFAllowAssembly
  chkAllowScreenReaders.Caption = .OptionsPDFAllowScreenReaders

  chkSignPDF.Caption = .OptionsPDFSigningSignPdfFile
  lblPFXFile.Caption = .OptionsPDFSigningPfxFile
  lblSignatureReason.Caption = .OptionsPDFSigningSignatureReason
  lblSignatureContact.Caption = .OptionsPDFSigningSignatureContact
  lblSignatureLocation.Caption = .OptionsPDFSigningSignatureLocation
  dmFraSignaturePosition.Caption = .OptionsPDFSigningSignaturePosition
  chkSignatureVisible.Caption = .OptionsPDFSigningSignatureVisible
  lblSignatureOnPage.Caption = .OptionsPDFSigningSignatureOnPage
  lblLeftX.Caption = .OptionsPDFSigningSignaturePositionLeftX
  lblLeftY.Caption = .OptionsPDFSigningSignaturePositionLeftY
  lblRightX.Caption = .OptionsPDFSigningSignaturePositionRightX
  lblRightY.Caption = .OptionsPDFSigningSignaturePositionRightY
  chkMultiSignature.Caption = .OptionsPDFSigningSignatureMultiSignature
 End With
End Sub

Public Sub SetOptions()
 With Options1
  chkAllowAssembly.value = .PDFAllowAssembly
  chkAllowDegradedPrinting.value = .PDFAllowDegradedPrinting
  chkAllowFillIn.value = .PDFAllowFillIn
  chkAllowScreenReaders.value = .PDFAllowScreenReaders
  chkPDFCMYKtoRGB.value = .PDFColorsCMYKToRGB
  cmbPDFColorModel.ListIndex = .PDFColorsColorModel
  chkPDFPreserveHalftone.value = .PDFColorsPreserveHalftone
  chkPDFPreserveOverprint.value = .PDFColorsPreserveOverprint
  chkPDFPreserveTransfer.value = .PDFColorsPreserveTransfer
  chkPDFColorComp.value = .PDFCompressionColorCompression
  cmbPDFColorComp.ListIndex = .PDFCompressionColorCompressionChoice
  chkPDFColorResample.value = .PDFCompressionColorResample
  cmbPDFColorResample.ListIndex = .PDFCompressionColorResampleChoice
  txtPDFColorRes.Text = .PDFCompressionColorResolution
  chkPDFGreyComp.value = .PDFCompressionGreyCompression
  cmbPDFGreyComp.ListIndex = .PDFCompressionGreyCompressionChoice
  chkPDFGreyResample.value = .PDFCompressionGreyResample
  cmbPDFGreyResample.ListIndex = .PDFCompressionGreyResampleChoice
  txtPDFGreyRes.Text = .PDFCompressionGreyResolution
  chkPDFMonoComp.value = .PDFCompressionMonoCompression
  cmbPDFMonoComp.ListIndex = .PDFCompressionMonoCompressionChoice
  chkPDFMonoResample.value = .PDFCompressionMonoResample
  cmbPDFMonoResample.ListIndex = .PDFCompressionMonoResampleChoice
  txtPDFMonoRes.Text = .PDFCompressionMonoResolution
  chkPDFTextComp.value = .PDFCompressionTextCompression
  chkAllowCopy.value = .PDFDisallowCopy
  chkAllowModifyAnnotations.value = .PDFDisallowModifyAnnotations
  chkAllowModifyContents.value = .PDFDisallowModifyContents
  chkAllowPrinting.value = .PDFDisallowPrinting
  cmbPDFEncryptor.ItemData(cmbPDFEncryptor.ListIndex) = .PDFEncryptor
  chkPDFEmbedAll.value = .PDFFontsEmbedAll
  chkPDFSubSetFonts.value = .PDFFontsSubSetFonts
  txtPDFSubSetPerc.Text = .PDFFontsSubSetFontsPercent
  chkPDFASCII85.value = .PDFGeneralASCII85
  cmbPDFRotate.ListIndex = .PDFGeneralAutorotate
  cmbPDFCompat.ListIndex = .PDFGeneralCompatibility
  cmbPDFDefaultSettings.ListIndex = .PDFGeneralDefault
  cmbPDFOverprint.ListIndex = .PDFGeneralOverprint
  txtPDFRes.Text = .PDFGeneralResolution
'  optEncHigh.value = .PDFHighEncryption
'  optEncLow.value = .PDFLowEncryption
  chkPDFOptimize.value = .PDFOptimize
  chkOwnerPass.value = .PDFOwnerPass
  chkUserPass.value = .PDFUserPass
  chkUseSecurity.value = .PDFUseSecurity

  chkSignPDF.value = .PDFSigningSignPDF
  txtPFXfile.Text = .PDFSigningPFXFile
  txtSignatureReason.Text = .PDFSigningSignatureReason
  txtSignatureContact.Text = .PDFSigningSignatureContact
  txtSignatureLocation.Text = .PDFSigningSignatureLocation

  chkSignatureVisible.value = .PDFSigningSignatureVisible
  txtSignatureOnPage.Text = .PDFSigningSignatureOnPage
  txtLeftX.Text = .PDFSigningSignatureLeftX
  txtLeftY.Text = .PDFSigningSignatureLeftY
  txtRightX.Text = .PDFSigningSignatureRightX
  txtRightY.Text = .PDFSigningSignatureRightY
  chkMultiSignature.value = .PDFSigningMultiSignature
 End With
 If chkSignPDF.value = 1 Then
   EnableControls True
  Else
   EnableControls False
 End If
End Sub

Public Sub GetOptions()
 With Options1
  .PDFAllowAssembly = Abs(chkAllowAssembly.value)
  .PDFAllowDegradedPrinting = Abs(chkAllowDegradedPrinting.value)
  .PDFAllowFillIn = Abs(chkAllowFillIn.value)
  .PDFAllowScreenReaders = Abs(chkAllowScreenReaders.value)
  .PDFColorsCMYKToRGB = Abs(chkPDFCMYKtoRGB.value)
  If LenB(CStr(cmbPDFColorModel.ListIndex)) > 0 Then
   .PDFColorsColorModel = cmbPDFColorModel.ListIndex
  End If
  .PDFColorsPreserveHalftone = Abs(chkPDFPreserveHalftone.value)
  .PDFColorsPreserveOverprint = Abs(chkPDFPreserveOverprint.value)
  .PDFColorsPreserveTransfer = Abs(chkPDFPreserveTransfer.value)
  .PDFCompressionColorCompression = Abs(chkPDFColorComp.value)
  If LenB(CStr(cmbPDFColorComp.ListIndex)) > 0 Then
   .PDFCompressionColorCompressionChoice = cmbPDFColorComp.ListIndex
  End If
  .PDFCompressionColorResample = Abs(chkPDFColorResample.value)
  If LenB(CStr(cmbPDFColorResample.ListIndex)) > 0 Then
   .PDFCompressionColorResampleChoice = cmbPDFColorResample.ListIndex
  End If
  If LenB(txtPDFColorRes.Text) > 0 Then
   .PDFCompressionColorResolution = txtPDFColorRes.Text
  End If
  .PDFCompressionGreyCompression = Abs(chkPDFGreyComp.value)
  If LenB(CStr(cmbPDFGreyComp.ListIndex)) > 0 Then
   .PDFCompressionGreyCompressionChoice = cmbPDFGreyComp.ListIndex
  End If
  .PDFCompressionGreyResample = Abs(chkPDFGreyResample.value)
  If LenB(CStr(cmbPDFGreyResample.ListIndex)) > 0 Then
   .PDFCompressionGreyResampleChoice = cmbPDFGreyResample.ListIndex
  End If
  If LenB(txtPDFGreyRes.Text) > 0 Then
   .PDFCompressionGreyResolution = txtPDFGreyRes.Text
  End If
  .PDFCompressionMonoCompression = Abs(chkPDFMonoComp.value)
  If LenB(CStr(cmbPDFMonoComp.ListIndex)) > 0 Then
   .PDFCompressionMonoCompressionChoice = cmbPDFMonoComp.ListIndex
  End If
  .PDFCompressionMonoResample = Abs(chkPDFMonoResample.value)
  If LenB(CStr(cmbPDFMonoResample.ListIndex)) > 0 Then
   .PDFCompressionMonoResampleChoice = cmbPDFMonoResample.ListIndex
  End If
  If LenB(txtPDFMonoRes.Text) > 0 Then
   .PDFCompressionMonoResolution = txtPDFMonoRes.Text
  End If
  .PDFCompressionTextCompression = Abs(chkPDFTextComp.value)
  .PDFDisallowCopy = Abs(chkAllowCopy.value)
  .PDFDisallowModifyAnnotations = Abs(chkAllowModifyAnnotations.value)
  .PDFDisallowModifyContents = Abs(chkAllowModifyContents.value)
  .PDFDisallowPrinting = Abs(chkAllowPrinting.value)
  If cmbPDFEncryptor.ListIndex < 0 Then
    .PDFEncryptor = 0
   Else
    .PDFEncryptor = cmbPDFEncryptor.ItemData(cmbPDFEncryptor.ListIndex)
  End If
  .PDFFontsEmbedAll = Abs(chkPDFEmbedAll.value)
  .PDFFontsSubSetFonts = Abs(chkPDFSubSetFonts.value)
  If LenB(txtPDFSubSetPerc.Text) > 0 Then
   .PDFFontsSubSetFontsPercent = txtPDFSubSetPerc.Text
  End If
  .PDFGeneralASCII85 = Abs(chkPDFASCII85.value)
  If LenB(CStr(cmbPDFRotate.ListIndex)) > 0 Then
   .PDFGeneralAutorotate = cmbPDFRotate.ListIndex
  End If
  If LenB(CStr(cmbPDFCompat.ListIndex)) > 0 Then
   .PDFGeneralCompatibility = cmbPDFCompat.ListIndex
  End If
  If LenB(CStr(cmbPDFDefaultSettings.ListIndex)) > 0 Then
   .PDFGeneralDefault = cmbPDFDefaultSettings.ListIndex
  End If
  If LenB(CStr(cmbPDFOverprint.ListIndex)) > 0 Then
   .PDFGeneralOverprint = cmbPDFOverprint.ListIndex
  End If
  If LenB(txtPDFRes.Text) > 0 Then
   .PDFGeneralResolution = txtPDFRes.Text
  End If
  .PDFHighEncryption = Abs(optEncHigh.value)
  .PDFLowEncryption = Abs(optEncLow.value)
  .PDFOptimize = Abs(chkPDFOptimize.value)
  .PDFOwnerPass = Abs(chkOwnerPass.value)
  .PDFUserPass = Abs(chkUserPass.value)
  .PDFUseSecurity = Abs(chkUseSecurity.value)

  .PDFSigningSignPDF = Abs(chkSignPDF.value)
  .PDFSigningPFXFile = txtPFXfile.Text
  .PDFSigningSignatureReason = txtSignatureReason.Text
  .PDFSigningSignatureContact = txtSignatureContact.Text
  .PDFSigningSignatureLocation = txtSignatureLocation.Text

  .PDFSigningSignatureVisible = Abs(chkSignatureVisible.value)
  If LenB(txtSignatureOnPage.Text) > 0 Then
   .PDFSigningSignatureOnPage = txtSignatureOnPage.Text
  End If
  If LenB(txtLeftX.Text) > 0 Then
   .PDFSigningSignatureLeftX = txtLeftX.Text
  End If
  If LenB(txtLeftY.Text) > 0 Then
   .PDFSigningSignatureLeftY = txtLeftY.Text
  End If
  If LenB(txtRightX.Text) > 0 Then
   .PDFSigningSignatureRightX = txtRightX.Text
  End If
  If LenB(txtRightY.Text) > 0 Then
   .PDFSigningSignatureRightY = txtRightY.Text
  End If
  .PDFSigningMultiSignature = Abs(chkMultiSignature.value)
 End With
End Sub

Private Sub tbstrPDFOptions_Click()
 dmFraPDFGeneral.Visible = False
 dmfraPDFCompress.Visible = False
 dmFraPDFFonts.Visible = False
 dmFraPDFColors.Visible = False
 dmFraPDFColorOptions.Visible = False
 dmFraPDFSecurity.Visible = False
 dmFraPDFSigning.Visible = False
 dmFraPDFGeneral.Enabled = False
 dmfraPDFCompress.Enabled = False
 dmFraPDFFonts.Enabled = False
 dmFraPDFColors.Enabled = False
 dmFraPDFColorOptions.Enabled = False
 dmFraPDFSecurity.Enabled = False
 dmFraPDFSigning.Enabled = False
 Select Case tbstrPDFOptions.SelectedItem.Index
  Case 1:
   dmFraPDFGeneral.Visible = True
   dmFraPDFGeneral.Enabled = True
  Case 2:
   dmfraPDFCompress.Visible = True
   dmfraPDFCompress.Enabled = True
   dmFraPDFColor.Visible = True
   dmFraPDFColor.Enabled = True
   dmFraPDFGrey.Visible = True
   dmFraPDFGrey.Enabled = True
   dmFraPDFMono.Visible = True
   dmFraPDFMono.Enabled = True
  Case 3:
   dmFraPDFFonts.Visible = True
   dmFraPDFFonts.Enabled = True
  Case 4:
   dmFraPDFColors.Visible = True
   dmFraPDFColorOptions.Visible = True
   dmFraPDFColors.Enabled = True
   dmFraPDFColorOptions.Enabled = True
  Case 5:
   dmFraPDFSecurity.Visible = True
   dmFraPDFSecurity.Enabled = True
   dmFraPDFEncryptor.Visible = True
   dmFraPDFEncryptor.Enabled = True
   dmFraPDFEncLevel.Visible = True
   dmFraPDFEncLevel.Enabled = True
   dmFraSecurityPass.Visible = True
   dmFraSecurityPass.Enabled = True
   dmFraPDFPermissions.Visible = True
   dmFraPDFPermissions.Enabled = True
   dmFraPDFHighPermissions.Visible = True
   dmFraPDFHighPermissions.Enabled = True
   UpdateSecurityFields
   If cmbPDFCompat.ListIndex < 2 Then
     optEncLow.Enabled = True
     optEncHigh.Enabled = False
    Else
     optEncLow.Enabled = False
     optEncHigh.Enabled = True
   End If
   If SecurityIsPossible = False Then
    MsgBox LanguageStrings.MessagesMsg19
   End If
  Case 6:
   dmFraPDFSigning.Visible = True
   dmFraPDFSigning.Enabled = True
   If PDFSigningIsPossible = False Then
    chkSignPDF.Enabled = False
    EnableControls False
    MsgBox LanguageStrings.MessagesMsg39
   End If
 End Select
End Sub

Private Sub chkOwnerPass_Click()
 If chkUserPass.value = 0 Then
  If chkOwnerPass.value = 0 Then
   chkOwnerPass.value = 1
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

Private Sub SetPDFColorComprSettings()
 If chkPDFColorComp.value = 1 Then
   cmbPDFColorComp.Enabled = True
   If cmbPDFColorComp.ListIndex = 0 Then
     chkPDFColorResample.Enabled = False
     cmbPDFColorResample.Enabled = False
     lblPDFColorRes.Enabled = False
     txtPDFColorRes.Enabled = False
    Else
     chkPDFColorResample.Enabled = True
     If chkPDFColorResample.value = 1 Then
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
 If chkPDFGreyComp.value = 1 Then
   cmbPDFGreyComp.Enabled = True
   If cmbPDFGreyComp.ListIndex = 0 Then
     chkPDFGreyResample.Enabled = False
     cmbPDFGreyResample.Enabled = False
     lblPDFGreyRes.Enabled = False
     txtPDFGreyRes.Enabled = False
    Else
     chkPDFGreyResample.Enabled = True
     If chkPDFGreyResample.value = 1 Then
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
 If chkPDFMonoComp.value = 1 Then
   cmbPDFMonoComp.Enabled = True
   chkPDFMonoResample.Enabled = True
   If chkPDFMonoResample.value = 1 Then
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

Private Sub cmbPDFColorComp_Click()
 SetPDFColorComprSettings
End Sub

Private Sub cmbPDFGreyComp_Click()
 SetPDFGreyComprSettings
End Sub

Private Sub cmbPDFMonoComp_Click()
 SetPDFMonoComprSettings
End Sub

Private Sub UpdateSecurityFields()
 If chkUseSecurity.value = False Then
   dmFraPDFEncryptor.Enabled = False
   cmbPDFEncryptor.Enabled = False

   dmFraPDFEncLevel.Enabled = False
   optEncHigh.Enabled = False
   optEncLow.Enabled = False

   dmFraSecurityPass.Enabled = False
   chkUserPass.Enabled = False
   chkOwnerPass.Enabled = False

   dmFraPDFPermissions.Enabled = False
   chkAllowPrinting.Enabled = False
   chkAllowCopy.Enabled = False
   chkAllowModifyAnnotations.Enabled = False
   chkAllowModifyContents.Enabled = False

   dmFraPDFHighPermissions.Enabled = False
   chkAllowDegradedPrinting.Enabled = False
   chkAllowFillIn.Enabled = False
   chkAllowScreenReaders.Enabled = False
   chkAllowAssembly.Enabled = False
  Else
   dmFraPDFEncryptor.Enabled = True
   cmbPDFEncryptor.Enabled = True
   dmFraPDFEncLevel.Enabled = True

   dmFraSecurityPass.Enabled = True
   chkUserPass.Enabled = True
   chkOwnerPass.Enabled = True

   dmFraPDFPermissions.Enabled = True
   chkAllowPrinting.Enabled = True
   chkAllowCopy.Enabled = True
   chkAllowModifyAnnotations.Enabled = True
   chkAllowModifyContents.Enabled = True

   If cmbPDFCompat.ListIndex < 2 Then
     optEncLow.Enabled = True
     optEncHigh.Enabled = False
     optEncLow.value = True
     chkAllowDegradedPrinting.Enabled = False
     chkAllowFillIn.Enabled = False
     chkAllowScreenReaders.Enabled = False
     chkAllowAssembly.Enabled = False
     dmFraPDFHighPermissions.Enabled = False
    Else
     optEncLow.Enabled = False
     optEncHigh.Enabled = True
     optEncHigh.value = True
     chkAllowDegradedPrinting.Enabled = True
     chkAllowFillIn.Enabled = True
     chkAllowScreenReaders.Enabled = True
     chkAllowAssembly.Enabled = True
     dmFraPDFHighPermissions.Enabled = True
   End If
 End If

 If chkOwnerPass.value = 0 And chkUserPass.value = 0 Then
  chkOwnerPass.value = 1
  Options.PDFOwnerPass = 1
 End If
End Sub

Private Sub chkUserPass_Click()
 If chkOwnerPass.value = 0 Then
  If chkUserPass.value = 0 Then
   chkUserPass.value = 1
   chkOwnerPass.value = 1
  End If
  SavePasswordsForThisSession = False
 End If
End Sub

Private Sub chkUseSecurity_Click()
 UpdateSecurityFields
End Sub

Public Property Get PDFOptionsIndex()
 PDFOptionsIndex = tbstrPDFOptions.SelectedItem.Index
End Property

Private Sub chkSignPDF_Click()
 If chkSignPDF.value = 1 Then
   EnableControls True
  Else
   EnableControls False
 End If
End Sub

Private Sub EnableControls(value As Boolean)
 lblPFXFile.Enabled = value
 txtPFXfile.Enabled = value
 cmdGetPFXFile.Enabled = value
 txtPFXFilePreview.Enabled = value
 lblSignatureReason.Enabled = value
 txtSignatureReason.Enabled = value
 lblSignatureContact.Enabled = value
 txtSignatureContact.Enabled = value
 lblSignatureLocation.Enabled = value
 dmFraSignaturePosition.Enabled = value
 chkSignatureVisible.Enabled = value
 If chkSignatureVisible.value = 1 Then
   EnableSignPositionControls True
  Else
   EnableSignPositionControls False
 End If
 chkMultiSignature.Enabled = value
End Sub

Private Sub EnableSignPositionControls(value As Boolean)
 lblLeftX.Enabled = value
 txtLeftX.Enabled = value
 lblLeftY.Enabled = value
 txtLeftY.Enabled = value
 lblRightX.Enabled = value
 txtRightX.Enabled = value
 lblRightY.Enabled = value
 txtRightY.Enabled = value
End Sub

Private Sub chkSignatureVisible_Click()
 If chkSignatureVisible.value = 1 Then
   EnableSignPositionControls True
  Else
   EnableSignPositionControls False
 End If
End Sub

Private Sub cmdGetPFXFile_Click()
 Dim res As Long, files As Collection, certFilename As String
 With Options
  If LenB(.PDFSigningPFXFile) = 0 Then
    res = OpenFileDialog(files, "", "PFX\P12 files (*.pfx,*.p12)|*.pfx;*.p12|PFX files (*.pfx)|*pfx|P12 files (*.p12|*.p12", "*.pfx;*.p12", "C:\", "Choose a certificate", OFN_FILEMUSTEXIST Or OFN_EXPLORER Or OFN_LONGNAMES Or OFN_PATHMUSTEXIST, 0, 1)
    If res > 0 Then
     certFilename = files(1)
    End If
   Else
    certFilename = .PDFSigningPFXFile
  End If
  txtPFXfile.Text = certFilename
 End With
End Sub

Private Sub txtPFXfile_Change()
 txtPFXFilePreview.Text = txtPFXfile.Text
End Sub

Private Sub txtSignatureOnPage_KeyPress(KeyAscii As Integer)
 KeyAscii = AllowedKeypressChars(KeyAscii)
End Sub

Private Sub txtLeftX_KeyPress(KeyAscii As Integer)
 KeyAscii = AllowedKeypressChars(KeyAscii)
End Sub

Private Sub txtLeftY_KeyPress(KeyAscii As Integer)
 KeyAscii = AllowedKeypressChars(KeyAscii)
End Sub

Private Sub txtRightX_KeyPress(KeyAscii As Integer)
 KeyAscii = AllowedKeypressChars(KeyAscii)
End Sub

Private Sub txtRightY_KeyPress(KeyAscii As Integer)
 KeyAscii = AllowedKeypressChars(KeyAscii)
End Sub

Private Sub cmbPDFCompat_Click()
 If cmbPDFCompat.ListIndex < 2 Then
   optEncLow.value = True
  Else
   optEncHigh.value = True
 End If
End Sub
