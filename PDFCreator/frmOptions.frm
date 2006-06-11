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
   Begin PDFCreator.dmFrame dmFraProgDocument 
      Height          =   2250
      Left            =   2640
      TabIndex        =   45
      Top             =   1800
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3969
      Caption         =   "Document"
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
      Begin VB.CheckBox chkOnePagePerFile 
         Appearance      =   0  '2D
         Caption         =   "One page per file"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   224
         Top             =   1890
         Width           =   6000
      End
      Begin VB.ComboBox cmbAuthorTokens 
         Appearance      =   0  '2D
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmOptions.frx":000C
         Left            =   3720
         List            =   "frmOptions.frx":000E
         Style           =   2  'Dropdown-Liste
         TabIndex        =   50
         Top             =   840
         Width           =   2415
      End
      Begin VB.CheckBox chkUseStandardAuthor 
         Appearance      =   0  '2D
         Caption         =   "Use Standardauthor"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   360
         Width           =   6000
      End
      Begin VB.TextBox txtStandardAuthor 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   120
         TabIndex        =   48
         Top             =   840
         Width           =   3495
      End
      Begin VB.CheckBox chkUseCreationDateNow 
         Appearance      =   0  '2D
         Caption         =   "Use the current Date/Time for 'Creation Date'"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   1440
         Width           =   6000
      End
      Begin VB.Label lblAuthorTokens 
         AutoSize        =   -1  'True
         Caption         =   "Add a Author-Token"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3720
         TabIndex        =   46
         Top             =   600
         Width           =   1440
      End
   End
   Begin PDFCreator.dmFrame dmFraPDFGeneral 
      Height          =   4365
      Left            =   2625
      TabIndex        =   90
      Top             =   1890
      Visible         =   0   'False
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   7699
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
      Begin VB.CheckBox chkPDFOptimize 
         Appearance      =   0  '2D
         Caption         =   "Fast web view"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   223
         Top             =   3990
         Width           =   5880
      End
      Begin VB.CheckBox chkPDFASCII85 
         Appearance      =   0  '2D
         Caption         =   "Convert binary data to ASCII85"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   95
         Top             =   3570
         Width           =   5880
      End
      Begin VB.ComboBox cmbPDFOverprint 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "frmOptions.frx":0010
         Left            =   120
         List            =   "frmOptions.frx":0012
         Style           =   2  'Dropdown-Liste
         TabIndex        =   94
         Top             =   2940
         Width           =   2655
      End
      Begin VB.TextBox txtPDFRes 
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         Height          =   285
         Left            =   120
         TabIndex        =   93
         Text            =   "600"
         Top             =   2205
         Width           =   615
      End
      Begin VB.ComboBox cmbPDFCompat 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "frmOptions.frx":0014
         Left            =   120
         List            =   "frmOptions.frx":0016
         Style           =   2  'Dropdown-Liste
         TabIndex        =   92
         Top             =   735
         Width           =   2655
      End
      Begin VB.ComboBox cmbPDFRotate 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "frmOptions.frx":0018
         Left            =   120
         List            =   "frmOptions.frx":001A
         Style           =   2  'Dropdown-Liste
         TabIndex        =   91
         Tag             =   "None|All|PageByPage"
         Top             =   1470
         Width           =   2655
      End
      Begin VB.Label lblPDFDPI 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "dpi"
         Height          =   195
         Left            =   800
         TabIndex        =   100
         Top             =   2250
         Width           =   210
      End
      Begin VB.Label lblPDFOverprint 
         AutoSize        =   -1  'True
         Caption         =   "Overprint:"
         Height          =   195
         Left            =   120
         TabIndex        =   99
         Top             =   2730
         Width           =   690
      End
      Begin VB.Label lblPDFResolution 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resolution:"
         Height          =   195
         Left            =   120
         TabIndex        =   98
         Top             =   1995
         Width           =   795
      End
      Begin VB.Label lblPDFCompat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compatibility:"
         Height          =   195
         Left            =   120
         TabIndex        =   97
         Top             =   540
         Width           =   915
      End
      Begin VB.Label lblPDFAutoRotate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Auto-Rotate Pages:"
         Height          =   195
         Left            =   120
         TabIndex        =   96
         Top             =   1260
         Width           =   1395
      End
   End
   Begin PDFCreator.dmFrame dmFraProgGeneral1 
      Height          =   4110
      Left            =   2625
      TabIndex        =   9
      Top             =   420
      Visible         =   0   'False
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   7250
      Caption         =   "General 1"
      Caption3D       =   2
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
      TextShaddowColor=   12582912
      Begin VB.ComboBox cmbSendMailMethod 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown-Liste
         TabIndex        =   222
         Top             =   3675
         Width           =   2580
      End
      Begin VB.CheckBox chkNoProcessingAtStartup 
         Appearance      =   0  '2D
         Caption         =   "No processing at startup"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   120
         TabIndex        =   177
         Top             =   2280
         Width           =   5775
      End
      Begin VB.CommandButton cmdTestpage 
         Caption         =   "&Print testpage"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   2580
      End
      Begin PDFCreator.Line3D Line3D1 
         Height          =   30
         Index           =   0
         Left            =   105
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1200
         Width           =   5925
         _ExtentX        =   10451
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
         TabIndex        =   4
         Top             =   2640
         Width           =   5775
      End
      Begin MSComctlLib.Slider sldProcessPriority 
         Height          =   495
         Left            =   120
         TabIndex        =   3
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
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2160
         Width           =   5925
         _ExtentX        =   10451
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
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   3240
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   53
         LineType        =   1
         3DHighlight     =   -2147483628
         3DShadow        =   -2147483632
         DrawStyle       =   0
      End
      Begin VB.Label lblSendMailMethod 
         AutoSize        =   -1  'True
         Caption         =   "Methode to send an email"
         Height          =   195
         Left            =   120
         TabIndex        =   221
         Top             =   3465
         Width           =   1830
      End
      Begin VB.Label lblProcessPriority 
         AutoSize        =   -1  'True
         Caption         =   "Processpriority: Normal"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1605
      End
   End
   Begin PDFCreator.dmFrame dmFraShellIntegration 
      Height          =   1065
      Left            =   2640
      TabIndex        =   11
      Top             =   5565
      Visible         =   0   'False
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   1879
      Caption         =   "Shell integration"
      Caption3D       =   2
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
      TextShaddowColor=   12582912
      Enabled         =   0   'False
      Begin VB.CommandButton cmdShellintegration 
         Caption         =   "Integrate PDFCreator into shell"
         Height          =   495
         Index           =   1
         Left            =   3150
         TabIndex        =   7
         Top             =   420
         Width           =   2910
      End
      Begin VB.CommandButton cmdShellintegration 
         Caption         =   "Integrate PDFCreator into shell"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   420
         Width           =   2910
      End
   End
   Begin PDFCreator.dmFrame dmFraProgGeneral2 
      Height          =   2745
      Left            =   3150
      TabIndex        =   215
      Top             =   1470
      Visible         =   0   'False
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   4842
      Caption         =   "General 2"
      Caption3D       =   2
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
      TextShaddowColor=   12582912
      Begin VB.CommandButton cmdAsso 
         Caption         =   "&Associate PDFCreator with Postscript files"
         Height          =   495
         Left            =   120
         TabIndex        =   218
         Top             =   480
         Width           =   2580
      End
      Begin VB.CheckBox chkShowAnimation 
         Appearance      =   0  '2D
         Caption         =   "Show animation"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   120
         TabIndex        =   217
         Top             =   2220
         Width           =   5775
      End
      Begin VB.ComboBox cmbOptionsDesign 
         Height          =   315
         ItemData        =   "frmOptions.frx":001C
         Left            =   120
         List            =   "frmOptions.frx":001E
         Style           =   2  'Dropdown-Liste
         TabIndex        =   216
         Top             =   1620
         Width           =   3870
      End
      Begin PDFCreator.Line3D Line3D1 
         Height          =   30
         Index           =   3
         Left            =   105
         TabIndex        =   219
         TabStop         =   0   'False
         Top             =   1200
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   53
         LineType        =   1
         3DHighlight     =   -2147483628
         3DShadow        =   -2147483632
         DrawStyle       =   0
      End
      Begin VB.Label lblOptionsDesign 
         AutoSize        =   -1  'True
         Caption         =   "Frame color of the setting dialog"
         Height          =   195
         Left            =   120
         TabIndex        =   220
         Top             =   1380
         Width           =   2250
      End
   End
   Begin PDFCreator.dmFrame dmFraProgSave 
      Height          =   2670
      Left            =   2640
      TabIndex        =   51
      Top             =   2160
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4710
      Caption         =   "Save"
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
      Begin VB.ComboBox cmbStandardSaveFormat 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "frmOptions.frx":0020
         Left            =   120
         List            =   "frmOptions.frx":0022
         Style           =   2  'Dropdown-Liste
         TabIndex        =   213
         Top             =   2100
         Width           =   1050
      End
      Begin VB.CheckBox chkSpaces 
         Appearance      =   0  '2D
         Caption         =   "Remove leading and trailing spaces"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   1440
         Value           =   1  'Aktiviert
         Width           =   6015
      End
      Begin VB.TextBox txtSaveFilename 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   120
         TabIndex        =   54
         Text            =   "<Title>"
         Top             =   600
         Width           =   3495
      End
      Begin VB.ComboBox cmbSaveFilenameTokens 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "frmOptions.frx":0024
         Left            =   3720
         List            =   "frmOptions.frx":0026
         Style           =   2  'Dropdown-Liste
         TabIndex        =   53
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtSavePreview 
         Appearance      =   0  '2D
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   960
         Width           =   6015
      End
      Begin VB.Label lblStandardSaveformat 
         AutoSize        =   -1  'True
         Caption         =   "Standard save format"
         Height          =   195
         Left            =   120
         TabIndex        =   212
         Top             =   1890
         Width           =   1515
      End
      Begin VB.Label lblSaveFilename 
         AutoSize        =   -1  'True
         Caption         =   "Filename"
         Height          =   195
         Left            =   120
         TabIndex        =   57
         Top             =   360
         Width           =   630
      End
      Begin VB.Label lblSaveFilenameTokens 
         AutoSize        =   -1  'True
         Caption         =   "Add a Filename-Token"
         Height          =   195
         Left            =   3720
         TabIndex        =   56
         Top             =   360
         Width           =   1605
      End
   End
   Begin PDFCreator.dmFrame dmFraProgPrint 
      Height          =   3930
      Left            =   2940
      TabIndex        =   203
      Top             =   2205
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6932
      Caption         =   "Print after saving"
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
      Begin VB.ComboBox cmbPrintAfterSavingTumble 
         Height          =   315
         Left            =   420
         Style           =   2  'Dropdown-Liste
         TabIndex        =   211
         Top             =   3360
         Width           =   4470
      End
      Begin VB.CheckBox chkPrintAfterSavingDuplex 
         Appearance      =   0  '2D
         Caption         =   "Duplex"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   210
         Top             =   3045
         Width           =   6015
      End
      Begin VB.CheckBox chkPrintAfterSavingNoCancel 
         Appearance      =   0  '2D
         Caption         =   "No cancel dialog"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   209
         Top             =   2625
         Width           =   6015
      End
      Begin VB.ComboBox cmbPrintAfterSavingQueryUser 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown-Liste
         TabIndex        =   208
         Top             =   1995
         Width           =   4770
      End
      Begin VB.ComboBox cmbPrintAfterSavingPrinter 
         Height          =   315
         Left            =   105
         TabIndex        =   206
         Top             =   1155
         Width           =   4770
      End
      Begin VB.CheckBox chkPrintAfterSaving 
         Appearance      =   0  '2D
         Caption         =   "Print after saving"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   204
         Top             =   420
         Width           =   6015
      End
      Begin VB.Label lblPrintAfterSavingQueryUser 
         AutoSize        =   -1  'True
         Caption         =   "Query user"
         Height          =   195
         Left            =   120
         TabIndex        =   207
         Top             =   1785
         Width           =   765
      End
      Begin VB.Label lblPrintAfterSavingPrinter 
         AutoSize        =   -1  'True
         Caption         =   "Printer"
         Height          =   195
         Left            =   120
         TabIndex        =   205
         Top             =   945
         Width           =   450
      End
   End
   Begin PDFCreator.dmFrame dmFraProgActions 
      Height          =   4560
      Left            =   2625
      TabIndex        =   178
      Top             =   1050
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   8043
      Caption         =   "Actions"
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
      Begin PDFCreator.dmFrame dmFraProgActionsRunProgramAfterSaving 
         Height          =   3510
         Left            =   360
         TabIndex        =   189
         Top             =   2400
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   6191
         Caption         =   "Run a program/script after saving"
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
         Begin VB.CommandButton cmdRunProgramAfterSavingPrognameEdit 
            Height          =   300
            Left            =   5520
            Picture         =   "frmOptions.frx":0028
            Style           =   1  'Grafisch
            TabIndex        =   200
            Top             =   1155
            Width           =   435
         End
         Begin VB.ComboBox cmbRunProgramAfterSavingWindowstyle 
            Height          =   315
            Left            =   210
            Style           =   2  'Dropdown-Liste
            TabIndex        =   195
            Top             =   3045
            Width           =   5370
         End
         Begin VB.CheckBox chkRunProgramAfterSavingWaitUntilReady 
            Appearance      =   0  '2D
            Caption         =   "Wait until ready"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   210
            TabIndex        =   194
            Top             =   2415
            Width           =   5805
         End
         Begin VB.TextBox txtRunProgramAfterSavingProgramParameters 
            Appearance      =   0  '2D
            Height          =   285
            Left            =   210
            TabIndex        =   193
            Top             =   1890
            Width           =   5805
         End
         Begin VB.CommandButton cmdRunProgramAfterSavingPrognameChoice 
            Caption         =   "..."
            Height          =   300
            Left            =   5040
            TabIndex        =   192
            Top             =   1155
            Width           =   435
         End
         Begin VB.ComboBox cmbRunProgramAfterSavingProgramname 
            Height          =   315
            Left            =   210
            TabIndex        =   191
            Top             =   1155
            Width           =   4770
         End
         Begin VB.CheckBox chkRunProgramAfterSaving 
            Appearance      =   0  '2D
            Caption         =   "Run a program/script after saving"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   210
            TabIndex        =   190
            Top             =   420
            Width           =   5805
         End
         Begin VB.Label lblRunProgramAfterSavingWindowstyle 
            AutoSize        =   -1  'True
            Caption         =   "Windowstyle"
            Height          =   195
            Left            =   210
            TabIndex        =   198
            Top             =   2835
            Width           =   900
         End
         Begin VB.Label lblRunProgramAfterSavingProgramParameters 
            AutoSize        =   -1  'True
            Caption         =   "Program parameters"
            Height          =   195
            Left            =   210
            TabIndex        =   197
            Top             =   1680
            Width           =   1410
         End
         Begin VB.Label lblRunProgramAfterSavingProgramname 
            AutoSize        =   -1  'True
            Caption         =   "Program/Script"
            Height          =   195
            Left            =   210
            TabIndex        =   196
            Top             =   945
            Width           =   1065
         End
      End
      Begin PDFCreator.dmFrame dmFraProgActionsRunProgramBeforeSaving 
         Height          =   3510
         Left            =   210
         TabIndex        =   179
         Top             =   735
         Visible         =   0   'False
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   6191
         Caption         =   "Run a program/script before saving"
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
         Begin VB.CommandButton cmdRunProgramBeforeSavingPrognameEdit 
            Height          =   300
            Left            =   5520
            Picture         =   "frmOptions.frx":05B2
            Style           =   1  'Grafisch
            TabIndex        =   201
            Top             =   1155
            Width           =   435
         End
         Begin VB.ComboBox cmbRunProgramBeforeSavingWindowstyle 
            Height          =   315
            Left            =   210
            Style           =   2  'Dropdown-Liste
            TabIndex        =   187
            Top             =   3045
            Width           =   2790
         End
         Begin VB.CheckBox chkRunProgramBeforeSavingWaitUntilReady 
            Appearance      =   0  '2D
            Caption         =   "Wait until ready"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   210
            TabIndex        =   186
            Top             =   2415
            Width           =   5580
         End
         Begin VB.TextBox txtRunProgramBeforeSavingProgramParameters 
            Appearance      =   0  '2D
            Height          =   285
            Left            =   210
            TabIndex        =   184
            Top             =   1890
            Width           =   5580
         End
         Begin VB.CommandButton cmdRunProgramBeforeSavingPrognameChoice 
            Caption         =   "..."
            Height          =   300
            Left            =   5040
            TabIndex        =   183
            Top             =   1155
            Width           =   435
         End
         Begin VB.ComboBox cmbRunProgramBeforeSavingProgramname 
            Height          =   315
            Left            =   210
            TabIndex        =   182
            Top             =   1155
            Width           =   4770
         End
         Begin VB.CheckBox chkRunProgramBeforeSaving 
            Appearance      =   0  '2D
            Caption         =   "Run a program/script before saving"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   210
            TabIndex        =   180
            Top             =   420
            Width           =   5385
         End
         Begin VB.Label lblRunProgramBeforeSavingWindowstyle 
            AutoSize        =   -1  'True
            Caption         =   "Windowstyle"
            Height          =   195
            Left            =   210
            TabIndex        =   188
            Top             =   2835
            Width           =   900
         End
         Begin VB.Label lblRunProgramBeforeSavingProgramParameters 
            AutoSize        =   -1  'True
            Caption         =   "Program parameters"
            Height          =   195
            Left            =   210
            TabIndex        =   185
            Top             =   1680
            Width           =   1410
         End
         Begin VB.Label lblRunProgramBeforeSavingProgramname 
            AutoSize        =   -1  'True
            Caption         =   "Program/Script"
            Height          =   195
            Left            =   210
            TabIndex        =   181
            Top             =   945
            Width           =   1065
         End
      End
      Begin MSComctlLib.TabStrip tbstrProgActions 
         Height          =   4110
         Left            =   105
         TabIndex        =   199
         Top             =   315
         Width           =   6210
         _ExtentX        =   10954
         _ExtentY        =   7250
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin PDFCreator.dmFrame dmFraProgAutosave 
      Height          =   5085
      Left            =   2640
      TabIndex        =   30
      Top             =   1200
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   8969
      Caption         =   "Autosave"
      Caption3D       =   2
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
      TextShaddowColor=   12582912
      Begin VB.CheckBox chkAutosaveSendEmail 
         Appearance      =   0  '2D
         Caption         =   "Send an email after auto-saving"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   202
         Top             =   4680
         Width           =   5895
      End
      Begin VB.CheckBox chkAutosaveStartStandardProgram 
         Appearance      =   0  '2D
         Caption         =   "After auto-saving open the document with the default program."
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   120
         TabIndex        =   171
         Top             =   4095
         Width           =   5895
      End
      Begin VB.CommandButton cmdGetAutosaveDirectory 
         Caption         =   "..."
         Height          =   300
         Left            =   5760
         TabIndex        =   167
         Top             =   3120
         Width           =   375
      End
      Begin VB.ComboBox cmbAutosaveFormat 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown-Liste
         TabIndex        =   38
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox cmbAutoSaveFilenameTokens 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "frmOptions.frx":0B3C
         Left            =   3690
         List            =   "frmOptions.frx":0B3E
         Style           =   2  'Dropdown-Liste
         TabIndex        =   37
         Top             =   1785
         Width           =   2460
      End
      Begin VB.TextBox txtAutosaveFilename 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   120
         TabIndex        =   36
         Text            =   "<DateTime>"
         Top             =   1800
         Width           =   3495
      End
      Begin VB.TextBox txtAutosaveDirectory 
         Appearance      =   0  '2D
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   35
         Top             =   3120
         Width           =   5535
      End
      Begin VB.CheckBox chkUseAutosaveDirectory 
         Appearance      =   0  '2D
         Caption         =   "For autosave use this directory"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   2760
         Width           =   5895
      End
      Begin VB.CheckBox chkUseAutosave 
         Appearance      =   0  '2D
         Caption         =   "Use Autosave"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   6015
      End
      Begin VB.TextBox txtAutoSaveFilenamePreview 
         Appearance      =   0  '2D
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   2145
         Width           =   6015
      End
      Begin VB.TextBox txtAutoSaveDirectoryPreview 
         Appearance      =   0  '2D
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   3450
         Width           =   6015
      End
      Begin VB.Label lblAutosaveFilenameTokens 
         AutoSize        =   -1  'True
         Caption         =   "Add a Filename-Token"
         Height          =   195
         Left            =   3720
         TabIndex        =   41
         Top             =   1560
         Width           =   1605
      End
      Begin VB.Label lblAutosaveformat 
         Caption         =   "Autosaveformat"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label lblAutosaveFilename 
         AutoSize        =   -1  'True
         Caption         =   "Filename"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   1560
         Width           =   630
      End
   End
   Begin PDFCreator.dmFrame dmFraProgGhostscript 
      Height          =   3150
      Left            =   2625
      TabIndex        =   12
      Top             =   945
      Visible         =   0   'False
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   5556
      Caption         =   "Ghostscript"
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
      TextShaddowColor=   12582912
      Begin VB.CheckBox chkAddWindowsFontpath 
         Appearance      =   0  '2D
         Caption         =   "Add Windows fontpath"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   105
         TabIndex        =   176
         Top             =   2730
         Width           =   6105
      End
      Begin VB.TextBox txtAdditionalGhostscriptSearchpath 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   105
         TabIndex        =   174
         Top             =   2100
         Width           =   6105
      End
      Begin VB.TextBox txtAdditionalGhostscriptParameters 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   105
         TabIndex        =   173
         Top             =   1365
         Width           =   6105
      End
      Begin VB.ComboBox cmbGhostscript 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   105
         Style           =   2  'Dropdown-Liste
         TabIndex        =   21
         Top             =   630
         Width           =   4215
      End
      Begin VB.CommandButton cmdGetgsbinDirectory 
         Caption         =   "..."
         Height          =   255
         Left            =   5625
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   4890
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.CommandButton cmdGetgslibDirectory 
         Caption         =   "..."
         Height          =   255
         Left            =   5625
         TabIndex        =   16
         Top             =   4290
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdGetgsfontsDirectory 
         Caption         =   "..."
         Height          =   255
         Left            =   5625
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   5490
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.CommandButton cmdGetgsresourceDirectory 
         Caption         =   "..."
         Height          =   255
         Left            =   5625
         TabIndex        =   13
         Top             =   5490
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblAdditionalGhostscriptSearchpath 
         AutoSize        =   -1  'True
         Caption         =   "Additional Ghostscript searchpath"
         Height          =   195
         Left            =   105
         TabIndex        =   175
         Top             =   1890
         Width           =   2370
      End
      Begin VB.Label lblAdditionalGhostscriptParameters 
         AutoSize        =   -1  'True
         Caption         =   "Additional Ghostscript parameters"
         Height          =   195
         Left            =   105
         TabIndex        =   172
         Top             =   1155
         Width           =   2355
      End
      Begin VB.Label lblGhostscriptversion 
         AutoSize        =   -1  'True
         Caption         =   "Ghostscriptversion"
         Height          =   195
         Left            =   105
         TabIndex        =   26
         Top             =   420
         Width           =   1305
      End
      Begin VB.Label lblGSbin 
         AutoSize        =   -1  'True
         Caption         =   "Ghostscript Binaries"
         Height          =   195
         Left            =   105
         TabIndex        =   25
         Top             =   3450
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label lblGSlib 
         AutoSize        =   -1  'True
         Caption         =   "Ghostscript Libraries"
         Height          =   195
         Left            =   105
         TabIndex        =   24
         Top             =   4050
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label lblGSfonts 
         AutoSize        =   -1  'True
         Caption         =   "Ghostscript Fonts"
         Height          =   195
         Left            =   105
         TabIndex        =   23
         Top             =   4650
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label lblGhostscriptResource 
         AutoSize        =   -1  'True
         Caption         =   "Ghostscript Resource"
         Height          =   195
         Left            =   105
         TabIndex        =   22
         Top             =   5250
         Visible         =   0   'False
         Width           =   1530
      End
   End
   Begin PDFCreator.dmFrame dmFraProgDirectories 
      Height          =   1410
      Left            =   2640
      TabIndex        =   42
      Top             =   1320
      Visible         =   0   'False
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   2487
      Caption         =   "Directories"
      Caption3D       =   2
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
      TextShaddowColor=   12582912
      Begin VB.TextBox txtTemppathPreview 
         Appearance      =   0  '2D
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   170
         Top             =   945
         Width           =   5910
      End
      Begin VB.CommandButton cmdGetTemppath 
         Caption         =   "..."
         Height          =   300
         Left            =   5154
         TabIndex        =   165
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtTemppath 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   120
         TabIndex        =   43
         Top             =   600
         Width           =   4965
      End
      Begin VB.CommandButton cmdUsertempPath 
         Height          =   300
         Left            =   5640
         Picture         =   "frmOptions.frx":0B40
         Style           =   1  'Grafisch
         TabIndex        =   166
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblPrintTempPath 
         AutoSize        =   -1  'True
         Caption         =   "Temppath"
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   720
      End
   End
   Begin PDFCreator.dmFrame dmFraPSGeneral 
      Height          =   1095
      Left            =   2640
      TabIndex        =   86
      Top             =   1920
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   1931
      Caption         =   "Postscript"
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
      Begin VB.ComboBox cmbEPSLanguageLevel 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown-Liste
         TabIndex        =   88
         Top             =   600
         Width           =   735
      End
      Begin VB.ComboBox cmbPSLanguageLevel 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown-Liste
         TabIndex        =   87
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblLangLevel 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         Caption         =   "Language Level:"
         Height          =   195
         Left            =   735
         TabIndex        =   89
         Top             =   510
         Width           =   1200
      End
   End
   Begin PDFCreator.dmFrame dmFraBitmapGeneral 
      Height          =   1935
      Left            =   2640
      TabIndex        =   73
      Top             =   1680
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3413
      Caption         =   "Bitmap"
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
      Begin VB.ComboBox cmbTIFFColors 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown-Liste
         TabIndex        =   80
         Top             =   1440
         Width           =   2175
      End
      Begin VB.ComboBox cmbPCXColors 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown-Liste
         TabIndex        =   79
         Top             =   1320
         Width           =   2175
      End
      Begin VB.ComboBox cmbBMPColors 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown-Liste
         TabIndex        =   78
         Top             =   1200
         Width           =   2175
      End
      Begin VB.ComboBox cmbJPEGColors 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown-Liste
         TabIndex        =   77
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtJPEGQuality 
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         Height          =   285
         Left            =   1920
         TabIndex        =   76
         Text            =   "75"
         Top             =   1440
         Width           =   495
      End
      Begin VB.ComboBox cmbPNGColors 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown-Liste
         TabIndex        =   75
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox txtBitmapResolution 
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         Height          =   285
         Left            =   1920
         TabIndex        =   74
         Text            =   "72"
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblJPEQQualityProzent 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   2520
         TabIndex        =   85
         Top             =   1485
         Width           =   120
      End
      Begin VB.Label lblJPEGQuality 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         Caption         =   "Quality:"
         Height          =   195
         Left            =   1290
         TabIndex        =   84
         Top             =   1485
         Width           =   525
      End
      Begin VB.Label lblBitmapColors 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         Caption         =   "Colors:"
         Height          =   195
         Left            =   1335
         TabIndex        =   83
         Top             =   1020
         Width           =   480
      End
      Begin VB.Label lblBitmapDPI 
         AutoSize        =   -1  'True
         Caption         =   "dpi"
         Height          =   195
         Left            =   2520
         TabIndex        =   82
         Top             =   525
         Width           =   210
      End
      Begin VB.Label lblBitmapResolution 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         Caption         =   "Resolution:"
         Height          =   195
         Left            =   1020
         TabIndex        =   81
         Top             =   525
         Width           =   795
      End
   End
   Begin PDFCreator.dmFrame dmFraPDFFonts 
      Height          =   1695
      Left            =   2760
      TabIndex        =   124
      Top             =   2400
      Visible         =   0   'False
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
      Begin VB.CheckBox chkPDFEmbedAll 
         Appearance      =   0  '2D
         Caption         =   "Embed all Fonts"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   127
         Top             =   360
         Width           =   5955
      End
      Begin VB.CheckBox chkPDFSubSetFonts 
         Appearance      =   0  '2D
         Caption         =   "Subset Fonts, when percentage of used characters below:"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   126
         Top             =   780
         Width           =   5955
      End
      Begin VB.TextBox txtPDFSubSetPerc 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   400
         TabIndex        =   125
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblPDFPerc 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   960
         TabIndex        =   128
         Top             =   1365
         Width           =   120
      End
   End
   Begin PDFCreator.dmFrame dmFraProgFont 
      Height          =   4695
      Left            =   2640
      TabIndex        =   64
      Top             =   1440
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   8281
      Caption         =   "Programfont"
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
      Begin VB.CommandButton cmdCancelTest 
         Caption         =   "C&ancel test"
         Height          =   495
         Left            =   2310
         TabIndex        =   169
         Top             =   4095
         Width           =   1755
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "&Test"
         Height          =   495
         Left            =   120
         TabIndex        =   168
         Top             =   4095
         Width           =   1755
      End
      Begin VB.TextBox txtTest 
         Appearance      =   0  '2D
         Height          =   2655
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertikal
         TabIndex        =   68
         Top             =   1320
         Width           =   6135
      End
      Begin VB.ComboBox cmbCharset 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   3000
         TabIndex        =   67
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
         TabIndex        =   66
         Top             =   600
         Width           =   2775
      End
      Begin VB.ComboBox cmbProgramFontsize 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   5400
         TabIndex        =   65
         Text            =   "8"
         Top             =   600
         Width           =   765
      End
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         Caption         =   "Size"
         Height          =   195
         Left            =   5400
         TabIndex        =   72
         Top             =   360
         Width           =   300
      End
      Begin VB.Label lblTesttext 
         AutoSize        =   -1  'True
         Caption         =   "Here you can test the font."
         Height          =   195
         Left            =   120
         TabIndex        =   71
         Top             =   1080
         Width           =   1890
      End
      Begin VB.Label lblProgcharset 
         AutoSize        =   -1  'True
         Caption         =   "Charset"
         Height          =   195
         Left            =   3000
         TabIndex        =   70
         Top             =   360
         Width           =   540
      End
      Begin VB.Label lblProgfont 
         AutoSize        =   -1  'True
         Caption         =   "Programfont"
         Height          =   195
         Left            =   120
         TabIndex        =   69
         Top             =   360
         Width           =   855
      End
   End
   Begin PDFCreator.dmFrame dmFraPDFSecurity 
      Height          =   5535
      Left            =   2730
      TabIndex        =   136
      Top             =   2205
      Visible         =   0   'False
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
      Begin PDFCreator.dmFrame dmFraPDFHighPermissions 
         Height          =   855
         Left            =   120
         TabIndex        =   151
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
         Begin VB.CheckBox chkAllowAssembly 
            Appearance      =   0  '2D
            Caption         =   "Allow changes to the Assembly"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3100
            TabIndex        =   155
            Top             =   525
            Width           =   2760
         End
         Begin VB.CheckBox chkAllowScreenReaders 
            Appearance      =   0  '2D
            Caption         =   "Allow Screen Readers"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   154
            Top             =   525
            Width           =   2865
         End
         Begin VB.CheckBox chkAllowFillIn 
            Appearance      =   0  '2D
            Caption         =   "Allow filling in form fields"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3100
            TabIndex        =   153
            Top             =   300
            Width           =   2760
         End
         Begin VB.CheckBox chkAllowDegradedPrinting 
            Appearance      =   0  '2D
            Caption         =   "Allow printing in low resolution"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   152
            Top             =   300
            Width           =   2865
         End
      End
      Begin PDFCreator.dmFrame dmFraPDFPermissions 
         Height          =   855
         Left            =   120
         TabIndex        =   146
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
         Begin VB.CheckBox chkAllowModifyAnnotations 
            Appearance      =   0  '2D
            Caption         =   "modify comments"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3100
            TabIndex        =   150
            Top             =   525
            Width           =   2760
         End
         Begin VB.CheckBox chkAllowModifyContents 
            Appearance      =   0  '2D
            Caption         =   "modify the document"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3100
            TabIndex        =   149
            Top             =   300
            Width           =   2760
         End
         Begin VB.CheckBox chkAllowCopy 
            Appearance      =   0  '2D
            Caption         =   "copy text and images"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   148
            Top             =   525
            Width           =   2865
         End
         Begin VB.CheckBox chkAllowPrinting 
            Appearance      =   0  '2D
            Caption         =   "print the document"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   147
            Top             =   300
            Width           =   2865
         End
      End
      Begin PDFCreator.dmFrame dmFraSecurityPass 
         Height          =   855
         Left            =   120
         TabIndex        =   143
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
         Begin VB.CheckBox chkOwnerPass 
            Appearance      =   0  '2D
            Caption         =   "Password required to change Permissions and Passwords"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   145
            Top             =   525
            Width           =   5700
         End
         Begin VB.CheckBox chkUserPass 
            Appearance      =   0  '2D
            Caption         =   "Password required to open document"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   144
            Top             =   300
            Width           =   5700
         End
      End
      Begin PDFCreator.dmFrame dmFraPDFEncLevel 
         Height          =   855
         Left            =   120
         TabIndex        =   140
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
         Begin VB.OptionButton optEncHigh 
            Appearance      =   0  '2D
            Caption         =   "High (128 Bit - Adobe Acrobat 5.0 and above)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   142
            Top             =   525
            Width           =   5775
         End
         Begin VB.OptionButton optEncLow 
            Appearance      =   0  '2D
            Caption         =   "Low (40 Bit - Adobe Acrobat 3.0 and above)"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   141
            Top             =   300
            Width           =   5775
         End
      End
      Begin PDFCreator.dmFrame dmFraPDFEncryptor 
         Height          =   855
         Left            =   120
         TabIndex        =   138
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
            ItemData        =   "frmOptions.frx":10CA
            Left            =   120
            List            =   "frmOptions.frx":10CC
            Style           =   2  'Dropdown-Liste
            TabIndex        =   139
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
         TabIndex        =   137
         Top             =   360
         Width           =   5535
      End
   End
   Begin PDFCreator.dmFrame dmfraPDFCompress 
      Height          =   4335
      Left            =   2760
      TabIndex        =   101
      Top             =   1920
      Visible         =   0   'False
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
      Begin PDFCreator.dmFrame dmFraPDFMono 
         Height          =   1095
         Left            =   120
         TabIndex        =   117
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
         Begin VB.CheckBox chkPDFMonoComp 
            Appearance      =   0  '2D
            Caption         =   "Compress"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   122
            Top             =   360
            Width           =   2325
         End
         Begin VB.ComboBox cmbPDFMonoComp 
            Appearance      =   0  '2D
            Height          =   315
            ItemData        =   "frmOptions.frx":10CE
            Left            =   120
            List            =   "frmOptions.frx":10D0
            Style           =   2  'Dropdown-Liste
            TabIndex        =   121
            Top             =   660
            Width           =   2370
         End
         Begin VB.CheckBox chkPDFMonoResample 
            Appearance      =   0  '2D
            Caption         =   "Resample"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2520
            TabIndex        =   120
            Top             =   360
            Width           =   2295
         End
         Begin VB.ComboBox cmbPDFMonoResample 
            Appearance      =   0  '2D
            Height          =   315
            ItemData        =   "frmOptions.frx":10D2
            Left            =   2520
            List            =   "frmOptions.frx":10D4
            Style           =   2  'Dropdown-Liste
            TabIndex        =   119
            Tag             =   "Bicubic|Subsample|Average"
            Top             =   660
            Width           =   2370
         End
         Begin VB.TextBox txtPDFMonoRes 
            Appearance      =   0  '2D
            Height          =   285
            Left            =   4935
            TabIndex        =   118
            Top             =   660
            Width           =   735
         End
         Begin VB.Label lblPDFMonoRes 
            AutoSize        =   -1  'True
            Caption         =   "Resolution"
            Height          =   195
            Left            =   4935
            TabIndex        =   123
            Top             =   360
            Width           =   750
         End
      End
      Begin PDFCreator.dmFrame dmFraPDFGrey 
         Height          =   1095
         Left            =   120
         TabIndex        =   110
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
         Begin VB.TextBox txtPDFGreyRes 
            Appearance      =   0  '2D
            Height          =   285
            Left            =   4935
            TabIndex        =   115
            Top             =   660
            Width           =   735
         End
         Begin VB.ComboBox cmbPDFGreyResample 
            Appearance      =   0  '2D
            Height          =   315
            ItemData        =   "frmOptions.frx":10D6
            Left            =   2520
            List            =   "frmOptions.frx":10D8
            Style           =   2  'Dropdown-Liste
            TabIndex        =   114
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
            TabIndex        =   113
            Top             =   360
            Width           =   2295
         End
         Begin VB.ComboBox cmbPDFGreyComp 
            Appearance      =   0  '2D
            Height          =   315
            ItemData        =   "frmOptions.frx":10DA
            Left            =   120
            List            =   "frmOptions.frx":10DC
            Style           =   2  'Dropdown-Liste
            TabIndex        =   112
            Top             =   660
            Width           =   2370
         End
         Begin VB.CheckBox chkPDFGreyComp 
            Appearance      =   0  '2D
            Caption         =   "Compress"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   111
            Top             =   360
            Width           =   2325
         End
         Begin VB.Label lblPDFGreyRes 
            AutoSize        =   -1  'True
            Caption         =   "Resolution"
            Height          =   195
            Left            =   4935
            TabIndex        =   116
            Top             =   360
            Width           =   750
         End
      End
      Begin PDFCreator.dmFrame dmFraPDFColor 
         Height          =   1095
         Left            =   120
         TabIndex        =   103
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
         Begin VB.TextBox txtPDFColorRes 
            Appearance      =   0  '2D
            Height          =   285
            Left            =   4935
            TabIndex        =   108
            Top             =   660
            Width           =   735
         End
         Begin VB.ComboBox cmbPDFColorResample 
            Appearance      =   0  '2D
            Height          =   315
            ItemData        =   "frmOptions.frx":10DE
            Left            =   2520
            List            =   "frmOptions.frx":10E0
            Style           =   2  'Dropdown-Liste
            TabIndex        =   107
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
            TabIndex        =   106
            Top             =   360
            Width           =   2295
         End
         Begin VB.ComboBox cmbPDFColorComp 
            Appearance      =   0  '2D
            Height          =   315
            ItemData        =   "frmOptions.frx":10E2
            Left            =   120
            List            =   "frmOptions.frx":10E4
            Style           =   2  'Dropdown-Liste
            TabIndex        =   105
            Top             =   660
            Width           =   2370
         End
         Begin VB.CheckBox chkPDFColorComp 
            Appearance      =   0  '2D
            Caption         =   "Compress"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   104
            Top             =   360
            Width           =   2325
         End
         Begin VB.Label lblPDFColorRes 
            AutoSize        =   -1  'True
            Caption         =   "Resolution"
            Height          =   195
            Left            =   4935
            TabIndex        =   109
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
         TabIndex        =   102
         Top             =   360
         Width           =   5910
      End
   End
   Begin PDFCreator.dmFrame dmFraPDFColors 
      Height          =   1215
      Left            =   2760
      TabIndex        =   129
      Top             =   2760
      Visible         =   0   'False
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
      Begin VB.ComboBox cmbPDFColorModel 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "frmOptions.frx":10E6
         Left            =   120
         List            =   "frmOptions.frx":10E8
         Style           =   2  'Dropdown-Liste
         TabIndex        =   131
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
         TabIndex        =   130
         Top             =   840
         Width           =   5880
      End
   End
   Begin PDFCreator.dmFrame dmFraPDFColorOptions 
      Height          =   1455
      Left            =   2760
      TabIndex        =   132
      Top             =   4080
      Visible         =   0   'False
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
      Begin VB.CheckBox chkPDFPreserveHalftone 
         Appearance      =   0  '2D
         Caption         =   "Preserve Halftone Information"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   135
         Top             =   1050
         Width           =   5910
      End
      Begin VB.CheckBox chkPDFPreserveTransfer 
         Appearance      =   0  '2D
         Caption         =   "Preserve Transfer Functions"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   134
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
         TabIndex        =   133
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
      TabIndex        =   159
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
            NumListImages   =   19
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":10EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":1684
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":1C1E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":21B8
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":2752
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":2AEC
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":3086
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":3960
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":3EFA
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":4494
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":4A2E
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":4FC8
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":5562
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":5AFC
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":6096
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":6630
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":6BCA
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":7164
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":7A3E
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
   Begin PDFCreator.dmFrame dmFraFilenameSubstitutions 
      Height          =   2535
      Left            =   2640
      TabIndex        =   58
      Top             =   4200
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4471
      Caption         =   "Filename substitutions"
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
      Begin VB.CommandButton cmdFilenameSubst 
         Caption         =   "&Add"
         Height          =   375
         Index           =   0
         Left            =   4440
         TabIndex        =   162
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdFilenameSubstMove 
         Height          =   435
         Index           =   0
         Left            =   120
         Picture         =   "frmOptions.frx":8318
         Style           =   1  'Grafisch
         TabIndex        =   160
         Top             =   915
         Width           =   375
      End
      Begin VB.TextBox txtFilenameSubst 
         Appearance      =   0  '2D
         Height          =   285
         Index           =   1
         Left            =   2640
         TabIndex        =   61
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox chkFilenameSubst 
         Appearance      =   0  '2D
         Caption         =   "Substitutions only in <Title>"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   60
         Top             =   2160
         Value           =   1  'Aktiviert
         Width           =   3255
      End
      Begin VB.TextBox txtFilenameSubst 
         Appearance      =   0  '2D
         Height          =   285
         Index           =   0
         Left            =   600
         TabIndex        =   59
         Top             =   360
         Width           =   1695
      End
      Begin MSComctlLib.ListView lsvFilenameSubst 
         Height          =   1335
         Left            =   600
         TabIndex        =   62
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
         Picture         =   "frmOptions.frx":86A2
         Style           =   1  'Grafisch
         TabIndex        =   161
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton cmdFilenameSubst 
         Caption         =   "C&hange"
         Height          =   375
         Index           =   1
         Left            =   4440
         TabIndex        =   163
         Top             =   1155
         Width           =   1455
      End
      Begin VB.CommandButton cmdFilenameSubst 
         Caption         =   "&Delete"
         Height          =   375
         Index           =   2
         Left            =   4440
         TabIndex        =   164
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblEqual 
         Caption         =   "="
         Height          =   255
         Left            =   2400
         TabIndex        =   63
         Top             =   360
         Width           =   135
      End
   End
   Begin PDFCreator.dmFrame dmFraDescription 
      Height          =   1065
      Left            =   2640
      TabIndex        =   156
      Top             =   105
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1879
      Caption         =   ""
      BarColorFrom    =   8421631
      BarColorTo      =   192
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
         TabIndex        =   157
         Top             =   420
         Width           =   615
      End
      Begin VB.Label lblOptions 
         Height          =   615
         Left            =   735
         TabIndex        =   158
         Top             =   420
         Width           =   5655
      End
   End
   Begin MSComctlLib.TabStrip tbstrPDFOptions 
      Height          =   4935
      Left            =   2640
      TabIndex        =   8
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
   Begin MSComctlLib.TabStrip tbstrProgGeneral 
      Height          =   4935
      Left            =   2600
      TabIndex        =   214
      Top             =   420
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

Private Sub chkPrintAfterSaving_Click()
 If chkPrintAfterSaving.Value = 1 Then
   ViewPrintAfterSaving True
  Else
   ViewPrintAfterSaving False
 End If
End Sub

Private Sub chkPrintAfterSavingDuplex_Click()
 If chkPrintAfterSavingDuplex.Value = 1 Then
   ViewPrintAfterTumple True
  Else
   ViewPrintAfterTumple False
 End If
End Sub

Private Sub chkRunProgramBeforeSaving_Click()
 If chkRunProgramBeforeSaving.Value = 1 Then
   ViewRunProgramBeforeSaving True
  Else
   ViewRunProgramBeforeSaving False
 End If
End Sub

Private Sub chkRunProgramAfterSaving_Click()
 If chkRunProgramAfterSaving.Value = 1 Then
   ViewRunProgramAfterSaving True
  Else
   ViewRunProgramAfterSaving False
 End If
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

Private Sub cmbAutosaveFormat_Click()
 Dim Ext As String
 txtAutosaveFilename.ToolTipText = txtAutosaveFilename.Text
 txtAutoSaveFilenamePreview.Text = GetSubstFilename("B:\dummy.dum", txtAutosaveFilename.Text, , True) & _
  GetAutosaveFormatExtension
 If IsValidPath("C:\" & txtAutoSaveFilenamePreview.Text) = False Then
   txtAutoSaveFilenamePreview.ForeColor = vbRed
  Else
   txtAutoSaveFilenamePreview.ForeColor = &H80000008
 End If
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
   txtGSresource.Text = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryResource"))
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
     If InStrRev(Path, "\") > 0 Then
      txtGSlib.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
      txtGSfonts.Text = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
      If tsf(UBound(tsf)) <> "8.00" Then
       txtGSresource.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "Resource"
      End If
     End If
    End If
   End If
   If InStr(UCase$(gsv), "GNU") Then
    If InStr(gsv, " ") > 0 Then
     tsf = Split(gsv, " ")
     reg.KeyRoot = "SOFTWARE\GNU Ghostscript\" & tsf(UBound(tsf))
     tStr = reg.GetRegistryValue("GS_DLL")
     SplitPath tStr, , Path
     txtGSbin.Text = CompletePath(Path)
     If InStrRev(Path, "\") > 0 Then
      txtGSlib.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
      txtGSfonts.Text = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
      txtGSresource.Text = ""
     End If
    End If
   End If
   If InStr(UCase$(gsv), "GPL") Then
    If InStr(gsv, " ") > 0 Then
     tsf = Split(gsv, " ")
     reg.KeyRoot = "SOFTWARE\GPL Ghostscript\" & tsf(UBound(tsf))
     tStr = reg.GetRegistryValue("GS_DLL")
     SplitPath tStr, , Path
     txtGSbin.Text = CompletePath(Path)
     If InStrRev(Path, "\") > 0 Then
      txtGSlib.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
      txtGSfonts.Text = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
      txtGSresource.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "Resource"
     End If
    End If
   End If
 End If
 Set reg = Nothing
End Sub

Private Sub cmbOptionsDesign_Click()
 Options.OptionsDesign = cmbOptionsDesign.ListIndex
 SetFrames
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

Private Sub cmbRunProgramAfterSavingProgramname_Change()
 Dim Program As String, drv As String
 Program = RemoveLeadingAndTrailingQuotes(cmbRunProgramAfterSavingProgramname.Text)
 SplitPath Program, drv
 If LenB(drv) = 0 Then
  Program = ResolveRelativePath(Program, GetPDFCreatorApplicationPath)
 End If
 If FileExists(Program) = True Then
   If IsFileEditable(Program) Then
     cmdRunProgramAfterSavingPrognameEdit.Enabled = True
    Else
     cmdRunProgramAfterSavingPrognameEdit.Enabled = False
   End If
  Else
   cmdRunProgramAfterSavingPrognameEdit.Enabled = False
 End If
End Sub

Private Sub cmbRunProgramAfterSavingProgramname_Click()
 With cmbRunProgramAfterSavingProgramname
  If .ListCount > 0 Then
   .Text = "Scripts\RunProgramAfterSaving\" & .List(.ListIndex)
  End If
 End With
End Sub

Private Sub cmbRunProgramBeforeSavingProgramname_Change()
 Dim Program As String, drv As String
 Program = RemoveLeadingAndTrailingQuotes(cmbRunProgramBeforeSavingProgramname.Text)
 SplitPath Program, drv
 If LenB(drv) = 0 Then
  Program = ResolveRelativePath(Program, GetPDFCreatorApplicationPath)
 End If
 If FileExists(Program) = True Then
   If IsFileEditable(Program) Then
     cmdRunProgramBeforeSavingPrognameEdit.Enabled = True
    Else
     cmdRunProgramBeforeSavingPrognameEdit.Enabled = False
   End If
  Else
   cmdRunProgramBeforeSavingPrognameEdit.Enabled = False
 End If
End Sub

Private Sub cmbRunProgramBeforeSavingProgramname_Click()
 With cmbRunProgramBeforeSavingProgramname
  If .ListCount > 0 Then
   .Text = "Scripts\RunProgramBeforeSaving\" & .List(.ListIndex)
  End If
 End With
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
  ieb.Refresh
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
 End Select
End Sub

Private Sub cmdFilenameSubstMove_Click(Index As Integer)
 Select Case Index
  Case 0: ' Up
   MoveUpFilenameSubstitutions
  Case 1: ' Down
   MoveDownFilenameSubstitutions
 End Select
End Sub

Private Sub cmdGetAutosaveDirectory_Click()
 Dim strFolder As String
 strFolder = BrowseForFolderFiles(Me.hwnd, LanguageStrings.OptionsAutosaveDirectoryPrompt)
 If Len(strFolder) = 0 Then
  Exit Sub
 End If
 txtAutosaveDirectory.Text = CompletePath(strFolder)
End Sub

Private Sub cmdGetgsbinDirectory_Click()
 Dim strFolder As String, aw As Long
 strFolder = BrowseForFolderFiles(Me.hwnd, LanguageStrings.OptionsGhostscriptBinariesDirectoryPrompt)
 If Len(strFolder) = 0 Then
  Exit Sub
 End If
 strFolder = CompletePath(strFolder)
 If FileExists(strFolder & GsDll) = False Then
  MsgBox LanguageStrings.MessagesMsg15
  Exit Sub
 End If
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
 With txtGSbin
  .Text = strFolder
  .ToolTipText = .Text
 End With
End Sub

Private Sub cmdGetgsfontsDirectory_Click()
 Dim strFolder As String
 strFolder = BrowseForFolderFiles(Me.hwnd, LanguageStrings.OptionsGhostscriptFontsDirectoryPrompt)
 If Len(strFolder) = 0 Then Exit Sub
 strFolder = CompletePath(strFolder)
 If LenB(Dir(strFolder & "*.afm", vbNormal)) = 0 And LenB(Dir(strFolder & "*.pfb", vbNormal)) = 0 Then
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
 strFolder = BrowseForFolderFiles(Me.hwnd, LanguageStrings.OptionsGhostscriptLibrariesDirectoryPrompt)
 If Len(strFolder) = 0 Then Exit Sub
 strFolder = CompletePath(strFolder)
 If LenB(Dir(strFolder & "*.*", vbNormal)) = 0 Then
  MsgBox LanguageStrings.MessagesMsg17
  Exit Sub
 End If
 With txtGSlib
  .Text = strFolder
  .ToolTipText = .Text
 End With
End Sub

Private Sub cmdGetgsresourceDirectory_Click()
 Dim strFolder As String
 strFolder = BrowseForFolderFiles(Me.hwnd, LanguageStrings.OptionsGhostscriptResourceDirectoryPrompt)
 If Len(strFolder) = 0 Then Exit Sub
 strFolder = CompletePath(strFolder)
 With txtGSresource
  .Text = strFolder
  .ToolTipText = .Text
 End With
End Sub

Private Sub cmdGetTemppath_Click()
 Dim strFolder As String
 strFolder = BrowseForFolderFiles(Me.hwnd, LanguageStrings.OptionsPrintertempDirectoryPrompt)
 If Len(strFolder) = 0 Then Exit Sub
 strFolder = CompletePath(strFolder)
 With txtTemppath
  .Text = strFolder
  .ToolTipText = .Text
 End With
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
   ieb.Refresh
  End With
 End If
End Sub

Private Sub cmdRunProgramAfterSavingPrognameChoice_Click()
 Dim Filename As String
 Filename = BrowseForFolderFiles(Me.hwnd, LanguageStrings.OptionsProgramRunProgramAfterSavingCaption, False)
 If LenB(Filename) > 0 Then
  cmbRunProgramAfterSavingProgramname.Text = Filename
 End If
 If FileExists(Filename) = True Then
   If IsFileEditable(Filename) Then
     cmdRunProgramAfterSavingPrognameEdit.Enabled = True
    Else
     cmdRunProgramAfterSavingPrognameEdit.Enabled = False
   End If
  Else
   cmdRunProgramAfterSavingPrognameEdit.Enabled = False
 End If
End Sub

Private Sub cmdRunProgramAfterSavingPrognameEdit_Click()
 Dim Program As String, drv As String
 Program = RemoveLeadingAndTrailingQuotes(cmbRunProgramAfterSavingProgramname.Text)
 SplitPath Program, drv
 If LenB(drv) = 0 Then
  Program = ResolveRelativePath(Program, GetPDFCreatorApplicationPath)
 End If
 If FileExists(Program) = True Then
  If IsFileEditable(Program) Then
   EditDocument Program
  End If
 End If
End Sub

Private Sub cmdRunProgramBeforeSavingPrognameChoice_Click()
 Dim Filename As String
 Filename = BrowseForFolderFiles(Me.hwnd, LanguageStrings.OptionsProgramRunProgramBeforeSavingCaption, False)
 If LenB(Filename) > 0 Then
  cmbRunProgramBeforeSavingProgramname.Text = Filename
 End If
 If FileExists(Filename) = True Then
   If IsFileEditable(Filename) Then
     cmdRunProgramBeforeSavingPrognameEdit.Enabled = True
    Else
     cmdRunProgramBeforeSavingPrognameEdit.Enabled = False
   End If
  Else
   cmdRunProgramBeforeSavingPrognameEdit.Enabled = False
 End If
End Sub

Private Sub cmdRunProgramBeforeSavingPrognameEdit_Click()
 Dim Program As String, drv As String
 Program = RemoveLeadingAndTrailingQuotes(cmbRunProgramBeforeSavingProgramname.Text)
 SplitPath Program, drv
 If LenB(drv) = 0 Then
  Program = ResolveRelativePath(Program, GetPDFCreatorApplicationPath)
 End If
 If FileExists(Program) = True Then
  If IsFileEditable(Program) Then
   EditDocument Program
  End If
 End If
End Sub

Private Sub cmdSave_Click()
 Dim tRestart As Boolean
 tRestart = False
 If UCase$(Options.DirectoryGhostscriptBinaries) <> UCase$(txtGSbin.Text) Then
  tRestart = True
 End If
 CorrectCmbCharset
 GetOptions Me, Options
 CorrectOptions
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
 Dim tCharset As Long, tStr As String, tFontSize As Long, tFontname As String, _
  tFontCharset As Long
 If InStr(1, cmbCharset.Text, ",", vbTextCompare) > 0 Then
   tStr = Trim$(Mid$(cmbCharset.Text, 1, InStr(1, cmbCharset.Text, ",", vbTextCompare) - 1))
  Else
   tStr = Trim$(cmbCharset.Text)
 End If
 If Len(tStr) = 0 Then
  cmbCharset.Text = 0
  Exit Sub
 End If
 If IsNumeric(tStr) = False Then
  cmbCharset.Text = 0
  Exit Sub
 End If
 tCharset = tStr
 With cmdTest.Font
  tFontname = .Name
  tFontSize = .Size
  tFontCharset = .Charset
 End With
 SetFont Me, cmbFonts.List(cmbFonts.ListIndex), CLng(tStr), cmbProgramFontsize.Text
 cmbCharset.Text = tCharset
 SetFont frmMain, cmbFonts.List(cmbFonts.ListIndex), CLng(tStr), cmbProgramFontsize.Text
 ieb.Refresh
 With cmdTest.Font
  .Name = tFontname
  .Size = tFontSize
  .Charset = tFontCharset
 End With
 With cmdCancelTest
  .Font.Name = tFontname
  .Font.Size = tFontSize
  .Font.Charset = tFontCharset
  .Enabled = True
 End With
End Sub

Private Sub cmdTestpage_Click()
 PrintTestpage frmMain
End Sub

Private Sub cmdUsertempPath_Click()
 Dim Temppath As String
 Temppath = "<Temp>PDFCreator\"
 If DirExists(ResolveEnvironment(GetSubstFilename2(Temppath))) = False Then
  MakePath ResolveEnvironment(GetSubstFilename2(Temppath))
 End If
 With txtTemppath
  .Text = Temppath
  .ToolTipText = ResolveEnvironment(GetSubstFilename2(Temppath))
 End With
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF1 Then
  KeyCode = 0
    Call HTMLHelp_ShowTopic("html\generalsettings.htm")
'  MsgBox ieb.GetSelectedGroup & vbCrLf & ieb.GetSelectedItem
   Select Case ieb.GetSelectedGroup
    Case 1
     Select Case ieb.GetSelectedItem
      Case 1
       Call HTMLHelp_ShowTopic("html\generalsettings.htm")
      Case 2
       Call HTMLHelp_ShowTopic("html\ghostscript.htm")
      Case 3
       Call HTMLHelp_ShowTopic("html\docproperties.htm")
      Case 4
       Call HTMLHelp_ShowTopic("html\savesettings.htm")
      Case 5
       Call HTMLHelp_ShowTopic("html\autosave.htm")
      Case 6
       Call HTMLHelp_ShowTopic("html\directories.htm")
      Case 7
       Call HTMLHelp_ShowTopic("html\fontsetting.htm")
      Case Else
       Call HTMLHelp_ShowTopic("html\generalsettings.htm")
     End Select
    Case 2
     Select Case ieb.GetSelectedItem
      Case 1
       Select Case tbstrPDFOptions.SelectedItem.Index
        Case 1
         Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
        Case 2
         Call HTMLHelp_ShowTopic("html\pdfcompression.htm")
        Case 3
         Call HTMLHelp_ShowTopic("html\pdffonts.htm")
        Case 4
         Call HTMLHelp_ShowTopic("html\pdfcolors.htm")
        Case 5
         Call HTMLHelp_ShowTopic("html\pdfsecurity.htm")
        Case Else
         Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
       End Select
      Case 2
       Call HTMLHelp_ShowTopic("html\pngsettings.htm")
      Case 3
       Call HTMLHelp_ShowTopic("html\jpegsettings.htm")
      Case 4
       Call HTMLHelp_ShowTopic("html\bmpsettings.htm")
      Case 5
       Call HTMLHelp_ShowTopic("html\pcxsettings.htm")
      Case 6
       Call HTMLHelp_ShowTopic("html\tiffsettings.htm")
      Case 7
       Call HTMLHelp_ShowTopic("html\pssettings.htm")
      Case 8
       Call HTMLHelp_ShowTopic("html\epssettings.htm")
      Case Else
       Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
     End Select
   End Select
 End If
End Sub

Private Sub Form_Load()
 Const fraPDFTop = 1360, fraPDFLeft = 2960
 Dim pic As New StdPicture, i As Long, tStr As String, gsvers As Collection, _
  fc As Long, reg As clsRegistry, tsf() As String, tStr2 As String, Files As Collection, _
  Path As String, Filename As String, Ext As String, p As Printer
 
 Me.Icon = LoadResPicture(2120, vbResIcon)
 KeyPreview = True

 With Screen
  .MousePointer = vbHourglass
  Move (.Width - Width) / 2, (.Height - Height) / 2
 End With

 SetFrames

 With dmFraDescription
  .Caption = LanguageStrings.OptionsTreeProgram
  .Visible = True
 End With
 tbstrProgGeneral.Visible = True
 With dmFraProgGeneral1
  .Visible = True
  .Top = dmFraDescription.Top + dmFraDescription.Height + 50
  .Left = dmFraDescription.Left
  dmFraProgGeneral2.Top = .Top
  dmFraProgGeneral2.Left = .Left
  dmFraProgGeneral2.Width = .Width
  dmFraShellIntegration.Width = dmFraProgGeneral2.Width
  dmFraProgGhostscript.Top = .Top
  dmFraProgGhostscript.Left = .Left
  dmFraProgGhostscript.Width = .Width
  dmFraProgAutosave.Top = .Top
  dmFraProgAutosave.Left = .Left
  dmFraProgAutosave.Width = .Width
  dmFraProgDirectories.Top = .Top
  dmFraProgDirectories.Left = .Left
  dmFraProgDirectories.Width = .Width
  dmFraProgDocument.Top = .Top
  dmFraProgDocument.Left = .Left
  dmFraProgDocument.Width = .Width
  dmFraProgSave.Top = .Top
  dmFraProgSave.Left = .Left
  dmFraProgSave.Width = .Width
  dmFraFilenameSubstitutions.Top = dmFraProgSave.Top + dmFraProgSave.Height + 50
  dmFraFilenameSubstitutions.Left = .Left
  dmFraFilenameSubstitutions.Width = .Width
  dmFraProgFont.Top = .Top
  dmFraProgFont.Left = .Left
  dmFraProgFont.Width = .Width
  dmFraProgActions.Top = .Top
  dmFraProgActions.Left = .Left
  dmFraProgActions.Width = .Width
  dmFraProgPrint.Top = .Top
  dmFraProgPrint.Left = .Left
  dmFraProgPrint.Width = .Width
  dmFraBitmapGeneral.Top = .Top
  dmFraBitmapGeneral.Left = .Left
  dmFraBitmapGeneral.Width = .Width
  dmFraPSGeneral.Top = .Top
  dmFraPSGeneral.Left = .Left
  dmFraPSGeneral.Width = .Width

  dmFraProgActionsRunProgramAfterSaving.Top = dmFraProgActionsRunProgramBeforeSaving.Top
  dmFraProgActionsRunProgramAfterSaving.Left = dmFraProgActionsRunProgramBeforeSaving.Left

  cmdCancel.Left = .Left
  cmdReset.Left = .Left + (.Width - cmdReset.Width) / 2
  cmdSave.Left = .Left + .Width - cmdSave.Width
 End With

 With tbstrProgGeneral
  .Top = dmFraDescription.Top + dmFraDescription.Height + 50
  .Left = dmFraDescription.Left
  .Height = cmdCancel.Top - tbstrProgGeneral.Top - 50
  .Width = dmFraDescription.Width
 End With
 
 With dmFraProgGeneral1
  .Top = tbstrProgGeneral.ClientTop + 100
  .Left = tbstrProgGeneral.Left + (tbstrProgGeneral.Width - .Width) / 2
 End With
 With dmFraProgGeneral2
  .Top = tbstrProgGeneral.ClientTop + 100
  .Left = tbstrProgGeneral.Left + (tbstrProgGeneral.Width - .Width) / 2
 End With
 With dmFraShellIntegration
  .Top = dmFraProgGeneral2.Top + dmFraProgGeneral2.Height + 50
  .Left = tbstrProgGeneral.Left + (tbstrProgGeneral.Width - .Width) / 2
 End With
 
 With tbstrPDFOptions
  .Top = dmFraDescription.Top + dmFraDescription.Height + 50
  .Left = dmFraDescription.Left
  .Height = cmdCancel.Top - tbstrPDFOptions.Top - 50
  .Width = dmFraDescription.Width
 End With

 With dmFraPDFGeneral
  .Top = tbstrPDFOptions.ClientTop + 100
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
 End With

 cmbEPSLanguageLevel.Top = cmbPSLanguageLevel.Top
 cmbEPSLanguageLevel.Left = cmbPSLanguageLevel.Left

 ieb.DisableUpdates True
 ieb.ClearStructure
 ieb.SetImageList imlIeb
 With LanguageStrings
  ieb.AddGroup "Program", .OptionsTreeProgram, 0
  ieb.AddItem "Program", "General", .OptionsProgramGeneralSymbol, 1
  ieb.AddItem "Program", "Ghostscript", .OptionsProgramGhostscriptSymbol, 2
  ieb.AddItem "Program", "Document", .OptionsProgramDocumentSymbol, 3
  ieb.AddItem "Program", "Save", .OptionsProgramSaveSymbol, 4
  ieb.AddItem "Program", "AutoSave", .OptionsProgramAutosaveSymbol, 5
  ieb.AddItem "Program", "Directories", .OptionsProgramDirectoriesSymbol, 6
  ieb.AddItem "Program", "Actions", .OptionsProgramActionsSymbol, 7
  ieb.AddItem "Program", "Print", .OptionsProgramPrintSymbol, 8
  ieb.AddItem "Program", "Fonts", .OptionsProgramFontSymbol, 9
  ieb.AddGroup "Formats", .OptionsTreeFormats, 0
  ieb.AddItem "Formats", "PDF", .OptionsPDFSymbol, 10
  ieb.AddItem "Formats", "PNG", .OptionsPNGSymbol, 11
  ieb.AddItem "Formats", "JPEG", .OptionsJPEGSymbol, 12
  ieb.AddItem "Formats", "BMP", .OptionsBMPSymbol, 13
  ieb.AddItem "Formats", "PCX", .OptionsPCXSymbol, 14
  ieb.AddItem "Formats", "TIFF", .OptionsTIFFSymbol, 15
  ieb.AddItem "Formats", "PS", .OptionsPSSymbol, 16
  ieb.AddItem "Formats", "EPS", .OptionsEPSSymbol, 17
  ieb.DisableUpdates False

  Set picOptions = LoadResPicture(2101, vbResIcon)
  dmFraProgGeneral1.Visible = True

  dmFraProgGeneral1.Caption = .OptionsProgramGeneralDescription1
  dmFraProgGeneral2.Caption = .OptionsProgramGeneralDescription2
  With tbstrProgGeneral.Tabs
   .Clear
   .Add , , LanguageStrings.OptionsProgramGeneralDescription1
   .Add , , LanguageStrings.OptionsProgramGeneralDescription2
  End With
  dmFraShellIntegration.Caption = .OptionsShellIntegration
  dmFraProgGhostscript.Caption = .OptionsProgramGhostscriptSymbol
  dmFraProgAutosave.Caption = .OptionsProgramAutosaveSymbol
  dmFraProgDirectories.Caption = .OptionsProgramDirectoriesSymbol
  dmFraProgDocument.Caption = .OptionsProgramDocumentSymbol
  dmFraProgFont.Caption = .OptionsProgramFontSymbol
  dmFraProgSave.Caption = .OptionsProgramSaveSymbol
  dmFraProgActions.Caption = .OptionsProgramActionsSymbol
  dmFraProgPrint.Caption = .OptionsProgramPrintSymbol

  cmdShellintegration(0).Caption = .OptionsShellIntegrationAdd
  cmdShellintegration(1).Caption = .OptionsShellIntegrationRemove
  If IsWin9xMe = False Then
   If IsAdmin = False Then
    cmdShellintegration(0).Enabled = False
    cmdShellintegration(1).Enabled = False
   End If
  End If
  
  lblSendMailMethod.Caption = .OptionsSendMailMethod
  cmbSendMailMethod.AddItem .OptionsSendMailMethodAutomatic
  cmbSendMailMethod.AddItem .OptionsSendMailMethodMapi
  cmbSendMailMethod.AddItem .OptionsSendMailMethodSendmailDLL

  lblGhostscriptversion.Caption = .OptionsGhostscriptversion
  lblAdditionalGhostscriptParameters.Caption = .OptionsAdditionalGhostscriptParameters
  lblAdditionalGhostscriptSearchpath.Caption = .OptionsAdditionalGhostscriptSearchpath
  chkAddWindowsFontpath.Caption = .OptionsAddWindowsFontpath

  lblSaveFilename.Caption = .OptionsSaveFilename
  lblSaveFilenameTokens.Caption = .OptionsSaveFilenameTokens
  dmFraFilenameSubstitutions.Caption = .OptionsSaveFilenameSubstitutions
  chkFilenameSubst.Caption = .OptionsSaveFilenameSubstitutionsTitle
  cmdFilenameSubst(0).Caption = .OptionsSaveFilenameAdd
  cmdFilenameSubst(1).Caption = .OptionsSaveFilenameChange
  cmdFilenameSubst(2).Caption = .OptionsSaveFilenameDelete

  chkSpaces.Caption = .OptionsRemoveSpaces
  chkNoConfirmMessageSwitchingDefaultprinter.Caption = .OptionsProgramSwitchingDefaultprinter
  chkNoProcessingAtStartup.Caption = .OptionsProgramNoProcessingAtStartup
  lblOptionsDesign.Caption = .OptionsProgramOptionsDesign
  cmbOptionsDesign.AddItem .OptionsProgramOptionsDesignGradient
  cmbOptionsDesign.AddItem .OptionsProgramOptionsDesignSimple
  chkShowAnimation.Caption = .OptionsProgramShowAnimation

  lblGSbin.Caption = .OptionsDirectoriesGSBin
  lblGSlib.Caption = .OptionsDirectoriesGSLibraries
  lblGSfonts.Caption = .OptionsDirectoriesGSFonts
  lblPrintTempPath.Caption = .OptionsDirectoriesTempPath

  chkOnePagePerFile.Caption = .OptionsOnePagePerFile
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
  chkAutosaveStartStandardProgram.Caption = .OptionsAutosaveStartStandardProgram
  chkAutosaveSendEmail.Caption = .OptionsSendEmailAfterAutosave
  lblStandardSaveformat.Caption = .OptionsStandardSaveFormat

  dmFraProgActionsRunProgramAfterSaving.Caption = .OptionsProgramRunProgramAfterSavingCaption
  chkRunProgramAfterSaving.Caption = .OptionsProgramRunProgramAfterSavingCaption
  lblRunProgramAfterSavingProgramname.Caption = .OptionsProgramRunProgramAfterSavingProgram
  lblRunProgramAfterSavingProgramParameters.Caption = .OptionsProgramRunProgramAfterSavingProgramParameters
  chkRunProgramAfterSavingWaitUntilReady.Caption = .OptionsProgramRunProgramAfterSavingWaitUntilReady
  lblRunProgramAfterSavingWindowstyle.Caption = .OptionsProgramRunProgramAfterSavingWindowstyle
  cmbRunProgramAfterSavingWindowstyle.AddItem .OptionsProgramRunProgramAfterSavingWindowstyleHide
  cmbRunProgramAfterSavingWindowstyle.AddItem .OptionsProgramRunProgramAfterSavingWindowstyleNormalFocus
  cmbRunProgramAfterSavingWindowstyle.AddItem .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedFocus
  cmbRunProgramAfterSavingWindowstyle.AddItem .OptionsProgramRunProgramAfterSavingWindowstyleMaximizedFocus
  cmbRunProgramAfterSavingWindowstyle.AddItem .OptionsProgramRunProgramAfterSavingWindowstyleNormalNoFocus
  cmbRunProgramAfterSavingWindowstyle.AddItem .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedNoFocus

  With tbstrProgActions.Tabs
   .Clear
   .Add , , LanguageStrings.OptionsProgramRunProgramBeforeSavingCaption
   .Add , , LanguageStrings.OptionsProgramRunProgramAfterSavingCaption
  End With

  dmFraProgActionsRunProgramBeforeSaving.Caption = .OptionsProgramRunProgramBeforeSavingCaption
  chkRunProgramBeforeSaving.Caption = .OptionsProgramRunProgramBeforeSavingCaption
  lblRunProgramBeforeSavingProgramname.Caption = .OptionsProgramRunProgramBeforeSavingProgram
  lblRunProgramBeforeSavingProgramParameters.Caption = .OptionsProgramRunProgramBeforeSavingProgramParameters
  chkRunProgramBeforeSavingWaitUntilReady.Caption = .OptionsProgramRunProgramBeforeSavingWaitUntilReady
  lblRunProgramBeforeSavingWindowstyle.Caption = .OptionsProgramRunProgramBeforeSavingWindowstyle
  cmbRunProgramBeforeSavingWindowstyle.AddItem .OptionsProgramRunProgramBeforeSavingWindowstyleHide
  cmbRunProgramBeforeSavingWindowstyle.AddItem .OptionsProgramRunProgramBeforeSavingWindowstyleNormalFocus
  cmbRunProgramBeforeSavingWindowstyle.AddItem .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedFocus
  cmbRunProgramBeforeSavingWindowstyle.AddItem .OptionsProgramRunProgramBeforeSavingWindowstyleMaximizedFocus
  cmbRunProgramBeforeSavingWindowstyle.AddItem .OptionsProgramRunProgramBeforeSavingWindowstyleNormalNoFocus
  cmbRunProgramBeforeSavingWindowstyle.AddItem .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedNoFocus

  chkPrintAfterSaving.Caption = .OptionsPrintAfterSaving
  lblPrintAfterSavingPrinter.Caption = .OptionsPrintAfterSavingPrinter
  
  For Each p In Printers
   cmbPrintAfterSavingPrinter.AddItem p.DeviceName
  Next p
  
  lblPrintAfterSavingQueryUser.Caption = .OptionsPrintAfterSavingQueryUser
  cmbPrintAfterSavingQueryUser.AddItem .OptionsPrintAfterSavingQueryUserOff
  cmbPrintAfterSavingQueryUser.AddItem .OptionsPrintAfterSavingQueryUserStandardPrinterDialog
  cmbPrintAfterSavingQueryUser.AddItem .OptionsPrintAfterSavingQueryUserPrinterSetupDialog
  cmbPrintAfterSavingQueryUser.AddItem .OptionsPrintAfterSavingQueryUserDefaultPrinter

  chkPrintAfterSavingNoCancel.Caption = .OptionsPrintAfterSavingNoCancel
  chkPrintAfterSavingDuplex.Caption = .OptionsPrintAfterSavingDuplex
  cmbPrintAfterSavingTumble.AddItem .OptionsPrintAfterSavingDuplexTumbleOff
  cmbPrintAfterSavingTumble.AddItem .OptionsPrintAfterSavingDuplexTumbleOn
  
  With cmbStandardSaveFormat
   .AddItem "PDF"
   .AddItem "PNG"
   .AddItem "JPEG"
   .AddItem "BMP"
   .AddItem "PCX"
   .AddItem "TIFF"
   .AddItem "PS"
   .AddItem "EPS"
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
  With cmbSaveFilenameTokens
   .AddItem "<Author>"
   .AddItem "<Computername>"
   .AddItem "<DateTime>"
   .AddItem "<Title>"
   .AddItem "<Username>"
   .AddItem "<REDMON_DOCNAME>"
   .AddItem "<REDMON_DOCNAME_FILE>"
   .AddItem "<REDMON_DOCNAME_PATH>"
   .AddItem "<REDMON_JOB>"
   .AddItem "<REDMON_MACHINE>"
   .AddItem "<REDMON_PORT>"
   .AddItem "<REDMON_PRINTER>"
   .AddItem "<REDMON_SESSIONID>"
   .AddItem "<REDMON_USER>"
   .ListIndex = 0
  End With
  With cmbAuthorTokens
   .AddItem "<Computername>"
   .AddItem "<ClientComputer>"
   .AddItem "<DateTime>"
   .AddItem "<Title>"
   .AddItem "<Username>"
   .AddItem "<REDMON_DOCNAME>"
   .AddItem "<REDMON_DOCNAME_FILE>"
   .AddItem "<REDMON_DOCNAME_PATH>"
   .AddItem "<REDMON_JOB>"
   .AddItem "<REDMON_MACHINE>"
   .AddItem "<REDMON_PORT>"
   .AddItem "<REDMON_PRINTER>"
   .AddItem "<REDMON_SESSIONID>"
   .AddItem "<REDMON_USER>"
   .ListIndex = 0
  End With
  With cmbAutoSaveFilenameTokens
   .AddItem "<Author>"
   .AddItem "<Computername>"
   .AddItem "<ClientComputer>"
   .AddItem "<DateTime>"
   .AddItem "<Title>"
   .AddItem "<Username>"
   .AddItem "<REDMON_DOCNAME>"
   .AddItem "<REDMON_DOCNAME_FILE>"
   .AddItem "<REDMON_DOCNAME_PATH>"
   .AddItem "<REDMON_JOB>"
   .AddItem "<REDMON_MACHINE>"
   .AddItem "<REDMON_PORT>"
   .AddItem "<REDMON_PRINTER>"
   .AddItem "<REDMON_SESSIONID>"
   .AddItem "<REDMON_USER>"
   .ListIndex = 0
  End With
  Me.Caption = .DialogPrinterOptions
  cmdCancel.Caption = .OptionsCancel
  cmdReset.Caption = .OptionsReset
  cmdSave.Caption = .OptionsSave
  tbstrPDFOptions.Tabs.Clear
  tbstrPDFOptions.Tabs.Add , "General", .OptionsPDFGeneral
  tbstrPDFOptions.Tabs.Add , "Compression", .OptionsPDFCompression
  tbstrPDFOptions.Tabs.Add , "Fonts", .OptionsPDFFonts
  tbstrPDFOptions.Tabs.Add , "Colors", .OptionsPDFColors
  tbstrPDFOptions.Tabs.Add , "Security", .OptionsPDFSecurity
  dmFraPDFGeneral.Caption = .OptionsPDFGeneralCaption
  chkPDFOptimize.Caption = .OptionsPDFOptimize
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

  dmfraPDFCompress.Caption = .OptionsPDFCompressionCaption
  chkPDFTextComp.Caption = .OptionsPDFCompressionTextComp
  dmFraPDFColor.Caption = .OptionsPDFCompressionColor
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
'  cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp08
  cmbPDFColorResample.Clear
  cmbPDFColorResample.AddItem .OptionsPDFCompressionColorResample01
  cmbPDFColorResample.AddItem .OptionsPDFCompressionColorResample02
'  cmbPDFColorResample.AddItem .OptionsPDFCompressionColorResample03
  dmFraPDFGrey.Caption = .OptionsPDFCompressionGrey
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
'  cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp08
  cmbPDFGreyResample.Clear
  cmbPDFGreyResample.AddItem .OptionsPDFCompressionGreyResample01
  cmbPDFGreyResample.AddItem .OptionsPDFCompressionGreyResample02
'  cmbPDFGreyResample.AddItem .OptionsPDFCompressionGreyResample03
  dmFraPDFMono.Caption = .OptionsPDFCompressionMono
  chkPDFMonoComp.Caption = .OptionsPDFCompressionMonoComp
  chkPDFMonoResample.Caption = .OptionsPDFCompressionMonoResample
  lblPDFMonoRes.Caption = .OptionsPDFCompressionMonoRes
  cmbPDFMonoComp.Clear
  cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp01
  cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp02
  cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp03
'  cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp04
  cmbPDFMonoResample.Clear
  cmbPDFMonoResample.AddItem .OptionsPDFCompressionMonoResample01
  cmbPDFMonoResample.AddItem .OptionsPDFCompressionMonoResample02
'  cmbPDFMonoResample.AddItem .OptionsPDFCompressionMonoResample03

  dmFraPDFFonts.Caption = .OptionsPDFFontsCaption
  chkPDFEmbedAll.Caption = .OptionsPDFFontsEmbedAll
  chkPDFSubSetFonts.Caption = .OptionsPDFFontsSubSetFonts

  dmFraPDFColors.Caption = .OptionsPDFColorsCaption
  chkPDFCMYKtoRGB.Caption = .OptionsPDFColorsCMYKtoRGB
  dmFraPDFColorOptions.Caption = .OptionsPDFColorsColorOptions
  chkPDFPreserveOverprint.Caption = .OptionsPDFColorsPreserveOverprint
  chkPDFPreserveTransfer.Caption = .OptionsPDFColorsPreserveTransfer
  chkPDFPreserveHalftone.Caption = .OptionsPDFColorsPreserveHalftone
  cmbPDFColorModel.Clear
  cmbPDFColorModel.AddItem .OptionsPDFColorsColorModel01
  cmbPDFColorModel.AddItem .OptionsPDFColorsColorModel02
  cmbPDFColorModel.AddItem .OptionsPDFColorsColorModel03

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

  dmFraBitmapGeneral.Caption = .OptionsImageSettings
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

 cmdFilenameSubst(0).Top = lsvFilenameSubst.Top
 cmdFilenameSubst(1).Top = lsvFilenameSubst.Top + (lsvFilenameSubst.Height - cmdFilenameSubst(1).Height) / 2
 cmdFilenameSubst(2).Top = lsvFilenameSubst.Top + lsvFilenameSubst.Height - cmdFilenameSubst(2).Height

 If chkUseStandardAuthor.Value = 1 Then
   txtStandardAuthor.Enabled = True
   txtStandardAuthor.BackColor = &H80000005
  Else
   txtStandardAuthor.Enabled = False
   txtStandardAuthor.BackColor = &H8000000F
 End If
 With Options
  SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
 End With
 ieb.Refresh
 If chkUseAutosave.Value = 1 Then
   ViewAutosave True
  Else
   ViewAutosave False
 End If
 If chkPrintAfterSaving.Value = 1 Then
   ViewPrintAfterSaving True
  Else
   ViewPrintAfterSaving False
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
  .ToolTipText = ResolveEnvironment(GetSubstFilename2(.Text))
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

 If Options.RunProgramAfterSaving Then
   ViewRunProgramAfterSaving True
  Else
   ViewRunProgramAfterSaving False
 End If
 If Options.RunProgramBeforeSaving Then
   ViewRunProgramBeforeSaving True
  Else
   ViewRunProgramBeforeSaving False
 End If

 Set Files = GetFiles(GetPDFCreatorApplicationPath & "Scripts\RunProgramAfterSaving\", "*.*", SortedByName)
 For i = 1 To Files.Count
  tsf = Split(Files(i), "|")
  SplitPath tsf(1), , Path, Filename, , Ext
  If UCase$(Ext) <> "TXT" And UCase$(Ext) <> "PDF" And UCase$(Ext) <> "PNG" And _
   UCase$(Ext) <> "JPG" And UCase$(Ext) <> "BMP" And UCase$(Ext) <> "PCX" And _
   UCase$(Ext) <> "TIF" And UCase$(Ext) <> "EPS" And UCase$(Ext) <> "PS" Then
   If UCase$(tsf(0)) <> UCase$(GetPDFCreatorApplicationPath & "Scripts\RunProgramAfterSaving\") Then
     cmbRunProgramAfterSavingProgramname.AddItem tsf(0)
    Else
     cmbRunProgramAfterSavingProgramname.AddItem Filename
   End If
  End If
 Next i

 Set Files = GetFiles(GetPDFCreatorApplicationPath & "Scripts\RunProgramBeforeSaving\", "*.*", SortedByName)
 For i = 1 To Files.Count
  tsf = Split(Files(i), "|")
  SplitPath tsf(1), , Path, Filename, , Ext
  If UCase$(Ext) <> "TXT" And UCase$(Ext) <> "PDF" And UCase$(Ext) <> "PNG" And _
   UCase$(Ext) <> "JPG" And UCase$(Ext) <> "BMP" And UCase$(Ext) <> "PCX" And _
   UCase$(Ext) <> "TIF" And UCase$(Ext) <> "EPS" And UCase$(Ext) <> "PS" Then
   If UCase$(tsf(0)) <> UCase$(GetPDFCreatorApplicationPath & "Scripts\RunProgramBeforeSaving\") Then
     cmbRunProgramBeforeSavingProgramname.AddItem tsf(0)
    Else
     cmbRunProgramBeforeSavingProgramname.AddItem Filename
   End If
  End If
 Next i

 tStr2 = CompletePath(UCase$(Trim$(Options.DirectoryGhostscriptBinaries)))
 cmbGhostscript.Clear: Set reg = New clsRegistry
 reg.hkey = HKEY_LOCAL_MACHINE

 Set gsvers = GetAllGhostscriptversions

 If gsvers.Count = 0 Then
   cmbGhostscript.Enabled = False
  Else
   For i = 1 To gsvers.Count
    cmbGhostscript.AddItem gsvers.Item(i)
   Next i
   cmbGhostscript.ListIndex = cmbGhostscript.ListCount - 1
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
      If InStr(UCase$(cmbGhostscript.List(i)), "GPL") Then
       reg.KeyRoot = "SOFTWARE\GPL Ghostscript"
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

 tbstrPDFOptions.ZOrder 1
 tbstrProgActions.ZOrder 1

 If ShowOnlyOptions = True Then
  FormInTaskbar Me, True, True
  Caption = "PDFCreator - " & Caption
 End If

 ShowAcceleratorsInForm Me, True

 ShowOptions Me, Options
 Timer1.Enabled = True
 Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Resize()
 With cmbCharset
  .Top = cmbFonts.Top
  .Left = lblProgcharset.Left
  .Width = 2295
  .SelStart = 0
  .SelLength = 0
 End With
 With cmbProgramFontsize
  .Top = cmbFonts.Top
  .Left = lblSize.Left
  .Width = 765
  .SelStart = 0
  .SelLength = 0
 End With
 With cmbGhostscript
  .Top = lblGhostscriptversion.Top + lblGhostscriptversion.Height + 20
  .Left = lblGhostscriptversion.Left
  .Width = 4215
 End With
End Sub

Private Sub ieb_ItemClick(sGroup As String, sItemKey As String)
 Dim ctl As Control
 lblJPEGQuality.Visible = False
 cmbPNGColors.Visible = False
 cmbJPEGColors.Visible = False
 cmbBMPColors.Visible = False
 cmbPCXColors.Visible = False
 cmbTIFFColors.Visible = False
 tbstrPDFOptions.Visible = False
 tbstrProgGeneral.Visible = False
 For Each ctl In Controls
  If TypeOf ctl Is dmFrame Then
   ctl.Visible = False
   ctl.Enabled = False
  End If
 Next
 dmFraDescription.Visible = True
 dmFraDescription.Enabled = True
 tbstrPDFOptions.Enabled = False
 txtJPEGQuality.Visible = False
 lblJPEQQualityProzent.Visible = False
 dmFraPSGeneral.Visible = False
 cmbPSLanguageLevel.Visible = False
 cmbEPSLanguageLevel.Visible = False

 Select Case UCase$(sGroup)
  Case "PROGRAM"
   Select Case UCase$(sItemKey)
    Case "GENERAL"
     Set picOptions = LoadResPicture(2101, vbResIcon)
     lblOptions = LanguageStrings.OptionsProgramGeneralDescription
     tbstrProgGeneral.Enabled = True
     tbstrProgGeneral.Visible = True
     Select Case tbstrProgGeneral.SelectedItem.Index
      Case 1
       dmFraProgGeneral1.Enabled = True
       dmFraProgGeneral1.Visible = True
      Case 2
       dmFraProgGeneral2.Enabled = True
       dmFraProgGeneral2.Visible = True
       dmFraShellIntegration.Enabled = True
       dmFraShellIntegration.Visible = True
     End Select
     dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
    Case "GHOSTSCRIPT"
     Set picOptions = LoadResPicture(2119, vbResIcon)
     lblOptions = LanguageStrings.OptionsProgramGhostscriptDescription
     dmFraProgGhostscript.Enabled = True
     dmFraProgGhostscript.Visible = True
     dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
    Case "DOCUMENT"
     Set picOptions = LoadResPicture(2105, vbResIcon)
     lblOptions = LanguageStrings.OptionsProgramDocumentDescription
     dmFraProgDocument.Enabled = True
     dmFraProgDocument.Visible = True
     dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
    Case "SAVE"
     Set picOptions = LoadResPicture(2106, vbResIcon)
     lblOptions = LanguageStrings.OptionsProgramSaveDescription
     dmFraProgSave.Enabled = True
     dmFraProgSave.Visible = True
     dmFraFilenameSubstitutions.Visible = True
     dmFraFilenameSubstitutions.Enabled = True
     dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
    Case "AUTOSAVE"
     Set picOptions = LoadResPicture(2103, vbResIcon)
     lblOptions = LanguageStrings.OptionsProgramAutosaveDescription
     dmFraProgAutosave.Enabled = True
     dmFraProgAutosave.Visible = True
     dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
    Case "DIRECTORIES"
     Set picOptions = LoadResPicture(2104, vbResIcon)
     lblOptions = LanguageStrings.OptionsProgramDirectoriesDescription
     dmFraProgDirectories.Enabled = True
     dmFraProgDirectories.Visible = True
     dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
    Case "ACTIONS"
     Set picOptions = LoadResPicture(2121, vbResIcon)
     lblOptions = LanguageStrings.OptionsProgramActionsDescription
     dmFraProgActions.Enabled = True
     dmFraProgActions.Visible = True
     ViewProgActions
     dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
    Case "PRINT"
     Set picOptions = LoadResPicture(2122, vbResIcon)
     lblOptions = LanguageStrings.OptionsProgramPrintDescription
     dmFraProgPrint.Enabled = True
     dmFraProgPrint.Visible = True
     dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
    Case "FONTS"
     Set picOptions = LoadResPicture(2102, vbResIcon)
     lblOptions = LanguageStrings.OptionsProgramFontDescription
     dmFraProgFont.Enabled = True
     dmFraProgFont.Visible = True
     dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
   End Select
  Case "FORMATS"
   Select Case UCase$(sItemKey)
    Case "PDF"
     Set picOptions = LoadResPicture(2111, vbResIcon)
     lblOptions = LanguageStrings.OptionsPDFDescription
     tbstrPDFOptions.Enabled = True
     tbstrPDFOptions.Visible = True
     dmFraPDFGeneral.Enabled = True
     dmFraPDFGeneral.Visible = True
     dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
     dmFraPDFGeneral.Enabled = True
    Case "PNG"
     Set picOptions = LoadResPicture(2112, vbResIcon)
     lblOptions = LanguageStrings.OptionsPNGDescription
     dmFraBitmapGeneral.Enabled = True
     dmFraBitmapGeneral.Visible = True
     cmbPNGColors.Visible = True
     dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
    Case "JPEG"
     Set picOptions = LoadResPicture(2113, vbResIcon)
     lblOptions = LanguageStrings.OptionsJPEGDescription
     dmFraBitmapGeneral.Enabled = True
     dmFraBitmapGeneral.Visible = True
     lblJPEGQuality.Caption = LanguageStrings.OptionsJPEGQuality
     lblJPEGQuality.Visible = True
     txtJPEGQuality.Visible = True
     lblJPEQQualityProzent.Visible = True
     lblJPEQQualityProzent.Left = txtJPEGQuality.Left + txtJPEGQuality.Width + 100
     cmbJPEGColors.Visible = True
     dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
    Case "BMP"
     Set picOptions = LoadResPicture(2114, vbResIcon)
     lblOptions = LanguageStrings.OptionsBMPDescription
     dmFraBitmapGeneral.Enabled = True
     dmFraBitmapGeneral.Visible = True
     cmbBMPColors.Visible = True
     dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
    Case "PCX"
     Set picOptions = LoadResPicture(2115, vbResIcon)
     lblOptions = LanguageStrings.OptionsPCXDescription
     dmFraBitmapGeneral.Enabled = True
     dmFraBitmapGeneral.Visible = True
     cmbPCXColors.Visible = True
     dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
    Case "TIFF"
     Set picOptions = LoadResPicture(2116, vbResIcon)
     lblOptions = LanguageStrings.OptionsTIFFDescription
     dmFraBitmapGeneral.Enabled = True
     dmFraBitmapGeneral.Visible = True
     cmbTIFFColors.Visible = True
     dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
    Case "PS"
     Set picOptions = LoadResPicture(2117, vbResIcon)
     lblOptions.Caption = LanguageStrings.OptionsPSDescription
     dmFraPSGeneral.Enabled = True
     dmFraPSGeneral.Visible = True
     cmbPSLanguageLevel.Visible = True
     dmFraPSGeneral.Caption = LanguageStrings.OptionsPSDescription
     dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
    Case "EPS"
     Set picOptions = LoadResPicture(2118, vbResIcon)
     lblOptions.Caption = LanguageStrings.OptionsEPSDescription
     dmFraPSGeneral.Enabled = True
     dmFraPSGeneral.Visible = True
     cmbEPSLanguageLevel.Visible = True
     dmFraPSGeneral.Caption = LanguageStrings.OptionsEPSDescription
     dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
   End Select
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
 dmFraPDFGeneral.Visible = False
 dmfraPDFCompress.Visible = False
 dmFraPDFFonts.Visible = False
 dmFraPDFColors.Visible = False
 dmFraPDFColorOptions.Visible = False
 dmFraPDFSecurity.Visible = False
 dmFraPDFGeneral.Enabled = False
 dmfraPDFCompress.Enabled = False
 dmFraPDFFonts.Enabled = False
 dmFraPDFColors.Enabled = False
 dmFraPDFColorOptions.Enabled = False
 dmFraPDFSecurity.Enabled = False
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
   If SecurityIsPossible = False Then
    MsgBox LanguageStrings.MessagesMsg19
   End If
 End Select
End Sub

Private Sub tbstrProgActions_Click()
 ViewProgActions
End Sub

Private Sub tbstrProgGeneral_Click()
 Select Case tbstrProgGeneral.SelectedItem.Index
  Case 1
   dmFraProgGeneral1.Enabled = True
   dmFraProgGeneral1.Visible = True
   dmFraProgGeneral2.Enabled = False
   dmFraProgGeneral2.Visible = False
   dmFraShellIntegration.Enabled = False
   dmFraShellIntegration.Visible = False
  Case 2
   dmFraProgGeneral2.Enabled = True
   dmFraProgGeneral2.Visible = True
   dmFraShellIntegration.Enabled = True
   dmFraShellIntegration.Visible = True
   dmFraProgGeneral1.Enabled = False
   dmFraProgGeneral1.Visible = False
 End Select
End Sub

Private Sub Timer1_Timer()
 Dim i As Long, fi As Long, tStr As String, SMF As Collection, _
  cSystem As clsSystem, ctl As Control
 Timer1.Enabled = False
 Set cSystem = New clsSystem
 Set SMF = cSystem.GetSystemFont(Me, Menu)
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
 With cmbProgramFontsize
  .AddItem "8"
  .AddItem "9"
  .AddItem "10"
  .AddItem "11"
  .AddItem "12"
  .AddItem "14"
  .AddItem "16"
  .AddItem "18"
  .AddItem "20"
  .AddItem "22"
  .AddItem "24"
  .AddItem "26"
  .AddItem "28"
  .AddItem "36"
  .AddItem "48"
  .AddItem "72"
 End With
 cmbProgramFontsize.Text = 8
 cmbCharset.Text = cmbCharset.ItemData(0)
 cmbCharset.Text = Options.ProgramFontCharset
 For Each ctl In Controls
  If TypeOf ctl Is ComboBox Then
   ComboSetListWidth ctl
  End If
 Next ctl

 SetOptimalComboboxHeigth cmbCharset, Me
 SetOptimalComboboxHeigth cmbProgramFontsize, Me
 SetOptimalComboboxHeigth cmbGhostscript, Me

 Form_Resize

 cmbProgramFontsize.Width = txtTest.Width - _
  (cmbProgramFontsize.Left - txtTest.Left)

 If fi >= 0 Then
  cmbFonts.ListIndex = fi
  cmbCharset.Text = SMF(1)(2)
  cmbProgramFontsize.Text = SMF(1)(1)
  txtTest.Font = cmbFonts.List(cmbFonts.ListIndex)
  txtTest.Font.Charset = cmbCharset.Text
 End If

 ShowOptions Me, Options

 If Options.UseAutosaveDirectory = "1" Then
   ViewAutosaveDirectory True
  Else
   ViewAutosaveDirectory False
 End If
 If Options.UseAutosave = "1" Then
   ViewAutosave True
  Else
   ViewAutosave False
 End If

 CheckCmdFilenameSubst
 CorrectCmbCharset
 tbstrProgActions.Tabs(2).Selected = True
' Call ieb_ItemClick("PROGRAM", "GENERAL")
End Sub

Private Sub txtAutosaveDirectory_Change()
 txtAutosaveDirectory.ToolTipText = txtAutosaveDirectory.Text
 With txtAutoSaveDirectoryPreview
  .Text = GetSubstFilename2(txtAutosaveDirectory.Text)
  .ToolTipText = .Text
  If IsValidPath(.Text) = False Then
    .ForeColor = vbRed
   Else
    .ForeColor = &H80000008
  End If
 End With
End Sub

Private Sub txtAutosaveFilename_Change()
 Dim Ext As String
 txtAutosaveFilename.ToolTipText = txtAutosaveFilename.Text
 With txtAutoSaveFilenamePreview
  .Text = GetSubstFilename("B:\dummy.dum", txtAutosaveFilename.Text, , True) & GetAutosaveFormatExtension
  .ToolTipText = .Text
  If IsValidPath("C:\" & .Text) = False Then
    .ForeColor = vbRed
   Else
    .ForeColor = &H80000008
  End If
 End With
End Sub

Private Sub cmbProgramFontSize_Change()
 Dim tL As Long
If Trim$(cmbProgramFontsize.Text) = vbNullString Then
  cmbProgramFontsize.Text = 8
 End If
 tL = CLng(cmbProgramFontsize.Text)
 If tL <= 0 Then
  tL = 1
 End If
 If tL > 72 Then
  tL = 72
 End If
 cmbProgramFontsize.Text = tL
 txtTest.Font.Size = tL
End Sub

Private Sub cmbProgramFontSize_KeyPress(KeyAscii As Integer)
 Dim allow As String, tStr As String

 allow = "0123456789" & Chr$(8) & Chr$(13)

 tStr = Chr$(KeyAscii)

 If InStr(1, allow, tStr) = 0 Then
   KeyAscii = 0
 End If
End Sub

Private Sub cmbProgramFontsize_Click()
 Dim tL As Long
If Trim$(cmbProgramFontsize.Text) = vbNullString Then
  cmbProgramFontsize.Text = 8
 End If
 tL = CLng(cmbProgramFontsize.Text)
 If tL <= 0 Then
  tL = 1
 End If
 If tL > 72 Then
  tL = 72
 End If
 cmbProgramFontsize.Text = tL
 txtTest.Font.Size = tL
End Sub

Private Sub ViewAutosave(ViewIt As Boolean)
 lblAutosaveformat.Enabled = ViewIt
 cmbAutosaveFormat.Enabled = ViewIt
 lblAutosaveFilename.Enabled = ViewIt
 txtAutosaveFilename.Enabled = ViewIt
 txtAutoSaveFilenamePreview.Enabled = ViewIt
 lblAutosaveFilenameTokens.Enabled = ViewIt
 cmbAutoSaveFilenameTokens.Enabled = ViewIt
 chkUseAutosaveDirectory.Enabled = ViewIt
 txtAutoSaveDirectoryPreview.Enabled = ViewIt
 chkAutosaveStartStandardProgram.Enabled = ViewIt
 chkAutosaveSendEmail.Enabled = ViewIt

 If ViewIt Then
   cmbAutosaveFormat.BackColor = &H80000005
   cmbAutoSaveFilenameTokens.BackColor = &H80000005
   txtAutosaveFilename.BackColor = &H80000005
   txtAutosaveDirectory.BackColor = &H80000005
  Else
   cmbAutosaveFormat.BackColor = &H8000000F
   cmbAutoSaveFilenameTokens.BackColor = &H8000000F
   txtAutosaveFilename.BackColor = &H8000000F
   txtAutosaveDirectory.BackColor = &H8000000F
 End If
 If chkUseAutosaveDirectory.Value = 1 And ViewIt Then
   ViewAutosaveDirectory True
  Else
   ViewAutosaveDirectory False
 End If
End Sub

Private Sub ViewAutosaveDirectory(ViewIt As Boolean)
 txtAutosaveDirectory.Enabled = ViewIt
 txtAutoSaveDirectoryPreview.Enabled = ViewIt
 cmdGetAutosaveDirectory.Enabled = ViewIt
 If ViewIt = True Then
   txtAutosaveDirectory.BackColor = &H80000005
  Else
   txtAutosaveDirectory.BackColor = &H8000000F
 End If
End Sub

Private Sub ViewProgActions()
 Select Case tbstrProgActions.SelectedItem.Index
  Case 1
   dmFraProgActionsRunProgramBeforeSaving.Visible = True
   dmFraProgActionsRunProgramBeforeSaving.Enabled = True
   dmFraProgActionsRunProgramAfterSaving.Visible = False
   dmFraProgActionsRunProgramAfterSaving.Enabled = False
  Case 2
   dmFraProgActionsRunProgramAfterSaving.Visible = True
   dmFraProgActionsRunProgramAfterSaving.Enabled = True
   dmFraProgActionsRunProgramBeforeSaving.Visible = False
   dmFraProgActionsRunProgramBeforeSaving.Enabled = False
 End Select
End Sub

Private Sub ViewRunProgramAfterSaving(ViewIt As Boolean)
 lblRunProgramAfterSavingProgramname.Enabled = ViewIt
 cmbRunProgramAfterSavingProgramname.Enabled = ViewIt
 lblRunProgramAfterSavingProgramParameters.Enabled = ViewIt
 txtRunProgramAfterSavingProgramParameters.Enabled = ViewIt
 chkRunProgramAfterSavingWaitUntilReady.Enabled = ViewIt
 lblRunProgramAfterSavingWindowstyle.Enabled = ViewIt
 cmbRunProgramAfterSavingWindowstyle.Enabled = ViewIt
 cmdRunProgramAfterSavingPrognameChoice.Enabled = ViewIt
 cmdRunProgramAfterSavingPrognameEdit.Enabled = ViewIt

 If ViewIt Then
   cmbRunProgramAfterSavingProgramname.BackColor = &H80000005
   cmbRunProgramAfterSavingWindowstyle.BackColor = &H80000005
   txtRunProgramAfterSavingProgramParameters.BackColor = &H80000005
  Else
   cmbRunProgramAfterSavingProgramname.BackColor = &H8000000F
   cmbRunProgramAfterSavingWindowstyle.BackColor = &H8000000F
   txtRunProgramAfterSavingProgramParameters.BackColor = &H8000000F
 End If

 cmbRunProgramAfterSavingProgramname_Change
End Sub

Private Sub ViewRunProgramBeforeSaving(ViewIt As Boolean)
 lblRunProgramBeforeSavingProgramname.Enabled = ViewIt
 cmbRunProgramBeforeSavingProgramname.Enabled = ViewIt
 lblRunProgramBeforeSavingProgramParameters.Enabled = ViewIt
 txtRunProgramBeforeSavingProgramParameters.Enabled = ViewIt
 chkRunProgramBeforeSavingWaitUntilReady.Enabled = ViewIt
 lblRunProgramBeforeSavingWindowstyle.Enabled = ViewIt
 cmbRunProgramBeforeSavingWindowstyle.Enabled = ViewIt
 cmdRunProgramBeforeSavingPrognameChoice.Enabled = ViewIt
 cmdRunProgramBeforeSavingPrognameEdit.Enabled = ViewIt

 If ViewIt Then
   cmbRunProgramBeforeSavingProgramname.BackColor = &H80000005
   cmbRunProgramBeforeSavingWindowstyle.BackColor = &H80000005
   txtRunProgramBeforeSavingProgramParameters.BackColor = &H80000005
  Else
   cmbRunProgramBeforeSavingProgramname.BackColor = &H8000000F
   cmbRunProgramBeforeSavingWindowstyle.BackColor = &H8000000F
   txtRunProgramBeforeSavingProgramParameters.BackColor = &H8000000F
 End If

 cmbRunProgramBeforeSavingProgramname_Change
End Sub

Private Sub ViewPrintAfterSaving(ViewIt As Boolean)
 lblPrintAfterSavingPrinter.Enabled = ViewIt
 cmbPrintAfterSavingPrinter.Enabled = ViewIt
 lblPrintAfterSavingQueryUser.Enabled = ViewIt
 cmbPrintAfterSavingQueryUser.Enabled = ViewIt
 chkPrintAfterSavingNoCancel.Enabled = ViewIt
 chkPrintAfterSavingDuplex.Enabled = ViewIt

 If ViewIt Then
   cmbPrintAfterSavingPrinter.BackColor = &H80000005
   cmbPrintAfterSavingQueryUser.BackColor = &H80000005
  Else
   cmbPrintAfterSavingPrinter.BackColor = &H8000000F
   cmbPrintAfterSavingQueryUser.BackColor = &H8000000F
 End If
 
 If chkPrintAfterSavingDuplex.Value = 1 And ViewIt Then
   ViewPrintAfterTumple True
  Else
   ViewPrintAfterTumple False
 End If
End Sub

Private Sub ViewPrintAfterTumple(ViewIt As Boolean)
 cmbPrintAfterSavingTumble.Enabled = ViewIt
 
 If ViewIt Then
   cmbPrintAfterSavingTumble.BackColor = &H80000005
  Else
   cmbPrintAfterSavingTumble.BackColor = &H8000000F
 End If
End Sub

Private Sub UpdateSecurityFields()
 If cmbPDFCompat.ListIndex < 2 Then
  optEncLow.Value = True
 End If
 If chkUseSecurity.Value = False Then
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
   If cmbPDFCompat.ListIndex >= 2 Then
     optEncHigh.Enabled = True
    Else
     optEncHigh.Enabled = False
   End If
   optEncLow.Enabled = True

   dmFraSecurityPass.Enabled = True
   chkUserPass.Enabled = True
   chkOwnerPass.Enabled = True

   dmFraPDFPermissions.Enabled = True
   chkAllowPrinting.Enabled = True
   chkAllowCopy.Enabled = True
   chkAllowModifyAnnotations.Enabled = True
   chkAllowModifyContents.Enabled = True

   If optEncHigh.Value = True Then
     dmFraPDFHighPermissions.Enabled = True
     chkAllowDegradedPrinting.Enabled = True
     chkAllowFillIn.Enabled = True
     chkAllowScreenReaders.Enabled = True
     chkAllowAssembly.Enabled = True
    Else
     dmFraPDFHighPermissions.Enabled = False
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
        Case 0, 2:
   lsvFilenameSubst.ListItems.Add , , txtFilenameSubst(0).Text
   lsvFilenameSubst.ListItems(lsvFilenameSubst.ListItems.Count).SubItems(1) = txtFilenameSubst(1).Text
   lsvFilenameSubst.ListItems(lsvFilenameSubst.ListItems.Count).Selected = True
   lsvFilenameSubst.ListItems(lsvFilenameSubst.ListItems.Count).EnsureVisible
   Set_txtFilenameSubst
'  Case 2:
'   MsgBox LanguageStrings.MessagesMsg12 & _
    vbCrLf & vbTab & "\ / : * ? < > | """
  Case 3:
   MsgBox LanguageStrings.MessagesMsg11
 End Select
End Sub

Private Sub ChangeFilenameSubstitutions()
 Dim i As Long, res As Long
 res = CheckFilenameSubstitutions(lsvFilenameSubst.SelectedItem.Index)
 Select Case res
        Case 0, 2:
   lsvFilenameSubst.SelectedItem.Text = txtFilenameSubst(0).Text
   lsvFilenameSubst.SelectedItem.SubItems(1) = txtFilenameSubst(1).Text
'  Case 2:
'   MsgBox LanguageStrings.MessagesMsg12 & _
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
   cmdFilenameSubstMove(0).Enabled = True
   cmdFilenameSubstMove(1).Enabled = True
  Else
   cmdFilenameSubstMove(0).Enabled = False
   cmdFilenameSubstMove(1).Enabled = False
 End If
 If lsvFilenameSubst.ListItems.Count > 0 Then
  If lsvFilenameSubst.SelectedItem.Index = 1 Then
   cmdFilenameSubstMove(0).Enabled = False
  End If
  If lsvFilenameSubst.SelectedItem.Index = lsvFilenameSubst.ListItems.Count Then
   cmdFilenameSubstMove(1).Enabled = False
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
 With txtSavePreview
  .Text = GetSubstFilename("C:\test.pdf", txtSaveFilename.Text, , True) & ".pdf"
  .ToolTipText = .Text
 End With
End Sub

Private Function GetAutosaveFormatExtension() As String
 Select Case cmbAutosaveFormat.ListIndex
  Case -1, 0
   GetAutosaveFormatExtension = ".pdf"
  Case 1
   GetAutosaveFormatExtension = ".png"
  Case 2
   GetAutosaveFormatExtension = ".jpg"
  Case 3
   GetAutosaveFormatExtension = ".bmp"
  Case 4
   GetAutosaveFormatExtension = ".pcx"
  Case 5
   GetAutosaveFormatExtension = ".tif"
  Case 6
   GetAutosaveFormatExtension = ".ps"
  Case 7
   GetAutosaveFormatExtension = ".eps"
 End Select
End Function

Public Sub CorrectCmbCharset()
 Dim tStrf() As String
 If InStr(1, cmbCharset.Text, ",", vbTextCompare) > 0 Then
   tStrf = Split(cmbCharset.Text, ",")
   If Len(tStrf(0)) = 0 Then
     cmbCharset.Text = 0
    Else
     If IsNumeric(tStrf(0)) = False Then
       cmbCharset.Text = 0
      Else
       cmbCharset.Text = CLng(tStrf(0))
     End If
   End If
  Else
   If Len(cmbCharset.Text) = 0 Then
     cmbCharset.Text = 0
    Else
     If IsNumeric(cmbCharset.Text) = False Then
       cmbCharset.Text = 0
      Else
       cmbCharset.Text = CLng(cmbCharset.Text)
     End If
   End If
 End If
End Sub

Private Sub txtStandardAuthor_Change()
 txtStandardAuthor.ToolTipText = txtStandardAuthor.Text
End Sub

Private Sub txtTemppath_Change()
 txtTemppath.ToolTipText = txtTemppath.Text
 With txtTemppathPreview
  .Text = ResolveEnvironment(GetSubstFilename2(txtTemppath.Text))
  .ToolTipText = .Text
 End With
End Sub

Private Sub SetFrames()
 Dim ctl As Control
 For Each ctl In Controls
  If TypeOf ctl Is dmFrame Then
   ctl.Font.Size = 10
   ctl.TextShaddowColor = &HC00000
   If ComputerScreenResolution <= 8 Or Options.OptionsDesign = 1 Then
     ctl.UseGradient = False: ctl.Caption3D = [Flat Caption]
     If UCase$(ctl.Name) = "DMFRADESCRIPTION" Then
       ctl.BarColorFrom = vbRed
      Else
       ctl.BarColorFrom = vbBlue
     End If
    Else
     ctl.UseGradient = True: ctl.Caption3D = [Raised Caption]
     If UCase$(ctl.Name) = "DMFRADESCRIPTION" Then
       ctl.BarColorFrom = &H8080FF
       ctl.BarColorTo = &HC0&
      Else
       ctl.BarColorFrom = &HFF8080
       ctl.BarColorTo = &H400000
     End If
   End If
  End If
 Next ctl
End Sub
