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
         Left            =   1785
         TabIndex        =   188
         Top             =   735
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
            Picture         =   "frmOptions.frx":000C
            Style           =   1  'Grafisch
            TabIndex        =   199
            Top             =   1155
            Width           =   435
         End
         Begin VB.ComboBox cmbRunProgramAfterSavingWindowstyle 
            Height          =   315
            Left            =   210
            Style           =   2  'Dropdown-Liste
            TabIndex        =   194
            Top             =   2625
            Width           =   5370
         End
         Begin VB.CheckBox chkRunProgramAfterSavingWaitUntilReady 
            Appearance      =   0  '2D
            Caption         =   "Wait until ready"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   210
            TabIndex        =   193
            Top             =   3150
            Width           =   5805
         End
         Begin VB.TextBox txtRunProgramAfterSavingProgramParameters 
            Appearance      =   0  '2D
            Height          =   285
            Left            =   210
            TabIndex        =   192
            Top             =   1890
            Width           =   5805
         End
         Begin VB.CommandButton cmdRunProgramAfterSavingPrognameChoice 
            Caption         =   "..."
            Height          =   300
            Left            =   5040
            TabIndex        =   191
            Top             =   1155
            Width           =   435
         End
         Begin VB.ComboBox cmbRunProgramAfterSavingProgramname 
            Height          =   315
            Left            =   210
            TabIndex        =   190
            Top             =   1155
            Width           =   4770
         End
         Begin VB.CheckBox chkRunProgramAfterSaving 
            Appearance      =   0  '2D
            Caption         =   "Run a program/script after saving"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   210
            TabIndex        =   189
            Top             =   420
            Width           =   5805
         End
         Begin VB.Label lblRunProgramAfterSavingWindowstyle 
            AutoSize        =   -1  'True
            Caption         =   "Windowstyle"
            Height          =   195
            Left            =   210
            TabIndex        =   197
            Top             =   2415
            Width           =   900
         End
         Begin VB.Label lblRunProgramAfterSavingProgramParameters 
            AutoSize        =   -1  'True
            Caption         =   "Program parameters"
            Height          =   195
            Left            =   210
            TabIndex        =   196
            Top             =   1680
            Width           =   1410
         End
         Begin VB.Label lblRunProgramAfterSavingProgramname 
            AutoSize        =   -1  'True
            Caption         =   "Program/Script"
            Height          =   195
            Left            =   210
            TabIndex        =   195
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
            Picture         =   "frmOptions.frx":0596
            Style           =   1  'Grafisch
            TabIndex        =   200
            Top             =   1155
            Width           =   435
         End
         Begin VB.ComboBox cmbRunProgramBeforeSavingWindowstyle 
            Height          =   315
            Left            =   210
            Style           =   2  'Dropdown-Liste
            TabIndex        =   186
            Top             =   2625
            Width           =   2790
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
            TabIndex        =   187
            Top             =   2415
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
         TabIndex        =   198
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
   Begin PDFCreator.dmFrame dmFraProgSave 
      Height          =   2670
      Left            =   2730
      TabIndex        =   51
      Top             =   2835
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
         ItemData        =   "frmOptions.frx":0B20
         Left            =   120
         List            =   "frmOptions.frx":0B22
         Style           =   2  'Dropdown-Liste
         TabIndex        =   212
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
         ItemData        =   "frmOptions.frx":0B24
         Left            =   3720
         List            =   "frmOptions.frx":0B26
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
         TabIndex        =   211
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
         TabIndex        =   201
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
         ItemData        =   "frmOptions.frx":0B28
         Left            =   3690
         List            =   "frmOptions.frx":0B2A
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
         TabIndex        =   223
         Top             =   1890
         Width           =   6000
      End
      Begin VB.ComboBox cmbAuthorTokens 
         Appearance      =   0  '2D
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmOptions.frx":0B2C
         Left            =   3720
         List            =   "frmOptions.frx":0B2E
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
         TabIndex        =   222
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
         ItemData        =   "frmOptions.frx":0B30
         Left            =   120
         List            =   "frmOptions.frx":0B32
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
         ItemData        =   "frmOptions.frx":0B34
         Left            =   120
         List            =   "frmOptions.frx":0B36
         Style           =   2  'Dropdown-Liste
         TabIndex        =   92
         Top             =   735
         Width           =   2655
      End
      Begin VB.ComboBox cmbPDFRotate 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "frmOptions.frx":0B38
         Left            =   120
         List            =   "frmOptions.frx":0B3A
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
         TabIndex        =   221
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
         TabIndex        =   220
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
      TabIndex        =   214
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
         TabIndex        =   217
         Top             =   480
         Width           =   2580
      End
      Begin VB.CheckBox chkShowAnimation 
         Appearance      =   0  '2D
         Caption         =   "Show animation"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   120
         TabIndex        =   216
         Top             =   2220
         Width           =   5775
      End
      Begin VB.ComboBox cmbOptionsDesign 
         Height          =   315
         ItemData        =   "frmOptions.frx":0B3C
         Left            =   120
         List            =   "frmOptions.frx":0B3E
         Style           =   2  'Dropdown-Liste
         TabIndex        =   215
         Top             =   1620
         Width           =   3870
      End
      Begin PDFCreator.Line3D Line3D1 
         Height          =   30
         Index           =   3
         Left            =   105
         TabIndex        =   218
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
         TabIndex        =   219
         Top             =   1380
         Width           =   2250
      End
   End
   Begin PDFCreator.dmFrame dmFraProgPrint 
      Height          =   3930
      Left            =   2940
      TabIndex        =   202
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
         TabIndex        =   210
         Top             =   3360
         Width           =   4470
      End
      Begin VB.CheckBox chkPrintAfterSavingDuplex 
         Appearance      =   0  '2D
         Caption         =   "Duplex"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   209
         Top             =   3045
         Width           =   6015
      End
      Begin VB.CheckBox chkPrintAfterSavingNoCancel 
         Appearance      =   0  '2D
         Caption         =   "No cancel dialog"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   208
         Top             =   2625
         Width           =   6015
      End
      Begin VB.ComboBox cmbPrintAfterSavingQueryUser 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown-Liste
         TabIndex        =   207
         Top             =   1995
         Width           =   4770
      End
      Begin VB.ComboBox cmbPrintAfterSavingPrinter 
         Height          =   315
         Left            =   105
         TabIndex        =   205
         Top             =   1155
         Width           =   4770
      End
      Begin VB.CheckBox chkPrintAfterSaving 
         Appearance      =   0  '2D
         Caption         =   "Print after saving"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   203
         Top             =   420
         Width           =   6015
      End
      Begin VB.Label lblPrintAfterSavingQueryUser 
         AutoSize        =   -1  'True
         Caption         =   "Query user"
         Height          =   195
         Left            =   120
         TabIndex        =   206
         Top             =   1785
         Width           =   765
      End
      Begin VB.Label lblPrintAfterSavingPrinter 
         AutoSize        =   -1  'True
         Caption         =   "Printer"
         Height          =   195
         Left            =   120
         TabIndex        =   204
         Top             =   945
         Width           =   450
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
      TabIndex        =   213
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

Private Sub chkPrintAfterSaving_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkPrintAfterSaving.Value = 1 Then
50020    ViewPrintAfterSaving True
50030   Else
50040    ViewPrintAfterSaving False
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "chkPrintAfterSaving_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkPrintAfterSavingDuplex_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkPrintAfterSavingDuplex.Value = 1 Then
50020    ViewPrintAfterTumple True
50030   Else
50040    ViewPrintAfterTumple False
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "chkPrintAfterSavingDuplex_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkRunProgramBeforeSaving_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkRunProgramBeforeSaving.Value = 1 Then
50020    ViewRunProgramBeforeSaving True
50030   Else
50040    ViewRunProgramBeforeSaving False
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "chkRunProgramBeforeSaving_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkRunProgramAfterSaving_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkRunProgramAfterSaving.Value = 1 Then
50020    ViewRunProgramAfterSaving True
50030   Else
50040    ViewRunProgramAfterSaving False
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "chkRunProgramAfterSaving_Click")
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

Private Sub cmbAutosaveFormat_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txtAutosaveFilename.ToolTipText = txtAutosaveFilename.Text
50020  txtAutoSaveFilenamePreview.Text = GetSubstFilename("B:\dummy.dum", txtAutosaveFilename.Text, , True) & _
  GetSaveAutosaveFormatExtension(cmbAutosaveFormat.ListIndex)
50040  If IsValidPath("C:\" & txtAutoSaveFilenamePreview.Text) = False Then
50050    txtAutoSaveFilenamePreview.ForeColor = vbRed
50060   Else
50070    txtAutoSaveFilenamePreview.ForeColor = &H80000008
50080  End If
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
 If Err.Number = 380 Then
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
50110    txtGSresource.Text = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryResource"))
50120    Set reg = Nothing
50130    Exit Sub
50140   Else
50150    If InStr(UCase$(gsv), "AFPL") Then
50160     If InStr(gsv, " ") > 0 Then
50170      tsf = Split(gsv, " ")
50180      reg.KeyRoot = "SOFTWARE\AFPL Ghostscript\" & tsf(UBound(tsf))
50190      tStr = reg.GetRegistryValue("GS_DLL")
50200      SplitPath tStr, , Path
50210      txtGSbin.Text = CompletePath(Path)
50220      If InStrRev(Path, "\") > 0 Then
50230       txtGSlib.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50240       txtGSfonts.Text = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50250       If tsf(UBound(tsf)) <> "8.00" Then
50260        txtGSresource.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "Resource"
50270       End If
50280      End If
50290     End If
50300    End If
50310    If InStr(UCase$(gsv), "GNU") Then
50320     If InStr(gsv, " ") > 0 Then
50330      tsf = Split(gsv, " ")
50340      reg.KeyRoot = "SOFTWARE\GNU Ghostscript\" & tsf(UBound(tsf))
50350      tStr = reg.GetRegistryValue("GS_DLL")
50360      SplitPath tStr, , Path
50370      txtGSbin.Text = CompletePath(Path)
50380      If InStrRev(Path, "\") > 0 Then
50390       txtGSlib.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50400       txtGSfonts.Text = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50410       txtGSresource.Text = ""
50420      End If
50430     End If
50440    End If
50450    If InStr(UCase$(gsv), "GPL") Then
50460     If InStr(gsv, " ") > 0 Then
50470      tsf = Split(gsv, " ")
50480      reg.KeyRoot = "SOFTWARE\GPL Ghostscript\" & tsf(UBound(tsf))
50490      tStr = reg.GetRegistryValue("GS_DLL")
50500      SplitPath tStr, , Path
50510      txtGSbin.Text = CompletePath(Path)
50520      If InStrRev(Path, "\") > 0 Then
50530       txtGSlib.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50540       txtGSfonts.Text = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50550       txtGSresource.Text = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "Resource"
50560      End If
50570     End If
50580    End If
50590  End If
50600  Set reg = Nothing
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

Private Sub cmbOptionsDesign_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Options.OptionsDesign = cmbOptionsDesign.ListIndex
50020  SetFrames
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmbOptionsDesign_Click")
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

Private Sub cmbRunProgramAfterSavingProgramname_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Program As String, drv As String
50020  Program = RemoveLeadingAndTrailingQuotes(cmbRunProgramAfterSavingProgramname.Text)
50030  SplitPath Program, drv
50040  If LenB(drv) = 0 Then
50050   Program = ResolveRelativePath(Program, GetPDFCreatorApplicationPath)
50060  End If
50070  If FileExists(Program) = True Then
50080    If IsFileEditable(Program) Then
50090      cmdRunProgramAfterSavingPrognameEdit.Enabled = True
50100     Else
50110      cmdRunProgramAfterSavingPrognameEdit.Enabled = False
50120    End If
50130   Else
50140    cmdRunProgramAfterSavingPrognameEdit.Enabled = False
50150  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmbRunProgramAfterSavingProgramname_Change")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbRunProgramAfterSavingProgramname_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With cmbRunProgramAfterSavingProgramname
50020   If .ListCount > 0 Then
50030    .Text = "Scripts\RunProgramAfterSaving\" & .List(.ListIndex)
50040   End If
50050  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmbRunProgramAfterSavingProgramname_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbRunProgramBeforeSavingProgramname_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Program As String, drv As String
50020  Program = RemoveLeadingAndTrailingQuotes(cmbRunProgramBeforeSavingProgramname.Text)
50030  SplitPath Program, drv
50040  If LenB(drv) = 0 Then
50050   Program = ResolveRelativePath(Program, GetPDFCreatorApplicationPath)
50060  End If
50070  If FileExists(Program) = True Then
50080    If IsFileEditable(Program) Then
50090      cmdRunProgramBeforeSavingPrognameEdit.Enabled = True
50100     Else
50110      cmdRunProgramBeforeSavingPrognameEdit.Enabled = False
50120    End If
50130   Else
50140    cmdRunProgramBeforeSavingPrognameEdit.Enabled = False
50150  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmbRunProgramBeforeSavingProgramname_Change")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmbRunProgramBeforeSavingProgramname_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With cmbRunProgramBeforeSavingProgramname
50020   If .ListCount > 0 Then
50030    .Text = "Scripts\RunProgramBeforeSaving\" & .List(.ListIndex)
50040   End If
50050  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmbRunProgramBeforeSavingProgramname_Click")
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
50020  txtSavePreview.ToolTipText = txtSavePreview.Text
50030  txtSavePreview.Text = GetSubstFilename("B:\dummy.dum", txtSaveFilename.Text, , True) & _
  GetSaveAutosaveFormatExtension(cmbStandardSaveFormat.ListIndex)
50050  If IsValidPath("C:\" & txtSavePreview.Text) = False Then
50060    txtSavePreview.ForeColor = vbRed
50070   Else
50080    txtSavePreview.ForeColor = &H80000008
50090  End If
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

Private Sub cmbStandardSaveFormat_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txtSavePreview.ToolTipText = txtSavePreview.Text
50020  txtSavePreview.Text = GetSubstFilename("B:\dummy.dum", txtSaveFilename.Text, , True) & _
  GetSaveAutosaveFormatExtension(cmbStandardSaveFormat.ListIndex)
50040  If IsValidPath("C:\" & txtSavePreview.Text) = False Then
50050    txtSavePreview.ForeColor = vbRed
50060   Else
50070    txtSavePreview.ForeColor = &H80000008
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmbStandardSaveFormat_Click")
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
50050   ieb.Refresh
50060  End With
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
50080  End Select
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

Private Sub cmdFilenameSubstMove_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case Index
        Case 0: ' Up
50030    MoveUpFilenameSubstitutions
50040   Case 1: ' Down
50050    MoveDownFilenameSubstitutions
50060  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdFilenameSubstMove_Click")
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
50010  Dim strFolder As String
50020  strFolder = BrowseForFolderFiles(Me.hwnd, LanguageStrings.OptionsAutosaveDirectoryPrompt)
50030  If Len(strFolder) = 0 Then
50040   Exit Sub
50050  End If
50060  txtAutosaveDirectory.Text = CompletePath(strFolder)
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
50020  strFolder = BrowseForFolderFiles(Me.hwnd, LanguageStrings.OptionsGhostscriptBinariesDirectoryPrompt)
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
50020  strFolder = BrowseForFolderFiles(Me.hwnd, LanguageStrings.OptionsGhostscriptFontsDirectoryPrompt)
50030  If Len(strFolder) = 0 Then Exit Sub
50040  strFolder = CompletePath(strFolder)
50050  If LenB(Dir(strFolder & "*.afm", vbNormal)) = 0 And LenB(Dir(strFolder & "*.pfb", vbNormal)) = 0 Then
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
50020  strFolder = BrowseForFolderFiles(Me.hwnd, LanguageStrings.OptionsGhostscriptLibrariesDirectoryPrompt)
50030  If Len(strFolder) = 0 Then Exit Sub
50040  strFolder = CompletePath(strFolder)
50050  If LenB(Dir(strFolder & "*.*", vbNormal)) = 0 Then
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
50020  strFolder = BrowseForFolderFiles(Me.hwnd, LanguageStrings.OptionsGhostscriptResourceDirectoryPrompt)
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
50020  strFolder = BrowseForFolderFiles(Me.hwnd, LanguageStrings.OptionsPrintertempDirectoryPrompt)
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
50100    ieb.Refresh
50110   End With
50120  End If
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

Private Sub cmdRunProgramAfterSavingPrognameChoice_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Filename As String
50020  Filename = BrowseForFolderFiles(Me.hwnd, LanguageStrings.OptionsProgramRunProgramAfterSavingCaption, False)
50030  If LenB(Filename) > 0 Then
50040   cmbRunProgramAfterSavingProgramname.Text = Filename
50050  End If
50060  If FileExists(Filename) = True Then
50070    If IsFileEditable(Filename) Then
50080      cmdRunProgramAfterSavingPrognameEdit.Enabled = True
50090     Else
50100      cmdRunProgramAfterSavingPrognameEdit.Enabled = False
50110    End If
50120   Else
50130    cmdRunProgramAfterSavingPrognameEdit.Enabled = False
50140  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdRunProgramAfterSavingPrognameChoice_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdRunProgramAfterSavingPrognameEdit_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Program As String, drv As String
50020  Program = RemoveLeadingAndTrailingQuotes(cmbRunProgramAfterSavingProgramname.Text)
50030  SplitPath Program, drv
50040  If LenB(drv) = 0 Then
50050   Program = ResolveRelativePath(Program, GetPDFCreatorApplicationPath)
50060  End If
50070  If FileExists(Program) = True Then
50080   If IsFileEditable(Program) Then
50090    EditDocument Program
50100   End If
50110  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdRunProgramAfterSavingPrognameEdit_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdRunProgramBeforeSavingPrognameChoice_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Filename As String
50020  Filename = BrowseForFolderFiles(Me.hwnd, LanguageStrings.OptionsProgramRunProgramBeforeSavingCaption, False)
50030  If LenB(Filename) > 0 Then
50040   cmbRunProgramBeforeSavingProgramname.Text = Filename
50050  End If
50060  If FileExists(Filename) = True Then
50070    If IsFileEditable(Filename) Then
50080      cmdRunProgramBeforeSavingPrognameEdit.Enabled = True
50090     Else
50100      cmdRunProgramBeforeSavingPrognameEdit.Enabled = False
50110    End If
50120   Else
50130    cmdRunProgramBeforeSavingPrognameEdit.Enabled = False
50140  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdRunProgramBeforeSavingPrognameChoice_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdRunProgramBeforeSavingPrognameEdit_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Program As String, drv As String
50020  Program = RemoveLeadingAndTrailingQuotes(cmbRunProgramBeforeSavingProgramname.Text)
50030  SplitPath Program, drv
50040  If LenB(drv) = 0 Then
50050   Program = ResolveRelativePath(Program, GetPDFCreatorApplicationPath)
50060  End If
50070  If FileExists(Program) = True Then
50080   If IsFileEditable(Program) Then
50090    EditDocument Program
50100   End If
50110  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdRunProgramBeforeSavingPrognameEdit_Click")
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
50080  CorrectOptions
50090  SaveOptions Options
50100  If IsWin9xMe = False Then
50111   Select Case Options.ProcessPriority
               Case 0: 'Idle
50130     SetProcessPriority Idle
50140    Case 1: 'Normal
50150     SetProcessPriority Normal
50160    Case 2: 'High
50170     SetProcessPriority High
50180    Case 3: 'Realtime
50190     SetProcessPriority RealTime
50200   End Select
50210  End If
50220  If tRestart = True Then
50230   Restart = True
50240  End If
50250  Unload Me
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
50170  With cmdTest.Font
50180   tFontname = .Name
50190   tFontSize = .Size
50200   tFontCharset = .Charset
50210  End With
50220  SetFont Me, cmbFonts.List(cmbFonts.ListIndex), CLng(tStr), cmbProgramFontsize.Text
50230  cmbCharset.Text = tCharset
50240  SetFont frmMain, cmbFonts.List(cmbFonts.ListIndex), CLng(tStr), cmbProgramFontsize.Text
50250  ieb.Refresh
50260  With cmdTest.Font
50270   .Name = tFontname
50280   .Size = tFontSize
50290   .Charset = tFontCharset
50300  End With
50310  With cmdCancelTest
50320   .Font.Name = tFontname
50330   .Font.Size = tFontSize
50340   .Font.Charset = tFontCharset
50350   .Enabled = True
50360  End With
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
50010  PrintTestpage frmMain
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
50020  Temppath = "<Temp>PDFCreator\"
50030  If DirExists(ResolveEnvironment(GetSubstFilename2(Temppath))) = False Then
50040   MakePath ResolveEnvironment(GetSubstFilename2(Temppath))
50050  End If
50060  With txtTemppath
50070   .Text = Temppath
50080   .ToolTipText = ResolveEnvironment(GetSubstFilename2(Temppath))
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
50030     Call HTMLHelp_ShowTopic("html\generalsettings.htm")
50040 '  MsgBox ieb.GetSelectedGroup & vbCrLf & ieb.GetSelectedItem
50051    Select Case ieb.GetSelectedGroup
          Case 1
50071      Select Case ieb.GetSelectedItem
            Case 1
50090        Call HTMLHelp_ShowTopic("html\generalsettings.htm")
50100       Case 2
50110        Call HTMLHelp_ShowTopic("html\ghostscript.htm")
50120       Case 3
50130        Call HTMLHelp_ShowTopic("html\docproperties.htm")
50140       Case 4
50150        Call HTMLHelp_ShowTopic("html\savesettings.htm")
50160       Case 5
50170        Call HTMLHelp_ShowTopic("html\autosave.htm")
50180       Case 6
50190        Call HTMLHelp_ShowTopic("html\directories.htm")
50200       Case 7
50210        Call HTMLHelp_ShowTopic("html\fontsetting.htm")
50220       Case Else
50230        Call HTMLHelp_ShowTopic("html\generalsettings.htm")
50240      End Select
50250     Case 2
50261      Select Case ieb.GetSelectedItem
            Case 1
50281        Select Case tbstrPDFOptions.SelectedItem.Index
              Case 1
50300          Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
50310         Case 2
50320          Call HTMLHelp_ShowTopic("html\pdfcompression.htm")
50330         Case 3
50340          Call HTMLHelp_ShowTopic("html\pdffonts.htm")
50350         Case 4
50360          Call HTMLHelp_ShowTopic("html\pdfcolors.htm")
50370         Case 5
50380          Call HTMLHelp_ShowTopic("html\pdfsecurity.htm")
50390         Case Else
50400          Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
50410        End Select
50420       Case 2
50430        Call HTMLHelp_ShowTopic("html\pngsettings.htm")
50440       Case 3
50450        Call HTMLHelp_ShowTopic("html\jpegsettings.htm")
50460       Case 4
50470        Call HTMLHelp_ShowTopic("html\bmpsettings.htm")
50480       Case 5
50490        Call HTMLHelp_ShowTopic("html\pcxsettings.htm")
50500       Case 6
50510        Call HTMLHelp_ShowTopic("html\tiffsettings.htm")
50520       Case 7
50530        Call HTMLHelp_ShowTopic("html\pssettings.htm")
50540       Case 8
50550        Call HTMLHelp_ShowTopic("html\epssettings.htm")
50560       Case Else
50570        Call HTMLHelp_ShowTopic("html\pdfgeneral.htm")
50580      End Select
50590    End Select
50600  End If
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
  fc As Long, reg As clsRegistry, tsf() As String, tStr2 As String, Files As Collection, _
  Path As String, Filename As String, Ext As String, p As Printer
50050
50060  Me.Icon = LoadResPicture(2120, vbResIcon)
50070  KeyPreview = True
50080
50090  With Screen
50100   .MousePointer = vbHourglass
50110   Move (.Width - Width) / 2, (.Height - Height) / 2
50120  End With
50130
50140  SetFrames
50150
50160  With dmFraDescription
50170   .Caption = LanguageStrings.OptionsTreeProgram
50180   .Visible = True
50190  End With
50200  tbstrProgGeneral.Visible = True
50210  With dmFraProgGeneral1
50220   .Visible = True
50230   .Top = dmFraDescription.Top + dmFraDescription.Height + 50
50240   .Left = dmFraDescription.Left
50250   dmFraProgGeneral2.Top = .Top
50260   dmFraProgGeneral2.Left = .Left
50270   dmFraProgGeneral2.Width = .Width
50280   dmFraShellIntegration.Width = dmFraProgGeneral2.Width
50290   dmFraProgGhostscript.Top = .Top
50300   dmFraProgGhostscript.Left = .Left
50310   dmFraProgGhostscript.Width = dmFraDescription.Width
50320   dmFraProgAutosave.Top = .Top
50330   dmFraProgAutosave.Left = .Left
50340   dmFraProgAutosave.Width = dmFraDescription.Width
50350   dmFraProgDirectories.Top = .Top
50360   dmFraProgDirectories.Left = .Left
50370   dmFraProgDirectories.Width = dmFraDescription.Width
50380   dmFraProgDocument.Top = .Top
50390   dmFraProgDocument.Left = .Left
50400   dmFraProgDocument.Width = dmFraDescription.Width
50410   dmFraProgSave.Top = .Top
50420   dmFraProgSave.Left = .Left
50430   dmFraProgSave.Width = dmFraDescription.Width
50440   dmFraFilenameSubstitutions.Top = dmFraProgSave.Top + dmFraProgSave.Height + 50
50450   dmFraFilenameSubstitutions.Left = .Left
50460   dmFraFilenameSubstitutions.Width = dmFraDescription.Width
50470   dmFraProgFont.Top = .Top
50480   dmFraProgFont.Left = .Left
50490   dmFraProgFont.Width = dmFraDescription.Width
50500   dmFraProgActions.Top = .Top
50510   dmFraProgActions.Left = .Left
50520   dmFraProgActions.Width = dmFraDescription.Width
50530   dmFraProgPrint.Top = .Top
50540   dmFraProgPrint.Left = .Left
50550   dmFraProgPrint.Width = dmFraDescription.Width
50560   dmFraBitmapGeneral.Top = .Top
50570   dmFraBitmapGeneral.Left = .Left
50580   dmFraBitmapGeneral.Width = dmFraDescription.Width
50590   dmFraPSGeneral.Top = .Top
50600   dmFraPSGeneral.Left = .Left
50610   dmFraPSGeneral.Width = dmFraDescription.Width
50620
50630   dmFraProgActionsRunProgramAfterSaving.Top = dmFraProgActionsRunProgramBeforeSaving.Top
50640   dmFraProgActionsRunProgramAfterSaving.Left = dmFraProgActionsRunProgramBeforeSaving.Left
50650
50660   cmdCancel.Left = .Left
50670   cmdReset.Left = .Left + (dmFraDescription.Width - cmdReset.Width) / 2
50680   cmdSave.Left = .Left + dmFraDescription.Width - cmdSave.Width
50690  End With
50700
50710  With tbstrProgGeneral
50720   .Top = dmFraDescription.Top + dmFraDescription.Height + 50
50730   .Left = dmFraDescription.Left
50740   .Height = cmdCancel.Top - tbstrProgGeneral.Top - 50
50750   .Width = dmFraDescription.Width
50760  End With
50770
50780  With dmFraProgGeneral1
50790   .Top = tbstrProgGeneral.ClientTop + 100
50800   .Left = tbstrProgGeneral.Left + (tbstrProgGeneral.Width - .Width) / 2
50810  End With
50820  With dmFraProgGeneral2
50830   .Top = tbstrProgGeneral.ClientTop + 100
50840   .Left = tbstrProgGeneral.Left + (tbstrProgGeneral.Width - .Width) / 2
50850  End With
50860  With dmFraShellIntegration
50870   .Top = dmFraProgGeneral2.Top + dmFraProgGeneral2.Height + 50
50880   .Left = tbstrProgGeneral.Left + (tbstrProgGeneral.Width - .Width) / 2
50890  End With
50900
50910  With tbstrPDFOptions
50920   .Top = dmFraDescription.Top + dmFraDescription.Height + 50
50930   .Left = dmFraDescription.Left
50940   .Height = cmdCancel.Top - tbstrPDFOptions.Top - 50
50950   .Width = dmFraDescription.Width
50960  End With
50970
50980  With dmFraPDFGeneral
50990   .Top = tbstrPDFOptions.ClientTop + 100
51000   .Left = tbstrPDFOptions.Left + (tbstrPDFOptions.Width - .Width) / 2
51010   dmfraPDFCompress.Top = .Top
51020   dmfraPDFCompress.Left = .Left
51030   dmFraPDFFonts.Top = .Top
51040   dmFraPDFFonts.Left = .Left
51050   dmFraPDFColors.Top = .Top
51060   dmFraPDFColors.Left = .Left
51070   dmFraPDFColorOptions.Top = dmFraPDFColors.Top + dmFraPDFColors.Height + 50
51080   dmFraPDFColorOptions.Left = .Left
51090   dmFraPDFSecurity.Top = .Top
51100   dmFraPDFSecurity.Left = .Left
51110  End With
51120
51130  cmbEPSLanguageLevel.Top = cmbPSLanguageLevel.Top
51140  cmbEPSLanguageLevel.Left = cmbPSLanguageLevel.Left
51150
51160  ieb.DisableUpdates True
51170  ieb.ClearStructure
51180  ieb.SetImageList imlIeb
51190  With LanguageStrings
51200   ieb.AddGroup "Program", .OptionsTreeProgram, 0
51210   ieb.AddItem "Program", "General", .OptionsProgramGeneralSymbol, 1
51220   ieb.AddItem "Program", "Ghostscript", .OptionsProgramGhostscriptSymbol, 2
51230   ieb.AddItem "Program", "Document", .OptionsProgramDocumentSymbol, 3
51240   ieb.AddItem "Program", "Save", .OptionsProgramSaveSymbol, 4
51250   ieb.AddItem "Program", "AutoSave", .OptionsProgramAutosaveSymbol, 5
51260   ieb.AddItem "Program", "Directories", .OptionsProgramDirectoriesSymbol, 6
51270   ieb.AddItem "Program", "Actions", .OptionsProgramActionsSymbol, 7
51280   ieb.AddItem "Program", "Print", .OptionsProgramPrintSymbol, 8
51290   ieb.AddItem "Program", "Fonts", .OptionsProgramFontSymbol, 9
51300   ieb.AddGroup "Formats", .OptionsTreeFormats, 0
51310   ieb.AddItem "Formats", "PDF", .OptionsPDFSymbol, 10
51320   ieb.AddItem "Formats", "PNG", .OptionsPNGSymbol, 11
51330   ieb.AddItem "Formats", "JPEG", .OptionsJPEGSymbol, 12
51340   ieb.AddItem "Formats", "BMP", .OptionsBMPSymbol, 13
51350   ieb.AddItem "Formats", "PCX", .OptionsPCXSymbol, 14
51360   ieb.AddItem "Formats", "TIFF", .OptionsTIFFSymbol, 15
51370   ieb.AddItem "Formats", "PS", .OptionsPSSymbol, 16
51380   ieb.AddItem "Formats", "EPS", .OptionsEPSSymbol, 17
51390   ieb.DisableUpdates False
51400
51410   Set picOptions = LoadResPicture(2101, vbResIcon)
51420   dmFraProgGeneral1.Visible = True
51430
51440   dmFraProgGeneral1.Caption = .OptionsProgramGeneralDescription1
51450   dmFraProgGeneral2.Caption = .OptionsProgramGeneralDescription2
51460   With tbstrProgGeneral.Tabs
51470    .Clear
51480    .Add , , LanguageStrings.OptionsProgramGeneralDescription1
51490    .Add , , LanguageStrings.OptionsProgramGeneralDescription2
51500   End With
51510   dmFraShellIntegration.Caption = .OptionsShellIntegration
51520   dmFraProgGhostscript.Caption = .OptionsProgramGhostscriptSymbol
51530   dmFraProgAutosave.Caption = .OptionsProgramAutosaveSymbol
51540   dmFraProgDirectories.Caption = .OptionsProgramDirectoriesSymbol
51550   dmFraProgDocument.Caption = .OptionsProgramDocumentSymbol
51560   dmFraProgFont.Caption = .OptionsProgramFontSymbol
51570   dmFraProgSave.Caption = .OptionsProgramSaveSymbol
51580   dmFraProgActions.Caption = .OptionsProgramActionsSymbol
51590   dmFraProgPrint.Caption = .OptionsProgramPrintSymbol
51600
51610   cmdShellintegration(0).Caption = .OptionsShellIntegrationAdd
51620   cmdShellintegration(1).Caption = .OptionsShellIntegrationRemove
51630   If IsWin9xMe = False Then
51640    If IsAdmin = False Then
51650     cmdShellintegration(0).Enabled = False
51660     cmdShellintegration(1).Enabled = False
51670    End If
51680   End If
51690
51700   lblSendMailMethod.Caption = .OptionsSendMailMethod
51710   cmbSendMailMethod.AddItem .OptionsSendMailMethodAutomatic
51720   cmbSendMailMethod.AddItem .OptionsSendMailMethodMapi
51730   cmbSendMailMethod.AddItem .OptionsSendMailMethodSendmailDLL
51740
51750   lblGhostscriptversion.Caption = .OptionsGhostscriptversion
51760   lblAdditionalGhostscriptParameters.Caption = .OptionsAdditionalGhostscriptParameters
51770   lblAdditionalGhostscriptSearchpath.Caption = .OptionsAdditionalGhostscriptSearchpath
51780   chkAddWindowsFontpath.Caption = .OptionsAddWindowsFontpath
51790
51800   lblSaveFilename.Caption = .OptionsSaveFilename
51810   lblSaveFilenameTokens.Caption = .OptionsSaveFilenameTokens
51820   dmFraFilenameSubstitutions.Caption = .OptionsSaveFilenameSubstitutions
51830   chkFilenameSubst.Caption = .OptionsSaveFilenameSubstitutionsTitle
51840   cmdFilenameSubst(0).Caption = .OptionsSaveFilenameAdd
51850   cmdFilenameSubst(1).Caption = .OptionsSaveFilenameChange
51860   cmdFilenameSubst(2).Caption = .OptionsSaveFilenameDelete
51870
51880   chkSpaces.Caption = .OptionsRemoveSpaces
51890   chkNoConfirmMessageSwitchingDefaultprinter.Caption = .OptionsProgramSwitchingDefaultprinter
51900   chkNoProcessingAtStartup.Caption = .OptionsProgramNoProcessingAtStartup
51910   lblOptionsDesign.Caption = .OptionsProgramOptionsDesign
51920   cmbOptionsDesign.AddItem .OptionsProgramOptionsDesignGradient
51930   cmbOptionsDesign.AddItem .OptionsProgramOptionsDesignSimple
51940   chkShowAnimation.Caption = .OptionsProgramShowAnimation
51950
51960   lblGSbin.Caption = .OptionsDirectoriesGSBin
51970   lblGSlib.Caption = .OptionsDirectoriesGSLibraries
51980   lblGSfonts.Caption = .OptionsDirectoriesGSFonts
51990   lblPrintTempPath.Caption = .OptionsDirectoriesTempPath
52000
52010   chkOnePagePerFile.Caption = .OptionsOnePagePerFile
52020   lblOptions = .OptionsProgramGeneralDescription
52030   lblAutosaveformat.Caption = .OptionsAutosaveFormat
52040   chkUseStandardAuthor.Caption = .OptionsUseStandardauthor
52050   chkUseCreationDateNow.Caption = .OptionsUseCreationDateNow
52060   chkUseAutosave.Caption = .OptionsUseAutosave
52070   cmdTestpage.Caption = .OptionsPrintTestpage
52080   lblAutosaveFilename.Caption = .OptionsAutosaveFilename
52090   lblAutosaveFilenameTokens.Caption = .OptionsAutosaveFilenameTokens
52100   chkUseAutosaveDirectory.Caption = .OptionsUseAutosaveDirectory
52110   lblAuthorTokens.Caption = .OptionsStandardAuthorToken
52120   chkAutosaveStartStandardProgram.Caption = .OptionsAutosaveStartStandardProgram
52130   chkAutosaveSendEmail.Caption = .OptionsSendEmailAfterAutosave
52140   lblStandardSaveformat.Caption = .OptionsStandardSaveFormat
52150
52160   dmFraProgActionsRunProgramAfterSaving.Caption = .OptionsProgramRunProgramAfterSavingCaption
52170   chkRunProgramAfterSaving.Caption = .OptionsProgramRunProgramAfterSavingCaption
52180   lblRunProgramAfterSavingProgramname.Caption = .OptionsProgramRunProgramAfterSavingProgram
52190   lblRunProgramAfterSavingProgramParameters.Caption = .OptionsProgramRunProgramAfterSavingProgramParameters
52200   chkRunProgramAfterSavingWaitUntilReady.Caption = .OptionsProgramRunProgramAfterSavingWaitUntilReady
52210   lblRunProgramAfterSavingWindowstyle.Caption = .OptionsProgramRunProgramAfterSavingWindowstyle
52220   cmbRunProgramAfterSavingWindowstyle.AddItem .OptionsProgramRunProgramAfterSavingWindowstyleHide
52230   cmbRunProgramAfterSavingWindowstyle.AddItem .OptionsProgramRunProgramAfterSavingWindowstyleNormalFocus
52240   cmbRunProgramAfterSavingWindowstyle.AddItem .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedFocus
52250   cmbRunProgramAfterSavingWindowstyle.AddItem .OptionsProgramRunProgramAfterSavingWindowstyleMaximizedFocus
52260   cmbRunProgramAfterSavingWindowstyle.AddItem .OptionsProgramRunProgramAfterSavingWindowstyleNormalNoFocus
52270   cmbRunProgramAfterSavingWindowstyle.AddItem .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedNoFocus
52280
52290   With tbstrProgActions.Tabs
52300    .Clear
52310    .Add , , LanguageStrings.OptionsProgramRunProgramBeforeSavingCaption
52320    .Add , , LanguageStrings.OptionsProgramRunProgramAfterSavingCaption
52330   End With
52340
52350   dmFraProgActionsRunProgramBeforeSaving.Caption = .OptionsProgramRunProgramBeforeSavingCaption
52360   chkRunProgramBeforeSaving.Caption = .OptionsProgramRunProgramBeforeSavingCaption
52370   lblRunProgramBeforeSavingProgramname.Caption = .OptionsProgramRunProgramBeforeSavingProgram
52380   lblRunProgramBeforeSavingProgramParameters.Caption = .OptionsProgramRunProgramBeforeSavingProgramParameters
52390   lblRunProgramBeforeSavingWindowstyle.Caption = .OptionsProgramRunProgramBeforeSavingWindowstyle
52400   cmbRunProgramBeforeSavingWindowstyle.AddItem .OptionsProgramRunProgramBeforeSavingWindowstyleHide
52410   cmbRunProgramBeforeSavingWindowstyle.AddItem .OptionsProgramRunProgramBeforeSavingWindowstyleNormalFocus
52420   cmbRunProgramBeforeSavingWindowstyle.AddItem .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedFocus
52430   cmbRunProgramBeforeSavingWindowstyle.AddItem .OptionsProgramRunProgramBeforeSavingWindowstyleMaximizedFocus
52440   cmbRunProgramBeforeSavingWindowstyle.AddItem .OptionsProgramRunProgramBeforeSavingWindowstyleNormalNoFocus
52450   cmbRunProgramBeforeSavingWindowstyle.AddItem .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedNoFocus
52460
52470   chkPrintAfterSaving.Caption = .OptionsPrintAfterSaving
52480   lblPrintAfterSavingPrinter.Caption = .OptionsPrintAfterSavingPrinter
52490
52500   For Each p In Printers
52510    cmbPrintAfterSavingPrinter.AddItem p.DeviceName
52520   Next p
52530
52540   lblPrintAfterSavingQueryUser.Caption = .OptionsPrintAfterSavingQueryUser
52550   cmbPrintAfterSavingQueryUser.AddItem .OptionsPrintAfterSavingQueryUserOff
52560   cmbPrintAfterSavingQueryUser.AddItem .OptionsPrintAfterSavingQueryUserStandardPrinterDialog
52570   cmbPrintAfterSavingQueryUser.AddItem .OptionsPrintAfterSavingQueryUserPrinterSetupDialog
52580   cmbPrintAfterSavingQueryUser.AddItem .OptionsPrintAfterSavingQueryUserDefaultPrinter
52590
52600   chkPrintAfterSavingNoCancel.Caption = .OptionsPrintAfterSavingNoCancel
52610   chkPrintAfterSavingDuplex.Caption = .OptionsPrintAfterSavingDuplex
52620   cmbPrintAfterSavingTumble.AddItem .OptionsPrintAfterSavingDuplexTumbleOff
52630   cmbPrintAfterSavingTumble.AddItem .OptionsPrintAfterSavingDuplexTumbleOn
52640
52650   With cmbStandardSaveFormat
52660    .AddItem "PDF"
52670    .AddItem "PNG"
52680    .AddItem "JPEG"
52690    .AddItem "BMP"
52700    .AddItem "PCX"
52710    .AddItem "TIFF"
52720    .AddItem "PS"
52730    .AddItem "EPS"
52740   End With
52750   With cmbAutosaveFormat
52760    .AddItem "PDF"
52770    .AddItem "PNG"
52780    .AddItem "JPEG"
52790    .AddItem "BMP"
52800    .AddItem "PCX"
52810    .AddItem "TIFF"
52820    .AddItem "PS"
52830    .AddItem "EPS"
52840   End With
52850   With cmbSaveFilenameTokens
52860    .AddItem "<Author>"
52870    .AddItem "<Computername>"
52880    .AddItem "<DateTime>"
52890    .AddItem "<Title>"
52900    .AddItem "<Username>"
52910    .AddItem "<REDMON_DOCNAME>"
52920    .AddItem "<REDMON_DOCNAME_FILE>"
52930    .AddItem "<REDMON_DOCNAME_PATH>"
52940    .AddItem "<REDMON_JOB>"
52950    .AddItem "<REDMON_MACHINE>"
52960    .AddItem "<REDMON_PORT>"
52970    .AddItem "<REDMON_PRINTER>"
52980    .AddItem "<REDMON_SESSIONID>"
52990    .AddItem "<REDMON_USER>"
53000    .ListIndex = 0
53010   End With
53020   With cmbAuthorTokens
53030    .AddItem "<Computername>"
53040    .AddItem "<ClientComputer>"
53050    .AddItem "<DateTime>"
53060    .AddItem "<Title>"
53070    .AddItem "<Username>"
53080    .AddItem "<REDMON_DOCNAME>"
53090    .AddItem "<REDMON_DOCNAME_FILE>"
53100    .AddItem "<REDMON_DOCNAME_PATH>"
53110    .AddItem "<REDMON_JOB>"
53120    .AddItem "<REDMON_MACHINE>"
53130    .AddItem "<REDMON_PORT>"
53140    .AddItem "<REDMON_PRINTER>"
53150    .AddItem "<REDMON_SESSIONID>"
53160    .AddItem "<REDMON_USER>"
53170    .ListIndex = 0
53180   End With
53190   With cmbAutoSaveFilenameTokens
53200    .AddItem "<Author>"
53210    .AddItem "<Computername>"
53220    .AddItem "<ClientComputer>"
53230    .AddItem "<DateTime>"
53240    .AddItem "<Title>"
53250    .AddItem "<Username>"
53260    .AddItem "<REDMON_DOCNAME>"
53270    .AddItem "<REDMON_DOCNAME_FILE>"
53280    .AddItem "<REDMON_DOCNAME_PATH>"
53290    .AddItem "<REDMON_JOB>"
53300    .AddItem "<REDMON_MACHINE>"
53310    .AddItem "<REDMON_PORT>"
53320    .AddItem "<REDMON_PRINTER>"
53330    .AddItem "<REDMON_SESSIONID>"
53340    .AddItem "<REDMON_USER>"
53350    .ListIndex = 0
53360   End With
53370   Me.Caption = .DialogPrinterOptions
53380   cmdCancel.Caption = .OptionsCancel
53390   cmdReset.Caption = .OptionsReset
53400   cmdSave.Caption = .OptionsSave
53410   tbstrPDFOptions.Tabs.Clear
53420   tbstrPDFOptions.Tabs.Add , "General", .OptionsPDFGeneral
53430   tbstrPDFOptions.Tabs.Add , "Compression", .OptionsPDFCompression
53440   tbstrPDFOptions.Tabs.Add , "Fonts", .OptionsPDFFonts
53450   tbstrPDFOptions.Tabs.Add , "Colors", .OptionsPDFColors
53460   tbstrPDFOptions.Tabs.Add , "Security", .OptionsPDFSecurity
53470   dmFraPDFGeneral.Caption = .OptionsPDFGeneralCaption
53480   chkPDFOptimize.Caption = .OptionsPDFOptimize
53490   lblPDFCompat.Caption = .OptionsPDFGeneralCompatibility
53500   lblPDFAutoRotate.Caption = .OptionsPDFGeneralAutorotate
53510   lblPDFResolution.Caption = .OptionsPDFGeneralResolution
53520   lblPDFOverprint.Caption = .OptionsPDFGeneralOverprint
53530   lblProgfont.Caption = .OptionsProgramFont
53540   lblProgcharset.Caption = .OptionsProgramFontcharset
53550   lblSize.Caption = .OptionsProgramFontSize
53560   lblTesttext = .OptionsProgramFontTestdescription
53570   cmdTest.Caption = .OptionsProgramFontTest
53580   cmdCancelTest.Caption = .OptionsProgramFontCancelTest
53590   chkPDFASCII85.Caption = .OptionsPDFGeneralASCII85
53600   cmbPDFCompat.Clear
53610   cmbPDFCompat.AddItem .OptionsPDFGeneralCompatibility01
53620   cmbPDFCompat.AddItem .OptionsPDFGeneralCompatibility02
53630   cmbPDFCompat.AddItem .OptionsPDFGeneralCompatibility03
53640   cmbPDFRotate.Clear
53650   cmbPDFRotate.AddItem .OptionsPDFGeneralRotate01
53660   cmbPDFRotate.AddItem .OptionsPDFGeneralRotate02
53670   cmbPDFRotate.AddItem .OptionsPDFGeneralRotate03
53680   cmbPDFOverprint.Clear
53690   cmbPDFOverprint.AddItem .OptionsPDFGeneralOverprint01
53700   cmbPDFOverprint.AddItem .OptionsPDFGeneralOverprint02
53710
53720   dmfraPDFCompress.Caption = .OptionsPDFCompressionCaption
53730   chkPDFTextComp.Caption = .OptionsPDFCompressionTextComp
53740   dmFraPDFColor.Caption = .OptionsPDFCompressionColor
53750   chkPDFColorComp.Caption = .OptionsPDFCompressionColorComp
53760   chkPDFColorResample.Caption = .OptionsPDFCompressionColorResample
53770   lblPDFColorRes.Caption = .OptionsPDFCompressionColorRes
53780   cmbPDFColorComp.Clear
53790   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp01
53800   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp02
53810   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp03
53820   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp04
53830   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp05
53840   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp06
53850   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp07
53860 '  cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp08
53870   cmbPDFColorResample.Clear
53880   cmbPDFColorResample.AddItem .OptionsPDFCompressionColorResample01
53890   cmbPDFColorResample.AddItem .OptionsPDFCompressionColorResample02
53900 '  cmbPDFColorResample.AddItem .OptionsPDFCompressionColorResample03
53910   dmFraPDFGrey.Caption = .OptionsPDFCompressionGrey
53920   chkPDFGreyComp.Caption = .OptionsPDFCompressionGreyComp
53930   chkPDFGreyResample.Caption = .OptionsPDFCompressionGreyResample
53940   lblPDFGreyRes.Caption = .OptionsPDFCompressionGreyRes
53950   cmbPDFGreyComp.Clear
53960   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp01
53970   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp02
53980   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp03
53990   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp04
54000   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp05
54010   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp06
54020   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp07
54030 '  cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp08
54040   cmbPDFGreyResample.Clear
54050   cmbPDFGreyResample.AddItem .OptionsPDFCompressionGreyResample01
54060   cmbPDFGreyResample.AddItem .OptionsPDFCompressionGreyResample02
54070 '  cmbPDFGreyResample.AddItem .OptionsPDFCompressionGreyResample03
54080   dmFraPDFMono.Caption = .OptionsPDFCompressionMono
54090   chkPDFMonoComp.Caption = .OptionsPDFCompressionMonoComp
54100   chkPDFMonoResample.Caption = .OptionsPDFCompressionMonoResample
54110   lblPDFMonoRes.Caption = .OptionsPDFCompressionMonoRes
54120   cmbPDFMonoComp.Clear
54130   cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp01
54140   cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp02
54150   cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp03
54160 '  cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp04
54170   cmbPDFMonoResample.Clear
54180   cmbPDFMonoResample.AddItem .OptionsPDFCompressionMonoResample01
54190   cmbPDFMonoResample.AddItem .OptionsPDFCompressionMonoResample02
54200 '  cmbPDFMonoResample.AddItem .OptionsPDFCompressionMonoResample03
54210
54220   dmFraPDFFonts.Caption = .OptionsPDFFontsCaption
54230   chkPDFEmbedAll.Caption = .OptionsPDFFontsEmbedAll
54240   chkPDFSubSetFonts.Caption = .OptionsPDFFontsSubSetFonts
54250
54260   dmFraPDFColors.Caption = .OptionsPDFColorsCaption
54270   chkPDFCMYKtoRGB.Caption = .OptionsPDFColorsCMYKtoRGB
54280   dmFraPDFColorOptions.Caption = .OptionsPDFColorsColorOptions
54290   chkPDFPreserveOverprint.Caption = .OptionsPDFColorsPreserveOverprint
54300   chkPDFPreserveTransfer.Caption = .OptionsPDFColorsPreserveTransfer
54310   chkPDFPreserveHalftone.Caption = .OptionsPDFColorsPreserveHalftone
54320   cmbPDFColorModel.Clear
54330   cmbPDFColorModel.AddItem .OptionsPDFColorsColorModel01
54340   cmbPDFColorModel.AddItem .OptionsPDFColorsColorModel02
54350   cmbPDFColorModel.AddItem .OptionsPDFColorsColorModel03
54360
54370   dmFraPDFEncryptor.Caption = .OptionsPDFEncryptor
54380   dmFraPDFSecurity.Caption = .OptionsPDFSecurityCaption
54390   chkUseSecurity.Caption = .OptionsPDFUseSecurity
54400   dmFraPDFEncLevel.Caption = .OptionsPDFEncryptionLevel
54410   optEncHigh.Caption = .OptionsPDFEncryptionHigh
54420   optEncLow.Caption = .OptionsPDFEncryptionLow
54430   dmFraSecurityPass.Caption = .OptionsPDFPasswords
54440   chkUserPass.Caption = .OptionsPDFUserPass
54450   chkOwnerPass.Caption = .OptionsPDFOwnerPass
54460   dmFraPDFPermissions.Caption = .OptionsPDFDisallowUser
54470   dmFraPDFHighPermissions.Caption = .OptionsPDFEnhancedPermissions
54480   chkAllowPrinting.Caption = .OptionsPDFDisallowPrint
54490   chkAllowModifyContents.Caption = .OptionsPDFDisallowModify
54500   chkAllowCopy.Caption = .OptionsPDFDisallowCopy
54510   chkAllowModifyAnnotations.Caption = .OptionsPDFDisallowModifyComments
54520   chkAllowDegradedPrinting.Caption = .OptionsPDFAllowDegradedPrinting
54530   chkAllowFillIn.Caption = .OptionsPDFAllowFillIn
54540   chkAllowAssembly.Caption = .OptionsPDFAllowAssembly
54550   chkAllowScreenReaders.Caption = .OptionsPDFAllowScreenReaders
54560
54570   cmbPNGColors.AddItem .OptionsPNGColorscount01
54580   cmbPNGColors.AddItem .OptionsPNGColorscount02
54590   cmbPNGColors.AddItem .OptionsPNGColorscount03
54600   cmbPNGColors.AddItem .OptionsPNGColorscount04
54610   cmbJPEGColors.Left = cmbPNGColors.Left
54620   cmbJPEGColors.Width = cmbPNGColors.Width
54630   cmbJPEGColors.Top = cmbPNGColors.Top
54640   cmbJPEGColors.AddItem .OptionsJPEGColorscount01
54650   cmbJPEGColors.AddItem .OptionsJPEGColorscount02
54660   cmbBMPColors.Left = cmbPNGColors.Left
54670   cmbBMPColors.Width = cmbPNGColors.Width
54680   cmbBMPColors.Top = cmbPNGColors.Top
54690   cmbBMPColors.AddItem .OptionsBMPColorscount01
54700   cmbBMPColors.AddItem .OptionsBMPColorscount02
54710   cmbBMPColors.AddItem .OptionsBMPColorscount03
54720   cmbBMPColors.AddItem .OptionsBMPColorscount04
54730   cmbBMPColors.AddItem .OptionsBMPColorscount05
54740   cmbBMPColors.AddItem .OptionsBMPColorscount06
54750   cmbBMPColors.AddItem .OptionsBMPColorscount07
54760   cmbPCXColors.Left = cmbPNGColors.Left
54770   cmbPCXColors.Width = cmbPNGColors.Width
54780   cmbPCXColors.Top = cmbPNGColors.Top
54790   cmbPCXColors.AddItem .OptionsPCXColorscount01
54800   cmbPCXColors.AddItem .OptionsPCXColorscount02
54810   cmbPCXColors.AddItem .OptionsPCXColorscount03
54820   cmbPCXColors.AddItem .OptionsPCXColorscount04
54830   cmbPCXColors.AddItem .OptionsPCXColorscount05
54840   cmbPCXColors.AddItem .OptionsPCXColorscount06
54850   cmbTIFFColors.Left = cmbPNGColors.Left
54860   cmbTIFFColors.Width = cmbPNGColors.Width
54870   cmbTIFFColors.Top = cmbPNGColors.Top
54880   cmbTIFFColors.AddItem .OptionsTIFFColorscount01
54890   cmbTIFFColors.AddItem .OptionsTIFFColorscount02
54900   cmbTIFFColors.AddItem .OptionsTIFFColorscount03
54910   cmbTIFFColors.AddItem .OptionsTIFFColorscount04
54920   cmbTIFFColors.AddItem .OptionsTIFFColorscount05
54930   cmbTIFFColors.AddItem .OptionsTIFFColorscount06
54940   cmbTIFFColors.AddItem .OptionsTIFFColorscount07
54950   cmbTIFFColors.AddItem .OptionsTIFFColorscount08
54960
54970   dmFraBitmapGeneral.Caption = .OptionsImageSettings
54980   lblBitmapResolution = .OptionsBitmapResolution
54990   lblJPEGQuality = .OptionsJPEGQuality
55000   lblBitmapColors = .OptionsPDFColors
55010   lblProcessPriority.Caption = .OptionsProcesspriority
55020   lblLangLevel.Caption = .OptionsPSLanguageLevel
55030
55040   cmdAsso.Caption = .OptionsAssociatePSFiles
55050  End With
55060
55070  If IsPsAssociate = False Then
55080    cmdAsso.Enabled = True
55090   Else
55100    cmdAsso.Enabled = False
55110  End If
55120
55130  txtPDFRes.Text = 600
55140  cmbPDFCompat.ListIndex = 1
55150  cmbPDFRotate.ListIndex = 0
55160  cmbPDFOverprint.ListIndex = 0
55170  chkPDFASCII85.Value = 0
55180
55190  chkPDFTextComp.Value = 1
55200
55210  chkPDFColorComp.Value = 1
55220  chkPDFColorResample.Value = 0
55230  cmbPDFColorComp.ListIndex = 0
55240  cmbPDFColorResample.ListIndex = 0
55250  txtPDFColorRes.Text = 300
55260
55270  chkPDFGreyComp.Value = 1
55280  chkPDFGreyResample.Value = 0
55290  cmbPDFGreyComp.ListIndex = 0
55300  cmbPDFGreyResample.ListIndex = 0
55310  txtPDFGreyRes.Text = 300
55320
55330  chkPDFMonoComp.Value = 1
55340  chkPDFMonoResample.Value = 0
55350  cmbPDFMonoComp.ListIndex = 0
55360  cmbPDFMonoResample.ListIndex = 0
55370  txtPDFMonoRes.Text = 1200
55380
55390  chkPDFEmbedAll.Value = 1
55400  chkPDFSubSetFonts.Value = 1
55410  txtPDFSubSetPerc.Text = 100
55420
55430  cmbPDFColorModel.ListIndex = 1
55440  chkPDFCMYKtoRGB.Value = 1
55450  chkPDFPreserveOverprint.Value = 1
55460  chkPDFPreserveTransfer.Value = 1
55470  chkPDFPreserveHalftone.Value = 0
55480
55490  cmbPNGColors.ListIndex = 0
55500  cmbJPEGColors.ListIndex = 0
55510  cmbBMPColors.ListIndex = 0
55520  cmbPCXColors.ListIndex = 0
55530  cmbTIFFColors.ListIndex = 0
55540  txtBitmapResolution.Text = 150
55550
55560 ' chkUseStandardAuthor.Value = 1
55570  txtStandardAuthor.Text = vbNullString
55580
55590  With cmbPSLanguageLevel
55600   .AddItem "1"
55610   .AddItem "1.5"
55620   .AddItem "2"
55630   .AddItem "3"
55640  End With
55650  With cmbEPSLanguageLevel
55660   .AddItem "1"
55670   .AddItem "1.5"
55680   .AddItem "2"
55690   .AddItem "3"
55700  End With
55710
55720  With lsvFilenameSubst
55730   .Appearance = ccFlat
55740   .ColumnHeaders.Clear
55750   .ColumnHeaders.Add , "Str1", "", lsvFilenameSubst.Width / 2 - 140
55760   .ColumnHeaders.Add , "Str2", "", lsvFilenameSubst.Width / 2 - 140
55770   .HideColumnHeaders = True
55780   .GridLines = True
55790   .FullRowSelect = True
55800   .HideSelection = False
55810  End With
55820
55830  With cmbPDFEncryptor
55840   .Clear
55850   .AddItem "Ghostscript (>= 8.14)"
55860   .ItemData(.NewIndex) = 0
55870   .AddItem "PDFEnc"
55880   .ItemData(.NewIndex) = 1
55890
55900   SecurityIsPossible = True
55910
55920   If FileExists(GetPDFCreatorApplicationPath & "pdfenc.exe") = False Then
55930    .RemoveItem 1
55940    .ListIndex = 0
55950    Options.PDFEncryptor = .ItemData(.ListIndex)
55960   End If
55970   If GhostScriptSecurity = False Then
55980    .RemoveItem 0
55990   End If
56000   If .ListCount = 0 Then
56010     chkUseSecurity.Value = 0
56020     chkUseSecurity.Enabled = False
56030     SecurityIsPossible = False
56040    Else
56050     For i = 0 To .ListCount - 1
56060      If .ItemData(i) = Options.PDFEncryptor Then
56070       .ListIndex = i
56080       Exit For
56090      End If
56100     Next i
56110     If .ListIndex = -1 Then
56120      .ListIndex = 0
56130      Options.PDFEncryptor = .ItemData(.ListIndex)
56140     End If
56150   End If
56160  End With
56170
56180  If Options.PDFHighEncryption <> 0 Then
56190    optEncHigh.Value = True
56200   Else
56210    optEncLow.Value = True
56220  End If
56230
56240  cmdFilenameSubst(0).Top = lsvFilenameSubst.Top
56250  cmdFilenameSubst(1).Top = lsvFilenameSubst.Top + (lsvFilenameSubst.Height - cmdFilenameSubst(1).Height) / 2
56260  cmdFilenameSubst(2).Top = lsvFilenameSubst.Top + lsvFilenameSubst.Height - cmdFilenameSubst(2).Height
56270
56280  If chkUseStandardAuthor.Value = 1 Then
56290    txtStandardAuthor.Enabled = True
56300    txtStandardAuthor.BackColor = &H80000005
56310   Else
56320    txtStandardAuthor.Enabled = False
56330    txtStandardAuthor.BackColor = &H8000000F
56340  End If
56350  With Options
56360   SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
56370  End With
56380  ieb.Refresh
56390  If chkUseAutosave.Value = 1 Then
56400    ViewAutosave True
56410   Else
56420    ViewAutosave False
56430  End If
56440  If chkPrintAfterSaving.Value = 1 Then
56450    ViewPrintAfterSaving True
56460   Else
56470    ViewPrintAfterSaving False
56480  End If
56490
56500  With txtGSbin
56510   .ToolTipText = .Text
56520  End With
56530  With txtGSlib
56540   .ToolTipText = .Text
56550  End With
56560  With txtGSfonts
56570   .ToolTipText = .Text
56580  End With
56590  With txtTemppath
56600   .ToolTipText = ResolveEnvironment(GetSubstFilename2(.Text))
56610  End With
56620
56630  With sldProcessPriority
56640   .TextPosition = sldBelowRight
56650   .TickFrequency = 1
56660   .TickStyle = sldTopLeft
56671   Select Case .Value
         Case 0: 'Idle
56690     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityIdle
56700    Case 1: 'Normal
56710     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityNormal
56720    Case 2: 'High
56730     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityHigh
56740    Case 3: 'Realtime
56750     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityRealtime
56760   End Select
56770  End With
56780
56790  If IsWin9xMe = False Then
56800    lblProcessPriority.Enabled = True
56810    sldProcessPriority.Enabled = True
56820   Else
56830    lblProcessPriority.Enabled = False
56840    sldProcessPriority.Enabled = False
56850  End If
56860  UpdateSecurityFields
56870
56880  If Options.RunProgramAfterSaving Then
56890    ViewRunProgramAfterSaving True
56900   Else
56910    ViewRunProgramAfterSaving False
56920  End If
56930  If Options.RunProgramBeforeSaving Then
56940    ViewRunProgramBeforeSaving True
56950   Else
56960    ViewRunProgramBeforeSaving False
56970  End If
56980
56990  Set Files = GetFiles(GetPDFCreatorApplicationPath & "Scripts\RunProgramAfterSaving\", "*.*", SortedByName)
57000  For i = 1 To Files.Count
57010   tsf = Split(Files(i), "|")
57020   SplitPath tsf(1), , Path, Filename, , Ext
57030   If UCase$(Ext) <> "TXT" And UCase$(Ext) <> "PDF" And UCase$(Ext) <> "PNG" And _
   UCase$(Ext) <> "JPG" And UCase$(Ext) <> "BMP" And UCase$(Ext) <> "PCX" And _
   UCase$(Ext) <> "TIF" And UCase$(Ext) <> "EPS" And UCase$(Ext) <> "PS" Then
57060    If UCase$(tsf(0)) <> UCase$(GetPDFCreatorApplicationPath & "Scripts\RunProgramAfterSaving\") Then
57070      cmbRunProgramAfterSavingProgramname.AddItem tsf(0)
57080     Else
57090      cmbRunProgramAfterSavingProgramname.AddItem Filename
57100    End If
57110   End If
57120  Next i
57130
57140  Set Files = GetFiles(GetPDFCreatorApplicationPath & "Scripts\RunProgramBeforeSaving\", "*.*", SortedByName)
57150  For i = 1 To Files.Count
57160   tsf = Split(Files(i), "|")
57170   SplitPath tsf(1), , Path, Filename, , Ext
57180   If UCase$(Ext) <> "TXT" And UCase$(Ext) <> "PDF" And UCase$(Ext) <> "PNG" And _
   UCase$(Ext) <> "JPG" And UCase$(Ext) <> "BMP" And UCase$(Ext) <> "PCX" And _
   UCase$(Ext) <> "TIF" And UCase$(Ext) <> "EPS" And UCase$(Ext) <> "PS" Then
57210    If UCase$(tsf(0)) <> UCase$(GetPDFCreatorApplicationPath & "Scripts\RunProgramBeforeSaving\") Then
57220      cmbRunProgramBeforeSavingProgramname.AddItem tsf(0)
57230     Else
57240      cmbRunProgramBeforeSavingProgramname.AddItem Filename
57250    End If
57260   End If
57270  Next i
57280
57290  tStr2 = CompletePath(UCase$(Trim$(Options.DirectoryGhostscriptBinaries)))
 cmbGhostscript.Clear: Set reg = New clsRegistry
57310  reg.hkey = HKEY_LOCAL_MACHINE
57320
57330  Set gsvers = GetAllGhostscriptversions
57340
57350  If gsvers.Count = 0 Then
57360    cmbGhostscript.Enabled = False
57370   Else
57380    For i = 1 To gsvers.Count
57390     cmbGhostscript.AddItem gsvers.Item(i)
57400    Next i
57410    cmbGhostscript.ListIndex = cmbGhostscript.ListCount - 1
57420    For i = 0 To cmbGhostscript.ListCount - 1
57430     tStr = ""
57440     If InStr(cmbGhostscript.List(i), ":") Then
57450       reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
57460       If tStr2 = CompletePath(UCase$(Trim$(reg.GetRegistryValue("GhostscriptDirectoryBinaries")))) Then
57470        cmbGhostscript.ListIndex = i
57480        Exit For
57490       End If
57500      Else
57510       If InStr(UCase$(cmbGhostscript.List(i)), "AFPL") Then
57520        reg.KeyRoot = "SOFTWARE\AFPL Ghostscript"
57530        If InStr(cmbGhostscript.List(i), " ") > 0 Then
57540         tsf = Split(cmbGhostscript.List(i), " ")
57550         reg.Subkey = tsf(UBound(tsf))
57560         tStr = reg.GetRegistryValue("GS_DLL")
57570         If tStr2 & "GSDLL32.DLL" = UCase$(tStr) Then
57580          cmbGhostscript.ListIndex = i
57590          Exit For
57600         End If
57610        End If
57620       End If
57630       If InStr(UCase$(cmbGhostscript.List(i)), "GNU") Then
57640        reg.KeyRoot = "SOFTWARE\GNU Ghostscript"
57650        If InStr(cmbGhostscript.List(i), " ") > 0 Then
57660         tsf = Split(cmbGhostscript.List(i), " ")
57670         reg.Subkey = tsf(UBound(tsf))
57680         tStr = reg.GetRegistryValue("GS_DLL")
57690         If tStr2 & "GSDLL32.DLL" = UCase$(tStr) Then
57700          cmbGhostscript.ListIndex = i
57710          Exit For
57720         End If
57730        End If
57740       End If
57750       If InStr(UCase$(cmbGhostscript.List(i)), "GPL") Then
57760        reg.KeyRoot = "SOFTWARE\GPL Ghostscript"
57770        If InStr(cmbGhostscript.List(i), " ") > 0 Then
57780         tsf = Split(cmbGhostscript.List(i), " ")
57790         reg.Subkey = tsf(UBound(tsf))
57800         tStr = reg.GetRegistryValue("GS_DLL")
57810         If tStr2 & "GSDLL32.DLL" = UCase$(tStr) Then
57820          cmbGhostscript.ListIndex = i
57830          Exit For
57840         End If
57850        End If
57860       End If
57870     End If
57880    Next i
57890  End If
57900  Set reg = Nothing
57910  With cmbGhostscript
57920   If .ListCount = 0 Then
57930    .Enabled = False
57940    .BackColor = &H8000000F
57950   End If
57960  End With
57970
57980  tbstrPDFOptions.ZOrder 1
57990  tbstrProgActions.ZOrder 1
58000
58010  If ShowOnlyOptions = True Then
58020   FormInTaskbar Me, True, True
58030   Caption = "PDFCreator - " & Caption
58040  End If
58050
58060  ShowAcceleratorsInForm Me, True
58070
58080  ShowOptions Me, Options
58090  Timer1.Enabled = True
58100  Screen.MousePointer = vbNormal
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

Private Sub ieb_ItemClick(sGroup As String, sItemKey As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ctl As Control
50020  lblJPEGQuality.Visible = False
50030  cmbPNGColors.Visible = False
50040  cmbJPEGColors.Visible = False
50050  cmbBMPColors.Visible = False
50060  cmbPCXColors.Visible = False
50070  cmbTIFFColors.Visible = False
50080  tbstrPDFOptions.Visible = False
50090  tbstrProgGeneral.Visible = False
50100  For Each ctl In Controls
50110   If TypeOf ctl Is dmFrame Then
50120    ctl.Visible = False
50130    ctl.Enabled = False
50140   End If
50150  Next
50160  dmFraDescription.Visible = True
50170  dmFraDescription.Enabled = True
50180  tbstrPDFOptions.Enabled = False
50190  txtJPEGQuality.Visible = False
50200  lblJPEQQualityProzent.Visible = False
50210  dmFraPSGeneral.Visible = False
50220  cmbPSLanguageLevel.Visible = False
50230  cmbEPSLanguageLevel.Visible = False
50240
50251  Select Case UCase$(sGroup)
        Case "PROGRAM"
50271    Select Case UCase$(sItemKey)
          Case "GENERAL"
50290      Set picOptions = LoadResPicture(2101, vbResIcon)
50300      lblOptions = LanguageStrings.OptionsProgramGeneralDescription
50310      tbstrProgGeneral.Enabled = True
50320      tbstrProgGeneral.Visible = True
50331      Select Case tbstrProgGeneral.SelectedItem.Index
            Case 1
50350        dmFraProgGeneral1.Enabled = True
50360        dmFraProgGeneral1.Visible = True
50370       Case 2
50380        dmFraProgGeneral2.Enabled = True
50390        dmFraProgGeneral2.Visible = True
50400        dmFraShellIntegration.Enabled = True
50410        dmFraShellIntegration.Visible = True
50420      End Select
50430      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50440     Case "GHOSTSCRIPT"
50450      Set picOptions = LoadResPicture(2119, vbResIcon)
50460      lblOptions = LanguageStrings.OptionsProgramGhostscriptDescription
50470      dmFraProgGhostscript.Enabled = True
50480      dmFraProgGhostscript.Visible = True
50490      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50500     Case "DOCUMENT"
50510      Set picOptions = LoadResPicture(2105, vbResIcon)
50520      lblOptions = LanguageStrings.OptionsProgramDocumentDescription
50530      dmFraProgDocument.Enabled = True
50540      dmFraProgDocument.Visible = True
50550      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50560     Case "SAVE"
50570      Set picOptions = LoadResPicture(2106, vbResIcon)
50580      lblOptions = LanguageStrings.OptionsProgramSaveDescription
50590      dmFraProgSave.Enabled = True
50600      dmFraProgSave.Visible = True
50610      dmFraFilenameSubstitutions.Visible = True
50620      dmFraFilenameSubstitutions.Enabled = True
50630      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50640     Case "AUTOSAVE"
50650      Set picOptions = LoadResPicture(2103, vbResIcon)
50660      lblOptions = LanguageStrings.OptionsProgramAutosaveDescription
50670      dmFraProgAutosave.Enabled = True
50680      dmFraProgAutosave.Visible = True
50690      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50700     Case "DIRECTORIES"
50710      Set picOptions = LoadResPicture(2104, vbResIcon)
50720      lblOptions = LanguageStrings.OptionsProgramDirectoriesDescription
50730      dmFraProgDirectories.Enabled = True
50740      dmFraProgDirectories.Visible = True
50750      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50760     Case "ACTIONS"
50770      Set picOptions = LoadResPicture(2121, vbResIcon)
50780      lblOptions = LanguageStrings.OptionsProgramActionsDescription
50790      dmFraProgActions.Enabled = True
50800      dmFraProgActions.Visible = True
50810      ViewProgActions
50820      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50830     Case "PRINT"
50840      Set picOptions = LoadResPicture(2122, vbResIcon)
50850      lblOptions = LanguageStrings.OptionsProgramPrintDescription
50860      dmFraProgPrint.Enabled = True
50870      dmFraProgPrint.Visible = True
50880      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50890     Case "FONTS"
50900      Set picOptions = LoadResPicture(2102, vbResIcon)
50910      lblOptions = LanguageStrings.OptionsProgramFontDescription
50920      dmFraProgFont.Enabled = True
50930      dmFraProgFont.Visible = True
50940    End Select
50950   Case "FORMATS"
50961    Select Case UCase$(sItemKey)
          Case "PDF"
50980      Set picOptions = LoadResPicture(2111, vbResIcon)
50990      lblOptions = LanguageStrings.OptionsPDFDescription
51000      tbstrPDFOptions.Enabled = True
51010      tbstrPDFOptions.Visible = True
51020      tbstrPDFOptions_Click
51030      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51040     Case "PNG"
51050      Set picOptions = LoadResPicture(2112, vbResIcon)
51060      lblOptions = LanguageStrings.OptionsPNGDescription
51070      dmFraBitmapGeneral.Enabled = True
51080      dmFraBitmapGeneral.Visible = True
51090      cmbPNGColors.Visible = True
51100      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51110     Case "JPEG"
51120      Set picOptions = LoadResPicture(2113, vbResIcon)
51130      lblOptions = LanguageStrings.OptionsJPEGDescription
51140      dmFraBitmapGeneral.Enabled = True
51150      dmFraBitmapGeneral.Visible = True
51160      lblJPEGQuality.Caption = LanguageStrings.OptionsJPEGQuality
51170      lblJPEGQuality.Visible = True
51180      txtJPEGQuality.Visible = True
51190      lblJPEQQualityProzent.Visible = True
51200      lblJPEQQualityProzent.Left = txtJPEGQuality.Left + txtJPEGQuality.Width + 100
51210      cmbJPEGColors.Visible = True
51220      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51230     Case "BMP"
51240      Set picOptions = LoadResPicture(2114, vbResIcon)
51250      lblOptions = LanguageStrings.OptionsBMPDescription
51260      dmFraBitmapGeneral.Enabled = True
51270      dmFraBitmapGeneral.Visible = True
51280      cmbBMPColors.Visible = True
51290      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51300     Case "PCX"
51310      Set picOptions = LoadResPicture(2115, vbResIcon)
51320      lblOptions = LanguageStrings.OptionsPCXDescription
51330      dmFraBitmapGeneral.Enabled = True
51340      dmFraBitmapGeneral.Visible = True
51350      cmbPCXColors.Visible = True
51360      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51370     Case "TIFF"
51380      Set picOptions = LoadResPicture(2116, vbResIcon)
51390      lblOptions = LanguageStrings.OptionsTIFFDescription
51400      dmFraBitmapGeneral.Enabled = True
51410      dmFraBitmapGeneral.Visible = True
51420      cmbTIFFColors.Visible = True
51430      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51440     Case "PS"
51450      Set picOptions = LoadResPicture(2117, vbResIcon)
51460      lblOptions.Caption = LanguageStrings.OptionsPSDescription
51470      dmFraPSGeneral.Enabled = True
51480      dmFraPSGeneral.Visible = True
51490      cmbPSLanguageLevel.Visible = True
51500      dmFraPSGeneral.Caption = LanguageStrings.OptionsPSDescription
51510      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51520     Case "EPS"
51530      Set picOptions = LoadResPicture(2118, vbResIcon)
51540      lblOptions.Caption = LanguageStrings.OptionsEPSDescription
51550      dmFraPSGeneral.Enabled = True
51560      dmFraPSGeneral.Visible = True
51570      cmbEPSLanguageLevel.Visible = True
51580      dmFraPSGeneral.Caption = LanguageStrings.OptionsEPSDescription
51590      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51600    End Select
51610  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "ieb_ItemClick")
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
50010  dmFraPDFGeneral.Visible = False
50020  dmfraPDFCompress.Visible = False
50030  dmFraPDFFonts.Visible = False
50040  dmFraPDFColors.Visible = False
50050  dmFraPDFColorOptions.Visible = False
50060  dmFraPDFSecurity.Visible = False
50070  dmFraPDFGeneral.Enabled = False
50080  dmfraPDFCompress.Enabled = False
50090  dmFraPDFFonts.Enabled = False
50100  dmFraPDFColors.Enabled = False
50110  dmFraPDFColorOptions.Enabled = False
50120  dmFraPDFSecurity.Enabled = False
50131  Select Case tbstrPDFOptions.SelectedItem.Index
        Case 1:
50150    dmFraPDFGeneral.Visible = True
50160    dmFraPDFGeneral.Enabled = True
50170   Case 2:
50180    dmfraPDFCompress.Visible = True
50190    dmfraPDFCompress.Enabled = True
50200    dmFraPDFColor.Visible = True
50210    dmFraPDFColor.Enabled = True
50220    dmFraPDFGrey.Visible = True
50230    dmFraPDFGrey.Enabled = True
50240    dmFraPDFMono.Visible = True
50250    dmFraPDFMono.Enabled = True
50260   Case 3:
50270    dmFraPDFFonts.Visible = True
50280    dmFraPDFFonts.Enabled = True
50290   Case 4:
50300    dmFraPDFColors.Visible = True
50310    dmFraPDFColorOptions.Visible = True
50320    dmFraPDFColors.Enabled = True
50330    dmFraPDFColorOptions.Enabled = True
50340   Case 5:
50350    dmFraPDFSecurity.Visible = True
50360    dmFraPDFSecurity.Enabled = True
50370    dmFraPDFEncryptor.Visible = True
50380    dmFraPDFEncryptor.Enabled = True
50390    dmFraPDFEncLevel.Visible = True
50400    dmFraPDFEncLevel.Enabled = True
50410    dmFraSecurityPass.Visible = True
50420    dmFraSecurityPass.Enabled = True
50430    dmFraPDFPermissions.Visible = True
50440    dmFraPDFPermissions.Enabled = True
50450    dmFraPDFHighPermissions.Visible = True
50460    dmFraPDFHighPermissions.Enabled = True
50470    If SecurityIsPossible = False Then
50480     MsgBox LanguageStrings.MessagesMsg19
50490    End If
50500  End Select
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

Private Sub tbstrProgActions_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ViewProgActions
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "tbstrProgActions_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub tbstrProgGeneral_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case tbstrProgGeneral.SelectedItem.Index
        Case 1
50030    dmFraProgGeneral1.Enabled = True
50040    dmFraProgGeneral1.Visible = True
50050    dmFraProgGeneral2.Enabled = False
50060    dmFraProgGeneral2.Visible = False
50070    dmFraShellIntegration.Enabled = False
50080    dmFraShellIntegration.Visible = False
50090   Case 2
50100    dmFraProgGeneral2.Enabled = True
50110    dmFraProgGeneral2.Visible = True
50120    dmFraShellIntegration.Enabled = True
50130    dmFraShellIntegration.Visible = True
50140    dmFraProgGeneral1.Enabled = False
50150    dmFraProgGeneral1.Visible = False
50160  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "tbstrProgGeneral_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Timer1_Timer()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, fi As Long, tStr As String, SMF As Collection, _
  cSystem As clsSystem, ctl As Control
50030  Timer1.Enabled = False
50040  Set cSystem = New clsSystem
50050  Set SMF = cSystem.GetSystemFont(Me, Menu)
50060  txtTest.Text = vbNullString
50070  For i = 33 To 255
50080   txtTest.Text = txtTest.Text & Chr$(i)
50090  Next i
50100  fi = -1
50110  With cmbFonts
50120   .Clear
50130   For i = 1 To Screen.FontCount
50140    tStr = Trim$(Screen.Fonts(i))
50150    If Len(tStr) > 0 Then
50160     cmbFonts.AddItem tStr
50170    End If
50180   Next i
50190   If .ListCount > 0 Then
50200     For i = 0 To cmbFonts.ListCount - 1
50210      If SMF.Count > 0 Then
50220       If UCase$(cmbFonts.List(i)) = UCase$(SMF(1)(0)) Then
50230        fi = i
50240       End If
50250      End If
50260     Next i
50270    Else
50280    .ListIndex = 0
50290   End If
50300  End With
50310  With cmbCharset
50320   .Clear
50330   .AddItem "0, Western": .ItemData(.NewIndex) = 0
50340   .AddItem "2, Symbol": .ItemData(.NewIndex) = 2
50350   .AddItem "77, Mac": .ItemData(.NewIndex) = 77
50360   .AddItem "161, Greek": .ItemData(.NewIndex) = 161
50370   .AddItem "162, Turkish": .ItemData(.NewIndex) = 162
50380   .AddItem "177, Hebrew": .ItemData(.NewIndex) = 177
50390   .AddItem "178, Arabic": .ItemData(.NewIndex) = 178
50400   .AddItem "186, Baltic": .ItemData(.NewIndex) = 186
50410   .AddItem "204, Cyrillic": .ItemData(.NewIndex) = 204
50420   .AddItem "238, Central European": .ItemData(.NewIndex) = 238
50430   .AddItem "255, DOS/OEM": .ItemData(.NewIndex) = 255
50440   .Text = 0
50450  End With
50460  With cmbProgramFontsize
50470   .AddItem "8"
50480   .AddItem "9"
50490   .AddItem "10"
50500   .AddItem "11"
50510   .AddItem "12"
50520   .AddItem "14"
50530   .AddItem "16"
50540   .AddItem "18"
50550   .AddItem "20"
50560   .AddItem "22"
50570   .AddItem "24"
50580   .AddItem "26"
50590   .AddItem "28"
50600   .AddItem "36"
50610   .AddItem "48"
50620   .AddItem "72"
50630  End With
50640  cmbProgramFontsize.Text = 8
50650  cmbCharset.Text = cmbCharset.ItemData(0)
50660  cmbCharset.Text = Options.ProgramFontCharset
50670  For Each ctl In Controls
50680   If TypeOf ctl Is ComboBox Then
50690    ComboSetListWidth ctl
50700   End If
50710  Next ctl
50720
50730  SetOptimalComboboxHeigth cmbCharset, Me
50740  SetOptimalComboboxHeigth cmbProgramFontsize, Me
50750  SetOptimalComboboxHeigth cmbGhostscript, Me
50760
50770  Form_Resize
50780
50790  cmbProgramFontsize.Width = txtTest.Width - _
  (cmbProgramFontsize.Left - txtTest.Left)
50810
50820  If fi >= 0 Then
50830   cmbFonts.ListIndex = fi
50840   cmbCharset.Text = SMF(1)(2)
50850   cmbProgramFontsize.Text = SMF(1)(1)
50860   txtTest.Font = cmbFonts.List(cmbFonts.ListIndex)
50870   txtTest.Font.Charset = cmbCharset.Text
50880  End If
50890
50900  ShowOptions Me, Options
50910
50920  If Options.UseAutosaveDirectory = "1" Then
50930    ViewAutosaveDirectory True
50940   Else
50950    ViewAutosaveDirectory False
50960  End If
50970  If Options.UseAutosave = "1" Then
50980    ViewAutosave True
50990   Else
51000    ViewAutosave False
51010  End If
51020
51030  CheckCmdFilenameSubst
51040  CorrectCmbCharset
51050  tbstrProgActions.Tabs(2).Selected = True
51060 ' Call ieb_ItemClick("PROGRAM", "GENERAL")
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "Timer1_Timer")
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
50010  txtAutosaveDirectory.ToolTipText = txtAutosaveDirectory.Text
50020  With txtAutoSaveDirectoryPreview
50030   .Text = GetSubstFilename2(txtAutosaveDirectory.Text)
50040   .ToolTipText = .Text
50050   If IsValidPath(.Text) = False Then
50060     .ForeColor = vbRed
50070    Else
50080     .ForeColor = &H80000008
50090   End If
50100  End With
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
50010  Dim Ext As String
50020  txtAutosaveFilename.ToolTipText = txtAutosaveFilename.Text
50030  With txtAutoSaveFilenamePreview
50040   .Text = GetSubstFilename("B:\dummy.dum", txtAutosaveFilename.Text, , True) & GetSaveAutosaveFormatExtension(cmbAutosaveFormat.ListIndex)
50050   .ToolTipText = .Text
50060   If IsValidPath("C:\" & .Text) = False Then
50070     .ForeColor = vbRed
50080    Else
50090     .ForeColor = &H80000008
50100   End If
50110  End With
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
50100  chkAutosaveStartStandardProgram.Enabled = ViewIt
50110  chkAutosaveSendEmail.Enabled = ViewIt
50120
50130  If ViewIt Then
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
50240  If chkUseAutosaveDirectory.Value = 1 And ViewIt Then
50250    ViewAutosaveDirectory True
50260   Else
50270    ViewAutosaveDirectory False
50280  End If
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
50010  txtAutosaveDirectory.Enabled = ViewIt
50020  txtAutoSaveDirectoryPreview.Enabled = ViewIt
50030  cmdGetAutosaveDirectory.Enabled = ViewIt
50040  If ViewIt = True Then
50050    txtAutosaveDirectory.BackColor = &H80000005
50060   Else
50070    txtAutosaveDirectory.BackColor = &H8000000F
50080  End If
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

Private Sub ViewProgActions()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case tbstrProgActions.SelectedItem.Index
        Case 1
50030    dmFraProgActionsRunProgramBeforeSaving.Visible = True
50040    dmFraProgActionsRunProgramBeforeSaving.Enabled = True
50050    dmFraProgActionsRunProgramAfterSaving.Visible = False
50060    dmFraProgActionsRunProgramAfterSaving.Enabled = False
50070   Case 2
50080    dmFraProgActionsRunProgramAfterSaving.Visible = True
50090    dmFraProgActionsRunProgramAfterSaving.Enabled = True
50100    dmFraProgActionsRunProgramBeforeSaving.Visible = False
50110    dmFraProgActionsRunProgramBeforeSaving.Enabled = False
50120  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "ViewProgActions")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ViewRunProgramAfterSaving(ViewIt As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  lblRunProgramAfterSavingProgramname.Enabled = ViewIt
50020  cmbRunProgramAfterSavingProgramname.Enabled = ViewIt
50030  lblRunProgramAfterSavingProgramParameters.Enabled = ViewIt
50040  txtRunProgramAfterSavingProgramParameters.Enabled = ViewIt
50050  chkRunProgramAfterSavingWaitUntilReady.Enabled = ViewIt
50060  lblRunProgramAfterSavingWindowstyle.Enabled = ViewIt
50070  cmbRunProgramAfterSavingWindowstyle.Enabled = ViewIt
50080  cmdRunProgramAfterSavingPrognameChoice.Enabled = ViewIt
50090  cmdRunProgramAfterSavingPrognameEdit.Enabled = ViewIt
50100
50110  If ViewIt Then
50120    cmbRunProgramAfterSavingProgramname.BackColor = &H80000005
50130    cmbRunProgramAfterSavingWindowstyle.BackColor = &H80000005
50140    txtRunProgramAfterSavingProgramParameters.BackColor = &H80000005
50150   Else
50160    cmbRunProgramAfterSavingProgramname.BackColor = &H8000000F
50170    cmbRunProgramAfterSavingWindowstyle.BackColor = &H8000000F
50180    txtRunProgramAfterSavingProgramParameters.BackColor = &H8000000F
50190  End If
50200
50210  cmbRunProgramAfterSavingProgramname_Change
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "ViewRunProgramAfterSaving")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ViewRunProgramBeforeSaving(ViewIt As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  lblRunProgramBeforeSavingProgramname.Enabled = ViewIt
50020  cmbRunProgramBeforeSavingProgramname.Enabled = ViewIt
50030  lblRunProgramBeforeSavingProgramParameters.Enabled = ViewIt
50040  txtRunProgramBeforeSavingProgramParameters.Enabled = ViewIt
50050  lblRunProgramBeforeSavingWindowstyle.Enabled = ViewIt
50060  cmbRunProgramBeforeSavingWindowstyle.Enabled = ViewIt
50070  cmdRunProgramBeforeSavingPrognameChoice.Enabled = ViewIt
50080  cmdRunProgramBeforeSavingPrognameEdit.Enabled = ViewIt
50090
50100  If ViewIt Then
50110    cmbRunProgramBeforeSavingProgramname.BackColor = &H80000005
50120    cmbRunProgramBeforeSavingWindowstyle.BackColor = &H80000005
50130    txtRunProgramBeforeSavingProgramParameters.BackColor = &H80000005
50140   Else
50150    cmbRunProgramBeforeSavingProgramname.BackColor = &H8000000F
50160    cmbRunProgramBeforeSavingWindowstyle.BackColor = &H8000000F
50170    txtRunProgramBeforeSavingProgramParameters.BackColor = &H8000000F
50180  End If
50190
50200  cmbRunProgramBeforeSavingProgramname_Change
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "ViewRunProgramBeforeSaving")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ViewPrintAfterSaving(ViewIt As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  lblPrintAfterSavingPrinter.Enabled = ViewIt
50020  cmbPrintAfterSavingPrinter.Enabled = ViewIt
50030  lblPrintAfterSavingQueryUser.Enabled = ViewIt
50040  cmbPrintAfterSavingQueryUser.Enabled = ViewIt
50050  chkPrintAfterSavingNoCancel.Enabled = ViewIt
50060  chkPrintAfterSavingDuplex.Enabled = ViewIt
50070
50080  If ViewIt Then
50090    cmbPrintAfterSavingPrinter.BackColor = &H80000005
50100    cmbPrintAfterSavingQueryUser.BackColor = &H80000005
50110   Else
50120    cmbPrintAfterSavingPrinter.BackColor = &H8000000F
50130    cmbPrintAfterSavingQueryUser.BackColor = &H8000000F
50140  End If
50150
50160  If chkPrintAfterSavingDuplex.Value = 1 And ViewIt Then
50170    ViewPrintAfterTumple True
50180   Else
50190    ViewPrintAfterTumple False
50200  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "ViewPrintAfterSaving")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ViewPrintAfterTumple(ViewIt As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  cmbPrintAfterSavingTumble.Enabled = ViewIt
50020
50030  If ViewIt Then
50040    cmbPrintAfterSavingTumble.BackColor = &H80000005
50050   Else
50060    cmbPrintAfterSavingTumble.BackColor = &H8000000F
50070  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "ViewPrintAfterTumple")
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
50010  If cmbPDFCompat.ListIndex < 2 Then
50020   optEncLow.Value = True
50030  End If
50040  If chkUseSecurity.Value = False Then
50050    dmFraPDFEncryptor.Enabled = False
50060    cmbPDFEncryptor.Enabled = False
50070
50080    dmFraPDFEncLevel.Enabled = False
50090    optEncHigh.Enabled = False
50100    optEncLow.Enabled = False
50110
50120    dmFraSecurityPass.Enabled = False
50130    chkUserPass.Enabled = False
50140    chkOwnerPass.Enabled = False
50150
50160    dmFraPDFPermissions.Enabled = False
50170    chkAllowPrinting.Enabled = False
50180    chkAllowCopy.Enabled = False
50190    chkAllowModifyAnnotations.Enabled = False
50200    chkAllowModifyContents.Enabled = False
50210
50220    dmFraPDFHighPermissions.Enabled = False
50230    chkAllowDegradedPrinting.Enabled = False
50240    chkAllowFillIn.Enabled = False
50250    chkAllowScreenReaders.Enabled = False
50260    chkAllowAssembly.Enabled = False
50270   Else
50280    dmFraPDFEncryptor.Enabled = True
50290    cmbPDFEncryptor.Enabled = True
50300
50310    dmFraPDFEncLevel.Enabled = True
50320    If cmbPDFCompat.ListIndex >= 2 Then
50330      optEncHigh.Enabled = True
50340     Else
50350      optEncHigh.Enabled = False
50360    End If
50370    optEncLow.Enabled = True
50380
50390    dmFraSecurityPass.Enabled = True
50400    chkUserPass.Enabled = True
50410    chkOwnerPass.Enabled = True
50420
50430    dmFraPDFPermissions.Enabled = True
50440    chkAllowPrinting.Enabled = True
50450    chkAllowCopy.Enabled = True
50460    chkAllowModifyAnnotations.Enabled = True
50470    chkAllowModifyContents.Enabled = True
50480
50490    If optEncHigh.Value = True Then
50500      dmFraPDFHighPermissions.Enabled = True
50510      chkAllowDegradedPrinting.Enabled = True
50520      chkAllowFillIn.Enabled = True
50530      chkAllowScreenReaders.Enabled = True
50540      chkAllowAssembly.Enabled = True
50550     Else
50560      dmFraPDFHighPermissions.Enabled = False
50570      chkAllowDegradedPrinting.Enabled = False
50580      chkAllowFillIn.Enabled = False
50590      chkAllowScreenReaders.Enabled = False
50600      chkAllowAssembly.Enabled = False
50610    End If
50620  End If
50630  If chkOwnerPass.Value = 0 And chkUserPass.Value = 0 Then
50640   chkOwnerPass.Value = 1: Options.PDFOwnerPass = 1
50650  End If
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
50090    cmdFilenameSubstMove(0).Enabled = True
50100    cmdFilenameSubstMove(1).Enabled = True
50110   Else
50120    cmdFilenameSubstMove(0).Enabled = False
50130    cmdFilenameSubstMove(1).Enabled = False
50140  End If
50150  If lsvFilenameSubst.ListItems.Count > 0 Then
50160   If lsvFilenameSubst.SelectedItem.Index = 1 Then
50170    cmdFilenameSubstMove(0).Enabled = False
50180   End If
50190   If lsvFilenameSubst.SelectedItem.Index = lsvFilenameSubst.ListItems.Count Then
50200    cmdFilenameSubstMove(1).Enabled = False
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
50020  With txtSavePreview
50030   .Text = GetSubstFilename("C:\test.pdf", txtSaveFilename.Text, , True) & ".pdf"
50040   .ToolTipText = .Text
50050  End With
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

Private Function GetSaveAutosaveFormatExtension(Index As Long) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case Index
        Case -1, 0
50030    GetSaveAutosaveFormatExtension = ".pdf"
50040   Case 1
50050    GetSaveAutosaveFormatExtension = ".png"
50060   Case 2
50070    GetSaveAutosaveFormatExtension = ".jpg"
50080   Case 3
50090    GetSaveAutosaveFormatExtension = ".bmp"
50100   Case 4
50110    GetSaveAutosaveFormatExtension = ".pcx"
50120   Case 5
50130    GetSaveAutosaveFormatExtension = ".tif"
50140   Case 6
50150    GetSaveAutosaveFormatExtension = ".ps"
50160   Case 7
50170    GetSaveAutosaveFormatExtension = ".eps"
50180  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "GetSaveAutosaveFormatExtension")
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

Private Sub txtStandardAuthor_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txtStandardAuthor.ToolTipText = txtStandardAuthor.Text
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "txtStandardAuthor_Change")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtTemppath_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txtTemppath.ToolTipText = txtTemppath.Text
50020  With txtTemppathPreview
50030   .Text = ResolveEnvironment(GetSubstFilename2(txtTemppath.Text))
50040   .ToolTipText = .Text
50050  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "txtTemppath_Change")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetFrames()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ctl As Control
50020  For Each ctl In Controls
50030   If TypeOf ctl Is dmFrame Then
50040    ctl.Font.Size = 10
50050    ctl.TextShaddowColor = &HC00000
50060    If ComputerScreenResolution <= 8 Or Options.OptionsDesign = 1 Then
50070      ctl.UseGradient = False: ctl.Caption3D = [Flat Caption]
50080      If UCase$(ctl.Name) = "DMFRADESCRIPTION" Then
50090        ctl.BarColorFrom = vbRed
50100       Else
50110        ctl.BarColorFrom = vbBlue
50120      End If
50130     Else
50140      ctl.UseGradient = True: ctl.Caption3D = [Raised Caption]
50150      If UCase$(ctl.Name) = "DMFRADESCRIPTION" Then
50160        ctl.BarColorFrom = &H8080FF
50170        ctl.BarColorTo = &HC0&
50180       Else
50190        ctl.BarColorFrom = &HFF8080
50200        ctl.BarColorTo = &H400000
50210      End If
50220    End If
50230   End If
50240  Next ctl
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "SetFrames")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
