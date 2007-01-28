VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Options"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9165
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin PDFCreator.dmFrame dmFraProgLanguage 
      Height          =   5895
      Left            =   2640
      TabIndex        =   244
      Top             =   1320
      Visible         =   0   'False
      Width           =   6375
      _extentx        =   11245
      _extenty        =   10398
      caption         =   "Language"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":000C
      Begin VB.CommandButton cmdLanguageRemove 
         Caption         =   "Remove"
         Height          =   315
         Left            =   3990
         TabIndex        =   251
         Top             =   630
         Width           =   1575
      End
      Begin VB.ComboBox cmbCurrentLanguage 
         Height          =   315
         Left            =   105
         Style           =   2  'Dropdown-Liste
         TabIndex        =   249
         Top             =   630
         Width           =   3795
      End
      Begin VB.CommandButton cmdLanguageInstall 
         Caption         =   "Install"
         Height          =   375
         Left            =   4080
         TabIndex        =   247
         Top             =   2025
         Width           =   1575
      End
      Begin VB.CommandButton cmdLanguageRefresh 
         Caption         =   "Refresh List"
         Height          =   375
         Left            =   4080
         TabIndex        =   246
         Top             =   1545
         Width           =   1575
      End
      Begin MSComctlLib.ListView lsvTranslations 
         Height          =   4230
         Left            =   120
         TabIndex        =   245
         Top             =   1545
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   7461
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label lblLanguagesFromInternet 
         AutoSize        =   -1  'True
         Caption         =   "Load more languages from the internet"
         Height          =   195
         Left            =   105
         TabIndex        =   250
         Top             =   1260
         Width           =   2715
      End
      Begin VB.Label lblCurrentLanguage 
         AutoSize        =   -1  'True
         Caption         =   "Current language"
         Height          =   195
         Left            =   105
         TabIndex        =   248
         Top             =   420
         Width           =   1215
      End
   End
   Begin PDFCreator.dmFrame dmFraProgDocument2 
      Height          =   2610
      Left            =   2640
      TabIndex        =   234
      Top             =   3720
      Visible         =   0   'False
      Width           =   6195
      _extentx        =   10927
      _extenty        =   4604
      caption         =   "Document 2"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":0038
      Begin VB.TextBox txtCustomPapersizeHeight 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   1920
         TabIndex        =   240
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtCustomPapersizeWidth 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   640
         TabIndex        =   239
         Top             =   1800
         Width           =   975
      End
      Begin VB.CheckBox chkUseCustomPapersize 
         Appearance      =   0  '2D
         Caption         =   "Use a custom papersize"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   238
         Top             =   1200
         Width           =   5760
      End
      Begin VB.ComboBox cmbDocumentPapersizes 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown-Liste
         TabIndex        =   237
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox chkUseFixPaperSize 
         Appearance      =   0  '2D
         Caption         =   "Use a fix papersize"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   235
         Top             =   360
         Width           =   6000
      End
      Begin VB.Label lblCustomPapersizeInfo 
         AutoSize        =   -1  'True
         Caption         =   "Units of 1/72 of an inch."
         Height          =   195
         Left            =   640
         TabIndex        =   243
         Top             =   2280
         Width           =   1725
      End
      Begin VB.Label lblCustomPapersizeHeight 
         AutoSize        =   -1  'True
         Caption         =   "Height"
         Height          =   195
         Left            =   1920
         TabIndex        =   242
         Top             =   1560
         Width           =   465
      End
      Begin VB.Label lblCustomPapersizeWidth 
         AutoSize        =   -1  'True
         Caption         =   "Width"
         Height          =   195
         Left            =   640
         TabIndex        =   241
         Top             =   1560
         Width           =   420
      End
   End
   Begin PDFCreator.dmFrame dmFraProgDocument1 
      Height          =   2250
      Left            =   2640
      TabIndex        =   45
      Top             =   1800
      Visible         =   0   'False
      Width           =   6195
      _extentx        =   10927
      _extenty        =   3969
      caption         =   "Document 1"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":0064
      Begin VB.CheckBox chkOnePagePerFile 
         Appearance      =   0  '2D
         Caption         =   "One page per file"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   223
         Top             =   1890
         Width           =   5985
      End
      Begin VB.ComboBox cmbAuthorTokens 
         Appearance      =   0  '2D
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmOptions.frx":0090
         Left            =   3720
         List            =   "frmOptions.frx":0092
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
         Width           =   5985
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
         Width           =   5985
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
   Begin PDFCreator.dmFrame dmFraProgStamp 
      Height          =   2610
      Left            =   2640
      TabIndex        =   224
      Top             =   3360
      Visible         =   0   'False
      Width           =   6195
      _extentx        =   10927
      _extenty        =   4604
      caption         =   "Stamp"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":0094
      Begin VB.TextBox txtOutlineFontThickness 
         Alignment       =   1  'Rechts
         Appearance      =   0  '2D
         Height          =   255
         Left            =   2040
         TabIndex        =   232
         Text            =   "0"
         Top             =   2040
         Width           =   375
      End
      Begin VB.CommandButton cmdStampFont 
         Caption         =   "..."
         Height          =   315
         Left            =   3720
         TabIndex        =   230
         Top             =   720
         Width           =   495
      End
      Begin VB.CheckBox chkStampUseOutlineFont 
         Appearance      =   0  '2D
         Caption         =   "Use outline font"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   229
         Top             =   1560
         Width           =   5895
      End
      Begin VB.PictureBox picStampFontColor 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4800
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   228
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox txtStampString 
         Appearance      =   0  '2D
         Height          =   315
         Left            =   120
         TabIndex        =   226
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label lblFontNameSize 
         AutoSize        =   -1  'True
         Caption         =   "Arial, 12"
         Height          =   195
         Left            =   120
         TabIndex        =   233
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label lblOutlineFontThickness 
         AutoSize        =   -1  'True
         Caption         =   "Outline font thickness"
         Height          =   195
         Left            =   390
         TabIndex        =   231
         Top             =   2040
         Width           =   1530
      End
      Begin VB.Label lblStampFontcolor 
         AutoSize        =   -1  'True
         Caption         =   "Font-color"
         Height          =   195
         Left            =   4800
         TabIndex        =   227
         Top             =   480
         Width           =   705
      End
      Begin VB.Label lblStampString 
         AutoSize        =   -1  'True
         Caption         =   "Stampstring"
         Height          =   195
         Left            =   120
         TabIndex        =   225
         Top             =   480
         Width           =   825
      End
   End
   Begin PDFCreator.dmFrame dmFraProgActions 
      Height          =   4560
      Left            =   2625
      TabIndex        =   178
      Top             =   1050
      Visible         =   0   'False
      Width           =   6375
      _extentx        =   11245
      _extenty        =   8043
      caption         =   "Actions"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":00C0
      Begin PDFCreator.dmFrame dmFraProgActionsRunProgramAfterSaving 
         Height          =   3510
         Left            =   1785
         TabIndex        =   188
         Top             =   735
         Width           =   6165
         _extentx        =   10874
         _extenty        =   6191
         caption         =   "Run a program/script after saving"
         barcolorfrom    =   16744576
         barcolorto      =   4194304
         font            =   "frmOptions.frx":00EC
         Begin VB.CommandButton cmdRunProgramAfterSavingPrognameEdit 
            Height          =   300
            Left            =   5520
            Picture         =   "frmOptions.frx":0118
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
         _extentx        =   10874
         _extenty        =   6191
         caption         =   "Run a program/script before saving"
         barcolorfrom    =   16744576
         barcolorto      =   4194304
         font            =   "frmOptions.frx":06A2
         Begin VB.CommandButton cmdRunProgramBeforeSavingPrognameEdit 
            Height          =   300
            Left            =   5520
            Picture         =   "frmOptions.frx":06CE
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
      _extentx        =   11245
      _extenty        =   4710
      caption         =   "Save"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":0C58
      Begin VB.ComboBox cmbStandardSaveFormat 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "frmOptions.frx":0C84
         Left            =   120
         List            =   "frmOptions.frx":0C86
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
         ItemData        =   "frmOptions.frx":0C88
         Left            =   3720
         List            =   "frmOptions.frx":0C8A
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
      _extentx        =   11245
      _extenty        =   8969
      caption         =   "Autosave"
      caption3d       =   2
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":0C8C
      textshaddowcolor=   12582912
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
         ItemData        =   "frmOptions.frx":0CB8
         Left            =   3690
         List            =   "frmOptions.frx":0CBA
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
   Begin PDFCreator.dmFrame dmFraPDFGeneral 
      Height          =   4365
      Left            =   2625
      TabIndex        =   90
      Top             =   1890
      Visible         =   0   'False
      Width           =   6195
      _extentx        =   10927
      _extenty        =   7699
      caption         =   "General Options"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":0CBC
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
         ItemData        =   "frmOptions.frx":0CE8
         Left            =   120
         List            =   "frmOptions.frx":0CEA
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
         ItemData        =   "frmOptions.frx":0CEC
         Left            =   120
         List            =   "frmOptions.frx":0CEE
         Style           =   2  'Dropdown-Liste
         TabIndex        =   92
         Top             =   735
         Width           =   2655
      End
      Begin VB.ComboBox cmbPDFRotate 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "frmOptions.frx":0CF0
         Left            =   120
         List            =   "frmOptions.frx":0CF2
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
      _extentx        =   10927
      _extenty        =   7250
      caption         =   "General 1"
      caption3d       =   2
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":0CF4
      textshaddowcolor=   12582912
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
         _extentx        =   10451
         _extenty        =   53
         linetype        =   1
         3dhighlight     =   -2147483628
         3dshadow        =   -2147483632
         drawstyle       =   0
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
         _extentx        =   10451
         _extenty        =   53
         linetype        =   1
         3dhighlight     =   -2147483628
         3dshadow        =   -2147483632
         drawstyle       =   0
      End
      Begin PDFCreator.Line3D Line3D1 
         Height          =   30
         Index           =   2
         Left            =   120
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   3240
         Width           =   5925
         _extentx        =   10451
         _extenty        =   53
         linetype        =   1
         3dhighlight     =   -2147483628
         3dshadow        =   -2147483632
         drawstyle       =   0
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
      _extentx        =   10927
      _extenty        =   1879
      caption         =   "Shell integration"
      caption3d       =   2
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":0D20
      textshaddowcolor=   12582912
      enabled         =   0   'False
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
      _extentx        =   10927
      _extenty        =   4842
      caption         =   "General 2"
      caption3d       =   2
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":0D4C
      textshaddowcolor=   12582912
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
         ItemData        =   "frmOptions.frx":0D78
         Left            =   120
         List            =   "frmOptions.frx":0D7A
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
         _extentx        =   10451
         _extenty        =   53
         linetype        =   1
         3dhighlight     =   -2147483628
         3dshadow        =   -2147483632
         drawstyle       =   0
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
      _extentx        =   11245
      _extenty        =   6932
      caption         =   "Print after saving"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":0D7C
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
      _extentx        =   11324
      _extenty        =   5556
      caption         =   "Ghostscript"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":0DA8
      textshaddowcolor=   12582912
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
      _extentx        =   11456
      _extenty        =   2487
      caption         =   "Directories"
      caption3d       =   2
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":0DD4
      textshaddowcolor=   12582912
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
         Picture         =   "frmOptions.frx":0E00
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
      _extentx        =   11245
      _extenty        =   1931
      caption         =   "Postscript"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":138A
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
      _extentx        =   11245
      _extenty        =   3413
      caption         =   "Bitmap"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":13B6
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
      _extentx        =   10927
      _extenty        =   2990
      caption         =   "Font options"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":13E2
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
      _extentx        =   11245
      _extenty        =   8281
      caption         =   "Programfont"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":140E
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
      _extentx        =   10927
      _extenty        =   9763
      caption         =   "Security"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":143A
      Begin PDFCreator.dmFrame dmFraPDFHighPermissions 
         Height          =   855
         Left            =   120
         TabIndex        =   151
         Top             =   4560
         Width           =   5955
         _extentx        =   10504
         _extenty        =   1508
         caption         =   "Enhanced permissions (128 Bit only)"
         barcolorfrom    =   16744576
         barcolorto      =   4194304
         font            =   "frmOptions.frx":1466
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
         _extentx        =   10504
         _extenty        =   1508
         caption         =   "Disallow user to"
         barcolorfrom    =   16744576
         barcolorto      =   4194304
         font            =   "frmOptions.frx":1492
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
         _extentx        =   10504
         _extenty        =   1508
         caption         =   "Passwords"
         barcolorfrom    =   16744576
         barcolorto      =   4194304
         font            =   "frmOptions.frx":14BE
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
         _extentx        =   10504
         _extenty        =   1508
         caption         =   "Encryption level"
         barcolorfrom    =   16744576
         barcolorto      =   4194304
         font            =   "frmOptions.frx":14EA
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
         _extentx        =   10504
         _extenty        =   1508
         caption         =   "Encryptor"
         barcolorfrom    =   16744576
         barcolorto      =   4194304
         font            =   "frmOptions.frx":1516
         Begin VB.ComboBox cmbPDFEncryptor 
            Appearance      =   0  '2D
            Height          =   315
            ItemData        =   "frmOptions.frx":1542
            Left            =   120
            List            =   "frmOptions.frx":1544
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
      _extentx        =   10927
      _extenty        =   7646
      caption         =   "Compression"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":1546
      Begin PDFCreator.dmFrame dmFraPDFMono 
         Height          =   1095
         Left            =   120
         TabIndex        =   117
         Top             =   3120
         Width           =   5955
         _extentx        =   10504
         _extenty        =   1931
         caption         =   "Monochrome images"
         barcolorfrom    =   16744576
         barcolorto      =   4194304
         font            =   "frmOptions.frx":1572
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
            ItemData        =   "frmOptions.frx":159E
            Left            =   120
            List            =   "frmOptions.frx":15A0
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
            ItemData        =   "frmOptions.frx":15A2
            Left            =   2520
            List            =   "frmOptions.frx":15A4
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
         _extentx        =   10504
         _extenty        =   1931
         caption         =   "Greyscale images"
         barcolorfrom    =   16744576
         barcolorto      =   4194304
         font            =   "frmOptions.frx":15A6
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
            ItemData        =   "frmOptions.frx":15D2
            Left            =   2520
            List            =   "frmOptions.frx":15D4
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
            ItemData        =   "frmOptions.frx":15D6
            Left            =   120
            List            =   "frmOptions.frx":15D8
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
         _extentx        =   10504
         _extenty        =   1931
         caption         =   "Color images"
         barcolorfrom    =   16744576
         barcolorto      =   4194304
         font            =   "frmOptions.frx":15DA
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
            ItemData        =   "frmOptions.frx":1606
            Left            =   2520
            List            =   "frmOptions.frx":1608
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
            ItemData        =   "frmOptions.frx":160A
            Left            =   120
            List            =   "frmOptions.frx":160C
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
      _extentx        =   10927
      _extenty        =   2143
      caption         =   "Color options"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":160E
      Begin VB.ComboBox cmbPDFColorModel 
         Appearance      =   0  '2D
         Height          =   315
         ItemData        =   "frmOptions.frx":163A
         Left            =   120
         List            =   "frmOptions.frx":163C
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
      _extentx        =   10927
      _extenty        =   2566
      caption         =   "Options"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":163E
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
      Height          =   7935
      Left            =   0
      TabIndex        =   159
      Top             =   0
      Width           =   2535
      _extentx        =   4471
      _extenty        =   13996
      fontname        =   "MS Sans Serif"
      fontcharset     =   0
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   840
         Top             =   7245
      End
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
            NumListImages   =   20
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":166A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":1C04
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":219E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":2738
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":2CD2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":306C
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":3606
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":3EE0
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":447A
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":4A14
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":4FAE
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":5548
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":5AE2
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":607C
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":6616
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":6BB0
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":714A
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":76E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":7C7E
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOptions.frx":8558
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
      _extentx        =   11245
      _extenty        =   4471
      caption         =   "Filename substitutions"
      barcolorfrom    =   16744576
      barcolorto      =   4194304
      font            =   "frmOptions.frx":8E32
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
         Picture         =   "frmOptions.frx":8E5E
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
         Picture         =   "frmOptions.frx":91E8
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
      _extentx        =   11324
      _extenty        =   1879
      caption         =   ""
      barcolorfrom    =   8421631
      barcolorto      =   192
      font            =   "frmOptions.frx":9572
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
   Begin MSComctlLib.TabStrip tbstrProgDocument 
      Height          =   4935
      Left            =   0
      TabIndex        =   236
      Top             =   0
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

Private WithEvents dl As clsDownload
Attribute dl.VB_VarHelpID = -1

Private UnloadForm As Boolean, TimerReady As Boolean, _
 Languages As Collection, LangFiles As Collection, oldLanguage As String

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

Private Sub chkStampUseOutlineFont_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkStampUseOutlineFont.Value = 1 Then
50020    lblOutlineFontThickness.Enabled = True
50030    txtOutlineFontThickness.Enabled = True
50040    txtOutlineFontThickness.BackColor = &H80000005
50050   Else
50060    lblOutlineFontThickness.Enabled = False
50070    txtOutlineFontThickness.Enabled = False
50080    txtOutlineFontThickness.BackColor = &H8000000F
50090  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "chkStampUseOutlineFont_Click")
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

Private Sub chkUseCustomPapersize_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkUseCustomPapersize.Value = 1 Then
50020    lblCustomPapersizeWidth.Enabled = True
50030    lblCustomPapersizeHeight.Enabled = True
50040    txtCustomPapersizeWidth.Enabled = True
50050    txtCustomPapersizeWidth.BackColor = &H80000005
50060    txtCustomPapersizeHeight.Enabled = True
50070    txtCustomPapersizeHeight.BackColor = &H80000005
50080    lblCustomPapersizeInfo.Enabled = True
50090    cmbDocumentPapersizes.Enabled = False
50100   Else
50110    cmbDocumentPapersizes.Enabled = True
50120    lblCustomPapersizeWidth.Enabled = False
50130    lblCustomPapersizeHeight.Enabled = False
50140    txtCustomPapersizeWidth.Enabled = False
50150    txtCustomPapersizeWidth.BackColor = &H8000000F
50160    txtCustomPapersizeHeight.Enabled = False
50170    txtCustomPapersizeHeight.BackColor = &H8000000F
50180    lblCustomPapersizeInfo.Enabled = False
50190  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "chkUseCustomPapersize_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub chkUseFixPaperSize_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If chkUseFixPaperSize.Value = 1 Then
50020    cmbDocumentPapersizes.Enabled = True
50030    chkUseCustomPapersize.Enabled = True
50040    If chkUseCustomPapersize.Value = 1 Then
50050      lblCustomPapersizeWidth.Enabled = True
50060      lblCustomPapersizeHeight.Enabled = True
50070      txtCustomPapersizeWidth.Enabled = True
50080      txtCustomPapersizeWidth.BackColor = &H80000005
50090      txtCustomPapersizeHeight.Enabled = True
50100      txtCustomPapersizeHeight.BackColor = &H80000005
50110      lblCustomPapersizeInfo.Enabled = True
50120      cmbDocumentPapersizes.Enabled = False
50130     Else
50140      cmbDocumentPapersizes.Enabled = True
50150      lblCustomPapersizeWidth.Enabled = False
50160      lblCustomPapersizeHeight.Enabled = False
50170      txtCustomPapersizeWidth.Enabled = False
50180      txtCustomPapersizeWidth.BackColor = &H8000000F
50190      txtCustomPapersizeHeight.Enabled = False
50200      txtCustomPapersizeHeight.BackColor = &H8000000F
50210      lblCustomPapersizeInfo.Enabled = False
50220    End If
50230   Else
50240    cmbDocumentPapersizes.Enabled = False
50250    chkUseCustomPapersize.Enabled = False
50260    lblCustomPapersizeWidth.Enabled = False
50270    lblCustomPapersizeHeight.Enabled = False
50280    txtCustomPapersizeWidth.Enabled = False
50290    txtCustomPapersizeWidth.BackColor = &H8000000F
50300    txtCustomPapersizeHeight.Enabled = False
50310    txtCustomPapersizeHeight.BackColor = &H8000000F
50320    lblCustomPapersizeInfo.Enabled = False
50330  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "chkUseFixPaperSize_Click")
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
50010  KeyAscii = AllowedKeypressChars(KeyAscii)
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

Private Sub cmbCurrentLanguage_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim f As Form
50020  If InStr(1, LangFiles(cmbCurrentLanguage.ListIndex + 1), LanguagePath, vbTextCompare) = 1 Then
50030    cmdLanguageRemove.Enabled = False
50040   Else
50050    cmdLanguageRemove.Enabled = True
50060  End If
50070  SetLanguage Languages(cmbCurrentLanguage.ListIndex + 1)
50080  LoadLanguage LangFiles(cmbCurrentLanguage.ListIndex + 1)
50090  For Each f In Forms
50100   f.ChangeLanguage
50110  Next
50120  lblOptions.Caption = LanguageStrings.OptionsProgramLanguagesDescription
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmbCurrentLanguage_Click")
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
50010  Dim f As Form, LanguagePath As String
50020
50030  SetLanguage oldLanguage
50040
50050  LanguagePath = CompletePath(GetPDFCreatorApplicationPath) & "Languages\"
50060  LoadLanguage LanguagePath & oldLanguage & ".ini"
50070
50080  For Each f In Forms
50090   f.ChangeLanguage
50100  Next
50110  Unload Me
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

Private Sub cmdLanguageInstall_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim strInstallPath As String
50020  Const strDownloadPath = "http://www.pdfforge.org/files/translations/"
50030  strInstallPath = CompletePath(GetMyAppData()) & "PDFCreator\Languages"
50040  If Not DirExists(CompletePath(GetMyAppData()) & "PDFCreator") Then
50050   CreateDir CompletePath(GetMyAppData()) & "PDFCreator"
50060  End If
50070
50080  If Not DirExists(GetMyAppData() & "\PDFCreator\Languages") Then
50090   CreateDir GetMyAppData() & "\PDFCreator\Languages"
50100  End If
50110  If lsvTranslations.SelectedItem Is Nothing Then
50120   Exit Sub
50130  End If
50140  InstallInternetLanguageFile lsvTranslations.SelectedItem.Text, lsvTranslations.SelectedItem.SubItems(1), strDownloadPath, strInstallPath
50150  ReadAllLanguages LanguagePath, True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdLanguageInstall_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdLanguageRefresh_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim strLanguages() As String, strFile() As String, i As Long
50020  Const strDownloadURL = "http://www.pdfforge.org/products/pdfcreator/translations/list"
50030  MousePointer = vbHourglass
50040  Set dl = New clsDownload
50050  strLanguages = Split(dl.DownloadString(strDownloadURL), vbLf)
50060  Set dl = Nothing
50070  lsvTranslations.ListItems.Clear
50080  For i = LBound(strLanguages) To UBound(strLanguages)
50090   If ((strLanguages(i) <> vbNullString) And (InStr(1, strLanguages(i), ":"))) Then
50100    strFile = Split(strLanguages(i), ":")
50110    lsvTranslations.ListItems.Add(, , strFile(0)).SubItems(1) = strFile(1)
50120   End If
50130  Next i
50140  MousePointer = vbDefault
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdLanguageRefresh_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub cmdLanguageRemove_Click()
 On Error GoTo ErrorHandler
 Kill LangFiles(cmbCurrentLanguage.ListIndex + 1)
 If StrComp(LangFiles(cmbCurrentLanguage.ListIndex + 1), oldLanguage, vbTextCompare) <> 0 Then
  Options.Language = "english"
 End If
 ReadAllLanguages LanguagePath, True
 Exit Sub
ErrorHandler:
 MsgBox Err.Description
End Sub

Private Sub cmdReset_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim res As Long
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
50010  Dim tRestart As Boolean, tOpt As tOptions, newLanguage As String
50020  tRestart = False
50030  If UCase$(Options.DirectoryGhostscriptBinaries) <> UCase$(txtGSbin.Text) Then
50040   tRestart = True
50050  End If
50060  CorrectCmbCharset
50070
50080  tOpt = Options
50090  newLanguage = Options.Language
50100  GetOptions Me, Options
50110  Options.Language = newLanguage
50120  Options.StampFontname = tOpt.StampFontname
50130  Options.StampFontsize = tOpt.StampFontsize
50140  SaveOptions Options
50150
50160  If IsWin9xMe = False Then
50171   Select Case Options.ProcessPriority
         Case 0: 'Idle
50190     SetProcessPriority Idle
50200    Case 1: 'Normal
50210     SetProcessPriority Normal
50220    Case 2: 'High
50230     SetProcessPriority High
50240    Case 3: 'Realtime
50250     SetProcessPriority RealTime
50260   End Select
50270  End If
50280  If tRestart = True Then
50290   Restart = True
50300  End If
50310  Unload Me
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

Private Sub cmdStampFont_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim res As Long, Font As tFont
50020  Font.Name = Options.StampFontname
50030  Font.Size = Options.StampFontsize
50040  If OpenFontDialog(Font, Me.hwnd) > 0 Then
50050   Options.StampFontname = Font.Name
50060   Options.StampFontsize = Font.Size
50070   lblFontNameSize.Caption = Font.Name & ", " & Font.Size
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "cmdStampFont_Click")
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
50060  UnloadForm = False
50070  TimerReady = True
50080  Me.Icon = LoadResPicture(2120, vbResIcon)
50090  KeyPreview = True
50100
50110  oldLanguage = Options.Language
50120
50130  With Screen
50140   .MousePointer = vbHourglass
50150   Move (.Width - Width) / 2, (.Height - Height) / 2
50160  End With
50170
50180  SetFrames
50190
50200  With dmFraDescription
50210   .Caption = LanguageStrings.OptionsTreeProgram
50220   .Visible = True
50230  End With
50240  tbstrProgGeneral.Visible = True
50250  With dmFraProgGeneral1
50260   .Visible = True
50270   .Top = dmFraDescription.Top + dmFraDescription.Height + 50
50280   .Left = dmFraDescription.Left
50290   dmFraProgGeneral2.Top = .Top
50300   dmFraProgGeneral2.Left = .Left
50310   dmFraProgGeneral2.Width = .Width
50320   dmFraShellIntegration.Width = dmFraProgGeneral2.Width
50330   dmFraProgGhostscript.Top = .Top
50340   dmFraProgGhostscript.Left = .Left
50350   dmFraProgGhostscript.Width = dmFraDescription.Width
50360   dmFraProgAutosave.Top = .Top
50370   dmFraProgAutosave.Left = .Left
50380   dmFraProgAutosave.Width = dmFraDescription.Width
50390   dmFraProgDirectories.Top = .Top
50400   dmFraProgDirectories.Left = .Left
50410   dmFraProgDirectories.Width = dmFraDescription.Width
50420   dmFraProgDocument1.Top = .Top
50430   dmFraProgDocument1.Left = .Left
50440   dmFraProgStamp.Top = dmFraProgDocument1.Top + dmFraProgDocument1.Height + 50
50450   dmFraProgStamp.Left = .Left
50460   dmFraProgDocument2.Top = .Top
50470   dmFraProgDocument2.Left = .Left
50480   dmFraProgSave.Top = .Top
50490   dmFraProgSave.Left = .Left
50500   dmFraProgSave.Width = dmFraDescription.Width
50510   dmFraFilenameSubstitutions.Top = dmFraProgSave.Top + dmFraProgSave.Height + 50
50520   dmFraFilenameSubstitutions.Left = .Left
50530   dmFraFilenameSubstitutions.Width = dmFraDescription.Width
50540   dmFraProgFont.Top = .Top
50550   dmFraProgFont.Left = .Left
50560   dmFraProgFont.Width = dmFraDescription.Width
50570   dmFraProgLanguage.Top = .Top
50580   dmFraProgLanguage.Left = .Left
50590   dmFraProgLanguage.Width = dmFraDescription.Width
50600   dmFraProgActions.Top = .Top
50610   dmFraProgActions.Left = .Left
50620   dmFraProgActions.Width = dmFraDescription.Width
50630   dmFraProgPrint.Top = .Top
50640   dmFraProgPrint.Left = .Left
50650   dmFraProgPrint.Width = dmFraDescription.Width
50660   dmFraBitmapGeneral.Top = .Top
50670   dmFraBitmapGeneral.Left = .Left
50680   dmFraBitmapGeneral.Width = dmFraDescription.Width
50690   dmFraPSGeneral.Top = .Top
50700   dmFraPSGeneral.Left = .Left
50710   dmFraPSGeneral.Width = dmFraDescription.Width
50720
50730   dmFraProgActionsRunProgramAfterSaving.Top = dmFraProgActionsRunProgramBeforeSaving.Top
50740   dmFraProgActionsRunProgramAfterSaving.Left = dmFraProgActionsRunProgramBeforeSaving.Left
50750
50760   cmdCancel.Left = .Left
50770   cmdReset.Left = .Left + (dmFraDescription.Width - cmdReset.Width) / 2
50780   cmdSave.Left = .Left + dmFraDescription.Width - cmdSave.Width
50790  End With
50800
50810  With tbstrProgGeneral
50820   .Top = dmFraDescription.Top + dmFraDescription.Height + 50
50830   .Left = dmFraDescription.Left
50840   .Height = cmdCancel.Top - tbstrProgGeneral.Top - 50
50850   .Width = dmFraDescription.Width
50860  End With
50870
50880  With dmFraProgGeneral1
50890   .Top = tbstrProgGeneral.ClientTop + 100
50900   .Left = tbstrProgGeneral.Left + (tbstrProgGeneral.Width - .Width) / 2
50910  End With
50920  With dmFraProgGeneral2
50930   .Top = tbstrProgGeneral.ClientTop + 100
50940   .Left = tbstrProgGeneral.Left + (tbstrProgGeneral.Width - .Width) / 2
50950  End With
50960  With dmFraShellIntegration
50970   .Top = dmFraProgGeneral2.Top + dmFraProgGeneral2.Height + 50
50980   .Left = tbstrProgGeneral.Left + (tbstrProgGeneral.Width - .Width) / 2
50990  End With
51000
51010  With tbstrProgDocument
51020   .Top = dmFraDescription.Top + dmFraDescription.Height + 50
51030   .Left = dmFraDescription.Left
51040   .Height = cmdCancel.Top - tbstrProgDocument.Top - 50
51050   .Width = dmFraDescription.Width
51060  End With
51070  With dmFraProgDocument1
51080   .Top = tbstrProgDocument.ClientTop + 100
51090   .Left = tbstrProgDocument.Left + (tbstrProgDocument.Width - .Width) / 2
51100  End With
51110  With dmFraProgStamp
51120   .Top = dmFraProgDocument1.Top + dmFraProgDocument1.Height + 50
51130   .Left = tbstrProgDocument.Left + (tbstrProgDocument.Width - .Width) / 2
51140  End With
51150
51160  With dmFraProgDocument2
51170   .Top = dmFraProgDocument1.Top
51180   .Left = dmFraProgDocument1.Left
51190  End With
51200
51210  With tbstrPDFOptions
51220   .Top = dmFraDescription.Top + dmFraDescription.Height + 50
51230   .Left = dmFraDescription.Left
51240   .Height = cmdCancel.Top - tbstrPDFOptions.Top - 50
51250   .Width = dmFraDescription.Width
51260  End With
51270
51280  With dmFraPDFGeneral
51290   .Top = tbstrPDFOptions.ClientTop + 100
51300   .Left = tbstrPDFOptions.Left + (tbstrPDFOptions.Width - .Width) / 2
51310   dmfraPDFCompress.Top = .Top
51320   dmfraPDFCompress.Left = .Left
51330   dmFraPDFFonts.Top = .Top
51340   dmFraPDFFonts.Left = .Left
51350   dmFraPDFColors.Top = .Top
51360   dmFraPDFColors.Left = .Left
51370   dmFraPDFColorOptions.Top = dmFraPDFColors.Top + dmFraPDFColors.Height + 50
51380   dmFraPDFColorOptions.Left = .Left
51390   dmFraPDFSecurity.Top = .Top
51400   dmFraPDFSecurity.Left = .Left
51410  End With
51420
51430  cmbEPSLanguageLevel.Top = cmbPSLanguageLevel.Top
51440  cmbEPSLanguageLevel.Left = cmbPSLanguageLevel.Left
51450
51460  ieb.DisableUpdates True
51470  ieb.ClearStructure
51480  ieb.SetImageList imlIeb
51490  With LanguageStrings
51500   ieb.AddGroup "Program", .OptionsTreeProgram, 0
51510   ieb.AddItem "Program", "General", .OptionsProgramGeneralSymbol, 1
51520   ieb.AddItem "Program", "Ghostscript", .OptionsProgramGhostscriptSymbol, 2
51530   ieb.AddItem "Program", "Document", .OptionsProgramDocumentSymbol, 3
51540   ieb.AddItem "Program", "Save", .OptionsProgramSaveSymbol, 4
51550   ieb.AddItem "Program", "AutoSave", .OptionsProgramAutosaveSymbol, 5
51560   ieb.AddItem "Program", "Directories", .OptionsProgramDirectoriesSymbol, 6
51570   ieb.AddItem "Program", "Actions", .OptionsProgramActionsSymbol, 7
51580   ieb.AddItem "Program", "Print", .OptionsProgramPrintSymbol, 8
51590   ieb.AddItem "Program", "Fonts", .OptionsProgramFontSymbol, 9
51600   ieb.AddItem "Program", "Language", .OptionsProgramLanguagesSymbol, 10
51610
51620   ieb.AddGroup "Formats", .OptionsTreeFormats, 0
51630   ieb.AddItem "Formats", "PDF", .OptionsPDFSymbol, 10
51640   ieb.AddItem "Formats", "PNG", .OptionsPNGSymbol, 11
51650   ieb.AddItem "Formats", "JPEG", .OptionsJPEGSymbol, 12
51660   ieb.AddItem "Formats", "BMP", .OptionsBMPSymbol, 13
51670   ieb.AddItem "Formats", "PCX", .OptionsPCXSymbol, 14
51680   ieb.AddItem "Formats", "TIFF", .OptionsTIFFSymbol, 15
51690   ieb.AddItem "Formats", "PS", .OptionsPSSymbol, 16
51700   ieb.AddItem "Formats", "EPS", .OptionsEPSSymbol, 17
51710   ieb.DisableUpdates False
51720
51730   Set picOptions = LoadResPicture(2101, vbResIcon)
51740   dmFraProgGeneral1.Visible = True
51750
51760   dmFraProgGeneral1.Caption = .OptionsProgramGeneralDescription1
51770   dmFraProgGeneral2.Caption = .OptionsProgramGeneralDescription2
51780   With tbstrProgGeneral.Tabs
51790    .Clear
51800    .Add , , LanguageStrings.OptionsProgramGeneralDescription1
51810    .Add , , LanguageStrings.OptionsProgramGeneralDescription2
51820   End With
51830   With tbstrProgDocument.Tabs
51840    .Clear
51850    .Add , , LanguageStrings.OptionsProgramDocumentDescription1
51860    .Add , , LanguageStrings.OptionsProgramDocumentDescription2
51870   End With
51880   dmFraShellIntegration.Caption = .OptionsShellIntegration
51890   dmFraProgGhostscript.Caption = .OptionsProgramGhostscriptSymbol
51900   dmFraProgAutosave.Caption = .OptionsProgramAutosaveSymbol
51910   dmFraProgDirectories.Caption = .OptionsProgramDirectoriesSymbol
51920   dmFraProgDocument1.Caption = .OptionsProgramDocumentDescription1
51930   dmFraProgDocument2.Caption = .OptionsProgramDocumentDescription2
51940   dmFraProgStamp.Caption = .OptionsStamp
51950   dmFraProgFont.Caption = .OptionsProgramFontSymbol
51960   dmFraProgSave.Caption = .OptionsProgramSaveSymbol
51970   dmFraProgActions.Caption = .OptionsProgramActionsSymbol
51980   dmFraProgPrint.Caption = .OptionsProgramPrintSymbol
51990   dmFraProgLanguage.Caption = .OptionsProgramLanguagesSymbol
52000
52010   cmdShellintegration(0).Caption = .OptionsShellIntegrationAdd
52020   cmdShellintegration(1).Caption = .OptionsShellIntegrationRemove
52030   If IsWin9xMe = False Then
52040    If IsAdmin = False Then
52050     cmdShellintegration(0).Enabled = False
52060     cmdShellintegration(1).Enabled = False
52070    End If
52080   End If
52090
52100   lblSendMailMethod.Caption = .OptionsSendMailMethod
52110   cmbSendMailMethod.AddItem .OptionsSendMailMethodAutomatic
52120   cmbSendMailMethod.AddItem .OptionsSendMailMethodMapi
52130   cmbSendMailMethod.AddItem .OptionsSendMailMethodSendmailDLL
52140
52150   cmdLanguageInstall.Caption = .OptionsLanguagesInstall
52160   cmdLanguageRefresh.Caption = .OptionsLanguagesRefresh
52170   lblLanguagesFromInternet.Caption = .OptionsLanguagesDownloadMoreLanguages
52180
52190   lblGhostscriptversion.Caption = .OptionsGhostscriptversion
52200   lblAdditionalGhostscriptParameters.Caption = .OptionsAdditionalGhostscriptParameters
52210   lblAdditionalGhostscriptSearchpath.Caption = .OptionsAdditionalGhostscriptSearchpath
52220   chkAddWindowsFontpath.Caption = .OptionsAddWindowsFontpath
52230
52240   lblSaveFilename.Caption = .OptionsSaveFilename
52250   lblSaveFilenameTokens.Caption = .OptionsSaveFilenameTokens
52260   dmFraFilenameSubstitutions.Caption = .OptionsSaveFilenameSubstitutions
52270   chkFilenameSubst.Caption = .OptionsSaveFilenameSubstitutionsTitle
52280   cmdFilenameSubst(0).Caption = .OptionsSaveFilenameAdd
52290   cmdFilenameSubst(1).Caption = .OptionsSaveFilenameChange
52300   cmdFilenameSubst(2).Caption = .OptionsSaveFilenameDelete
52310
52320   chkSpaces.Caption = .OptionsRemoveSpaces
52330   chkNoConfirmMessageSwitchingDefaultprinter.Caption = .OptionsProgramSwitchingDefaultprinter
52340   chkNoProcessingAtStartup.Caption = .OptionsProgramNoProcessingAtStartup
52350   lblOptionsDesign.Caption = .OptionsProgramOptionsDesign
52360   cmbOptionsDesign.AddItem .OptionsProgramOptionsDesignGradient
52370   cmbOptionsDesign.AddItem .OptionsProgramOptionsDesignSimple
52380   chkShowAnimation.Caption = .OptionsProgramShowAnimation
52390
52400   lblGSbin.Caption = .OptionsDirectoriesGSBin
52410   lblGSlib.Caption = .OptionsDirectoriesGSLibraries
52420   lblGSfonts.Caption = .OptionsDirectoriesGSFonts
52430   lblPrintTempPath.Caption = .OptionsDirectoriesTempPath
52440
52450   chkOnePagePerFile.Caption = .OptionsOnePagePerFile
52460   lblOptions = .OptionsProgramGeneralDescription
52470   lblAutosaveformat.Caption = .OptionsAutosaveFormat
52480   chkUseStandardAuthor.Caption = .OptionsUseStandardauthor
52490   chkUseCreationDateNow.Caption = .OptionsUseCreationDateNow
52500   chkUseAutosave.Caption = .OptionsUseAutosave
52510   cmdTestpage.Caption = .OptionsPrintTestpage
52520   lblAutosaveFilename.Caption = .OptionsAutosaveFilename
52530   lblAutosaveFilenameTokens.Caption = .OptionsAutosaveFilenameTokens
52540   chkUseAutosaveDirectory.Caption = .OptionsUseAutosaveDirectory
52550   lblAuthorTokens.Caption = .OptionsStandardAuthorToken
52560   chkAutosaveStartStandardProgram.Caption = .OptionsAutosaveStartStandardProgram
52570   chkAutosaveSendEmail.Caption = .OptionsSendEmailAfterAutosave
52580   lblStandardSaveformat.Caption = .OptionsStandardSaveFormat
52590
52600   dmFraProgActionsRunProgramAfterSaving.Caption = .OptionsProgramRunProgramAfterSavingCaption
52610   chkRunProgramAfterSaving.Caption = .OptionsProgramRunProgramAfterSavingCaption
52620   lblRunProgramAfterSavingProgramname.Caption = .OptionsProgramRunProgramAfterSavingProgram
52630   lblRunProgramAfterSavingProgramParameters.Caption = .OptionsProgramRunProgramAfterSavingProgramParameters
52640   chkRunProgramAfterSavingWaitUntilReady.Caption = .OptionsProgramRunProgramAfterSavingWaitUntilReady
52650   lblRunProgramAfterSavingWindowstyle.Caption = .OptionsProgramRunProgramAfterSavingWindowstyle
52660   cmbRunProgramAfterSavingWindowstyle.AddItem .OptionsProgramRunProgramAfterSavingWindowstyleHide
52670   cmbRunProgramAfterSavingWindowstyle.AddItem .OptionsProgramRunProgramAfterSavingWindowstyleNormalFocus
52680   cmbRunProgramAfterSavingWindowstyle.AddItem .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedFocus
52690   cmbRunProgramAfterSavingWindowstyle.AddItem .OptionsProgramRunProgramAfterSavingWindowstyleMaximizedFocus
52700   cmbRunProgramAfterSavingWindowstyle.AddItem .OptionsProgramRunProgramAfterSavingWindowstyleNormalNoFocus
52710   cmbRunProgramAfterSavingWindowstyle.AddItem .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedNoFocus
52720
52730   With tbstrProgActions.Tabs
52740    .Clear
52750    .Add , , LanguageStrings.OptionsProgramRunProgramBeforeSavingCaption
52760    .Add , , LanguageStrings.OptionsProgramRunProgramAfterSavingCaption
52770   End With
52780
52790   dmFraProgActionsRunProgramBeforeSaving.Caption = .OptionsProgramRunProgramBeforeSavingCaption
52800   chkRunProgramBeforeSaving.Caption = .OptionsProgramRunProgramBeforeSavingCaption
52810   lblRunProgramBeforeSavingProgramname.Caption = .OptionsProgramRunProgramBeforeSavingProgram
52820   lblRunProgramBeforeSavingProgramParameters.Caption = .OptionsProgramRunProgramBeforeSavingProgramParameters
52830   lblRunProgramBeforeSavingWindowstyle.Caption = .OptionsProgramRunProgramBeforeSavingWindowstyle
52840   cmbRunProgramBeforeSavingWindowstyle.AddItem .OptionsProgramRunProgramBeforeSavingWindowstyleHide
52850   cmbRunProgramBeforeSavingWindowstyle.AddItem .OptionsProgramRunProgramBeforeSavingWindowstyleNormalFocus
52860   cmbRunProgramBeforeSavingWindowstyle.AddItem .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedFocus
52870   cmbRunProgramBeforeSavingWindowstyle.AddItem .OptionsProgramRunProgramBeforeSavingWindowstyleMaximizedFocus
52880   cmbRunProgramBeforeSavingWindowstyle.AddItem .OptionsProgramRunProgramBeforeSavingWindowstyleNormalNoFocus
52890   cmbRunProgramBeforeSavingWindowstyle.AddItem .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedNoFocus
52900
52910   chkPrintAfterSaving.Caption = .OptionsPrintAfterSaving
52920   lblPrintAfterSavingPrinter.Caption = .OptionsPrintAfterSavingPrinter
52930
52940   For Each p In Printers
52950    cmbPrintAfterSavingPrinter.AddItem p.DeviceName
52960   Next p
52970
52980   lblPrintAfterSavingQueryUser.Caption = .OptionsPrintAfterSavingQueryUser
52990   cmbPrintAfterSavingQueryUser.AddItem .OptionsPrintAfterSavingQueryUserOff
53000   cmbPrintAfterSavingQueryUser.AddItem .OptionsPrintAfterSavingQueryUserStandardPrinterDialog
53010   cmbPrintAfterSavingQueryUser.AddItem .OptionsPrintAfterSavingQueryUserPrinterSetupDialog
53020   cmbPrintAfterSavingQueryUser.AddItem .OptionsPrintAfterSavingQueryUserDefaultPrinter
53030
53040   chkPrintAfterSavingNoCancel.Caption = .OptionsPrintAfterSavingNoCancel
53050   chkPrintAfterSavingDuplex.Caption = .OptionsPrintAfterSavingDuplex
53060   cmbPrintAfterSavingTumble.AddItem .OptionsPrintAfterSavingDuplexTumbleOff
53070   cmbPrintAfterSavingTumble.AddItem .OptionsPrintAfterSavingDuplexTumbleOn
53080
53090   With cmbStandardSaveFormat
53100    .AddItem "PDF"
53110    .AddItem "PNG"
53120    .AddItem "JPEG"
53130    .AddItem "BMP"
53140    .AddItem "PCX"
53150    .AddItem "TIFF"
53160    .AddItem "PS"
53170    .AddItem "EPS"
53180   End With
53190   With cmbAutosaveFormat
53200    .AddItem "PDF"
53210    .AddItem "PNG"
53220    .AddItem "JPEG"
53230    .AddItem "BMP"
53240    .AddItem "PCX"
53250    .AddItem "TIFF"
53260    .AddItem "PS"
53270    .AddItem "EPS"
53280   End With
53290   With cmbSaveFilenameTokens
53300    .AddItem "<Author>"
53310    .AddItem "<Computername>"
53320    .AddItem "<DateTime>"
53330    .AddItem "<Title>"
53340    .AddItem "<Username>"
53350    .AddItem "<REDMON_DOCNAME>"
53360    .AddItem "<REDMON_DOCNAME_FILE>"
53370    .AddItem "<REDMON_DOCNAME_PATH>"
53380    .AddItem "<REDMON_JOB>"
53390    .AddItem "<REDMON_MACHINE>"
53400    .AddItem "<REDMON_PORT>"
53410    .AddItem "<REDMON_PRINTER>"
53420    .AddItem "<REDMON_SESSIONID>"
53430    .AddItem "<REDMON_USER>"
53440    .ListIndex = 0
53450   End With
53460   With cmbAuthorTokens
53470    .AddItem "<Computername>"
53480    .AddItem "<ClientComputer>"
53490    .AddItem "<DateTime>"
53500    .AddItem "<Title>"
53510    .AddItem "<Username>"
53520    .AddItem "<REDMON_DOCNAME>"
53530    .AddItem "<REDMON_DOCNAME_FILE>"
53540    .AddItem "<REDMON_DOCNAME_PATH>"
53550    .AddItem "<REDMON_JOB>"
53560    .AddItem "<REDMON_MACHINE>"
53570    .AddItem "<REDMON_PORT>"
53580    .AddItem "<REDMON_PRINTER>"
53590    .AddItem "<REDMON_SESSIONID>"
53600    .AddItem "<REDMON_USER>"
53610    .ListIndex = 0
53620   End With
53630   With cmbAutoSaveFilenameTokens
53640    .AddItem "<Author>"
53650    .AddItem "<Computername>"
53660    .AddItem "<ClientComputer>"
53670    .AddItem "<DateTime>"
53680    .AddItem "<Title>"
53690    .AddItem "<Username>"
53700    .AddItem "<REDMON_DOCNAME>"
53710    .AddItem "<REDMON_DOCNAME_FILE>"
53720    .AddItem "<REDMON_DOCNAME_PATH>"
53730    .AddItem "<REDMON_JOB>"
53740    .AddItem "<REDMON_MACHINE>"
53750    .AddItem "<REDMON_PORT>"
53760    .AddItem "<REDMON_PRINTER>"
53770    .AddItem "<REDMON_SESSIONID>"
53780    .AddItem "<REDMON_USER>"
53790    .ListIndex = 0
53800   End With
53810   Me.Caption = .DialogPrinterOptions
53820   cmdCancel.Caption = .OptionsCancel
53830   cmdReset.Caption = .OptionsReset
53840   cmdSave.Caption = .OptionsSave
53850   tbstrPDFOptions.Tabs.Clear
53860   tbstrPDFOptions.Tabs.Add , "General", .OptionsPDFGeneral
53870   tbstrPDFOptions.Tabs.Add , "Compression", .OptionsPDFCompression
53880   tbstrPDFOptions.Tabs.Add , "Fonts", .OptionsPDFFonts
53890   tbstrPDFOptions.Tabs.Add , "Colors", .OptionsPDFColors
53900   tbstrPDFOptions.Tabs.Add , "Security", .OptionsPDFSecurity
53910   dmFraPDFGeneral.Caption = .OptionsPDFGeneralCaption
53920   chkPDFOptimize.Caption = .OptionsPDFOptimize
53930   lblPDFCompat.Caption = .OptionsPDFGeneralCompatibility
53940   lblPDFAutoRotate.Caption = .OptionsPDFGeneralAutorotate
53950   lblPDFResolution.Caption = .OptionsPDFGeneralResolution
53960   lblPDFOverprint.Caption = .OptionsPDFGeneralOverprint
53970   lblProgfont.Caption = .OptionsProgramFont
53980   lblProgcharset.Caption = .OptionsProgramFontcharset
53990   lblSize.Caption = .OptionsProgramFontSize
54000   lblTesttext = .OptionsProgramFontTestdescription
54010   cmdTest.Caption = .OptionsProgramFontTest
54020   cmdCancelTest.Caption = .OptionsProgramFontCancelTest
54030   chkPDFASCII85.Caption = .OptionsPDFGeneralASCII85
54040   cmbPDFCompat.Clear
54050   cmbPDFCompat.AddItem .OptionsPDFGeneralCompatibility01
54060   cmbPDFCompat.AddItem .OptionsPDFGeneralCompatibility02
54070   cmbPDFCompat.AddItem .OptionsPDFGeneralCompatibility03
54080   cmbPDFRotate.Clear
54090   cmbPDFRotate.AddItem .OptionsPDFGeneralRotate01
54100   cmbPDFRotate.AddItem .OptionsPDFGeneralRotate02
54110   cmbPDFRotate.AddItem .OptionsPDFGeneralRotate03
54120   cmbPDFOverprint.Clear
54130   cmbPDFOverprint.AddItem .OptionsPDFGeneralOverprint01
54140   cmbPDFOverprint.AddItem .OptionsPDFGeneralOverprint02
54150
54160   dmfraPDFCompress.Caption = .OptionsPDFCompressionCaption
54170   chkPDFTextComp.Caption = .OptionsPDFCompressionTextComp
54180   dmFraPDFColor.Caption = .OptionsPDFCompressionColor
54190   chkPDFColorComp.Caption = .OptionsPDFCompressionColorComp
54200   chkPDFColorResample.Caption = .OptionsPDFCompressionColorResample
54210   lblPDFColorRes.Caption = .OptionsPDFCompressionColorRes
54220   cmbPDFColorComp.Clear
54230   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp01
54240   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp02
54250   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp03
54260   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp04
54270   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp05
54280   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp06
54290   cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp07
54300 '  cmbPDFColorComp.AddItem .OptionsPDFCompressionColorComp08
54310   cmbPDFColorResample.Clear
54320   cmbPDFColorResample.AddItem .OptionsPDFCompressionColorResample01
54330   cmbPDFColorResample.AddItem .OptionsPDFCompressionColorResample02
54340 '  cmbPDFColorResample.AddItem .OptionsPDFCompressionColorResample03
54350   dmFraPDFGrey.Caption = .OptionsPDFCompressionGrey
54360   chkPDFGreyComp.Caption = .OptionsPDFCompressionGreyComp
54370   chkPDFGreyResample.Caption = .OptionsPDFCompressionGreyResample
54380   lblPDFGreyRes.Caption = .OptionsPDFCompressionGreyRes
54390   cmbPDFGreyComp.Clear
54400   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp01
54410   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp02
54420   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp03
54430   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp04
54440   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp05
54450   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp06
54460   cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp07
54470 '  cmbPDFGreyComp.AddItem .OptionsPDFCompressionGreyComp08
54480   cmbPDFGreyResample.Clear
54490   cmbPDFGreyResample.AddItem .OptionsPDFCompressionGreyResample01
54500   cmbPDFGreyResample.AddItem .OptionsPDFCompressionGreyResample02
54510 '  cmbPDFGreyResample.AddItem .OptionsPDFCompressionGreyResample03
54520   dmFraPDFMono.Caption = .OptionsPDFCompressionMono
54530   chkPDFMonoComp.Caption = .OptionsPDFCompressionMonoComp
54540   chkPDFMonoResample.Caption = .OptionsPDFCompressionMonoResample
54550   lblPDFMonoRes.Caption = .OptionsPDFCompressionMonoRes
54560   cmbPDFMonoComp.Clear
54570   cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp01
54580   cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp02
54590   cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp03
54600 '  cmbPDFMonoComp.AddItem .OptionsPDFCompressionMonoComp04
54610   cmbPDFMonoResample.Clear
54620   cmbPDFMonoResample.AddItem .OptionsPDFCompressionMonoResample01
54630   cmbPDFMonoResample.AddItem .OptionsPDFCompressionMonoResample02
54640 '  cmbPDFMonoResample.AddItem .OptionsPDFCompressionMonoResample03
54650
54660   dmFraPDFFonts.Caption = .OptionsPDFFontsCaption
54670   chkPDFEmbedAll.Caption = .OptionsPDFFontsEmbedAll
54680   chkPDFSubSetFonts.Caption = .OptionsPDFFontsSubSetFonts
54690
54700   dmFraPDFColors.Caption = .OptionsPDFColorsCaption
54710   chkPDFCMYKtoRGB.Caption = .OptionsPDFColorsCMYKtoRGB
54720   dmFraPDFColorOptions.Caption = .OptionsPDFColorsColorOptions
54730   chkPDFPreserveOverprint.Caption = .OptionsPDFColorsPreserveOverprint
54740   chkPDFPreserveTransfer.Caption = .OptionsPDFColorsPreserveTransfer
54750   chkPDFPreserveHalftone.Caption = .OptionsPDFColorsPreserveHalftone
54760   cmbPDFColorModel.Clear
54770   cmbPDFColorModel.AddItem .OptionsPDFColorsColorModel01
54780   cmbPDFColorModel.AddItem .OptionsPDFColorsColorModel02
54790   cmbPDFColorModel.AddItem .OptionsPDFColorsColorModel03
54800
54810   dmFraPDFEncryptor.Caption = .OptionsPDFEncryptor
54820   dmFraPDFSecurity.Caption = .OptionsPDFSecurityCaption
54830   chkUseSecurity.Caption = .OptionsPDFUseSecurity
54840   dmFraPDFEncLevel.Caption = .OptionsPDFEncryptionLevel
54850   optEncHigh.Caption = .OptionsPDFEncryptionHigh
54860   optEncLow.Caption = .OptionsPDFEncryptionLow
54870   dmFraSecurityPass.Caption = .OptionsPDFPasswords
54880   chkUserPass.Caption = .OptionsPDFUserPass
54890   chkOwnerPass.Caption = .OptionsPDFOwnerPass
54900   dmFraPDFPermissions.Caption = .OptionsPDFDisallowUser
54910   dmFraPDFHighPermissions.Caption = .OptionsPDFEnhancedPermissions
54920   chkAllowPrinting.Caption = .OptionsPDFDisallowPrint
54930   chkAllowModifyContents.Caption = .OptionsPDFDisallowModify
54940   chkAllowCopy.Caption = .OptionsPDFDisallowCopy
54950   chkAllowModifyAnnotations.Caption = .OptionsPDFDisallowModifyComments
54960   chkAllowDegradedPrinting.Caption = .OptionsPDFAllowDegradedPrinting
54970   chkAllowFillIn.Caption = .OptionsPDFAllowFillIn
54980   chkAllowAssembly.Caption = .OptionsPDFAllowAssembly
54990   chkAllowScreenReaders.Caption = .OptionsPDFAllowScreenReaders
55000
55010   cmbPNGColors.AddItem .OptionsPNGColorscount01
55020   cmbPNGColors.AddItem .OptionsPNGColorscount02
55030   cmbPNGColors.AddItem .OptionsPNGColorscount03
55040   cmbPNGColors.AddItem .OptionsPNGColorscount04
55050   cmbJPEGColors.Left = cmbPNGColors.Left
55060   cmbJPEGColors.Width = cmbPNGColors.Width
55070   cmbJPEGColors.Top = cmbPNGColors.Top
55080   cmbJPEGColors.AddItem .OptionsJPEGColorscount01
55090   cmbJPEGColors.AddItem .OptionsJPEGColorscount02
55100   cmbBMPColors.Left = cmbPNGColors.Left
55110   cmbBMPColors.Width = cmbPNGColors.Width
55120   cmbBMPColors.Top = cmbPNGColors.Top
55130   cmbBMPColors.AddItem .OptionsBMPColorscount01
55140   cmbBMPColors.AddItem .OptionsBMPColorscount02
55150   cmbBMPColors.AddItem .OptionsBMPColorscount03
55160   cmbBMPColors.AddItem .OptionsBMPColorscount04
55170   cmbBMPColors.AddItem .OptionsBMPColorscount05
55180   cmbBMPColors.AddItem .OptionsBMPColorscount06
55190   cmbBMPColors.AddItem .OptionsBMPColorscount07
55200   cmbPCXColors.Left = cmbPNGColors.Left
55210   cmbPCXColors.Width = cmbPNGColors.Width
55220   cmbPCXColors.Top = cmbPNGColors.Top
55230   cmbPCXColors.AddItem .OptionsPCXColorscount01
55240   cmbPCXColors.AddItem .OptionsPCXColorscount02
55250   cmbPCXColors.AddItem .OptionsPCXColorscount03
55260   cmbPCXColors.AddItem .OptionsPCXColorscount04
55270   cmbPCXColors.AddItem .OptionsPCXColorscount05
55280   cmbPCXColors.AddItem .OptionsPCXColorscount06
55290   cmbTIFFColors.Left = cmbPNGColors.Left
55300   cmbTIFFColors.Width = cmbPNGColors.Width
55310   cmbTIFFColors.Top = cmbPNGColors.Top
55320   cmbTIFFColors.AddItem .OptionsTIFFColorscount01
55330   cmbTIFFColors.AddItem .OptionsTIFFColorscount02
55340   cmbTIFFColors.AddItem .OptionsTIFFColorscount03
55350   cmbTIFFColors.AddItem .OptionsTIFFColorscount04
55360   cmbTIFFColors.AddItem .OptionsTIFFColorscount05
55370   cmbTIFFColors.AddItem .OptionsTIFFColorscount06
55380   cmbTIFFColors.AddItem .OptionsTIFFColorscount07
55390   cmbTIFFColors.AddItem .OptionsTIFFColorscount08
55400
55410   dmFraBitmapGeneral.Caption = .OptionsImageSettings
55420   lblBitmapResolution = .OptionsBitmapResolution
55430   lblJPEGQuality = .OptionsJPEGQuality
55440   lblBitmapColors = .OptionsPDFColors
55450   lblProcessPriority.Caption = .OptionsProcesspriority
55460   lblLangLevel.Caption = .OptionsPSLanguageLevel
55470
55480   cmdAsso.Caption = .OptionsAssociatePSFiles
55490
55500   lblStampString.Caption = .OptionsStampString
55510   lblStampFontcolor.Caption = .OptionsStampFontColor
55520   chkStampUseOutlineFont.Caption = .OptionsStampUseOutlineFont
55530   lblOutlineFontThickness.Caption = .OptionsStampOutlineFontThickness
55540
55550   chkUseFixPaperSize.Caption = .OptionsUseFixPapersize
55560   chkUseCustomPapersize.Caption = .OptionsUseCustomPapersize
55570   lblCustomPapersizeWidth.Caption = .OptionsCustomPapersizeWidth
55580   lblCustomPapersizeHeight.Caption = .OptionsCustomPapersizeHeight
55590   lblCustomPapersizeInfo.Caption = .OptionsCustomPapersizeInfo
55600
55610   lsvTranslations.ColumnHeaders.Add , , .OptionsLanguagesTranslation
55620   lsvTranslations.ColumnHeaders.Add , , .OptionsLanguagesVersion
55630
55640  End With
55650  lsvTranslations.ColumnHeaders(1).Width = 2000
55660  lsvTranslations.ColumnHeaders(2).Width = 1500
55670
55680  With cmbDocumentPapersizes
55690   .AddItem "11x17"
55700   .AddItem "ledger"
55710   .AddItem "legal"
55720   .AddItem "letter"
55730   .AddItem "lettersmall"
55740   .AddItem "archE"
55750   .AddItem "archD"
55760   .AddItem "archC"
55770   .AddItem "archB"
55780   .AddItem "archA"
55790   .AddItem "a0"
55800   .AddItem "a1"
55810   .AddItem "a2"
55820   .AddItem "a3"
55830   .AddItem "a4"
55840   .AddItem "a4small"
55850   .AddItem "a5"
55860   .AddItem "a6"
55870   .AddItem "a7"
55880   .AddItem "a8"
55890   .AddItem "a9"
55900   .AddItem "a10"
55910   .AddItem "isob0"
55920   .AddItem "isob1"
55930   .AddItem "isob2"
55940   .AddItem "isob3"
55950   .AddItem "isob4"
55960   .AddItem "isob5"
55970   .AddItem "isob6"
55980   .AddItem "c0"
55990   .AddItem "c1"
56000   .AddItem "c2"
56010   .AddItem "c3"
56020   .AddItem "c4"
56030   .AddItem "c5"
56040   .AddItem "c6"
56050   .AddItem "jisb0"
56060   .AddItem "jisb1"
56070   .AddItem "jisb2"
56080   .AddItem "jisb3"
56090   .AddItem "jisb4"
56100   .AddItem "jisb5"
56110   .AddItem "jisb6"
56120   .AddItem "b0"
56130   .AddItem "b1"
56140   .AddItem "b2"
56150   .AddItem "b3"
56160   .AddItem "b4"
56170   .AddItem "b5"
56180   .AddItem "flsa"
56190   .AddItem "flse"
56200   .AddItem "halfletter"
56210   .ListIndex = 0
56220  End With
56230
56240  If IsPsAssociate = False Then
56250    cmdAsso.Enabled = True
56260   Else
56270    cmdAsso.Enabled = False
56280  End If
56290
56300  txtPDFRes.Text = 600
56310  cmbPDFCompat.ListIndex = 1
56320  cmbPDFRotate.ListIndex = 0
56330  cmbPDFOverprint.ListIndex = 0
56340  chkPDFASCII85.Value = 0
56350
56360  chkPDFTextComp.Value = 1
56370
56380  chkPDFColorComp.Value = 1
56390  chkPDFColorResample.Value = 0
56400  cmbPDFColorComp.ListIndex = 0
56410  cmbPDFColorResample.ListIndex = 0
56420  txtPDFColorRes.Text = 300
56430
56440  chkPDFGreyComp.Value = 1
56450  chkPDFGreyResample.Value = 0
56460  cmbPDFGreyComp.ListIndex = 0
56470  cmbPDFGreyResample.ListIndex = 0
56480  txtPDFGreyRes.Text = 300
56490
56500  chkPDFMonoComp.Value = 1
56510  chkPDFMonoResample.Value = 0
56520  cmbPDFMonoComp.ListIndex = 0
56530  cmbPDFMonoResample.ListIndex = 0
56540  txtPDFMonoRes.Text = 1200
56550
56560  chkPDFEmbedAll.Value = 1
56570  chkPDFSubSetFonts.Value = 1
56580  txtPDFSubSetPerc.Text = 100
56590
56600  cmbPDFColorModel.ListIndex = 1
56610  chkPDFCMYKtoRGB.Value = 1
56620  chkPDFPreserveOverprint.Value = 1
56630  chkPDFPreserveTransfer.Value = 1
56640  chkPDFPreserveHalftone.Value = 0
56650
56660  cmbPNGColors.ListIndex = 0
56670  cmbJPEGColors.ListIndex = 0
56680  cmbBMPColors.ListIndex = 0
56690  cmbPCXColors.ListIndex = 0
56700  cmbTIFFColors.ListIndex = 0
56710  txtBitmapResolution.Text = 150
56720
56730 ' chkUseStandardAuthor.Value = 1
56740  txtStandardAuthor.Text = vbNullString
56750
56760  With cmbPSLanguageLevel
56770   .AddItem "1"
56780   .AddItem "1.5"
56790   .AddItem "2"
56800   .AddItem "3"
56810  End With
56820  With cmbEPSLanguageLevel
56830   .AddItem "1"
56840   .AddItem "1.5"
56850   .AddItem "2"
56860   .AddItem "3"
56870  End With
56880
56890  With lsvFilenameSubst
56900   .Appearance = ccFlat
56910   .ColumnHeaders.Clear
56920   .ColumnHeaders.Add , "Str1", "", lsvFilenameSubst.Width / 2 - 140
56930   .ColumnHeaders.Add , "Str2", "", lsvFilenameSubst.Width / 2 - 140
56940   .HideColumnHeaders = True
56950   .GridLines = True
56960   .FullRowSelect = True
56970   .HideSelection = False
56980  End With
56990
57000  With cmbPDFEncryptor
57010   .Clear
57020   .AddItem "Ghostscript (>= 8.14)"
57030   .ItemData(.NewIndex) = 0
57040   .AddItem "PDFEnc"
57050   .ItemData(.NewIndex) = 1
57060
57070   SecurityIsPossible = True
57080
57090   If FileExists(GetPDFCreatorApplicationPath & "pdfenc.exe") = False Then
57100    .RemoveItem 1
57110    .ListIndex = 0
57120    Options.PDFEncryptor = .ItemData(.ListIndex)
57130   End If
57140   If GhostScriptSecurity = False Then
57150    .RemoveItem 0
57160   End If
57170   If .ListCount = 0 Then
57180     chkUseSecurity.Value = 0
57190     chkUseSecurity.Enabled = False
57200     SecurityIsPossible = False
57210    Else
57220     For i = 0 To .ListCount - 1
57230      If .ItemData(i) = Options.PDFEncryptor Then
57240       .ListIndex = i
57250       Exit For
57260      End If
57270     Next i
57280     If .ListIndex = -1 Then
57290      .ListIndex = 0
57300      Options.PDFEncryptor = .ItemData(.ListIndex)
57310     End If
57320   End If
57330  End With
57340
57350  If Options.PDFHighEncryption <> 0 Then
57360    optEncHigh.Value = True
57370   Else
57380    optEncLow.Value = True
57390  End If
57400
57410  cmdFilenameSubst(0).Top = lsvFilenameSubst.Top
57420  cmdFilenameSubst(1).Top = lsvFilenameSubst.Top + (lsvFilenameSubst.Height - cmdFilenameSubst(1).Height) / 2
57430  cmdFilenameSubst(2).Top = lsvFilenameSubst.Top + lsvFilenameSubst.Height - cmdFilenameSubst(2).Height
57440
57450  If chkUseStandardAuthor.Value = 1 Then
57460    txtStandardAuthor.Enabled = True
57470    txtStandardAuthor.BackColor = &H80000005
57480   Else
57490    txtStandardAuthor.Enabled = False
57500    txtStandardAuthor.BackColor = &H8000000F
57510  End If
57520  With Options
57530   SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
57540  End With
57550  ieb.Refresh
57560  If chkUseAutosave.Value = 1 Then
57570    ViewAutosave True
57580   Else
57590    ViewAutosave False
57600  End If
57610  If chkPrintAfterSaving.Value = 1 Then
57620    ViewPrintAfterSaving True
57630   Else
57640    ViewPrintAfterSaving False
57650  End If
57660
57670  With txtGSbin
57680   .ToolTipText = .Text
57690  End With
57700  With txtGSlib
57710   .ToolTipText = .Text
57720  End With
57730  With txtGSfonts
57740   .ToolTipText = .Text
57750  End With
57760  With txtTemppath
57770   .ToolTipText = ResolveEnvironment(GetSubstFilename2(.Text))
57780  End With
57790
57800  With sldProcessPriority
57810   .TextPosition = sldBelowRight
57820   .TickFrequency = 1
57830   .TickStyle = sldTopLeft
57841   Select Case .Value
         Case 0: 'Idle
57860     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityIdle
57870    Case 1: 'Normal
57880     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityNormal
57890    Case 2: 'High
57900     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityHigh
57910    Case 3: 'Realtime
57920     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityRealtime
57930   End Select
57940  End With
57950
57960  If IsWin9xMe = False Then
57970    lblProcessPriority.Enabled = True
57980    sldProcessPriority.Enabled = True
57990   Else
58000    lblProcessPriority.Enabled = False
58010    sldProcessPriority.Enabled = False
58020  End If
58030  UpdateSecurityFields
58040
58050  If Options.RunProgramAfterSaving Then
58060    ViewRunProgramAfterSaving True
58070   Else
58080    ViewRunProgramAfterSaving False
58090  End If
58100  If Options.RunProgramBeforeSaving Then
58110    ViewRunProgramBeforeSaving True
58120   Else
58130    ViewRunProgramBeforeSaving False
58140  End If
58150
58160  Set Files = GetFiles(GetPDFCreatorApplicationPath & "Scripts\RunProgramAfterSaving\", "*.*", SortedByName)
58170  For i = 1 To Files.Count
58180   tsf = Split(Files(i), "|")
58190   SplitPath tsf(1), , Path, Filename, , Ext
58200   If UCase$(Ext) <> "TXT" And UCase$(Ext) <> "PDF" And UCase$(Ext) <> "PNG" And _
   UCase$(Ext) <> "JPG" And UCase$(Ext) <> "BMP" And UCase$(Ext) <> "PCX" And _
   UCase$(Ext) <> "TIF" And UCase$(Ext) <> "EPS" And UCase$(Ext) <> "PS" Then
58230    If UCase$(tsf(0)) <> UCase$(GetPDFCreatorApplicationPath & "Scripts\RunProgramAfterSaving\") Then
58240      cmbRunProgramAfterSavingProgramname.AddItem tsf(0)
58250     Else
58260      cmbRunProgramAfterSavingProgramname.AddItem Filename
58270    End If
58280   End If
58290  Next i
58300
58310  Set Files = GetFiles(GetPDFCreatorApplicationPath & "Scripts\RunProgramBeforeSaving\", "*.*", SortedByName)
58320  For i = 1 To Files.Count
58330   tsf = Split(Files(i), "|")
58340   SplitPath tsf(1), , Path, Filename, , Ext
58350   If UCase$(Ext) <> "TXT" And UCase$(Ext) <> "PDF" And UCase$(Ext) <> "PNG" And _
   UCase$(Ext) <> "JPG" And UCase$(Ext) <> "BMP" And UCase$(Ext) <> "PCX" And _
   UCase$(Ext) <> "TIF" And UCase$(Ext) <> "EPS" And UCase$(Ext) <> "PS" Then
58380    If UCase$(tsf(0)) <> UCase$(GetPDFCreatorApplicationPath & "Scripts\RunProgramBeforeSaving\") Then
58390      cmbRunProgramBeforeSavingProgramname.AddItem tsf(0)
58400     Else
58410      cmbRunProgramBeforeSavingProgramname.AddItem Filename
58420    End If
58430   End If
58440  Next i
58450
58460  tStr2 = CompletePath(UCase$(Trim$(Options.DirectoryGhostscriptBinaries)))
 cmbGhostscript.Clear: Set reg = New clsRegistry
58480  reg.hkey = HKEY_LOCAL_MACHINE
58490
58500  Set gsvers = GetAllGhostscriptversions
58510
58520  If gsvers.Count = 0 Then
58530    cmbGhostscript.Enabled = False
58540   Else
58550    For i = 1 To gsvers.Count
58560     cmbGhostscript.AddItem gsvers.Item(i)
58570    Next i
58580    cmbGhostscript.ListIndex = cmbGhostscript.ListCount - 1
58590    For i = 0 To cmbGhostscript.ListCount - 1
58600     tStr = ""
58610     If InStr(cmbGhostscript.List(i), ":") Then
58620       reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
58630       If tStr2 = CompletePath(UCase$(Trim$(reg.GetRegistryValue("GhostscriptDirectoryBinaries")))) Then
58640        cmbGhostscript.ListIndex = i
58650        Exit For
58660       End If
58670      Else
58680       If InStr(UCase$(cmbGhostscript.List(i)), "AFPL") Then
58690        reg.KeyRoot = "SOFTWARE\AFPL Ghostscript"
58700        If InStr(cmbGhostscript.List(i), " ") > 0 Then
58710         tsf = Split(cmbGhostscript.List(i), " ")
58720         reg.Subkey = tsf(UBound(tsf))
58730         tStr = reg.GetRegistryValue("GS_DLL")
58740         If tStr2 & "GSDLL32.DLL" = UCase$(tStr) Then
58750          cmbGhostscript.ListIndex = i
58760          Exit For
58770         End If
58780        End If
58790       End If
58800       If InStr(UCase$(cmbGhostscript.List(i)), "GNU") Then
58810        reg.KeyRoot = "SOFTWARE\GNU Ghostscript"
58820        If InStr(cmbGhostscript.List(i), " ") > 0 Then
58830         tsf = Split(cmbGhostscript.List(i), " ")
58840         reg.Subkey = tsf(UBound(tsf))
58850         tStr = reg.GetRegistryValue("GS_DLL")
58860         If tStr2 & "GSDLL32.DLL" = UCase$(tStr) Then
58870          cmbGhostscript.ListIndex = i
58880          Exit For
58890         End If
58900        End If
58910       End If
58920       If InStr(UCase$(cmbGhostscript.List(i)), "GPL") Then
58930        reg.KeyRoot = "SOFTWARE\GPL Ghostscript"
58940        If InStr(cmbGhostscript.List(i), " ") > 0 Then
58950         tsf = Split(cmbGhostscript.List(i), " ")
58960         reg.Subkey = tsf(UBound(tsf))
58970         tStr = reg.GetRegistryValue("GS_DLL")
58980         If tStr2 & "GSDLL32.DLL" = UCase$(tStr) Then
58990          cmbGhostscript.ListIndex = i
59000          Exit For
59010         End If
59020        End If
59030       End If
59040     End If
59050    Next i
59060  End If
59070  Set reg = Nothing
59080  With cmbGhostscript
59090   If .ListCount = 0 Then
59100    .Enabled = False
59110    .BackColor = &H8000000F
59120   End If
59130  End With
59140
59150  lblFontNameSize.Caption = Options.StampFontname & ", " & Options.StampFontsize
59160  If lblOutlineFontThickness.Left + lblOutlineFontThickness.Width + 50 + txtOutlineFontThickness.Width > dmFraProgStamp.Width Then
59170    txtOutlineFontThickness.Left = dmFraProgStamp.Width - txtOutlineFontThickness.Width - 10
59180   Else
59190    txtOutlineFontThickness.Left = lblOutlineFontThickness.Left + lblOutlineFontThickness.Width + 50
59200  End If
59210  txtOutlineFontThickness.Top = lblOutlineFontThickness.Top + (lblOutlineFontThickness.Height - txtOutlineFontThickness.Height) / 2
59220
59230  tbstrPDFOptions.ZOrder 1
59240  tbstrProgActions.ZOrder 1
59250
59260  If ShowOnlyOptions = True Then
59270   FormInTaskbar Me, True, True
59280   Caption = "PDFCreator - " & Caption
59290  End If
59300
59310  ShowAcceleratorsInForm Me, True
59320
59330  ShowOptions Me, Options
59340  If chkStampUseOutlineFont.Value = 1 Then
59350    lblOutlineFontThickness.Enabled = True
59360    txtOutlineFontThickness.Enabled = True
59370    txtOutlineFontThickness.BackColor = &H80000005
59380   Else
59390    lblOutlineFontThickness.Enabled = False
59400    txtOutlineFontThickness.Enabled = False
59410    txtOutlineFontThickness.BackColor = &H8000000F
59420  End If
59430  If chkUseFixPaperSize.Value = 1 Then
59440    cmbDocumentPapersizes.Enabled = True
59450    chkUseCustomPapersize.Enabled = True
59460    If chkUseCustomPapersize.Value = 1 Then
59470      lblCustomPapersizeWidth.Enabled = True
59480      lblCustomPapersizeHeight.Enabled = True
59490      txtCustomPapersizeWidth.Enabled = True
59500      txtCustomPapersizeWidth.BackColor = &H80000005
59510      txtCustomPapersizeHeight.Enabled = True
59520      txtCustomPapersizeHeight.BackColor = &H80000005
59530      lblCustomPapersizeInfo.Enabled = True
59540      cmbDocumentPapersizes.Enabled = True
59550      lblCustomPapersizeInfo.Enabled = True
59560     Else
59570      cmbDocumentPapersizes.Enabled = True
59580      lblCustomPapersizeWidth.Enabled = False
59590      lblCustomPapersizeHeight.Enabled = False
59600      txtCustomPapersizeWidth.Enabled = False
59610      txtCustomPapersizeWidth.BackColor = &H8000000F
59620      txtCustomPapersizeHeight.Enabled = False
59630      txtCustomPapersizeHeight.BackColor = &H8000000F
59640      lblCustomPapersizeInfo.Enabled = False
59650      lblCustomPapersizeInfo.Enabled = False
59660    End If
59670   Else
59680    cmbDocumentPapersizes.Enabled = False
59690    chkUseCustomPapersize.Enabled = False
59700    lblCustomPapersizeWidth.Enabled = False
59710    lblCustomPapersizeHeight.Enabled = False
59720    txtCustomPapersizeWidth.Enabled = False
59730    txtCustomPapersizeWidth.BackColor = &H8000000F
59740    txtCustomPapersizeHeight.Enabled = False
59750    txtCustomPapersizeHeight.BackColor = &H8000000F
59760    lblCustomPapersizeInfo.Enabled = False
59770  End If
59780  ReadAllLanguages LanguagePath, True
59790  Screen.MousePointer = vbNormal
59800  Timer1.Enabled = True
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  UnloadForm = True
50020  Me.Visible = False
50030  If TimerReady = False Then
50040   Timer2.Enabled = True
50050   Cancel = True
50060  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "Form_QueryUnload")
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
50100  tbstrProgDocument.Visible = False
50110  For Each ctl In Controls
50120   If TypeOf ctl Is dmFrame Then
50130    ctl.Visible = False
50140    ctl.Enabled = False
50150   End If
50160  Next
50170  tbstrPDFOptions.Enabled = False
50180  dmFraDescription.Visible = True
50190  dmFraDescription.Enabled = True
50200  txtJPEGQuality.Visible = False
50210  lblJPEQQualityProzent.Visible = False
50220  dmFraPSGeneral.Visible = False
50230  cmbPSLanguageLevel.Visible = False
50240  cmbEPSLanguageLevel.Visible = False
50251  Select Case UCase$(sGroup)
        Case "PROGRAM"
50271    Select Case UCase$(sItemKey)
          Case "GENERAL"
50290      Set picOptions = LoadResPicture(2101, vbResIcon)
50300      lblOptions.Caption = LanguageStrings.OptionsProgramGeneralDescription
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
50460      lblOptions.Caption = LanguageStrings.OptionsProgramGhostscriptDescription
50470      dmFraProgGhostscript.Enabled = True
50480      dmFraProgGhostscript.Visible = True
50490      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50500     Case "DOCUMENT"
50510      Set picOptions = LoadResPicture(2105, vbResIcon)
50520      lblOptions.Caption = LanguageStrings.OptionsProgramDocumentDescription
50530      tbstrProgDocument.Enabled = True
50540      tbstrProgDocument.Visible = True
50551      Select Case tbstrProgDocument.SelectedItem.Index
            Case 1
50570        dmFraProgDocument1.Enabled = True
50580        dmFraProgDocument1.Visible = True
50590        dmFraProgStamp.Enabled = True
50600        dmFraProgStamp.Visible = True
50610       Case 2
50620        dmFraProgDocument2.Enabled = True
50630        dmFraProgDocument2.Visible = True
50640      End Select
50650      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50660     Case "SAVE"
50670      Set picOptions = LoadResPicture(2106, vbResIcon)
50680      lblOptions.Caption = LanguageStrings.OptionsProgramSaveDescription
50690      dmFraProgSave.Enabled = True
50700      dmFraProgSave.Visible = True
50710      dmFraFilenameSubstitutions.Visible = True
50720      dmFraFilenameSubstitutions.Enabled = True
50730      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50740     Case "AUTOSAVE"
50750      Set picOptions = LoadResPicture(2103, vbResIcon)
50760      lblOptions.Caption = LanguageStrings.OptionsProgramAutosaveDescription
50770      dmFraProgAutosave.Enabled = True
50780      dmFraProgAutosave.Visible = True
50790      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50800     Case "DIRECTORIES"
50810      Set picOptions = LoadResPicture(2104, vbResIcon)
50820      lblOptions.Caption = LanguageStrings.OptionsProgramDirectoriesDescription
50830      dmFraProgDirectories.Enabled = True
50840      dmFraProgDirectories.Visible = True
50850      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50860     Case "ACTIONS"
50870      Set picOptions = LoadResPicture(2121, vbResIcon)
50880      lblOptions.Caption = LanguageStrings.OptionsProgramActionsDescription
50890      dmFraProgActions.Enabled = True
50900      dmFraProgActions.Visible = True
50910      ViewProgActions
50920      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50930     Case "PRINT"
50940      Set picOptions = LoadResPicture(2122, vbResIcon)
50950      lblOptions.Caption = LanguageStrings.OptionsProgramPrintDescription
50960      dmFraProgPrint.Enabled = True
50970      dmFraProgPrint.Visible = True
50980      dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50990     Case "FONTS"
51000      Set picOptions = LoadResPicture(2102, vbResIcon)
51010      lblOptions.Caption = LanguageStrings.OptionsProgramFontDescription
51020      dmFraProgFont.Enabled = True
51030      dmFraProgFont.Visible = True
51040     Case "LANGUAGE"
51050      Set picOptions = LoadResPicture(2123, vbResIcon)
51060      lblOptions.Caption = LanguageStrings.OptionsProgramLanguagesDescription
51070      dmFraProgLanguage.Enabled = True
51080      dmFraProgLanguage.Visible = True
51090    End Select
51100   Case "FORMATS"
51111    Select Case UCase$(sItemKey)
          Case "PDF"
51130      Set picOptions = LoadResPicture(2111, vbResIcon)
51140      lblOptions.Caption = LanguageStrings.OptionsPDFDescription
51150      tbstrPDFOptions.Enabled = True
51160      tbstrPDFOptions.Visible = True
51170      tbstrPDFOptions_Click
51180      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51190     Case "PNG"
51200      Set picOptions = LoadResPicture(2112, vbResIcon)
51210      lblOptions.Caption = LanguageStrings.OptionsPNGDescription
51220      dmFraBitmapGeneral.Enabled = True
51230      dmFraBitmapGeneral.Visible = True
51240      cmbPNGColors.Visible = True
51250      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51260     Case "JPEG"
51270      Set picOptions = LoadResPicture(2113, vbResIcon)
51280      lblOptions.Caption = LanguageStrings.OptionsJPEGDescription
51290      dmFraBitmapGeneral.Enabled = True
51300      dmFraBitmapGeneral.Visible = True
51310      lblJPEGQuality.Caption = LanguageStrings.OptionsJPEGQuality
51320      lblJPEGQuality.Visible = True
51330      txtJPEGQuality.Visible = True
51340      lblJPEQQualityProzent.Visible = True
51350      lblJPEQQualityProzent.Left = txtJPEGQuality.Left + txtJPEGQuality.Width + 100
51360      cmbJPEGColors.Visible = True
51370      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51380     Case "BMP"
51390      Set picOptions = LoadResPicture(2114, vbResIcon)
51400      lblOptions.Caption = LanguageStrings.OptionsBMPDescription
51410      dmFraBitmapGeneral.Enabled = True
51420      dmFraBitmapGeneral.Visible = True
51430      cmbBMPColors.Visible = True
51440      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51450     Case "PCX"
51460      Set picOptions = LoadResPicture(2115, vbResIcon)
51470      lblOptions.Caption = LanguageStrings.OptionsPCXDescription
51480      dmFraBitmapGeneral.Enabled = True
51490      dmFraBitmapGeneral.Visible = True
51500      cmbPCXColors.Visible = True
51510      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51520     Case "TIFF"
51530      Set picOptions = LoadResPicture(2116, vbResIcon)
51540      lblOptions.Caption = LanguageStrings.OptionsTIFFDescription
51550      dmFraBitmapGeneral.Enabled = True
51560      dmFraBitmapGeneral.Visible = True
51570      cmbTIFFColors.Visible = True
51580      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51590     Case "PS"
51600      Set picOptions = LoadResPicture(2117, vbResIcon)
51610      lblOptions.Caption = LanguageStrings.OptionsPSDescription
51620      dmFraPSGeneral.Enabled = True
51630      dmFraPSGeneral.Visible = True
51640      cmbPSLanguageLevel.Visible = True
51650      dmFraPSGeneral.Caption = LanguageStrings.OptionsPSDescription
51660      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51670     Case "EPS"
51680      Set picOptions = LoadResPicture(2118, vbResIcon)
51690      lblOptions.Caption = LanguageStrings.OptionsEPSDescription
51700      dmFraPSGeneral.Enabled = True
51710      dmFraPSGeneral.Visible = True
51720      cmbEPSLanguageLevel.Visible = True
51730      dmFraPSGeneral.Caption = LanguageStrings.OptionsEPSDescription
51740      dmFraDescription.Caption = LanguageStrings.OptionsTreeFormats
51750    End Select
51760  End Select
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

Private Sub picStampFontColor_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim c As OLE_COLOR
50020  If OpenColorDialog(c, Me.hwnd) = 1 Then
50030   picStampFontColor.BackColor = c
50040   Options.StampFontColor = OleColorToHTMLColor(c)
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "picStampFontColor_Click")
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

Private Sub tbstrProgDocument_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case tbstrProgDocument.SelectedItem.Index
        Case 1
50030    dmFraProgDocument2.Enabled = False
50040    dmFraProgDocument2.Visible = False
50050    dmFraProgDocument1.Enabled = True
50060    dmFraProgDocument1.Visible = True
50070    dmFraProgStamp.Enabled = True
50080    dmFraProgStamp.Visible = True
50090   Case 2
50100    dmFraProgDocument1.Enabled = False
50110    dmFraProgDocument1.Visible = False
50120    dmFraProgStamp.Enabled = False
50130    dmFraProgStamp.Visible = False
50140    dmFraProgDocument2.Enabled = True
50150    dmFraProgDocument2.Visible = True
50160  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "tbstrProgDocument_Click")
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
50030    dmFraProgGeneral2.Enabled = False
50040    dmFraProgGeneral2.Visible = False
50050    dmFraShellIntegration.Enabled = False
50060    dmFraShellIntegration.Visible = False
50070    dmFraProgGeneral1.Enabled = True
50080    dmFraProgGeneral1.Visible = True
50090   Case 2
50100    dmFraProgGeneral1.Enabled = False
50110    dmFraProgGeneral1.Visible = False
50120    dmFraProgGeneral2.Enabled = True
50130    dmFraProgGeneral2.Visible = True
50140    dmFraShellIntegration.Enabled = True
50150    dmFraShellIntegration.Visible = True
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
50040  TimerReady = False
50050  Set cSystem = New clsSystem
50060  Set SMF = cSystem.GetSystemFont(Me, Menu)
50070  txtTest.Text = vbNullString
50080  For i = 33 To 255
50090   txtTest.Text = txtTest.Text & Chr$(i)
50100   If UnloadForm Then
50110    TimerReady = True
50120    Exit Sub
50130   End If
50140   DoEvents
50150  Next i
50160  With cmbCharset
50170   .Clear
50180   .AddItem "0, Western": .ItemData(.NewIndex) = 0
50190   .AddItem "2, Symbol": .ItemData(.NewIndex) = 2
50200   .AddItem "77, Mac": .ItemData(.NewIndex) = 77
50210   .AddItem "161, Greek": .ItemData(.NewIndex) = 161
50220   .AddItem "162, Turkish": .ItemData(.NewIndex) = 162
50230   .AddItem "177, Hebrew": .ItemData(.NewIndex) = 177
50240   .AddItem "178, Arabic": .ItemData(.NewIndex) = 178
50250   .AddItem "186, Baltic": .ItemData(.NewIndex) = 186
50260   .AddItem "204, Cyrillic": .ItemData(.NewIndex) = 204
50270   .AddItem "238, Central European": .ItemData(.NewIndex) = 238
50280   .AddItem "255, DOS/OEM": .ItemData(.NewIndex) = 255
50290   .Text = 0
50300  End With
50310  With cmbProgramFontsize
50320   .AddItem "8"
50330   .AddItem "9"
50340   .AddItem "10"
50350   .AddItem "11"
50360   .AddItem "12"
50370   .AddItem "14"
50380   .AddItem "16"
50390   .AddItem "18"
50400   .AddItem "20"
50410   .AddItem "22"
50420   .AddItem "24"
50430   .AddItem "26"
50440   .AddItem "28"
50450   .AddItem "36"
50460   .AddItem "48"
50470   .AddItem "72"
50480  End With
50490  cmbProgramFontsize.Text = 8
50500  cmbCharset.Text = cmbCharset.ItemData(0)
50510  cmbCharset.Text = Options.ProgramFontCharset
50520  fi = -1
50530  With cmbFonts
50540   For i = 1 To Screen.FontCount
50550    tStr = Trim$(Screen.Fonts(i))
50560    If LenB(tStr) > 0 Then
50570     cmbFonts.AddItem tStr
50580     If UCase$(cmbFonts.List(i)) = UCase$(SMF(1)(0)) Then
50590      fi = i
50600     End If
50610    End If
50620    If UnloadForm Then
50630     TimerReady = True
50640     Exit Sub
50650    End If
50660    DoEvents
50670   Next i
50680  End With
50690  For Each ctl In Controls
50700   If UnloadForm Then
50710    TimerReady = True
50720    Exit Sub
50730   End If
50740   DoEvents
50750   If TypeOf ctl Is ComboBox Then
50760    ComboSetListWidth ctl
50770   End If
50780  Next ctl
50790
50800  SetOptimalComboboxHeigth cmbCharset, Me
50810  SetOptimalComboboxHeigth cmbProgramFontsize, Me
50820  SetOptimalComboboxHeigth cmbGhostscript, Me
50830
50840  Form_Resize
50850
50860  cmbProgramFontsize.Width = txtTest.Width - _
  (cmbProgramFontsize.Left - txtTest.Left)
50880
50890  If fi >= 0 Then
50900   cmbFonts.ListIndex = fi
50910   cmbCharset.Text = SMF(1)(2)
50920   cmbProgramFontsize.Text = SMF(1)(1)
50930   txtTest.Font = cmbFonts.List(cmbFonts.ListIndex)
50940   txtTest.Font.Charset = cmbCharset.Text
50950  End If
50960
50970  ShowOptions Me, Options
50980
50990  If Options.UseAutosaveDirectory = "1" Then
51000    ViewAutosaveDirectory True
51010   Else
51020    ViewAutosaveDirectory False
51030  End If
51040  If Options.UseAutosave = "1" Then
51050    ViewAutosave True
51060   Else
51070    ViewAutosave False
51080  End If
51090
51100  CheckCmdFilenameSubst
51110  CorrectCmbCharset
51120  tbstrProgActions.Tabs(2).Selected = True
51130
51140  TimerReady = True
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

Private Sub Timer2_Timer()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If TimerReady Then
50020   Timer2.Enabled = False
50030  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "Timer2_Timer")
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
50010  KeyAscii = AllowedKeypressChars(KeyAscii)
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

Private Sub ViewAutosave(Viewit As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  lblAutosaveformat.Enabled = Viewit
50020  cmbAutosaveFormat.Enabled = Viewit
50030  lblAutosaveFilename.Enabled = Viewit
50040  txtAutosaveFilename.Enabled = Viewit
50050  txtAutoSaveFilenamePreview.Enabled = Viewit
50060  lblAutosaveFilenameTokens.Enabled = Viewit
50070  cmbAutoSaveFilenameTokens.Enabled = Viewit
50080  chkUseAutosaveDirectory.Enabled = Viewit
50090  txtAutoSaveDirectoryPreview.Enabled = Viewit
50100  chkAutosaveStartStandardProgram.Enabled = Viewit
50110  chkAutosaveSendEmail.Enabled = Viewit
50120
50130  If Viewit Then
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
50240  If chkUseAutosaveDirectory.Value = 1 And Viewit Then
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

Private Sub ViewAutosaveDirectory(Viewit As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txtAutosaveDirectory.Enabled = Viewit
50020  txtAutoSaveDirectoryPreview.Enabled = Viewit
50030  cmdGetAutosaveDirectory.Enabled = Viewit
50040  If Viewit = True Then
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

Private Sub ViewRunProgramAfterSaving(Viewit As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  lblRunProgramAfterSavingProgramname.Enabled = Viewit
50020  cmbRunProgramAfterSavingProgramname.Enabled = Viewit
50030  lblRunProgramAfterSavingProgramParameters.Enabled = Viewit
50040  txtRunProgramAfterSavingProgramParameters.Enabled = Viewit
50050  chkRunProgramAfterSavingWaitUntilReady.Enabled = Viewit
50060  lblRunProgramAfterSavingWindowstyle.Enabled = Viewit
50070  cmbRunProgramAfterSavingWindowstyle.Enabled = Viewit
50080  cmdRunProgramAfterSavingPrognameChoice.Enabled = Viewit
50090  cmdRunProgramAfterSavingPrognameEdit.Enabled = Viewit
50100
50110  If Viewit Then
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

Private Sub ViewRunProgramBeforeSaving(Viewit As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  lblRunProgramBeforeSavingProgramname.Enabled = Viewit
50020  cmbRunProgramBeforeSavingProgramname.Enabled = Viewit
50030  lblRunProgramBeforeSavingProgramParameters.Enabled = Viewit
50040  txtRunProgramBeforeSavingProgramParameters.Enabled = Viewit
50050  lblRunProgramBeforeSavingWindowstyle.Enabled = Viewit
50060  cmbRunProgramBeforeSavingWindowstyle.Enabled = Viewit
50070  cmdRunProgramBeforeSavingPrognameChoice.Enabled = Viewit
50080  cmdRunProgramBeforeSavingPrognameEdit.Enabled = Viewit
50090
50100  If Viewit Then
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

Private Sub ViewPrintAfterSaving(Viewit As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  lblPrintAfterSavingPrinter.Enabled = Viewit
50020  cmbPrintAfterSavingPrinter.Enabled = Viewit
50030  lblPrintAfterSavingQueryUser.Enabled = Viewit
50040  cmbPrintAfterSavingQueryUser.Enabled = Viewit
50050  chkPrintAfterSavingNoCancel.Enabled = Viewit
50060  chkPrintAfterSavingDuplex.Enabled = Viewit
50070
50080  If Viewit Then
50090    cmbPrintAfterSavingPrinter.BackColor = &H80000005
50100    cmbPrintAfterSavingQueryUser.BackColor = &H80000005
50110   Else
50120    cmbPrintAfterSavingPrinter.BackColor = &H8000000F
50130    cmbPrintAfterSavingQueryUser.BackColor = &H8000000F
50140  End If
50150
50160  If chkPrintAfterSavingDuplex.Value = 1 And Viewit Then
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

Private Sub ViewPrintAfterTumple(Viewit As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  cmbPrintAfterSavingTumble.Enabled = Viewit
50020
50030  If Viewit Then
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

Private Sub txtCustomPapersizeHeight_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  KeyAscii = AllowedKeypressChars(KeyAscii)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "txtCustomPapersizeHeight_KeyPress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtCustomPapersizeWidth_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  KeyAscii = AllowedKeypressChars(KeyAscii)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "txtCustomPapersizeWidth_KeyPress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtOutlineFontThickness_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  KeyAscii = AllowedKeypressChars(KeyAscii)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "txtOutlineFontThickness_KeyPress")
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

Private Sub txtStampString_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Viewit As Boolean
50020  If LenB(txtStampString.Text) > 0 Then
50030    Viewit = True
50040   Else
50050    Viewit = False
50060  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "txtStampString_Change")
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

Public Sub ChangeLanguage()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  dmFraDescription.Caption = LanguageStrings.OptionsTreeProgram
50020
50030  With LanguageStrings
50040   ieb.DisableUpdates True
50050   ieb.SetGroupCaption "Program", .OptionsTreeProgram
50060   ieb.SetItemText "Program", "General", .OptionsProgramGeneralSymbol
50070   ieb.SetItemText "Program", "Ghostscript", .OptionsProgramGhostscriptSymbol
50080   ieb.SetItemText "Program", "Document", .OptionsProgramDocumentSymbol
50090   ieb.SetItemText "Program", "Save", .OptionsProgramSaveSymbol
50100   ieb.SetItemText "Program", "AutoSave", .OptionsProgramAutosaveSymbol
50110   ieb.SetItemText "Program", "Directories", .OptionsProgramDirectoriesSymbol
50120
50130   ieb.SetItemText "Program", "Actions", .OptionsProgramActionsSymbol
50140   ieb.SetItemText "Program", "Print", .OptionsProgramPrintSymbol
50150   ieb.SetItemText "Program", "Fonts", .OptionsProgramFontSymbol
50160   ieb.SetItemText "Program", "Language", .OptionsProgramLanguagesSymbol
50170
50180   ieb.SetGroupCaption "Formats", .OptionsTreeFormats
50190   ieb.SetItemText "Formats", "PDF", .OptionsPDFSymbol
50200   ieb.SetItemText "Formats", "PNG", .OptionsPNGSymbol
50210   ieb.SetItemText "Formats", "JPEG", .OptionsJPEGSymbol
50220   ieb.SetItemText "Formats", "BMP", .OptionsBMPSymbol
50230   ieb.SetItemText "Formats", "PCX", .OptionsPCXSymbol
50240   ieb.SetItemText "Formats", "TIFF", .OptionsTIFFSymbol
50250   ieb.SetItemText "Formats", "PS", .OptionsPSSymbol
50260   ieb.SetItemText "Formats", "EPS", .OptionsEPSSymbol
50270   ieb.DisableUpdates False
50280
50290   dmFraProgGeneral1.Caption = .OptionsProgramGeneralDescription1
50300   dmFraProgGeneral2.Caption = .OptionsProgramGeneralDescription2
50310   With tbstrProgGeneral
50320    .Tabs(1).Caption = LanguageStrings.OptionsProgramGeneralDescription1
50330    .Tabs(2).Caption = LanguageStrings.OptionsProgramGeneralDescription2
50340   End With
50350   With tbstrProgDocument
50360    .Tabs(1).Caption = LanguageStrings.OptionsProgramDocumentDescription1
50370    .Tabs(2).Caption = LanguageStrings.OptionsProgramDocumentDescription2
50380   End With
50390   dmFraShellIntegration.Caption = .OptionsShellIntegration
50400   dmFraProgGhostscript.Caption = .OptionsProgramGhostscriptSymbol
50410   dmFraProgAutosave.Caption = .OptionsProgramAutosaveSymbol
50420   dmFraProgDirectories.Caption = .OptionsProgramDirectoriesSymbol
50430   dmFraProgDocument1.Caption = .OptionsProgramDocumentDescription1
50440   dmFraProgDocument2.Caption = .OptionsProgramDocumentDescription2
50450   dmFraProgStamp.Caption = .OptionsStamp
50460   dmFraProgFont.Caption = .OptionsProgramFontSymbol
50470   dmFraProgSave.Caption = .OptionsProgramSaveSymbol
50480   dmFraProgActions.Caption = .OptionsProgramActionsSymbol
50490   dmFraProgPrint.Caption = .OptionsProgramPrintSymbol
50500   dmFraProgLanguage.Caption = .OptionsProgramLanguagesSymbol
50510
50520   cmdShellintegration(0).Caption = .OptionsShellIntegrationAdd
50530   cmdShellintegration(1).Caption = .OptionsShellIntegrationRemove
50540
50550   lblSendMailMethod.Caption = .OptionsSendMailMethod
50560   cmbSendMailMethod.List(0) = .OptionsSendMailMethodAutomatic
50570   cmbSendMailMethod.List(1) = .OptionsSendMailMethodMapi
50580   cmbSendMailMethod.List(2) = .OptionsSendMailMethodSendmailDLL
50590
50600   cmdLanguageInstall.Caption = .OptionsLanguagesInstall
50610   cmdLanguageRefresh.Caption = .OptionsLanguagesRefresh
50620   lblLanguagesFromInternet.Caption = .OptionsLanguagesDownloadMoreLanguages
50630   lsvTranslations.ColumnHeaders(1).Text = .OptionsLanguagesTranslation
50640   lsvTranslations.ColumnHeaders(2).Text = .OptionsLanguagesVersion
50650
50660   lblCurrentLanguage.Caption = .OptionsLanguagesCurrentLanguage
50670
50680   lblGhostscriptversion.Caption = .OptionsGhostscriptversion
50690   lblAdditionalGhostscriptParameters.Caption = .OptionsAdditionalGhostscriptParameters
50700   lblAdditionalGhostscriptSearchpath.Caption = .OptionsAdditionalGhostscriptSearchpath
50710   chkAddWindowsFontpath.Caption = .OptionsAddWindowsFontpath
50720
50730   lblSaveFilename.Caption = .OptionsSaveFilename
50740   lblSaveFilenameTokens.Caption = .OptionsSaveFilenameTokens
50750   dmFraFilenameSubstitutions.Caption = .OptionsSaveFilenameSubstitutions
50760   chkFilenameSubst.Caption = .OptionsSaveFilenameSubstitutionsTitle
50770   cmdFilenameSubst(0).Caption = .OptionsSaveFilenameAdd
50780   cmdFilenameSubst(1).Caption = .OptionsSaveFilenameChange
50790   cmdFilenameSubst(2).Caption = .OptionsSaveFilenameDelete
50800
50810   chkSpaces.Caption = .OptionsRemoveSpaces
50820   chkNoConfirmMessageSwitchingDefaultprinter.Caption = .OptionsProgramSwitchingDefaultprinter
50830   chkNoProcessingAtStartup.Caption = .OptionsProgramNoProcessingAtStartup
50840   lblOptionsDesign.Caption = .OptionsProgramOptionsDesign
50850   cmbOptionsDesign.List(0) = .OptionsProgramOptionsDesignGradient
50860   cmbOptionsDesign.List(1) = .OptionsProgramOptionsDesignSimple
50870   chkShowAnimation.Caption = .OptionsProgramShowAnimation
50880
50890   lblGSbin.Caption = .OptionsDirectoriesGSBin
50900   lblGSlib.Caption = .OptionsDirectoriesGSLibraries
50910   lblGSfonts.Caption = .OptionsDirectoriesGSFonts
50920   lblPrintTempPath.Caption = .OptionsDirectoriesTempPath
50930
50940   chkOnePagePerFile.Caption = .OptionsOnePagePerFile
50950   lblOptions = .OptionsProgramGeneralDescription
50960   lblAutosaveformat.Caption = .OptionsAutosaveFormat
50970   chkUseStandardAuthor.Caption = .OptionsUseStandardauthor
50980   chkUseCreationDateNow.Caption = .OptionsUseCreationDateNow
50990   chkUseAutosave.Caption = .OptionsUseAutosave
51000   cmdTestpage.Caption = .OptionsPrintTestpage
51010   lblAutosaveFilename.Caption = .OptionsAutosaveFilename
51020   lblAutosaveFilenameTokens.Caption = .OptionsAutosaveFilenameTokens
51030   chkUseAutosaveDirectory.Caption = .OptionsUseAutosaveDirectory
51040   lblAuthorTokens.Caption = .OptionsStandardAuthorToken
51050   chkAutosaveStartStandardProgram.Caption = .OptionsAutosaveStartStandardProgram
51060   chkAutosaveSendEmail.Caption = .OptionsSendEmailAfterAutosave
51070   lblStandardSaveformat.Caption = .OptionsStandardSaveFormat
51080
51090   dmFraProgActionsRunProgramAfterSaving.Caption = .OptionsProgramRunProgramAfterSavingCaption
51100   chkRunProgramAfterSaving.Caption = .OptionsProgramRunProgramAfterSavingCaption
51110   lblRunProgramAfterSavingProgramname.Caption = .OptionsProgramRunProgramAfterSavingProgram
51120   lblRunProgramAfterSavingProgramParameters.Caption = .OptionsProgramRunProgramAfterSavingProgramParameters
51130   chkRunProgramAfterSavingWaitUntilReady.Caption = .OptionsProgramRunProgramAfterSavingWaitUntilReady
51140   lblRunProgramAfterSavingWindowstyle.Caption = .OptionsProgramRunProgramAfterSavingWindowstyle
51150
51160   cmbRunProgramAfterSavingWindowstyle.List(0) = .OptionsProgramRunProgramAfterSavingWindowstyleHide
51170   cmbRunProgramAfterSavingWindowstyle.List(1) = .OptionsProgramRunProgramAfterSavingWindowstyleNormalFocus
51180   cmbRunProgramAfterSavingWindowstyle.List(2) = .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedFocus
51190   cmbRunProgramAfterSavingWindowstyle.List(3) = .OptionsProgramRunProgramAfterSavingWindowstyleMaximizedFocus
51200   cmbRunProgramAfterSavingWindowstyle.List(4) = .OptionsProgramRunProgramAfterSavingWindowstyleNormalNoFocus
51210   cmbRunProgramAfterSavingWindowstyle.List(5) = .OptionsProgramRunProgramAfterSavingWindowstyleMinimizedNoFocus
51220
51230   With tbstrProgActions
51240    .Tabs(1).Caption = LanguageStrings.OptionsProgramRunProgramBeforeSavingCaption
51250    .Tabs(2).Caption = LanguageStrings.OptionsProgramRunProgramAfterSavingCaption
51260   End With
51270
51280   dmFraProgActionsRunProgramBeforeSaving.Caption = .OptionsProgramRunProgramBeforeSavingCaption
51290   chkRunProgramBeforeSaving.Caption = .OptionsProgramRunProgramBeforeSavingCaption
51300   lblRunProgramBeforeSavingProgramname.Caption = .OptionsProgramRunProgramBeforeSavingProgram
51310   lblRunProgramBeforeSavingProgramParameters.Caption = .OptionsProgramRunProgramBeforeSavingProgramParameters
51320   lblRunProgramBeforeSavingWindowstyle.Caption = .OptionsProgramRunProgramBeforeSavingWindowstyle
51330   cmbRunProgramBeforeSavingWindowstyle.List(0) = .OptionsProgramRunProgramBeforeSavingWindowstyleHide
51340   cmbRunProgramBeforeSavingWindowstyle.List(1) = .OptionsProgramRunProgramBeforeSavingWindowstyleNormalFocus
51350   cmbRunProgramBeforeSavingWindowstyle.List(2) = .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedFocus
51360   cmbRunProgramBeforeSavingWindowstyle.List(3) = .OptionsProgramRunProgramBeforeSavingWindowstyleMaximizedFocus
51370   cmbRunProgramBeforeSavingWindowstyle.List(4) = .OptionsProgramRunProgramBeforeSavingWindowstyleNormalNoFocus
51380   cmbRunProgramBeforeSavingWindowstyle.List(5) = .OptionsProgramRunProgramBeforeSavingWindowstyleMinimizedNoFocus
51390
51400   chkPrintAfterSaving.Caption = .OptionsPrintAfterSaving
51410   lblPrintAfterSavingPrinter.Caption = .OptionsPrintAfterSavingPrinter
51420
51430   lblPrintAfterSavingQueryUser.Caption = .OptionsPrintAfterSavingQueryUser
51440   cmbPrintAfterSavingQueryUser.List(0) = .OptionsPrintAfterSavingQueryUserOff
51450   cmbPrintAfterSavingQueryUser.List(1) = .OptionsPrintAfterSavingQueryUserStandardPrinterDialog
51460   cmbPrintAfterSavingQueryUser.List(2) = .OptionsPrintAfterSavingQueryUserPrinterSetupDialog
51470   cmbPrintAfterSavingQueryUser.List(3) = .OptionsPrintAfterSavingQueryUserDefaultPrinter
51480
51490   chkPrintAfterSavingNoCancel.Caption = .OptionsPrintAfterSavingNoCancel
51500   chkPrintAfterSavingDuplex.Caption = .OptionsPrintAfterSavingDuplex
51510   cmbPrintAfterSavingTumble.List(0) = .OptionsPrintAfterSavingDuplexTumbleOff
51520   cmbPrintAfterSavingTumble.List(1) = .OptionsPrintAfterSavingDuplexTumbleOn
51530
51540   Me.Caption = .DialogPrinterOptions
51550   cmdCancel.Caption = .OptionsCancel
51560   cmdReset.Caption = .OptionsReset
51570   cmdSave.Caption = .OptionsSave
51580   tbstrPDFOptions.Tabs(1).Caption = .OptionsPDFGeneral
51590   tbstrPDFOptions.Tabs(2).Caption = .OptionsPDFCompression
51600   tbstrPDFOptions.Tabs(3).Caption = .OptionsPDFFonts
51610   tbstrPDFOptions.Tabs(4).Caption = .OptionsPDFColors
51620   tbstrPDFOptions.Tabs(5).Caption = .OptionsPDFSecurity
51630   dmFraPDFGeneral.Caption = .OptionsPDFGeneralCaption
51640   chkPDFOptimize.Caption = .OptionsPDFOptimize
51650   lblPDFCompat.Caption = .OptionsPDFGeneralCompatibility
51660   lblPDFAutoRotate.Caption = .OptionsPDFGeneralAutorotate
51670   lblPDFResolution.Caption = .OptionsPDFGeneralResolution
51680   lblPDFOverprint.Caption = .OptionsPDFGeneralOverprint
51690   lblProgfont.Caption = .OptionsProgramFont
51700   lblProgcharset.Caption = .OptionsProgramFontcharset
51710   lblSize.Caption = .OptionsProgramFontSize
51720   lblTesttext = .OptionsProgramFontTestdescription
51730   cmdTest.Caption = .OptionsProgramFontTest
51740   cmdCancelTest.Caption = .OptionsProgramFontCancelTest
51750   chkPDFASCII85.Caption = .OptionsPDFGeneralASCII85
51760   cmbPDFCompat.List(0) = .OptionsPDFGeneralCompatibility01
51770   cmbPDFCompat.List(1) = .OptionsPDFGeneralCompatibility02
51780   cmbPDFCompat.List(2) = .OptionsPDFGeneralCompatibility03
51790   cmbPDFRotate.List(0) = .OptionsPDFGeneralRotate01
51800   cmbPDFRotate.List(1) = .OptionsPDFGeneralRotate02
51810   cmbPDFRotate.List(2) = .OptionsPDFGeneralRotate03
51820   cmbPDFOverprint.List(0) = .OptionsPDFGeneralOverprint01
51830   cmbPDFOverprint.List(1) = .OptionsPDFGeneralOverprint02
51840
51850   dmfraPDFCompress.Caption = .OptionsPDFCompressionCaption
51860   chkPDFTextComp.Caption = .OptionsPDFCompressionTextComp
51870   dmFraPDFColor.Caption = .OptionsPDFCompressionColor
51880   chkPDFColorComp.Caption = .OptionsPDFCompressionColorComp
51890   chkPDFColorResample.Caption = .OptionsPDFCompressionColorResample
51900   lblPDFColorRes.Caption = .OptionsPDFCompressionColorRes
51910   cmbPDFColorComp.List(0) = .OptionsPDFCompressionColorComp01
51920   cmbPDFColorComp.List(1) = .OptionsPDFCompressionColorComp02
51930   cmbPDFColorComp.List(2) = .OptionsPDFCompressionColorComp03
51940   cmbPDFColorComp.List(3) = .OptionsPDFCompressionColorComp04
51950   cmbPDFColorComp.List(4) = .OptionsPDFCompressionColorComp05
51960   cmbPDFColorComp.List(5) = .OptionsPDFCompressionColorComp06
51970   cmbPDFColorComp.List(6) = .OptionsPDFCompressionColorComp07
51980 '  cmbPDFColorComp.List(7) = .OptionsPDFCompressionColorComp08
51990   cmbPDFColorResample.List(0) = .OptionsPDFCompressionColorResample01
52000   cmbPDFColorResample.List(1) = .OptionsPDFCompressionColorResample02
52010 '  cmbPDFColorResample.List(2) = .OptionsPDFCompressionColorResample03
52020   dmFraPDFGrey.Caption = .OptionsPDFCompressionGrey
52030   chkPDFGreyComp.Caption = .OptionsPDFCompressionGreyComp
52040   chkPDFGreyResample.Caption = .OptionsPDFCompressionGreyResample
52050   lblPDFGreyRes.Caption = .OptionsPDFCompressionGreyRes
52060   cmbPDFGreyComp.List(0) = .OptionsPDFCompressionGreyComp01
52070   cmbPDFGreyComp.List(1) = .OptionsPDFCompressionGreyComp02
52080   cmbPDFGreyComp.List(2) = .OptionsPDFCompressionGreyComp03
52090   cmbPDFGreyComp.List(3) = .OptionsPDFCompressionGreyComp04
52100   cmbPDFGreyComp.List(4) = .OptionsPDFCompressionGreyComp05
52110   cmbPDFGreyComp.List(5) = .OptionsPDFCompressionGreyComp06
52120   cmbPDFGreyComp.List(6) = .OptionsPDFCompressionGreyComp07
52130 '  cmbPDFGreyComp.List(7) = .OptionsPDFCompressionGreyComp08
52140   cmbPDFGreyResample.List(0) = .OptionsPDFCompressionGreyResample01
52150   cmbPDFGreyResample.List(1) = .OptionsPDFCompressionGreyResample02
52160 '  cmbPDFGreyResample.List(2) = .OptionsPDFCompressionGreyResample03
52170   dmFraPDFMono.Caption = .OptionsPDFCompressionMono
52180   chkPDFMonoComp.Caption = .OptionsPDFCompressionMonoComp
52190   chkPDFMonoResample.Caption = .OptionsPDFCompressionMonoResample
52200   lblPDFMonoRes.Caption = .OptionsPDFCompressionMonoRes
52210   cmbPDFMonoComp.List(0) = .OptionsPDFCompressionMonoComp01
52220   cmbPDFMonoComp.List(1) = .OptionsPDFCompressionMonoComp02
52230   cmbPDFMonoComp.List(2) = .OptionsPDFCompressionMonoComp03
52240 '  cmbPDFMonoComp.List(3) = .OptionsPDFCompressionMonoComp04
52250   cmbPDFMonoResample.List(0) = .OptionsPDFCompressionMonoResample01
52260   cmbPDFMonoResample.List(1) = .OptionsPDFCompressionMonoResample02
52270 '  cmbPDFMonoResample.List(2) = .OptionsPDFCompressionMonoResample03
52280
52290   dmFraPDFFonts.Caption = .OptionsPDFFontsCaption
52300   chkPDFEmbedAll.Caption = .OptionsPDFFontsEmbedAll
52310   chkPDFSubSetFonts.Caption = .OptionsPDFFontsSubSetFonts
52320
52330   dmFraPDFColors.Caption = .OptionsPDFColorsCaption
52340   chkPDFCMYKtoRGB.Caption = .OptionsPDFColorsCMYKtoRGB
52350   dmFraPDFColorOptions.Caption = .OptionsPDFColorsColorOptions
52360   chkPDFPreserveOverprint.Caption = .OptionsPDFColorsPreserveOverprint
52370   chkPDFPreserveTransfer.Caption = .OptionsPDFColorsPreserveTransfer
52380   chkPDFPreserveHalftone.Caption = .OptionsPDFColorsPreserveHalftone
52390   cmbPDFColorModel.List(0) = .OptionsPDFColorsColorModel01
52400   cmbPDFColorModel.List(1) = .OptionsPDFColorsColorModel02
52410   cmbPDFColorModel.List(2) = .OptionsPDFColorsColorModel03
52420
52430   dmFraPDFEncryptor.Caption = .OptionsPDFEncryptor
52440   dmFraPDFSecurity.Caption = .OptionsPDFSecurityCaption
52450   chkUseSecurity.Caption = .OptionsPDFUseSecurity
52460   dmFraPDFEncLevel.Caption = .OptionsPDFEncryptionLevel
52470   optEncHigh.Caption = .OptionsPDFEncryptionHigh
52480   optEncLow.Caption = .OptionsPDFEncryptionLow
52490   dmFraSecurityPass.Caption = .OptionsPDFPasswords
52500   chkUserPass.Caption = .OptionsPDFUserPass
52510   chkOwnerPass.Caption = .OptionsPDFOwnerPass
52520   dmFraPDFPermissions.Caption = .OptionsPDFDisallowUser
52530   dmFraPDFHighPermissions.Caption = .OptionsPDFEnhancedPermissions
52540   chkAllowPrinting.Caption = .OptionsPDFDisallowPrint
52550   chkAllowModifyContents.Caption = .OptionsPDFDisallowModify
52560   chkAllowCopy.Caption = .OptionsPDFDisallowCopy
52570   chkAllowModifyAnnotations.Caption = .OptionsPDFDisallowModifyComments
52580   chkAllowDegradedPrinting.Caption = .OptionsPDFAllowDegradedPrinting
52590   chkAllowFillIn.Caption = .OptionsPDFAllowFillIn
52600   chkAllowAssembly.Caption = .OptionsPDFAllowAssembly
52610   chkAllowScreenReaders.Caption = .OptionsPDFAllowScreenReaders
52620
52630   cmbPNGColors.List(0) = .OptionsPNGColorscount01
52640   cmbPNGColors.List(1) = .OptionsPNGColorscount02
52650   cmbPNGColors.List(2) = .OptionsPNGColorscount03
52660   cmbPNGColors.List(3) = .OptionsPNGColorscount04
52670   cmbJPEGColors.List(0) = .OptionsJPEGColorscount01
52680   cmbJPEGColors.List(1) = .OptionsJPEGColorscount02
52690   cmbBMPColors.List(0) = .OptionsBMPColorscount01
52700   cmbBMPColors.List(1) = .OptionsBMPColorscount02
52710   cmbBMPColors.List(2) = .OptionsBMPColorscount03
52720   cmbBMPColors.List(3) = .OptionsBMPColorscount04
52730   cmbBMPColors.List(4) = .OptionsBMPColorscount05
52740   cmbBMPColors.List(5) = .OptionsBMPColorscount06
52750   cmbBMPColors.List(6) = .OptionsBMPColorscount07
52760   cmbPCXColors.List(0) = .OptionsPCXColorscount01
52770   cmbPCXColors.List(1) = .OptionsPCXColorscount02
52780   cmbPCXColors.List(2) = .OptionsPCXColorscount03
52790   cmbPCXColors.List(3) = .OptionsPCXColorscount04
52800   cmbPCXColors.List(4) = .OptionsPCXColorscount05
52810   cmbPCXColors.List(5) = .OptionsPCXColorscount06
52820   cmbTIFFColors.List(0) = .OptionsTIFFColorscount01
52830   cmbTIFFColors.List(1) = .OptionsTIFFColorscount02
52840   cmbTIFFColors.List(2) = .OptionsTIFFColorscount03
52850   cmbTIFFColors.List(3) = .OptionsTIFFColorscount04
52860   cmbTIFFColors.List(4) = .OptionsTIFFColorscount05
52870   cmbTIFFColors.List(5) = .OptionsTIFFColorscount06
52880   cmbTIFFColors.List(6) = .OptionsTIFFColorscount07
52890   cmbTIFFColors.List(7) = .OptionsTIFFColorscount08
52900
52910   dmFraBitmapGeneral.Caption = .OptionsImageSettings
52920   lblBitmapResolution = .OptionsBitmapResolution
52930   lblJPEGQuality = .OptionsJPEGQuality
52940   lblBitmapColors = .OptionsPDFColors
52950   lblProcessPriority.Caption = .OptionsProcesspriority
52960   lblLangLevel.Caption = .OptionsPSLanguageLevel
52970
52980   cmdAsso.Caption = .OptionsAssociatePSFiles
52990
53000   lblStampString.Caption = .OptionsStampString
53010   lblStampFontcolor.Caption = .OptionsStampFontColor
53020   chkStampUseOutlineFont.Caption = .OptionsStampUseOutlineFont
53030   lblOutlineFontThickness.Caption = .OptionsStampOutlineFontThickness
53040
53050   chkUseFixPaperSize.Caption = .OptionsUseFixPapersize
53060   chkUseCustomPapersize.Caption = .OptionsUseCustomPapersize
53070   lblCustomPapersizeWidth.Caption = .OptionsCustomPapersizeWidth
53080   lblCustomPapersizeHeight.Caption = .OptionsCustomPapersizeHeight
53090   lblCustomPapersizeInfo.Caption = .OptionsCustomPapersizeInfo
53100  End With
53110
53120  With sldProcessPriority
53131   Select Case .Value
         Case 0: 'Idle
53150     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityIdle
53160    Case 1: 'Normal
53170     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityNormal
53180    Case 2: 'High
53190     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityHigh
53200    Case 3: 'Realtime
53210     lblProcessPriority.Caption = LanguageStrings.OptionsProcesspriority & ": " & LanguageStrings.OptionsProcesspriorityRealtime
53220   End Select
53230  End With
53240  If ShowOnlyOptions = True Then
53250   Caption = "PDFCreator - " & Caption
53260  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "ChangeLanguage")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ReadAllLanguages(LanguagePath As String, Optional UserPath As Boolean = False)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Languagename As String, ini As clsINI, UserLangFiles As Collection, _
  i As Long, j As Long, found As Boolean, Version As String, Filename As String, _
  UserLanguagePath As String
50040
50050  cmbCurrentLanguage.Clear
50060  cmbCurrentLanguage.AddItem "No languages available."
50070  Set LangFiles = GetAllLanguagesFiles(LanguagePath)
50080
50090  If UserPath Then
50100   UserLanguagePath = GetMyAppData() & "\PDFCreator\Languages"
50110   Set UserLangFiles = GetAllLanguagesFiles(UserLanguagePath)
50120   For i = 1 To UserLangFiles.Count
50130    For j = 1 To LangFiles.Count
50140     If GetFilenameFromPath(LangFiles(j)) = GetFilenameFromPath(UserLangFiles(i)) Then
50150      LangFiles.Remove j
50160      Exit For
50170     End If
50180    Next j
50190    LangFiles.Add UserLangFiles(i)
50200   Next i
50210  End If
50220
50230  Set Languages = New Collection
50240  For i = 1 To LangFiles.Count
50250   SplitPath LangFiles(i), , , , Filename
50260   Languages.Add Filename
50270  Next i
50280
50290  Set ini = New clsINI
50300  For i = 1 To LangFiles.Count
50310   ini.Filename = LangFiles.Item(i)
50320   ini.Section = "Common"
50330   Languagename = ini.GetKeyFromSection("Languagename")
50340   Version = ini.GetKeyFromSection("Version")
50350   If Len(Languagename) = 0 Then
50360    Languagename = "No name available."
50370   End If
50380   If IsCompatibleLanguageVersion(Version) = True Then
50390     If i = 1 Then
50400       cmbCurrentLanguage.List(0) = Languagename
50410      Else
50420       cmbCurrentLanguage.AddItem Languagename
50430     End If
50440    Else
50450     If i = 1 Then
50460       cmbCurrentLanguage.List(0) = Languagename & " [" & Version & "]"
50470      Else
50480       cmbCurrentLanguage.AddItem Languagename & " [" & Version & "]"
50490     End If
50500   End If
50510 '  cmbCurrentLanguage.ItemData(cmbCurrentLanguage.ListCount - 1) = LangFiles.Item(i)
50520   SplitPath LangFiles.Item(i), , , , Filename
50530   If UCase$(Options.Language) = UCase$(Filename) Then
50540    cmbCurrentLanguage.ListIndex = i - 1
50550   End If
50560   DoEvents
50570  Next i
50580  Set ini = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "ReadAllLanguages")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function GetAllLanguagesFiles(LanguagePath As String) As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tColl1 As Collection, tColl2 As Collection, i As Long, tStrf() As String, ini As clsINI, _
  Languagename As String
50030  Set GetAllLanguagesFiles = New Collection
50040  Set tColl1 = GetFiles(LanguagePath, "*.ini", SortedByName)
50050  Set tColl2 = New Collection
50060  For i = 1 To tColl1.Count
50070   tStrf = Split(tColl1(i), "|")
50080   Set ini = New clsINI
50090   ini.Filename = tStrf(1)
50100   ini.Section = "Common"
50110   Languagename = ini.GetKeyFromSection("Languagename")
50120   If Len(Languagename) = 0 Then
50130    Languagename = "No name available."
50140   End If
50150   AddSortedStr tColl2, Languagename & "|" & tStrf(1)
50160  Next i
50170  For i = 1 To tColl2.Count
50180   tStrf() = Split(tColl2(i), "|")
50190   GetAllLanguagesFiles.Add tStrf(1)
50200  Next i
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "GetAllLanguagesFiles")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function IsCompatibleLanguageVersion(Version As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Byte, delim As String, fVers() As String, fCVers() As String, _
  ProgVersion As String, fPVers() As String
50030  IsCompatibleLanguageVersion = False
50040  delim = "."
50050  ProgVersion = GetProgramRelease
50060  If Len(CompatibleLanguageVersion) = 0 Or Len(Version) = 0 Or Len(ProgVersion) = 0 Then
50070   Exit Function
50080  End If
50090  If InStr(1, CompatibleLanguageVersion, delim) = 0 Or _
    InStr(1, Version, delim) = 0 Or _
    InStr(1, ProgVersion, delim) = 0 Then
50120   Exit Function
50130  End If
50140  fVers = Split(Version, delim)
50150  fCVers = Split(CompatibleLanguageVersion, delim)
50160  fPVers = Split(ProgVersion, delim)
50170  If UBound(fVers) < 2 Or UBound(fCVers) < 2 Or UBound(fPVers) < 2 Then
50180   Exit Function
50190  End If
50200  For i = 0 To 2
50210   If IsNumeric(fVers(i)) = False Or IsNumeric(fCVers(i)) = False Or _
   IsNumeric(fPVers(i)) = False Then
50230    Exit Function
50240   End If
50250  Next i
50260  If CLng(fVers(0)) >= CLng(fCVers(0)) And CLng(fVers(0)) <= CLng(fPVers(0)) Then
50270   If CLng(fVers(1)) >= CLng(fCVers(1)) And CLng(fVers(1)) <= CLng(fPVers(1)) Then
50280    If CLng(fVers(2)) >= CLng(fCVers(2)) And CLng(fVers(2)) <= CLng(fPVers(2)) Then
50290     IsCompatibleLanguageVersion = True
50300    End If
50310   End If
50320  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "IsCompatibleLanguageVersion")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function InstallInternetLanguageFile(File As String, Version As String, DownloadURL As String, ProgramPath As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim strLangFile As String, strFile As String
50020  InstallInternetLanguageFile = False
50030  If (File = vbNullString) Or (Version = vbNullString) Or (DownloadURL = vbNullString) Or (ProgramPath = vbNullString) Then
50040   Exit Function
50050  End If
50060  ProgramPath = CompletePath(ProgramPath)
50070  If Right$(DownloadURL, 1) <> "/" Then
50080   DownloadURL = DownloadURL & "/"
50090  End If
50100
50110  Set dl = New clsDownload
50120  strLangFile = dl.DownloadString(DownloadURL & Version & "/" & File)
50130  Set dl = Nothing
50140
50150  If InStr(1, strLangFile, "[Common]", vbTextCompare) = 0 Then
50160   MsgBox LanguageStrings.MessagesMsg37, vbCritical
50170   Exit Function
50180  End If
50190
50200  If Not DirExists(ProgramPath) Then
50210   MsgBox LanguageStrings.MessagesMsg10, vbCritical
50220   Exit Function
50230  End If
50240
50250  strFile = ProgramPath & File
50260
50270  If FileExists(strFile) Then
50280   If MsgBox(LanguageStrings.MessagesMsg05, vbYesNo) = vbNo Then
50290    Exit Function
50300   End If
50310  End If
50320
50330  Open strFile For Output As #1
50340  Print #1, strLangFile
50350  Close #1
50360
50370  MsgBox LanguageStrings.MessagesMsg38, vbInformation
50380
50390  InstallInternetLanguageFile = True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmOptions", "InstallInternetLanguageFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

