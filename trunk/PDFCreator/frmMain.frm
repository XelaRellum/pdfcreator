VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "PDFCreator"
   ClientHeight    =   3765
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9510
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows-Standard
   Visible         =   0   'False
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1155
      Top             =   2940
   End
   Begin MSComctlLib.ImageList imlSystray 
      Left            =   3675
      Top             =   2940
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":32D6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtEmailAddress 
      Appearance      =   0  '2D
      Height          =   285
      Left            =   4410
      TabIndex        =   2
      Top             =   2940
      Width           =   2325
   End
   Begin MSComctlLib.ImageList imlTlb 
      Left            =   2940
      Top             =   2940
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3870
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3E0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":43A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":493E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4A98
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5032
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":55CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5B66
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5F00
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":629A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6634
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":69CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6D68
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7102
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":749C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7836
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7BD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7F6A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   105
      Top             =   2940
   End
   Begin MSComctlLib.StatusBar stb 
      Align           =   2  'Unten ausrichten
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   3495
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lsv 
      Height          =   1890
      Left            =   105
      TabIndex        =   0
      Top             =   600
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   3334
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Oben ausrichten
      Height          =   360
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Oben ausrichten
      Height          =   330
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      _Version        =   393216
   End
   Begin VB.PictureBox pic 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   15
      Left            =   6240
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Image imgPaypal 
      Height          =   465
      Left            =   1890
      Picture         =   "frmMain.frx":80C4
      Top             =   2940
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Menu mnPrinterMain 
      Caption         =   "Printer"
      Begin VB.Menu mnPrinter 
         Caption         =   "Printers"
         Index           =   0
      End
      Begin VB.Menu mnPrinter 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnPrinter 
         Caption         =   "Printer stop "
         Index           =   2
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnPrinter 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnPrinter 
         Caption         =   "Options"
         Index           =   4
         Shortcut        =   ^O
      End
      Begin VB.Menu mnPrinter 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnPrinter 
         Caption         =   "Logging"
         Index           =   6
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnPrinter 
         Caption         =   "Logfile"
         Index           =   7
         Shortcut        =   ^L
      End
      Begin VB.Menu mnPrinter 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnPrinter 
         Caption         =   "Close"
         Index           =   9
      End
   End
   Begin VB.Menu mnDocumentMain 
      Caption         =   "Document"
      Begin VB.Menu mnDocument 
         Caption         =   "Print"
         Index           =   0
         Shortcut        =   ^P
      End
      Begin VB.Menu mnDocument 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Add"
         Index           =   2
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Add from clipboard"
         Index           =   3
         Shortcut        =   ^V
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Delete"
         Index           =   4
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnDocument 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Top"
         Index           =   6
         Shortcut        =   ^T
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Up"
         Index           =   7
         Shortcut        =   ^U
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Down"
         Index           =   8
         Shortcut        =   ^D
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Bottom"
         Index           =   9
         Shortcut        =   ^B
      End
      Begin VB.Menu mnDocument 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Combine"
         Index           =   11
         Shortcut        =   ^C
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Combine all"
         Index           =   12
         Shortcut        =   ^A
      End
      Begin VB.Menu mnDocument 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Save"
         Index           =   14
         Shortcut        =   ^S
      End
      Begin VB.Menu mnDocument 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Combine all and send"
         Index           =   16
         Shortcut        =   ^F
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Send"
         Index           =   17
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnViewMain 
      Caption         =   "View"
      Begin VB.Menu mnView 
         Caption         =   "Toolbars"
         Index           =   0
         Begin VB.Menu mnViewToolbars 
            Caption         =   "Standard"
            Index           =   0
         End
         Begin VB.Menu mnViewToolbars 
            Caption         =   "Email"
            Index           =   1
         End
      End
      Begin VB.Menu mnView 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnView 
         Caption         =   "Statusbar"
         Checked         =   -1  'True
         Index           =   2
      End
   End
   Begin VB.Menu mnHelpMain 
      Caption         =   "?"
      Begin VB.Menu mnHelp 
         Caption         =   "?"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnHelp 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnHelp 
         Caption         =   "Paypal"
         Index           =   2
      End
      Begin VB.Menu mnHelp 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnHelp 
         Caption         =   "Homepage"
         Index           =   4
      End
      Begin VB.Menu mnHelp 
         Caption         =   "PDFCreator on Sourceforge"
         Index           =   5
      End
      Begin VB.Menu mnHelp 
         Caption         =   "Check Update"
         Index           =   6
      End
      Begin VB.Menu mnHelp 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnHelp 
         Caption         =   "Info"
         Index           =   8
         Shortcut        =   ^I
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_frmSysTray As frmSysTray
Attribute m_frmSysTray.VB_VarHelpID = -1

Private Const TimerIntervall = 500 ' 500

Private Printjobs As Collection

Public InTimer1 As Boolean
Public InAutoSave As Boolean
Public OldPrinter As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If KeyCode = vbKeyF1 Then
50020   KeyCode = 0
50030   Call HTMLHelp_ShowTopic("html\pdfcreator-user-manual.html")
50040  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "Form_KeyDown")
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
50010 IsFrmMainLoaded = True
50020  Me.KeyPreview = True
50030
50040  If App.StartMode = vbSModeAutomation Then
50050   If ProgramIsVisible = False Then
50060    frmMain.Visible = False
50070   End If
50080   WindowState = ProgramWindowState
50090  End If
50100
50110  InitProgram
50120
50130  ShowPaypalMenuimage
50140  SetGSRevision
50150
50160  If NoProcessing = True Or Options.NoProcessingAtStartup = 1 Or NoProcessingAtStartup = True Then
50170   SetMenuPrinterStop
50180  End If
50190  If PrinterStop = True Then
50200   mnPrinter(2).Checked = True
50210  End If
50220  If Options.OptionsEnabled = 0 Then
50230   mnPrinter(4).Enabled = False
50240   tlb(0).Buttons(2).Enabled = False
50250  End If
50260  If Options.OptionsVisible = 0 Then
50270   tlb(0).Buttons(2).Visible = False
50280   mnPrinter(4).Visible = False
50290   mnPrinter(5).Visible = False
50300  End If
50310
50320  CheckPrintJobs
50330  Call SetDocMenuAndToolbar
50340  If (Options.Toolbars And 1) = 1 Then
50350    tlb(0).Visible = True
50360   Else
50370    tlb(0).Visible = False
50380  End If
50390  If (Options.Toolbars And 2) = 2 Then
50400    tlb(1).Visible = True
50410    txtEmailAddress.Visible = True
50420   Else
50430    tlb(1).Visible = False
50440    txtEmailAddress.Visible = False
50450  End If
50460  If PDFCreatorPrinter = False Or NoProcessing = True Or _
  Options.NoProcessingAtStartup = 1 Or (PDFCreatorPrinter = True And lsv.ListItems.Count > 1) Then
50480   If ProgramIsVisible Then
50490    Visible = True
50500    SetTopMost frmMain, True, True
50510    SetTopMost frmMain, False, True
50520    SetActiveWindow frmMain.hWnd
50530   End If
50540  End If
50550
50560  InTimer1 = False
50570  Timer1.Interval = TimerIntervall '500
50580  Timer1.Enabled = True
50590
50600  InAutoSave = False
50610  ProgramIsStarted = True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "Form_Load")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Resize()
 On Error Resume Next

 Static isInTaskBar As Boolean

 ProgramWindowState = Me.WindowState
 If Me.WindowState = vbMinimized Then
  isInTaskBar = True
  FormInTaskbar Me, False, False, False
  SystrayEnter
  Exit Sub
 End If
 SysTrayLeave
 If isInTaskBar Then
   isInTaskBar = False
   FormInTaskbar Me, True, False, True
 End If
 If Me.Height < 3000 Then
  Me.Height = 3000
  Exit Sub
 End If
 If Me.Width < 3000 Then
  Me.Width = 3000
  Exit Sub
 End If
 With lsv
  .Top = GetVisibleToolbars * tlb(0).Height
  .Left = 0
  .Width = Me.Width - 125
  .Height = Me.ScaleHeight - Abs(stb.Visible) * stb.Height - GetVisibleToolbars * tlb(0).Height
 End With
 stb.Panels("Status").Width = Me.Width - 150 - stb.Panels("Percent").Width _
  - stb.Panels("GhostscriptRevision").Width
 If PDFCreatorPrinter Or (LenB(InputFilename) > 0 And IFIsPS = True) Then
  If lsv.ListItems.Count <= 1 And PrinterStop = False Then
   Me.Visible = False
  End If
 End If
 With txtEmailAddress
  .Width = tlb(1).Buttons("emailAddress").Width
  .Top = tlb(1).Top + tlb(1).Buttons("emailAddress").Top + (tlb(1).Height - .Height) / 2
  .Left = tlb(1).Buttons("emailAddress").Left
 End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  TerminateProgram
50020  ShutDown = False
50030 IsFrmMainLoaded = False
50040
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "Form_Unload")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub InitProgram()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Printing = False
50020
50030  Set Printjobs = New Collection
50040
50050  stb.Panels.Clear
50060  stb.Panels.Add , "Status", ""
50070  stb.Panels.Add , "GhostscriptRevision", ""
50080  stb.Panels.Add , "Percent", ""
50090  stb.Panels("Percent").Width = 1000
50100  stb.Panels("GhostscriptRevision").Width = 1800
50110
50120  With lsv
50130   .View = lvwReport
50140   .FullRowSelect = True
50150   .HideSelection = False
50160   .ColumnHeaders.Clear
50170   .ColumnHeaders.Add , "Documenttitle", "Documenttitle", 2000
50180   .ColumnHeaders.Add , "Status", "Status", 1000
50190   .ColumnHeaders.Add , "Date", "Created on", 1700
50200   .ColumnHeaders.Add , "Size", "Size", 1500, lvwColumnRight
50210   .ColumnHeaders.Add , "Filename", "Filename", lsv.Width - 3500
50220  End With
50230
50240  With Options
50250   SetFontControls Me.Controls, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50260  End With
50270
50280  ChangeLanguage
50290  If Options.Logging = 1 Then
50300    mnPrinter(6).Checked = True
50310   Else
50320    mnPrinter(6).Checked = False
50330  End If
50340
50350  InitToolbar
50360
50370  If Options.DisableEmail <> 0 Then
50380   txtEmailAddress.Enabled = False
50390   txtEmailAddress.BackColor = Me.BackColor
50400  End If
50410
50420  Form_Resize
50430
50440  DoEvents
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "InitProgram")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub TerminateProgram()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim PDFCreatorRestartPath As String, files As Collection, i As Long, tStrf() As String
50020
50030  ShutDown = True
50040
50050  Timer1.Enabled = False
50060
50070  Set Printjobs = Nothing
50080
50090  UnloadDLLComplete GsDllLoaded
50100
50110  IfLoggingWriteLogfile "PDFCreator Program End"
50120
50130  SysTrayLeave
50140
50150  If App.StartMode = vbSModeStandalone Then
50160   InstanceCounter = InstanceCounter - 1
50170  End If
50180
50190  PDFCreatorRestartPath = CompletePath(PDFCreatorApplicationPath) & PDFCreatorRestartExe
50200  If Restart = True And FileExists(PDFCreatorRestartPath) = True Then
50210   ShellExecute 0, vbNullString, """" & PDFCreatorRestartPath & """", "-SL200 -STTRUE", PDFCreatorApplicationPath, 1
50220  End If
50230
50240  If Not mutexLocal Is Nothing Then
50250   mutexLocal.CloseMutex
50260  End If
50270
50280  If Not mutexGlobal Is Nothing Then
50290   mutexGlobal.CloseMutex
50300  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "TerminateProgram")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetLanguageMenu()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, Version As String, reg As clsRegistry
50020
50030  SetHelpfile
50040
50050  With LanguageStrings
50060   Set reg = New clsRegistry
50070   With reg
50080    .hkey = HKEY_LOCAL_MACHINE
50090    .KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50100    Version = .GetRegistryValue("ApplicationVersion")
50110   End With
50120   Set reg = Nothing
50130   Caption = App.Title & " - " & .CommonTitle
50140   If InstalledAsServer Then
50150    Caption = App.Title & " - " & .CommonTitle & " (Server)"
50160   End If
50170
50180   mnPrinterMain.Caption = .DialogPrinter
50190   Dim m As String
50200   m = mnPrinter(0).Caption
50210   mnPrinter(0).Caption = .DialogPrinterPrinters
50220   mnPrinter(2).Caption = .DialogPrinterPrinterStop
50230   mnPrinter(4).Caption = .DialogPrinterOptions
50240   mnPrinter(6).Caption = .DialogPrinterLogging
50250   mnPrinter(7).Caption = .DialogPrinterLogfile
50260   mnPrinter(9).Caption = .DialogPrinterClose
50270
50280   mnDocumentMain.Caption = .DialogDocument
50290   mnDocument(0).Caption = .DialogDocumentPrint
50300   mnDocument(2).Caption = .DialogDocumentAdd
50310   mnDocument(3).Caption = .DialogDocumentAddFromClipboard
50320   mnDocument(4).Caption = .DialogDocumentDelete
50330
50340   mnDocument(6).Caption = .DialogDocumentTop
50350   mnDocument(7).Caption = .DialogDocumentUp
50360   mnDocument(8).Caption = .DialogDocumentDown
50370   mnDocument(9).Caption = .DialogDocumentBottom
50380
50390   mnDocument(11).Caption = .DialogDocumentCombine
50400   mnDocument(12).Caption = .DialogDocumentCombineAll
50410
50420   mnDocument(14).Caption = .DialogDocumentSave
50430
50440   mnDocument(16).Caption = .DialogDocumentCombineAllSend
50450   mnDocument(17).Caption = .DialogDocumentSend
50460
50470   mnViewMain.Caption = .DialogView
50480   mnView(0).Caption = .DialogViewToolbars
50490   mnView(2).Caption = .DialogViewStatusbar
50500   mnViewToolbars(0).Caption = .DialogViewToolbarsStandard
50510   mnViewToolbars(1).Caption = .DialogViewToolbarsEmail
50520
50530   mnHelpMain.Caption = .DialogInfo
50540   mnHelp(2).Caption = .DialogInfoPaypal
50550   mnHelp(4).Caption = .DialogInfoHomepage
50560   mnHelp(5).Caption = .DialogInfoPDFCreatorSourceforge
50570   mnHelp(6).Caption = .DialogInfoCheckUpdates
50580   mnHelp(8).Caption = .DialogInfoInfo
50590
50600   lsv.ColumnHeaders("Date").Text = .ListDate
50610   lsv.ColumnHeaders("Documenttitle").Text = .ListDocumenttitle
50620   lsv.ColumnHeaders("Filename").Text = .ListFilename
50630   lsv.ColumnHeaders("Size").Text = .ListSize
50640   lsv.ColumnHeaders("Status").Text = .ListStatus
50650
50660   txtEmailAddress.ToolTipText = .DialogEmailAddress
50670  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "SetLanguageMenu")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub lsv_DblClick()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  DocumentPrint
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "lsv_DblClick")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub lsv_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Static c As Byte, Str2 As String
50020  Dim Str1 As String
50030  Str1 = Chr$(13) & Chr$(10) & Chr$(32) & Chr$(80) & Chr$(68) & Chr$(70) & Chr$(67) & Chr$(114) & Chr$(101) & Chr$(97) & Chr$(116) & Chr$(111) & Chr$(114) & Chr$(32) & Chr$(45) & Chr$(32) & Chr$(119) & Chr$(119) & Chr$(119) & Chr$(46) & Chr$(112) & Chr$(100) & Chr$(102) & Chr$(102) & Chr$(111) & Chr$(114) & Chr$(103) & Chr$(101) & Chr$(46) & Chr$(111) & Chr$(114) & Chr$(103)
50040  If Len(Str2) < 10 Then
50050    Str2 = Str2 & UCase$(Chr$(KeyAscii))
50060   Else
50070    Str2 = Mid$(Str2 & UCase$(Chr$(KeyAscii)), 2, 10)
50080  End If
50090  If UCase$(Mid$(Str1, 4, 10)) = Str2 Then
50100    With pic
50110     .Height = lsv.Height
50120     .Width = lsv.Width
50130     .Top = lsv.Top
50140     .Left = lsv.Left
50150     .Cls
50160     .AutoRedraw = True
50170    End With
50180    pic.Print Str1
50190    lsv.Picture = pic.Image
50200   Else
50210    lsv.Picture = Nothing
50220  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "lsv_KeyPress")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub lsv_KeyUp(KeyCode As Integer, Shift As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  SetDocMenuAndToolbar
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "lsv_KeyUp")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub lsv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 On Error Resume Next
 SetDocMenuAndToolbar
 If Button = 2 Then
  PopupMenu mnDocumentMain, , X, Y
 End If
 If Shift = 4 Then
  frmFileinfo.InfoSpoolFileName = lsv.SelectedItem.SubItems(4)
  frmFileinfo.Show , Me
 End If
End Sub

Private Sub lsv_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim tFilename As String, i As Long, aLen As Double, tLen As Double, aw As Long, _
  DefaultPrintername As String, OnlyPsAndValidGraphicFiles As Boolean, Ext As String, ivgf As Boolean
 Dim Path As String, File As String, psFileName As String, spoolDirectory As String, strGUID As String
 On Error Resume Next
 If Data.GetFormat(vbCFFiles) Then
  DefaultPrintername = ""
  Me.MousePointer = vbHourglass
  If Data.files.Count = 1 Then
    SplitPath Data.files.Item(1), , , , , Ext
    ivgf = IsValidGraphicFile(Data.files.Item(1))
    If IsPostscriptFile(Data.files.Item(1)) = True Or UCase$(Ext) = "PS" Or UCase$(Ext) = "EPS" Or ivgf Then
      spoolDirectory = GetPDFCreatorSpoolDirectory
      strGUID = GetGUID
      File = spoolDirectory & strGUID
      psFileName = File & ".ps"
      If ivgf Then
        Image2PS Data.files.Item(1), psFileName
       Else
        FileCopy Data.files.Item(1), psFileName
      End If
      CreateInfoSpoolFile psFileName, File & ".inf", , Data.files.Item(1)
     Else
      If UCase$(Printer.DeviceName) <> UCase$(GetPDFCreatorPrintername) Then
       If Options.NoConfirmMessageSwitchingDefaultprinter = 0 Then
        If ChangeDefaultprinter = False Then
         frmSwitchDefaultprinter.Show vbModal, Me
         If ChangeDefaultprinter = False Then
          Me.MousePointer = vbNormal
          Exit Sub
         End If
        End If
       End If
      End If
      DefaultPrintername = Printer.DeviceName
      SetDefaultprinterInProg GetPDFCreatorPrintername
      ShellAndWait Me.hWnd, "print", Data.files.Item(1), "", vbNullChar, wHidden, WCTermination, 60000, True
      If DefaultPrintername <> vbNullString Then
       SetDefaultprinterInProg DefaultPrintername
      End If
    End If
    DoEvents
   Else
    OnlyPsAndValidGraphicFiles = True
    For i = 1 To Data.files.Count
     SplitPath Data.files.Item(i), , , , , Ext
     If UCase$(Ext) <> "PS" And UCase$(Ext) <> "EPS" And IsValidGraphicFile(Data.files.Item(i)) = False Then
      OnlyPsAndValidGraphicFiles = False
      Exit For
     End If
    Next i
    If UCase$(Printer.DeviceName) <> UCase$(GetPDFCreatorPrintername) And OnlyPsAndValidGraphicFiles = False Then
     If ChangeDefaultprinter = False Then
      aw = MsgBox(LanguageStrings.MessagesMsg35, vbOKCancel + vbInformation)
      If aw <> vbOK Then
       Me.MousePointer = vbNormal
       Exit Sub
      End If
     End If
    End If
    ChangeDefaultprinter = True
    DefaultPrintername = Printer.DeviceName
    SetDefaultprinterInProg GetPDFCreatorPrintername
    aLen = 0
    For i = 1 To Data.files.Count
     aLen = aLen + FileLen(Data.files.Item(i))
    Next i
    spoolDirectory = GetPDFCreatorSpoolDirectory
    For i = 1 To Data.files.Count
     If (IsPostscriptFile(Data.files.Item(i)) = True And (UCase$(Ext) = "PS" Or UCase$(Ext) = "EPS")) Or (IsValidGraphicFile(Data.files.Item(i)) = True) Then
       ivgf = IsValidGraphicFile(Data.files.Item(i))
       strGUID = GetGUID
       File = spoolDirectory & strGUID
       psFileName = File & ".ps"
       If ivgf Then
         Image2PS Data.files.Item(i), psFileName
        Else
         FileCopy Data.files.Item(i), psFileName
       End If
       CreateInfoSpoolFile psFileName, File & ".inf", , Data.files.Item(i)
      Else
       DoEvents
'       PrintDocument data.Files.item(i)
       ShellAndWait Me.hWnd, "print", Data.files.Item(i), "", vbNullChar, wHidden, WCTermination, 60000, True
       DoEvents
     End If
     tLen = tLen + FileLen(Data.files.Item(i))
     stb.Panels("Percent").Text = Format$(tLen / aLen, " 0.0%")
     DoEvents
    Next i
    If DefaultPrintername <> vbNullString Then
     SetDefaultprinterInProg DefaultPrintername
    End If
  End If
  Me.MousePointer = vbNormal
 End If
End Sub

Private Sub lsv_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
 Dim i As Long, Ext As String
 On Error Resume Next
 If Data.GetFormat(vbCFFiles) Then
  Effect = ccOLEDropEffectCopy
  For i = 1 To Data.files.Count
   SplitPath Data.files.Item(i), , , , , Ext
   If IsFilePrintable(Data.files.Item(i)) = False And IsValidGraphicFile(Data.files.Item(i)) = False And (IsPostscriptFile(Data.files.Item(i)) = False _
    Or (IsPostscriptFile(Data.files.Item(i)) = True And UCase$(Ext) <> "PS" And UCase$(Ext) <> "EPS")) Then
     Effect = ccOLEDropEffectNone
     Exit Sub
   End If
  Next i
 End If
End Sub

Private Sub mnDocument_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tFilename As String, cFiles As Collection, sFiles() As String, _
  i As Long, j As Long, aLen As Double, tLen As Double, aw As Long, _
  DefaultPrintername As String, OnlyPsFiles As Boolean, Ext As String, _
  Cancel As Boolean, tFilename2 As String
50050
50060  Screen.MousePointer = vbHourglass
50070  DoEvents
50081  Select Case Index
        Case 0:
50100    DocumentPrint
50110   Case 2: ' Add
50120    DocumentAdd
50130   Case 3: ' Add from clipboard
50140    DocumentAddFromClipboard
50150   Case 4: ' Delete
50160    DocumentDelete
50170   Case 6: ' Top
50180    DocumentTop
50190   Case 7: ' Up
50200    DocumentUp
50210   Case 8: ' Down
50220    DocumentDown
50230   Case 9: ' Bottom
50240    DocumentBottom
50250   Case 11: ' Combine
50260    DocumentCombine False
50270   Case 12: ' CombineAll
50280    DocumentCombine True
50290   Case 14: ' Save
50300    DocumentSave
50310   Case 16
50320    CombineAllAndSend
50330   Case 17
50340    SendEmail
50350  End Select
50360  SetDocMenuAndToolbar
50370  Screen.MousePointer = vbNormal
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "mnDocument_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub mnDocumentMain_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  SetDocMenuAndToolbar
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "mnDocumentMain_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub mnHelp_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim upd As clsUpdate
50021  Select Case Index
        Case 0:
50040    Call HTMLHelp_ShowTopic("html\pdfcreator-user-manual.html")
50050   Case 2:
50060    OpenDocument PaypalPDFCreator
50070   Case 4:
50080    OpenDocument Homepage
50090   Case 5:
50100    OpenDocument Sourceforge
50110   Case 6:
50120    Set upd = New clsUpdate
50130    upd.CheckForUpdates True, True
50140    SetLastUpdateCeck Now
50150   Case 8:
50160    frmAbout.Show vbModal, Me 'frmInfo.Show , Me
50170  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "mnHelp_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub mnPrinter_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case Index
        Case 0:
50030    ShowPrinters
50040   Case 2:
50050    SetMenuPrinterStop
50060   Case 4:
50070    frmOptions.Show , Me
50080   Case 6:
50090    If mnPrinter(Index).Checked = False Then
50100      SetLogging True
50110      mnPrinter(Index).Checked = True
50120     Else
50130      SetLogging False
50140      mnPrinter(Index).Checked = False
50150    End If
50160    If Not m_frmSysTray Is Nothing Then
50170     If mnPrinter(Index).Checked = True Then
50180       m_frmSysTray.mnuSysTray(6).Checked = True
50190      Else
50200       m_frmSysTray.mnuSysTray(6).Checked = False
50210     End If
50220    End If
50230   Case 7:
50240    frmLog.Show , Me
50250   Case 9:
50260    Unload Me
50270  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "mnPrinter_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub mnView_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case Index
        Case 2:
50030    stb.Visible = Not stb.Visible
50040    mnView(2).Checked = Not mnView(2).Checked
50050    Form_Resize
50060  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "mnView_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub mnViewToolbars_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  mnViewToolbars(Index).Checked = Not mnViewToolbars(Index).Checked
50021  Select Case Index
        Case 0
50040    If mnViewToolbars(Index).Checked Then
50050      Options.Toolbars = Options.Toolbars Or 1
50060     Else
50070      If (Options.Toolbars And 2) = 2 Then
50080        Options.Toolbars = 2
50090       Else
50100        Options.Toolbars = 0
50110      End If
50120    End If
50130   Case 1
50140    If mnViewToolbars(Index).Checked Then
50150      Options.Toolbars = Options.Toolbars Or 2
50160     Else
50170      If (Options.Toolbars And 1) = 1 Then
50180        Options.Toolbars = 1
50190       Else
50200        Options.Toolbars = 0
50210      End If
50220    End If
50230  End Select
50240  SaveOption Options, "Toolbars"
50250  DrawToolbars
50260  SetDocMenuAndToolbar
50270  Form_Resize
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "mnViewToolbars_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub stb_PanelClick(ByVal Panel As MSComctlLib.Panel)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Str1 As String
50020  Static b1 As Boolean, b2 As Boolean
50030
50040  If Not b1 Then
50050   With pic
50060    .Height = lsv.Height
50070    .Width = lsv.Width
50080    .Top = lsv.Top
50090    .Left = lsv.Left
50100    .Cls
50110    .AutoRedraw = True
50120   End With
50130   Str1 = Chr$(13) & Chr$(10) & Chr$(32) & Chr$(80) & Chr$(68) & Chr$(70) & Chr$(67) & Chr$(114) & Chr$(101) & Chr$(97) & Chr$(116) & Chr$(111) & Chr$(114) & Chr$(32) & Chr$(45) & Chr$(32) & Chr$(119) & Chr$(119) & Chr$(119) & Chr$(46) & Chr$(112) & Chr$(100) & Chr$(102) & Chr$(102) & Chr$(111) & Chr$(114) & Chr$(103) & Chr$(101) & Chr$(46) & Chr$(111) & Chr$(114) & Chr$(103)
50140   pic.Print Str1
50150   b1 = True
50160  End If
50170
50180  If UCase$(Panel.key) = UCase(stb.Panels(3).key) Then
50190   If lsv.ListItems.Count = 17 Then
50200    If Not b2 Then
50210      Set lsv.Picture = pic.Image
50220     Else
50230      lsv.Picture = Nothing
50240    End If
50250    b2 = Not b2
50260   End If
50270  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "stb_PanelClick")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub CreateMutexIfNecessary()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If mutexLocal Is Nothing Then
50020   Set mutexLocal = New clsMutex
50030  End If
50040  If mutexLocal.MutexHandle = 0 Then
50050   ' Create a lokal mutex
50060    mutexLocal.CreateMutex PDFCreator_GUID
50070  End If
50080  If mutexGlobal Is Nothing Then
50090   Set mutexGlobal = New clsMutex
50100  End If
50110  If mutexGlobal.MutexHandle = 0 Then
50120  ' Create a global mutex
50130    mutexGlobal.CreateMutex "Global\" & PDFCreator_GUID
50140  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "CreateMutexIfNecessary")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub CheckClipboardForBitmap()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If Clipboard.GetFormat(vbCFBitmap) = True Then
50020    mnDocument(3).Enabled = True
50030    tlb(0).Buttons(7).Enabled = True
50040   Else
50050    mnDocument(3).Enabled = False
50060    tlb(0).Buttons(7).Enabled = False
50070  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "CheckClipboardForBitmap")
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
50010  PrinterStop = True
50020  PrintSelectedJobs = False
50030  If Timer3.Enabled = True Then
50040   Timer3.Enabled = False
50050   If In_eActionTimer = True Then
50060    If ShutDown = False Then
50070     ShutDown = True
50080    End If
50090    Cancel = True
50100   End If
50110  End If
50120  If Timer1.Enabled = True Then
50130   Timer1.Enabled = False
50140   If InTimer1 Then
50150    If ShutDown = False Then
50160     ShutDown = True
50170    End If
50180    Cancel = True
50190   End If
50200  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "Form_QueryUnload")
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
50010  If InTimer1 Then
50020   Exit Sub
50030  End If
50040
50050  InTimer1 = True
50060
50070  CreateMutexIfNecessary
50080  CheckClipboardForBitmap
50090
50100  DoEvents
50110  If FileExists(PDFCreatorApplicationPath & "Unload.tmp") = True Or Restart = True Then
50120   InTimer1 = False
50130   Unload Me
50140   Exit Sub
50150  End If
50160  CheckPrintJobs
50170  If Not NoProcessing Then
50180   CheckForPrinting
50190  End If
50200  If lsv.ListItems.Count = 0 And LenB(CommandSwitch("IF", True)) > 0 And ShellAndWaitingIsRunning = False Then
50210   InTimer1 = False
50220   Unload Me
50230   Exit Sub
50240  End If
50250
50260  If lsv.ListItems.Count = 1 Then
50270   If lsv.SelectedItem.Index <> 1 Then
50280    lsv.ListItems(1).Selected = True
50290   End If
50300  End If
50310  DoEvents
50320
50330  InTimer1 = False
50340  If ShutDown And In_eActionTimer = False Then
50350   Timer3.Enabled = False
50360   Timer1.Enabled = False
50370   Unload Me
50380  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "Timer1_Timer")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub tlb_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case Index
        Case 0
50031    Select Case Button.Index
          Case 1
50050      SetMenuPrinterStop
50060     Case 2
50070      frmOptions.Show , Me
50080     Case 3
50090      frmLog.Show , Me
50100     Case 5
50110      DocumentPrint
50120     Case 6
50130      DocumentAdd
50140     Case 7
50150      DocumentAddFromClipboard
50160     Case 8
50170      DocumentDelete
50180     Case 9
50190      DocumentTop
50200     Case 10
50210      DocumentUp
50220     Case 11
50230      DocumentDown
50240     Case 12
50250      DocumentBottom
50260     Case 13
50270      DocumentCombine False
50280     Case 14
50290      DocumentCombine True
50300     Case 15
50310      ' DocumentSave ' Not possible for the moment
50320     Case 17
50330      Call HTMLHelp_ShowTopic("html\pdfcreator-user-manual.html")
50340    End Select
50350   Case 1
50361    Select Case Button.Index
          Case 1
50380      CombineAllAndSend
50390     Case 2
50400      SendEmail
50410    End Select
50420  End Select
50430  SetDocMenuAndToolbar
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "tlb_ButtonClick")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtEmailAddress_Change()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  txtEmailAddress.ToolTipText = LanguageStrings.DialogEmailAddress & ": " & txtEmailAddress.Text
50020  If InStr(1, txtEmailAddress.Text, "@") = 0 And LenB(Options.StandardMailDomain) > 0 Then
50030   txtEmailAddress.ToolTipText = txtEmailAddress.ToolTipText & "@" & Options.StandardMailDomain
50040  End If
50050  SetDocMenuAndToolbar
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "txtEmailAddress_Change")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub CheckForPrinting()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim isf As clsInfoSpoolFile, PrinterDefaultProfile As String, opt As tOptions, opt2 As tOptions
50020
50030  If lsv.ListItems.Count > 0 Then
50040   If mnPrinter(2).Checked = True Then
50050     If PrintSelectedJobs = True Then
50060       If lsv.ListItems(1).SubItems(1) = LanguageStrings.ListPrinting Then
50070         'PDFSpoolfile = lsv.ListItems(1).SubItems(4)
50080         CurrentInfoSpoolFile = lsv.ListItems(1).SubItems(4)
50090         Set isf = New clsInfoSpoolFile
50100         isf.ReadInfoFile CurrentInfoSpoolFile
50110         DoEvents
50120         If LenB(isf.FirstPrinterName) > 0 And UCase$(OldPrinter) <> UCase$(isf.FirstPrinterName) Then
50130          OldPrinter = isf.FirstPrinterName
50140          Options = ReadOptions(True, , GetPrinterDefaultProfile(isf.FirstPrinterName))
50150         End If
50160         frmPrinting.PrinterProfile = GetPrinterDefaultProfile(isf.FirstPrinterName)
50170         frmPrinting.Show vbModal, Me
50180        Else
50190         PrintSelectedJobs = False
50200       End If
50210      Else
50220       If lsv.ListItems(1).SubItems(1) <> LanguageStrings.ListWaiting Then
50230        lsv.ListItems(1).SubItems(1) = LanguageStrings.ListWaiting
50240       End If
50250     End If
50260    Else
50270     If lsv.ListItems(1).SubItems(1) <> LanguageStrings.ListPrinting Then
50280      lsv.ListItems(1).SubItems(1) = LanguageStrings.ListPrinting
50290     End If
50300     ' PDFSpoolfile = lsv.ListItems(1).SubItems(4)
50310     CurrentInfoSpoolFile = lsv.ListItems(1).SubItems(4)
50320     If PrinterStop = False Then
50330      If IsFormLoaded(frmPrinting) = False Then
50340       Set isf = New clsInfoSpoolFile
50350       isf.ReadInfoFile CurrentInfoSpoolFile
50360       PrinterDefaultProfile = Trim$(GetPrinterDefaultProfile(isf.FirstPrinterName))
50370       If LenB(PrinterDefaultProfile) > 0 Then
50380         opt = ReadOptions(True, , PrinterDefaultProfile)
50390        Else
50400         opt = Options
50410       End If
50420       If opt.UseAutosave = 1 Then
50430         opt2 = Options
50440         Options = opt
50450         Autosave CurrentInfoSpoolFile
50460         SaveOption Options, "Counter", PrinterDefaultProfile
50470         opt2.Counter = Options.Counter
50480         Options = opt2
50490        Else
50500         If LenB(isf.FirstPrinterName) > 0 And UCase$(OldPrinter) <> UCase$(isf.FirstPrinterName) Then
50510          OldPrinter = isf.FirstPrinterName
50520         End If
50530         frmPrinting.PrinterProfile = GetPrinterDefaultProfile(isf.FirstPrinterName)
50540         frmPrinting.Show , Me
50550       End If
50560      End If
50570     End If
50580     If PrinterStop = False And NoProcessing = False Then
50590       mnPrinter(2).Checked = False
50600       tlb(0).Buttons(1).Image = 1
50610      Else
50620       mnPrinter(2).Checked = True
50630       tlb(0).Buttons(1).Image = 2
50640     End If
50650   End If
50660  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "CheckForPrinting")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub CheckPrintJobs()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim lItem As ListItem, files As Collection, tColl As Collection, _
  tFile() As String, i As Long, j As Long, kB As Long, MB As Long, GB As Long, _
  ind As Long, tStr As String, tB As Boolean
50040  Dim isf As clsInfoSpoolFile, isfi As clsInfoSpoolFileInfo
50050  kB = 1024: MB = kB * 1024: GB = MB * 1024
50060  Set files = New Collection
50070
50080  ' Get all files from spool directory
50090  FindFiles GetPDFCreatorSpoolDirectory, files, "*.inf", , False, True
50100
50110  ' Remove all files with file length 0, possible in use
50120  For i = files.Count To 1 Step -1
50130   tFile = Split(files(i), "|")
50140   If FileLen(tFile(1)) = 0 Then
50150    files.Remove i
50160   End If
50170  Next i
50180
50190  If files.Count = 0 And lsv.ListItems.Count > 0 Then
50200   lsv.ListItems.Clear
50210   SetDocMenuAndToolbar
50220   Exit Sub
50230  End If
50240
50250  tB = False
50260
50270  Set tColl = New Collection
50280  For j = 1 To lsv.ListItems.Count
50290   ind = 0
50300   For i = 1 To files.Count
50310    tFile = Split(files.Item(i), "|")
50320    If UCase$(tFile(1)) = UCase$(lsv.ListItems(j).SubItems(4)) Then
50330     ind = i
50340     Exit For
50350    End If
50360   Next i
50370   If ind = 0 Then
50380    tColl.Add j
50390   End If
50400  Next j
50410  If tColl.Count > 0 Then
50420   tB = True
50430  End If
50440  For j = 1 To tColl.Count
50450   lsv.ListItems.Remove tColl(j) - (j - 1)
50460  Next j
50470
50480  Set tColl = New Collection
50490  For j = 1 To files.Count
50500   tFile = Split(files.Item(j), "|")
50510   ind = 0
50520   For i = 1 To lsv.ListItems.Count
50530    If UCase$(tFile(1)) = UCase$(lsv.ListItems(i).SubItems(4)) Then
50540     ind = i
50550     Exit For
50560    End If
50570   Next i
50580   If ind > 0 And ind < lsv.ListItems.Count + 1 Then
50590    tColl.Add j
50600   End If
50610  Next j
50620  For j = 1 To tColl.Count
50630   files.Remove tColl(j) - (j - 1)
50640  Next j
50650
50660  For j = 1 To files.Count
50670   tFile = Split(files.Item(j), "|")
50680   ind = 0
50690   For i = 1 To lsv.ListItems.Count
50700    If UCase$(tFile(1)) = UCase$(lsv.ListItems(i).SubItems(4)) Then
50710     ind = i
50720     Exit For
50730    End If
50740   Next i
50750   If ind = 0 Then
50760    tB = True
50770    If FormISLoaded("frmOptions") = False And Me.Visible = True Then
50780     Me.Show
50790    End If
50800
50810    Set isf = New clsInfoSpoolFile
50820    isf.ReadInfoFile tFile(1)
50830    Set lItem = lsv.ListItems.Add(, , isf.FirstDocumentTitle)
50840
50850    lItem.SubItems(1) = LanguageStrings.ListWaiting
50860    lItem.SubItems(2) = tFile(3)
50870
50880    If CLng(tFile(2)) > GB Then
50890      lItem.SubItems(3) = Format$(CDbl(isf.sumFileSizes) / GB, "#,##0.00 " & LanguageStrings.ListGBytes)
50900     Else
50910      If CLng(tFile(2)) > MB Then
50920        lItem.SubItems(3) = Format$(CDbl(isf.sumFileSizes) / MB, "#,##0.00 " & LanguageStrings.ListMBytes)
50930       Else
50940        If CLng(tFile(2)) > kB Then
50950          lItem.SubItems(3) = Format$(CDbl(isf.sumFileSizes) / kB, "#,##0.00 " & LanguageStrings.ListKBytes)
50960         Else
50970          lItem.SubItems(3) = Format$(CDbl(isf.sumFileSizes), "#,##0 " & LanguageStrings.ListBytes)
50980        End If
50990     End If
51000    End If
51010    lItem.SubItems(4) = tFile(1)
51020    DoEvents
51030   End If
51040  Next j
51050  If lsv.ListItems.Count = 1 Then
51060    tStr = LanguageStrings.ListStatus & ": " & lsv.ListItems.Count & " " & LanguageStrings.MessagesMsg01
51070   Else
51080    tStr = LanguageStrings.ListStatus & ": " & lsv.ListItems.Count & " " & LanguageStrings.MessagesMsg02
51090  End If
51100  If tStr <> stb.Panels("Status").Text Then
51110   stb.Panels("Status").Text = tStr
51120  End If
51130  If tB = True Then
51140   SetDocMenuAndToolbar
51150  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "CheckPrintJobs")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetDocMenuAndToolbar()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim c As Long
50020  For c = 0 To mnDocument.Count - 1
50030   If c <> 14 Then ' Disable PS-Save menu item, because doesn't work with the new monitor dll
50040     If mnDocument(c).Enabled = False Then mnDocument(c).Enabled = True
50050    Else
50060     If mnDocument(c).Enabled = True Then mnDocument(c).Enabled = False
50070     If mnDocument(c).Visible = True Then mnDocument(c).Visible = False
50080   End If
50090  Next c
50100  For c = 1 To tlb(0).Buttons.Count
50110   If c <> 15 Then ' Disable PS-Save toolbar button, because doesn't work with the new monitor dll
50120     If tlb(0).Buttons(c).Enabled = False Then tlb(0).Buttons(c).Enabled = True
50130    Else
50140     If tlb(0).Buttons(c).Enabled = True Then tlb(0).Buttons(c).Enabled = False
50150     If tlb(0).Buttons(c).Visible = True Then tlb(0).Buttons(c).Visible = False
50160   End If
50170  Next c
50180  If (Options.Toolbars And 2) = 2 Then
50190   For c = 1 To tlb(1).Buttons.Count
50200    If tlb(1).Buttons(c).Enabled = False Then tlb(1).Buttons(c).Enabled = True
50210   Next c
50220  End If
50230  If Clipboard.GetFormat(vbCFBitmap) = True Then
50240    If mnDocument(3).Enabled = False Then mnDocument(3).Enabled = True
50250    If tlb(0).Buttons(7).Enabled = False Then tlb(0).Buttons(7).Enabled = True
50260   Else
50270    If mnDocument(3).Enabled = True Then mnDocument(3).Enabled = False
50280    If tlb(0).Buttons(7).Enabled = True Then tlb(0).Buttons(7).Enabled = False
50290  End If
50301  Select Case True
        Case lsv.ListItems.Count = 0, LvwGetCountSelectedItems(lsv, False) = 0
50320    With mnDocument
50330     If .Item(0).Enabled = True Then .Item(0).Enabled = False
50340     If .Item(4).Enabled = True Then .Item(4).Enabled = False
50350     If .Item(6).Enabled = True Then .Item(6).Enabled = False
50360     If .Item(7).Enabled = True Then .Item(7).Enabled = False
50370     If .Item(8).Enabled = True Then .Item(8).Enabled = False
50380     If .Item(9).Enabled = True Then .Item(9).Enabled = False
50390     If .Item(11).Enabled = True Then .Item(11).Enabled = False
50400     If .Item(12).Enabled = True Then .Item(12).Enabled = False
50410     If .Item(14).Enabled = True Then .Item(14).Enabled = False
50420     If (Options.Toolbars And 2) = 2 Then
50430      If .Item(16).Enabled = True Then .Item(16).Enabled = False
50440      If .Item(17).Enabled = True Then .Item(17).Enabled = False
50450     End If
50460    End With
50470    With tlb(0)
50480     If .Buttons(5).Enabled = True Then .Buttons(5).Enabled = False ' print
50490     For c = 8 To 15
50500      If .Buttons(c).Enabled = True Then .Buttons(c).Enabled = False
50510     Next c
50520    End With
50530    With tlb(1)
50540     If (Options.Toolbars And 2) = 2 Then
50550      If .Buttons(1).Enabled = True Then .Buttons(1).Enabled = False
50560      If .Buttons(2).Enabled = True Then .Buttons(2).Enabled = False
50570     End If
50580    End With
50590    Exit Sub
50600   Case lsv.ListItems.Count = 1
50610    With mnDocument
50620     If .Item(6).Enabled = True Then .Item(6).Enabled = False
50630     If .Item(7).Enabled = True Then .Item(7).Enabled = False
50640     If .Item(8).Enabled = True Then .Item(8).Enabled = False
50650     If .Item(9).Enabled = True Then .Item(9).Enabled = False
50660     If .Item(11).Enabled = True Then .Item(11).Enabled = False
50670     If .Item(12).Enabled = True Then .Item(12).Enabled = False
50680     If (Options.Toolbars And 2) = 2 Then
50690      If .Item(16).Enabled = True Then .Item(16).Enabled = False
50700      If LenB(txtEmailAddress.Text) = 0 Or Options.DisableEmail <> 0 Then
50710       If .Item(17).Enabled = True Then .Item(17).Enabled = False
50720      End If
50730     End If
50740    End With
50750    With tlb(0)
50760     For c = 9 To 14
50770      If .Buttons(c).Enabled = True Then .Buttons(c).Enabled = False
50780     Next c
50790    End With
50800    With tlb(1)
50810     If (Options.Toolbars And 2) = 2 Then
50820      If LenB(txtEmailAddress.Text) = 0 Then
50830       If .Buttons(2).Enabled = True Then .Buttons(2).Enabled = False
50840      End If
50850     End If
50860    End With
50870   Case lsv.ListItems.Count > 1
50880    With mnDocument
50890     If AllSelectedListitemsAtTop Then
50900      .Item(6).Enabled = False
50910      .Item(7).Enabled = False
50920     End If
50930     If AllSelectedListitemsAtBottom Then
50940      .Item(8).Enabled = False
50950      .Item(9).Enabled = False
50960     End If
50970     If LvwGetCountSelectedItems(lsv, False) = 1 Then
50980      .Item(11).Enabled = False
50990     End If
51000     If LvwGetCountSelectedItems(lsv, False) > 1 Then
51010      .Item(14).Enabled = False
51020     End If
51030     If ((Options.Toolbars And 2) = 2 And LenB(txtEmailAddress.Text) = 0) Or Options.DisableEmail <> 0 Then
51040      .Item(16).Enabled = False
51050      .Item(17).Enabled = False
51060     End If
51070    End With
51080    With tlb(0)
51090     If AllSelectedListitemsAtTop Then
51100      If .Buttons(9).Enabled = True Then .Buttons(9).Enabled = False
51110      If .Buttons(10).Enabled = True Then .Buttons(10).Enabled = False
51120     End If
51130     If AllSelectedListitemsAtBottom Then
51140      If .Buttons(11).Enabled = True Then .Buttons(11).Enabled = False
51150      If .Buttons(12).Enabled = True Then .Buttons(12).Enabled = False
51160     End If
51170     If LvwGetCountSelectedItems(lsv, False) = 1 Then
51180      If .Buttons(13).Enabled = True Then .Buttons(13).Enabled = False
51190     End If
51200     If LvwGetCountSelectedItems(lsv, False) > 1 Then
51210      If .Buttons(15).Enabled = True Then .Buttons(15).Enabled = False
51220     End If
51230    End With
51240    With tlb(1)
51250     If (Options.Toolbars And 2) = 2 And LenB(txtEmailAddress.Text) = 0 Then
51260      If .Buttons(1).Enabled = True Then .Buttons(1).Enabled = False
51270      If .Buttons(2).Enabled = True Then .Buttons(2).Enabled = False
51280     End If
51290    End With
51300  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "SetDocMenuAndToolbar")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Autosave(Optional InfoSpoolFileName As String = vbNullString)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tColl As Collection, i As Long, tFile() As String, Pathname As String, _
  OutputFilename As String, PDFDocInfo As tPDFDocInfo, tStr As String, _
  PSHeader As tPSHeader, tDate As Date, mail As clsPDFCreatorMail
50040  Dim PSFile As String, isf As clsInfoSpoolFile
50050  Dim PDFDocInfoFile As String, StampFile As String
50060
50070  InAutoSave = True
50080  IsConverted = False
50090  Set tColl = New Collection
50100
50110  If Len(InfoSpoolFileName) > 0 Then
50120    If FileExists(InfoSpoolFileName) = True Then
50130     SplitPath InfoSpoolFileName, , Pathname
50140     tColl.Add Pathname & "|" & InfoSpoolFileName & "|" & FileLen(InfoSpoolFileName) & "|" & FileDateTime(InfoSpoolFileName)
50150    End If
50160   Else
50170    FindFiles GetPDFCreatorSpoolDirectory, tColl, "*.inf", , True, True
50180  End If
50190
50200  tStr = "Autosavemodus: " & tColl.Count & "files"
50210  IfLoggingWriteLogfile tStr
50220  WriteToSpecialLogfile tStr
50230  Do While tColl.Count > 0
50240   For i = 1 To tColl.Count
50250    tFile = Split(tColl.Item(i), "|")
50260    If FileExists(tFile(1)) And Not FileInUse(tFile(1)) Then
50270     OutputFilename = GetAutosaveFilename(tFile(1))
50280     SplitPath OutputFilename, , Pathname
50290     If IsValidPath(Pathname) = True Then
50300       If DirExists(Pathname) = False Then
50310        MakePath (Pathname)
50320       End If
50330       tStr = "Autosavemodus: Create File '" & OutputFilename & "'"
50340       IfLoggingWriteLogfile tStr
50350       WriteToSpecialLogfile tStr
50360
50370       Set isf = New clsInfoSpoolFile
50380       isf.ReadInfoFile tFile(1)
50390
50400       PSHeader = GetPSHeader(isf.FirstSpoolFileName)
50410       tDate = Now
50420       With PDFDocInfo
50430        If Options.UseStandardAuthor = 1 Then
50440          .Author = GetSubstFilename(isf.FirstSpoolFileName, _
          RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardAuthor)), True)
50460         Else
50470          .Author = isf.FirstUserName
50480        End If
50490        If LenB(PSHeader.CreationDate.Comment) > 0 Then
50500          tStr = PSHeader.CreationDate.Comment
50510         Else
50520          tStr = CStr(tDate)
50530        End If
50540        .CreationDate = GetDocDate(Trim$(Options.StandardCreationdate), Options.StandardDateformat, FormatPrintDocumentDate(tStr))
50550        .Creator = App.EXEName & " Version " & App.Major & "." & App.Minor & "." & App.Revision
50560        .Keywords = GetSubstFilename(tFile(1), RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardKeywords)), , , True)
50570        'tStr = CStr(tDate)
50580        .ModifyDate = GetDocDate(Trim$(Options.StandardModifydate), Options.StandardDateformat, FormatPrintDocumentDate(tStr))
50590        .Subject = GetSubstFilename(isf.FirstSpoolFileName, RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardSubject)), , , True)
50600        If Len(Options.StandardTitle) > 0 Then
50610          .Title = GetSubstFilename(isf.FirstSpoolFileName, RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardTitle)), , , True)
50620         Else
50630          .Title = GetSubstFilename(isf.FirstSpoolFileName, Options.SaveFilename, , , True)
50640        End If
50650       End With
50660
50670       PDFDocInfoFile = CreatePDFDocInfoFile(tFile(1), PDFDocInfo)
50680       StampFile = CreateStampFile(tFile(1))
50690
50700       If Options.RunProgramBeforeSaving = 1 Then
50710        RunProgramBeforeSaving Me.hWnd, tFile(1), _
        Options.RunProgramBeforeSavingProgramParameters, _
        Options.RunProgramBeforeSavingWindowstyle
50740       End If
50750       CallGScript tFile(1), OutputFilename, Options, Options.AutosaveFormat, PDFDocInfoFile, StampFile
50760
50770       If FileExists(OutputFilename) = True Then
50780         IsConverted = True
50790         tStr = "Autosavemodus: Create File '" & OutputFilename & "' success"
50800         IfLoggingWriteLogfile tStr
50810         WriteToSpecialLogfile tStr
50820         If Options.RunProgramAfterSaving = 1 Then
50830          RunProgramAfterSaving Me.hWnd, OutputFilename, _
         Options.RunProgramAfterSavingProgramParameters, _
         Options.RunProgramAfterSavingWindowstyle, tFile(1)
50860         End If
50870         If Options.SendEmailAfterAutoSaving = 1 Then
50880          Set mail = New clsPDFCreatorMail
50890          If mail.Send(OutputFilename, PDFDocInfo.Subject, Options.SendMailMethod) <> 0 Then
50900           MsgBox LanguageStrings.MessagesMsg04, vbCritical, App.EXEName
50910          End If
50920          Set mail = Nothing
50930         End If
50940         Options.Counter = Options.Counter + 1
50950         If Options.AutosaveStartStandardProgram = 1 Then
50960          If Options.OnePagePerFile = 1 Then
50970            OpenDocument Replace$(OutputFilename, "%d", "1", , , vbTextCompare)
50980           Else
50990            OpenDocument OutputFilename
51000          End If
51010         End If
51020        Else
51030         tStr = "Autosavemodus: Create File '" & OutputFilename & "' failed"
51040         IfLoggingWriteLogfile tStr
51050         WriteToSpecialLogfile tStr
51060       End If
51070      Else
51080       IfLoggingWriteLogfile "Error: Invalid autosave pathname, spoolfile will be deleted! > " & OutputFilename
51090     End If
51100     CheckForPrintingAfterSaving tFile(1), Options
51110
51120     KillInfoSpoolFiles tFile(1)
51130
51140     ConvertedOutputFilename = OutputFilename
51150     ReadyConverting = True
51160    End If
51170   Next i
51180   Set tColl = GetFiles(GetPDFCreatorTempfolder, "PDF*.tmp", SortedByDate)
51190  Loop
51200  InAutoSave = False
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "Autosave")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ShowPaypalMenuimage()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim h1 As Long, h2 As Long, com As Long
50020  h1 = GetMenu(Me.hWnd): h2 = GetSubMenu(h1, 4)
50030  com = GetMenuItemID(h2, 2)
50040  ModifyMenu h2, com, MF_BYCOMMAND Or MF_BITMAP, com, CLng(imgPaypal.Picture)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "ShowPaypalMenuimage")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetGSRevision()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim RNum As String
50020  If Len(CStr(GSRevision.intRevision)) >= 3 Then
50030    RNum = Mid(CStr(GSRevision.intRevision), Len(CStr(GSRevision.intRevision)) - 1, 2)
50040    RNum = Mid(CStr(GSRevision.intRevision), 1, Len(CStr(GSRevision.intRevision)) - 2) & "." & RNum
50050   Else
50060    RNum = ""
50070  End If
50080  If Len(GSRevision.strProduct) > 0 Then
50090    stb.Panels("GhostscriptRevision").Text = GSRevision.strProduct & " " & RNum
50100   Else
50110    stb.Panels("GhostscriptRevision").Text = "-"
50120  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "SetGSRevision")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub InitToolbar()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim btn As MSComctlLib.Button
50020  With tlb(0)
50030   Set .ImageList = imlTlb
50040   .Buttons.Clear
50050   .Buttons.Add , , , , 1
50060   .Buttons.Add , , , , 3
50070   .Buttons.Add , , , , 4
50080   .Buttons.Add , , , tbrSeparator
50090   .Buttons.Add , , , , 5
50100   .Buttons.Add , , , , 6
50110   .Buttons.Add , , , , 18
50120   .Buttons.Add , , , , 7
50130   .Buttons.Add , , , , 8
50140   .Buttons.Add , , , , 9
50150   .Buttons.Add , , , , 10
50160   .Buttons.Add , , , , 11
50170   .Buttons.Add , , , , 12
50180   .Buttons.Add , , , , 15
50190   .Buttons.Add , , , , 13
50200   .Buttons.Add , , , tbrSeparator
50210   .Buttons.Add , , , , 14
50220  End With
50230  With tlb(1)
50240   Set .ImageList = imlTlb
50250   .Buttons.Clear
50260   .Buttons.Add , , , , 16
50270   .Buttons.Add , , , , 17
50280   .Buttons.Add , , , tbrSeparator
50290   Set btn = .Buttons.Add(, "emailAddress", , tbrPlaceholder)
50300   btn.Width = (tlb(0).Buttons(15).Left + tlb(0).Buttons(15).Width) - btn.Left
50310  End With
50320  With txtEmailAddress
50330   .Top = tlb(1).Buttons("emailAddress").Top - (.Height - tlb(1).Height) / 2
50340   .Left = tlb(1).Buttons("emailAddress").Left
50350  End With
50360
50370  SetLanguageToolbar
50380  DrawToolbars
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "InitToolbar")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub DrawToolbars()
 If (Options.Toolbars And 2) <> 2 Then
   txtEmailAddress.Enabled = False
   txtEmailAddress.Visible = False
   On Error Resume Next
   mnDocument(15).Enabled = False
   mnDocument(16).Enabled = False
   mnDocument(17).Enabled = False
   mnDocument(14).Visible = False
   mnDocument(15).Visible = False
   mnDocument(16).Visible = False
   mnDocument(17).Visible = False
   On Error GoTo 0
 End If
 If (Options.Toolbars And 1) = 1 Then
   tlb(0).Visible = True
   mnViewToolbars(0).Checked = True
  Else
   tlb(0).Visible = False
   mnViewToolbars(0).Checked = False
 End If
 If (Options.Toolbars And 2) = 2 Then
   tlb(1).Visible = True
   txtEmailAddress.Visible = True
   mnViewToolbars(1).Checked = True
   mnDocument(16).Visible = True
   mnDocument(17).Visible = True
  Else
   tlb(1).Visible = False
   txtEmailAddress.Visible = False
   mnViewToolbars(1).Checked = False
 End If
End Sub

Private Sub DocumentPrint()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, j As Long
50020  For j = 1 To LvwGetCountSelectedItems(lsv, False)
50030   DoEvents
50040   For i = lsv.ListItems.Count To 1 Step -1
50050    If lsv.ListItems(i).Selected = True Then
50060     lsv.ListItems(i).SubItems(1) = LanguageStrings.ListPrinting
50070     LvwListItemToTop lsv, i, True
50080     Exit For
50090    End If
50100   Next i
50110  Next j
50120  PrintSelectedJobs = True
50130 ' SetPrinterStop False
50140 ' mnPrinter(2).Checked = False
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "DocumentPrint")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub DocumentAddFromClipboard()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim File As String
50020  If Clipboard.GetFormat(vbCFBitmap) = True Then
50030   File = GetPDFCreatorSpoolDirectory & GetGUID
50040   ConvertStandardImageFromPicture Clipboard.GetData(2), File & ".ps", "Screenshot"
50050   CreateInfoSpoolFile File & ".inf"
50060  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "DocumentAddFromClipboard")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub DocumentAdd()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, Cancel As Boolean, cFiles As Collection, aLen As Double, _
  OnlyPsFiles As Boolean, Ext As String, DefaultPrintername As String, _
  tFilename As String, tLen As Double
50040  Dim Path As String, psFileName As String, ini As clsINI, Title As String, strGUID As String, _
  spoolDirectory As String
50060  Set cFiles = GetFilename("", GetMyFiles, 0, _
  LanguageStrings.ListPostscriptFiles & " (*.ps)|*.ps|" & LanguageStrings.ListAllFiles & " (*.*)|*.*", _
   OpenFile, Cancel, Me)
50090  If Cancel = True Then
50100   Screen.MousePointer = vbNormal
50110   Exit Sub
50120  End If
50130  aLen = 0
50140  For i = 1 To cFiles.Count
50150   aLen = aLen + FileLen(cFiles.Item(i))
50160  Next i
50170  OnlyPsFiles = True
50180  For i = 1 To cFiles.Count
50190   SplitPath cFiles.Item(i), , , , , Ext
50200   If UCase$(Ext) <> "PS" And UCase$(Ext) <> "EPS" Then
50210    OnlyPsFiles = False
50220    Exit For
50230   End If
50240  Next i
50250  If UCase$(Printer.DeviceName) <> UCase$(GetPDFCreatorPrintername) And OnlyPsFiles = False Then
50260   If Options.NoConfirmMessageSwitchingDefaultprinter = 0 Then
50270    If ChangeDefaultprinter = False Then
50280     frmSwitchDefaultprinter.Show vbModal, Me
50290     If ChangeDefaultprinter = False Then
50300      Screen.MousePointer = vbNormal
50310      Exit Sub
50320     End If
50330    End If
50340   End If
50350  End If
50360  ChangeDefaultprinter = True
50370  DefaultPrintername = Printer.DeviceName
50380  SetDefaultprinterInProg GetPDFCreatorPrintername
50390  aLen = 0
50400  For i = 1 To cFiles.Count
50410   aLen = aLen + FileLen(cFiles.Item(i))
50420  Next i
50430  For i = 1 To cFiles.Count
50440   SplitPath cFiles.Item(i), , , , , Ext
50450   If IsPostscriptFile(cFiles.Item(i)) = True And (UCase$(Ext) = "PS" Or UCase$(Ext) = "EPS") Then
50460     spoolDirectory = GetPDFCreatorSpoolDirectory
50470     strGUID = GetGUID
50480     tFilename = spoolDirectory & strGUID & ".inf"
50490     psFileName = spoolDirectory & strGUID & ".ps"
50500     FileCopy cFiles.Item(i), psFileName
50510     CreateInfoSpoolFile psFileName, tFilename
50520    Else
50530     DoEvents
50540     ShellAndWait Me.hWnd, "print", cFiles.Item(i), "", vbNullChar, wHidden, WCTermination, 60000, True
50550   End If
50560   DoEvents
50570   tLen = tLen + FileLen(cFiles.Item(i))
50580   stb.Panels("Percent").Text = Format$(tLen / aLen, " 0.0%")
50590   DoEvents
50600  Next i
50610  If DefaultPrintername <> vbNullString Then
50620   SetDefaultprinterInProg DefaultPrintername
50630  End If
50640  stb.Panels("Percent").Text = vbNullString
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "DocumentAdd")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub DocumentDelete()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long
50020  For i = lsv.ListItems.Count To 1 Step -1
50030   If lsv.ListItems(i).Selected = True Then
50040    KillInfoSpoolFiles lsv.ListItems(i).SubItems(4)
50050   End If
50060   DoEvents
50070  Next i
50080  LvwRemoveSelectedItems lsv, True
50090  If lsv.ListItems.Count > 0 Then
50100   If lsv.SelectedItem.Index > lsv.ListItems.Count Then
50110     lsv.ListItems(lsv.SelectedItem.Index - 1).Selected = True
50120    Else
50130     lsv.ListItems(lsv.SelectedItem.Index).Selected = True
50140   End If
50150  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "DocumentDelete")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub DocumentTop()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, j As Long
50020  For j = 1 To LvwGetCountSelectedItems(lsv, False)
50030   For i = lsv.ListItems.Count To 1 Step -1
50040    If lsv.ListItems(i).Selected = True Then
50050     LvwListItemToTop lsv, i, True
50060     Exit For
50070    End If
50080   Next i
50090  Next j
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "DocumentTop")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub DocumentUp()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  LvwListItemUp lsv, , True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "DocumentUp")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub DocumentDown()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  LvwListItemDown lsv, , True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "DocumentDown")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub DocumentBottom()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, j As Long
50020  For j = 1 To LvwGetCountSelectedItems(lsv, False)
50030   For i = 1 To lsv.ListItems.Count
50040    If lsv.ListItems(i).Selected = True Then
50050     LvwListItemToBottom lsv, i, True
50060     Exit For
50070    End If
50080   Next i
50090  Next j
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "DocumentBottom")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function DocumentCombine(all As Boolean) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, tFilename As String, tFilename2 As String
50020  Dim isf As clsInfoSpoolFile, isfi As clsInfoSpoolFileInfo, c As Long, ini As clsINI, j As Long
50030
50040  Screen.MousePointer = vbHourglass
50050  LockWindowUpdate lsv.hWnd
50060
50070  c = 1
50080  tFilename = GetPDFCreatorSpoolDirectory & GetGUID & ".inf"
50090  Set ini = New clsINI
50100  ini.filename = tFilename
50110  Set isf = New clsInfoSpoolFile
50120  For i = 1 To lsv.ListItems.Count
50130   If all = True Or lsv.ListItems(i).Selected = True Then
50140    isf.ReadInfoFile lsv.ListItems(i).SubItems(4)
50150    ini.Section = CStr(c)
50160    For j = 1 To isf.InfoFiles.Count
50170     Set isfi = isf.InfoFiles(j)
50180     ini.SaveKey isfi.ClientComputer, "ClientComputer"
50190     ini.SaveKey isfi.DocumentTitle, "DocumentTitle"
50200     ini.SaveKey isfi.JobID, "JobId"
50210     ini.SaveKey isfi.PrinterName, "Printername"
50220     ini.SaveKey isfi.SessionID, "SessionID"
50230     ini.SaveKey isfi.SpoolFileName, "SpoolFilename"
50240     ini.SaveKey isfi.UserName, "UserName"
50250     ini.SaveKey isfi.WinStation, "WinStation"
50260    Next j
50270    c = c + 1
50280    KillFile lsv.ListItems(i).SubItems(4)
50290   End If
50300  Next i
50310  DocumentCombine = tFilename
50320  LockWindowUpdate 0&
50330  Screen.MousePointer = vbNormal
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "DocumentCombine")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub DocumentSave()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010   Dim tFilename As String, cFiles As Collection, Cancel As Boolean
50020  tFilename = ReplaceForbiddenChars(GetPSTitle(lsv.ListItems(lsv.SelectedItem.Index).SubItems(4)), , ".")
50030  If LenB(tFilename) = 0 Then
50040   SplitPath lsv.ListItems(lsv.SelectedItem.Index).SubItems(4), , , , tFilename
50050  End If
50060  Set cFiles = GetFilename(tFilename, GetMyFiles, 0, _
  LanguageStrings.ListPostscriptFiles & " (*.ps)|*.ps", _
   SaveFile, Cancel, Me)
50090  If Cancel = True Then
50100   Screen.MousePointer = vbNormal
50110   Exit Sub
50120  End If
50130  If cFiles.Count > 0 Then
50140   FileCopy lsv.ListItems(lsv.SelectedItem.Index).SubItems(4), cFiles.Item(1)
50150  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "DocumentSave")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SetMenuPrinterStop()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If mnPrinter(2).Checked = False Or NoProcessing = True Then
50020    SetPrinterStop True
50030    mnPrinter(2).Checked = True
50040    tlb(0).Buttons(1).Image = 2
50050   Else
50060    SetPrinterStop False
50070    mnPrinter(2).Checked = False
50080    tlb(0).Buttons(1).Image = 1
50090  End If
50100  If Not m_frmSysTray Is Nothing Then
50110   If mnPrinter(2).Checked = True Then
50120     m_frmSysTray.IconHandle = imlSystray.ListImages(1).Picture.handle
50130     m_frmSysTray.mnuSysTray(2).Checked = True
50140    Else
50150     m_frmSysTray.IconHandle = imlSystray.ListImages(2).Picture.handle
50160     m_frmSysTray.mnuSysTray(2).Checked = False
50170   End If
50180  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "SetMenuPrinterStop")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetLanguageToolbar()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With tlb(0)
50020   .Buttons(1).ToolTipText = LanguageStrings.DialogPrinterPrinterStop
50030   .Buttons(2).ToolTipText = LanguageStrings.DialogPrinterOptions
50040   .Buttons(3).ToolTipText = LanguageStrings.DialogPrinterLogfile
50050   .Buttons(5).ToolTipText = LanguageStrings.DialogDocumentPrint
50060   .Buttons(6).ToolTipText = LanguageStrings.DialogDocumentAdd
50070   .Buttons(7).ToolTipText = LanguageStrings.DialogDocumentAddFromClipboard
50080   .Buttons(8).ToolTipText = LanguageStrings.DialogDocumentDelete
50090   .Buttons(9).ToolTipText = LanguageStrings.DialogDocumentTop
50100   .Buttons(10).ToolTipText = LanguageStrings.DialogDocumentUp
50110   .Buttons(11).ToolTipText = LanguageStrings.DialogDocumentDown
50120   .Buttons(12).ToolTipText = LanguageStrings.DialogDocumentBottom
50130   .Buttons(13).ToolTipText = LanguageStrings.DialogDocumentCombine
50140   .Buttons(14).ToolTipText = LanguageStrings.DialogDocumentCombineAll
50150   .Buttons(15).ToolTipText = LanguageStrings.DialogDocumentSave
50160   .Buttons(17).ToolTipText = "?"
50170  End With
50180  With tlb(1)
50190   .Buttons(1).ToolTipText = LanguageStrings.DialogDocumentCombineAllSend
50200   .Buttons(2).ToolTipText = LanguageStrings.DialogDocumentSend
50210  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "SetLanguageToolbar")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function GetVisibleToolbars() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim c As Long
50020  c = 0
50030  If (Options.Toolbars And 1) = 1 Then
50040   c = c + 1
50050  End If
50060  If (Options.Toolbars And 2) = 2 Then
50070   c = c + 1
50080  End If
50090  GetVisibleToolbars = c
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "GetVisibleToolbars")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub SendEmailImmediately(InputFilename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim OutputFilename As String, mail As clsPDFCreatorMail, rec As String
50020  If LenB(InputFilename) > 0 Then
50030   If FileExists(InputFilename) = True And FileInUse(InputFilename) = False Then
50040    rec = LTrim(txtEmailAddress.Text)
50050    If LenB(LTrim(Options.StandardMailDomain)) > 0 And InStr(1, LTrim(txtEmailAddress.Text), "@") <= 0 Then
50060     rec = rec & "@" & Options.StandardMailDomain
50070    End If
50080    OutputFilename = CompletePath(GetPDFCreatorTempfolder) & txtEmailAddress.Text & ".pdf"
50090
50100    If Options.RunProgramBeforeSaving = 1 Then
50110     RunProgramBeforeSaving Me.hWnd, GetShortName(InputFilename), _
    Options.RunProgramBeforeSavingProgramParameters, _
    Options.RunProgramBeforeSavingWindowstyle
50140    End If
50150    ConvertFile InputFilename, OutputFilename
50160    If Len(OutputFilename) > 0 And FileExists(OutputFilename) = True Then
50170     If Options.RunProgramAfterSaving = 1 Then
50180      RunProgramAfterSaving Me.hWnd, OutputFilename, _
      Options.RunProgramAfterSavingProgramParameters, _
      Options.RunProgramAfterSavingWindowstyle, InputFilename
50210     End If
50220     Set mail = New clsPDFCreatorMail
50230     If mail.Send(OutputFilename, Options.StandardSubject, Options.SendMailMethod, rec) <> 0 Then
50240      MsgBox LanguageStrings.MessagesMsg04, vbCritical, App.EXEName
50250     End If
50260     Set mail = Nothing
50270    End If
50280    Options.Counter = Options.Counter + 1
50290    KillFile InputFilename
50300   End If
50310  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "SendEmailImmediately")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub CombineAllAndSend()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If Options.ShowAnimation = 1 Then
50020   ShowAnimationWindow = True
50030   frmAnimation.Show
50040  End If
50050  SendEmailImmediately DocumentCombine(True)
50060  If Options.ShowAnimation = 1 Then
50070   ShowAnimationWindow = False
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "CombineAllAndSend")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SendEmail()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If lsv.ListItems.Count > 0 Then
50020   If lsv.SelectedItem.Index >= 0 Then
50030    If FileExists(lsv.SelectedItem.ListSubItems(4)) = True And FileInUse(lsv.SelectedItem.ListSubItems(4)) = False Then
50040     If Options.ShowAnimation = 1 Then
50050      ShowAnimationWindow = True
50060      frmAnimation.Show
50070     End If
50080     SendEmailImmediately lsv.SelectedItem.ListSubItems(4)
50090     If Options.ShowAnimation = 1 Then
50100      ShowAnimationWindow = False
50110     End If
50120    End If
50130   End If
50140  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "SendEmail")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function AllSelectedListitemsAtTop() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, tB As Boolean
50020  AllSelectedListitemsAtTop = False
50030  If lsv.ListItems.Count > 0 Then
50040   If lsv.ListItems(1).Selected = False Then
50050    Exit Function
50060   End If
50070   AllSelectedListitemsAtTop = True
50080   tB = False
50090   For i = 2 To lsv.ListItems.Count
50100    If tB = True And lsv.ListItems(i).Selected = True Then
50110     AllSelectedListitemsAtTop = False
50120     Exit For
50130    End If
50140    If lsv.ListItems(i).Selected = False Then
50150     tB = True
50160    End If
50170   Next i
50180  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "AllSelectedListitemsAtTop")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function AllSelectedListitemsAtBottom() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, tB As Boolean
50020  AllSelectedListitemsAtBottom = False
50030  If lsv.ListItems.Count > 0 Then
50040   If lsv.ListItems(lsv.ListItems.Count).Selected = False Then
50050    Exit Function
50060   End If
50070   AllSelectedListitemsAtBottom = True
50080   tB = False
50090   For i = lsv.ListItems.Count - 1 To 1 Step -1
50100    If tB = True And lsv.ListItems(i).Selected = True Then
50110     AllSelectedListitemsAtBottom = False
50120     Exit For
50130    End If
50140    If lsv.ListItems(i).Selected = False Then
50150     tB = True
50160    End If
50170   Next i
50180  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "AllSelectedListitemsAtBottom")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub m_frmSysTray_MenuClick(ByVal lIndex As Long, ByVal sKey As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case lIndex
        Case 0
50030    If Not ExistsAnModalForm Then
50040     ShowFrmMain
50050    End If
50060   Case Is > 1
50070    mnPrinter_Click CInt(lIndex - 2)
50080  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "m_frmSysTray_MenuClick")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub m_frmSysTray_SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If Not ExistsAnModalForm Then
50020   ShowFrmMain
50030  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "m_frmSysTray_SysTrayDoubleClick")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub m_frmSysTray_SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If (eButton = vbRightButton) Then
50020   m_frmSysTray.ShowMenu
50030  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "m_frmSysTray_SysTrayMouseDown")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SystrayEnter()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long
50020  If m_frmSysTray Is Nothing Then
50030   Set m_frmSysTray = New frmSysTray
50040   With m_frmSysTray
50050    .AddMenuItem App.EXEName, , True
50060    .AddMenuItem "-"
50070    For i = mnPrinter.LBound To mnPrinter.UBound
50080     .AddMenuItem mnPrinter(i).Caption
50090    Next i
50100    If mnPrinter(6).Checked = True Then
50110     .mnuSysTray(6).Checked = True
50120    End If
50130    If mnPrinter(2).Checked = True Then
50140     .mnuSysTray(2).Checked = True
50150    End If
50160    .ToolTip = Me.Caption
50170   End With
50180   If mnPrinter(2).Checked = True Then
50190     m_frmSysTray.IconHandle = imlSystray.ListImages(1).Picture.handle
50200    Else
50210     m_frmSysTray.IconHandle = imlSystray.ListImages(2).Picture.handle
50220   End If
50230   Me.Hide
50240  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "SystrayEnter")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SysTrayLeave()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If Not m_frmSysTray Is Nothing Then
50020   Unload m_frmSysTray
50030   Set m_frmSysTray = Nothing
50040  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "SysTrayLeave")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SetSystrayIcon(Index As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If Not m_frmSysTray Is Nothing Then
50020   m_frmSysTray.IconHandle = frmMain.imlSystray.ListImages(Index).Picture.handle
50030  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "SetSystrayIcon")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub ShowFrmMain()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With Me
50020   .ZOrder
50030   .WindowState = vbNormal
50040   .Show
50050  End With
50060  SysTrayLeave
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "ShowFrmMain")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtEmailAddress_KeyPress(KeyAscii As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If KeyAscii = vbKeyReturn Then
50020   SendEmail
50030  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "txtEmailAddress_KeyPress")
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
50010  SetLanguageMenu
50020 ' SetLanguageToolbar
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "ChangeLanguage")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub ShowPrinters()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  frmPrinters.Show vbModal, Me
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "ShowPrinters")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
