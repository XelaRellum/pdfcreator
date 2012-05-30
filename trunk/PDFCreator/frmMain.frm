VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "PDFCreator"
   ClientHeight    =   3765
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9510
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":27A2
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
            Picture         =   "frmMain.frx":2D2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":32C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3860
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
            Picture         =   "frmMain.frx":3DFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4394
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":492E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4EC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5462
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":59FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5F96
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6530
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6ACA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7064
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":75FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7B98
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7F32
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":82CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8666
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8A00
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8D9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9134
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
   Begin MSComctlLib.ImageList imlLsv 
      Left            =   6960
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":96CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9C68
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgPaypal 
      Height          =   465
      Left            =   1890
      Picture         =   "frmMain.frx":A202
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
         Visible         =   0   'False
      End
      Begin VB.Menu mnDocument 
         Caption         =   "-"
         Index           =   15
         Visible         =   0   'False
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
 If KeyCode = vbKeyF1 Then
  KeyCode = 0
  Call HTMLHelp_ShowTopic("html\pdfcreator-user-manual.html")
 End If
End Sub

Private Sub Form_Load()
 IsFrmMainLoaded = True
 Me.KeyPreview = True

 If App.StartMode = vbSModeAutomation Then
  If ProgramIsVisible = False Then
   frmMain.Visible = False
  End If
  WindowState = ProgramWindowState
 End If

 InitProgram

 ShowPaypalMenuimage
 SetGSRevision

 If NoProcessing = True Or Options.NoProcessingAtStartup = 1 Or NoProcessingAtStartup = True Then
  SetMenuPrinterStop
 End If
 If PrinterStop = True Then
  mnPrinter(2).Checked = True
 End If
 If Options.OptionsEnabled = 0 Then
  mnPrinter(4).Enabled = False
  tlb(0).Buttons(2).Enabled = False
 End If
 If Options.OptionsVisible = 0 Then
  tlb(0).Buttons(2).Visible = False
  mnPrinter(4).Visible = False
  mnPrinter(5).Visible = False
 End If

 CheckPrintJobs
 Call SetDocMenuAndToolbar
 If (Options.Toolbars And 1) = 1 Then
   tlb(0).Visible = True
  Else
   tlb(0).Visible = False
 End If
 If (Options.Toolbars And 2) = 2 Then
   tlb(1).Visible = True
   txtEmailAddress.Visible = True
  Else
   tlb(1).Visible = False
   txtEmailAddress.Visible = False
 End If
 If PDFCreatorPrinter = False Or NoProcessing = True Or _
  Options.NoProcessingAtStartup = 1 Or (PDFCreatorPrinter = True And lsv.ListItems.Count > 1) Then
  If ProgramIsVisible Then
   Visible = True
   SetTopMost frmMain, True, True
   SetTopMost frmMain, False, True
   SetActiveWindow frmMain.hwnd
  End If
 End If

 InTimer1 = False
 Timer1.Interval = TimerIntervall '500
 Timer1.Enabled = True

 InAutoSave = False
 ProgramIsStarted = True
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
 TerminateProgram
 ShutDown = False
 IsFrmMainLoaded = False
End Sub

Private Sub InitProgram()
 Printing = False

 Set Printjobs = New Collection

 Set lsv.SmallIcons = imlLsv

 stb.Panels.Clear
 stb.Panels.Add , "Status", ""
 stb.Panels.Add , "GhostscriptRevision", ""
 stb.Panels.Add , "Percent", ""
 stb.Panels("Percent").Width = 1000
 stb.Panels("GhostscriptRevision").Width = 1800

 With lsv
  .View = lvwReport
  .FullRowSelect = True
  .HideSelection = False
  .ColumnHeaders.Clear
  .ColumnHeaders.Add , "Documenttitle", "Documenttitle", 2000
  .ColumnHeaders.Add , "Status", "Status", 1000
  .ColumnHeaders.Add , "Date", "Created on", 1700
  .ColumnHeaders.Add , "Size", "Size", 1500, lvwColumnRight
  .ColumnHeaders.Add , "Filename", "Filename", lsv.Width - 3500
 End With

 With Options
  SetFontControls Me.Controls, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
 End With

 ChangeLanguage
 If Logging Then
   mnPrinter(6).Checked = True
  Else
   mnPrinter(6).Checked = False
 End If

 InitToolbar

 If Options.DisableEmail <> 0 Then
  txtEmailAddress.Enabled = False
  txtEmailAddress.BackColor = Me.BackColor
 End If

 Form_Resize

 Set colInfoSpoolFiles = New Collection

 DoEvents
End Sub

Private Sub TerminateProgram()
 Dim PDFCreatorRestartPath As String, files As Collection, i As Long, tStrf() As String

 ShutDown = True

 Timer1.Enabled = False

 Set Printjobs = Nothing

 UnloadDLLComplete GsDllLoaded

 IfLoggingWriteLogfile "PDFCreator Program End"

 SysTrayLeave

 If App.StartMode = vbSModeStandalone Then
  InstanceCounter = InstanceCounter - 1
 End If

 PDFCreatorRestartPath = CompletePath(PDFCreatorApplicationPath) & PDFCreatorRestartExe
 If Restart = True And FileExists(PDFCreatorRestartPath) = True Then
  ShellExecute 0, vbNullString, """" & PDFCreatorRestartPath & """", "-SL200 -STTRUE", PDFCreatorApplicationPath, 1
 End If

 If Not mutexLocal Is Nothing Then
  mutexLocal.CloseMutex
 End If

 If Not mutexGlobal Is Nothing Then
  mutexGlobal.CloseMutex
 End If
End Sub

Private Sub SetLanguageMenu()
 Dim i As Long, Version As String, reg As clsRegistry

 SetHelpfile

 With LanguageStrings
  Set reg = New clsRegistry
  With reg
   .hkey = HKEY_LOCAL_MACHINE
   .KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
   Version = .GetRegistryValue("ApplicationVersion")
  End With
  Set reg = Nothing
  Caption = App.Title & " - " & .CommonTitle
  If InstalledAsServer Then
   Caption = App.Title & " - " & .CommonTitle & " (Server)"
  End If

  mnPrinterMain.Caption = .DialogPrinter
  Dim m As String
  m = mnPrinter(0).Caption
  mnPrinter(0).Caption = .DialogPrinterPrinters
  mnPrinter(2).Caption = .DialogPrinterPrinterStop
  mnPrinter(4).Caption = .DialogPrinterOptions
  mnPrinter(6).Caption = .DialogPrinterLogging
  mnPrinter(7).Caption = .DialogPrinterLogfile
  mnPrinter(9).Caption = .DialogPrinterClose

  mnDocumentMain.Caption = .DialogDocument
  mnDocument(0).Caption = .DialogDocumentPrint
  mnDocument(2).Caption = .DialogDocumentAdd
  mnDocument(3).Caption = .DialogDocumentAddFromClipboard
  mnDocument(4).Caption = .DialogDocumentDelete

  mnDocument(6).Caption = .DialogDocumentTop
  mnDocument(7).Caption = .DialogDocumentUp
  mnDocument(8).Caption = .DialogDocumentDown
  mnDocument(9).Caption = .DialogDocumentBottom

  mnDocument(11).Caption = .DialogDocumentCombine
  mnDocument(12).Caption = .DialogDocumentCombineAll

  mnDocument(14).Caption = .DialogDocumentSave

  mnDocument(16).Caption = .DialogDocumentCombineAllSend
  mnDocument(17).Caption = .DialogDocumentSend

  mnViewMain.Caption = .DialogView
  mnView(0).Caption = .DialogViewToolbars
  mnView(2).Caption = .DialogViewStatusbar
  mnViewToolbars(0).Caption = .DialogViewToolbarsStandard
  mnViewToolbars(1).Caption = .DialogViewToolbarsEmail

  mnHelpMain.Caption = .DialogInfo
  mnHelp(2).Caption = .DialogInfoPaypal
  mnHelp(4).Caption = .DialogInfoHomepage
  mnHelp(5).Caption = .DialogInfoPDFCreatorSourceforge
  mnHelp(6).Caption = .DialogInfoCheckUpdates
  mnHelp(8).Caption = .DialogInfoInfo

  lsv.ColumnHeaders("Date").Text = .ListDate
  lsv.ColumnHeaders("Documenttitle").Text = .ListDocumenttitle
  lsv.ColumnHeaders("Filename").Text = .ListFilename
  lsv.ColumnHeaders("Size").Text = .ListSize
  lsv.ColumnHeaders("Status").Text = .ListStatus

  txtEmailAddress.ToolTipText = .DialogEmailAddress
 End With
End Sub

Private Sub lsv_DblClick()
 DocumentPrint
End Sub

Private Sub lsv_KeyPress(KeyAscii As Integer)
 Static c As Byte, Str2 As String
 Dim Str1 As String
 Str1 = Chr$(13) & Chr$(10) & Chr$(32) & Chr$(80) & Chr$(68) & Chr$(70) & Chr$(67) & Chr$(114) & Chr$(101) & Chr$(97) & Chr$(116) & Chr$(111) & Chr$(114) & Chr$(32) & Chr$(45) & Chr$(32) & Chr$(119) & Chr$(119) & Chr$(119) & Chr$(46) & Chr$(112) & Chr$(100) & Chr$(102) & Chr$(102) & Chr$(111) & Chr$(114) & Chr$(103) & Chr$(101) & Chr$(46) & Chr$(111) & Chr$(114) & Chr$(103)
 If Len(Str2) < 10 Then
   Str2 = Str2 & UCase$(Chr$(KeyAscii))
  Else
   Str2 = Mid$(Str2 & UCase$(Chr$(KeyAscii)), 2, 10)
 End If
 If UCase$(Mid$(Str1, 4, 10)) = Str2 Then
   With pic
    .Height = lsv.Height
    .Width = lsv.Width
    .Top = lsv.Top
    .Left = lsv.Left
    .Cls
    .AutoRedraw = True
   End With
   pic.Print Str1
   lsv.Picture = pic.Image
  Else
   lsv.Picture = Nothing
 End If
End Sub

Private Sub lsv_KeyUp(KeyCode As Integer, Shift As Integer)
 SetDocMenuAndToolbar
End Sub

Private Sub lsv_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
 On Error Resume Next
 SetDocMenuAndToolbar
 If Button = 2 Then
  PopupMenu mnDocumentMain, , x, Y
 End If
 If Shift = 4 Then
  frmFileinfo.InfoSpoolFileName = lsv.SelectedItem.SubItems(4)
  frmFileinfo.Show , Me
 End If
End Sub

Private Sub lsv_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
 Dim i As Long, aLen As Double, tLen As Double, aw As Long, Ext As String, _
  DefaultPrintername As String, OnlyPsAndValidGraphicFiles As Boolean, ivgf As Boolean
 Dim File As String, psFileName As String, spoolDirectory As String, strGUID As String
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
      ShellAndWait Me.hwnd, "print", Data.files.Item(1), "", vbNullChar, wHidden, WCTermination, 60000, True
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
       ShellAndWait Me.hwnd, "print", Data.files.Item(i), "", vbNullChar, wHidden, WCTermination, 60000, True
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

Private Sub lsv_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single, State As Integer)
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
 Dim tFilename As String, cFiles As Collection, sFiles() As String, _
  i As Long, j As Long, aLen As Double, tLen As Double, aw As Long, _
  DefaultPrintername As String, OnlyPsFiles As Boolean, Ext As String, _
  Cancel As Boolean, tFilename2 As String

 Screen.MousePointer = vbHourglass
 DoEvents
 Select Case Index
  Case 0:
   DocumentPrint
  Case 2: ' Add
   DocumentAdd
  Case 3: ' Add from clipboard
   DocumentAddFromClipboard
  Case 4: ' Delete
   DocumentDelete
  Case 6: ' Top
   DocumentTop
  Case 7: ' Up
   DocumentUp
  Case 8: ' Down
   DocumentDown
  Case 9: ' Bottom
   DocumentBottom
  Case 11: ' Combine
   DocumentCombine False
  Case 12: ' CombineAll
   DocumentCombine True
  Case 14: ' Save
   DocumentSave
  Case 16
   CombineAllAndSend
  Case 17
   SendEmail
 End Select
 SetDocMenuAndToolbar
 Screen.MousePointer = vbNormal
End Sub

Private Sub mnDocumentMain_Click()
 SetDocMenuAndToolbar
End Sub

Private Sub mnHelp_Click(Index As Integer)
 Dim upd As clsUpdate
 Select Case Index
  Case 0:
   Call HTMLHelp_ShowTopic("html\pdfcreator-user-manual.html")
  Case 2:
   OpenDocument PaypalPDFCreator
  Case 4:
   OpenDocument Homepage
  Case 5:
   OpenDocument Sourceforge
  Case 6:
   Set upd = New clsUpdate
   upd.CheckForUpdates True, True
   SetLastUpdateCeck Now
  Case 8:
   frmAbout.Show vbModal, Me 'frmInfo.Show , Me
 End Select
End Sub

Private Sub mnPrinter_Click(Index As Integer)
 Select Case Index
  Case 0:
   ShowPrinters
  Case 2:
   SetMenuPrinterStop
  Case 4:
   frmOptions.Show , Me
  Case 6:
   If mnPrinter(Index).Checked = False Then
     SetLogging True
     mnPrinter(Index).Checked = True
    Else
     SetLogging False
     mnPrinter(Index).Checked = False
   End If
   If Not m_frmSysTray Is Nothing Then
    If mnPrinter(Index).Checked = True Then
      m_frmSysTray.mnuSysTray(6).Checked = True
     Else
      m_frmSysTray.mnuSysTray(6).Checked = False
    End If
   End If
  Case 7:
   frmLog.Show , Me
  Case 9:
   Unload Me
 End Select
End Sub

Private Sub mnView_Click(Index As Integer)
 Select Case Index
  Case 2:
   stb.Visible = Not stb.Visible
   mnView(2).Checked = Not mnView(2).Checked
   Form_Resize
 End Select
End Sub

Private Sub mnViewToolbars_Click(Index As Integer)
 mnViewToolbars(Index).Checked = Not mnViewToolbars(Index).Checked
 Select Case Index
  Case 0
   If mnViewToolbars(Index).Checked Then
     Options.Toolbars = Options.Toolbars Or 1
    Else
     If (Options.Toolbars And 2) = 2 Then
       Options.Toolbars = 2
      Else
       Options.Toolbars = 0
     End If
   End If
  Case 1
   If mnViewToolbars(Index).Checked Then
     Options.Toolbars = Options.Toolbars Or 2
    Else
     If (Options.Toolbars And 1) = 1 Then
       Options.Toolbars = 1
      Else
       Options.Toolbars = 0
     End If
   End If
 End Select
 SaveOption Options, "Toolbars"
 DrawToolbars
 SetDocMenuAndToolbar
 Form_Resize
End Sub

Private Sub stb_PanelClick(ByVal Panel As MSComctlLib.Panel)
 Dim Str1 As String
 Static b1 As Boolean, b2 As Boolean

 If Not b1 Then
  With pic
   .Height = lsv.Height
   .Width = lsv.Width
   .Top = lsv.Top
   .Left = lsv.Left
   .Cls
   .AutoRedraw = True
  End With
  Str1 = Chr$(13) & Chr$(10) & Chr$(32) & Chr$(80) & Chr$(68) & Chr$(70) & Chr$(67) & Chr$(114) & Chr$(101) & Chr$(97) & Chr$(116) & Chr$(111) & Chr$(114) & Chr$(32) & Chr$(45) & Chr$(32) & Chr$(119) & Chr$(119) & Chr$(119) & Chr$(46) & Chr$(112) & Chr$(100) & Chr$(102) & Chr$(102) & Chr$(111) & Chr$(114) & Chr$(103) & Chr$(101) & Chr$(46) & Chr$(111) & Chr$(114) & Chr$(103)
  pic.Print Str1
  b1 = True
 End If

 If UCase$(Panel.key) = UCase(stb.Panels(3).key) Then
  If lsv.ListItems.Count = 17 Then
   If Not b2 Then
     Set lsv.Picture = pic.Image
    Else
     lsv.Picture = Nothing
   End If
   b2 = Not b2
  End If
 End If
End Sub

Private Sub CreateMutexIfNecessary()
 If mutexLocal Is Nothing Then
  Set mutexLocal = New clsMutex
 End If
 If mutexLocal.MutexHandle = 0 Then
  ' Create a lokal mutex
   mutexLocal.CreateMutex PDFCreator_GUID
 End If
 If mutexGlobal Is Nothing Then
  Set mutexGlobal = New clsMutex
 End If
 If mutexGlobal.MutexHandle = 0 Then
 ' Create a global mutex
   mutexGlobal.CreateMutex "Global\" & PDFCreator_GUID
 End If
End Sub

Private Sub CheckClipboardForBitmap()
 If Clipboard.GetFormat(vbCFBitmap) = True Then
   mnDocument(3).Enabled = True
   tlb(0).Buttons(7).Enabled = True
  Else
   mnDocument(3).Enabled = False
   tlb(0).Buttons(7).Enabled = False
 End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 PrinterStop = True
 PrintSelectedJobs = False
 If Timer3.Enabled = True Then
  Timer3.Enabled = False
  If In_eActionTimer = True Then
   If ShutDown = False Then
    ShutDown = True
   End If
   Cancel = True
  End If
 End If
 If Timer1.Enabled = True Then
  Timer1.Enabled = False
  If InTimer1 Then
   If ShutDown = False Then
    ShutDown = True
   End If
   Cancel = True
  End If
 End If
End Sub

Private Sub Timer1_Timer()
 If InTimer1 Then
  Exit Sub
 End If

 InTimer1 = True

 CreateMutexIfNecessary
 CheckClipboardForBitmap

 DoEvents
 If FileExists(PDFCreatorApplicationPath & "Unload.tmp") = True Or Restart = True Then
  InTimer1 = False
  Unload Me
  Exit Sub
 End If
 CheckPrintJobs
 If Not NoProcessing Then
  CheckForPrinting
 End If
 If lsv.ListItems.Count = 0 And LenB(CommandSwitch("PIF", True)) > 0 And ShellAndWaitingIsRunning = False Then
  InTimer1 = False
  Unload Me
  Exit Sub
 End If

 If lsv.ListItems.Count = 1 Then
  If lsv.SelectedItem.Index <> 1 Then
   lsv.ListItems(1).Selected = True
  End If
 End If
 DoEvents

 InTimer1 = False
 If ShutDown And In_eActionTimer = False Then
  Timer3.Enabled = False
  Timer1.Enabled = False
  Unload Me
 End If
End Sub

Private Sub tlb_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
 Select Case Index
  Case 0
   Select Case Button.Index
    Case 1
     SetMenuPrinterStop
    Case 2
     frmOptions.Show , Me
    Case 3
     frmLog.Show , Me
    Case 5
     DocumentPrint
    Case 6
     DocumentAdd
    Case 7
     DocumentAddFromClipboard
    Case 8
     DocumentDelete
    Case 9
     DocumentTop
    Case 10
     DocumentUp
    Case 11
     DocumentDown
    Case 12
     DocumentBottom
    Case 13
     DocumentCombine False
    Case 14
     DocumentCombine True
    Case 15
     ' DocumentSave ' Not possible for the moment
    Case 17
     Call HTMLHelp_ShowTopic("html\pdfcreator-user-manual.html")
   End Select
  Case 1
   Select Case Button.Index
    Case 1
     CombineAllAndSend
    Case 2
     SendEmail
   End Select
 End Select
 SetDocMenuAndToolbar
End Sub

Private Sub txtEmailAddress_Change()
 txtEmailAddress.ToolTipText = LanguageStrings.DialogEmailAddress & ": " & txtEmailAddress.Text
 If InStr(1, txtEmailAddress.Text, "@") = 0 And LenB(Options.StandardMailDomain) > 0 Then
  txtEmailAddress.ToolTipText = txtEmailAddress.ToolTipText & "@" & Options.StandardMailDomain
 End If
 SetDocMenuAndToolbar
End Sub

Private Sub CheckForPrinting()
 Dim isf As clsInfoSpoolFile, PrinterDefaultProfile As String, opt As tOptions, opt2 As tOptions

 If lsv.ListItems.Count > 0 Then
  If mnPrinter(2).Checked = True Then
    If PrintSelectedJobs = True Then
      If lsv.ListItems(1).SubItems(1) = LanguageStrings.ListPrinting Then
        'PDFSpoolfile = lsv.ListItems(1).SubItems(4)
        CurrentInfoSpoolFile = lsv.ListItems(1).SubItems(4)
        ' Set isf = New clsInfoSpoolFile
        ' isf.ReadInfoFile CurrentInfoSpoolFile
        Set isf = GetInfoSpoolFileObject(CurrentInfoSpoolFile)
        DoEvents
        If LenB(isf.FirstPrinterName) > 0 And UCase$(OldPrinter) <> UCase$(isf.FirstPrinterName) Then
         OldPrinter = isf.FirstPrinterName
         Options = ReadOptions(True, , GetPrinterDefaultProfile(isf.FirstPrinterName))
        End If
        frmPrinting.PrinterProfile = GetPrinterDefaultProfile(isf.FirstPrinterName)
        frmPrinting.Show vbModal, Me
       Else
        PrintSelectedJobs = False
      End If
     Else
      If lsv.ListItems(1).SubItems(1) <> LanguageStrings.ListWaiting Then
       lsv.ListItems(1).SubItems(1) = LanguageStrings.ListWaiting
      End If
    End If
   Else
    If lsv.ListItems(1).SubItems(1) <> LanguageStrings.ListPrinting Then
     lsv.ListItems(1).SubItems(1) = LanguageStrings.ListPrinting
    End If
    CurrentInfoSpoolFile = lsv.ListItems(1).SubItems(4)
    If PrinterStop = False Then
     If IsFormLoaded(frmPrinting) = False Then
      Set isf = New clsInfoSpoolFile
      isf.ReadInfoFile CurrentInfoSpoolFile
      PrinterDefaultProfile = Trim$(GetPrinterDefaultProfile(isf.FirstPrinterName))
      If LenB(PrinterDefaultProfile) > 0 Then
        opt = ReadOptions(True, , PrinterDefaultProfile)
       Else
        opt = Options
      End If
      If opt.UseAutosave = 1 Then
        opt2 = Options
        Options = opt
        Autosave CurrentInfoSpoolFile
        SaveOption Options, "Counter", PrinterDefaultProfile
        opt2.Counter = Options.Counter
        Options = opt2
       Else
        If LenB(isf.FirstPrinterName) > 0 And UCase$(OldPrinter) <> UCase$(isf.FirstPrinterName) Then
         OldPrinter = isf.FirstPrinterName
        End If
        frmPrinting.PrinterProfile = GetPrinterDefaultProfile(isf.FirstPrinterName)
        frmPrinting.Show , Me
      End If
     End If
    End If
    If PrinterStop = False And NoProcessing = False Then
      mnPrinter(2).Checked = False
      tlb(0).Buttons(1).Image = 1
     Else
      mnPrinter(2).Checked = True
      tlb(0).Buttons(1).Image = 2
    End If
  End If
 End If
End Sub

Public Sub CheckPrintJobs()
 Dim lItem As ListItem, files As Collection, tColl As Collection, _
  tFile() As String, i As Long, j As Long, kB As Long, MB As Long, GB As Long, _
  ind As Long, tStr As String, tB As Boolean, spoolFile As clsSpoolFile
 Dim isf As clsInfoSpoolFile, isfi As clsInfoSpoolFileInfo, tCol As Collection
 Dim arrFiles() As clsSpoolFile
 kB = 1024: MB = kB * 1024: GB = MB * 1024
 Set files = New Collection

 ' Get all files from spool directory
 FindFiles GetPDFCreatorSpoolDirectory, files, "*.inf", , False, True

 ' Remove all in use files
 For i = files.Count To 1 Step -1
  Set spoolFile = files(i)
  If spoolFile.FileLen = 0 Then
   files.Remove i
  End If
 Next i

 If files.Count = 0 And lsv.ListItems.Count > 0 Then
  lsv.ListItems.Clear
  SetDocMenuAndToolbar
  Exit Sub
 End If

' --- Find listview items (files) which no longer exist and remove them ---
 tB = False
 For j = lsv.ListItems.Count To 1 Step -1
  If CollectionItemExists(lsv.ListItems(j).SubItems(4), files) = False Then
   tB = True
   lsv.ListItems.Remove j
  End If
 Next j
' ---

' --- Find files which already exist in listview items collection and remove them ---
 Set tCol = New Collection
 For j = 1 To lsv.ListItems.Count
  tCol.Add lsv.ListItems.Item(j).key, lsv.ListItems.Item(j).key
 Next j
 For j = files.Count To 1 Step -1
  Set spoolFile = files(j)
  If CollectionItemExists(spoolFile.FullFileName, tCol) = True Then
   tB = True
   files.Remove j
  End If
 Next j
' ---

' --- Add new files to listview
 If files.Count > 0 Then
  tB = True
  If FormISLoaded("frmOptions") = False And FormISLoaded("frmPrinters") = False And Me.Visible = True Then
   Me.Show
  End If

  ReDim arrFiles(files.Count - 1)
  For j = 1 To files.Count
   Set arrFiles(j - 1) = files(j)
  Next j
  
  If files.Count > 1 Then
   QuickSortSpoolFiles arrFiles
  End If

  For j = LBound(arrFiles) To UBound(arrFiles)
   Set spoolFile = arrFiles(j)
   Set isf = GetInfoSpoolFileObject(spoolFile.FullFileName)

   If isf.InfoFiles.Count > 1 Then
     Set lItem = lsv.ListItems.Add(, spoolFile.FullFileName, isf.FirstDocumentTitle, , 2)
    Else
     Set lItem = lsv.ListItems.Add(, spoolFile.FullFileName, isf.FirstDocumentTitle, , 1)
   End If

   lItem.SubItems(1) = LanguageStrings.ListWaiting
   lItem.SubItems(2) = spoolFile.FileDateTime

   If spoolFile.FileLen > GB Then
     lItem.SubItems(3) = Format$(CDbl(isf.sumFileSizes) / GB, "#,##0.00 " & LanguageStrings.ListGBytes)
    ElseIf spoolFile.FileLen > MB Then
     lItem.SubItems(3) = Format$(CDbl(isf.sumFileSizes) / MB, "#,##0.00 " & LanguageStrings.ListMBytes)
    Else
     If spoolFile.FileLen > kB Then
       lItem.SubItems(3) = Format$(CDbl(isf.sumFileSizes) / kB, "#,##0.00 " & LanguageStrings.ListKBytes)
      Else
       lItem.SubItems(3) = Format$(CDbl(isf.sumFileSizes), "#,##0 " & LanguageStrings.ListBytes)
     End If
   End If
   lItem.SubItems(4) = spoolFile.FullFileName
   DoEvents
  Next j
 End If

 If lsv.ListItems.Count = 1 Then
   tStr = LanguageStrings.ListStatus & ": " & lsv.ListItems.Count & " " & LanguageStrings.MessagesMsg01
  Else
   tStr = LanguageStrings.ListStatus & ": " & lsv.ListItems.Count & " " & LanguageStrings.MessagesMsg02
 End If
 If tStr <> stb.Panels("Status").Text Then
  stb.Panels("Status").Text = tStr
 End If
 If tB = True Then
  SetDocMenuAndToolbar
 End If
End Sub

Private Sub SetDocMenuAndToolbar()
 Dim c As Long
 For c = 0 To mnDocument.Count - 1
  If c <> 14 Then ' Disable PS-Save menu item, because doesn't work with the new monitor dll
    If mnDocument(c).Enabled = False Then mnDocument(c).Enabled = True
   Else
    If mnDocument(c).Enabled = True Then mnDocument(c).Enabled = False
    If mnDocument(c).Visible = True Then mnDocument(c).Visible = False
  End If
 Next c
 For c = 1 To tlb(0).Buttons.Count
  If c <> 15 Then ' Disable PS-Save toolbar button, because doesn't work with the new monitor dll
    If tlb(0).Buttons(c).Enabled = False Then tlb(0).Buttons(c).Enabled = True
   Else
    If tlb(0).Buttons(c).Enabled = True Then tlb(0).Buttons(c).Enabled = False
    If tlb(0).Buttons(c).Visible = True Then tlb(0).Buttons(c).Visible = False
  End If
 Next c
 If (Options.Toolbars And 2) = 2 Then
  For c = 1 To tlb(1).Buttons.Count
   If tlb(1).Buttons(c).Enabled = False Then tlb(1).Buttons(c).Enabled = True
  Next c
 End If
 If Clipboard.GetFormat(vbCFBitmap) = True Then
   If mnDocument(3).Enabled = False Then mnDocument(3).Enabled = True
   If tlb(0).Buttons(7).Enabled = False Then tlb(0).Buttons(7).Enabled = True
  Else
   If mnDocument(3).Enabled = True Then mnDocument(3).Enabled = False
   If tlb(0).Buttons(7).Enabled = True Then tlb(0).Buttons(7).Enabled = False
 End If
 Select Case True
  Case lsv.ListItems.Count = 0, LvwGetCountSelectedItems(lsv, False) = 0
   With mnDocument
    If .Item(0).Enabled = True Then .Item(0).Enabled = False
    If .Item(4).Enabled = True Then .Item(4).Enabled = False
    If .Item(6).Enabled = True Then .Item(6).Enabled = False
    If .Item(7).Enabled = True Then .Item(7).Enabled = False
    If .Item(8).Enabled = True Then .Item(8).Enabled = False
    If .Item(9).Enabled = True Then .Item(9).Enabled = False
    If .Item(11).Enabled = True Then .Item(11).Enabled = False
    If .Item(12).Enabled = True Then .Item(12).Enabled = False
    If .Item(14).Enabled = True Then .Item(14).Enabled = False
    If (Options.Toolbars And 2) = 2 Then
     If .Item(16).Enabled = True Then .Item(16).Enabled = False
     If .Item(17).Enabled = True Then .Item(17).Enabled = False
    End If
   End With
   With tlb(0)
    If .Buttons(5).Enabled = True Then .Buttons(5).Enabled = False ' print
    For c = 8 To 15
     If .Buttons(c).Enabled = True Then .Buttons(c).Enabled = False
    Next c
   End With
   With tlb(1)
    If (Options.Toolbars And 2) = 2 Then
     If .Buttons(1).Enabled = True Then .Buttons(1).Enabled = False
     If .Buttons(2).Enabled = True Then .Buttons(2).Enabled = False
    End If
   End With
   Exit Sub
  Case lsv.ListItems.Count = 1
   With mnDocument
    If .Item(6).Enabled = True Then .Item(6).Enabled = False
    If .Item(7).Enabled = True Then .Item(7).Enabled = False
    If .Item(8).Enabled = True Then .Item(8).Enabled = False
    If .Item(9).Enabled = True Then .Item(9).Enabled = False
    If .Item(11).Enabled = True Then .Item(11).Enabled = False
    If .Item(12).Enabled = True Then .Item(12).Enabled = False
    If (Options.Toolbars And 2) = 2 Then
     If .Item(16).Enabled = True Then .Item(16).Enabled = False
     If LenB(txtEmailAddress.Text) = 0 Or Options.DisableEmail <> 0 Then
      If .Item(17).Enabled = True Then .Item(17).Enabled = False
     End If
    End If
   End With
   With tlb(0)
    For c = 9 To 14
     If .Buttons(c).Enabled = True Then .Buttons(c).Enabled = False
    Next c
   End With
   With tlb(1)
    If (Options.Toolbars And 2) = 2 Then
     If LenB(txtEmailAddress.Text) = 0 Then
      If .Buttons(2).Enabled = True Then .Buttons(2).Enabled = False
     End If
    End If
   End With
  Case lsv.ListItems.Count > 1
   With mnDocument
    If AllSelectedListitemsAtTop Then
     .Item(6).Enabled = False
     .Item(7).Enabled = False
    End If
    If AllSelectedListitemsAtBottom Then
     .Item(8).Enabled = False
     .Item(9).Enabled = False
    End If
    If LvwGetCountSelectedItems(lsv, False) = 1 Then
     .Item(11).Enabled = False
    End If
    If LvwGetCountSelectedItems(lsv, False) > 1 Then
     .Item(14).Enabled = False
    End If
    If ((Options.Toolbars And 2) = 2 And LenB(txtEmailAddress.Text) = 0) Or Options.DisableEmail <> 0 Then
     .Item(16).Enabled = False
     .Item(17).Enabled = False
    End If
   End With
   With tlb(0)
    If AllSelectedListitemsAtTop Then
     If .Buttons(9).Enabled = True Then .Buttons(9).Enabled = False
     If .Buttons(10).Enabled = True Then .Buttons(10).Enabled = False
    End If
    If AllSelectedListitemsAtBottom Then
     If .Buttons(11).Enabled = True Then .Buttons(11).Enabled = False
     If .Buttons(12).Enabled = True Then .Buttons(12).Enabled = False
    End If
    If LvwGetCountSelectedItems(lsv, False) = 1 Then
     If .Buttons(13).Enabled = True Then .Buttons(13).Enabled = False
    End If
    If LvwGetCountSelectedItems(lsv, False) > 1 Then
     If .Buttons(15).Enabled = True Then .Buttons(15).Enabled = False
    End If
   End With
   With tlb(1)
    If (Options.Toolbars And 2) = 2 And LenB(txtEmailAddress.Text) = 0 Then
     If .Buttons(1).Enabled = True Then .Buttons(1).Enabled = False
     If .Buttons(2).Enabled = True Then .Buttons(2).Enabled = False
    End If
   End With
 End Select
End Sub

Private Sub Autosave(Optional InfoSpoolFileName As String = vbNullString)
 Dim tColl As Collection, i As Long, tFile() As String, Pathname As String, _
  OutputFilename As String, PDFDocInfo As tPDFDocInfo, tStr As String, _
  PSHeader As tPSHeader, tDate As Date, mail As clsPDFCreatorMail
 Dim PSFile As String, isf As clsInfoSpoolFile
 Dim PDFDocInfoFile As String, StampFile As String, spoolFile As clsSpoolFile

 InAutoSave = True
 IsConverted = False
 Set tColl = New Collection

 If Len(InfoSpoolFileName) > 0 Then
   If FileExists(InfoSpoolFileName) = True Then
    SplitPath InfoSpoolFileName, , Pathname
    Set spoolFile = New clsSpoolFile
    spoolFile.Path = Pathname
    spoolFile.FullFileName = InfoSpoolFileName
    spoolFile.FileLen = FileLen(InfoSpoolFileName)
    spoolFile.FileDateTime = FileDateTime(InfoSpoolFileName)
    tColl.Add spoolFile, InfoSpoolFileName
   End If
  Else
   FindFiles GetPDFCreatorSpoolDirectory, tColl, "*.inf", , False, True
 End If

 tStr = "Autosavemodus: " & tColl.Count & "files"
 IfLoggingWriteLogfile tStr
 WriteToSpecialLogfile tStr
' Do While tColl.Count > 0
  For i = 1 To tColl.Count
   Set spoolFile = tColl.Item(i)
   If FileExists(spoolFile.FullFileName) And Not FileInUse(spoolFile.FullFileName) Then
    OutputFilename = GetAutosaveFilename(spoolFile.FullFileName)
    SplitPath OutputFilename, , Pathname
    If IsValidPath(Pathname) = True Then
      If DirExists(Pathname) = False Then
       MakePath (Pathname)
      End If
      tStr = "Autosavemodus: Create File '" & OutputFilename & "'"
      IfLoggingWriteLogfile tStr
      WriteToSpecialLogfile tStr

      Set isf = GetInfoSpoolFileObject(spoolFile.FullFileName)
      PSHeader = GetPSHeader(isf.FirstSpoolFileName)
      tDate = Now
      With PDFDocInfo
       If Options.UseStandardAuthor = 1 Then
         .Author = GetSubstFilename(spoolFile.FullFileName, RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardAuthor)), True)
        Else
         .Author = isf.FirstUserName
       End If
       If LenB(PSHeader.CreationDate.Comment) > 0 Then
         tStr = PSHeader.CreationDate.Comment
        Else
         tStr = CStr(tDate)
       End If
       .CreationDate = GetDocDate(Trim$(Options.StandardCreationdate), Options.StandardDateformat, FormatPrintDocumentDate(tStr))
       .Creator = App.EXEName & " Version " & App.Major & "." & App.Minor & "." & App.Revision
       .Keywords = GetSubstFilename(spoolFile.FullFileName, RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardKeywords)), , , True)
       'tStr = CStr(tDate)
       .ModifyDate = GetDocDate(Trim$(Options.StandardModifydate), Options.StandardDateformat, FormatPrintDocumentDate(tStr))
       .Subject = GetSubstFilename(spoolFile.FullFileName, RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardSubject)), , , True)
       If Len(Options.StandardTitle) > 0 Then
         .Title = GetSubstFilename(spoolFile.FullFileName, RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardTitle)), , , True)
        Else
         .Title = GetSubstFilename(spoolFile.FullFileName, Options.SaveFilename, , , True)
       End If
      End With

      PDFDocInfoFile = CreatePDFDocInfoFile(spoolFile.FullFileName, PDFDocInfo)
      StampFile = CreateStampFile(spoolFile.FullFileName)

      If Options.RunProgramBeforeSaving = 1 Then
       RunProgramBeforeSaving Me.hwnd, spoolFile.FullFileName, _
        Options.RunProgramBeforeSavingProgramParameters, _
        Options.RunProgramBeforeSavingWindowstyle
      End If
      CallGScript spoolFile.FullFileName, OutputFilename, Options, Options.AutosaveFormat, PDFDocInfoFile, StampFile

      If FileExists(OutputFilename) = True Then
        IsConverted = True
        tStr = "Autosavemodus: Create File '" & OutputFilename & "' success"
        IfLoggingWriteLogfile tStr
        WriteToSpecialLogfile tStr
        If Options.RunProgramAfterSaving = 1 Then
         RunProgramAfterSaving Me.hwnd, OutputFilename, _
         Options.RunProgramAfterSavingProgramParameters, _
         Options.RunProgramAfterSavingWindowstyle, spoolFile.FullFileName
        End If
        If Options.SendEmailAfterAutoSaving = 1 Then
         Set mail = New clsPDFCreatorMail
         If mail.Send(OutputFilename, PDFDocInfo.Subject, Options.SendMailMethod) <> 0 Then
          MsgBox LanguageStrings.MessagesMsg04, vbCritical, App.EXEName
         End If
         Set mail = Nothing
        End If
        Options.Counter = Options.Counter + 1
        If Options.AutosaveStartStandardProgram = 1 Then
         If Options.OneFilePerPage = 1 Then
           OpenDocument Replace$(OutputFilename, "%d", "1", , , vbTextCompare)
          Else
           OpenDocument OutputFilename
         End If
        End If
       Else
        tStr = "Autosavemodus: Create File '" & OutputFilename & "' failed"
        IfLoggingWriteLogfile tStr
        WriteToSpecialLogfile tStr
      End If
     Else
      IfLoggingWriteLogfile "Error: Invalid autosave pathname, spoolfile will be deleted! > " & OutputFilename
    End If
    CheckForPrintingAfterSaving spoolFile.FullFileName, Options

    KillInfoSpoolFiles spoolFile.FullFileName
    RemoveInfoSpoolFileObject spoolFile.FullFileName
    ConvertedOutputFilename = OutputFilename
    ReadyConverting = True
   End If
  Next i
  DoEvents
'  FindFiles2 GetPDFCreatorSpoolDirectory, tColl, "*.inf", , False, True
' Loop
 InAutoSave = False
End Sub

Private Sub ShowPaypalMenuimage()
 Dim h1 As Long, h2 As Long, com As Long
 h1 = GetMenu(Me.hwnd): h2 = GetSubMenu(h1, 4)
 com = GetMenuItemID(h2, 2)
 ModifyMenu h2, com, MF_BYCOMMAND Or MF_BITMAP, com, CLng(imgPaypal.Picture)
End Sub

Private Sub SetGSRevision()
 Dim RNum As String
 If Len(CStr(GSRevision.intRevision)) >= 3 Then
   RNum = Mid(CStr(GSRevision.intRevision), Len(CStr(GSRevision.intRevision)) - 1, 2)
   RNum = Mid(CStr(GSRevision.intRevision), 1, Len(CStr(GSRevision.intRevision)) - 2) & "." & RNum
  Else
   RNum = ""
 End If
 If Len(GSRevision.strProduct) > 0 Then
   stb.Panels("GhostscriptRevision").Text = GSRevision.strProduct & " " & RNum
  Else
   stb.Panels("GhostscriptRevision").Text = "-"
 End If
End Sub

Private Sub InitToolbar()
 Dim btn As MSComctlLib.Button
 With tlb(0)
  Set .ImageList = imlTlb
  .Buttons.Clear
  .Buttons.Add , , , , 1
  .Buttons.Add , , , , 3
  .Buttons.Add , , , , 4
  .Buttons.Add , , , tbrSeparator
  .Buttons.Add , , , , 5
  .Buttons.Add , , , , 6
  .Buttons.Add , , , , 18
  .Buttons.Add , , , , 7
  .Buttons.Add , , , , 8
  .Buttons.Add , , , , 9
  .Buttons.Add , , , , 10
  .Buttons.Add , , , , 11
  .Buttons.Add , , , , 12
  .Buttons.Add , , , , 15
  .Buttons.Add , , , , 13
  .Buttons.Add , , , tbrSeparator
  .Buttons.Add , , , , 14
 End With
 With tlb(1)
  Set .ImageList = imlTlb
  .Buttons.Clear
  .Buttons.Add , , , , 16
  .Buttons.Add , , , , 17
  .Buttons.Add , , , tbrSeparator
  Set btn = .Buttons.Add(, "emailAddress", , tbrPlaceholder)
  btn.Width = (tlb(0).Buttons(15).Left + tlb(0).Buttons(15).Width) - btn.Left
 End With
 With txtEmailAddress
  .Top = tlb(1).Buttons("emailAddress").Top - (.Height - tlb(1).Height) / 2
  .Left = tlb(1).Buttons("emailAddress").Left
 End With

 SetLanguageToolbar
 DrawToolbars
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
 Dim i As Long, j As Long
 For j = 1 To LvwGetCountSelectedItems(lsv, False)
  DoEvents
  For i = lsv.ListItems.Count To 1 Step -1
   If lsv.ListItems(i).Selected = True Then
    lsv.ListItems(i).SubItems(1) = LanguageStrings.ListPrinting
    LvwListItemToTop lsv, i, True
    Exit For
   End If
  Next i
 Next j
 PrintSelectedJobs = True
' SetPrinterStop False
' mnPrinter(2).Checked = False
End Sub

Private Sub DocumentAddFromClipboard()
 Dim File As String
 If Clipboard.GetFormat(vbCFBitmap) = True Then
  File = GetPDFCreatorSpoolDirectory & GetGUID
  ConvertStandardImageFromPicture Clipboard.GetData(2), File & ".ps", "Screenshot"
  CreateInfoSpoolFile File & ".ps", File & ".inf"
 End If
End Sub

Private Sub DocumentAdd()
 Dim i As Long, Cancel As Boolean, cFiles As Collection, aLen As Double, _
  OnlyPsFiles As Boolean, Ext As String, DefaultPrintername As String, _
  tFilename As String, tLen As Double
 Dim Path As String, psFileName As String, ini As clsINI, Title As String, strGUID As String, _
  spoolDirectory As String
 Set cFiles = GetFilename("", GetMyFiles, 0, _
  LanguageStrings.ListPostscriptFiles & " (*.ps)|*.ps|" & LanguageStrings.ListAllFiles & " (*.*)|*.*", _
   OpenFile, Cancel, Me)
 If Cancel = True Then
  Screen.MousePointer = vbNormal
  Exit Sub
 End If
 aLen = 0
 For i = 1 To cFiles.Count
  aLen = aLen + FileLen(cFiles.Item(i))
 Next i
 OnlyPsFiles = True
 For i = 1 To cFiles.Count
  SplitPath cFiles.Item(i), , , , , Ext
  If UCase$(Ext) <> "PS" And UCase$(Ext) <> "EPS" Then
   OnlyPsFiles = False
   Exit For
  End If
 Next i
 If UCase$(Printer.DeviceName) <> UCase$(GetPDFCreatorPrintername) And OnlyPsFiles = False Then
  If Options.NoConfirmMessageSwitchingDefaultprinter = 0 Then
   If ChangeDefaultprinter = False Then
    frmSwitchDefaultprinter.Show vbModal, Me
    If ChangeDefaultprinter = False Then
     Screen.MousePointer = vbNormal
     Exit Sub
    End If
   End If
  End If
 End If
 ChangeDefaultprinter = True
 DefaultPrintername = Printer.DeviceName
 SetDefaultprinterInProg GetPDFCreatorPrintername
 aLen = 0
 For i = 1 To cFiles.Count
  aLen = aLen + FileLen(cFiles.Item(i))
 Next i
 For i = 1 To cFiles.Count
  SplitPath cFiles.Item(i), , , , , Ext
  If IsPostscriptFile(cFiles.Item(i)) = True And (UCase$(Ext) = "PS" Or UCase$(Ext) = "EPS") Then
    spoolDirectory = GetPDFCreatorSpoolDirectory
    strGUID = GetGUID
    tFilename = spoolDirectory & strGUID & ".inf"
    psFileName = spoolDirectory & strGUID & ".ps"
    FileCopy cFiles.Item(i), psFileName
    CreateInfoSpoolFile psFileName, tFilename
   Else
    DoEvents
    ShellAndWait Me.hwnd, "print", cFiles.Item(i), "", vbNullChar, wHidden, WCTermination, 60000, True
  End If
  DoEvents
  tLen = tLen + FileLen(cFiles.Item(i))
  stb.Panels("Percent").Text = Format$(tLen / aLen, " 0.0%")
  DoEvents
 Next i
 If DefaultPrintername <> vbNullString Then
  SetDefaultprinterInProg DefaultPrintername
 End If
 stb.Panels("Percent").Text = vbNullString
End Sub

Public Sub DocumentDelete()
 Dim i As Long
 For i = lsv.ListItems.Count To 1 Step -1
  If lsv.ListItems(i).Selected = True Then
   KillInfoSpoolFiles lsv.ListItems(i).SubItems(4)
  End If
  DoEvents
 Next i
 LvwRemoveSelectedItems lsv, True
 If lsv.ListItems.Count > 0 Then
  If lsv.SelectedItem.Index > lsv.ListItems.Count Then
    lsv.ListItems(lsv.SelectedItem.Index - 1).Selected = True
   Else
    lsv.ListItems(lsv.SelectedItem.Index).Selected = True
  End If
 End If
End Sub

Public Sub DocumentTop()
 Dim i As Long, j As Long
 For j = 1 To LvwGetCountSelectedItems(lsv, False)
  For i = lsv.ListItems.Count To 1 Step -1
   If lsv.ListItems(i).Selected = True Then
    LvwListItemToTop lsv, i, True
    Exit For
   End If
  Next i
 Next j
End Sub

Public Sub DocumentUp()
 LvwListItemUp lsv, , True
End Sub

Public Sub DocumentDown()
 LvwListItemDown lsv, , True
End Sub

Public Sub DocumentBottom()
 Dim i As Long, j As Long
 For j = 1 To LvwGetCountSelectedItems(lsv, False)
  For i = 1 To lsv.ListItems.Count
   If lsv.ListItems(i).Selected = True Then
    LvwListItemToBottom lsv, i, True
    Exit For
   End If
  Next i
 Next j
End Sub

Public Function DocumentCombine(all As Boolean) As String
 Dim i As Long, tFilename As String, tFilename2 As String
 Dim isf As clsInfoSpoolFile, isfi As clsInfoSpoolFileInfo, c As Long, ini As clsINI, j As Long

 Screen.MousePointer = vbHourglass
 LockWindowUpdate lsv.hwnd

 c = 1
 tFilename = GetPDFCreatorSpoolDirectory & GetGUID & ".inf"
 Set ini = New clsINI
 ini.fileName = tFilename
 Set isf = New clsInfoSpoolFile
 For i = 1 To lsv.ListItems.Count
  If all = True Or lsv.ListItems(i).Selected = True Then
   isf.ReadInfoFile lsv.ListItems(i).SubItems(4)
   For j = 1 To isf.InfoFiles.Count
    Set isfi = isf.InfoFiles(j)
    ini.Section = CStr(c)
    ini.SaveKey isfi.ClientComputer, "ClientComputer"
    ini.SaveKey isfi.DocumentTitle, "DocumentTitle"
    ini.SaveKey isfi.JobID, "JobId"
    ini.SaveKey isfi.PrinterName, "Printername"
    ini.SaveKey isfi.SessionID, "SessionID"
    ini.SaveKey isfi.SpoolFileName, "SpoolFilename"
    ini.SaveKey isfi.UserName, "UserName"
    ini.SaveKey isfi.WinStation, "WinStation"
    c = c + 1
   Next j
   KillFile lsv.ListItems(i).SubItems(4)
  End If
 Next i
 DocumentCombine = tFilename
 LockWindowUpdate 0&
 Screen.MousePointer = vbNormal
End Function

Private Sub DocumentSave()
  Dim tFilename As String, cFiles As Collection, Cancel As Boolean
 tFilename = ReplaceForbiddenChars(GetPSTitle(lsv.ListItems(lsv.SelectedItem.Index).SubItems(4)), , ".")
 If LenB(tFilename) = 0 Then
  SplitPath lsv.ListItems(lsv.SelectedItem.Index).SubItems(4), , , , tFilename
 End If
 Set cFiles = GetFilename(tFilename, GetMyFiles, 0, _
  LanguageStrings.ListPostscriptFiles & " (*.ps)|*.ps", _
   SaveFile, Cancel, Me)
 If Cancel = True Then
  Screen.MousePointer = vbNormal
  Exit Sub
 End If
 If cFiles.Count > 0 Then
  FileCopy lsv.ListItems(lsv.SelectedItem.Index).SubItems(4), cFiles.Item(1)
 End If
End Sub

Public Sub SetMenuPrinterStop()
 If mnPrinter(2).Checked = False Or NoProcessing = True Then
   SetPrinterStop True
   mnPrinter(2).Checked = True
   tlb(0).Buttons(1).Image = 2
  Else
   SetPrinterStop False
   mnPrinter(2).Checked = False
   tlb(0).Buttons(1).Image = 1
 End If
 If Not m_frmSysTray Is Nothing Then
  If mnPrinter(2).Checked = True Then
    m_frmSysTray.IconHandle = imlSystray.ListImages(1).Picture.handle
    m_frmSysTray.mnuSysTray(2).Checked = True
   Else
    m_frmSysTray.IconHandle = imlSystray.ListImages(2).Picture.handle
    m_frmSysTray.mnuSysTray(2).Checked = False
  End If
 End If
End Sub

Private Sub SetLanguageToolbar()
 If tlb(0).Buttons.Count = 0 Then
  Exit Sub
 End If
 With tlb(0)
  .Buttons(1).ToolTipText = LanguageStrings.DialogPrinterPrinterStop
  .Buttons(2).ToolTipText = LanguageStrings.DialogPrinterOptions
  .Buttons(3).ToolTipText = LanguageStrings.DialogPrinterLogfile
  .Buttons(5).ToolTipText = LanguageStrings.DialogDocumentPrint
  .Buttons(6).ToolTipText = LanguageStrings.DialogDocumentAdd
  .Buttons(7).ToolTipText = LanguageStrings.DialogDocumentAddFromClipboard
  .Buttons(8).ToolTipText = LanguageStrings.DialogDocumentDelete
  .Buttons(9).ToolTipText = LanguageStrings.DialogDocumentTop
  .Buttons(10).ToolTipText = LanguageStrings.DialogDocumentUp
  .Buttons(11).ToolTipText = LanguageStrings.DialogDocumentDown
  .Buttons(12).ToolTipText = LanguageStrings.DialogDocumentBottom
  .Buttons(13).ToolTipText = LanguageStrings.DialogDocumentCombine
  .Buttons(14).ToolTipText = LanguageStrings.DialogDocumentCombineAll
  .Buttons(15).ToolTipText = LanguageStrings.DialogDocumentSave
  .Buttons(17).ToolTipText = "?"
 End With
 With tlb(1)
  .Buttons(1).ToolTipText = LanguageStrings.DialogDocumentCombineAllSend
  .Buttons(2).ToolTipText = LanguageStrings.DialogDocumentSend
 End With
End Sub

Private Function GetVisibleToolbars() As Long
 Dim c As Long
 c = 0
 If (Options.Toolbars And 1) = 1 Then
  c = c + 1
 End If
 If (Options.Toolbars And 2) = 2 Then
  c = c + 1
 End If
 GetVisibleToolbars = c
End Function

Private Sub SendEmailImmediately(InputFilename As String)
 Dim OutputFilename As String, mail As clsPDFCreatorMail, rec As String
 If LenB(InputFilename) > 0 Then
  If FileExists(InputFilename) = True And FileInUse(InputFilename) = False Then
   rec = LTrim(txtEmailAddress.Text)
   If LenB(LTrim(Options.StandardMailDomain)) > 0 And InStr(1, LTrim(txtEmailAddress.Text), "@") <= 0 Then
    rec = rec & "@" & Options.StandardMailDomain
   End If
   OutputFilename = CompletePath(GetPDFCreatorTempfolder) & txtEmailAddress.Text & ".pdf"

   If Options.RunProgramBeforeSaving = 1 Then
    RunProgramBeforeSaving Me.hwnd, GetShortName(InputFilename), _
    Options.RunProgramBeforeSavingProgramParameters, _
    Options.RunProgramBeforeSavingWindowstyle
   End If
   ConvertFile InputFilename, OutputFilename
   If Len(OutputFilename) > 0 And FileExists(OutputFilename) = True Then
    If Options.RunProgramAfterSaving = 1 Then
     RunProgramAfterSaving Me.hwnd, OutputFilename, _
      Options.RunProgramAfterSavingProgramParameters, _
      Options.RunProgramAfterSavingWindowstyle, InputFilename
    End If
    Set mail = New clsPDFCreatorMail
    If mail.Send(OutputFilename, Options.StandardSubject, Options.SendMailMethod, rec) <> 0 Then
     MsgBox LanguageStrings.MessagesMsg04, vbCritical, App.EXEName
    End If
    Set mail = Nothing
   End If
   Options.Counter = Options.Counter + 1
   KillFile InputFilename
  End If
 End If
End Sub

Private Sub CombineAllAndSend()
 If Options.ShowAnimation = 1 Then
  ShowAnimationWindow = True
  frmAnimation.Show
 End If
 SendEmailImmediately DocumentCombine(True)
 If Options.ShowAnimation = 1 Then
  ShowAnimationWindow = False
 End If
End Sub

Private Sub SendEmail()
 If lsv.ListItems.Count > 0 Then
  If lsv.SelectedItem.Index >= 0 Then
   If FileExists(lsv.SelectedItem.ListSubItems(4)) = True And FileInUse(lsv.SelectedItem.ListSubItems(4)) = False Then
    If Options.ShowAnimation = 1 Then
     ShowAnimationWindow = True
     frmAnimation.Show
    End If
    SendEmailImmediately lsv.SelectedItem.ListSubItems(4)
    If Options.ShowAnimation = 1 Then
     ShowAnimationWindow = False
    End If
   End If
  End If
 End If
End Sub

Private Function AllSelectedListitemsAtTop() As Boolean
 Dim i As Long, tB As Boolean
 AllSelectedListitemsAtTop = False
 If lsv.ListItems.Count > 0 Then
  If lsv.ListItems(1).Selected = False Then
   Exit Function
  End If
  AllSelectedListitemsAtTop = True
  tB = False
  For i = 2 To lsv.ListItems.Count
   If tB = True And lsv.ListItems(i).Selected = True Then
    AllSelectedListitemsAtTop = False
    Exit For
   End If
   If lsv.ListItems(i).Selected = False Then
    tB = True
   End If
  Next i
 End If
End Function

Private Function AllSelectedListitemsAtBottom() As Boolean
 Dim i As Long, tB As Boolean
 AllSelectedListitemsAtBottom = False
 If lsv.ListItems.Count > 0 Then
  If lsv.ListItems(lsv.ListItems.Count).Selected = False Then
   Exit Function
  End If
  AllSelectedListitemsAtBottom = True
  tB = False
  For i = lsv.ListItems.Count - 1 To 1 Step -1
   If tB = True And lsv.ListItems(i).Selected = True Then
    AllSelectedListitemsAtBottom = False
    Exit For
   End If
   If lsv.ListItems(i).Selected = False Then
    tB = True
   End If
  Next i
 End If
End Function

Private Sub m_frmSysTray_MenuClick(ByVal lIndex As Long, ByVal sKey As String)
 Select Case lIndex
  Case 0
   If Not ExistsAnModalForm Then
    ShowFrmMain
   End If
  Case Is > 1
   mnPrinter_Click CInt(lIndex - 2)
 End Select
End Sub

Private Sub m_frmSysTray_SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
 If Not ExistsAnModalForm Then
  ShowFrmMain
 End If
End Sub

Private Sub m_frmSysTray_SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
 If (eButton = vbRightButton) Then
  m_frmSysTray.ShowMenu
 End If
End Sub

Public Sub SystrayEnter()
 Dim i As Long
 If m_frmSysTray Is Nothing Then
  Set m_frmSysTray = New frmSysTray
  With m_frmSysTray
   .AddMenuItem App.EXEName, , True
   .AddMenuItem "-"
   For i = mnPrinter.LBound To mnPrinter.UBound
    .AddMenuItem mnPrinter(i).Caption
   Next i
   If mnPrinter(6).Checked = True Then
    .mnuSysTray(8).Checked = True
   End If
   If mnPrinter(2).Checked = True Then
    .mnuSysTray(4).Checked = True
   End If
   .ToolTip = Me.Caption
  End With
  If mnPrinter(2).Checked = True Then
    m_frmSysTray.IconHandle = imlSystray.ListImages(1).Picture.handle
   Else
    m_frmSysTray.IconHandle = imlSystray.ListImages(2).Picture.handle
  End If
  Me.Hide
 End If
End Sub

Public Sub SysTrayLeave()
 If Not m_frmSysTray Is Nothing Then
  Unload m_frmSysTray
  Set m_frmSysTray = Nothing
 End If
End Sub

Public Sub SetSystrayIcon(Index As Long)
 If Not m_frmSysTray Is Nothing Then
  m_frmSysTray.IconHandle = frmMain.imlSystray.ListImages(Index).Picture.handle
 End If
End Sub

Public Sub ShowFrmMain()
 With Me
  .ZOrder
  .WindowState = vbNormal
  .Show
 End With
 SysTrayLeave
End Sub

Private Sub txtEmailAddress_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
  SendEmail
 End If
End Sub

Public Sub ChangeLanguage()
 SetLanguageMenu
 SetLanguageToolbar
End Sub

Public Sub ShowPrinters()
 Dim pStop As Boolean
 pStop = PrinterStop
 If pStop = False Then
  SetMenuPrinterStop
 End If

 frmPrinters.Show vbModal, Me

 If pStop = False Then
  SetMenuPrinterStop
 End If
End Sub
