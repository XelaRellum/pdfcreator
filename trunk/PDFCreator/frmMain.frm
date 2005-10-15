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
            Picture         =   "frmMain.frx":548A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5A24
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5FBE
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
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6558
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6AF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":708C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7626
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7780
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7D1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":82B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":884E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8BE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8F82
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":931C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":96B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9A50
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9DEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A384
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A71E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AAB8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   630
      Top             =   2940
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
      Top             =   630
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
      Height          =   330
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   582
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
      Top             =   330
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      _Version        =   393216
   End
   Begin VB.Image imgPaypal 
      Height          =   465
      Left            =   1890
      Picture         =   "frmMain.frx":AE52
      Top             =   2940
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Menu mnPrinterMain 
      Caption         =   "Printer"
      Begin VB.Menu mnPrinter 
         Caption         =   "Printer stop "
         Index           =   0
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnPrinter 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnPrinter 
         Caption         =   "Options"
         Index           =   2
         Shortcut        =   ^O
      End
      Begin VB.Menu mnPrinter 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnPrinter 
         Caption         =   "Logging"
         Index           =   4
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnPrinter 
         Caption         =   "Logfile"
         Index           =   5
         Shortcut        =   ^L
      End
      Begin VB.Menu mnPrinter 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnPrinter 
         Caption         =   "Close"
         Index           =   7
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
         Caption         =   "Delete"
         Index           =   3
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnDocument 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Top"
         Index           =   5
         Shortcut        =   ^T
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Up"
         Index           =   6
         Shortcut        =   ^U
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Down"
         Index           =   7
         Shortcut        =   ^D
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Bottom"
         Index           =   8
         Shortcut        =   ^B
      End
      Begin VB.Menu mnDocument 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Combine"
         Index           =   10
         Shortcut        =   ^C
      End
      Begin VB.Menu mnDocument 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Save"
         Index           =   12
         Shortcut        =   ^S
      End
      Begin VB.Menu mnDocument 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Combine all"
         Index           =   14
         Shortcut        =   ^A
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Combine all and send"
         Index           =   15
         Shortcut        =   ^F
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Send"
         Index           =   16
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnViewMain 
      Caption         =   "View"
      Begin VB.Menu mnView 
         Caption         =   "Statusbar"
         Checked         =   -1  'True
         Index           =   0
      End
   End
   Begin VB.Menu mnLanguageMain 
      Caption         =   "Language"
      Begin VB.Menu mnLanguage 
         Caption         =   "-"
         Index           =   0
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

Private WithEvents dl As clsDownload
Attribute dl.VB_VarHelpID = -1
Private WithEvents m_frmSysTray As frmSysTray
Attribute m_frmSysTray.VB_VarHelpID = -1

Private Const TimerIntervall = 500

Private Printjobs As Collection

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF1 Then
  KeyCode = 0
  Call HTMLHelp_ShowTopic("html\welcome.htm")
 End If
End Sub

Private Sub Form_Load()
 Me.KeyPreview = True

 If App.StartMode = vbSModeAutomation Then
  If ProgramIsVisible = False Then
   frmMain.Visible = False
  End If
  WindowState = ProgramWindowState
 End If

 ReadAllLanguages LanguagePath
 InitProgram

 ShowPaypalMenuimage
 SetGSRevision

 If NoProcessing = True Or Options.NoProcessingAtStartup = 1 Or NoProcessingAtStartup = True Then
  SetMenuPrinterStop
 End If
 If PrinterStop = True Then
  mnPrinter(0).Checked = True
 End If
 If Options.OptionsEnabled = 0 Then
  mnPrinter(2).Enabled = False
  tlb(0).Buttons(2).Enabled = False
 End If
 If Options.OptionsVisible = 0 Then
  tlb(0).Buttons(2).Visible = False
  mnPrinter(2).Visible = False
  mnPrinter(3).Visible = False
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

 ' Only for the first time set interval to 10 ms
 Timer1.Interval = 10
 Timer1.Enabled = True
 Timer2.Interval = 100
 Timer2.Enabled = True

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
End Sub

Private Sub InitProgram()
 Printing = False

 Set Printjobs = New Collection

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
  SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
 End With

 SetLanguageMenu
 If Options.Logging = 1 Then
   mnPrinter(4).Checked = True
  Else
   mnPrinter(4).Checked = False
 End If

 InitToolbar

 If Options.DisableEmail <> 0 Then
  txtEmailAddress.Enabled = False
  txtEmailAddress.BackColor = Me.BackColor
 End If
 txtEmailAddress.ToolTipText = LanguageStrings.DialogEmailAddress

 Form_Resize

 DoEvents
End Sub

Private Sub TerminateProgram()
 Dim PDFSpoolerPath As String, Files As Collection, i As Long, tStrf() As String
 Timer1.Enabled = False
 Timer2.Enabled = False

 Set Printjobs = Nothing

 If Not mutexLocal Is Nothing Then
  mutexLocal.CloseMutex
  Set mutexLocal = Nothing
 End If

 If Not mutexGlobal Is Nothing Then
  mutexGlobal.CloseMutex
  Set mutexGlobal = Nothing
 End If

 UnloadDLLComplete GsDllLoaded

 FindFiles CompletePath(GetPDFCreatorTempfolder), Files, "*.pdf", , False, True
 For i = 1 To Files.Count
  tStrf = Split(Files(i), "|")
  If FileExists(tStrf(1)) And Not FileInUse(tStrf(1)) Then
   KillFile tStrf(1)
  End If
 Next i

 IfLoggingWriteLogfile "PDFCreator Program End"
 SysTrayLeave
 If App.StartMode = vbSModeStandalone Then
  InstanceCounter = InstanceCounter - 1
 End If
 PDFSpoolerPath = GetPDFCreatorApplicationPath & "PDFSpooler.exe"
 If Restart = True And FileExists(PDFSpoolerPath) = True Then
  ShellExecute 0, vbNullString, """" & PDFSpoolerPath & """", "-SL200 -STTRUE", GetPDFCreatorApplicationPath, 1
 End If
End Sub

Private Function GetAllLanguagesFiles(LanguagePath As String) As Collection
 Dim tColl1 As Collection, tColl2 As Collection, i As Long, tStrf() As String, ini As clsINI, _
  Languagename As String
 Set GetAllLanguagesFiles = New Collection
 Set tColl1 = GetFiles(LanguagePath, "*.ini", SortedByName)
 Set tColl2 = New Collection
 For i = 1 To tColl1.Count
  tStrf = Split(tColl1(i), "|")
  Set ini = New clsINI
  ini.Filename = tStrf(1)
  ini.Section = "Common"
  Languagename = ini.GetKeyFromSection("Languagename")
  If Len(Languagename) = 0 Then
   Languagename = "No name available."
  End If
  AddSortedStr tColl2, Languagename & "|" & tStrf(1)
 Next i
 For i = 1 To tColl2.Count
  tStrf() = Split(tColl2(i), "|")
  GetAllLanguagesFiles.Add tStrf(1)
 Next i
End Function

Private Sub ReadAllLanguages(LanguagePath As String)
 Dim Languagename As String, ini As clsINI, LangFiles As Collection, _
  i As Long, Version As String, Filename As String
 mnLanguage(0).Caption = "No languages available."

 Set LangFiles = GetAllLanguagesFiles(LanguagePath)
 Set ini = New clsINI
 For i = 1 To LangFiles.Count
  ini.Filename = LangFiles.Item(i)
  ini.Section = "Common"
  Languagename = ini.GetKeyFromSection("Languagename")
  Version = ini.GetKeyFromSection("Version")
  If Len(Languagename) = 0 Then
   Languagename = "No name available."
  End If
  Load mnLanguage(mnLanguage.Count)
  If IsCompatibleLanguageVersion(Version) = True Then
    mnLanguage(mnLanguage.Count - 1).Caption = Languagename
   Else
    mnLanguage(mnLanguage.Count - 1).Caption = Languagename & " [" & Version & "]"
  End If
  mnLanguage(mnLanguage.Count - 1).Tag = LangFiles.Item(i)
  SplitPath LangFiles.Item(i), , , , Filename
  If UCase$(Options.Language) = UCase$(Filename) Then
   mnLanguage(mnLanguage.Count - 1).Checked = True
  End If
  DoEvents
 Next i

 If mnLanguage.Count > 1 Then
  mnLanguage(0).Caption = "No languages available."
  mnLanguage(0).Visible = False
 End If
 Set ini = Nothing
End Sub

Private Sub SetLanguageMenu()
 Dim i As Long, Version As String, reg As clsRegistry, Filename As String

 For i = mnLanguage.LBound To mnLanguage.UBound
  If UCase$(Languagefile) = UCase$(mnLanguage.Item(i).Tag) Then
    mnLanguage.Item(i).Checked = True
    SplitPath Languagefile, , , , Filename
   Else
    mnLanguage.Item(i).Checked = False
  End If
 Next i

 Select Case UCase$(Filename)
  Case "GERMAN"
   HelpFile = GetPDFCreatorApplicationPath & "PDFCreator_german.chm"
   If Not FileExists(HelpFile) Then
    HelpFile = GetPDFCreatorApplicationPath & "PDFCreator_english.chm"
   End If
  Case Else
   HelpFile = GetPDFCreatorApplicationPath & "PDFCreator_english.chm"
 End Select

 If Not FileExists(HelpFile) Then
  MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & vbCrLf & HelpFile, vbExclamation
  HelpFile = GetPDFCreatorApplicationPath & "PDFCreator_english.chm"
 End If

 With LanguageStrings
  Set reg = New clsRegistry
  With reg
   .hkey = HKEY_LOCAL_MACHINE
   .KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
   Version = .GetRegistryValue("ApplicationVersion")
  End With
  Set reg = Nothing
  Caption = App.Title & " - " & .CommonTitle

  mnPrinterMain.Caption = .DialogPrinter
  mnPrinter(0).Caption = .DialogPrinterPrinterStop
  mnPrinter(2).Caption = .DialogPrinterOptions
  mnPrinter(4).Caption = .DialogPrinterLogging
  mnPrinter(5).Caption = .DialogPrinterLogfile
  mnPrinter(7).Caption = .DialogPrinterClose

  mnDocumentMain.Caption = .DialogDocument
  mnDocument(0).Caption = .DialogDocumentPrint
  mnDocument(2).Caption = .DialogDocumentAdd
  mnDocument(3).Caption = .DialogDocumentDelete
  mnDocument(5).Caption = .DialogDocumentTop
  mnDocument(6).Caption = .DialogDocumentUp
  mnDocument(7).Caption = .DialogDocumentDown
  mnDocument(8).Caption = .DialogDocumentBottom
  mnDocument(10).Caption = .DialogDocumentCombine
  mnDocument(12).Caption = .DialogDocumentSave

  mnDocument(14).Caption = .DialogDocumentCombineAll
  mnDocument(15).Caption = .DialogDocumentCombineAllSend
  mnDocument(16).Caption = .DialogDocumentSend

  mnViewMain.Caption = .DialogView
  mnView(0).Caption = .DialogViewStatusbar

  mnLanguageMain.Caption = .DialogLanguage

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
 End With
End Sub

Private Sub lsv_DblClick()
 DocumentPrint
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
  frmFileinfo.SpoolFilename = lsv.SelectedItem.SubItems(4)
  frmFileinfo.Show , Me
 End If
End Sub

Private Sub lsv_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
 Dim tFilename As String, i As Long, aLen As Double, tLen As Double, aw As Long, _
  DefaultPrintername As String, OnlyPsFiles As Boolean, Ext As String
 On Error Resume Next
 If Data.GetFormat(vbCFFiles) Then
  DefaultPrintername = ""
  If Data.Files.Count = 1 Then
    SplitPath Data.Files.Item(1), , , , , Ext
    If IsPostscriptFile(Data.Files.Item(1)) = True And (UCase$(Ext) = "PS" Or UCase$(Ext) = "EPS") Then
      tFilename = GetTempFile(GetPDFCreatorTempfolder & PDFCreatorSpoolDirectory, "~PS")
      FileCopy Data.Files.Item(1), tFilename
     Else
      If UCase$(Printer.DeviceName) <> UCase$(GetPDFCreatorPrintername) Then
       If Options.NoConfirmMessageSwitchingDefaultprinter = 0 Then
        If ChangeDefaultprinter = False Then
         frmSwitchDefaultprinter.Show vbModal, Me
         If ChangeDefaultprinter = False Then
          Exit Sub
         End If
        End If
       End If
      End If
      DefaultPrintername = Printer.DeviceName
      SetDefaultprinterInProg GetPDFCreatorPrintername
      ShellAndWait Me.hwnd, "print", Data.Files.Item(1), "", vbNullChar, wHidden, WCTermination, 60000, True
      If DefaultPrintername <> vbNullString Then
       SetDefaultprinterInProg DefaultPrintername
      End If
    End If
    DoEvents
   Else
    OnlyPsFiles = True
    For i = 1 To Data.Files.Count
     SplitPath Data.Files.Item(i), , , , , Ext
     If UCase$(Ext) <> "PS" And UCase$(Ext) <> "EPS" Then
      OnlyPsFiles = False
      Exit For
     End If
    Next i
    If UCase$(Printer.DeviceName) <> UCase$(GetPDFCreatorPrintername) And OnlyPsFiles = False Then
     If ChangeDefaultprinter = False Then
      aw = MsgBox(LanguageStrings.MessagesMsg35, vbOKCancel + vbInformation)
      If aw <> vbOK Then
       Exit Sub
      End If
     End If
    End If
    ChangeDefaultprinter = True
    DefaultPrintername = Printer.DeviceName
    SetDefaultprinterInProg GetPDFCreatorPrintername
    aLen = 0
    For i = 1 To Data.Files.Count
     aLen = aLen + FileLen(Data.Files.Item(i))
    Next i
    For i = 1 To Data.Files.Count
     If IsPostscriptFile(Data.Files.Item(i)) = True And (UCase$(Ext) = "PS" Or UCase$(Ext) = "EPS") Then
       tFilename = GetTempFile(GetPDFCreatorTempfolder, "~PD")
       DoEvents
       FileCopy Data.Files.Item(i), tFilename
      Else
       DoEvents
'       PrintDocument data.Files.item(i)
       ShellAndWait Me.hwnd, "print", Data.Files.Item(i), "", vbNullChar, wHidden, WCTermination, 60000, True
       DoEvents
     End If
     tLen = tLen + FileLen(Data.Files.Item(i))
     stb.Panels("Percent").Text = Format$(tLen / aLen, " 0.0%")
     DoEvents
    Next i
    If DefaultPrintername <> vbNullString Then
     SetDefaultprinterInProg DefaultPrintername
    End If
  End If
 End If
End Sub

Private Sub lsv_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single, State As Integer)
 Dim i As Long, Ext As String
 On Error Resume Next
 If Data.GetFormat(vbCFFiles) Then
  Effect = ccOLEDropEffectCopy
  For i = 1 To Data.Files.Count
   SplitPath Data.Files.Item(i), , , , , Ext
   If IsFilePrintable(Data.Files.Item(i)) = False And (IsPostscriptFile(Data.Files.Item(i)) = False _
    Or (IsPostscriptFile(Data.Files.Item(i)) = True And UCase$(Ext) <> "PS" And UCase$(Ext) <> "EPS")) Then
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

 Timer1.Enabled = False
 Screen.MousePointer = vbHourglass
 DoEvents
 Select Case Index
  Case 0:
   DocumentPrint
  Case 2: ' Add
   DocumentAdd
  Case 3: ' Delete
   DocumentDelete
  Case 5: ' Top
   DocumentTop
  Case 6: ' Up
   DocumentUp
  Case 7: ' Down
   DocumentDown
  Case 8: ' Bottom
   DocumentBottom
  Case 10: ' Combine
   DocumentCombine
  Case 12: ' Save
   DocumentSave
  Case 14
   CombineAll
  Case 15
   CombineAllAndSend
  Case 16
   SendEmail
 End Select
 SetDocMenuAndToolbar
 Screen.MousePointer = vbNormal
 Timer1.Enabled = True
End Sub

Private Sub mnDocumentMain_Click()
 SetDocMenuAndToolbar
End Sub

Private Sub mnHelp_Click(Index As Integer)
 Dim updStr As String, updStrA() As String, aw As Long
 Set dl = New clsDownload
 Select Case Index
  Case 0:
   Call HTMLHelp_ShowTopic("html\welcome.htm")
  Case 2:
   OpenDocument Paypal
  Case 4:
   OpenDocument Homepage
  Case 5:
   OpenDocument Sourceforge
  Case 6:
   updStr = dl.DownloadString(UpdateURL)
'   updStr = dl.DownloadString("http://localhost:8080/update.txt")
   If Len(updStr) > 0 Then
     If CheckPDFCreatorVersion(updStr) > 0 Then
       updStrA = Split(updStr, ".")
       If updStrA(3) = 0 Then
         updStr = "[" & updStrA(0) & "." & updStrA(1) & "." & updStrA(2) & "]"
        Else
         updStr = "[" & updStrA(0) & "." & updStrA(1) & "." & updStrA(2) & " Beta " & updStrA(3) & "]"
       End If
       aw = MsgBox(Replace$(LanguageStrings.MessagesMsg32, "%1", updStr), vbYesNo + vbQuestion)
       If aw = vbYes Then
        OpenDocument "http://www.sourceforge.net/project/showfiles.php?group_id=57796"
       End If
      Else
       MsgBox LanguageStrings.MessagesMsg33, vbOKOnly + vbInformation
     End If
    Else
     MsgBox LanguageStrings.MessagesMsg31 & ": " & dl.ErrorDescription & " [" & dl.ErrorNumber & "]", vbOKOnly + vbExclamation
   End If
  Case 8:
   frmInfo.Show vbModal, Me
 End Select
End Sub

Private Sub mnPrinter_Click(Index As Integer)
 Select Case Index
  Case 0:
   SetMenuPrinterStop
  Case 2:
   frmOptions.Show , Me
  Case 4:
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
  Case 5:
   frmLog.Show , Me
  Case 7:
   Unload Me
 End Select
End Sub

Private Sub mnLanguage_Click(Index As Integer)
 Dim File As String
 Screen.MousePointer = vbHourglass
 LoadLanguage mnLanguage(Index).Tag
 Languagefile = mnLanguage(Index).Tag
 SetLanguageMenu
 SplitPath Languagefile, , , , File
 SetLanguage File
 ShowPaypalMenuimage
 Me.Refresh
 Screen.MousePointer = vbNormal
End Sub

Private Sub mnView_Click(Index As Integer)
 Select Case Index
  Case 0:
   stb.Visible = Not stb.Visible
   mnView(0).Checked = Not mnView(0).Checked
   Form_Resize
 End Select
End Sub

Private Sub Timer1_Timer()
 Timer1.Enabled = False
 DoEvents
 If FileExists(GetPDFCreatorApplicationPath & "Unload.tmp") = True Or Restart = True Then
  Unload Me
  Exit Sub
 End If
 CheckPrintJobs
 If Not NoProcessing Then
  CheckForPrinting
 End If
 If lsv.ListItems.Count = 0 And LenB(CommandSwitch("IF", True)) > 0 And ShellAndWaitingIsRunning = False Then
  Unload Me
  Exit Sub
 End If
 If lsv.ListItems.Count = 1 Then
  If lsv.SelectedItem.Index <> 1 Then
   lsv.ListItems(1).Selected = True
  End If
 End If
 DoEvents
 Timer1.Interval = TimerIntervall
 Timer1.Enabled = True
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
     DocumentDelete
    Case 8
     DocumentTop
    Case 9
     DocumentUp
    Case 10
     DocumentDown
    Case 11
     DocumentBottom
    Case 12
     DocumentCombine
    Case 13
     DocumentSave
    Case 15
     Call HTMLHelp_ShowTopic("html\welcome.htm")
   End Select
  Case 1
   Select Case Button.Index
    Case 1
     CombineAll
    Case 3
     CombineAllAndSend
    Case 4
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
 If lsv.ListItems.Count > 0 Then
  If mnPrinter(0).Checked = True Then
    If PrintSelectedJobs = True Then
      If lsv.ListItems(1).SubItems(1) = LanguageStrings.ListPrinting Then
        PDFSpoolfile = lsv.ListItems(1).SubItems(4)
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
    PDFSpoolfile = lsv.ListItems(1).SubItems(4)
    If PrinterStop = False Then
     If IsFormLoaded(frmPrinting) = False Then
      If InstalledAsServer Then ' This is necessary because there is no other way to tell the running server that the options are changed!
       Options = ReadOptions
      End If
      If Options.UseAutosave = 1 Then
        Autosave
       Else
        frmPrinting.Show , Me
      End If
     End If
    End If
    If PrinterStop = False And NoProcessing = False Then
      mnPrinter(0).Checked = False
      tlb(0).Buttons(1).Image = 1
     Else
      mnPrinter(0).Checked = True
      tlb(0).Buttons(1).Image = 2
    End If
  End If
 End If
End Sub

Public Sub CheckPrintJobs()
 Dim Temppath As String, LItem As ListItem, Files As Collection, tColl As Collection, _
  tFile() As String, i As Long, j As Long, kB As Long, MB As Long, GB As Long, _
  ind As Long, tStr As String, tB As Boolean
 kB = 1024: MB = kB * 1024: GB = MB * 1024
 Set Files = New Collection: Timer1.Enabled = False
 Temppath = CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory
 FindFiles Temppath, Files, "~PS*.tmp", , False, True

 If Files.Count = 0 And lsv.ListItems.Count > 0 Then
  lsv.ListItems.Clear
  SetDocMenuAndToolbar
  Exit Sub
 End If

 tB = False

 Set tColl = New Collection
 For j = 1 To lsv.ListItems.Count
  ind = 0
  For i = 1 To Files.Count
   tFile = Split(Files.Item(i), "|")
   If UCase$(tFile(1)) = UCase$(lsv.ListItems(j).SubItems(4)) Then
    ind = i
    Exit For
   End If
  Next i
  If ind = 0 Then
   tColl.Add j
  End If
 Next j
 If tColl.Count > 0 Then
  tB = True
 End If
 For j = 1 To tColl.Count
  lsv.ListItems.Remove tColl(j) - (j - 1)
 Next j


 Set tColl = New Collection
 For j = 1 To Files.Count
  tFile = Split(Files.Item(j), "|")
  ind = 0
  For i = 1 To lsv.ListItems.Count
   If UCase$(tFile(1)) = UCase$(lsv.ListItems(i).SubItems(4)) Then
    ind = i
    Exit For
   End If
  Next i
  If ind > 0 And ind < lsv.ListItems.Count + 1 Then
   tColl.Add j
  End If
 Next j
 For j = 1 To tColl.Count
  Files.Remove tColl(j) - (j - 1)
 Next j

 For j = 1 To Files.Count
  tFile = Split(Files.Item(j), "|")
  ind = 0
  For i = 1 To lsv.ListItems.Count
   If UCase$(tFile(1)) = UCase$(lsv.ListItems(i).SubItems(4)) Then
    ind = i
    Exit For
   End If
  Next i
  If ind = 0 Then
   tB = True
   If FormISLoaded("frmOptions") = False Then
    Me.Show
   End If
   Set LItem = lsv.ListItems.Add(, , GetPSTitle(tFile(1)))
   
   LItem.SubItems(1) = LanguageStrings.ListWaiting
   LItem.SubItems(2) = tFile(3)
   
   If CLng(tFile(2)) > GB Then
     LItem.SubItems(3) = Format$(CDbl(tFile(2)) / GB, "0.00 " & LanguageStrings.ListGBytes)
    Else
     If CLng(tFile(2)) > MB Then
       LItem.SubItems(3) = Format$(CDbl(tFile(2)) / MB, "0.00 " & LanguageStrings.ListMBytes)
      Else
       If CLng(tFile(2)) > kB Then
         LItem.SubItems(3) = Format$(CDbl(tFile(2)) / kB, "0.00 " & LanguageStrings.ListKBytes)
        Else
         LItem.SubItems(3) = Format$(tFile(2), "0 " & LanguageStrings.ListBytes)
       End If
    End If
   End If
   LItem.SubItems(4) = tFile(1)
   DoEvents
  End If
 Next j
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
  mnDocument(c).Enabled = True
 Next c
 For c = 1 To tlb(0).Buttons.Count
  tlb(0).Buttons(c).Enabled = True
 Next c
 If (Options.Toolbars And 2) = 2 Then
  For c = 1 To tlb(1).Buttons.Count
   tlb(1).Buttons(c).Enabled = True
  Next c
 End If
 Select Case True
  Case lsv.ListItems.Count = 0, LvwGetCountSelectedItems(lsv, False) = 0
   With mnDocument
    .Item(0).Enabled = False
    .Item(3).Enabled = False
    .Item(5).Enabled = False
    .Item(6).Enabled = False
    .Item(7).Enabled = False
    .Item(8).Enabled = False
    .Item(10).Enabled = False
    .Item(12).Enabled = False
    If (Options.Toolbars And 2) = 2 Then
     .Item(14).Enabled = False
     .Item(15).Enabled = False
     .Item(16).Enabled = False
    End If
   End With
   With tlb(0)
    .Buttons(5).Enabled = False
    For c = 7 To 13
     .Buttons(c).Enabled = False
    Next c
   End With
   With tlb(1)
    If (Options.Toolbars And 2) = 2 Then
     .Buttons(1).Enabled = False
     .Buttons(3).Enabled = False
     .Buttons(4).Enabled = False
    End If
   End With
   Exit Sub
  Case lsv.ListItems.Count = 1
   With mnDocument
    .Item(5).Enabled = False
    .Item(6).Enabled = False
    .Item(7).Enabled = False
    .Item(8).Enabled = False
    .Item(10).Enabled = False
    If (Options.Toolbars And 2) = 2 Then
     .Item(14).Enabled = False
     .Item(15).Enabled = False
     If LenB(txtEmailAddress.Text) = 0 Or Options.DisableEmail <> 0 Then
      .Item(16).Enabled = False
     End If
    End If
   End With
   With tlb(0)
    For c = 8 To 12
     .Buttons(c).Enabled = False
    Next c
   End With
   With tlb(1)
    If (Options.Toolbars And 2) = 2 Then
     .Buttons(1).Enabled = False
     .Buttons(3).Enabled = False
     If LenB(txtEmailAddress.Text) = 0 Then
      .Buttons(4).Enabled = False
     End If
    End If
   End With
  Case lsv.ListItems.Count > 1
   With mnDocument
    If AllSelectedListitemsAtTop Then
     .Item(5).Enabled = False
     .Item(6).Enabled = False
    End If
    If AllSelectedListitemsAtBottom Then
     .Item(7).Enabled = False
     .Item(8).Enabled = False
    End If
    If LvwGetCountSelectedItems(lsv, False) = 1 Then
     .Item(10).Enabled = False
    End If
    If LvwGetCountSelectedItems(lsv, False) > 1 Then
     .Item(12).Enabled = False
    End If
    If ((Options.Toolbars And 2) = 2 And LenB(txtEmailAddress.Text) = 0) Or Options.DisableEmail <> 0 Then
     .Item(15).Enabled = False
     .Item(16).Enabled = False
    End If
   End With
   With tlb(0)
    If AllSelectedListitemsAtTop Then
     .Buttons(8).Enabled = False
     .Buttons(9).Enabled = False
    End If
    If AllSelectedListitemsAtBottom Then
     .Buttons(10).Enabled = False
     .Buttons(11).Enabled = False
    End If
    If LvwGetCountSelectedItems(lsv, False) = 1 Then
     .Buttons(12).Enabled = False
    End If
    If LvwGetCountSelectedItems(lsv, False) > 1 Then
     .Buttons(13).Enabled = False
    End If
   End With
   With tlb(1)
    If (Options.Toolbars And 2) = 2 And LenB(txtEmailAddress.Text) = 0 Then
     .Buttons(3).Enabled = False
     .Buttons(4).Enabled = False
    End If
   End With
 End Select
End Sub

Private Sub Autosave(Optional Filename As String = vbNullString)
 Dim tColl As Collection, i As Long, tFile() As String, Pathname As String, _
  OutputFilename As String, PDFDocInfo As tPDFDocInfo, tStr As String, _
  PSHeader As tPSHeader, tDate As Date

 Set tColl = New Collection

 If Len(Filename) > 0 Then
   If FileExists(Filename) = True Then
    SplitPath Filename, , Pathname
    tColl.Add Pathname & "|" & Filename & "|" & FileLen(Filename) & "|" & FileDateTime(Filename)
   End If
  Else
   FindFiles CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory, tColl, "~PS*.tmp", , True, True
'   Set tColl = GetFiles(GetPDFCreatorTempfolder, "~P*.tmp")
 End If

 tStr = "Autosavemodus: " & tColl.Count & "files"
 IfLoggingWriteLogfile tStr
 WriteToSpecialLogfile tStr
 Do While tColl.Count > 0
  For i = 1 To tColl.Count
   tFile = Split(tColl.Item(i), "|")
   If FileExists(tFile(1)) And Not FileInUse(tFile(1)) Then
    OutputFilename = GetAutosaveFilename(tFile(1))
    SplitPath OutputFilename, , Pathname
    If IsValidPath(Pathname) = True Then
      If DirExists(Pathname) = False Then
       MakePath (Pathname)
      End If
      tStr = "Autosavemodus: Create File '" & OutputFilename & "'"
      IfLoggingWriteLogfile tStr
      WriteToSpecialLogfile tStr
      PSHeader = GetPSHeader(tFile(1))
      tDate = Now
      With PDFDocInfo
       If Options.UseStandardAuthor = 1 Then
         .Author = GetSubstFilename(tFile(1), RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardAuthor)), True)
        Else
         .Author = GetDocUsername(tFile(1), False)
       End If
       If LenB(PSHeader.CreationDate.Comment) > 0 Then
         tStr = PSHeader.CreationDate.Comment
        Else
         tStr = CStr(tDate)
       End If
       .CreationDate = GetDocDate(Trim$(Options.StandardCreationdate), Options.StandardDateformat, FormatPrintDocumentDate(tStr))
       .Creator = App.EXEName & " Version " & App.Major & "." & App.Minor & "." & App.Revision
       .Keywords = GetSubstFilename(tFile(1), RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardKeywords)))
       'tStr = CStr(tDate)
       .ModifyDate = GetDocDate(Trim$(Options.StandardModifydate), Options.StandardDateformat, FormatPrintDocumentDate(tStr))
       .Subject = GetSubstFilename(tFile(1), RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardSubject)))
       If Len(Options.StandardTitle) > 0 Then
         .Title = GetSubstFilename(tFile(1), RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardTitle)))
        Else
         .Title = GetSubstFilename(tFile(1), Options.SaveFilename)
       End If
      End With
      AppendPDFDocInfo tFile(1), PDFDocInfo
      CheckForStamping tFile(1)
      If Options.RunProgramBeforeSaving = 1 Then
       RunProgramBeforeSaving Me.hwnd, tFile(1), _
        Options.RunProgramBeforeSavingProgramParameters, _
        Options.RunProgramBeforeSavingWindowstyle
      End If
      CallGScript tFile(1), OutputFilename, Options, Options.AutosaveFormat
      If FileExists(OutputFilename) = True Then
        tStr = "Autosavemodus: Create File '" & OutputFilename & "' success"
        IfLoggingWriteLogfile tStr
        WriteToSpecialLogfile tStr
        If Options.RunProgramAfterSaving = 1 Then
         RunProgramAfterSaving Me.hwnd, OutputFilename, _
         Options.RunProgramAfterSavingProgramParameters, _
         Options.RunProgramAfterSavingWindowstyle, tFile(1)
        End If
       Else
        tStr = "Autosavemodus: Create File '" & OutputFilename & "' failed"
        IfLoggingWriteLogfile tStr
        WriteToSpecialLogfile tStr
      End If
     Else
      IfLoggingWriteLogfile "Error: Invalid autosave pathname, spoolfile will be deleted!"
    End If
    CheckForPrintingAfterSaving tFile(1), Options
    KillFile tFile(1)
    KillInfoSpoolfile tFile(1)
    ConvertedOutputFilename = OutputFilename
    ReadyConverting = True
   End If
  Next i
  Set tColl = GetFiles(GetPDFCreatorTempfolder, "~PS*.tmp", SortedByDate)
 Loop
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

Private Function IsCompatibleLanguageVersion(Version As String) As Boolean
 Dim i As Byte, delim As String, fVers() As String, fCVers() As String, _
  ProgVersion As String, fPVers() As String
 IsCompatibleLanguageVersion = False
 delim = "."
 ProgVersion = GetProgramRelease
 If Len(CompatibleLanguageVersion) = 0 Or Len(Version) = 0 Or Len(ProgVersion) = 0 Then
  Exit Function
 End If
 If InStr(1, CompatibleLanguageVersion, delim) = 0 Or _
    InStr(1, Version, delim) = 0 Or _
    InStr(1, ProgVersion, delim) = 0 Then
  Exit Function
 End If
 fVers = Split(Version, delim)
 fCVers = Split(CompatibleLanguageVersion, delim)
 fPVers = Split(ProgVersion, delim)
 If UBound(fVers) < 2 Or UBound(fCVers) < 2 Or UBound(fPVers) < 2 Then
  Exit Function
 End If
 For i = 0 To 2
  If IsNumeric(fVers(i)) = False Or IsNumeric(fCVers(i)) = False Or _
   IsNumeric(fPVers(i)) = False Then
   Exit Function
  End If
 Next i
 If CLng(fVers(0)) >= CLng(fCVers(0)) And CLng(fVers(0)) <= CLng(fPVers(0)) Then
  If CLng(fVers(1)) >= CLng(fCVers(1)) And CLng(fVers(1)) <= CLng(fPVers(1)) Then
   If CLng(fVers(2)) >= CLng(fCVers(2)) And CLng(fVers(2)) <= CLng(fPVers(2)) Then
    IsCompatibleLanguageVersion = True
   End If
  End If
 End If
End Function

Private Sub Timer2_Timer()
 On Error Resume Next
 If mutexLocal.CheckMutex(PDFCreator_GUID) = False Then
 ' Create a lokal mutex
   mutexLocal.CreateMutex PDFCreator_GUID
 End If
 If mutexGlobal.CheckMutex("Global\" & PDFCreator_GUID) = False Then
 ' Create a global mutex
   mutexGlobal.CreateMutex "Global\" & PDFCreator_GUID
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
  .Buttons.Add , , , , 7
  .Buttons.Add , , , , 8
  .Buttons.Add , , , , 9
  .Buttons.Add , , , , 10
  .Buttons.Add , , , , 11
  .Buttons.Add , , , , 12
  .Buttons.Add , , , , 13
  .Buttons.Add , , , tbrSeparator
  .Buttons.Add , , , , 14
 End With
 With tlb(1)
  Set .ImageList = imlTlb
  .Buttons.Clear
  .Buttons.Add , , , , 15
  .Buttons.Add , , , tbrSeparator
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
 If (Options.Toolbars And 2) <> 2 Then
  txtEmailAddress.Enabled = False
  txtEmailAddress.Visible = False
  mnDocument(14).Enabled = False
  mnDocument(15).Enabled = False
  mnDocument(16).Enabled = False
  mnDocument(13).Visible = False
  mnDocument(14).Visible = False
  mnDocument(15).Visible = False
  mnDocument(16).Visible = False
 End If
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
' mnPrinter(0).Checked = False
End Sub

Private Sub DocumentAdd()
 Dim i As Long, Cancel As Boolean, cFiles As Collection, aLen As Double, _
  OnlyPsFiles As Boolean, Ext As String, DefaultPrintername As String, _
  tFilename As String, tLen As Double
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
    tFilename = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory, "~PS")
    DoEvents
    FileCopy cFiles.Item(i), tFilename
   Else
    DoEvents
    ShellAndWait Me.hwnd, "print", cFiles.Item(i), "", vbNullChar, wHidden, WCTermination, 60000, True
    DoEvents
  End If
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
 For i = 1 To lsv.ListItems.Count
  If lsv.ListItems(i).Selected = True Then
   KillFile lsv.ListItems(i).SubItems(4)
   KillInfoSpoolfile lsv.ListItems(i).SubItems(4)
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

Public Sub DocumentCombine()
 Dim i As Long, cFiles As Collection, tFilename As String, tFilename2 As String
 Screen.MousePointer = vbHourglass
 LockWindowUpdate lsv.hwnd
 Set cFiles = New Collection
 For i = 1 To lsv.ListItems.Count
  If lsv.ListItems(i).Selected = True Then
   cFiles.Add lsv.ListItems(i).SubItems(4)
  End If
 Next i
 tFilename = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory, "~PS")
 KillFile tFilename
 If cFiles.Count > 1 Then
  CombineFiles tFilename, cFiles, , stb
  stb.Panels("Percent").Text = ""
 End If
 Set cFiles = Nothing
 LockWindowUpdate 0&
 Screen.MousePointer = vbNormal
End Sub

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
 If mnPrinter(0).Checked = False Or NoProcessing = True Then
   SetPrinterStop True
   mnPrinter(0).Checked = True
   tlb(0).Buttons(1).Image = 2
  Else
   SetPrinterStop False
   mnPrinter(0).Checked = False
   tlb(0).Buttons(1).Image = 1
 End If
 If Not m_frmSysTray Is Nothing Then
  If mnPrinter(0).Checked = True Then
    m_frmSysTray.IconHandle = imlSystray.ListImages(1).Picture.handle
    m_frmSysTray.mnuSysTray(2).Checked = True
   Else
    m_frmSysTray.IconHandle = imlSystray.ListImages(2).Picture.handle
    m_frmSysTray.mnuSysTray(2).Checked = False
  End If
 End If
End Sub

Private Sub SetLanguageToolbar()
 With tlb(0)
  .Buttons(1).ToolTipText = LanguageStrings.DialogPrinterPrinterStop
  .Buttons(2).ToolTipText = LanguageStrings.DialogPrinterOptions
  .Buttons(3).ToolTipText = LanguageStrings.DialogPrinterLogfile
  .Buttons(5).ToolTipText = LanguageStrings.DialogDocumentPrint
  .Buttons(6).ToolTipText = LanguageStrings.DialogDocumentAdd
  .Buttons(7).ToolTipText = LanguageStrings.DialogDocumentDelete
  .Buttons(8).ToolTipText = LanguageStrings.DialogDocumentTop
  .Buttons(9).ToolTipText = LanguageStrings.DialogDocumentUp
  .Buttons(10).ToolTipText = LanguageStrings.DialogDocumentDown
  .Buttons(11).ToolTipText = LanguageStrings.DialogDocumentBottom
  .Buttons(12).ToolTipText = LanguageStrings.DialogDocumentCombine
  .Buttons(13).ToolTipText = LanguageStrings.DialogDocumentSave
  .Buttons(15).ToolTipText = "?"
 End With
 With tlb(1)
  .Buttons(1).ToolTipText = LanguageStrings.DialogDocumentCombineAll
  .Buttons(3).ToolTipText = LanguageStrings.DialogDocumentCombineAllSend
  .Buttons(4).ToolTipText = LanguageStrings.DialogDocumentSend
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

Public Function CombineAll() As String
 Dim i As Long, cFiles As Collection, tFilename As String, tFilename2 As String
 Screen.MousePointer = vbHourglass
 LockWindowUpdate hwnd
 Timer1.Enabled = False
 Set cFiles = New Collection
 For i = 1 To lsv.ListItems.Count
  cFiles.Add lsv.ListItems(i).SubItems(4)
 Next i
 tFilename = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory, "~PT")
 KillFile tFilename
 If cFiles.Count > 1 Then
  CombineFiles tFilename, cFiles, , stb
  stb.Panels("Percent").Text = ""
 End If
 tFilename2 = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory, "~PS")
 KillFile tFilename2
 Name tFilename As tFilename2
 Set cFiles = Nothing
 CombineAll = tFilename2
 Timer1.Enabled = True
 LockWindowUpdate 0&
 Screen.MousePointer = vbNormal
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
   ConvertPostscriptFile InputFilename, OutputFilename
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
   KillFile InputFilename
  End If
 End If
 Timer1.Enabled = True
End Sub

Private Sub CombineAllAndSend()
 Timer1.Enabled = False
 If Options.ShowAnimation = 1 Then
  ShowAnimationWindow = True
  frmAnimation.Show
 End If
 SendEmailImmediately CombineAll
 If Options.ShowAnimation = 1 Then
  ShowAnimationWindow = False
 End If
 Timer1.Enabled = True
End Sub

Private Sub SendEmail()
 If lsv.ListItems.Count > 0 Then
  If lsv.SelectedItem.Index >= 0 Then
   If FileExists(lsv.SelectedItem.ListSubItems(4)) = True And FileInUse(lsv.SelectedItem.ListSubItems(4)) = False Then
    Timer1.Enabled = False
    If Options.ShowAnimation = 1 Then
     ShowAnimationWindow = True
     frmAnimation.Show
    End If
    SendEmailImmediately lsv.SelectedItem.ListSubItems(4)
    If Options.ShowAnimation = 1 Then
     ShowAnimationWindow = False
     Timer1.Enabled = True
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
   If mnPrinter(4).Checked = True Then
    .mnuSysTray(6).Checked = True
   End If
   If mnPrinter(0).Checked = True Then
    .mnuSysTray(2).Checked = True
   End If
   .ToolTip = Me.Caption
  End With
  If mnPrinter(0).Checked = True Then
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
