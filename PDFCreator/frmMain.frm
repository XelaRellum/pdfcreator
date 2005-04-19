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
      Interval        =   500
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
      Height          =   1260
      Left            =   105
      TabIndex        =   0
      Top             =   1260
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   2223
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
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If KeyCode = vbKeyF1 Then
50050   KeyCode = 0
50060   Call HTMLHelp_ShowTopic("html\welcome.htm")
50070  End If
50080 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50090 Exit Sub
ErrPtnr_OnError:
50111 Select Case ErrPtnr.OnError("frmMain", "Form_KeyDown")
      Case 0: Resume
50130 Case 1: Resume Next
50140 Case 2: Exit Sub
50150 Case 3: End
50160 End Select
50170 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Load()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim b(1) As Boolean
50050  Me.KeyPreview = True
50060
50070  ReadAllLanguages LanguagePath
50080  InitProgram
50090
50100  ShowPaypalMenuimage
50110  SetGSRevision
50120
50130  If NoProcessing = True Or Options.NoProcessingAtStartup = 1 Then
50140   SetMenuPrinterStop
50150  End If
50160  If PrinterStop = True Then
50170   mnPrinter(0).Checked = True
50180  End If
50190  If Options.OptionsEnabled = 0 Then
50200   mnPrinter(2).Enabled = False
50210   tlb(0).Buttons(2).Enabled = False
50220  End If
50230  If Options.OptionsVisible = 0 Then
50240   tlb(0).Buttons(2).Visible = False
50250   mnPrinter(2).Visible = False
50260   mnPrinter(3).Visible = False
50270  End If
50280
50290  CheckPrintJobs
50300  Call SetDocMenuAndToolbar
50310  If (Options.Toolbars And 1) = 1 Then
50320    tlb(0).Visible = True
50330   Else
50340    tlb(0).Visible = False
50350  End If
50360  If (Options.Toolbars And 2) = 2 Then
50370    tlb(1).Visible = True
50380    txtEmailAddress.Visible = True
50390   Else
50400    tlb(1).Visible = False
50410    txtEmailAddress.Visible = False
50420  End If
50430  If PDFCreatorPrinter = False Or NoProcessing = True Or _
  Options.NoProcessingAtStartup = 1 Or (PDFCreatorPrinter = True And lsv.ListItems.Count > 1) Then
50450   Visible = True
50460   SetTopMost frmMain, True, True
50470   SetTopMost frmMain, False, True
50480   SetActiveWindow frmMain.hwnd
50490  End If
50500
50510  ' Only for the first time set interval to 10 ms
50520  Timer1.Interval = 10
50530  Timer1.Enabled = True
50540  Timer2.Interval = 100
50550  Timer2.Enabled = True
50560 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50570 Exit Sub
ErrPtnr_OnError:
50591 Select Case ErrPtnr.OnError("frmMain", "Form_Load")
      Case 0: Resume
50610 Case 1: Resume Next
50620 Case 2: Exit Sub
50630 Case 3: End
50640 End Select
50650 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Resize()
 On Error Resume Next
 If Me.WindowState = vbMinimized Then
  FormInTaskbar Me, False, False, False
  SystrayEnter
  Exit Sub
 End If
 SysTrayLeave
 FormInTaskbar Me, True, False, True
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
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  TerminateProgram
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("frmMain", "Form_Unload")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub InitProgram()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Printing = False
50050
50060  Set Printjobs = New Collection
50070
50080  stb.Panels.Clear
50090  stb.Panels.Add , "Status", ""
50100  stb.Panels.Add , "GhostscriptRevision", ""
50110  stb.Panels.Add , "Percent", ""
50120  stb.Panels("Percent").Width = 1000
50130  stb.Panels("GhostscriptRevision").Width = 1800
50140
50150  With lsv
50160   .View = lvwReport
50170   .FullRowSelect = True
50180   .HideSelection = False
50190   .ColumnHeaders.Clear
50200   .ColumnHeaders.Add , "Documenttitle", "Documenttitle", 2000
50210   .ColumnHeaders.Add , "Status", "Status", 1000
50220   .ColumnHeaders.Add , "Date", "Created on", 1700
50230   .ColumnHeaders.Add , "Size", "Size", 1500, lvwColumnRight
50240   .ColumnHeaders.Add , "Filename", "Filename", lsv.Width - 3500
50250  End With
50260
50270  With Options
50280   SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50290  End With
50300
50310  SetLanguageMenu
50320  If Options.Logging = 1 Then
50330    mnPrinter(4).Checked = True
50340   Else
50350    mnPrinter(4).Checked = False
50360  End If
50370
50380  InitToolbar
50390
50400  txtEmailAddress.ToolTipText = LanguageStrings.DialogEmailAddress
50410
50420  Form_Resize
50430
50440  DoEvents
50450 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50460 Exit Sub
ErrPtnr_OnError:
50481 Select Case ErrPtnr.OnError("frmMain", "InitProgram")
      Case 0: Resume
50500 Case 1: Resume Next
50510 Case 2: Exit Sub
50520 Case 3: End
50530 End Select
50540 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub TerminateProgram()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim PDFSpoolerPath As String, Files As Collection, i As Long, _
  tStrf() As String
50060  Timer1.Enabled = False
50070  Timer2.Enabled = False
50080
50090  Set Printjobs = Nothing
50100
50110  If Not mutex Is Nothing Then
50120   mutex.CloseMutex
50130   Set mutex = Nothing
50140  End If
50150
50160  UnloadDLLComplete GsDllLoaded
50170
50180  FindFiles CompletePath(GetPDFCreatorTempfolder), Files, "*.pdf", , False, True
50190  For i = 1 To Files.Count
50200   tStrf = Split(Files(i), "|")
50210   KillFile tStrf(1)
50220  Next i
50230
50240  IfLoggingWriteLogfile "PDFCreator Program End"
50250  SysTrayLeave
50260  PDFSpoolerPath = CompletePath(GetSystemDirectory) & "PDFSpooler.exe"
50270  If Restart = True And FileExists(PDFSpoolerPath) = True Then
50280   ShellExecute 0, vbNullString, """" & PDFSpoolerPath & """", "-SL200 -STTRUE", App.Path, 1
50290  End If
50300 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50310 Exit Sub
ErrPtnr_OnError:
50331 Select Case ErrPtnr.OnError("frmMain", "TerminateProgram")
      Case 0: Resume
50350 Case 1: Resume Next
50360 Case 2: Exit Sub
50370 Case 3: End
50380 End Select
50390 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function GetAllLanguagesFiles(LanguagePath As String) As Collection
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim tColl1 As Collection, tColl2 As Collection, i As Long, tStrf() As String, ini As clsINI, _
  Languagename As String
50060  Set GetAllLanguagesFiles = New Collection
50070  Set tColl1 = GetFiles(LanguagePath, "*.ini", SortedByName)
50080  Set tColl2 = New Collection
50090  For i = 1 To tColl1.Count
50100   tStrf = Split(tColl1(i), "|")
50110   Set ini = New clsINI
50120   ini.Filename = tStrf(1)
50130   ini.Section = "Common"
50140   Languagename = ini.GetKeyFromSection("Languagename")
50150   If Len(Languagename) = 0 Then
50160    Languagename = "No name available."
50170   End If
50180   AddSortedStr tColl2, Languagename & "|" & tStrf(1)
50190  Next i
50200  For i = 1 To tColl2.Count
50210   tStrf() = Split(tColl2(i), "|")
50220   GetAllLanguagesFiles.Add tStrf(1)
50230  Next i
50240 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50250 Exit Function
ErrPtnr_OnError:
50271 Select Case ErrPtnr.OnError("frmMain", "GetAllLanguagesFiles")
      Case 0: Resume
50290 Case 1: Resume Next
50300 Case 2: Exit Function
50310 Case 3: End
50320 End Select
50330 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub ReadAllLanguages(LanguagePath As String)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim Languagename As String, ini As clsINI, LangFiles As Collection, _
  i As Long, Version As String, Filename As String
50060  mnLanguage(0).Caption = "No languages available."
50070
50080  Set LangFiles = GetAllLanguagesFiles(LanguagePath)
50090  Set ini = New clsINI
50100  For i = 1 To LangFiles.Count
50110   ini.Filename = LangFiles.Item(i)
50120   ini.Section = "Common"
50130   Languagename = ini.GetKeyFromSection("Languagename")
50140   Version = ini.GetKeyFromSection("Version")
50150   If Len(Languagename) = 0 Then
50160    Languagename = "No name available."
50170   End If
50180   Load mnLanguage(mnLanguage.Count)
50190   If IsCompatibleLanguageVersion(Version) = True Then
50200     mnLanguage(mnLanguage.Count - 1).Caption = Languagename
50210    Else
50220     mnLanguage(mnLanguage.Count - 1).Caption = Languagename & " [" & Version & "]"
50230   End If
50240   mnLanguage(mnLanguage.Count - 1).Tag = LangFiles.Item(i)
50250   SplitPath LangFiles.Item(i), , , , Filename
50260   If UCase$(Options.Language) = UCase$(Filename) Then
50270    mnLanguage(mnLanguage.Count - 1).Checked = True
50280   End If
50290   DoEvents
50300  Next i
50310
50320  If mnLanguage.Count > 1 Then
50330   mnLanguage(0).Caption = "No languages available."
50340   mnLanguage(0).Visible = False
50350  End If
50360  Set ini = Nothing
50370 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50380 Exit Sub
ErrPtnr_OnError:
50401 Select Case ErrPtnr.OnError("frmMain", "ReadAllLanguages")
      Case 0: Resume
50420 Case 1: Resume Next
50430 Case 2: Exit Sub
50440 Case 3: End
50450 End Select
50460 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetLanguageMenu()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim i As Long, Version As String, reg As clsRegistry, Filename As String
50050
50060  For i = mnLanguage.LBound To mnLanguage.UBound
50070   If UCase$(Languagefile) = UCase$(mnLanguage.Item(i).Tag) Then
50080     mnLanguage.Item(i).Checked = True
50090     SplitPath Languagefile, , , , Filename
50100    Else
50110     mnLanguage.Item(i).Checked = False
50120   End If
50130  Next i
50140
50151  Select Case UCase$(Filename)
        Case "GERMAN"
50170    HelpFile = App.Path & "\PDFCreator_german.chm"
50180    If Not FileExists(HelpFile) Then
50190     HelpFile = App.Path & "\PDFCreator_english.chm"
50200    End If
50210   Case Else
50220    HelpFile = App.Path & "\PDFCreator_english.chm"
50230  End Select
50240
50250  If Not FileExists(HelpFile) Then
50260   MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & vbCrLf & HelpFile, vbExclamation
50270   HelpFile = App.Path & "\PDFCreator_english.chm"
50280  End If
50290
50300  With LanguageStrings
50310   Set reg = New clsRegistry
50320   With reg
50330    .hkey = HKEY_LOCAL_MACHINE
50340    .KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50350    Version = .GetRegistryValue("ApplicationVersion")
50360   End With
50370   Set reg = Nothing
50380   Caption = App.Title & " " & GetProgramReleaseStr & " " & .CommonTitle
50390
50400   mnPrinterMain.Caption = .DialogPrinter
50410   mnPrinter(0).Caption = .DialogPrinterPrinterStop
50420   mnPrinter(2).Caption = .DialogPrinterOptions
50430   mnPrinter(4).Caption = .DialogPrinterLogging
50440   mnPrinter(5).Caption = .DialogPrinterLogfile
50450   mnPrinter(7).Caption = .DialogPrinterClose
50460
50470   mnDocumentMain.Caption = .DialogDocument
50480   mnDocument(0).Caption = .DialogDocumentPrint
50490   mnDocument(2).Caption = .DialogDocumentAdd
50500   mnDocument(3).Caption = .DialogDocumentDelete
50510   mnDocument(5).Caption = .DialogDocumentTop
50520   mnDocument(6).Caption = .DialogDocumentUp
50530   mnDocument(7).Caption = .DialogDocumentDown
50540   mnDocument(8).Caption = .DialogDocumentBottom
50550   mnDocument(10).Caption = .DialogDocumentCombine
50560   mnDocument(12).Caption = .DialogDocumentSave
50570
50580   mnDocument(14).Caption = .DialogDocumentCombineAll
50590   mnDocument(15).Caption = .DialogDocumentCombineAllSend
50600   mnDocument(16).Caption = .DialogDocumentSend
50610
50620   mnViewMain.Caption = .DialogView
50630   mnView(0).Caption = .DialogViewStatusbar
50640
50650   mnLanguageMain.Caption = .DialogLanguage
50660
50670   mnHelpMain.Caption = .DialogInfo
50680   mnHelp(2).Caption = .DialogInfoPaypal
50690   mnHelp(4).Caption = .DialogInfoHomepage
50700   mnHelp(5).Caption = .DialogInfoPDFCreatorSourceforge
50710   mnHelp(6).Caption = .DialogInfoCheckUpdates
50720   mnHelp(8).Caption = .DialogInfoInfo
50730
50740   lsv.ColumnHeaders("Date").Text = .ListDate
50750   lsv.ColumnHeaders("Documenttitle").Text = .ListDocumenttitle
50760   lsv.ColumnHeaders("Filename").Text = .ListFilename
50770   lsv.ColumnHeaders("Size").Text = .ListSize
50780   lsv.ColumnHeaders("Status").Text = .ListStatus
50790  End With
50800 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50810 Exit Sub
ErrPtnr_OnError:
50831 Select Case ErrPtnr.OnError("frmMain", "SetLanguageMenu")
      Case 0: Resume
50850 Case 1: Resume Next
50860 Case 2: Exit Sub
50870 Case 3: End
50880 End Select
50890 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub lsv_DblClick()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  DocumentPrint
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("frmMain", "lsv_DblClick")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub lsv_KeyUp(KeyCode As Integer, Shift As Integer)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  SetDocMenuAndToolbar
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("frmMain", "lsv_KeyUp")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub lsv_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
 On Error Resume Next
 SetDocMenuAndToolbar
 If Button = 2 Then
  PopupMenu mnDocumentMain, , x, Y
 End If
End Sub

Private Sub lsv_OLEDragDrop(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
 Dim tFilename As String, i As Long, aLen As Double, tLen As Double, aw As Long, _
  DefaultPrintername As String, OnlyPsFiles As Boolean, Ext As String
 On Error Resume Next
 If data.GetFormat(vbCFFiles) Then
  DefaultPrintername = ""
  If data.Files.Count = 1 Then
    If IsPostscriptFile(data.Files.Item(1)) = True Then
      tFilename = GetTempFile(GetPDFCreatorTempfolder & PDFCreatorSpoolDirectory, "~PS")
      FileCopy data.Files.Item(1), tFilename
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
      ShellAndWait Me.hwnd, "print", data.Files.Item(1), "", vbNullChar, wHidden, WCTermination, 60000, True
      If DefaultPrintername <> vbNullString Then
       SetDefaultprinterInProg DefaultPrintername
      End If
    End If
    DoEvents
   Else
    OnlyPsFiles = True
    For i = 1 To data.Files.Count
     SplitPath data.Files.Item(i), , , , , Ext
     If UCase$(Ext) <> "PS" Then
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
    For i = 1 To data.Files.Count
     aLen = aLen + FileLen(data.Files.Item(i))
    Next i
    For i = 1 To data.Files.Count
     If IsPostscriptFile(data.Files.Item(i)) = True Then
       tFilename = GetTempFile(GetPDFCreatorTempfolder, "~PD")
       DoEvents
       FileCopy data.Files.Item(i), tFilename
      Else
       DoEvents
'       PrintDocument data.Files.item(i)
       ShellAndWait Me.hwnd, "print", data.Files.Item(i), "", vbNullChar, wHidden, WCTermination, 60000, True
       DoEvents
     End If
     tLen = tLen + FileLen(data.Files.Item(i))
     stb.Panels("Percent").Text = Format$(tLen / aLen, " 0.0%")
     DoEvents
    Next i
    If DefaultPrintername <> vbNullString Then
     SetDefaultprinterInProg DefaultPrintername
    End If
  End If
 End If
End Sub

Private Sub lsv_OLEDragOver(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single, State As Integer)
 Dim i As Long
 On Error Resume Next
 If data.GetFormat(vbCFFiles) Then
  Effect = ccOLEDropEffectCopy
  For i = 1 To data.Files.Count
   If IsFilePrintable(data.Files.Item(i)) = False And IsPostscriptFile(data.Files.Item(i)) = False Then
    Effect = ccOLEDropEffectNone
    Exit Sub
   End If
  Next i
 End If
End Sub

Private Sub mnDocument_Click(Index As Integer)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim tFilename As String, cFiles As Collection, sFiles() As String, _
  i As Long, j As Long, aLen As Double, tLen As Double, aw As Long, _
  DefaultPrintername As String, OnlyPsFiles As Boolean, Ext As String, _
  Cancel As Boolean, tFilename2 As String
50080
50090  Timer1.Enabled = False
50100  Screen.MousePointer = vbHourglass
50110  DoEvents
50121  Select Case Index
        Case 0:
50140    DocumentPrint
50150   Case 2: ' Add
50160    DocumentAdd
50170   Case 3: ' Delete
50180    DocumentDelete
50190   Case 5: ' Top
50200    DocumentTop
50210   Case 6: ' Up
50220    DocumentUp
50230   Case 7: ' Down
50240    DocumentDown
50250   Case 8: ' Bottom
50260    DocumentBottom
50270   Case 10: ' Combine
50280    DocumentCombine
50290   Case 12: ' Save
50300    DocumentSave
50310   Case 14
50320    CombineAll
50330   Case 15
50340    CombineAllAndSend
50350   Case 16
50360    SendEmail
50370  End Select
50380  SetDocMenuAndToolbar
50390  Screen.MousePointer = vbNormal
50400  Timer1.Enabled = True
50410 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50420 Exit Sub
ErrPtnr_OnError:
50441 Select Case ErrPtnr.OnError("frmMain", "mnDocument_Click")
      Case 0: Resume
50460 Case 1: Resume Next
50470 Case 2: Exit Sub
50480 Case 3: End
50490 End Select
50500 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub mnDocumentMain_Click()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  SetDocMenuAndToolbar
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("frmMain", "mnDocumentMain_Click")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub mnHelp_Click(Index As Integer)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim updStr As String, updStrA() As String, aw As Long
50050  Set dl = New clsDownload
50061  Select Case Index
        Case 0:
50080    Call HTMLHelp_ShowTopic("html\welcome.htm")
50090   Case 2:
50100    OpenDocument Paypal
50110   Case 4:
50120    OpenDocument Homepage
50130   Case 5:
50140    OpenDocument Sourceforge
50150   Case 6:
50160    updStr = dl.DownloadString(UpdateURL)
50170 '   updStr = dl.DownloadString("http://localhost:8080/update.txt")
50180    If Len(updStr) > 0 Then
50190      If CheckPDFCreatorVersion(updStr) > 0 Then
50200        updStrA = Split(updStr, ".")
50210        If updStrA(3) = 0 Then
50220          updStr = "[" & updStrA(0) & "." & updStrA(1) & "." & updStrA(2) & "]"
50230         Else
50240          updStr = "[" & updStrA(0) & "." & updStrA(1) & "." & updStrA(2) & " Beta " & updStrA(3) & "]"
50250        End If
50260        aw = MsgBox(Replace$(LanguageStrings.MessagesMsg32, "%1", updStr), vbYesNo + vbQuestion)
50270        If aw = vbYes Then
50280         OpenDocument "http://www.sourceforge.net/project/showfiles.php?group_id=57796"
50290        End If
50300       Else
50310        MsgBox LanguageStrings.MessagesMsg33, vbOKOnly + vbInformation
50320      End If
50330     Else
50340      MsgBox LanguageStrings.MessagesMsg31 & ": " & dl.ErrorDescription & " [" & dl.ErrorNumber & "]", vbOKOnly + vbExclamation
50350    End If
50360   Case 8:
50370    frmInfo.Show vbModal, Me
50380  End Select
50390 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50400 Exit Sub
ErrPtnr_OnError:
50421 Select Case ErrPtnr.OnError("frmMain", "mnHelp_Click")
      Case 0: Resume
50440 Case 1: Resume Next
50450 Case 2: Exit Sub
50460 Case 3: End
50470 End Select
50480 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub mnPrinter_Click(Index As Integer)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50041  Select Case Index
        Case 0:
50060    SetMenuPrinterStop
50070   Case 2:
50080    frmOptions.Show , Me
50090   Case 4:
50100    If mnPrinter(Index).Checked = False Then
50110      SetLogging True
50120      mnPrinter(Index).Checked = True
50130     Else
50140      SetLogging False
50150      mnPrinter(Index).Checked = False
50160    End If
50170    If Not m_frmSysTray Is Nothing Then
50180     If mnPrinter(Index).Checked = True Then
50190       m_frmSysTray.mnuSysTray(6).Checked = True
50200      Else
50210       m_frmSysTray.mnuSysTray(6).Checked = False
50220     End If
50230    End If
50240   Case 5:
50250    frmLog.Show , Me
50260   Case 7:
50270    Unload Me
50280  End Select
50290 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50300 Exit Sub
ErrPtnr_OnError:
50321 Select Case ErrPtnr.OnError("frmMain", "mnPrinter_Click")
      Case 0: Resume
50340 Case 1: Resume Next
50350 Case 2: Exit Sub
50360 Case 3: End
50370 End Select
50380 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub mnLanguage_Click(Index As Integer)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim File As String
50050  Screen.MousePointer = vbHourglass
50060  LoadLanguage mnLanguage(Index).Tag
50070  Languagefile = mnLanguage(Index).Tag
50080  SetLanguageMenu
50090  SplitPath Languagefile, , , , File
50100  SetLanguage File
50110  ShowPaypalMenuimage
50120  Me.Refresh
50130  Screen.MousePointer = vbNormal
50140 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50150 Exit Sub
ErrPtnr_OnError:
50171 Select Case ErrPtnr.OnError("frmMain", "mnLanguage_Click")
      Case 0: Resume
50190 Case 1: Resume Next
50200 Case 2: Exit Sub
50210 Case 3: End
50220 End Select
50230 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub mnView_Click(Index As Integer)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50041  Select Case Index
        Case 0:
50060    stb.Visible = Not stb.Visible
50070    mnView(0).Checked = Not mnView(0).Checked
50080    Form_Resize
50090  End Select
50100 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50110 Exit Sub
ErrPtnr_OnError:
50131 Select Case ErrPtnr.OnError("frmMain", "mnView_Click")
      Case 0: Resume
50150 Case 1: Resume Next
50160 Case 2: Exit Sub
50170 Case 3: End
50180 End Select
50190 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Timer1_Timer()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Timer1.Enabled = False
50050  DoEvents
50060  If FileExists(CompletePath(App.Path) & "Unload.tmp") = True Or Restart = True Then
50070   Unload Me
50080   Exit Sub
50090  End If
50100  CheckPrintJobs
50110  If Not NoProcessing Then
50120   CheckForPrinting
50130  End If
50140  If lsv.ListItems.Count = 0 And LenB(CommandSwitch("IF", True)) > 0 And ShellAndWaitingIsRunning = False Then
50150   Unload Me
50160   Exit Sub
50170  End If
50180  If lsv.ListItems.Count = 1 Then
50190   If lsv.SelectedItem.Index <> 1 Then
50200    lsv.ListItems(1).Selected = True
50210   End If
50220  End If
50230  DoEvents
50240  Timer1.Interval = TimerIntervall
50250  Timer1.Enabled = True
50260 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50270 Exit Sub
ErrPtnr_OnError:
50291 Select Case ErrPtnr.OnError("frmMain", "Timer1_Timer")
      Case 0: Resume
50310 Case 1: Resume Next
50320 Case 2: Exit Sub
50330 Case 3: End
50340 End Select
50350 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub tlb_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50041  Select Case Index
        Case 0
50061    Select Case Button.Index
          Case 1
50080      SetMenuPrinterStop
50090     Case 2
50100      frmOptions.Show , Me
50110     Case 3
50120      frmLog.Show , Me
50130     Case 5
50140      DocumentPrint
50150     Case 6
50160      DocumentAdd
50170     Case 7
50180      DocumentDelete
50190     Case 8
50200      DocumentTop
50210     Case 9
50220      DocumentUp
50230     Case 10
50240      DocumentDown
50250     Case 11
50260      DocumentBottom
50270     Case 12
50280      DocumentCombine
50290     Case 13
50300      DocumentSave
50310     Case 15
50320      Call HTMLHelp_ShowTopic("html\welcome.htm")
50330    End Select
50340   Case 1
50351    Select Case Button.Index
          Case 1
50370      CombineAll
50380     Case 3
50390      CombineAllAndSend
50400     Case 4
50410      SendEmail
50420    End Select
50430  End Select
50440  SetDocMenuAndToolbar
50450 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50460 Exit Sub
ErrPtnr_OnError:
50481 Select Case ErrPtnr.OnError("frmMain", "tlb_ButtonClick")
      Case 0: Resume
50500 Case 1: Resume Next
50510 Case 2: Exit Sub
50520 Case 3: End
50530 End Select
50540 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub txtEmailAddress_Change()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  txtEmailAddress.ToolTipText = LanguageStrings.DialogEmailAddress & ": " & txtEmailAddress.Text
50050  If InStr(1, txtEmailAddress.Text, "@") = 0 And LenB(Options.StandardMailDomain) > 0 Then
50060   txtEmailAddress.ToolTipText = txtEmailAddress.ToolTipText & "@" & Options.StandardMailDomain
50070  End If
50080  SetDocMenuAndToolbar
50090 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50100 Exit Sub
ErrPtnr_OnError:
50121 Select Case ErrPtnr.OnError("frmMain", "txtEmailAddress_Change")
      Case 0: Resume
50140 Case 1: Resume Next
50150 Case 2: Exit Sub
50160 Case 3: End
50170 End Select
50180 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub CheckForPrinting()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If lsv.ListItems.Count > 0 Then
50050   If mnPrinter(0).Checked = True Then
50060     If PrintSelectedJobs = True Then
50070       If lsv.ListItems(1).SubItems(1) = LanguageStrings.ListPrinting Then
50080         PDFSpoolfile = lsv.ListItems(1).SubItems(4)
50090         frmPrinting.Show vbModal, Me
50100        Else
50110         PrintSelectedJobs = False
50120       End If
50130      Else
50140       If lsv.ListItems(1).SubItems(1) <> LanguageStrings.ListWaiting Then
50150        lsv.ListItems(1).SubItems(1) = LanguageStrings.ListWaiting
50160       End If
50170     End If
50180    Else
50190     If lsv.ListItems(1).SubItems(1) <> LanguageStrings.ListPrinting Then
50200      lsv.ListItems(1).SubItems(1) = LanguageStrings.ListPrinting
50210     End If
50220     PDFSpoolfile = lsv.ListItems(1).SubItems(4)
50230     If PrinterStop = False Then
50240      If IsFormLoaded(frmPrinting) = False Then
50250       If InstalledAsServer Then
50260        Options = ReadOptions
50270       End If
50280       If Options.UseAutosave = 1 Then
50290         Autosave
50300        Else
50310         frmPrinting.Show , Me
50320       End If
50330      End If
50340     End If
50350     If PrinterStop = False And NoProcessing = False Then
50360       mnPrinter(0).Checked = False
50370       tlb(0).Buttons(1).Image = 1
50380      Else
50390       mnPrinter(0).Checked = True
50400       tlb(0).Buttons(1).Image = 2
50410     End If
50420   End If
50430  End If
50440 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50450 Exit Sub
ErrPtnr_OnError:
50471 Select Case ErrPtnr.OnError("frmMain", "CheckForPrinting")
      Case 0: Resume
50490 Case 1: Resume Next
50500 Case 2: Exit Sub
50510 Case 3: End
50520 End Select
50530 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub CheckPrintJobs()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim Temppath As String, LItem As ListItem, Files As Collection, tColl As Collection, _
  tFile() As String, i As Long, j As Long, kB As Long, MB As Long, GB As Long, _
  ind As Long, tstr As String, tB As Boolean
50070  kB = 1024: MB = kB * 1024: GB = MB * 1024
50080  Set Files = New Collection
50090  Temppath = CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory
50100  FindFiles Temppath, Files, "~PS*.tmp", , False, True
50110
50120  If Files.Count = 0 And lsv.ListItems.Count > 0 Then
50130   lsv.ListItems.Clear
50140   SetDocMenuAndToolbar
50150   Exit Sub
50160  End If
50170
50180  tB = False
50190
50200  Set tColl = New Collection
50210  For j = 1 To lsv.ListItems.Count
50220   ind = 0
50230   For i = 1 To Files.Count
50240    tFile = Split(Files.Item(i), "|")
50250    If UCase$(tFile(1)) = UCase$(lsv.ListItems(j).SubItems(4)) Then
50260     ind = i
50270     Exit For
50280    End If
50290   Next i
50300   If ind = 0 Then
50310    tColl.Add j
50320   End If
50330  Next j
50340  If tColl.Count > 0 Then
50350   tB = True
50360  End If
50370  For j = 1 To tColl.Count
50380   lsv.ListItems.Remove tColl(j) - (j - 1)
50390  Next j
50400
50410
50420  Set tColl = New Collection
50430  For j = 1 To Files.Count
50440   tFile = Split(Files.Item(j), "|")
50450   ind = 0
50460   For i = 1 To lsv.ListItems.Count
50470    If UCase$(tFile(1)) = UCase$(lsv.ListItems(i).SubItems(4)) Then
50480     ind = i
50490     Exit For
50500    End If
50510   Next i
50520   If ind > 0 And ind < lsv.ListItems.Count + 1 Then
50530    tColl.Add j
50540   End If
50550  Next j
50560  For j = 1 To tColl.Count
50570   Files.Remove tColl(j) - (j - 1)
50580  Next j
50590
50600  For j = 1 To Files.Count
50610   tFile = Split(Files.Item(j), "|")
50620   ind = 0
50630   For i = 1 To lsv.ListItems.Count
50640    If UCase$(tFile(1)) = UCase$(lsv.ListItems(i).SubItems(4)) Then
50650     ind = i
50660     Exit For
50670    End If
50680   Next i
50690   If ind = 0 Then
50700    tB = True
50710    If frmOptions Is Nothing Then
50720     SetActiveWindow Me.hwnd
50730    End If
50740    Set LItem = lsv.ListItems.Add(, , GetPSTitle(tFile(1)))
50750    LItem.SubItems(1) = LanguageStrings.ListWaiting
50760    LItem.SubItems(2) = tFile(3)
50770    If CLng(tFile(2)) > GB Then
50780      LItem.SubItems(3) = Format$(CDbl(tFile(2)) / GB, "0.00 " & LanguageStrings.ListGBytes)
50790     Else
50800      If CLng(tFile(2)) > MB Then
50810        LItem.SubItems(3) = Format$(CDbl(tFile(2)) / MB, "0.00 " & LanguageStrings.ListMBytes)
50820       Else
50830        If CLng(tFile(2)) > kB Then
50840          LItem.SubItems(3) = Format$(CDbl(tFile(2)) / kB, "0.00 " & LanguageStrings.ListKBytes)
50850         Else
50860          LItem.SubItems(3) = Format$(tFile(2), "0 " & LanguageStrings.ListBytes)
50870        End If
50880     End If
50890    End If
50900    LItem.SubItems(4) = tFile(1)
50910    DoEvents
50920   End If
50930  Next j
50940  If lsv.ListItems.Count = 1 Then
50950    tstr = LanguageStrings.ListStatus & ": " & lsv.ListItems.Count & " " & LanguageStrings.MessagesMsg01
50960   Else
50970    tstr = LanguageStrings.ListStatus & ": " & lsv.ListItems.Count & " " & LanguageStrings.MessagesMsg02
50980  End If
50990  If tstr <> stb.Panels("Status").Text Then
51000   stb.Panels("Status").Text = tstr
51010  End If
51020  If tB = True Then
51030   SetDocMenuAndToolbar
51040  End If
51050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
51060 Exit Sub
ErrPtnr_OnError:
51081 Select Case ErrPtnr.OnError("frmMain", "CheckPrintJobs")
      Case 0: Resume
51100 Case 1: Resume Next
51110 Case 2: Exit Sub
51120 Case 3: End
51130 End Select
51140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub CheckPrintJobsOld()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim Temppath As String, LItem As ListItem, tColl As Collection, _
  tFile() As String, i As Long, j As Long, kB As Long, MB As Long, GB As Long
50060  kB = 1024: MB = kB * 1024: GB = MB * 1024
50070  Set tColl = New Collection
50080  Temppath = CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory
50090
50100  'Set tColl = GetFiles(GetPDFCreatorTempfolder, "~PD*.tmp")
50110  FindFiles Temppath, tColl, "~PS*.tmp", , False, True
50120  If tColl.Count = 0 And lsv.ListItems.Count > 0 Then
50130   lsv.ListItems.Clear
50140  End If
50150  For i = 1 To tColl.Count
50160   tFile = Split(tColl.Item(i), "|")
50170   For j = 1 To lsv.ListItems.Count
50180    If UCase$(tFile(1)) = UCase$(lsv.ListItems(j).SubItems(4)) Then
50190     Exit For
50200    End If
50210   Next j
50220   If j > lsv.ListItems.Count Then
50230     SetActiveWindow Me.hwnd
50240     Set LItem = lsv.ListItems.Add(, , GetPSTitle(tFile(1)))
50250     LItem.SubItems(1) = LanguageStrings.ListWaiting
50260     LItem.SubItems(2) = tFile(3)
50270     If CLng(tFile(2)) > GB Then
50280       LItem.SubItems(3) = Format$(CDbl(tFile(2)) / GB, "0.00 " & LanguageStrings.ListGBytes)
50290      Else
50300       If CLng(tFile(2)) > MB Then
50310         LItem.SubItems(3) = Format$(CDbl(tFile(2)) / MB, "0.00 " & LanguageStrings.ListMBytes)
50320        Else
50330         If CLng(tFile(2)) > kB Then
50340           LItem.SubItems(3) = Format$(CDbl(tFile(2)) / kB, "0.00 " & LanguageStrings.ListKBytes)
50350          Else
50360           LItem.SubItems(3) = Format$(tFile(2), "0 " & LanguageStrings.ListBytes)
50370         End If
50380      End If
50390     End If
50400     LItem.SubItems(4) = tFile(1)
50410     DoEvents
50420    Else
50430 '
50440   End If
50450  Next i
50460  i = 0
50470  Do Until i + 1 >= lsv.ListItems.Count
50480   i = i + 1
50490   For j = 1 To tColl.Count
50500    tFile = Split(tColl.Item(j), "|")
50510    If UCase$(tFile(1)) = UCase$(lsv.ListItems(i).SubItems(4)) Then
50520     Exit For
50530    End If
50540   Next j
50550   If j > tColl.Count Then
50560    lsv.ListItems.Remove i
50570   End If
50580   DoEvents
50590  Loop
50600  If lsv.ListItems.Count = 1 Then
50610    stb.Panels("Status").Text = LanguageStrings.ListStatus & ": " & lsv.ListItems.Count & " " & LanguageStrings.MessagesMsg01
50620   Else
50630    stb.Panels("Status").Text = LanguageStrings.ListStatus & ": " & lsv.ListItems.Count & " " & LanguageStrings.MessagesMsg02
50640  End If
50650  Set tColl = Nothing
50660  SetDocMenuAndToolbar
50670 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50680 Exit Sub
ErrPtnr_OnError:
50701 Select Case ErrPtnr.OnError("frmMain", "CheckPrintJobsOld")
      Case 0: Resume
50720 Case 1: Resume Next
50730 Case 2: Exit Sub
50740 Case 3: End
50750 End Select
50760 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetDocMenuAndToolbar()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim c As Long
50050  For c = 0 To mnDocument.Count - 1
50060   mnDocument(c).Enabled = True
50070  Next c
50080  For c = 1 To tlb(0).Buttons.Count
50090   tlb(0).Buttons(c).Enabled = True
50100  Next c
50110  If (Options.Toolbars And 2) = 2 Then
50120   For c = 1 To tlb(1).Buttons.Count
50130    tlb(1).Buttons(c).Enabled = True
50140   Next c
50150  End If
50161  Select Case True
        Case lsv.ListItems.Count = 0, LvwGetCountSelectedItems(lsv, True) = 0
50180    With mnDocument
50190     .Item(0).Enabled = False
50200     .Item(3).Enabled = False
50210     .Item(5).Enabled = False
50220     .Item(6).Enabled = False
50230     .Item(7).Enabled = False
50240     .Item(8).Enabled = False
50250     .Item(10).Enabled = False
50260     .Item(12).Enabled = False
50270     If (Options.Toolbars And 2) = 2 Then
50280      .Item(14).Enabled = False
50290      .Item(15).Enabled = False
50300      .Item(16).Enabled = False
50310     End If
50320    End With
50330    With tlb(0)
50340     .Buttons(5).Enabled = False
50350     For c = 7 To 13
50360      .Buttons(c).Enabled = False
50370     Next c
50380    End With
50390    With tlb(1)
50400     If (Options.Toolbars And 2) = 2 Then
50410      .Buttons(1).Enabled = False
50420      .Buttons(3).Enabled = False
50430      .Buttons(4).Enabled = False
50440     End If
50450    End With
50460    Exit Sub
50470   Case lsv.ListItems.Count = 1
50480    With mnDocument
50490     .Item(5).Enabled = False
50500     .Item(6).Enabled = False
50510     .Item(7).Enabled = False
50520     .Item(8).Enabled = False
50530     .Item(10).Enabled = False
50540     If (Options.Toolbars And 2) = 2 Then
50550      .Item(14).Enabled = False
50560      .Item(15).Enabled = False
50570      If LenB(txtEmailAddress.Text) = 0 Then
50580       .Item(16).Enabled = False
50590      End If
50600     End If
50610    End With
50620    With tlb(0)
50630     For c = 8 To 12
50640      .Buttons(c).Enabled = False
50650     Next c
50660    End With
50670    With tlb(1)
50680     If (Options.Toolbars And 2) = 2 Then
50690      .Buttons(1).Enabled = False
50700      .Buttons(3).Enabled = False
50710      If LenB(txtEmailAddress.Text) = 0 Then
50720       .Buttons(4).Enabled = False
50730      End If
50740     End If
50750    End With
50760   Case lsv.ListItems.Count > 1
50770    With mnDocument
50780     If AllSelectedListitemsAtTop Then
50790      .Item(5).Enabled = False
50800      .Item(6).Enabled = False
50810     End If
50820     If AllSelectedListitemsAtBottom Then
50830      .Item(7).Enabled = False
50840      .Item(8).Enabled = False
50850     End If
50860     If LvwGetCountSelectedItems(lsv, True) = 1 Then
50870      .Item(10).Enabled = False
50880     End If
50890     If LvwGetCountSelectedItems(lsv, True) > 1 Then
50900      .Item(12).Enabled = False
50910     End If
50920     If (Options.Toolbars And 2) = 2 And LenB(txtEmailAddress.Text) = 0 Then
50930      .Item(15).Enabled = False
50940      .Item(16).Enabled = False
50950     End If
50960    End With
50970    With tlb(0)
50980     If AllSelectedListitemsAtTop Then
50990      .Buttons(8).Enabled = False
51000      .Buttons(9).Enabled = False
51010     End If
51020     If AllSelectedListitemsAtBottom Then
51030      .Buttons(10).Enabled = False
51040      .Buttons(11).Enabled = False
51050     End If
51060     If LvwGetCountSelectedItems(lsv, True) = 1 Then
51070      .Buttons(12).Enabled = False
51080     End If
51090     If LvwGetCountSelectedItems(lsv, True) > 1 Then
51100      .Buttons(13).Enabled = False
51110     End If
51120    End With
51130    With tlb(1)
51140     If (Options.Toolbars And 2) = 2 And LenB(txtEmailAddress.Text) = 0 Then
51150      .Buttons(3).Enabled = False
51160      .Buttons(4).Enabled = False
51170     End If
51180    End With
51190  End Select
51200 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
51210 Exit Sub
ErrPtnr_OnError:
51231 Select Case ErrPtnr.OnError("frmMain", "SetDocMenuAndToolbar")
      Case 0: Resume
51250 Case 1: Resume Next
51260 Case 2: Exit Sub
51270 Case 3: End
51280 End Select
51290 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Autosave(Optional Filename As String = vbNullString)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim tColl As Collection, i As Long, tFile() As String, Pathname As String, _
  OutputFilename As String, PDFDocInfo As tPDFDocInfo, tstr As String, _
  PSHeader As tPSHeader, tDate As Date
50070
50080  Set tColl = New Collection
50090
50100  If Len(Filename) > 0 Then
50110    If FileExists(Filename) = True Then
50120     SplitPath Filename, , Pathname
50130     tColl.Add Pathname & "|" & Filename & "|" & FileLen(Filename) & "|" & FileDateTime(Filename)
50140    End If
50150   Else
50160    FindFiles CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory, tColl, "~PS*.tmp", , True
50170 '   Set tColl = GetFiles(GetPDFCreatorTempfolder, "~P*.tmp")
50180  End If
50190
50200  tstr = "Autosavemodus: " & tColl.Count & "files"
50210  IfLoggingWriteLogfile tstr
50220  WriteToSpecialLogfile tstr
50230  Do While tColl.Count > 0
50240   For i = 1 To tColl.Count
50250    tFile = Split(tColl.Item(i), "|")
50260    OutputFilename = GetAutosaveFilename(tFile(1))
50270    SplitPath OutputFilename, , Pathname
50280    If IsValidPath(Pathname) = True Then
50290      If DirExists(Pathname) = False Then
50300       MakePath (Pathname)
50310      End If
50320      tstr = "Autosavemodus: Create File '" & OutputFilename & "'"
50330      IfLoggingWriteLogfile tstr
50340      WriteToSpecialLogfile tstr
50350      PSHeader = GetPSHeader(tFile(1))
50360      tDate = Now
50370      With PDFDocInfo
50380       If Options.UseStandardAuthor = 1 Then
50390         .Author = GetSubstFilename(tFile(1), RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardAuthor)), True)
50400        Else
50410         .Author = GetDocUsername(tFile(1), False)
50420       End If
50430       If LenB(PSHeader.CreationDate.Comment) > 0 Then
50440         tstr = PSHeader.CreationDate.Comment
50450        Else
50460         tstr = CStr(tDate)
50470       End If
50480       .CreationDate = GetDocDate(Trim$(Options.StandardCreationdate), Options.StandardDateformat, FormatPrintDocumentDate(tstr))
50490       .Creator = App.EXEName & " Version " & App.Major & "." & App.Minor & "." & App.Revision
50500       .Keywords = GetSubstFilename(tFile(1), RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardKeywords)))
50510       'tStr = CStr(tDate)
50520       .ModifyDate = GetDocDate(Trim$(Options.StandardModifydate), Options.StandardDateformat, FormatPrintDocumentDate(tstr))
50530       .Subject = GetSubstFilename(tFile(1), RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardSubject)))
50540       If Len(Options.StandardTitle) > 0 Then
50550         .Title = GetSubstFilename(tFile(1), RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardTitle)))
50560        Else
50570         .Title = GetSubstFilename(tFile(1), Options.SaveFilename)
50580       End If
50590      End With
50600      AppendPDFDocInfo tFile(1), PDFDocInfo
50610      CheckForStamping tFile(1)
50620      If Options.RunProgramBeforeSaving = 1 Then
50630       RunProgramBeforeSaving Me.hwnd, GetShortName(tFile(1)), _
       Options.RunProgramBeforeSavingProgramParameters, _
       Options.RunProgramBeforeSavingWindowstyle
50660      End If
50670      CallGScript tFile(1), OutputFilename, Options, Options.AutosaveFormat
50680      If FileExists(OutputFilename) = True Then
50690        tstr = "Autosavemodus: Create File '" & OutputFilename & "' success"
50700        IfLoggingWriteLogfile tstr
50710        WriteToSpecialLogfile tstr
50720        If Options.RunProgramAfterSaving = 1 Then
50730         RunProgramAfterSaving Me.hwnd, GetShortName(OutputFilename), _
        Options.RunProgramAfterSavingProgramParameters, _
        Options.RunProgramAfterSavingWindowstyle
50760        End If
50770       Else
50780        tstr = "Autosavemodus: Create File '" & OutputFilename & "' failed"
50790        IfLoggingWriteLogfile tstr
50800        WriteToSpecialLogfile tstr
50810      End If
50820     Else
50830      IfLoggingWriteLogfile "Error: Invalid autosave pathname, spoolfile will be deleted!"
50840    End If
50850    KillFile tFile(1)
50860    KillInfoSpoolfile tFile(1)
50870   Next i
50880   Set tColl = GetFiles(GetPDFCreatorTempfolder, "~PS*.tmp", SortedByDate)
50890  Loop
50900 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50910 Exit Sub
ErrPtnr_OnError:
50931 Select Case ErrPtnr.OnError("frmMain", "Autosave")
      Case 0: Resume
50950 Case 1: Resume Next
50960 Case 2: Exit Sub
50970 Case 3: End
50980 End Select
50990 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub ShowPaypalMenuimage()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim h1 As Long, h2 As Long, com As Long
50050  h1 = GetMenu(Me.hwnd): h2 = GetSubMenu(h1, 4)
50060  com = GetMenuItemID(h2, 2)
50070  ModifyMenu h2, com, MF_BYCOMMAND Or MF_BITMAP, com, CLng(imgPaypal.Picture)
50080 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50090 Exit Sub
ErrPtnr_OnError:
50111 Select Case ErrPtnr.OnError("frmMain", "ShowPaypalMenuimage")
      Case 0: Resume
50130 Case 1: Resume Next
50140 Case 2: Exit Sub
50150 Case 3: End
50160 End Select
50170 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetGSRevision()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim RNum As String
50050  If Len(CStr(GSRevision.intRevision)) >= 3 Then
50060    RNum = Mid(CStr(GSRevision.intRevision), Len(CStr(GSRevision.intRevision)) - 1, 2)
50070    RNum = Mid(CStr(GSRevision.intRevision), 1, Len(CStr(GSRevision.intRevision)) - 2) & "." & RNum
50080   Else
50090    RNum = ""
50100  End If
50110  If Len(GSRevision.strProduct) > 0 Then
50120    stb.Panels("GhostscriptRevision").Text = GSRevision.strProduct & " " & RNum
50130   Else
50140    stb.Panels("GhostscriptRevision").Text = "-"
50150  End If
50160 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50170 Exit Sub
ErrPtnr_OnError:
50191 Select Case ErrPtnr.OnError("frmMain", "SetGSRevision")
      Case 0: Resume
50210 Case 1: Resume Next
50220 Case 2: Exit Sub
50230 Case 3: End
50240 End Select
50250 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function IsCompatibleLanguageVersion(Version As String) As Boolean
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim i As Byte, delim As String, fVers() As String, fCVers() As String, _
  ProgVersion As String, fPVers() As String
50060  IsCompatibleLanguageVersion = False
50070  delim = "."
50080  ProgVersion = GetProgramRelease
50090  If Len(CompatibleLanguageVersion) = 0 Or Len(Version) = 0 Or Len(ProgVersion) = 0 Then
50100   Exit Function
50110  End If
50120  If InStr(1, CompatibleLanguageVersion, delim) = 0 Or _
    InStr(1, Version, delim) = 0 Or _
    InStr(1, ProgVersion, delim) = 0 Then
50150   Exit Function
50160  End If
50170  fVers = Split(Version, delim)
50180  fCVers = Split(CompatibleLanguageVersion, delim)
50190  fPVers = Split(ProgVersion, delim)
50200  If UBound(fVers) < 2 Or UBound(fCVers) < 2 Or UBound(fPVers) < 2 Then
50210   Exit Function
50220  End If
50230  For i = 0 To 2
50240   If IsNumeric(fVers(i)) = False Or IsNumeric(fCVers(i)) = False Or _
   IsNumeric(fPVers(i)) = False Then
50260    Exit Function
50270   End If
50280  Next i
50290  If CLng(fVers(0)) >= CLng(fCVers(0)) And CLng(fVers(0)) <= CLng(fPVers(0)) Then
50300   If CLng(fVers(1)) >= CLng(fCVers(1)) And CLng(fVers(1)) <= CLng(fPVers(1)) Then
50310    If CLng(fVers(2)) >= CLng(fCVers(2)) And CLng(fVers(2)) <= CLng(fPVers(2)) Then
50320     IsCompatibleLanguageVersion = True
50330    End If
50340   End If
50350  End If
50360 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50370 Exit Function
ErrPtnr_OnError:
50391 Select Case ErrPtnr.OnError("frmMain", "IsCompatibleLanguageVersion")
      Case 0: Resume
50410 Case 1: Resume Next
50420 Case 2: Exit Function
50430 Case 3: End
50440 End Select
50450 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub Timer2_Timer()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim mutex As clsMutex
50050  Set mutex = New clsMutex
50060  If mutex.CheckMutex(PDFCreator_GUID) = False Then
50070  ' Create a mutex
50080    mutex.CreateMutex PDFCreator_GUID
50090  End If
50100 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50110 Exit Sub
ErrPtnr_OnError:
50131 Select Case ErrPtnr.OnError("frmMain", "Timer2_Timer")
      Case 0: Resume
50150 Case 1: Resume Next
50160 Case 2: Exit Sub
50170 Case 3: End
50180 End Select
50190 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub InitToolbar()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim btn As MSComctlLib.Button
50050  With tlb(0)
50060   Set .ImageList = imlTlb
50070   .Buttons.Clear
50080   .Buttons.Add , , , , 1
50090   .Buttons.Add , , , , 3
50100   .Buttons.Add , , , , 4
50110   .Buttons.Add , , , tbrSeparator
50120   .Buttons.Add , , , , 5
50130   .Buttons.Add , , , , 6
50140   .Buttons.Add , , , , 7
50150   .Buttons.Add , , , , 8
50160   .Buttons.Add , , , , 9
50170   .Buttons.Add , , , , 10
50180   .Buttons.Add , , , , 11
50190   .Buttons.Add , , , , 12
50200   .Buttons.Add , , , , 13
50210   .Buttons.Add , , , tbrSeparator
50220   .Buttons.Add , , , , 14
50230  End With
50240  With tlb(1)
50250   Set .ImageList = imlTlb
50260   .Buttons.Clear
50270   .Buttons.Add , , , , 15
50280   .Buttons.Add , , , tbrSeparator
50290   .Buttons.Add , , , , 16
50300   .Buttons.Add , , , , 17
50310   .Buttons.Add , , , tbrSeparator
50320   Set btn = .Buttons.Add(, "emailAddress", , tbrPlaceholder)
50330   btn.Width = (tlb(0).Buttons(15).Left + tlb(0).Buttons(15).Width) - btn.Left
50340  End With
50350  With txtEmailAddress
50360   .Top = tlb(1).Buttons("emailAddress").Top - (.Height - tlb(1).Height) / 2
50370   .Left = tlb(1).Buttons("emailAddress").Left
50380  End With
50390
50400  SetLanguageToolbar
50410  If (Options.Toolbars And 2) <> 2 Then
50420   txtEmailAddress.Enabled = False
50430   txtEmailAddress.Visible = False
50440   mnDocument(14).Enabled = False
50450   mnDocument(15).Enabled = False
50460   mnDocument(16).Enabled = False
50470   mnDocument(13).Visible = False
50480   mnDocument(14).Visible = False
50490   mnDocument(15).Visible = False
50500   mnDocument(16).Visible = False
50510  End If
50520  If (Options.Toolbars And 1) = 1 Then
50530    tlb(0).Visible = True
50540   Else
50550    tlb(0).Visible = False
50560  End If
50570  If (Options.Toolbars And 2) = 2 Then
50580    tlb(1).Visible = True
50590    txtEmailAddress.Visible = True
50600   Else
50610    tlb(1).Visible = False
50620    txtEmailAddress.Visible = False
50630  End If
50640 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50650 Exit Sub
ErrPtnr_OnError:
50671 Select Case ErrPtnr.OnError("frmMain", "InitToolbar")
      Case 0: Resume
50690 Case 1: Resume Next
50700 Case 2: Exit Sub
50710 Case 3: End
50720 End Select
50730 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub DocumentPrint()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim i As Long, j As Long
50050  For j = 1 To LvwGetCountSelectedItems(lsv, True)
50060   DoEvents
50070   For i = lsv.ListItems.Count To 1 Step -1
50080    If lsv.ListItems(i).Selected = True Then
50090     lsv.ListItems(i).SubItems(1) = LanguageStrings.ListPrinting
50100     LvwListItemToTop lsv, i, True
50110     Exit For
50120    End If
50130   Next i
50140  Next j
50150  PrintSelectedJobs = True
50160 ' SetPrinterStop False
50170 ' mnPrinter(0).Checked = False
50180 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50190 Exit Sub
ErrPtnr_OnError:
50211 Select Case ErrPtnr.OnError("frmMain", "DocumentPrint")
      Case 0: Resume
50230 Case 1: Resume Next
50240 Case 2: Exit Sub
50250 Case 3: End
50260 End Select
50270 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub DocumentAdd()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim i As Long, Cancel As Boolean, cFiles As Collection, aLen As Double, _
  OnlyPsFiles As Boolean, Ext As String, DefaultPrintername As String, _
  tFilename As String, tLen As Double
50070  Set cFiles = GetFilename("", GetMyFiles, 0, _
  LanguageStrings.ListPostscriptFiles & " (*.ps)|*.ps|" & LanguageStrings.ListAllFiles & " (*.*)|*.*", _
   OpenFile, Cancel, Me)
50100  If Cancel = True Then
50110   Screen.MousePointer = vbNormal
50120   Exit Sub
50130  End If
50140  aLen = 0
50150  For i = 1 To cFiles.Count
50160   aLen = aLen + FileLen(cFiles.Item(i))
50170  Next i
50180  OnlyPsFiles = True
50190  For i = 1 To cFiles.Count
50200   SplitPath cFiles.Item(i), , , , , Ext
50210   If UCase$(Ext) <> "PS" Then
50220    OnlyPsFiles = False
50230    Exit For
50240   End If
50250  Next i
50260  If UCase$(Printer.DeviceName) <> UCase$(GetPDFCreatorPrintername) And OnlyPsFiles = False Then
50270   If Options.NoConfirmMessageSwitchingDefaultprinter = 0 Then
50280    If ChangeDefaultprinter = False Then
50290     frmSwitchDefaultprinter.Show vbModal, Me
50300     If ChangeDefaultprinter = False Then
50310      Screen.MousePointer = vbNormal
50320      Exit Sub
50330     End If
50340    End If
50350   End If
50360  End If
50370  ChangeDefaultprinter = True
50380  DefaultPrintername = Printer.DeviceName
50390  SetDefaultprinterInProg GetPDFCreatorPrintername
50400  aLen = 0
50410  For i = 1 To cFiles.Count
50420   aLen = aLen + FileLen(cFiles.Item(i))
50430  Next i
50440  For i = 1 To cFiles.Count
50450   If IsPostscriptFile(cFiles.Item(i)) = True Then
50460     tFilename = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory, "~PS")
50470     DoEvents
50480     FileCopy cFiles.Item(i), tFilename
50490    Else
50500     DoEvents
50510     ShellAndWait Me.hwnd, "print", cFiles.Item(i), "", vbNullChar, wHidden, WCTermination, 60000, True
50520     DoEvents
50530   End If
50540   tLen = tLen + FileLen(cFiles.Item(i))
50550   stb.Panels("Percent").Text = Format$(tLen / aLen, " 0.0%")
50560   DoEvents
50570  Next i
50580  If DefaultPrintername <> vbNullString Then
50590   SetDefaultprinterInProg DefaultPrintername
50600  End If
50610  stb.Panels("Percent").Text = vbNullString
50620 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50630 Exit Sub
ErrPtnr_OnError:
50651 Select Case ErrPtnr.OnError("frmMain", "DocumentAdd")
      Case 0: Resume
50670 Case 1: Resume Next
50680 Case 2: Exit Sub
50690 Case 3: End
50700 End Select
50710 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub DocumentDelete()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim i As Long
50050  For i = 1 To lsv.ListItems.Count
50060   If lsv.ListItems(i).Selected = True Then
50070    KillFile lsv.ListItems(i).SubItems(4)
50080    KillInfoSpoolfile lsv.ListItems(i).SubItems(4)
50090   End If
50100   DoEvents
50110  Next i
50120  LvwRemoveSelectedItems lsv, True
50130  If lsv.ListItems.Count > 0 Then
50140   If lsv.SelectedItem.Index > lsv.ListItems.Count Then
50150     lsv.ListItems(lsv.SelectedItem.Index - 1).Selected = True
50160    Else
50170     lsv.ListItems(lsv.SelectedItem.Index).Selected = True
50180   End If
50190  End If
50200 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50210 Exit Sub
ErrPtnr_OnError:
50231 Select Case ErrPtnr.OnError("frmMain", "DocumentDelete")
      Case 0: Resume
50250 Case 1: Resume Next
50260 Case 2: Exit Sub
50270 Case 3: End
50280 End Select
50290 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub DocumentTop()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim i As Long, j As Long
50050  For j = 1 To LvwGetCountSelectedItems(lsv, True)
50060   For i = lsv.ListItems.Count To 1 Step -1
50070    If lsv.ListItems(i).Selected = True Then
50080     LvwListItemToTop lsv, i, True
50090     Exit For
50100    End If
50110   Next i
50120  Next j
50130 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50140 Exit Sub
ErrPtnr_OnError:
50161 Select Case ErrPtnr.OnError("frmMain", "DocumentTop")
      Case 0: Resume
50180 Case 1: Resume Next
50190 Case 2: Exit Sub
50200 Case 3: End
50210 End Select
50220 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub DocumentUp()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  LvwListItemUp lsv, , True
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("frmMain", "DocumentUp")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub DocumentDown()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  LvwListItemDown lsv, , True
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("frmMain", "DocumentDown")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub DocumentBottom()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim i As Long, j As Long
50050  For j = 1 To LvwGetCountSelectedItems(lsv, True)
50060   For i = 1 To lsv.ListItems.Count
50070    If lsv.ListItems(i).Selected = True Then
50080     LvwListItemToBottom lsv, i, True
50090     Exit For
50100    End If
50110   Next i
50120  Next j
50130 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50140 Exit Sub
ErrPtnr_OnError:
50161 Select Case ErrPtnr.OnError("frmMain", "DocumentBottom")
      Case 0: Resume
50180 Case 1: Resume Next
50190 Case 2: Exit Sub
50200 Case 3: End
50210 End Select
50220 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub DocumentCombine()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim i As Long, cFiles As Collection, tFilename As String, tFilename2 As String
50050  Screen.MousePointer = vbHourglass
50060  LockWindowUpdate lsv.hwnd
50070  Set cFiles = New Collection
50080  For i = 1 To lsv.ListItems.Count
50090   If lsv.ListItems(i).Selected = True Then
50100    cFiles.Add lsv.ListItems(i).SubItems(4)
50110   End If
50120  Next i
50130  tFilename = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory, "~PS")
50140  KillFile tFilename
50150  If cFiles.Count > 1 Then
50160   CombineFiles tFilename, cFiles, , stb
50170   stb.Panels("Percent").Text = ""
50180  End If
50190  Set cFiles = Nothing
50200  LockWindowUpdate 0&
50210  Screen.MousePointer = vbNormal
50220 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50230 Exit Sub
ErrPtnr_OnError:
50251 Select Case ErrPtnr.OnError("frmMain", "DocumentCombine")
      Case 0: Resume
50270 Case 1: Resume Next
50280 Case 2: Exit Sub
50290 Case 3: End
50300 End Select
50310 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub DocumentSave()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040   Dim tFilename As String, cFiles As Collection, Cancel As Boolean
50050  tFilename = ReplaceForbiddenChars(GetPSTitle(lsv.ListItems(lsv.SelectedItem.Index).SubItems(4)), , ".")
50060  If LenB(tFilename) = 0 Then
50070   SplitPath lsv.ListItems(lsv.SelectedItem.Index).SubItems(4), , , , tFilename
50080  End If
50090  Set cFiles = GetFilename(tFilename, GetMyFiles, 0, _
  LanguageStrings.ListPostscriptFiles & " (*.ps)|*.ps", _
   SaveFile, Cancel, Me)
50120  If Cancel = True Then
50130   Screen.MousePointer = vbNormal
50140   Exit Sub
50150  End If
50160  If cFiles.Count > 0 Then
50170   FileCopy lsv.ListItems(lsv.SelectedItem.Index).SubItems(4), cFiles.Item(1)
50180  End If
50190 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50200 Exit Sub
ErrPtnr_OnError:
50221 Select Case ErrPtnr.OnError("frmMain", "DocumentSave")
      Case 0: Resume
50240 Case 1: Resume Next
50250 Case 2: Exit Sub
50260 Case 3: End
50270 End Select
50280 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetMenuPrinterStop()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If mnPrinter(0).Checked = False Or NoProcessing = True Then
50050    SetPrinterStop True
50060    mnPrinter(0).Checked = True
50070    tlb(0).Buttons(1).Image = 2
50080   Else
50090    SetPrinterStop False
50100    mnPrinter(0).Checked = False
50110    tlb(0).Buttons(1).Image = 1
50120  End If
50130  If Not m_frmSysTray Is Nothing Then
50140   If mnPrinter(0).Checked = True Then
50150     m_frmSysTray.IconHandle = imlSystray.ListImages(0).Picture.handle
50160     m_frmSysTray.mnuSysTray(2).Checked = True
50170    Else
50180     m_frmSysTray.IconHandle = imlSystray.ListImages(1).Picture.handle
50190     m_frmSysTray.mnuSysTray(2).Checked = False
50200   End If
50210   m_frmSysTray.IconHandle = m_frmSysTray.Icon.handle
50220  End If
50230 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50240 Exit Sub
ErrPtnr_OnError:
50261 Select Case ErrPtnr.OnError("frmMain", "SetMenuPrinterStop")
      Case 0: Resume
50280 Case 1: Resume Next
50290 Case 2: Exit Sub
50300 Case 3: End
50310 End Select
50320 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetLanguageToolbar()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  With tlb(0)
50050   .Buttons(1).ToolTipText = LanguageStrings.DialogPrinterPrinterStop
50060   .Buttons(2).ToolTipText = LanguageStrings.DialogPrinterOptions
50070   .Buttons(3).ToolTipText = LanguageStrings.DialogPrinterLogfile
50080   .Buttons(5).ToolTipText = LanguageStrings.DialogDocumentPrint
50090   .Buttons(6).ToolTipText = LanguageStrings.DialogDocumentAdd
50100   .Buttons(7).ToolTipText = LanguageStrings.DialogDocumentDelete
50110   .Buttons(8).ToolTipText = LanguageStrings.DialogDocumentTop
50120   .Buttons(9).ToolTipText = LanguageStrings.DialogDocumentUp
50130   .Buttons(10).ToolTipText = LanguageStrings.DialogDocumentDown
50140   .Buttons(11).ToolTipText = LanguageStrings.DialogDocumentBottom
50150   .Buttons(12).ToolTipText = LanguageStrings.DialogDocumentCombine
50160   .Buttons(13).ToolTipText = LanguageStrings.DialogDocumentSave
50170   .Buttons(15).ToolTipText = "?"
50180  End With
50190  With tlb(1)
50200   .Buttons(1).ToolTipText = LanguageStrings.DialogDocumentCombineAll
50210   .Buttons(3).ToolTipText = LanguageStrings.DialogDocumentCombineAllSend
50220   .Buttons(4).ToolTipText = LanguageStrings.DialogDocumentSend
50230  End With
50240 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50250 Exit Sub
ErrPtnr_OnError:
50271 Select Case ErrPtnr.OnError("frmMain", "SetLanguageToolbar")
      Case 0: Resume
50290 Case 1: Resume Next
50300 Case 2: Exit Sub
50310 Case 3: End
50320 End Select
50330 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function GetVisibleToolbars() As Long
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim c As Long
50050  c = 0
50060  If (Options.Toolbars And 1) = 1 Then
50070   c = c + 1
50080  End If
50090  If (Options.Toolbars And 2) = 2 Then
50100   c = c + 1
50110  End If
50120  GetVisibleToolbars = c
50130 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50140 Exit Function
ErrPtnr_OnError:
50161 Select Case ErrPtnr.OnError("frmMain", "GetVisibleToolbars")
      Case 0: Resume
50180 Case 1: Resume Next
50190 Case 2: Exit Function
50200 Case 3: End
50210 End Select
50220 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function CombineAll() As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim i As Long, cFiles As Collection, tFilename As String, tFilename2 As String
50050  Screen.MousePointer = vbHourglass
50060  LockWindowUpdate hwnd
50070  Timer1.Enabled = False
50080  Set cFiles = New Collection
50090  For i = 1 To lsv.ListItems.Count
50100   cFiles.Add lsv.ListItems(i).SubItems(4)
50110  Next i
50120  tFilename = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory, "~PT")
50130  KillFile tFilename
50140  If cFiles.Count > 1 Then
50150   CombineFiles tFilename, cFiles, , stb
50160   stb.Panels("Percent").Text = ""
50170  End If
50180  tFilename2 = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory, "~PS")
50190  KillFile tFilename2
50200  Name tFilename As tFilename2
50210  Set cFiles = Nothing
50220  CombineAll = tFilename2
50230  Timer1.Enabled = True
50240  LockWindowUpdate 0&
50250  Screen.MousePointer = vbNormal
50260 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50270 Exit Function
ErrPtnr_OnError:
50291 Select Case ErrPtnr.OnError("frmMain", "CombineAll")
      Case 0: Resume
50310 Case 1: Resume Next
50320 Case 2: Exit Function
50330 Case 3: End
50340 End Select
50350 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub SendEmailImmediately(InputFilename As String)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim OutputFilename As String, mail As clsPDFCreatorMail, rec As String
50050  If LenB(InputFilename) > 0 Then
50060   If FileExists(InputFilename) = True And FileInUse(InputFilename) = False Then
50070    rec = LTrim(txtEmailAddress.Text)
50080    If LenB(LTrim(Options.StandardMailDomain)) > 0 And InStr(1, LTrim(txtEmailAddress.Text), "@") <= 0 Then
50090     rec = rec & "@" & Options.StandardMailDomain
50100    End If
50110    OutputFilename = CompletePath(GetPDFCreatorTempfolder) & txtEmailAddress.Text & ".pdf"
50120
50130    If Options.RunProgramBeforeSaving = 1 Then
50140     RunProgramBeforeSaving Me.hwnd, GetShortName(InputFilename), _
    Options.RunProgramBeforeSavingProgramParameters, _
    Options.RunProgramBeforeSavingWindowstyle
50170    End If
50180    ConvertPostscriptFile InputFilename, OutputFilename
50190    KillFile InputFilename
50200    If Len(OutputFilename) > 0 And FileExists(OutputFilename) = True Then
50210     If Options.RunProgramAfterSaving = 1 Then
50220      RunProgramAfterSaving Me.hwnd, GetShortName(OutputFilename), _
      Options.RunProgramAfterSavingProgramParameters, _
      Options.RunProgramAfterSavingWindowstyle
50250     End If
50260     Set mail = New clsPDFCreatorMail
50270     If mail.Send(OutputFilename, Options.StandardSubject, Options.SendMailMethod, rec) <> 0 Then
50280      MsgBox LanguageStrings.MessagesMsg04, vbCritical, App.EXEName
50290     End If
50300     Set mail = Nothing
50310    End If
50320   End If
50330  End If
50340  Timer1.Enabled = True
50350 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50360 Exit Sub
ErrPtnr_OnError:
50381 Select Case ErrPtnr.OnError("frmMain", "SendEmailImmediately")
      Case 0: Resume
50400 Case 1: Resume Next
50410 Case 2: Exit Sub
50420 Case 3: End
50430 End Select
50440 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub CombineAllAndSend()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Timer1.Enabled = False
50050  If Options.ShowAnimation = 1 Then
50060   ShowAnimationWindow = True
50070   frmAnimation.Show
50080  End If
50090  SendEmailImmediately CombineAll
50100  If Options.ShowAnimation = 1 Then
50110   ShowAnimationWindow = False
50120  End If
50130  Timer1.Enabled = True
50140 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50150 Exit Sub
ErrPtnr_OnError:
50171 Select Case ErrPtnr.OnError("frmMain", "CombineAllAndSend")
      Case 0: Resume
50190 Case 1: Resume Next
50200 Case 2: Exit Sub
50210 Case 3: End
50220 End Select
50230 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SendEmail()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If lsv.ListItems.Count > 0 Then
50050   If lsv.SelectedItem.Index >= 0 Then
50060    If FileExists(lsv.SelectedItem.ListSubItems(4)) = True And FileInUse(lsv.SelectedItem.ListSubItems(4)) = False Then
50070     Timer1.Enabled = False
50080     If Options.ShowAnimation = 1 Then
50090      ShowAnimationWindow = True
50100      frmAnimation.Show
50110     End If
50120     SendEmailImmediately lsv.SelectedItem.ListSubItems(4)
50130     If Options.ShowAnimation = 1 Then
50140      ShowAnimationWindow = False
50150      Timer1.Enabled = True
50160     End If
50170    End If
50180   End If
50190  End If
50200 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50210 Exit Sub
ErrPtnr_OnError:
50231 Select Case ErrPtnr.OnError("frmMain", "SendEmail")
      Case 0: Resume
50250 Case 1: Resume Next
50260 Case 2: Exit Sub
50270 Case 3: End
50280 End Select
50290 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function AllSelectedListitemsAtTop() As Boolean
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim i As Long, tB As Boolean
50050  AllSelectedListitemsAtTop = False
50060  If lsv.ListItems.Count > 0 Then
50070   If lsv.ListItems(1).Selected = False Then
50080    Exit Function
50090   End If
50100   AllSelectedListitemsAtTop = True
50110   tB = False
50120   For i = 2 To lsv.ListItems.Count
50130    If tB = True And lsv.ListItems(i).Selected = True Then
50140     AllSelectedListitemsAtTop = False
50150     Exit For
50160    End If
50170    If lsv.ListItems(i).Selected = False Then
50180     tB = True
50190    End If
50200   Next i
50210  End If
50220 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50230 Exit Function
ErrPtnr_OnError:
50251 Select Case ErrPtnr.OnError("frmMain", "AllSelectedListitemsAtTop")
      Case 0: Resume
50270 Case 1: Resume Next
50280 Case 2: Exit Function
50290 Case 3: End
50300 End Select
50310 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function AllSelectedListitemsAtBottom() As Boolean
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim i As Long, tB As Boolean
50050  AllSelectedListitemsAtBottom = False
50060  If lsv.ListItems.Count > 0 Then
50070   If lsv.ListItems(lsv.ListItems.Count).Selected = False Then
50080    Exit Function
50090   End If
50100   AllSelectedListitemsAtBottom = True
50110   tB = False
50120   For i = lsv.ListItems.Count - 1 To 1 Step -1
50130    If tB = True And lsv.ListItems(i).Selected = True Then
50140     AllSelectedListitemsAtBottom = False
50150     Exit For
50160    End If
50170    If lsv.ListItems(i).Selected = False Then
50180     tB = True
50190    End If
50200   Next i
50210  End If
50220 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50230 Exit Function
ErrPtnr_OnError:
50251 Select Case ErrPtnr.OnError("frmMain", "AllSelectedListitemsAtBottom")
      Case 0: Resume
50270 Case 1: Resume Next
50280 Case 2: Exit Function
50290 Case 3: End
50300 End Select
50310 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub m_frmSysTray_MenuClick(ByVal lIndex As Long, ByVal sKey As String)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50041  Select Case lIndex
        Case 0
50060    Me.ZOrder
50070    Me.WindowState = vbNormal
50080    Me.Show
50090    SysTrayLeave
50100   Case Is > 1
50110    mnPrinter_Click CInt(lIndex - 2)
50120  End Select
50130 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50140 Exit Sub
ErrPtnr_OnError:
50161 Select Case ErrPtnr.OnError("frmMain", "m_frmSysTray_MenuClick")
      Case 0: Resume
50180 Case 1: Resume Next
50190 Case 2: Exit Sub
50200 Case 3: End
50210 End Select
50220 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub m_frmSysTray_SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Me.WindowState = vbNormal
50050  Me.ZOrder
50060  Me.Show
50070  SysTrayLeave
50080 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50090 Exit Sub
ErrPtnr_OnError:
50111 Select Case ErrPtnr.OnError("frmMain", "m_frmSysTray_SysTrayDoubleClick")
      Case 0: Resume
50130 Case 1: Resume Next
50140 Case 2: Exit Sub
50150 Case 3: End
50160 End Select
50170 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub m_frmSysTray_SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If (eButton = vbRightButton) Then
50050   m_frmSysTray.ShowMenu
50060  End If
50070 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50080 Exit Sub
ErrPtnr_OnError:
50101 Select Case ErrPtnr.OnError("frmMain", "m_frmSysTray_SysTrayMouseDown")
      Case 0: Resume
50120 Case 1: Resume Next
50130 Case 2: Exit Sub
50140 Case 3: End
50150 End Select
50160 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SystrayEnter()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim i As Long
50050  If m_frmSysTray Is Nothing Then
50060   Set m_frmSysTray = New frmSysTray
50070   With m_frmSysTray
50080    .AddMenuItem App.EXEName, , True
50090    .AddMenuItem "-"
50100    For i = mnPrinter.LBound To mnPrinter.UBound
50110     .AddMenuItem mnPrinter(i).Caption
50120    Next i
50130    If mnPrinter(4).Checked = True Then
50140     .mnuSysTray(6).Checked = True
50150    End If
50160    If mnPrinter(0).Checked = True Then
50170     .mnuSysTray(2).Checked = True
50180    End If
50190    .ToolTip = Me.Caption
50200   End With
50210   If mnPrinter(0).Checked = True Then
50220     m_frmSysTray.IconHandle = imlSystray.ListImages(1).Picture.handle
50230 '    m_frmSysTray.Icon = LoadResPicture(2120, vbResIcon)
50240    Else
50250     m_frmSysTray.IconHandle = imlSystray.ListImages(2).Picture.handle
50260 '    m_frmSysTray.Icon = LoadResPicture(2121, vbResIcon)
50270   End If
50280   Me.Hide
50290 '  m_frmSysTray.ShowBalloonTip "Here I am in the SysTray", _
   "PDFCreator Print Monitor", NIIF_INFO
50310  End If
50320 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50330 Exit Sub
ErrPtnr_OnError:
50351 Select Case ErrPtnr.OnError("frmMain", "SystrayEnter")
      Case 0: Resume
50370 Case 1: Resume Next
50380 Case 2: Exit Sub
50390 Case 3: End
50400 End Select
50410 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SysTrayLeave()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If Not m_frmSysTray Is Nothing Then
50050   Unload m_frmSysTray
50060   Set m_frmSysTray = Nothing
50070  End If
50080 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50090 Exit Sub
ErrPtnr_OnError:
50111 Select Case ErrPtnr.OnError("frmMain", "SysTrayLeave")
      Case 0: Resume
50130 Case 1: Resume Next
50140 Case 2: Exit Sub
50150 Case 3: End
50160 End Select
50170 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SetSystrayIcon(Index As Long)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  If Not m_frmSysTray Is Nothing Then
50050   m_frmSysTray.IconHandle = frmMain.imlSystray.ListImages(Index).Picture.handle
50060  End If
50070 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50080 Exit Sub
ErrPtnr_OnError:
50101 Select Case ErrPtnr.OnError("frmMain", "SetSystrayIcon")
      Case 0: Resume
50120 Case 1: Resume Next
50130 Case 2: Exit Sub
50140 Case 3: End
50150 End Select
50160 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
