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
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If KeyCode = vbKeyF1 Then
50020   KeyCode = 0
50030   Call HTMLHelp_ShowTopic("html\welcome.htm")
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
50010  Me.KeyPreview = True
50020
50030  If App.StartMode = vbSModeAutomation Then
50040   If ProgramIsVisible = False Then
50050    frmMain.Visible = False
50060   End If
50070   WindowState = ProgramWindowState
50080  End If
50090
50100  ReadAllLanguages LanguagePath
50110  InitProgram
50120
50130  ShowPaypalMenuimage
50140  SetGSRevision
50150
50160  If NoProcessing = True Or Options.NoProcessingAtStartup = 1 Or NoProcessingAtStartup = True Then
50170   SetMenuPrinterStop
50180  End If
50190  If PrinterStop = True Then
50200   mnPrinter(0).Checked = True
50210  End If
50220  If Options.OptionsEnabled = 0 Then
50230   mnPrinter(2).Enabled = False
50240   tlb(0).Buttons(2).Enabled = False
50250  End If
50260  If Options.OptionsVisible = 0 Then
50270   tlb(0).Buttons(2).Visible = False
50280   mnPrinter(2).Visible = False
50290   mnPrinter(3).Visible = False
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
50520    SetActiveWindow frmMain.hwnd
50530   End If
50540  End If
50550
50560  ' Only for the first time set interval to 10 ms
50570  Timer1.Interval = 10
50580  Timer1.Enabled = True
50590  Timer2.Interval = 100
50600  Timer2.Enabled = True
50610
50620  ProgramIsStarted = True
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
50250   SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50260  End With
50270
50280  SetLanguageMenu
50290  If Options.Logging = 1 Then
50300    mnPrinter(4).Checked = True
50310   Else
50320    mnPrinter(4).Checked = False
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
50010  Dim PDFSpoolerPath As String, Files As Collection, i As Long, tStrf() As String
50020  Timer1.Enabled = False
50030  Timer2.Enabled = False
50040
50050  Set Printjobs = Nothing
50060
50070  If Not mutexLocal Is Nothing Then
50080   mutexLocal.CloseMutex
50090   Set mutexLocal = Nothing
50100  End If
50110
50120  If Not mutexGlobal Is Nothing Then
50130   mutexGlobal.CloseMutex
50140   Set mutexGlobal = Nothing
50150  End If
50160
50170  UnloadDLLComplete GsDllLoaded
50180
50190  FindFiles CompletePath(GetPDFCreatorTempfolder), Files, "*.pdf", , False, True
50200  For i = 1 To Files.Count
50210   tStrf = Split(Files(i), "|")
50220   If FileExists(tStrf(1)) And Not FileInUse(tStrf(1)) Then
50230    KillFile tStrf(1)
50240   End If
50250  Next i
50260
50270  IfLoggingWriteLogfile "PDFCreator Program End"
50280  SysTrayLeave
50290  If App.StartMode = vbSModeStandalone Then
50300   InstanceCounter = InstanceCounter - 1
50310  End If
50320  PDFSpoolerPath = GetPDFCreatorApplicationPath & "PDFSpooler.exe"
50330  If Restart = True And FileExists(PDFSpoolerPath) = True Then
50340   ShellExecute 0, vbNullString, """" & PDFSpoolerPath & """", "-SL200 -STTRUE", GetPDFCreatorApplicationPath, 1
50350  End If
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
Select Case ErrPtnr.OnError("frmMain", "GetAllLanguagesFiles")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub ReadAllLanguages(LanguagePath As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Languagename As String, ini As clsINI, LangFiles As Collection, _
  i As Long, Version As String, Filename As String
50030  mnLanguage(0).Caption = "No languages available."
50040
50050  Set LangFiles = GetAllLanguagesFiles(LanguagePath)
50060  Set ini = New clsINI
50070  For i = 1 To LangFiles.Count
50080   ini.Filename = LangFiles.Item(i)
50090   ini.Section = "Common"
50100   Languagename = ini.GetKeyFromSection("Languagename")
50110   Version = ini.GetKeyFromSection("Version")
50120   If Len(Languagename) = 0 Then
50130    Languagename = "No name available."
50140   End If
50150   Load mnLanguage(mnLanguage.Count)
50160   If IsCompatibleLanguageVersion(Version) = True Then
50170     mnLanguage(mnLanguage.Count - 1).Caption = Languagename
50180    Else
50190     mnLanguage(mnLanguage.Count - 1).Caption = Languagename & " [" & Version & "]"
50200   End If
50210   mnLanguage(mnLanguage.Count - 1).Tag = LangFiles.Item(i)
50220   SplitPath LangFiles.Item(i), , , , Filename
50230   If UCase$(Options.Language) = UCase$(Filename) Then
50240    mnLanguage(mnLanguage.Count - 1).Checked = True
50250   End If
50260   DoEvents
50270  Next i
50280
50290  If mnLanguage.Count > 1 Then
50300   mnLanguage(0).Caption = "No languages available."
50310   mnLanguage(0).Visible = False
50320  End If
50330  Set ini = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "ReadAllLanguages")
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
50010  Dim i As Long, Version As String, reg As clsRegistry, Filename As String
50020
50030  For i = mnLanguage.LBound To mnLanguage.UBound
50040   If UCase$(Languagefile) = UCase$(mnLanguage.Item(i).Tag) Then
50050     mnLanguage.Item(i).Checked = True
50060     SplitPath Languagefile, , , , Filename
50070    Else
50080     mnLanguage.Item(i).Checked = False
50090   End If
50100  Next i
50110
50121  Select Case UCase$(Filename)
        Case "GERMAN"
50140    HelpFile = GetPDFCreatorApplicationPath & "PDFCreator_german.chm"
50150    If Not FileExists(HelpFile) Then
50160     HelpFile = GetPDFCreatorApplicationPath & "PDFCreator_english.chm"
50170    End If
50180   Case Else
50190    HelpFile = GetPDFCreatorApplicationPath & "PDFCreator_english.chm"
50200  End Select
50210
50220  If Not FileExists(HelpFile) Then
50230   MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & vbCrLf & HelpFile, vbExclamation
50240   HelpFile = GetPDFCreatorApplicationPath & "PDFCreator_english.chm"
50250  End If
50260
50270  With LanguageStrings
50280   Set reg = New clsRegistry
50290   With reg
50300    .hkey = HKEY_LOCAL_MACHINE
50310    .KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50320    Version = .GetRegistryValue("ApplicationVersion")
50330   End With
50340   Set reg = Nothing
50350   Caption = App.Title & " - " & .CommonTitle
50360
50370   mnPrinterMain.Caption = .DialogPrinter
50380   mnPrinter(0).Caption = .DialogPrinterPrinterStop
50390   mnPrinter(2).Caption = .DialogPrinterOptions
50400   mnPrinter(4).Caption = .DialogPrinterLogging
50410   mnPrinter(5).Caption = .DialogPrinterLogfile
50420   mnPrinter(7).Caption = .DialogPrinterClose
50430
50440   mnDocumentMain.Caption = .DialogDocument
50450   mnDocument(0).Caption = .DialogDocumentPrint
50460   mnDocument(2).Caption = .DialogDocumentAdd
50470   mnDocument(3).Caption = .DialogDocumentDelete
50480   mnDocument(5).Caption = .DialogDocumentTop
50490   mnDocument(6).Caption = .DialogDocumentUp
50500   mnDocument(7).Caption = .DialogDocumentDown
50510   mnDocument(8).Caption = .DialogDocumentBottom
50520   mnDocument(10).Caption = .DialogDocumentCombine
50530   mnDocument(12).Caption = .DialogDocumentSave
50540
50550   mnDocument(14).Caption = .DialogDocumentCombineAll
50560   mnDocument(15).Caption = .DialogDocumentCombineAllSend
50570   mnDocument(16).Caption = .DialogDocumentSend
50580
50590   mnViewMain.Caption = .DialogView
50600   mnView(0).Caption = .DialogViewToolbars
50610   mnView(2).Caption = .DialogViewStatusbar
50620   mnViewToolbars(0).Caption = .DialogViewToolbarsStandard
50630   mnViewToolbars(1).Caption = .DialogViewToolbarsEmail
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
50790
50800   txtEmailAddress.ToolTipText = .DialogEmailAddress
50810  End With
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
    If IsPostscriptFile(Data.Files.Item(1)) = True Or UCase$(Ext) = "PS" Or UCase$(Ext) = "EPS" Then
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
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tFilename As String, cFiles As Collection, sFiles() As String, _
  i As Long, j As Long, aLen As Double, tLen As Double, aw As Long, _
  DefaultPrintername As String, OnlyPsFiles As Boolean, Ext As String, _
  Cancel As Boolean, tFilename2 As String
50050
50060  Timer1.Enabled = False
50070  Screen.MousePointer = vbHourglass
50080  DoEvents
50091  Select Case Index
        Case 0:
50110    DocumentPrint
50120   Case 2: ' Add
50130    DocumentAdd
50140   Case 3: ' Delete
50150    DocumentDelete
50160   Case 5: ' Top
50170    DocumentTop
50180   Case 6: ' Up
50190    DocumentUp
50200   Case 7: ' Down
50210    DocumentDown
50220   Case 8: ' Bottom
50230    DocumentBottom
50240   Case 10: ' Combine
50250    DocumentCombine
50260   Case 12: ' Save
50270    DocumentSave
50280   Case 14
50290    CombineAll
50300   Case 15
50310    CombineAllAndSend
50320   Case 16
50330    SendEmail
50340  End Select
50350  SetDocMenuAndToolbar
50360  Screen.MousePointer = vbNormal
50370  Timer1.Enabled = True
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
50010  Dim updStr As String, updStrA() As String, aw As Long
50020  Set dl = New clsDownload
50031  Select Case Index
        Case 0:
50050    Call HTMLHelp_ShowTopic("html\welcome.htm")
50060   Case 2:
50070    OpenDocument Paypal
50080   Case 4:
50090    OpenDocument Homepage
50100   Case 5:
50110    OpenDocument Sourceforge
50120   Case 6:
50130    updStr = dl.DownloadString(UpdateURL)
50140 '   updStr = dl.DownloadString("http://localhost:8080/update.txt")
50150    If Len(updStr) > 0 Then
50160      If CheckPDFCreatorVersion(updStr) > 0 Then
50170        updStrA = Split(updStr, ".")
50180        If updStrA(3) = 0 Then
50190          updStr = "[" & updStrA(0) & "." & updStrA(1) & "." & updStrA(2) & "]"
50200         Else
50210          updStr = "[" & updStrA(0) & "." & updStrA(1) & "." & updStrA(2) & " Beta " & updStrA(3) & "]"
50220        End If
50230        aw = MsgBox(Replace$(LanguageStrings.MessagesMsg32, "%1", updStr), vbYesNo + vbQuestion)
50240        If aw = vbYes Then
50250         OpenDocument "http://www.sourceforge.net/project/showfiles.php?group_id=57796"
50260        End If
50270       Else
50280        MsgBox LanguageStrings.MessagesMsg33, vbOKOnly + vbInformation
50290      End If
50300     Else
50310      MsgBox LanguageStrings.MessagesMsg31 & ": " & dl.ErrorDescription & " [" & dl.ErrorNumber & "]", vbOKOnly + vbExclamation
50320    End If
50330   Case 8:
50340    frmInfo.Show , Me
50350  End Select
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
50030    SetMenuPrinterStop
50040   Case 2:
50050    frmOptions.Show , Me
50060   Case 4:
50070    If mnPrinter(Index).Checked = False Then
50080      SetLogging True
50090      mnPrinter(Index).Checked = True
50100     Else
50110      SetLogging False
50120      mnPrinter(Index).Checked = False
50130    End If
50140    If Not m_frmSysTray Is Nothing Then
50150     If mnPrinter(Index).Checked = True Then
50160       m_frmSysTray.mnuSysTray(6).Checked = True
50170      Else
50180       m_frmSysTray.mnuSysTray(6).Checked = False
50190     End If
50200    End If
50210   Case 5:
50220    frmLog.Show , Me
50230   Case 7:
50240    Unload Me
50250  End Select
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

Private Sub mnLanguage_Click(Index As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim File As String
50020  Screen.MousePointer = vbHourglass
50030  LoadLanguage mnLanguage(Index).Tag
50040  Languagefile = mnLanguage(Index).Tag
50050  SetLanguageMenu
50060  SplitPath Languagefile, , , , File
50070  SetLanguage File
50080  ShowPaypalMenuimage
50090  Me.Refresh
50100  Screen.MousePointer = vbNormal
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "mnLanguage_Click")
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
50240  SaveOptions Options
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

Private Sub Timer1_Timer()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Timer1.Enabled = False
50020  DoEvents
50030  If FileExists(GetPDFCreatorApplicationPath & "Unload.tmp") = True Or Restart = True Then
50040   Unload Me
50050   Exit Sub
50060  End If
50070  CheckPrintJobs
50080  If Not NoProcessing Then
50090   CheckForPrinting
50100  End If
50110  If lsv.ListItems.Count = 0 And LenB(CommandSwitch("IF", True)) > 0 And ShellAndWaitingIsRunning = False Then
50120   Unload Me
50130   Exit Sub
50140  End If
50150  If lsv.ListItems.Count = 1 Then
50160   If lsv.SelectedItem.Index <> 1 Then
50170    lsv.ListItems(1).Selected = True
50180   End If
50190  End If
50200  DoEvents
50210  Timer1.Interval = TimerIntervall
50220  Timer1.Enabled = True
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
50150      DocumentDelete
50160     Case 8
50170      DocumentTop
50180     Case 9
50190      DocumentUp
50200     Case 10
50210      DocumentDown
50220     Case 11
50230      DocumentBottom
50240     Case 12
50250      DocumentCombine
50260     Case 13
50270      DocumentSave
50280     Case 15
50290      Call HTMLHelp_ShowTopic("html\welcome.htm")
50300    End Select
50310   Case 1
50321    Select Case Button.Index
          Case 1
50340      CombineAll
50350     Case 3
50360      CombineAllAndSend
50370     Case 4
50380      SendEmail
50390    End Select
50400  End Select
50410  SetDocMenuAndToolbar
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
50010  If lsv.ListItems.Count > 0 Then
50020   If mnPrinter(0).Checked = True Then
50030     If PrintSelectedJobs = True Then
50040       If lsv.ListItems(1).SubItems(1) = LanguageStrings.ListPrinting Then
50050         PDFSpoolfile = lsv.ListItems(1).SubItems(4)
50060         DoEvents
50070         frmPrinting.Show vbModal, Me
50080        Else
50090         PrintSelectedJobs = False
50100       End If
50110      Else
50120       If lsv.ListItems(1).SubItems(1) <> LanguageStrings.ListWaiting Then
50130        lsv.ListItems(1).SubItems(1) = LanguageStrings.ListWaiting
50140       End If
50150     End If
50160    Else
50170     If lsv.ListItems(1).SubItems(1) <> LanguageStrings.ListPrinting Then
50180      lsv.ListItems(1).SubItems(1) = LanguageStrings.ListPrinting
50190     End If
50200     PDFSpoolfile = lsv.ListItems(1).SubItems(4)
50210     If PrinterStop = False Then
50220      If IsFormLoaded(frmPrinting) = False Then
50230       If InstalledAsServer Then ' This is necessary because there is no other way to tell the running server that the options are changed!
50240        Options = ReadOptions
50250       End If
50260       If Options.UseAutosave = 1 Then
50270         Autosave
50280        Else
50290         frmPrinting.Show , Me
50300       End If
50310      End If
50320     End If
50330     If PrinterStop = False And NoProcessing = False Then
50340       mnPrinter(0).Checked = False
50350       tlb(0).Buttons(1).Image = 1
50360      Else
50370       mnPrinter(0).Checked = True
50380       tlb(0).Buttons(1).Image = 2
50390     End If
50400   End If
50410  End If
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
50010  Dim Temppath As String, LItem As ListItem, Files As Collection, tColl As Collection, _
  tFile() As String, i As Long, j As Long, kB As Long, MB As Long, GB As Long, _
  ind As Long, tStr As String, tB As Boolean
50040  kB = 1024: MB = kB * 1024: GB = MB * 1024
50050  Set Files = New Collection: Timer1.Enabled = False
50060  Temppath = CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory
50070  FindFiles Temppath, Files, "~PS*.tmp", , False, True
50080
50090  If Files.Count = 0 And lsv.ListItems.Count > 0 Then
50100   lsv.ListItems.Clear
50110   SetDocMenuAndToolbar
50120   Exit Sub
50130  End If
50140
50150  tB = False
50160
50170  Set tColl = New Collection
50180  For j = 1 To lsv.ListItems.Count
50190   ind = 0
50200   For i = 1 To Files.Count
50210    tFile = Split(Files.Item(i), "|")
50220    If UCase$(tFile(1)) = UCase$(lsv.ListItems(j).SubItems(4)) Then
50230     ind = i
50240     Exit For
50250    End If
50260   Next i
50270   If ind = 0 Then
50280    tColl.Add j
50290   End If
50300  Next j
50310  If tColl.Count > 0 Then
50320   tB = True
50330  End If
50340  For j = 1 To tColl.Count
50350   lsv.ListItems.Remove tColl(j) - (j - 1)
50360  Next j
50370
50380
50390  Set tColl = New Collection
50400  For j = 1 To Files.Count
50410   tFile = Split(Files.Item(j), "|")
50420   ind = 0
50430   For i = 1 To lsv.ListItems.Count
50440    If UCase$(tFile(1)) = UCase$(lsv.ListItems(i).SubItems(4)) Then
50450     ind = i
50460     Exit For
50470    End If
50480   Next i
50490   If ind > 0 And ind < lsv.ListItems.Count + 1 Then
50500    tColl.Add j
50510   End If
50520  Next j
50530  For j = 1 To tColl.Count
50540   Files.Remove tColl(j) - (j - 1)
50550  Next j
50560
50570  For j = 1 To Files.Count
50580   tFile = Split(Files.Item(j), "|")
50590   ind = 0
50600   For i = 1 To lsv.ListItems.Count
50610    If UCase$(tFile(1)) = UCase$(lsv.ListItems(i).SubItems(4)) Then
50620     ind = i
50630     Exit For
50640    End If
50650   Next i
50660   If ind = 0 Then
50670    tB = True
50680    If FormISLoaded("frmOptions") = False And Me.Visible = True Then
50690     Me.Show
50700    End If
50710    Set LItem = lsv.ListItems.Add(, , GetPSTitle(tFile(1)))
50720
50730    LItem.SubItems(1) = LanguageStrings.ListWaiting
50740    LItem.SubItems(2) = tFile(3)
50750
50760    If CLng(tFile(2)) > GB Then
50770      LItem.SubItems(3) = Format$(CDbl(tFile(2)) / GB, "0.00 " & LanguageStrings.ListGBytes)
50780     Else
50790      If CLng(tFile(2)) > MB Then
50800        LItem.SubItems(3) = Format$(CDbl(tFile(2)) / MB, "0.00 " & LanguageStrings.ListMBytes)
50810       Else
50820        If CLng(tFile(2)) > kB Then
50830          LItem.SubItems(3) = Format$(CDbl(tFile(2)) / kB, "0.00 " & LanguageStrings.ListKBytes)
50840         Else
50850          LItem.SubItems(3) = Format$(tFile(2), "0 " & LanguageStrings.ListBytes)
50860        End If
50870     End If
50880    End If
50890    LItem.SubItems(4) = tFile(1)
50900    DoEvents
50910   End If
50920  Next j
50930  If lsv.ListItems.Count = 1 Then
50940    tStr = LanguageStrings.ListStatus & ": " & lsv.ListItems.Count & " " & LanguageStrings.MessagesMsg01
50950   Else
50960    tStr = LanguageStrings.ListStatus & ": " & lsv.ListItems.Count & " " & LanguageStrings.MessagesMsg02
50970  End If
50980  If tStr <> stb.Panels("Status").Text Then
50990   stb.Panels("Status").Text = tStr
51000  End If
51010  If tB = True Then
51020   SetDocMenuAndToolbar
51030  End If
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
50030   mnDocument(c).Enabled = True
50040  Next c
50050  For c = 1 To tlb(0).Buttons.Count
50060   tlb(0).Buttons(c).Enabled = True
50070  Next c
50080  If (Options.Toolbars And 2) = 2 Then
50090   For c = 1 To tlb(1).Buttons.Count
50100    tlb(1).Buttons(c).Enabled = True
50110   Next c
50120  End If
50131  Select Case True
        Case lsv.ListItems.Count = 0, LvwGetCountSelectedItems(lsv, False) = 0
50150    With mnDocument
50160     .Item(0).Enabled = False
50170     .Item(3).Enabled = False
50180     .Item(5).Enabled = False
50190     .Item(6).Enabled = False
50200     .Item(7).Enabled = False
50210     .Item(8).Enabled = False
50220     .Item(10).Enabled = False
50230     .Item(12).Enabled = False
50240     If (Options.Toolbars And 2) = 2 Then
50250      .Item(14).Enabled = False
50260      .Item(15).Enabled = False
50270      .Item(16).Enabled = False
50280     End If
50290    End With
50300    With tlb(0)
50310     .Buttons(5).Enabled = False
50320     For c = 7 To 13
50330      .Buttons(c).Enabled = False
50340     Next c
50350    End With
50360    With tlb(1)
50370     If (Options.Toolbars And 2) = 2 Then
50380      .Buttons(1).Enabled = False
50390      .Buttons(3).Enabled = False
50400      .Buttons(4).Enabled = False
50410     End If
50420    End With
50430    Exit Sub
50440   Case lsv.ListItems.Count = 1
50450    With mnDocument
50460     .Item(5).Enabled = False
50470     .Item(6).Enabled = False
50480     .Item(7).Enabled = False
50490     .Item(8).Enabled = False
50500     .Item(10).Enabled = False
50510     If (Options.Toolbars And 2) = 2 Then
50520      .Item(14).Enabled = False
50530      .Item(15).Enabled = False
50540      If LenB(txtEmailAddress.Text) = 0 Or Options.DisableEmail <> 0 Then
50550       .Item(16).Enabled = False
50560      End If
50570     End If
50580    End With
50590    With tlb(0)
50600     For c = 8 To 12
50610      .Buttons(c).Enabled = False
50620     Next c
50630    End With
50640    With tlb(1)
50650     If (Options.Toolbars And 2) = 2 Then
50660      .Buttons(1).Enabled = False
50670      .Buttons(3).Enabled = False
50680      If LenB(txtEmailAddress.Text) = 0 Then
50690       .Buttons(4).Enabled = False
50700      End If
50710     End If
50720    End With
50730   Case lsv.ListItems.Count > 1
50740    With mnDocument
50750     If AllSelectedListitemsAtTop Then
50760      .Item(5).Enabled = False
50770      .Item(6).Enabled = False
50780     End If
50790     If AllSelectedListitemsAtBottom Then
50800      .Item(7).Enabled = False
50810      .Item(8).Enabled = False
50820     End If
50830     If LvwGetCountSelectedItems(lsv, False) = 1 Then
50840      .Item(10).Enabled = False
50850     End If
50860     If LvwGetCountSelectedItems(lsv, False) > 1 Then
50870      .Item(12).Enabled = False
50880     End If
50890     If ((Options.Toolbars And 2) = 2 And LenB(txtEmailAddress.Text) = 0) Or Options.DisableEmail <> 0 Then
50900      .Item(15).Enabled = False
50910      .Item(16).Enabled = False
50920     End If
50930    End With
50940    With tlb(0)
50950     If AllSelectedListitemsAtTop Then
50960      .Buttons(8).Enabled = False
50970      .Buttons(9).Enabled = False
50980     End If
50990     If AllSelectedListitemsAtBottom Then
51000      .Buttons(10).Enabled = False
51010      .Buttons(11).Enabled = False
51020     End If
51030     If LvwGetCountSelectedItems(lsv, False) = 1 Then
51040      .Buttons(12).Enabled = False
51050     End If
51060     If LvwGetCountSelectedItems(lsv, False) > 1 Then
51070      .Buttons(13).Enabled = False
51080     End If
51090    End With
51100    With tlb(1)
51110     If (Options.Toolbars And 2) = 2 And LenB(txtEmailAddress.Text) = 0 Then
51120      .Buttons(3).Enabled = False
51130      .Buttons(4).Enabled = False
51140     End If
51150    End With
51160  End Select
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

Private Sub Autosave(Optional Filename As String = vbNullString)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tColl As Collection, i As Long, tFile() As String, Pathname As String, _
  OutputFilename As String, PDFDocInfo As tPDFDocInfo, tStr As String, _
  PSHeader As tPSHeader, tDate As Date, mail As clsPDFCreatorMail
50040
50050  IsConverted = False
50060  Set tColl = New Collection
50070
50080  If Len(Filename) > 0 Then
50090    If FileExists(Filename) = True Then
50100     SplitPath Filename, , Pathname
50110     tColl.Add Pathname & "|" & Filename & "|" & FileLen(Filename) & "|" & FileDateTime(Filename)
50120    End If
50130   Else
50140    FindFiles CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory, tColl, "~PS*.tmp", , True, True
50150 '   Set tColl = GetFiles(GetPDFCreatorTempfolder, "~P*.tmp")
50160  End If
50170
50180  tStr = "Autosavemodus: " & tColl.Count & "files"
50190  IfLoggingWriteLogfile tStr
50200  WriteToSpecialLogfile tStr
50210  Do While tColl.Count > 0
50220   For i = 1 To tColl.Count
50230    tFile = Split(tColl.Item(i), "|")
50240    If FileExists(tFile(1)) And Not FileInUse(tFile(1)) Then
50250     OutputFilename = GetAutosaveFilename(tFile(1))
50260     SplitPath OutputFilename, , Pathname
50270     If IsValidPath(Pathname) = True Then
50280       If DirExists(Pathname) = False Then
50290        MakePath (Pathname)
50300       End If
50310       tStr = "Autosavemodus: Create File '" & OutputFilename & "'"
50320       IfLoggingWriteLogfile tStr
50330       WriteToSpecialLogfile tStr
50340       PSHeader = GetPSHeader(tFile(1))
50350       tDate = Now
50360       With PDFDocInfo
50370        If Options.UseStandardAuthor = 1 Then
50380          .Author = GetSubstFilename(tFile(1), RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardAuthor)), True)
50390         Else
50400          .Author = GetDocUsername(tFile(1), False)
50410        End If
50420        If LenB(PSHeader.CreationDate.Comment) > 0 Then
50430          tStr = PSHeader.CreationDate.Comment
50440         Else
50450          tStr = CStr(tDate)
50460        End If
50470        .CreationDate = GetDocDate(Trim$(Options.StandardCreationdate), Options.StandardDateformat, FormatPrintDocumentDate(tStr))
50480        .Creator = App.EXEName & " Version " & App.Major & "." & App.Minor & "." & App.Revision
50490        .Keywords = GetSubstFilename(tFile(1), RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardKeywords)), , , True)
50500        'tStr = CStr(tDate)
50510        .ModifyDate = GetDocDate(Trim$(Options.StandardModifydate), Options.StandardDateformat, FormatPrintDocumentDate(tStr))
50520        .Subject = GetSubstFilename(tFile(1), RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardSubject)), , , True)
50530        If Len(Options.StandardTitle) > 0 Then
50540          .Title = GetSubstFilename(tFile(1), RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardTitle)), , , True)
50550         Else
50560          .Title = GetSubstFilename(tFile(1), Options.SaveFilename, , , True)
50570        End If
50580       End With
50590       AppendPDFDocInfo tFile(1), PDFDocInfo
50600       CheckForStamping tFile(1)
50610       If Options.RunProgramBeforeSaving = 1 Then
50620        RunProgramBeforeSaving Me.hwnd, tFile(1), _
        Options.RunProgramBeforeSavingProgramParameters, _
        Options.RunProgramBeforeSavingWindowstyle
50650       End If
50660       CallGScript tFile(1), OutputFilename, Options, Options.AutosaveFormat
50670       If FileExists(OutputFilename) = True Then
50680         IsConverted = True
50690         tStr = "Autosavemodus: Create File '" & OutputFilename & "' success"
50700         IfLoggingWriteLogfile tStr
50710         WriteToSpecialLogfile tStr
50720         If Options.RunProgramAfterSaving = 1 Then
50730          RunProgramAfterSaving Me.hwnd, OutputFilename, _
         Options.RunProgramAfterSavingProgramParameters, _
         Options.RunProgramAfterSavingWindowstyle, tFile(1)
50760         End If
50770         If Options.SendEmailAfterAutoSaving = 1 Then
50780          Set mail = New clsPDFCreatorMail
50790          If mail.Send(OutputFilename, PDFDocInfo.Subject, Options.SendMailMethod) <> 0 Then
50800           MsgBox LanguageStrings.MessagesMsg04, vbCritical, App.EXEName
50810          End If
50820          Set mail = Nothing
50830         End If
50840         If Options.AutosaveStartStandardProgram = 1 Then
50850          If Options.OnePagePerFile = 1 Then
50860            OpenDocument Replace$(OutputFilename, "%d", "1", , , vbTextCompare)
50870           Else
50880            OpenDocument OutputFilename
50890          End If
50900         End If
50910        Else
50920         tStr = "Autosavemodus: Create File '" & OutputFilename & "' failed"
50930         IfLoggingWriteLogfile tStr
50940         WriteToSpecialLogfile tStr
50950       End If
50960      Else
50970       IfLoggingWriteLogfile "Error: Invalid autosave pathname, spoolfile will be deleted!"
50980     End If
50990     CheckForPrintingAfterSaving tFile(1), Options
51000     KillFile tFile(1)
51010     KillInfoSpoolfile tFile(1)
51020     ConvertedOutputFilename = OutputFilename
51030     ReadyConverting = True
51040    End If
51050   Next i
51060   Set tColl = GetFiles(GetPDFCreatorTempfolder, "~PS*.tmp", SortedByDate)
51070  Loop
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
50020  h1 = GetMenu(Me.hwnd): h2 = GetSubMenu(h1, 4)
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
Select Case ErrPtnr.OnError("frmMain", "IsCompatibleLanguageVersion")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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
50110   .Buttons.Add , , , , 7
50120   .Buttons.Add , , , , 8
50130   .Buttons.Add , , , , 9
50140   .Buttons.Add , , , , 10
50150   .Buttons.Add , , , , 11
50160   .Buttons.Add , , , , 12
50170   .Buttons.Add , , , , 13
50180   .Buttons.Add , , , tbrSeparator
50190   .Buttons.Add , , , , 14
50200  End With
50210  With tlb(1)
50220   Set .ImageList = imlTlb
50230   .Buttons.Clear
50240   .Buttons.Add , , , , 15
50250   .Buttons.Add , , , tbrSeparator
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
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If (Options.Toolbars And 2) <> 2 Then
50020   txtEmailAddress.Enabled = False
50030   txtEmailAddress.Visible = False
50040   mnDocument(14).Enabled = False
50050   mnDocument(15).Enabled = False
50060   mnDocument(16).Enabled = False
50070   mnDocument(13).Visible = False
50080   mnDocument(14).Visible = False
50090   mnDocument(15).Visible = False
50100   mnDocument(16).Visible = False
50110  End If
50120  If (Options.Toolbars And 1) = 1 Then
50130    tlb(0).Visible = True
50140    mnViewToolbars(0).Checked = True
50150   Else
50160    tlb(0).Visible = False
50170    mnViewToolbars(0).Checked = False
50180  End If
50190  If (Options.Toolbars And 2) = 2 Then
50200    tlb(1).Visible = True
50210    txtEmailAddress.Visible = True
50220    mnViewToolbars(1).Checked = True
50230   Else
50240    tlb(1).Visible = False
50250    txtEmailAddress.Visible = False
50260    mnViewToolbars(1).Checked = False
50270  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "DrawToolbars")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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
50140 ' mnPrinter(0).Checked = False
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

Private Sub DocumentAdd()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, Cancel As Boolean, cFiles As Collection, aLen As Double, _
  OnlyPsFiles As Boolean, Ext As String, DefaultPrintername As String, _
  tFilename As String, tLen As Double
50040  Set cFiles = GetFilename("", GetMyFiles, 0, _
  LanguageStrings.ListPostscriptFiles & " (*.ps)|*.ps|" & LanguageStrings.ListAllFiles & " (*.*)|*.*", _
   OpenFile, Cancel, Me)
50070  If Cancel = True Then
50080   Screen.MousePointer = vbNormal
50090   Exit Sub
50100  End If
50110  aLen = 0
50120  For i = 1 To cFiles.Count
50130   aLen = aLen + FileLen(cFiles.Item(i))
50140  Next i
50150  OnlyPsFiles = True
50160  For i = 1 To cFiles.Count
50170   SplitPath cFiles.Item(i), , , , , Ext
50180   If UCase$(Ext) <> "PS" And UCase$(Ext) <> "EPS" Then
50190    OnlyPsFiles = False
50200    Exit For
50210   End If
50220  Next i
50230  If UCase$(Printer.DeviceName) <> UCase$(GetPDFCreatorPrintername) And OnlyPsFiles = False Then
50240   If Options.NoConfirmMessageSwitchingDefaultprinter = 0 Then
50250    If ChangeDefaultprinter = False Then
50260     frmSwitchDefaultprinter.Show vbModal, Me
50270     If ChangeDefaultprinter = False Then
50280      Screen.MousePointer = vbNormal
50290      Exit Sub
50300     End If
50310    End If
50320   End If
50330  End If
50340  ChangeDefaultprinter = True
50350  DefaultPrintername = Printer.DeviceName
50360  SetDefaultprinterInProg GetPDFCreatorPrintername
50370  aLen = 0
50380  For i = 1 To cFiles.Count
50390   aLen = aLen + FileLen(cFiles.Item(i))
50400  Next i
50410  For i = 1 To cFiles.Count
50420   SplitPath cFiles.Item(i), , , , , Ext
50430   If IsPostscriptFile(cFiles.Item(i)) = True And (UCase$(Ext) = "PS" Or UCase$(Ext) = "EPS") Then
50440     tFilename = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory, "~PS")
50450     DoEvents
50460     FileCopy cFiles.Item(i), tFilename
50470    Else
50480     DoEvents
50490     ShellAndWait Me.hwnd, "print", cFiles.Item(i), "", vbNullChar, wHidden, WCTermination, 60000, True
50500     DoEvents
50510   End If
50520   tLen = tLen + FileLen(cFiles.Item(i))
50530   stb.Panels("Percent").Text = Format$(tLen / aLen, " 0.0%")
50540   DoEvents
50550  Next i
50560  If DefaultPrintername <> vbNullString Then
50570   SetDefaultprinterInProg DefaultPrintername
50580  End If
50590  stb.Panels("Percent").Text = vbNullString
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
50020  For i = 1 To lsv.ListItems.Count
50030   If lsv.ListItems(i).Selected = True Then
50040    KillFile lsv.ListItems(i).SubItems(4)
50050    KillInfoSpoolfile lsv.ListItems(i).SubItems(4)
50060   End If
50070   DoEvents
50080  Next i
50090  LvwRemoveSelectedItems lsv, True
50100  If lsv.ListItems.Count > 0 Then
50110   If lsv.SelectedItem.Index > lsv.ListItems.Count Then
50120     lsv.ListItems(lsv.SelectedItem.Index - 1).Selected = True
50130    Else
50140     lsv.ListItems(lsv.SelectedItem.Index).Selected = True
50150   End If
50160  End If
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

Public Sub DocumentCombine()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, cFiles As Collection, tFilename As String, tFilename2 As String
50020  Screen.MousePointer = vbHourglass
50030  LockWindowUpdate lsv.hwnd
50040  Set cFiles = New Collection
50050  For i = 1 To lsv.ListItems.Count
50060   If lsv.ListItems(i).Selected = True Then
50070    cFiles.Add lsv.ListItems(i).SubItems(4)
50080   End If
50090  Next i
50100  tFilename = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory, "~PS")
50110  KillFile tFilename
50120  If cFiles.Count > 1 Then
50130   CombineFiles tFilename, cFiles, , stb
50140   stb.Panels("Percent").Text = ""
50150  End If
50160  Set cFiles = Nothing
50170  LockWindowUpdate 0&
50180  Screen.MousePointer = vbNormal
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "DocumentCombine")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

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
50010  If mnPrinter(0).Checked = False Or NoProcessing = True Then
50020    SetPrinterStop True
50030    mnPrinter(0).Checked = True
50040    tlb(0).Buttons(1).Image = 2
50050   Else
50060    SetPrinterStop False
50070    mnPrinter(0).Checked = False
50080    tlb(0).Buttons(1).Image = 1
50090  End If
50100  If Not m_frmSysTray Is Nothing Then
50110   If mnPrinter(0).Checked = True Then
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
50070   .Buttons(7).ToolTipText = LanguageStrings.DialogDocumentDelete
50080   .Buttons(8).ToolTipText = LanguageStrings.DialogDocumentTop
50090   .Buttons(9).ToolTipText = LanguageStrings.DialogDocumentUp
50100   .Buttons(10).ToolTipText = LanguageStrings.DialogDocumentDown
50110   .Buttons(11).ToolTipText = LanguageStrings.DialogDocumentBottom
50120   .Buttons(12).ToolTipText = LanguageStrings.DialogDocumentCombine
50130   .Buttons(13).ToolTipText = LanguageStrings.DialogDocumentSave
50140   .Buttons(15).ToolTipText = "?"
50150  End With
50160  With tlb(1)
50170   .Buttons(1).ToolTipText = LanguageStrings.DialogDocumentCombineAll
50180   .Buttons(3).ToolTipText = LanguageStrings.DialogDocumentCombineAllSend
50190   .Buttons(4).ToolTipText = LanguageStrings.DialogDocumentSend
50200  End With
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

Public Function CombineAll() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, cFiles As Collection, tFilename As String, tFilename2 As String
50020  If lsv.ListItems.Count = 0 Then
50030   Exit Function
50040  End If
50050  If lsv.ListItems.Count = 1 Then
50060   CombineAll = lsv.ListItems(1).SubItems(4)
50070   Exit Function
50080  End If
50090  Screen.MousePointer = vbHourglass
50100  LockWindowUpdate hwnd
50110  Timer1.Enabled = False
50120  Set cFiles = New Collection
50130  For i = 1 To lsv.ListItems.Count
50140   cFiles.Add lsv.ListItems(i).SubItems(4)
50150  Next i
50160  tFilename = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory, "~PT")
50170  KillFile tFilename
50180  If cFiles.Count > 1 Then
50190   CombineFiles tFilename, cFiles, , stb
50200   stb.Panels("Percent").Text = ""
50210  End If
50220  tFilename2 = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory, "~PS")
50230  KillFile tFilename2
50240  Name tFilename As tFilename2
50250  Set cFiles = Nothing
50260  CombineAll = tFilename2
50270  Timer1.Enabled = True
50280  LockWindowUpdate 0&
50290  Screen.MousePointer = vbNormal
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "CombineAll")
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
50110     RunProgramBeforeSaving Me.hwnd, GetShortName(InputFilename), _
    Options.RunProgramBeforeSavingProgramParameters, _
    Options.RunProgramBeforeSavingWindowstyle
50140    End If
50150    ConvertPostscriptFile InputFilename, OutputFilename
50160    If Len(OutputFilename) > 0 And FileExists(OutputFilename) = True Then
50170     If Options.RunProgramAfterSaving = 1 Then
50180      RunProgramAfterSaving Me.hwnd, OutputFilename, _
      Options.RunProgramAfterSavingProgramParameters, _
      Options.RunProgramAfterSavingWindowstyle, InputFilename
50210     End If
50220     Set mail = New clsPDFCreatorMail
50230     If mail.Send(OutputFilename, Options.StandardSubject, Options.SendMailMethod, rec) <> 0 Then
50240      MsgBox LanguageStrings.MessagesMsg04, vbCritical, App.EXEName
50250     End If
50260     Set mail = Nothing
50270    End If
50280    KillFile InputFilename
50290   End If
50300  End If
50310  Timer1.Enabled = True
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
50010  Timer1.Enabled = False
50020  If Options.ShowAnimation = 1 Then
50030   ShowAnimationWindow = True
50040   frmAnimation.Show
50050  End If
50060  SendEmailImmediately CombineAll
50070  If Options.ShowAnimation = 1 Then
50080   ShowAnimationWindow = False
50090  End If
50100  Timer1.Enabled = True
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
50040     Timer1.Enabled = False
50050     If Options.ShowAnimation = 1 Then
50060      ShowAnimationWindow = True
50070      frmAnimation.Show
50080     End If
50090     SendEmailImmediately lsv.SelectedItem.ListSubItems(4)
50100     If Options.ShowAnimation = 1 Then
50110      ShowAnimationWindow = False
50120      Timer1.Enabled = True
50130     End If
50140    End If
50150   End If
50160  End If
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
50100    If mnPrinter(4).Checked = True Then
50110     .mnuSysTray(6).Checked = True
50120    End If
50130    If mnPrinter(0).Checked = True Then
50140     .mnuSysTray(2).Checked = True
50150    End If
50160    .ToolTip = Me.Caption
50170   End With
50180   If mnPrinter(0).Checked = True Then
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
