VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "PDFCreator"
   ClientHeight    =   2970
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9510
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows-Standard
   Begin MSComctlLib.ImageList imlTlb 
      Left            =   3360
      Top             =   2100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
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
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6558
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":66B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6C4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":71E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7780
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7B1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7EB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":824E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":85E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8982
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8D1C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Oben ausrichten
      Height          =   630
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   840
      Top             =   2160
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   2160
   End
   Begin MSComctlLib.StatusBar stb 
      Align           =   2  'Unten ausrichten
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   2700
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
      Top             =   735
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
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin VB.Image imgPaypal 
      Height          =   465
      Left            =   2070
      Picture         =   "frmMain.frx":92B6
      Top             =   2025
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
50040  Me.KeyPreview = True
50050
50060  ReadAllLanguages LanguagePath
50070  InitProgram
50080
50090  ShowPaypalMenuimage
50100  SetGSRevision
50110
50120  If NoProcessing = True Or Options.NoProcessingAtStartup = 1 Then
50130   SetMenuPrinterStop
50140  End If
50150  If PrinterStop = True Then
50160   mnPrinter(0).Checked = True
50170  End If
50180  If Options.OptionsEnabled = 0 Then
50190   mnPrinter(2).Enabled = False
50200   tlb(0).Buttons(2).Enabled = False
50210  End If
50220  If Options.OptionsVisible = 0 Then
50230   tlb(0).Buttons(2).Visible = False
50240   mnPrinter(2).Visible = False
50250   mnPrinter(3).Visible = False
50260  End If
50270
50280  CheckPrintJobs
50290
50300  SetTopMost frmMain, True, True
50310  SetTopMost frmMain, False, True
50320  SetActiveWindow frmMain.hwnd
50330
50340  ' Only for the first time set interval to 10 ms
50350  Timer1.Interval = 10
50360  Timer1.Enabled = True
50370  Timer2.Interval = 100
50380  Timer2.Enabled = True
50390 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50400 Exit Sub
ErrPtnr_OnError:
50421 Select Case ErrPtnr.OnError("frmMain", "Form_Load")
      Case 0: Resume
50440 Case 1: Resume Next
50450 Case 2: Exit Sub
50460 Case 3: End
50470 End Select
50480 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub Form_Resize()
 On Error Resume Next
 If Me.WindowState = vbMinimized Then
  Exit Sub
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
  .Top = tlb(0).Height: .Left = 0
  .Width = Me.Width - 125
  .Height = Me.ScaleHeight - Abs(stb.Visible) * stb.Height - tlb(0).Height
 End With
 stb.Panels("Status").Width = Me.Width - 150 - stb.Panels("Percent").Width _
  - stb.Panels("GhostscriptRevision").Width
 If PDFCreatorPrinter Or (LenB(InputFilename) > 0 And IFIsPS = True) Then
  If lsv.ListItems.Count <= 1 And PrinterStop = False Then
   Me.Visible = False
  End If
 End If
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
50370  CheckPrintJobs
50380  DoEvents
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
50010  Dim PDFSpoolerPath As String
50020  Timer1.Enabled = False
50030  Timer2.Enabled = False
50040  Set Printjobs = Nothing
50050  If Not mutex Is Nothing Then
50060   mutex.CloseMutex
50070   Set mutex = Nothing
50080  End If
50090  IfLoggingWriteLogfile "PDFCreator Program End"
50100  UnloadDLLComplete GsDllLoaded
50110  PDFSpoolerPath = CompletePath(GetSystemDirectory) & "PDFSpooler.exe"
50120  If Restart = True And FileExists(PDFSpoolerPath) = True Then
50130   ShellExecute 0, vbNullString, """" & PDFSpoolerPath & """", "-SL200 -STTRUE", App.Path, 1
50140  End If
50150  End
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
50010  Dim coll As Collection, i As Long, tStrf() As String
50020  Set GetAllLanguagesFiles = New Collection
50030  Set coll = GetFiles(LanguagePath, "*.ini")
50040  For i = 1 To coll.Count
50050   tStrf = Split(coll(i), "|")
50060   GetAllLanguagesFiles.Add tStrf(1)
50070  Next i
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
50080   ini.Filename = LangFiles.item(i)
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
50210   mnLanguage(mnLanguage.Count - 1).Tag = LangFiles.item(i)
50220   SplitPath LangFiles.item(i), , , , Filename
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
50010  Dim i As Long, Version As String, reg As clsRegistry
50020
50030  For i = mnLanguage.LBound To mnLanguage.UBound
50040   If UCase$(Languagefile) = UCase$(mnLanguage.item(i).Tag) Then
50050     mnLanguage.item(i).Checked = True
50060    Else
50070     mnLanguage.item(i).Checked = False
50080   End If
50090  Next i
50100
50110  With LanguageStrings
50120   Set reg = New clsRegistry
50130   With reg
50140    .hkey = HKEY_LOCAL_MACHINE
50150    .KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50160    Version = .GetRegistryValue("ApplicationVersion")
50170   End With
50180   Set reg = Nothing
50190   Caption = App.Title & " " & GetProgramReleaseStr & " " & .CommonTitle
50200
50210   mnPrinterMain.Caption = .DialogPrinter
50220   mnPrinter(0).Caption = .DialogPrinterPrinterStop
50230   mnPrinter(2).Caption = .DialogPrinterOptions
50240   mnPrinter(4).Caption = .DialogPrinterLogging
50250   mnPrinter(5).Caption = .DialogPrinterLogfile
50260   mnPrinter(7).Caption = .DialogPrinterClose
50270
50280   mnDocumentMain.Caption = .DialogDocument
50290   mnDocument(0).Caption = .DialogDocumentPrint
50300   mnDocument(2).Caption = .DialogDocumentAdd
50310   mnDocument(3).Caption = .DialogDocumentDelete
50320   mnDocument(5).Caption = .DialogDocumentTop
50330   mnDocument(6).Caption = .DialogDocumentUp
50340   mnDocument(7).Caption = .DialogDocumentDown
50350   mnDocument(8).Caption = .DialogDocumentBottom
50360   mnDocument(10).Caption = .DialogDocumentCombine
50370   mnDocument(12).Caption = .DialogDocumentSave
50380
50390   mnViewMain.Caption = .DialogView
50400   mnView(0).Caption = .DialogViewStatusbar
50410
50420   mnLanguageMain.Caption = .DialogLanguage
50430
50440   mnHelpMain.Caption = .DialogInfo
50450   mnHelp(2).Caption = .DialogInfoPaypal
50460   mnHelp(4).Caption = .DialogInfoHomepage
50470   mnHelp(5).Caption = .DialogInfoPDFCreatorSourceforge
50480   mnHelp(6).Caption = .DialogInfoCheckUpdates
50490   mnHelp(8).Caption = .DialogInfoInfo
50500
50510   lsv.ColumnHeaders("Date").Text = .ListDate
50520   lsv.ColumnHeaders("Documenttitle").Text = .ListDocumenttitle
50530   lsv.ColumnHeaders("Filename").Text = .ListFilename
50540   lsv.ColumnHeaders("Size").Text = .ListSize
50550   lsv.ColumnHeaders("Status").Text = .ListStatus
50560  End With
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

Private Sub lsv_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
 On Error Resume Next
 SetDocumentMenu
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
    If IsPostscriptFile(data.Files.item(1)) = True Then
      tFilename = GetTempFile(GetPDFCreatorTempfolder & PDFCreatorSpoolDirectory & "\" & GetUsername, "~PD")
      FileCopy data.Files.item(1), tFilename
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
      ShellAndWait Me.hwnd, "print", data.Files.item(1), "", vbNullChar, wHidden, WCTermination, 60000, True
      If DefaultPrintername <> vbNullString Then
       SetDefaultprinterInProg DefaultPrintername
      End If
    End If
    DoEvents
   Else
    OnlyPsFiles = True
    For i = 1 To data.Files.Count
     SplitPath data.Files.item(i), , , , , Ext
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
     aLen = aLen + FileLen(data.Files.item(i))
    Next i
    For i = 1 To data.Files.Count
     If IsPostscriptFile(data.Files.item(i)) = True Then
       tFilename = GetTempFile(GetPDFCreatorTempfolder, "~PD")
       DoEvents
       FileCopy data.Files.item(i), tFilename
      Else
       DoEvents
'       PrintDocument data.Files.item(i)
       ShellAndWait Me.hwnd, "print", data.Files.item(i), "", vbNullChar, wHidden, WCTermination, 60000, True
       DoEvents
     End If
     tLen = tLen + FileLen(data.Files.item(i))
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
   If IsFilePrintable(data.Files.item(i)) = False And IsPostscriptFile(data.Files.item(i)) = False Then
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
50280  End Select
50290  SetDocumentMenu
50300  Screen.MousePointer = vbNormal
50310  Timer1.Enabled = True
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
50010  SetDocumentMenu
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
50340    frmInfo.Show vbModal, Me
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
50140   Case 5:
50150    frmLog.Show , Me
50160   Case 7:
50170    End
50180  End Select
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
        Case 0:
50030    stb.Visible = Not stb.Visible
50040    mnView(0).Checked = Not mnView(0).Checked
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

Private Sub Timer1_Timer()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Timer1.Enabled = False
50020  DoEvents
50030  If FileExists(CompletePath(App.Path) & "Unload.tmp") = True Or Restart = True Then
50040   Unload Me
50050  End If
50060  CheckPrintJobs
50070  If Not NoProcessing Then
50080   CheckForPrinting
50090  End If
50100  If lsv.ListItems.Count = 0 And LenB(CommandSwitch("IF", True)) > 0 Then
50110   Unload Me
50120  End If
50130  If lsv.ListItems.Count = 1 Then
50140   lsv.ListItems(1).Selected = True
50150  End If
50160  DoEvents
50170  Timer1.Interval = TimerIntervall
50180  Timer1.Enabled = True
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

Private Sub CheckForPrinting()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If lsv.ListItems.Count > 0 Then
50020   If mnPrinter(0).Checked = True Then
50030     If PrintSelectedJobs = True Then
50040       If lsv.ListItems(1).SubItems(1) = LanguageStrings.ListPrinting Then
50050         PDFSpoolfile = lsv.ListItems(1).SubItems(4)
50060         frmPrinting.Show vbModal, Me
50070        Else
50080         PrintSelectedJobs = False
50090       End If
50100      Else
50110       lsv.ListItems(1).SubItems(1) = LanguageStrings.ListWaiting
50120     End If
50130    Else
50140     lsv.ListItems(1).SubItems(1) = LanguageStrings.ListPrinting
50150     PDFSpoolfile = lsv.ListItems(1).SubItems(4)
50160     If PrinterStop = False Then
50170      If IsFormLoaded(frmPrinting) = False Then
50180       If InstalledAsServer Then
50190        Options = ReadOptions
50200       End If
50210       If Options.UseAutosave = 1 Then
50220         Autosave
50230        Else
50240         frmPrinting.Show , Me
50250       End If
50260 '      If Me.Visible = True Then
50270 '       Me.Show
50280 '      End If
50290      End If
50300     End If
50310     If PrinterStop = False And NoProcessing = False Then
50320       mnPrinter(0).Checked = False
50330       tlb(0).Buttons(1).Image = 1
50340      Else
50350       mnPrinter(0).Checked = True
50360       tlb(0).Buttons(1).Image = 2
50370     End If
50380   End If
50390  End If
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
50010  Dim Temppath As String, LItem As ListItem, tColl As Collection, _
  tFile() As String, i As Long, j As Long, kB As Long, MB As Long, GB As Long
50030  kB = 1024: MB = kB * 1024: GB = MB * 1024
50040  Set tColl = New Collection
50050  Temppath = CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory & "\"
50060
50070  'Set tColl = GetFiles(GetPDFCreatorTempfolder, "~PD*.tmp")
50080  FindFiles Temppath, tColl, "~PD*.tmp", , True
50090  If tColl.Count = 0 And lsv.ListItems.Count > 0 Then
50100   lsv.ListItems.Clear
50110  End If
50120  For i = 1 To tColl.Count
50130   tFile = Split(tColl.item(i), "|")
50140   For j = 1 To lsv.ListItems.Count
50150    If UCase$(tFile(1)) = UCase$(lsv.ListItems(j).SubItems(4)) Then
50160     Exit For
50170    End If
50180   Next j
50190   If j > lsv.ListItems.Count Then
50200     SetActiveWindow Me.hwnd
50210     Set LItem = lsv.ListItems.Add(, , GetPDFTitle(tFile(1)))
50220     LItem.SubItems(1) = LanguageStrings.ListWaiting
50230     LItem.SubItems(2) = tFile(3)
50240     If CLng(tFile(2)) > GB Then
50250       LItem.SubItems(3) = Format$(CDbl(tFile(2)) / GB, "0.00 " & LanguageStrings.ListGBytes)
50260      Else
50270       If CLng(tFile(2)) > MB Then
50280         LItem.SubItems(3) = Format$(CDbl(tFile(2)) / MB, "0.00 " & LanguageStrings.ListMBytes)
50290        Else
50300         If CLng(tFile(2)) > kB Then
50310           LItem.SubItems(3) = Format$(CDbl(tFile(2)) / kB, "0.00 " & LanguageStrings.ListKBytes)
50320          Else
50330           LItem.SubItems(3) = Format$(tFile(2), "0 " & LanguageStrings.ListBytes)
50340         End If
50350      End If
50360     End If
50370     LItem.SubItems(4) = tFile(1)
50380     DoEvents
50390    Else
50400 '
50410   End If
50420  Next i
50430  i = 0
50440  Do Until i + 1 >= lsv.ListItems.Count
50450   i = i + 1
50460   For j = 1 To tColl.Count
50470    tFile = Split(tColl.item(j), "|")
50480    If UCase$(tFile(1)) = UCase$(lsv.ListItems(i).SubItems(4)) Then
50490     Exit For
50500    End If
50510   Next j
50520   If j > tColl.Count Then
50530    lsv.ListItems.Remove i
50540   End If
50550   DoEvents
50560  Loop
50570  If lsv.ListItems.Count = 1 Then
50580    stb.Panels("Status").Text = LanguageStrings.ListStatus & ": " & lsv.ListItems.Count & " " & LanguageStrings.MessagesMsg01
50590   Else
50600    stb.Panels("Status").Text = LanguageStrings.ListStatus & ": " & lsv.ListItems.Count & " " & LanguageStrings.MessagesMsg02
50610  End If
50620  Set tColl = Nothing
50630  SetDocumentMenu
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

Private Sub SetDocumentMenu()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim c As Long
50020  SetToolbar
50030  If lsv.ListItems.Count = 0 Then
50040    mnDocument(0).Enabled = False
50050    mnDocument(3).Enabled = False
50060    mnDocument(5).Enabled = False
50070    mnDocument(6).Enabled = False
50080    mnDocument(7).Enabled = False
50090    mnDocument(8).Enabled = False
50100    mnDocument(10).Enabled = False
50110    mnDocument(12).Enabled = False
50120    Exit Sub
50130   Else
50140    If lsv.ListItems.Count = 1 Then
50150     mnDocument(0).Enabled = True
50160     mnDocument(3).Enabled = True
50170     mnDocument(5).Enabled = False
50180     mnDocument(6).Enabled = False
50190     mnDocument(7).Enabled = False
50200     mnDocument(8).Enabled = False
50210     mnDocument(10).Enabled = False
50220     mnDocument(12).Enabled = True
50230     Exit Sub
50240    End If
50250  End If
50260  mnDocument(0).Enabled = True
50270  mnDocument(3).Enabled = True
50280  mnDocument(5).Enabled = True
50290  mnDocument(6).Enabled = True
50300  mnDocument(7).Enabled = True
50310  mnDocument(8).Enabled = True
50320  mnDocument(10).Enabled = False
50330  mnDocument(12).Enabled = True
50340  c = LvwGetCountSelectedItems(lsv, True)
50350  If c > 1 Then
50360   mnDocument(10).Enabled = True
50370   mnDocument(12).Enabled = False
50380  End If
50390  If c = lsv.ListItems.Count Then
50400  End If
50410  If lsv.SelectedItem.Index = 1 And c <= 1 Then
50420   mnDocument(5).Enabled = False
50430   mnDocument(6).Enabled = False
50440   mnDocument(7).Enabled = True
50450   mnDocument(8).Enabled = True
50460  End If
50470  If lsv.SelectedItem.Index = lsv.ListItems.Count And c <= 1 Then
50480   mnDocument(5).Enabled = True
50490   mnDocument(6).Enabled = True
50500   mnDocument(7).Enabled = False
50510   mnDocument(8).Enabled = False
50520  End If
50530  If c = lsv.ListItems.Count Then
50540   mnDocument(5).Enabled = False
50550   mnDocument(6).Enabled = False
50560   mnDocument(7).Enabled = False
50570   mnDocument(8).Enabled = False
50580  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "SetDocumentMenu")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetToolbar()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim c As Long
50020  With tlb(0)
50030   If lsv.ListItems.Count = 0 Then
50040     .Buttons(5).Enabled = False
50050     .Buttons(7).Enabled = False
50060     .Buttons(8).Enabled = False
50070     .Buttons(9).Enabled = False
50080     .Buttons(10).Enabled = False
50090     .Buttons(11).Enabled = False
50100     .Buttons(12).Enabled = False
50110     .Buttons(13).Enabled = False
50120     Exit Sub
50130    Else
50140     If lsv.ListItems.Count = 1 Then
50150      .Buttons(5).Enabled = True
50160      .Buttons(7).Enabled = True
50170      .Buttons(8).Enabled = False
50180      .Buttons(9).Enabled = False
50190      .Buttons(10).Enabled = False
50200      .Buttons(11).Enabled = False
50210      .Buttons(12).Enabled = False
50220      .Buttons(13).Enabled = True
50230      Exit Sub
50240     End If
50250   End If
50260   .Buttons(5).Enabled = True
50270   .Buttons(7).Enabled = True
50280   .Buttons(8).Enabled = True
50290   .Buttons(9).Enabled = True
50300   .Buttons(10).Enabled = True
50310   .Buttons(11).Enabled = True
50320   .Buttons(12).Enabled = False
50330   .Buttons(13).Enabled = True
50340   c = LvwGetCountSelectedItems(lsv, True)
50350   If c > 1 Then
50360    .Buttons(12).Enabled = True
50370    .Buttons(13).Enabled = False
50380   End If
50390   If lsv.SelectedItem.Index = 1 And c <= 1 Then
50400    .Buttons(8).Enabled = False
50410    .Buttons(9).Enabled = False
50420    .Buttons(10).Enabled = True
50430    .Buttons(11).Enabled = True
50440   End If
50450   If lsv.SelectedItem.Index = lsv.ListItems.Count And c <= 1 Then
50460    .Buttons(8).Enabled = True
50470    .Buttons(9).Enabled = True
50480    .Buttons(10).Enabled = False
50490    .Buttons(11).Enabled = False
50500   End If
50510   If c = lsv.ListItems.Count Then
50520    .Buttons(8).Enabled = False
50530    .Buttons(9).Enabled = False
50540    .Buttons(10).Enabled = False
50550    .Buttons(11).Enabled = False
50560   End If
50570  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "SetToolbar")
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
  PSHeader As tPSHeader, tDate As Date
50040
50050  Set tColl = New Collection
50060
50070  If Len(Filename) > 0 Then
50080    If FileExists(Filename) = True Then
50090     SplitPath Filename, , Pathname
50100     tColl.Add Pathname & "|" & Filename & "|" & FileLen(Filename) & "|" & FileDateTime(Filename)
50110    End If
50120   Else
50130    FindFiles CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory, tColl, "~P*.tmp", , True
50140 '   Set tColl = GetFiles(GetPDFCreatorTempfolder, "~P*.tmp")
50150  End If
50160
50170  tStr = "Autosavemodus: " & tColl.Count & "files"
50180  IfLoggingWriteLogfile tStr
50190  WriteToSpecialLogfile tStr
50200  Do While tColl.Count > 0
50210   For i = 1 To tColl.Count
50220    tFile = Split(tColl.item(i), "|")
50230    OutputFilename = GetAutosaveFilename(tFile(1))
50240    SplitPath OutputFilename, , Pathname
50250    If IsValidPath(Pathname) = True Then
50260      If DirExists(Pathname) = False Then
50270       MakePath (Pathname)
50280      End If
50290      tStr = "Autosavemodus: Create File '" & OutputFilename & "'"
50300      IfLoggingWriteLogfile tStr
50310      WriteToSpecialLogfile tStr
50320      PSHeader = GetPSHeader(tFile(1))
50330      tDate = Now
50340      With PDFDocInfo
50350       If Options.UseStandardAuthor = 1 Then
50360         .Author = GetSubstFilename(tFile(1), RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardAuthor)), True)
50370        Else
50380         .Author = GetDocUsername(tFile(1), False)
50390       End If
50400       If LenB(PSHeader.CreationDate.Comment) > 0 Then
50410         tStr = PSHeader.CreationDate.Comment
50420        Else
50430         tStr = CStr(tDate)
50440       End If
50450       .CreationDate = GetDocDate(Trim$(Options.StandardCreationdate), Options.StandardDateformat, FormatPrintDocumentDate(tStr))
50460       .Creator = App.EXEName & " Version " & App.Major & "." & App.Minor & "." & App.Revision
50470       .Keywords = GetSubstFilename(tFile(1), RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardKeywords)))
50480       'tStr = CStr(tDate)
50490       .ModifyDate = GetDocDate(Trim$(Options.StandardModifydate), Options.StandardDateformat, FormatPrintDocumentDate(tStr))
50500       .Subject = GetSubstFilename(tFile(1), RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardSubject)))
50510       If Len(Options.StandardTitle) > 0 Then
50520         .Title = GetSubstFilename(tFile(1), RemoveLeadingAndTrailingQuotes(Trim$(Options.StandardTitle)))
50530        Else
50540         .Title = GetSubstFilename(tFile(1), Options.SaveFilename)
50550       End If
50560      End With
50570      AppendPDFDocInfo tFile(1), PDFDocInfo
50580      CheckForStamping tFile(1)
50590      CallGScript tFile(1), OutputFilename, Options, Options.AutosaveFormat
50600      If FileExists(OutputFilename) = True Then
50610        tStr = "Autosavemodus: Create File '" & OutputFilename & "' success"
50620        IfLoggingWriteLogfile tStr
50630        WriteToSpecialLogfile tStr
50640        If Options.RunProgramAfterSaving = 1 Then
50650         RunProgramAfterSaving Me.hwnd, GetShortName(OutputFilename), GetRunProgramAfterSavingProgramParameters, Options.RunProgramAfterSavingWindowstyle
50660        End If
50670       Else
50680        tStr = "Autosavemodus: Create File '" & OutputFilename & "' failed"
50690        IfLoggingWriteLogfile tStr
50700        WriteToSpecialLogfile tStr
50710      End If
50720     Else
50730      IfLoggingWriteLogfile "Error: Invalid autosave pathname, spoolfile will be deleted!"
50740    End If
50750    KillFile tFile(1)
50760   Next i
50770   Set tColl = GetFiles(GetPDFCreatorTempfolder, "~P*.tmp")
50780  Loop
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
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim mutex As clsMutex
50020  Set mutex = New clsMutex
50030  If mutex.CheckMutex(PDFCreator_GUID) = False Then
50040  ' Create a mutex
50050    mutex.CreateMutex PDFCreator_GUID
50060  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "Timer2_Timer")
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
50010  With tlb(0)
50020   Set .ImageList = imlTlb
50030   .Buttons.Clear
50040   .Buttons.Add , , , , 1
50050   .Buttons.Add , , , , 3
50060   .Buttons.Add , , , , 4
50070   .Buttons.Add , , , tbrSeparator
50080   .Buttons.Add , , , , 5
50090   .Buttons.Add , , , , 6
50100   .Buttons.Add , , , , 7
50110   .Buttons.Add , , , , 8
50120   .Buttons.Add , , , , 9
50130   .Buttons.Add , , , , 10
50140   .Buttons.Add , , , , 11
50150   .Buttons.Add , , , , 12
50160   .Buttons.Add , , , , 13
50170   .Buttons.Add , , , tbrSeparator
50180   .Buttons.Add , , , , 14
50190  End With
50200  SetLanguageToolbar
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

Private Sub DocumentPrint()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, j As Long
50020  For j = 1 To LvwGetCountSelectedItems(lsv, True)
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
50130   aLen = aLen + FileLen(cFiles.item(i))
50140  Next i
50150  OnlyPsFiles = True
50160  For i = 1 To cFiles.Count
50170   SplitPath cFiles.item(i), , , , , Ext
50180   If UCase$(Ext) <> "PS" Then
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
50390   aLen = aLen + FileLen(cFiles.item(i))
50400  Next i
50410  For i = 1 To cFiles.Count
50420   If IsPostscriptFile(cFiles.item(i)) = True Then
50430     tFilename = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory & "\" & GetUsername, "~PD")
50440     DoEvents
50450     FileCopy cFiles.item(i), tFilename
50460    Else
50470     DoEvents
50480     ShellAndWait Me.hwnd, "print", cFiles.item(i), "", vbNullChar, wHidden, WCTermination, 60000, True
50490     DoEvents
50500   End If
50510   tLen = tLen + FileLen(cFiles.item(i))
50520   stb.Panels("Percent").Text = Format$(tLen / aLen, " 0.0%")
50530   DoEvents
50540  Next i
50550  If DefaultPrintername <> vbNullString Then
50560   SetDefaultprinterInProg DefaultPrintername
50570  End If
50580  stb.Panels("Percent").Text = vbNullString
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

Private Sub DocumentDelete()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long
50020  For i = 1 To lsv.ListItems.Count
50030   If lsv.ListItems(i).Selected = True Then
50040    KillFile lsv.ListItems(i).SubItems(4)
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

Private Sub DocumentTop()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, j As Long
50020  For j = 1 To LvwGetCountSelectedItems(lsv, True)
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

Private Sub DocumentUp()
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

Private Sub DocumentDown()
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

Private Sub DocumentBottom()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, j As Long
50020  For j = 1 To LvwGetCountSelectedItems(lsv, True)
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

Private Sub DocumentCombine()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, cFiles As Collection, tFilename As String, tFilename2 As String
50020  Set cFiles = New Collection
50030  For i = 1 To lsv.ListItems.Count
50040   If lsv.ListItems(i).Selected = True Then
50050    cFiles.Add lsv.ListItems(i).SubItems(4)
50060   End If
50070  Next i
50080  tFilename = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory & "\" & GetUsername, "~PS")
50090  KillFile tFilename
50100  If cFiles.Count > 1 Then
50110   CombineFiles tFilename, cFiles, stb
50120  End If
50130  tFilename2 = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory & "\" & GetUsername, "~PD")
50140  KillFile tFilename2
50150  Name tFilename As tFilename2
50160  Set cFiles = Nothing
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
50020  tFilename = ReplaceForbiddenChars(GetPDFTitle(lsv.ListItems(lsv.SelectedItem.Index).SubItems(4)), , ".")
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
50140   FileCopy lsv.ListItems(lsv.SelectedItem.Index).SubItems(4), cFiles.item(1)
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
   
Private Sub SetMenuPrinterStop()
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

Private Sub tlb_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case Button.Index
        Case 1
50030    SetMenuPrinterStop
50040   Case 2
50050    frmOptions.Show , Me
50060   Case 3
50070    frmLog.Show , Me
50080   Case 5
50090    DocumentPrint
50100   Case 6
50110    DocumentAdd
50120   Case 7
50130    DocumentDelete
50140   Case 8
50150    DocumentTop
50160   Case 9
50170    DocumentUp
50180   Case 10
50190    DocumentDown
50200   Case 11
50210    DocumentBottom
50220   Case 12
50230    DocumentCombine
50240   Case 13
50250    DocumentSave
50260   Case 15
50270    Call HTMLHelp_ShowTopic("html\welcome.htm")
50280  End Select
50290  SetDocumentMenu
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
