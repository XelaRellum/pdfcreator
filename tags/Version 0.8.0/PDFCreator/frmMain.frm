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
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   2778
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
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
      Picture         =   "frmMain.frx":058A
      Top             =   2025
      Width           =   930
   End
   Begin VB.Menu mnPrinterMain 
      Caption         =   "Printer"
      Begin VB.Menu mnPrinter 
         Caption         =   "Printer stop "
         Index           =   0
      End
      Begin VB.Menu mnPrinter 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnPrinter 
         Caption         =   "Options"
         Index           =   2
      End
      Begin VB.Menu mnPrinter 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnPrinter 
         Caption         =   "Logging"
         Index           =   4
      End
      Begin VB.Menu mnPrinter 
         Caption         =   "Logfile"
         Index           =   5
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
      End
      Begin VB.Menu mnDocument 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Add"
         Index           =   2
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Delete"
         Index           =   3
      End
      Begin VB.Menu mnDocument 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Top"
         Index           =   5
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Up"
         Index           =   6
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Down"
         Index           =   7
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Bottom"
         Index           =   8
      End
      Begin VB.Menu mnDocument 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Combine"
         Index           =   10
      End
      Begin VB.Menu mnDocument 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnDocument 
         Caption         =   "Save"
         Index           =   12
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
         Caption         =   "Paypal"
         Index           =   0
      End
      Begin VB.Menu mnHelp 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnHelp 
         Caption         =   "Homepage"
         Index           =   2
      End
      Begin VB.Menu mnHelp 
         Caption         =   "PDFCreator on Sourceforge"
         Index           =   3
      End
      Begin VB.Menu mnHelp 
         Caption         =   "Check Update"
         Index           =   4
      End
      Begin VB.Menu mnHelp 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnHelp 
         Caption         =   "Info"
         Index           =   6
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'"On Error Resume Next" Functions -> Don't add the ErrorHandler
' modGeneral: makePath
' modPDFCreator: ClearCache
' modLvwListItem: LvwInsertListItemStore
' frmPrinting: Create_eDoc
' frmSave: SetNode

Private WithEvents dl As clsDownload
Attribute dl.VB_VarHelpID = -1

Private Const TimerIntervall = 500

Private LanguagePath As String, Languagefile As String, mutex As clsMutex, _
 Printjobs As Collection

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF1 Then
  KeyCode = 0
  Call HTMLHelp_ShowTopic("html\welcome.htm")
 End If
End Sub

Private Sub Form_Load()
 Dim fn As Long, stdio As clsStdIO, cinStr As String, Tempfile As String, _
  InFile As String, OutFile As String, Ext As String, IFIsPS As Boolean, _
  INIFilename As String, reg As clsRegistry, DefaultPrintername As String, _
  gsvers As Collection, tsf() As String, tStr As String, Path As String

'##############################################
'Performance Tools
 Dim LastStop As Currency
 LastStop = ExactTimer_Value()
'##############################################
 ReadVersionInfo

 ErrPtnr.SetProgInfo App.EXEName + " " & GetProgramReleaseStr

 Me.KeyPreview = True: Restart = False

 INIFilename = App.EXEName & ".ini"

 PDFCreatorLogfilePath = CompletePath(GetTempPath) & "PDFCreator\"

 Set reg = New clsRegistry
' reg.hkey = HKEY_CURRENT_USER
' reg.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
' MsgBox "PDFCreator:" & vbCrLf & _
        "TempPath:" & vbTab & CompletePath(GetTempPath) & vbCrLf & _
        "TempPathApi:" & vbTab & CompletePath(GetTempPathApi) & vbCrLf & _
        "MyAppData:" & vbTab & GetMyAppData & vbCrLf & _
        "MyFiles:" & vbTab & vbTab & GetMyFiles & vbCrLf & _
        "App.Path:" & vbTab & App.Path & vbCrLf & _
        "Reg-MyAppData" & vbTab & reg.GetRegistryValue("AppData") & vbCrLf & _
        "Reg-MyFiles" & vbTab & reg.GetRegistryValue("Personal")

 SavePasswordsForThisSession = False
 ChangeDefaultprinter = False
 SecurityIsPossible = False
 If CheckPath(GetMyAppData) = True Then
   If Len(Dir(CompletePath(GetMyAppData) & "PDFcreator", vbDirectory)) = 0 Then
    MakePath CompletePath(GetMyAppData) & "PDFcreator"
   End If
   PDFCreatorINIFile = CompletePath(GetMyAppData) & "PDFcreator\" & INIFilename
  Else
   PDFCreatorINIFile = CompletePath(App.Path) & INIFilename
 End If

 Options = ReadOptions

 LanguagePath = App.Path & "\Languages\"
 ReadAllLanguages LanguagePath
 Languagefile = LanguagePath & Options.Language & ".ini"
 LoadLanguage Languagefile

 IfLoggingWriteLogfile "PDFCreator Program Start"
 IfLoggingWriteLogfile "MyAppData: " & GetMyAppData
 IfLoggingWriteLogfile "PDFCreatorINIFile: " & PDFCreatorINIFile
 ' The program has commandswitches
 ' -IPTRUE : Install Printer
 ' -IPFALSE: UnInstall Printer
 ' -NSTRUE: No Start
 ' -ULTRUE: Unload all PDFCreator programs
 ' -PPDFCREATORPRINTER: The printer call the program
 ' -IF: Inputfile
 ' -OF: Outputfile
 ' -PF: a printable file
 ' -CLEARCACHE: Clear the temp. cache

 ' Check Installprinter

 Select Case UCase$(CommandSwitch("IP", True))
  Case "TRUE":
   Monitorname = "PDFCreator": Portname = "PDFCreator:": DriverName = "PDFCreator": Printername = "PDFCreator"
   InstallCompletePrinter
  Case "FALSE":
   Set reg = New clsRegistry
   reg.hkey = HKEY_LOCAL_MACHINE
   reg.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
   Printername = reg.GetRegistryValue("Printername", "PDFCreator")
   DriverName = reg.GetRegistryValue("Drivername", "PDFCreator")
   Monitorname = reg.GetRegistryValue("Monitorname", "PDFCreator")
   Portname = reg.GetRegistryValue("Portname", "PDFCreator:")
   UnInstallCompletePrinter
 End Select

 ' Initialize unload running program
 If UCase$(CommandSwitch("UL", True)) = "TRUE" Then
  fn = FreeFile
  If Len(Dir(App.Path & "\Unload.tmp")) > 0 Then
   Open App.Path & "\Unload.tmp" For Output As #fn
   Close #fn
  End If
 End If

 ' Clear the cache
 If UCase$(CommandSwitch("CLEAR", True)) = "CACHE" Then
  ClearCache
 End If

 ' print a printable file
 InFile = UCase$(CommandSwitch("PF", True))
 If Len(InFile) > 0 Then
  If DirExists(InFile) = True Then
   If UCase$(Printer.DeviceName) <> "PDFCREATOR" Then
    If Options.NoConfirmMessageSwitchingDefaultprinter = 0 Then
     If ChangeDefaultprinter = False Then
      frmSwitchDefaultprinter.Show vbModal, Me
      If ChangeDefaultprinter = False Then
       End
      End If
     End If
    End If
    DefaultPrintername = Printer.DeviceName
    SetDefaultprinterInProg "PDFCreator"
   End If
   DoEvents
   ShellAndWait "print", InFile, "", vbNullChar, WVersteckt, WCTermination, 60000, True
   DoEvents
   If DefaultPrintername <> "" Then
    SetDefaultprinterInProg DefaultPrintername
   End If
  End If
 End If
 InFile = ""

 ' NS: If NS=True Then end the program here
 ' It is necessary for uninstall.
 If UCase$(CommandSwitch("NS", True)) = "TRUE" Then
  End
 End If

 CreatePDFCreatorTempfolder

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

 GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)

 If GsDllLoaded = 0 Then
   Set gsvers = GetAllGhostscriptversions
   If gsvers.Count = 0 Then
     MsgBox LanguageStrings.MessagesMsg08
    Else
     Set reg = New clsRegistry: reg.hkey = HKEY_LOCAL_MACHINE
     If InStr(gsvers.item(1), ":") Then
       reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
       Options.DirectoryGhostscriptBinaries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryBinaries"))
       Options.DirectoryGhostscriptLibraries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryLibraries"))
       Options.DirectoryGhostscriptFonts = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
      Else
       If InStr(UCase$(gsvers.item(1)), "AFPL") Then
        If InStr(gsvers.item(1), " ") > 0 Then
         tsf = Split(gsvers.item(1), " ")
         reg.KeyRoot = "SOFTWARE\AFPL Ghostscript\" & tsf(UBound(tsf))
         tStr = reg.GetRegistryValue("GS_DLL")
         SplitPath tStr, , Path
         Options.DirectoryGhostscriptBinaries = CompletePath(Path)
         Options.DirectoryGhostscriptFonts = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
         Options.DirectoryGhostscriptLibraries = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
         Options.DirectoryGhostscriptFonts = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
        End If
       End If
       If InStr(UCase$(gsvers.item(1)), "GNU") Then
        If InStr(gsvers.item(1), " ") > 0 Then
         tsf = Split(gsvers.item(1), " ")
         ErrPtnr.CallStack "KeyRoot: " & "SOFTWARE\GNU Ghostscript\" & tsf(UBound(tsf))
         reg.KeyRoot = "SOFTWARE\GNU Ghostscript\" & tsf(UBound(tsf))
         tStr = reg.GetRegistryValue("GS_DLL")
         ErrPtnr.CallStack "Value (GsDll): " & tStr
         SplitPath tStr, , Path
         Options.DirectoryGhostscriptBinaries = CompletePath(Path)
         Options.DirectoryGhostscriptFonts = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
         Options.DirectoryGhostscriptLibraries = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
         Options.DirectoryGhostscriptFonts = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
        End If
       End If
     End If
     Set reg = Nothing
     GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
     If GsDllLoaded = 0 Then
       MsgBox LanguageStrings.MessagesMsg08
      Else
       GSRevision = GetGhostscriptRevision
     End If
   End If
  Else
   GSRevision = GetGhostscriptRevision
 End If

 IFIsPS = False
 InFile = UCase$(CommandSwitch("IF", True))
 If Len(InFile) > 0 Then
  If Dir(InFile) <> "" Then
    If Len(UCase$(CommandSwitch("OF", True))) > 0 Then
      If CheckIfPSFile(InFile) = True Then
       If GsDllLoaded = 0 Then
        End
       End If
       OutFile = CommandSwitch("OF", True)
       SplitPath OutFile, , , , , Ext
'       GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
       If GsDllLoaded = 0 Then
        MsgBox LanguageStrings.MessagesMsg08
       End If
       Select Case UCase$(Ext)
        Case "PDF"
         CallGScript InFile, OutFile, Options, PDFWriter
        Case "PNG"
         CallGScript PDFSpoolfile, OutFile, Options, PNGWriter
        Case "JPG"
         CallGScript PDFSpoolfile, OutFile, Options, JPEGWriter
        Case "BMP"
         CallGScript PDFSpoolfile, OutFile, Options, BMPWriter
        Case "PCX"
         CallGScript PDFSpoolfile, OutFile, Options, PCXWriter
        Case "TIF"
         CallGScript PDFSpoolfile, OutFile, Options, TIFFWriter
        Case "PS"
         CallGScript PDFSpoolfile, OutFile, Options, PSWriter
        Case "EPS"
         CallGScript PDFSpoolfile, OutFile, Options, EPSWriter
       End Select
      End If
      If GsDllLoaded <> 0 Then
       UnloadDLLComplete GsDllLoaded
      End If
      End
     Else
      If CheckIfPSFile(CommandSwitch("IF", True)) Then
        Tempfile = GetTempFile(GetPDFCreatorTempfolder, "~PD")
        FileCopy CommandSwitch("IF", True), Tempfile
        IFIsPS = True
       Else
        MsgBox LanguageStrings.MessagesMsg06
      End If
      DoEvents
    End If
   Else
    MsgBox LanguageStrings.MessagesMsg14
  End If
 End If

 ' Create a mutex; if mutex exists then exit
 Set mutex = New clsMutex
 If mutex.CheckMutex(PDFCreator_GUID) = False Then
   mutex.CreateMutex PDFCreator_GUID
  Else
   End
 End If

 ' Printer has started the program
 If UCase$(CommandSwitch("P", True)) = "PDFCREATORPRINTER" Then
  CheckAutosaveAndPrint
 End If

 InitProgram

 If UCase$(CommandSwitch("P", True)) = "PDFCREATORPRINTER" Or _
  (Len(CommandSwitch("IF", True)) > 0 And IFIsPS = True) Then
  If lsv.ListItems.Count <= 1 Then
   Me.Visible = False
  End If
 End If

'##############################################
'MsgBox "Programmstart: " & ExactTimer_Value() - LastStop & " Sekunden"
'LastStop = ExactTimer_Value()
'##############################################
 'IfLoggingWriteLogfile "PDFCreator started in " & ExactTimer_Value() - LastStop & " seconds"

 'Paypal Menu-image
 ShowPaypalMenuimage

 If Options.OptionsEnabled = 0 Then
  mnPrinter(2).Enabled = False
 End If
 If Options.OptionsVisible = 0 Then
  mnPrinter(2).Visible = False
  mnPrinter(3).Visible = False
 End If

 SetGSRevision

 ' Only for the first time set Interval to 10 ms
 Timer1.Interval = 10
 Timer1.Enabled = True
End Sub

Private Sub Form_Resize()
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
  .Top = 0: .Left = 0
  .Width = Me.Width - 125
  .Height = Me.ScaleHeight - Abs(stb.Visible) * stb.Height
 End With
 stb.Panels("Status").Width = Me.Width - 150 - stb.Panels("Percent").Width _
  - stb.Panels("GhostscriptRevision").Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
 TerminateProgram
End Sub

Private Sub InitProgram()
 Dim Filename As String, Tempfile As String

 Printing = False
 Filename = CommandSwitch("F", True)

 If Len(Dir(Filename)) > 0 And Len(Trim$(Filename)) > 0 Then
  If FileLen(Filename) > 0 Then
   Tempfile = GetTempFile(GetPDFCreatorTempfolder, "~PD")
   FileCopy Filename, Tempfile
  End If
 End If

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

 CheckPrintJobs
 DoEvents
End Sub

Private Sub TerminateProgram()
 Timer1.Enabled = False
 Set Printjobs = Nothing
 mutex.CloseMutex
 Set mutex = Nothing
 IfLoggingWriteLogfile "PDFCreator Program End"
 UnloadDLLComplete GsDllLoaded
 If Restart = True Then
  ShellExecute 0, vbNullString, """" & App.Path & "\PDFSpooler.exe""", "-SL100 -STTRUE", App.Path, 1
 End If
 End
End Sub

Private Function GetAllLanguagesFiles(LanguagePath As String) As Collection
 Dim Languagefile As String
 Set GetAllLanguagesFiles = New Collection
 Languagefile = Dir(LanguagePath & "*.ini")
 Do While Len(Languagefile) > 0
   GetAllLanguagesFiles.Add LanguagePath & Languagefile
   Languagefile = Dir()
  DoEvents
 Loop
End Function

Private Sub ReadAllLanguages(LanguagePath As String)
 Dim Languagename As String, ini As clsINI, LangFiles As Collection, i As Long, Version As String
 mnLanguage(0).Caption = "No languages available."

 Set LangFiles = GetAllLanguagesFiles(LanguagePath)
 Set ini = New clsINI
 For i = 1 To LangFiles.Count
  ini.Filename = LangFiles.item(i)
  ini.Section = "Common"
  Languagename = ini.GetKeyFromSection("Languagename")
  Version = ini.GetKeyFromSection("Version")
  If Len(Languagename) = 0 Then
   Languagename = "No name available."
  End If
  Load mnLanguage(mnLanguage.Count)
  If Version = GetProgramRelease(False) Then
    mnLanguage(mnLanguage.Count - 1).Caption = Languagename
   Else
    mnLanguage(mnLanguage.Count - 1).Caption = Languagename & " [" & Version & "]"
  End If
  mnLanguage(mnLanguage.Count - 1).Tag = LangFiles.item(i)
  DoEvents
 Next i

 If mnLanguage.Count > 1 Then
  mnLanguage(0).Caption = "No languages available."
  mnLanguage(0).Visible = False
 End If
 Set ini = Nothing
End Sub

Private Sub SetLanguageMenu()
 Dim i As Long, Version As String, reg As clsRegistry

 For i = mnLanguage.LBound To mnLanguage.UBound
  If UCase$(Languagefile) = UCase$(mnLanguage.item(i).Tag) Then
    mnLanguage.item(i).Checked = True
   Else
    mnLanguage.item(i).Checked = False
  End If
 Next i

 With LanguageStrings
  Set reg = New clsRegistry
  With reg
   .hkey = HKEY_LOCAL_MACHINE
   .KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
   Version = .GetRegistryValue("ApplicationVersion")
  End With
  Set reg = Nothing
  Caption = App.Title & " " & GetProgramReleaseStr & " " & .CommonTitle

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

  mnViewMain.Caption = .DialogView
  mnView(0).Caption = .DialogViewStatusbar

  mnLanguageMain.Caption = .DialogLanguage

  mnHelpMain.Caption = .DialogInfo
  mnHelp(0).Caption = .DialogInfoPaypal
  mnHelp(2).Caption = .DialogInfoHomepage
  mnHelp(3).Caption = .DialogInfoPDFCreatorSourceforge
  mnHelp(4).Caption = .DialogInfoCheckUpdates
  mnHelp(6).Caption = .DialogInfoInfo

  lsv.ColumnHeaders("Date").Text = .ListDate
  lsv.ColumnHeaders("Documenttitle").Text = .ListDocumenttitle
  lsv.ColumnHeaders("Filename").Text = .ListFilename
  lsv.ColumnHeaders("Size").Text = .ListSize
  lsv.ColumnHeaders("Status").Text = .ListStatus
 End With
End Sub

Private Sub lsv_KeyUp(KeyCode As Integer, Shift As Integer)
 Dim i As Long
 On Error Resume Next
 If KeyCode = 46 Then
  For i = 1 To lsv.ListItems.Count
   If lsv.ListItems(i).Selected = True Then
    Kill lsv.ListItems(i).SubItems(4)
   End If
   DoEvents
  Next i
  LvwRemoveSelectedItems lsv, True
 End If
End Sub

Private Sub lsv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 On Error Resume Next
 If Button = 2 Then
  SetDocumentMenu
  PopupMenu mnDocumentMain, , X, Y
 End If
End Sub

Private Sub lsv_OLEDragDrop(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim tFilename As String, i As Long, aLen As Double, tLen As Double, aw As Long, _
  DefaultPrintername As String, OnlyPsFiles As Boolean, Ext As String
 On Error Resume Next
 If data.GetFormat(vbCFFiles) Then
  DefaultPrintername = ""
  If data.Files.Count = 1 Then
    If CheckIfPSFile(data.Files.item(1)) = True Then
      tFilename = GetTempFile(GetPDFCreatorTempfolder, "~PD")
      FileCopy data.Files.item(1), tFilename
     Else
      If UCase$(Printer.DeviceName) <> "PDFCREATOR" Then
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
      SetDefaultprinterInProg "PDFCreator"
      ShellAndWait "print", data.Files.item(1), "", vbNullChar, WVersteckt, WCTermination, 60000, True
      If DefaultPrintername <> "" Then
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
    If UCase$(Printer.DeviceName) <> "PDFCREATOR" And OnlyPsFiles = False Then
     If ChangeDefaultprinter = False Then
      aw = MsgBox(LanguageStrings.MessagesMsg35, vbOKCancel + vbInformation)
      If aw <> vbOK Then
       Exit Sub
      End If
     End If
    End If
    ChangeDefaultprinter = True
    DefaultPrintername = Printer.DeviceName
    SetDefaultprinterInProg "PDFCreator"
    aLen = 0
    For i = 1 To data.Files.Count
     aLen = aLen + FileLen(data.Files.item(i))
    Next i
    For i = 1 To data.Files.Count
     If CheckIfPSFile(data.Files.item(i)) = True Then
       tFilename = GetTempFile(GetPDFCreatorTempfolder, "~PD")
       DoEvents
       FileCopy data.Files.item(i), tFilename
      Else
       DoEvents
'       PrintDocument data.Files.item(i)
       ShellAndWait "print", data.Files.item(i), "", vbNullChar, WVersteckt, WCTermination, 60000, True
       DoEvents
     End If
     tLen = tLen + FileLen(data.Files.item(i))
     stb.Panels("Percent").Text = Format$(tLen / aLen, " 0.0%")
     DoEvents
    Next i
    If DefaultPrintername <> "" Then
     SetDefaultprinterInProg DefaultPrintername
    End If
  End If
 End If
End Sub

Private Sub lsv_OLEDragOver(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
 Dim i As Long
 On Error Resume Next
 If data.GetFormat(vbCFFiles) Then
  Effect = ccOLEDropEffectCopy
  For i = 1 To data.Files.Count
   If FileISPrintable(data.Files.item(i)) = False And CheckIfPSFile(data.Files.item(i)) = False Then
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
  Cancel As Boolean

 Timer1.Enabled = False
 Screen.MousePointer = vbHourglass
 DoEvents
 Select Case Index
  Case 0:
   For j = 1 To LvwGetCountSelectedItems(lsv, True)
    DoEvents
    For i = lsv.ListItems.Count To 1 Step -1
     If lsv.ListItems(i).Selected = True Then
      lsv.ListItems(i).SubItems(1) = LanguageStrings.ListPrinting
      LvwListItemToTop lsv, i, True
      Exit For
     End If
    Next i
   Next j
   SetPrinterStop False
   mnPrinter(0).Checked = False
  Case 2: ' Add
   DoEvents
   Set cFiles = GetFilename("", GetMyFiles, 0, _
    LanguageStrings.ListPostscriptFiles & " (*.ps)|*.ps|" & LanguageStrings.ListAllFiles & " (*.*)|*.*", _
     OpenFile, Cancel, Me)
   If Cancel = True Then
    Screen.MousePointer = vbNormal
    Exit Sub
   End If
   aLen = 0
   For i = 1 To cFiles.Count
    aLen = aLen + FileLen(cFiles.item(i))
   Next i

   OnlyPsFiles = True
   For i = 1 To cFiles.Count
    SplitPath cFiles.item(i), , , , , Ext
    If UCase$(Ext) <> "PS" Then
     OnlyPsFiles = False
     Exit For
    End If
   Next i
   If UCase$(Printer.DeviceName) <> "PDFCREATOR" And OnlyPsFiles = False Then
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
   SetDefaultprinterInProg "PDFCreator"
   aLen = 0
   For i = 1 To cFiles.Count
    aLen = aLen + FileLen(cFiles.item(i))
   Next i
   For i = 1 To cFiles.Count
    If CheckIfPSFile(cFiles.item(i)) = True Then
      tFilename = GetTempFile(GetPDFCreatorTempfolder, "~PD")
      DoEvents
      FileCopy cFiles.item(i), tFilename
     Else
      DoEvents
      ShellAndWait "print", cFiles.item(i), "", vbNullChar, WVersteckt, WCTermination, 60000, True
      DoEvents
    End If
    tLen = tLen + FileLen(cFiles.item(i))
    stb.Panels("Percent").Text = Format$(tLen / aLen, " 0.0%")
    DoEvents
   Next i
   If DefaultPrintername <> "" Then
    SetDefaultprinterInProg DefaultPrintername
   End If
   stb.Panels("Percent").Text = vbNullString
  Case 3: ' Delete
   For i = 1 To lsv.ListItems.Count
    If lsv.ListItems(i).Selected = True Then
     Kill lsv.ListItems(i).SubItems(4)
    End If
    DoEvents
   Next i
   LvwRemoveSelectedItems lsv, True
  Case 5: ' Top
   For j = 1 To LvwGetCountSelectedItems(lsv, True)
    For i = lsv.ListItems.Count To 1 Step -1
     If lsv.ListItems(i).Selected = True Then
      LvwListItemToTop lsv, i, True
      Exit For
     End If
    Next i
   Next j
  Case 6: ' Up
   LvwListItemUp lsv, , True
  Case 7: ' Down
   LvwListItemDown lsv, , True
  Case 8: ' Bottom
   For j = 1 To LvwGetCountSelectedItems(lsv, True)
    For i = 1 To lsv.ListItems.Count
     If lsv.ListItems(i).Selected = True Then
      LvwListItemToBottom lsv, i, True
      Exit For
     End If
    Next i
   Next j
  Case 10: ' Combine
   Set cFiles = New Collection
   For i = 1 To lsv.ListItems.Count
    If lsv.ListItems(i).Selected = True Then
     cFiles.Add lsv.ListItems(i).SubItems(4)
    End If
   Next i
   tFilename = GetTempFile(GetPDFCreatorTempfolder, "~PD")
   Kill tFilename
   If cFiles.Count > 1 Then
    CombineFiles tFilename, cFiles, stb
   End If
   Set cFiles = Nothing
  Case 12: ' Save
   DoEvents
   SplitPath lsv.ListItems(lsv.SelectedItem.Index).SubItems(4), , , , tFilename
   Set cFiles = GetFilename(tFilename, GetMyFiles, 0, _
    LanguageStrings.ListPostscriptFiles & " (*.ps)|*.ps", _
     saveFile, Cancel, Me)
   If Cancel = True Then
    Screen.MousePointer = vbNormal
    Exit Sub
   End If
   If cFiles.Count > 0 Then
    FileCopy lsv.ListItems(lsv.SelectedItem.Index).SubItems(4), cFiles.item(1)
   End If
 End Select
 Screen.MousePointer = vbNormal
 Timer1.Enabled = True
End Sub

Private Sub mnDocumentMain_Click()
 SetDocumentMenu
End Sub

Private Sub mnHelp_Click(Index As Integer)
 Dim updStr As String, updStrA() As String, aw As Long
 Set dl = New clsDownload
 Select Case Index
  Case 0:
   OpenDocument Paypal
  Case 2:
   OpenDocument Homepage
  Case 3:
   OpenDocument Sourceforge
  Case 4:
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
  Case 6:
   frmInfo.Show vbModal, Me
 End Select
End Sub

Private Sub mnPrinter_Click(Index As Integer)
 Select Case Index
  Case 0:
   If mnPrinter(Index).Checked = False Then
     SetPrinterStop True
     mnPrinter(Index).Checked = True
    Else
     SetPrinterStop False
     mnPrinter(Index).Checked = False
   End If
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
  Case 5:
   frmLog.Show , Me
  Case 7:
   End
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
 If Len(Dir(App.Path & "\Unload.tmp")) > 0 Or Restart = True Then
  Unload Me
 End If
 CheckPrintJobs
 CheckForPrinting
 If lsv.ListItems.Count = 0 And UCase$(CommandSwitch("P", True)) = "PDFCREATORPRINTER" Then
  End
 End If
 If lsv.ListItems.Count = 1 Then
  lsv.ListItems(1).Selected = True
 End If
 DoEvents
 Timer1.Interval = TimerIntervall
 Timer1.Enabled = True
End Sub

Private Sub CheckForPrinting()
 If lsv.ListItems.Count > 0 Then
  If mnPrinter(0).Checked = True Then
    lsv.ListItems(1).SubItems(1) = LanguageStrings.ListWaiting
   Else
    lsv.ListItems(1).SubItems(1) = LanguageStrings.ListPrinting
    PDFSpoolfile = lsv.ListItems(1).SubItems(4)
    If PrinterStop = False Then
     If IsFormLoaded(frmPrinting) = False Then
      If Options.UseAutosave = 1 Then
        CheckAutosaveAndPrint
       Else
        frmPrinting.Show , Me
      End If
      If Me.Visible = True Then
       Me.Show
      End If
     End If
    End If
    If PrinterStop = False Then
      mnPrinter(0).Checked = False
     Else
      mnPrinter(0).Checked = True
    End If
  End If
 End If
End Sub

Private Sub CheckPrintJobs()
 Dim Temppath As String, LItem As ListItem, tColl As Collection, _
  tFile() As String, i As Long, j As Long, kB As Long, MB As Long, GB As Long
 kB = 1024: MB = kB * 1024: GB = MB * 1024
 Set tColl = New Collection
 Temppath = GetPDFCreatorTempfolder
 Set tColl = GetFiles(GetPDFCreatorTempfolder, "~PD*.tmp")
 If tColl.Count = 0 And lsv.ListItems.Count > 0 Then
  lsv.ListItems.Clear
 End If
 For i = 1 To tColl.Count
  tFile = Split(tColl.item(i), "|")
  For j = 1 To lsv.ListItems.Count
   If UCase$(tFile(1)) = UCase$(lsv.ListItems(j).SubItems(4)) Then
    Exit For
   End If
  Next j
  If j > lsv.ListItems.Count Then
    SetTopMost Me, True, True
    SetTopMost Me, False, True
    Set LItem = lsv.ListItems.Add(, , GetPDFTitle(tFile(1)))
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
   Else
'
  End If
 Next i
 i = 0
 Do Until i + 1 >= lsv.ListItems.Count
  i = i + 1
  For j = 1 To tColl.Count
   tFile = Split(tColl.item(j), "|")
   If UCase$(tFile(1)) = UCase$(lsv.ListItems(i).SubItems(4)) Then
    Exit For
   End If
  Next j
  If j > tColl.Count Then
   lsv.ListItems.Remove i
  End If
  DoEvents
 Loop
 If lsv.ListItems.Count = 1 Then
   stb.Panels("Status").Text = "Status: " & lsv.ListItems.Count & " " & LanguageStrings.MessagesMsg01
  Else
   stb.Panels("Status").Text = "Status: " & lsv.ListItems.Count & " " & LanguageStrings.MessagesMsg02
 End If
 Set tColl = Nothing
End Sub

Private Sub SetDocumentMenu()
 Dim c As Long
 If lsv.ListItems.Count = 0 Then
   mnDocument(0).Enabled = False
   mnDocument(3).Enabled = False
   mnDocument(5).Enabled = False
   mnDocument(6).Enabled = False
   mnDocument(7).Enabled = False
   mnDocument(8).Enabled = False
   mnDocument(10).Enabled = False
   mnDocument(12).Enabled = False
   Exit Sub
  Else
   If lsv.ListItems.Count = 1 Then
    mnDocument(0).Enabled = True
    mnDocument(3).Enabled = True
    mnDocument(5).Enabled = False
    mnDocument(6).Enabled = False
    mnDocument(7).Enabled = False
    mnDocument(8).Enabled = False
    mnDocument(10).Enabled = False
    mnDocument(12).Enabled = True
    Exit Sub
   End If
 End If
 mnDocument(0).Enabled = True
 mnDocument(3).Enabled = True
 mnDocument(5).Enabled = True
 mnDocument(6).Enabled = True
 mnDocument(7).Enabled = True
 mnDocument(8).Enabled = True
 mnDocument(10).Enabled = False
 mnDocument(12).Enabled = True
 c = LvwGetCountSelectedItems(lsv, True)
 If c > 1 Then
  mnDocument(10).Enabled = True
  mnDocument(12).Enabled = False
 End If
 If lsv.SelectedItem.Index = 1 And c <= 1 Then
  mnDocument(5).Enabled = False
  mnDocument(6).Enabled = False
  mnDocument(7).Enabled = True
  mnDocument(8).Enabled = True
 End If
 If lsv.SelectedItem.Index = lsv.ListItems.Count And c <= 1 Then
  mnDocument(5).Enabled = True
  mnDocument(6).Enabled = True
  mnDocument(7).Enabled = False
  mnDocument(8).Enabled = False
 End If
End Sub

Private Sub CheckAutosaveAndPrint()
 Dim tColl As Collection, i As Long, tFile() As String, Pathname As String

 If Options.UseAutosave = 1 Then
  Set tColl = GetFiles(GetPDFCreatorTempfolder, "~P*.tmp")

'  GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
  Do While tColl.Count > 0
   For i = 1 To tColl.Count
    tFile = Split(tColl.item(i), "|")
    SplitPath GetAutosaveFilename(tFile(1)), , Pathname
    If Len(Dir(Pathname, vbDirectory + vbHidden)) = 0 Then
      If Options.UseAutosaveDirectory = 1 Then
        IfLoggingWriteLogfile "Error: AutoSaveDirectory not found."
       Else
        IfLoggingWriteLogfile "Error: LastSaveDirectory not found."
      End If
     Else
      CallGScript tFile(1), GetAutosaveFilename(tFile(1)), Options, Options.AutosaveFormat
      If Len(Dir(tFile(1))) > 0 Then
       Kill tFile(1)
      End If
    End If
   Next i
   Set tColl = GetFiles(GetPDFCreatorTempfolder, "~P*.tmp")
  Loop
  If GsDllLoaded = 0 Then
    MsgBox LanguageStrings.MessagesMsg08
   Else
    UnloadDLLComplete GsDllLoaded
  End If
  End
 End If
End Sub

Private Sub ShowPaypalMenuimage()
 Dim h1 As Long, h2 As Long, com As Long
 h1 = GetMenu(Me.hwnd): h2 = GetSubMenu(h1, 4)
 com = GetMenuItemID(h2, 0)
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
