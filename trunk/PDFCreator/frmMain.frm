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
50010  Dim fn As Long, stdio As clsStdIO, cinStr As String, Tempfile As String, _
  InFile As String, OutFile As String, Ext As String, IFIsPS As Boolean, _
  INIFilename As String, reg As clsRegistry, DefaultPrintername As String, _
  gsvers As Collection, tsf() As String, tStr As String, Path As String
50050
50060 '##############################################
50070 'Performance Tools
50080  Dim LastStop As Currency
50090  LastStop = ExactTimer_Value()
50100 '##############################################
50110 ' ReadVersionInfo
50120
50130  ErrPtnr.SetProgInfo App.EXEName + " " & GetProgramReleaseStr
50140
50150  Me.KeyPreview = True: Restart = False
50160
50170  INIFilename = App.EXEName & ".ini"
50180
50190  PDFCreatorLogfilePath = CompletePath(GetTempPath) & "PDFCreator\"
50200
50210  Set reg = New clsRegistry
50220 ' reg.hkey = HKEY_CURRENT_USER
50230 ' reg.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
50240 ' MsgBox "PDFCreator:" & vbCrLf & _
        "TempPath:" & vbTab & CompletePath(GetTempPath) & vbCrLf & _
        "TempPathApi:" & vbTab & CompletePath(GetTempPathApi) & vbCrLf & _
        "MyAppData:" & vbTab & GetMyAppData & vbCrLf & _
        "MyFiles:" & vbTab & vbTab & GetMyFiles & vbCrLf & _
        "App.Path:" & vbTab & App.Path & vbCrLf & _
        "Reg-MyAppData" & vbTab & reg.GetRegistryValue("AppData") & vbCrLf & _
        "Reg-MyFiles" & vbTab & reg.GetRegistryValue("Personal")
50320
50330  SavePasswordsForThisSession = False
50340  ChangeDefaultprinter = False
50350  SecurityIsPossible = False
50360  If CheckPath(GetMyAppData) = True Then
50370    If Len(Dir(CompletePath(GetMyAppData) & "PDFcreator", vbDirectory)) = 0 Then
50380     MakePath CompletePath(GetMyAppData) & "PDFcreator"
50390    End If
50400    PDFCreatorINIFile = CompletePath(GetMyAppData) & "PDFcreator\" & INIFilename
50410   Else
50420    PDFCreatorINIFile = CompletePath(App.Path) & INIFilename
50430  End If
50440
50450  Options = ReadOptions
50460
50470  LanguagePath = App.Path & "\Languages\"
50480  ReadAllLanguages LanguagePath
50490  Languagefile = LanguagePath & Options.Language & ".ini"
50500  LoadLanguage Languagefile
50510
50520  IfLoggingWriteLogfile "PDFCreator Program Start"
50530  IfLoggingWriteLogfile "MyAppData: " & GetMyAppData
50540  IfLoggingWriteLogfile "PDFCreatorINIFile: " & PDFCreatorINIFile
50550  ' The program has commandswitches
50560  ' -IPTRUE : Install Printer
50570  ' -IPFALSE: UnInstall Printer
50580  ' -NSTRUE: No Start
50590  ' -ULTRUE: Unload all PDFCreator programs
50600  ' -PPDFCREATORPRINTER: The printer call the program
50610  ' -IF: Inputfile
50620  ' -OF: Outputfile
50630  ' -PF: a printable file
50640  ' -CLEARCACHE: Clear the temp. cache
50650
50660  ' Check Installprinter
50670
50681  Select Case UCase$(CommandSwitch("IP", True))
        Case "TRUE":
50700    Monitorname = "PDFCreator": Portname = "PDFCreator:": DriverName = "PDFCreator": Printername = "PDFCreator"
50710    InstallCompletePrinter
50720   Case "FALSE":
50730    Set reg = New clsRegistry
50740    reg.hkey = HKEY_LOCAL_MACHINE
50750    reg.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50760    Printername = reg.GetRegistryValue("Printername", "PDFCreator")
50770    DriverName = reg.GetRegistryValue("Drivername", "PDFCreator")
50780    Monitorname = reg.GetRegistryValue("Monitorname", "PDFCreator")
50790    Portname = reg.GetRegistryValue("Portname", "PDFCreator:")
50800    UnInstallCompletePrinter
50810  End Select
50820
50830  ' Initialize unload running program
50840  If UCase$(CommandSwitch("UL", True)) = "TRUE" Then
50850   fn = FreeFile
50860   If Len(Dir(App.Path & "\Unload.tmp")) > 0 Then
50870    Open App.Path & "\Unload.tmp" For Output As #fn
50880    Close #fn
50890   End If
50900  End If
50910
50920  ' Clear the cache
50930  If UCase$(CommandSwitch("CLEAR", True)) = "CACHE" Then
50940   ClearCache
50950  End If
50960
50970  ' print a printable file
50980  InFile = UCase$(CommandSwitch("PF", True))
50990  If Len(InFile) > 0 Then
51000   If DirExists(InFile) = True Then
51010    If UCase$(Printer.DeviceName) <> "PDFCREATOR" Then
51020     If Options.NoConfirmMessageSwitchingDefaultprinter = 0 Then
51030      If ChangeDefaultprinter = False Then
51040       frmSwitchDefaultprinter.Show vbModal, Me
51050       If ChangeDefaultprinter = False Then
51060        End
51070       End If
51080      End If
51090     End If
51100     DefaultPrintername = Printer.DeviceName
51110     SetDefaultprinterInProg "PDFCreator"
51120    End If
51130    DoEvents
51140    ShellAndWait "print", InFile, "", vbNullChar, WVersteckt, WCTermination, 60000, True
51150    DoEvents
51160    If DefaultPrintername <> "" Then
51170     SetDefaultprinterInProg DefaultPrintername
51180    End If
51190   End If
51200  End If
51210  InFile = ""
51220
51230  ' NS: If NS=True Then end the program here
51240  ' It is necessary for uninstall.
51250  If UCase$(CommandSwitch("NS", True)) = "TRUE" Then
51260   End
51270  End If
51280
51290  CreatePDFCreatorTempfolder
51300
51310  If IsWin9xMe = False Then
51321   Select Case Options.ProcessPriority
         Case 0: 'Idle
51340     SetProcessPriority Idle
51350    Case 1: 'Normal
51360     SetProcessPriority Normal
51370    Case 2: 'High
51380     SetProcessPriority High
51390    Case 3: 'Realtime
51400     SetProcessPriority RealTime
51410   End Select
51420  End If
51430
51440  GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
51450
51460  If GsDllLoaded = 0 Then
51470    Set gsvers = GetAllGhostscriptversions
51480    If gsvers.Count = 0 Then
51490      MsgBox LanguageStrings.MessagesMsg08
51500     Else
51510      Set reg = New clsRegistry: reg.hkey = HKEY_LOCAL_MACHINE
51520      If InStr(gsvers.item(1), ":") Then
51530        reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
51540        Options.DirectoryGhostscriptBinaries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryBinaries"))
51550        Options.DirectoryGhostscriptLibraries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryLibraries"))
51560        Options.DirectoryGhostscriptFonts = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
51570       Else
51580        If InStr(UCase$(gsvers.item(1)), "AFPL") Then
51590         If InStr(gsvers.item(1), " ") > 0 Then
51600          tsf = Split(gsvers.item(1), " ")
51610          reg.KeyRoot = "SOFTWARE\AFPL Ghostscript\" & tsf(UBound(tsf))
51620          tStr = reg.GetRegistryValue("GS_DLL")
51630          SplitPath tStr, , Path
51640          Options.DirectoryGhostscriptBinaries = CompletePath(Path)
51650          Options.DirectoryGhostscriptFonts = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
51660          Options.DirectoryGhostscriptLibraries = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
51670          Options.DirectoryGhostscriptFonts = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
51680         End If
51690        End If
51700        If InStr(UCase$(gsvers.item(1)), "GNU") Then
51710         If InStr(gsvers.item(1), " ") > 0 Then
51720          tsf = Split(gsvers.item(1), " ")
51730          ErrPtnr.CallStack "KeyRoot: " & "SOFTWARE\GNU Ghostscript\" & tsf(UBound(tsf))
51740          reg.KeyRoot = "SOFTWARE\GNU Ghostscript\" & tsf(UBound(tsf))
51750          tStr = reg.GetRegistryValue("GS_DLL")
51760          ErrPtnr.CallStack "Value (GsDll): " & tStr
51770          SplitPath tStr, , Path
51780          Options.DirectoryGhostscriptBinaries = CompletePath(Path)
51790          Options.DirectoryGhostscriptFonts = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
51800          Options.DirectoryGhostscriptLibraries = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
51810          Options.DirectoryGhostscriptFonts = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
51820         End If
51830        End If
51840        If InStr(UCase$(gsvers.item(1)), "GPL") Then
51850         If InStr(gsvers.item(1), " ") > 0 Then
51860          tsf = Split(gsvers.item(1), " ")
51870          ErrPtnr.CallStack "KeyRoot: " & "SOFTWARE\GPL Ghostscript\" & tsf(UBound(tsf))
51880          reg.KeyRoot = "SOFTWARE\GPL Ghostscript\" & tsf(UBound(tsf))
51890          tStr = reg.GetRegistryValue("GS_DLL")
51900          ErrPtnr.CallStack "Value (GsDll): " & tStr
51910          SplitPath tStr, , Path
51920          Options.DirectoryGhostscriptBinaries = CompletePath(Path)
51930          Options.DirectoryGhostscriptFonts = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
51940          Options.DirectoryGhostscriptLibraries = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
51950          Options.DirectoryGhostscriptFonts = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
51960         End If
51970        End If
51980      End If
51990      Set reg = Nothing
52000      GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
52010      If GsDllLoaded = 0 Then
52020        MsgBox LanguageStrings.MessagesMsg08
52030       Else
52040        GSRevision = GetGhostscriptRevision
52050      End If
52060    End If
52070   Else
52080    GSRevision = GetGhostscriptRevision
52090  End If
52100
52110  IFIsPS = False
52120  InFile = UCase$(CommandSwitch("IF", True))
52130  If Len(InFile) > 0 Then
52140   If Dir(InFile) <> "" Then
52150     If Len(UCase$(CommandSwitch("OF", True))) > 0 Then
52160       If CheckIfPSFile(InFile) = True Then
52170        If GsDllLoaded = 0 Then
52180         End
52190        End If
52200        OutFile = CommandSwitch("OF", True)
52210        SplitPath OutFile, , , , , Ext
52220 '       GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
52230        If GsDllLoaded = 0 Then
52240         MsgBox LanguageStrings.MessagesMsg08
52250        End If
52261        Select Case UCase$(Ext)
              Case "PDF"
52280          CallGScript InFile, OutFile, Options, PDFWriter
52290         Case "PNG"
52300          CallGScript PDFSpoolfile, OutFile, Options, PNGWriter
52310         Case "JPG"
52320          CallGScript PDFSpoolfile, OutFile, Options, JPEGWriter
52330         Case "BMP"
52340          CallGScript PDFSpoolfile, OutFile, Options, BMPWriter
52350         Case "PCX"
52360          CallGScript PDFSpoolfile, OutFile, Options, PCXWriter
52370         Case "TIF"
52380          CallGScript PDFSpoolfile, OutFile, Options, TIFFWriter
52390         Case "PS"
52400          CallGScript PDFSpoolfile, OutFile, Options, PSWriter
52410         Case "EPS"
52420          CallGScript PDFSpoolfile, OutFile, Options, EPSWriter
52430        End Select
52440       End If
52450       If GsDllLoaded <> 0 Then
52460        UnloadDLLComplete GsDllLoaded
52470       End If
52480       End
52490      Else
52500       If CheckIfPSFile(CommandSwitch("IF", True)) Then
52510         Tempfile = GetTempFile(GetPDFCreatorTempfolder, "~PD")
52520         FileCopy CommandSwitch("IF", True), Tempfile
52530         IFIsPS = True
52540        Else
52550         MsgBox LanguageStrings.MessagesMsg06
52560       End If
52570       DoEvents
52580     End If
52590    Else
52600     MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & vbCrLf & InFile
52610   End If
52620  End If
52630
52640  ' Create a mutex; if mutex exists then exit
52650  Set mutex = New clsMutex
52660  If mutex.CheckMutex(PDFCreator_GUID) = False Then
52670    mutex.CreateMutex PDFCreator_GUID
52680   Else
52690    End
52700  End If
52710
52720  ' Printer has started the program
52730  If UCase$(CommandSwitch("P", True)) = "PDFCREATORPRINTER" Then
52740   CheckAutosaveAndPrint
52750  End If
52760
52770  InitProgram
52780
52790  If UCase$(CommandSwitch("P", True)) = "PDFCREATORPRINTER" Or _
  (Len(CommandSwitch("IF", True)) > 0 And IFIsPS = True) Then
52810   If lsv.ListItems.Count <= 1 Then
52820    Me.Visible = False
52830   End If
52840  End If
52850
52860 '##############################################
52870 'MsgBox "Programmstart: " & ExactTimer_Value() - LastStop & " Sekunden"
52880 'LastStop = ExactTimer_Value()
52890 '##############################################
52900  'IfLoggingWriteLogfile "PDFCreator started in " & ExactTimer_Value() - LastStop & " seconds"
52910
52920  'Paypal Menu-image
52930  ShowPaypalMenuimage
52940
52950  If Options.OptionsEnabled = 0 Then
52960   mnPrinter(2).Enabled = False
52970  End If
52980  If Options.OptionsVisible = 0 Then
52990   mnPrinter(2).Visible = False
53000   mnPrinter(3).Visible = False
53010  End If
53020
53030  SetGSRevision
53040
53050  ' Only for the first time set Interval to 10 ms
53060  Timer1.Interval = 10
53070  Timer1.Enabled = True
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
50010  Dim Filename As String, Tempfile As String
50020
50030  Printing = False
50040  Filename = CommandSwitch("F", True)
50050
50060  If Len(Dir(Filename)) > 0 And Len(Trim$(Filename)) > 0 Then
50070   If FileLen(Filename) > 0 Then
50080    Tempfile = GetTempFile(GetPDFCreatorTempfolder, "~PD")
50090    FileCopy Filename, Tempfile
50100   End If
50110  End If
50120
50130  Set Printjobs = New Collection
50140
50150  stb.Panels.Clear
50160  stb.Panels.Add , "Status", ""
50170  stb.Panels.Add , "GhostscriptRevision", ""
50180  stb.Panels.Add , "Percent", ""
50190  stb.Panels("Percent").Width = 1000
50200  stb.Panels("GhostscriptRevision").Width = 1800
50210
50220  With lsv
50230   .View = lvwReport
50240   .FullRowSelect = True
50250   .HideSelection = False
50260   .ColumnHeaders.Clear
50270   .ColumnHeaders.Add , "Documenttitle", "Documenttitle", 2000
50280   .ColumnHeaders.Add , "Status", "Status", 1000
50290   .ColumnHeaders.Add , "Date", "Created on", 1700
50300   .ColumnHeaders.Add , "Size", "Size", 1500, lvwColumnRight
50310   .ColumnHeaders.Add , "Filename", "Filename", lsv.Width - 3500
50320  End With
50330
50340
50350  With Options
50360   SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50370  End With
50380
50390  SetLanguageMenu
50400  If Options.Logging = 1 Then
50410    mnPrinter(4).Checked = True
50420   Else
50430    mnPrinter(4).Checked = False
50440  End If
50450
50460  CheckPrintJobs
50470  DoEvents
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
50010  Timer1.Enabled = False
50020  Set Printjobs = Nothing
50030  mutex.CloseMutex
50040  Set mutex = Nothing
50050  IfLoggingWriteLogfile "PDFCreator Program End"
50060  UnloadDLLComplete GsDllLoaded
50070  If Restart = True Then
50080   ShellExecute 0, vbNullString, """" & App.Path & "\PDFSpooler.exe""", "-SL100 -STTRUE", App.Path, 1
50090  End If
50100  End
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
50010  Dim Languagefile As String
50020  Set GetAllLanguagesFiles = New Collection
50030  Languagefile = Dir(LanguagePath & "*.ini")
50040  Do While Len(Languagefile) > 0
50050    GetAllLanguagesFiles.Add LanguagePath & Languagefile
50060    Languagefile = Dir()
50070   DoEvents
50080  Loop
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
50010  Dim Languagename As String, ini As clsINI, LangFiles As Collection, i As Long, Version As String
50020  mnLanguage(0).Caption = "No languages available."
50030
50040  Set LangFiles = GetAllLanguagesFiles(LanguagePath)
50050  Set ini = New clsINI
50060  For i = 1 To LangFiles.Count
50070   ini.Filename = LangFiles.item(i)
50080   ini.Section = "Common"
50090   Languagename = ini.GetKeyFromSection("Languagename")
50100   Version = ini.GetKeyFromSection("Version")
50110   If Len(Languagename) = 0 Then
50120    Languagename = "No name available."
50130   End If
50140   Load mnLanguage(mnLanguage.Count)
50150   If IsCompatibleLanguageVersion(Version) = True Then
50160     mnLanguage(mnLanguage.Count - 1).Caption = Languagename
50170    Else
50180     mnLanguage(mnLanguage.Count - 1).Caption = Languagename & " [" & Version & "]"
50190   End If
50200   mnLanguage(mnLanguage.Count - 1).Tag = LangFiles.item(i)
50210   DoEvents
50220  Next i
50230
50240  If mnLanguage.Count > 1 Then
50250   mnLanguage(0).Caption = "No languages available."
50260   mnLanguage(0).Visible = False
50270  End If
50280  Set ini = Nothing
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
50450   mnHelp(0).Caption = .DialogInfoPaypal
50460   mnHelp(2).Caption = .DialogInfoHomepage
50470   mnHelp(3).Caption = .DialogInfoPDFCreatorSourceforge
50480   mnHelp(4).Caption = .DialogInfoCheckUpdates
50490   mnHelp(6).Caption = .DialogInfoInfo
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
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tFilename As String, cFiles As Collection, sFiles() As String, _
  i As Long, j As Long, aLen As Double, tLen As Double, aw As Long, _
  DefaultPrintername As String, OnlyPsFiles As Boolean, Ext As String, _
  Cancel As Boolean
50050
50060  Timer1.Enabled = False
50070  Screen.MousePointer = vbHourglass
50080  DoEvents
50091  Select Case Index
        Case 0:
50110    For j = 1 To LvwGetCountSelectedItems(lsv, True)
50120     DoEvents
50130     For i = lsv.ListItems.Count To 1 Step -1
50140      If lsv.ListItems(i).Selected = True Then
50150       lsv.ListItems(i).SubItems(1) = LanguageStrings.ListPrinting
50160       LvwListItemToTop lsv, i, True
50170       Exit For
50180      End If
50190     Next i
50200    Next j
50210    SetPrinterStop False
50220    mnPrinter(0).Checked = False
50230   Case 2: ' Add
50240    DoEvents
50250    Set cFiles = GetFilename("", GetMyFiles, 0, _
    LanguageStrings.ListPostscriptFiles & " (*.ps)|*.ps|" & LanguageStrings.ListAllFiles & " (*.*)|*.*", _
     OpenFile, Cancel, Me)
50280    If Cancel = True Then
50290     Screen.MousePointer = vbNormal
50300     Exit Sub
50310    End If
50320    aLen = 0
50330    For i = 1 To cFiles.Count
50340     aLen = aLen + FileLen(cFiles.item(i))
50350    Next i
50360
50370    OnlyPsFiles = True
50380    For i = 1 To cFiles.Count
50390     SplitPath cFiles.item(i), , , , , Ext
50400     If UCase$(Ext) <> "PS" Then
50410      OnlyPsFiles = False
50420      Exit For
50430     End If
50440    Next i
50450    If UCase$(Printer.DeviceName) <> "PDFCREATOR" And OnlyPsFiles = False Then
50460     If Options.NoConfirmMessageSwitchingDefaultprinter = 0 Then
50470      If ChangeDefaultprinter = False Then
50480       frmSwitchDefaultprinter.Show vbModal, Me
50490       If ChangeDefaultprinter = False Then
50500        Screen.MousePointer = vbNormal
50510        Exit Sub
50520       End If
50530      End If
50540     End If
50550    End If
50560    ChangeDefaultprinter = True
50570    DefaultPrintername = Printer.DeviceName
50580    SetDefaultprinterInProg "PDFCreator"
50590    aLen = 0
50600    For i = 1 To cFiles.Count
50610     aLen = aLen + FileLen(cFiles.item(i))
50620    Next i
50630    For i = 1 To cFiles.Count
50640     If CheckIfPSFile(cFiles.item(i)) = True Then
50650       tFilename = GetTempFile(GetPDFCreatorTempfolder, "~PD")
50660       DoEvents
50670       FileCopy cFiles.item(i), tFilename
50680      Else
50690       DoEvents
50700       ShellAndWait "print", cFiles.item(i), "", vbNullChar, WVersteckt, WCTermination, 60000, True
50710       DoEvents
50720     End If
50730     tLen = tLen + FileLen(cFiles.item(i))
50740     stb.Panels("Percent").Text = Format$(tLen / aLen, " 0.0%")
50750     DoEvents
50760    Next i
50770    If DefaultPrintername <> "" Then
50780     SetDefaultprinterInProg DefaultPrintername
50790    End If
50800    stb.Panels("Percent").Text = vbNullString
50810   Case 3: ' Delete
50820    For i = 1 To lsv.ListItems.Count
50830     If lsv.ListItems(i).Selected = True Then
50840      Kill lsv.ListItems(i).SubItems(4)
50850     End If
50860     DoEvents
50870    Next i
50880    LvwRemoveSelectedItems lsv, True
50890   Case 5: ' Top
50900    For j = 1 To LvwGetCountSelectedItems(lsv, True)
50910     For i = lsv.ListItems.Count To 1 Step -1
50920      If lsv.ListItems(i).Selected = True Then
50930       LvwListItemToTop lsv, i, True
50940       Exit For
50950      End If
50960     Next i
50970    Next j
50980   Case 6: ' Up
50990    LvwListItemUp lsv, , True
51000   Case 7: ' Down
51010    LvwListItemDown lsv, , True
51020   Case 8: ' Bottom
51030    For j = 1 To LvwGetCountSelectedItems(lsv, True)
51040     For i = 1 To lsv.ListItems.Count
51050      If lsv.ListItems(i).Selected = True Then
51060       LvwListItemToBottom lsv, i, True
51070       Exit For
51080      End If
51090     Next i
51100    Next j
51110   Case 10: ' Combine
51120    Set cFiles = New Collection
51130    For i = 1 To lsv.ListItems.Count
51140     If lsv.ListItems(i).Selected = True Then
51150      cFiles.Add lsv.ListItems(i).SubItems(4)
51160     End If
51170    Next i
51180    tFilename = GetTempFile(GetPDFCreatorTempfolder, "~PD")
51190    Kill tFilename
51200    If cFiles.Count > 1 Then
51210     CombineFiles tFilename, cFiles, stb
51220    End If
51230    Set cFiles = Nothing
51240   Case 12: ' Save
51250    DoEvents
51260    SplitPath lsv.ListItems(lsv.SelectedItem.Index).SubItems(4), , , , tFilename
51270    Set cFiles = GetFilename(tFilename, GetMyFiles, 0, _
    LanguageStrings.ListPostscriptFiles & " (*.ps)|*.ps", _
     saveFile, Cancel, Me)
51300    If Cancel = True Then
51310     Screen.MousePointer = vbNormal
51320     Exit Sub
51330    End If
51340    If cFiles.Count > 0 Then
51350     FileCopy lsv.ListItems(lsv.SelectedItem.Index).SubItems(4), cFiles.item(1)
51360    End If
51370  End Select
51380  Screen.MousePointer = vbNormal
51390  Timer1.Enabled = True
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
50050    OpenDocument Paypal
50060   Case 2:
50070    OpenDocument Homepage
50080   Case 3:
50090    OpenDocument Sourceforge
50100   Case 4:
50110    updStr = dl.DownloadString(UpdateURL)
50120 '   updStr = dl.DownloadString("http://localhost:8080/update.txt")
50130    If Len(updStr) > 0 Then
50140      If CheckPDFCreatorVersion(updStr) > 0 Then
50150        updStrA = Split(updStr, ".")
50160        If updStrA(3) = 0 Then
50170          updStr = "[" & updStrA(0) & "." & updStrA(1) & "." & updStrA(2) & "]"
50180         Else
50190          updStr = "[" & updStrA(0) & "." & updStrA(1) & "." & updStrA(2) & " Beta " & updStrA(3) & "]"
50200        End If
50210        aw = MsgBox(Replace$(LanguageStrings.MessagesMsg32, "%1", updStr), vbYesNo + vbQuestion)
50220        If aw = vbYes Then
50230         OpenDocument "http://www.sourceforge.net/project/showfiles.php?group_id=57796"
50240        End If
50250       Else
50260        MsgBox LanguageStrings.MessagesMsg33, vbOKOnly + vbInformation
50270      End If
50280     Else
50290      MsgBox LanguageStrings.MessagesMsg31 & ": " & dl.ErrorDescription & " [" & dl.ErrorNumber & "]", vbOKOnly + vbExclamation
50300    End If
50310   Case 6:
50320    frmInfo.Show vbModal, Me
50330  End Select
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
50030    If mnPrinter(Index).Checked = False Then
50040      SetPrinterStop True
50050      mnPrinter(Index).Checked = True
50060     Else
50070      SetPrinterStop False
50080      mnPrinter(Index).Checked = False
50090    End If
50100   Case 2:
50110    frmOptions.Show , Me
50120   Case 4:
50130    If mnPrinter(Index).Checked = False Then
50140      SetLogging True
50150      mnPrinter(Index).Checked = True
50160     Else
50170      SetLogging False
50180      mnPrinter(Index).Checked = False
50190    End If
50200   Case 5:
50210    frmLog.Show , Me
50220   Case 7:
50230    End
50240  End Select
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
50030  If Len(Dir(App.Path & "\Unload.tmp")) > 0 Or Restart = True Then
50040   Unload Me
50050  End If
50060  CheckPrintJobs
50070  CheckForPrinting
50080  If lsv.ListItems.Count = 0 And UCase$(CommandSwitch("P", True)) = "PDFCREATORPRINTER" Then
50090   End
50100  End If
50110  If lsv.ListItems.Count = 1 Then
50120   lsv.ListItems(1).Selected = True
50130  End If
50140  DoEvents
50150  Timer1.Interval = TimerIntervall
50160  Timer1.Enabled = True
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
50030     lsv.ListItems(1).SubItems(1) = LanguageStrings.ListWaiting
50040    Else
50050     lsv.ListItems(1).SubItems(1) = LanguageStrings.ListPrinting
50060     PDFSpoolfile = lsv.ListItems(1).SubItems(4)
50070     If PrinterStop = False Then
50080      If IsFormLoaded(frmPrinting) = False Then
50090       If Options.UseAutosave = 1 Then
50100         CheckAutosaveAndPrint
50110        Else
50120         frmPrinting.Show , Me
50130       End If
50140       If Me.Visible = True Then
50150        Me.Show
50160       End If
50170      End If
50180     End If
50190     If PrinterStop = False Then
50200       mnPrinter(0).Checked = False
50210      Else
50220       mnPrinter(0).Checked = True
50230     End If
50240   End If
50250  End If
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

Private Sub CheckPrintJobs()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Temppath As String, LItem As ListItem, tColl As Collection, _
  tFile() As String, i As Long, j As Long, kB As Long, MB As Long, GB As Long
50030  kB = 1024: MB = kB * 1024: GB = MB * 1024
50040  Set tColl = New Collection
50050  Temppath = GetPDFCreatorTempfolder
50060  Set tColl = GetFiles(GetPDFCreatorTempfolder, "~PD*.tmp")
50070  If tColl.Count = 0 And lsv.ListItems.Count > 0 Then
50080   lsv.ListItems.Clear
50090  End If
50100  For i = 1 To tColl.Count
50110   tFile = Split(tColl.item(i), "|")
50120   For j = 1 To lsv.ListItems.Count
50130    If UCase$(tFile(1)) = UCase$(lsv.ListItems(j).SubItems(4)) Then
50140     Exit For
50150    End If
50160   Next j
50170   If j > lsv.ListItems.Count Then
50180     SetTopMost Me, True, True
50190     SetTopMost Me, False, True
50200     Set LItem = lsv.ListItems.Add(, , GetPDFTitle(tFile(1)))
50210     LItem.SubItems(1) = LanguageStrings.ListWaiting
50220     LItem.SubItems(2) = tFile(3)
50230     If CLng(tFile(2)) > GB Then
50240       LItem.SubItems(3) = Format$(CDbl(tFile(2)) / GB, "0.00 " & LanguageStrings.ListGBytes)
50250      Else
50260       If CLng(tFile(2)) > MB Then
50270         LItem.SubItems(3) = Format$(CDbl(tFile(2)) / MB, "0.00 " & LanguageStrings.ListMBytes)
50280        Else
50290         If CLng(tFile(2)) > kB Then
50300           LItem.SubItems(3) = Format$(CDbl(tFile(2)) / kB, "0.00 " & LanguageStrings.ListKBytes)
50310          Else
50320           LItem.SubItems(3) = Format$(tFile(2), "0 " & LanguageStrings.ListBytes)
50330         End If
50340      End If
50350     End If
50360     LItem.SubItems(4) = tFile(1)
50370     DoEvents
50380    Else
50390 '
50400   End If
50410  Next i
50420  i = 0
50430  Do Until i + 1 >= lsv.ListItems.Count
50440   i = i + 1
50450   For j = 1 To tColl.Count
50460    tFile = Split(tColl.item(j), "|")
50470    If UCase$(tFile(1)) = UCase$(lsv.ListItems(i).SubItems(4)) Then
50480     Exit For
50490    End If
50500   Next j
50510   If j > tColl.Count Then
50520    lsv.ListItems.Remove i
50530   End If
50540   DoEvents
50550  Loop
50560  If lsv.ListItems.Count = 1 Then
50570    stb.Panels("Status").Text = LanguageStrings.ListStatus & ": " & lsv.ListItems.Count & " " & LanguageStrings.MessagesMsg01
50580   Else
50590    stb.Panels("Status").Text = LanguageStrings.ListStatus & ": " & lsv.ListItems.Count & " " & LanguageStrings.MessagesMsg02
50600  End If
50610  Set tColl = Nothing
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
50020  If lsv.ListItems.Count = 0 Then
50030    mnDocument(0).Enabled = False
50040    mnDocument(3).Enabled = False
50050    mnDocument(5).Enabled = False
50060    mnDocument(6).Enabled = False
50070    mnDocument(7).Enabled = False
50080    mnDocument(8).Enabled = False
50090    mnDocument(10).Enabled = False
50100    mnDocument(12).Enabled = False
50110    Exit Sub
50120   Else
50130    If lsv.ListItems.Count = 1 Then
50140     mnDocument(0).Enabled = True
50150     mnDocument(3).Enabled = True
50160     mnDocument(5).Enabled = False
50170     mnDocument(6).Enabled = False
50180     mnDocument(7).Enabled = False
50190     mnDocument(8).Enabled = False
50200     mnDocument(10).Enabled = False
50210     mnDocument(12).Enabled = True
50220     Exit Sub
50230    End If
50240  End If
50250  mnDocument(0).Enabled = True
50260  mnDocument(3).Enabled = True
50270  mnDocument(5).Enabled = True
50280  mnDocument(6).Enabled = True
50290  mnDocument(7).Enabled = True
50300  mnDocument(8).Enabled = True
50310  mnDocument(10).Enabled = False
50320  mnDocument(12).Enabled = True
50330  c = LvwGetCountSelectedItems(lsv, True)
50340  If c > 1 Then
50350   mnDocument(10).Enabled = True
50360   mnDocument(12).Enabled = False
50370  End If
50380  If lsv.SelectedItem.Index = 1 And c <= 1 Then
50390   mnDocument(5).Enabled = False
50400   mnDocument(6).Enabled = False
50410   mnDocument(7).Enabled = True
50420   mnDocument(8).Enabled = True
50430  End If
50440  If lsv.SelectedItem.Index = lsv.ListItems.Count And c <= 1 Then
50450   mnDocument(5).Enabled = True
50460   mnDocument(6).Enabled = True
50470   mnDocument(7).Enabled = False
50480   mnDocument(8).Enabled = False
50490  End If
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

Private Sub CheckAutosaveAndPrint()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tColl As Collection, i As Long, tFile() As String, Pathname As String
50020
50030  If Options.UseAutosave = 1 Then
50040   Set tColl = GetFiles(GetPDFCreatorTempfolder, "~P*.tmp")
50050
50060 '  GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
50070   Do While tColl.Count > 0
50080    For i = 1 To tColl.Count
50090     tFile = Split(tColl.item(i), "|")
50100     SplitPath GetAutosaveFilename(tFile(1)), , Pathname
50110     If Len(Dir(Pathname, vbDirectory + vbHidden)) = 0 Then
50120       If Options.UseAutosaveDirectory = 1 Then
50130         IfLoggingWriteLogfile "Error: AutoSaveDirectory not found."
50140        Else
50150         IfLoggingWriteLogfile "Error: LastSaveDirectory not found."
50160       End If
50170      Else
50180       CallGScript tFile(1), GetAutosaveFilename(tFile(1)), Options, Options.AutosaveFormat
50190       If Len(Dir(tFile(1))) > 0 Then
50200        Kill tFile(1)
50210       End If
50220     End If
50230    Next i
50240    Set tColl = GetFiles(GetPDFCreatorTempfolder, "~P*.tmp")
50250   Loop
50260   If GsDllLoaded = 0 Then
50270     MsgBox LanguageStrings.MessagesMsg08
50280    Else
50290     UnloadDLLComplete GsDllLoaded
50300   End If
50310   End
50320  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("frmMain", "CheckAutosaveAndPrint")
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
50030  com = GetMenuItemID(h2, 0)
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
