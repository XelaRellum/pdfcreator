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
      Picture         =   "frmMain.frx":8F72
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

Private LanguagePath As String, Languagefile As String, mutex As clsMutex, _
 Printjobs As Collection, NoProcessing As Boolean

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
  gsvers As Collection, tsf() As String, tStr As String, Path As String, _
  NoAbortIfRunning As Boolean, NoProccesing As Boolean, Files As Collection, _
  i As Long, tStrf() As String, PDFCreatorPrintername As String
50070
50080 '##############################################
50090 'Performance Tools
50100  Dim LastStop As Currency
50110  LastStop = ExactTimer_Value()
50120 '##############################################
50130 ' ReadVersionInfo
50140  ErrPtnr.SetProgInfo App.EXEName + " " & GetProgramReleaseStr
50150
50160  Me.KeyPreview = True: Restart = False
50170
50180  INIFilename = App.EXEName & ".ini"
50190
50200  PDFCreatorLogfilePath = CompletePath(GetTempPath) & "PDFCreator\"
50210
50220  Set reg = New clsRegistry
50230 ' reg.hkey = HKEY_CURRENT_USER
50240 ' reg.KeyRoot = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
50250 ' MsgBox "PDFCreator:" & vbCrLf & _
        "TempPath:" & vbTab & CompletePath(GetTempPath) & vbCrLf & _
        "TempPathApi:" & vbTab & CompletePath(GetTempPathApi) & vbCrLf & _
        "MyAppData:" & vbTab & GetMyAppData & vbCrLf & _
        "MyFiles:" & vbTab & vbTab & GetMyFiles & vbCrLf & _
        "App.Path:" & vbTab & App.Path & vbCrLf & _
        "Reg-MyAppData" & vbTab & reg.GetRegistryValue("AppData") & vbCrLf & _
        "Reg-MyFiles" & vbTab & reg.GetRegistryValue("Personal")
50330
50340  SavePasswordsForThisSession = False
50350  ChangeDefaultprinter = False
50360
50370  If InstalledAsServer = True Then
50380    PDFCreatorINIFile = CompletePath(GetPDFCreatorApplicationPath) & INIFilename
50390   Else
50400    If DirExists(GetMyAppData) = True Then
50410      tStr = CompletePath(GetMyAppData) & "PDFCreator"
50420      If DirExists(tStr) = False Then
50430       MakePath tStr
50440      End If
50450      PDFCreatorINIFile = CompletePath(tStr) & INIFilename
50460     Else
50470      PDFCreatorINIFile = CompletePath(GetPDFCreatorApplicationPath) & INIFilename
50480    End If
50490  End If
50500
50510  Options = ReadOptions
50520
50530  If UCase$(CommandSwitch("NO", True)) = "ABORTIFRUNNING" Then
50540    NoAbortIfRunning = True
50550   Else
50560    NoAbortIfRunning = False
50570  End If
50580  If UCase$(CommandSwitch("NO", True)) = "PROCESSING" Then
50590    NoProcessing = True
50600   Else
50610    NoProcessing = False
50620  End If
50630  If UCase$(CommandSwitch("L", True)) = "OG" Then
50640    enableSpecialLogging = True
50650   Else
50660    enableSpecialLogging = False
50670  End If
50680  If UCase$(CommandSwitch("SHOW", True)) = "ONLYOPTIONS" Then
50690    ShowOnlyOptions = True
50700    NoAbortIfRunning = True
50710    NoProcessing = True
50720   Else
50730    ShowOnlyOptions = False
50740  End If
50750  If UCase$(CommandSwitch("SHOW", True)) = "ONLYLOGFILE" Then
50760    ShowOnlyLogfile = True
50770    NoAbortIfRunning = True
50780    NoProcessing = True
50790   Else
50800    ShowOnlyLogfile = False
50810  End If
50820
50830  If IsWin9xMe = False Then
50841   Select Case Options.ProcessPriority
         Case 0: 'Idle
50860     SetProcessPriority Idle
50870    Case 1: 'Normal
50880     SetProcessPriority Normal
50890    Case 2: 'High
50900     SetProcessPriority High
50910    Case 3: 'Realtime
50920     SetProcessPriority RealTime
50930   End Select
50940  End If
50950
50960  LanguagePath = CompletePath(GetPDFCreatorApplicationPath) & "Languages\"
50970
50980  Languagefile = LanguagePath & Options.Language & ".ini"
50990  LoadLanguage Languagefile
51000
51010  IfLoggingWriteLogfile "PDFCreator Program Start"
51020  IfLoggingWriteLogfile "MyAppData: " & GetMyAppData
51030  IfLoggingWriteLogfile "PDFCreatorINIFile: " & PDFCreatorINIFile
51040  ' The program has commandswitches
51050  ' -IPTRUE : Install Printer
51060  ' -IPFALSE: UnInstall Printer
51070  ' -NSTRUE: No Start
51080  ' -ULTRUE: Unload all PDFCreator programs
51090  ' -PPDFCREATORPRINTER: The printer call the program
51100  ' -IF: Inputfile
51110  ' -OF: Outputfile
51120  ' -PF: a printable file
51130  ' -CLEARCACHE: Clear the temp. cache
51140
51150  ' Initialize unload running program
51160  If UCase$(CommandSwitch("UL", True)) = "TRUE" Then
51170   fn = FreeFile
51180   tStr = CompletePath(App.Path) & "Unload.tmp"
51190   If FileExists(tStr) = False Then
51200    Open tStr For Output As #fn
51210    Close #fn
51220   End If
51230  End If
51240
51250  ' Clear the cache
51260  If UCase$(CommandSwitch("CLEAR", True)) = "CACHE" Then
51270   ClearCache
51280  End If
51290
51300  ' print a printable file
51310  InFile = UCase$(CommandSwitch("PF", True))
51320  If Len(InFile) > 0 Then
51330   If UCase$(Printer.DeviceName) <> UCase$(GetPDFCreatorPrintername) Then
51340    If Options.NoConfirmMessageSwitchingDefaultprinter = 0 Then
51350     If ChangeDefaultprinter = False Then
51360      frmSwitchDefaultprinter.Show vbModal, Me
51370      If ChangeDefaultprinter = False Then
51380       End
51390      End If
51400     End If
51410    End If
51420    PDFCreatorPrintername = GetPDFCreatorPrintername
51430    If LenB(PDFCreatorPrintername) = 0 Then
51440     MsgBox LanguageStrings.MessagesMsg26 & " [1]"
51450     End
51460    End If
51470    DefaultPrintername = Printer.DeviceName
51480    SetDefaultprinterInProg PDFCreatorPrintername
51490   End If
51500   Set Files = GetFiles(InFile, "")
51510   If Files.Count > 0 Then
51520     DoEvents
51530     For i = 1 To Files.Count
51540      tStrf = Split(Files(i), "|")
51550      ShellAndWait "print", tStrf(1), "", tStrf(0), wHidden, WCTermination, 0, True
51560      DoEvents
51570     Next i
51580    Else
51590     MsgBox LanguageStrings.MessagesMsg14
51600   End If
51610   If DefaultPrintername <> vbNullString Then
51620    SetDefaultprinterInProg DefaultPrintername
51630   End If
51640  End If
51650  InFile = ""
51660
51670  If UCase$(CommandSwitch("NO", True)) = "START" Then
51680   End
51690  End If
51700
51710  CreatePDFCreatorTempfolder
51720
51730  GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
51740
51750  If GsDllLoaded = 0 Then
51760    IfLoggingWriteLogfile ("Cannot load " & CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
51770    If FileExists(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll) = False Then
51780      IfLoggingWriteLogfile ("Found " & CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll & " = False")
51790     Else
51800      IfLoggingWriteLogfile ("Found " & CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll & " = True")
51810    End If
51820    Set gsvers = GetAllGhostscriptversions
51830    If gsvers.Count = 0 Then
51840      SetPrinterStop True
51850      mnPrinter(0).Checked = True
51860      MsgBox LanguageStrings.MessagesMsg08
51870     Else
51880      Set reg = New clsRegistry: reg.hkey = HKEY_LOCAL_MACHINE
51890      If InStr(gsvers.item(1), ":") Then
51900        reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
51910        Options.DirectoryGhostscriptBinaries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryBinaries"))
51920        Options.DirectoryGhostscriptLibraries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryLibraries"))
51930        Options.DirectoryGhostscriptFonts = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
51940       Else
51950        If InStr(UCase$(gsvers.item(1)), "AFPL") Then
51960         If InStr(gsvers.item(1), " ") > 0 Then
51970          tsf = Split(gsvers.item(1), " ")
51980          reg.KeyRoot = "SOFTWARE\AFPL Ghostscript\" & tsf(UBound(tsf))
51990          tStr = reg.GetRegistryValue("GS_DLL")
52000          SplitPath tStr, , Path
52010          Options.DirectoryGhostscriptBinaries = CompletePath(Path)
52020          Options.DirectoryGhostscriptFonts = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
52030          Options.DirectoryGhostscriptLibraries = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
52040          Options.DirectoryGhostscriptFonts = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
52050          If tsf(UBound(tsf)) <> "8.00" Then
52060           Options.DirectoryGhostscriptResource = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "Resource"
52070          End If
52080         End If
52090        End If
52100        If InStr(UCase$(gsvers.item(1)), "GNU") Then
52110         If InStr(gsvers.item(1), " ") > 0 Then
52120          tsf = Split(gsvers.item(1), " ")
52130          reg.KeyRoot = "SOFTWARE\GNU Ghostscript\" & tsf(UBound(tsf))
52140          tStr = reg.GetRegistryValue("GS_DLL")
52150          SplitPath tStr, , Path
52160          Options.DirectoryGhostscriptBinaries = CompletePath(Path)
52170          Options.DirectoryGhostscriptFonts = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
52180          Options.DirectoryGhostscriptLibraries = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
52190          Options.DirectoryGhostscriptFonts = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
52200         End If
52210        End If
52220        If InStr(UCase$(gsvers.item(1)), "GPL") Then
52230         If InStr(gsvers.item(1), " ") > 0 Then
52240          tsf = Split(gsvers.item(1), " ")
52250          reg.KeyRoot = "SOFTWARE\GPL Ghostscript\" & tsf(UBound(tsf))
52260          tStr = reg.GetRegistryValue("GS_DLL")
52270          SplitPath tStr, , Path
52280          Options.DirectoryGhostscriptBinaries = CompletePath(Path)
52290          Options.DirectoryGhostscriptFonts = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
52300          Options.DirectoryGhostscriptLibraries = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
52310          Options.DirectoryGhostscriptFonts = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
52320          Options.DirectoryGhostscriptResource = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "Resource"
52330         End If
52340        End If
52350      End If
52360      Set reg = Nothing
52370      IfLoggingWriteLogfile ("Try to load alternative Ghostscript: " & CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
52380      If FileExists(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll) = False Then
52390        IfLoggingWriteLogfile ("Found alternative " & CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll & " = False")
52400       Else
52410        IfLoggingWriteLogfile ("Found alternative " & CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll & " = True")
52420      End If
52430      GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
52440      If GsDllLoaded = 0 Then
52450        SetPrinterStop True
52460        mnPrinter(0).Checked = True
52470        MsgBox LanguageStrings.MessagesMsg08
52480       Else
52490        GSRevision = GetGhostscriptRevision
52500      End If
52510    End If
52520   Else
52530    GSRevision = GetGhostscriptRevision
52540  End If
52550
52560  SecurityIsPossible = False
52570  If GSRevision.intRevision >= 814 Or FileExists(CompletePath(App.Path) & "pdfenc.exe") Then
52580   SecurityIsPossible = True
52590  End If
52600  IFIsPS = False
52610  InFile = UCase$(CommandSwitch("IF", True))
52620
52630  If FileExists(InFile) = True Then
52640    If Len(UCase$(CommandSwitch("OF", True))) > 0 Then
52650      If IsPostscriptFile(InFile) = True Then
52660       If GsDllLoaded = 0 Then
52670        End
52680       End If
52690       OutFile = CommandSwitch("OF", True)
52700       SplitPath OutFile, , , , , Ext
52710        GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
52720       If GsDllLoaded = 0 Then
52730        MsgBox LanguageStrings.MessagesMsg08
52740       End If
52751       Select Case UCase$(Ext)
             Case "PDF"
52770         CallGScript InFile, OutFile, Options, PDFWriter
52780        Case "PNG"
52790         CallGScript PDFSpoolfile, OutFile, Options, PNGWriter
52800        Case "JPG"
52810         CallGScript PDFSpoolfile, OutFile, Options, JPEGWriter
52820        Case "BMP"
52830         CallGScript PDFSpoolfile, OutFile, Options, BMPWriter
52840        Case "PCX"
52850         CallGScript PDFSpoolfile, OutFile, Options, PCXWriter
52860        Case "TIF"
52870         CallGScript PDFSpoolfile, OutFile, Options, TIFFWriter
52880        Case "PS"
52890         CallGScript PDFSpoolfile, OutFile, Options, PSWriter
52900        Case "EPS"
52910         CallGScript PDFSpoolfile, OutFile, Options, EPSWriter
52920       End Select
52930      End If
52940      If GsDllLoaded <> 0 Then
52950       UnloadDLLComplete GsDllLoaded
52960      End If
52970      End
52980     Else
52990      If IsPostscriptFile(CommandSwitch("IF", True)) = True Then
53000        IfLoggingWriteLogfile "Get inputfile: " & CommandSwitch("IF", True)
53010        WriteToSpecialLogfile "Get inputfile: " & CommandSwitch("IF", True)
53020        If DirExists(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory & "\" & Environ$("REDMON_USER")) = False Then
53030         MakePath CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory & "\" & Environ$("REDMON_USER")
53040        End If
53050        Tempfile = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory & "\" & Environ$("REDMON_USER"), "~PD")
53060        KillFile Tempfile
53070        If FileExists(Tempfile) = False Then
53080         'MsgBox ">" & CommandSwitch("IF", True) & vbCrLf & ">" & Tempfile
53090         If UCase$(CommandSwitch("P", True)) = "PDFCREATORPRINTER" Then
53100           IfLoggingWriteLogfile "Move the inputfile to " & Tempfile
53110           WriteToSpecialLogfile "Move the inputfile to " & Tempfile
53120           Name CommandSwitch("IF", True) As Tempfile
53130          Else
53140           IfLoggingWriteLogfile "Copy the inputfile to " & Tempfile
53150           WriteToSpecialLogfile "Copy the inputfile to " & Tempfile
53160           FileCopy CommandSwitch("IF", True), Tempfile
53170         End If
53180        End If
53190        IFIsPS = True
53200       Else
53210        MsgBox LanguageStrings.MessagesMsg06
53220      End If
53230      DoEvents
53240    End If
53250   Else
53260    If Len(InFile) > 0 Then
53270     MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & vbCrLf & _
     "InputFile -IF" & vbCrLf & ">" & InFile & "<", vbExclamation + vbOKOnly
53290    End If
53300  End If
53310
53320  If ProgramIsRunning(PDFCreator_GUID) = True Then
53330    If Not NoAbortIfRunning Then
53340     End
53350    End If
53360   Else
53370  ' Create a mutex
53380    Set mutex = New clsMutex
53390    mutex.CreateMutex PDFCreator_GUID
53400  End If
53410
53420
53430  ' Printer has started the program
53440 ' If UCase$(CommandSwitch("P", True)) = "PDFCREATORPRINTER" Then
53450 '  If Options.UseAutosave = 1 Then
53460 '   Call Autosave
53470 '   End
53480 '  End If
53490 ' End If
53500
53510  InitProgram
53520
53530  If UCase$(CommandSwitch("P", True)) = "PDFCREATORPRINTER" Or _
  (Len(CommandSwitch("IF", True)) > 0 And IFIsPS = True) Then
53550   If lsv.ListItems.Count <= 1 And PrinterStop = False Then
53560    Me.Visible = False
53570   End If
53580  End If
53590
53600 '##############################################
53610 'MsgBox "Programmstart: " & ExactTimer_Value() - LastStop & " Sekunden"
53620 'LastStop = ExactTimer_Value()
53630 '##############################################
53640  'IfLoggingWriteLogfile "PDFCreator started in " & ExactTimer_Value() - LastStop & " seconds"
53650
53660  'Paypal Menu-image
53670  ShowPaypalMenuimage
53680
53690  If Options.OptionsEnabled = 0 Then
53700   mnPrinter(2).Enabled = False
53710  End If
53720  If Options.OptionsVisible = 0 Then
53730   mnPrinter(2).Visible = False
53740   mnPrinter(3).Visible = False
53750  End If
53760
53770  SetGSRevision
53780
53790  ReadAllLanguages LanguagePath
53800  ' Only for the first time set Interval to 10 ms
53810  Timer1.Interval = 10
53820  Timer1.Enabled = True
53830  If ShowOnlyOptions = True Then
53840    Me.Visible = False
53850    frmOptions.Show vbModal, Me
53860    End
53870  End If
53880  If ShowOnlyLogfile = True Then
53890   Me.Visible = False
53900   frmLog.Show vbModal, Me
53910   End
53920  End If
53930  Timer2.Interval = 100
53940  Timer2.Enabled = True
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
50240
50250  With Options
50260   SetFont Me, .ProgramFont, .ProgramFontCharset, .ProgramFontSize
50270  End With
50280
50290  SetLanguageMenu
50300  If Options.Logging = 1 Then
50310    mnPrinter(4).Checked = True
50320   Else
50330    mnPrinter(4).Checked = False
50340  End If
50350
50360  CheckPrintJobs
50370  DoEvents
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
50030  Set Printjobs = Nothing
50040  If Not mutex Is Nothing Then
50050   mutex.CloseMutex
50060   Set mutex = Nothing
50070  End If
50080  IfLoggingWriteLogfile "PDFCreator Program End"
50090  UnloadDLLComplete GsDllLoaded
50100  PDFSpoolerPath = CompletePath(GetSystemDirectory) & "PDFSpooler.exe"
50110  If Restart = True And FileExists(PDFSpoolerPath) = True Then
50120   ShellExecute 0, vbNullString, """" & PDFSpoolerPath & """", "-SL200 -STTRUE", App.Path, 1
50130  End If
50140  End
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

Private Sub lsv_KeyUp(KeyCode As Integer, Shift As Integer)
 Dim i As Long
 On Error Resume Next
 If KeyCode = 46 Then
  For i = 1 To lsv.ListItems.Count
   If lsv.ListItems(i).Selected = True Then
    KillFile lsv.ListItems(i).SubItems(4)
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
      ShellAndWait "print", data.Files.item(1), "", vbNullChar, wHidden, WCTermination, 60000, True
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
       ShellAndWait "print", data.Files.item(i), "", vbNullChar, wHidden, WCTermination, 60000, True
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

Private Sub lsv_OLEDragOver(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
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
50450    If UCase$(Printer.DeviceName) <> UCase$(GetPDFCreatorPrintername) And OnlyPsFiles = False Then
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
50580    SetDefaultprinterInProg GetPDFCreatorPrintername
50590    aLen = 0
50600    For i = 1 To cFiles.Count
50610     aLen = aLen + FileLen(cFiles.item(i))
50620    Next i
50630    For i = 1 To cFiles.Count
50640     If IsPostscriptFile(cFiles.item(i)) = True Then
50650       tFilename = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory & "\" & GetUsername, "~PD")
50660       DoEvents
50670       FileCopy cFiles.item(i), tFilename
50680      Else
50690       DoEvents
50700       ShellAndWait "print", cFiles.item(i), "", vbNullChar, wHidden, WCTermination, 60000, True
50710       DoEvents
50720     End If
50730     tLen = tLen + FileLen(cFiles.item(i))
50740     stb.Panels("Percent").Text = Format$(tLen / aLen, " 0.0%")
50750     DoEvents
50760    Next i
50770    If DefaultPrintername <> vbNullString Then
50780     SetDefaultprinterInProg DefaultPrintername
50790    End If
50800    stb.Panels("Percent").Text = vbNullString
50810   Case 3: ' Delete
50820    For i = 1 To lsv.ListItems.Count
50830     If lsv.ListItems(i).Selected = True Then
50840      KillFile lsv.ListItems(i).SubItems(4)
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
51180    tFilename = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory & "\" & GetUsername, "~PS")
51190    KillFile tFilename
51200    If cFiles.Count > 1 Then
51210     CombineFiles tFilename, cFiles, stb
51220    End If
51230    tFilename2 = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory & "\" & GetUsername, "~PD")
51240    KillFile tFilename2
51250    Name tFilename As tFilename2
51260    Set cFiles = Nothing
51270   Case 12: ' Save
51280    DoEvents
51290    tFilename = ReplaceForbiddenChars(GetPDFTitle(lsv.ListItems(lsv.SelectedItem.Index).SubItems(4)), , ".")
51300    If LenB(tFilename) = 0 Then
51310     SplitPath lsv.ListItems(lsv.SelectedItem.Index).SubItems(4), , , , tFilename
51320    End If
51330    Set cFiles = GetFilename(tFilename, GetMyFiles, 0, _
    LanguageStrings.ListPostscriptFiles & " (*.ps)|*.ps", _
     SaveFile, Cancel, Me)
51360    If Cancel = True Then
51370     Screen.MousePointer = vbNormal
51380     Exit Sub
51390    End If
51400    If cFiles.Count > 0 Then
51410     FileCopy lsv.ListItems(lsv.SelectedItem.Index).SubItems(4), cFiles.item(1)
51420    End If
51430  End Select
51440  Screen.MousePointer = vbNormal
51450  Timer1.Enabled = True
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
50030  If FileExists(CompletePath(App.Path) & "Unload.tmp") = True Or Restart = True Then
50040   Unload Me
50050  End If
50060  CheckPrintJobs
50070  If Not NoProcessing Then
50080   CheckForPrinting
50090  End If
50100  If lsv.ListItems.Count = 0 And LenB(CommandSwitch("IF", True)) > 0 Then
50110   End
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
50030     lsv.ListItems(1).SubItems(1) = LanguageStrings.ListWaiting
50040    Else
50050     lsv.ListItems(1).SubItems(1) = LanguageStrings.ListPrinting
50060     PDFSpoolfile = lsv.ListItems(1).SubItems(4)
50070     If PrinterStop = False Then
50080      If IsFormLoaded(frmPrinting) = False Then
50090       If InstalledAsServer Then
50100        Options = ReadOptions
50110       End If
50120       If Options.UseAutosave = 1 Then
50130         Autosave
50140        Else
50150         frmPrinting.Show , Me
50160       End If
50170 '      If Me.Visible = True Then
50180 '       Me.Show
50190 '      End If
50200      End If
50210     End If
50220     If PrinterStop = False Then
50230       mnPrinter(0).Checked = False
50240      Else
50250       mnPrinter(0).Checked = True
50260     End If
50270   End If
50280  End If
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
50200     SetTopMost Me, True, True
50210     SetTopMost Me, False, True
50220     Set LItem = lsv.ListItems.Add(, , GetPDFTitle(tFile(1)))
50230     LItem.SubItems(1) = LanguageStrings.ListWaiting
50240     LItem.SubItems(2) = tFile(3)
50250     If CLng(tFile(2)) > GB Then
50260       LItem.SubItems(3) = Format$(CDbl(tFile(2)) / GB, "0.00 " & LanguageStrings.ListGBytes)
50270      Else
50280       If CLng(tFile(2)) > MB Then
50290         LItem.SubItems(3) = Format$(CDbl(tFile(2)) / MB, "0.00 " & LanguageStrings.ListMBytes)
50300        Else
50310         If CLng(tFile(2)) > kB Then
50320           LItem.SubItems(3) = Format$(CDbl(tFile(2)) / kB, "0.00 " & LanguageStrings.ListKBytes)
50330          Else
50340           LItem.SubItems(3) = Format$(tFile(2), "0 " & LanguageStrings.ListBytes)
50350         End If
50360      End If
50370     End If
50380     LItem.SubItems(4) = tFile(1)
50390     DoEvents
50400    Else
50410 '
50420   End If
50430  Next i
50440  i = 0
50450  Do Until i + 1 >= lsv.ListItems.Count
50460   i = i + 1
50470   For j = 1 To tColl.Count
50480    tFile = Split(tColl.item(j), "|")
50490    If UCase$(tFile(1)) = UCase$(lsv.ListItems(i).SubItems(4)) Then
50500     Exit For
50510    End If
50520   Next j
50530   If j > tColl.Count Then
50540    lsv.ListItems.Remove i
50550   End If
50560   DoEvents
50570  Loop
50580  If lsv.ListItems.Count = 1 Then
50590    stb.Panels("Status").Text = LanguageStrings.ListStatus & ": " & lsv.ListItems.Count & " " & LanguageStrings.MessagesMsg01
50600   Else
50610    stb.Panels("Status").Text = LanguageStrings.ListStatus & ": " & lsv.ListItems.Count & " " & LanguageStrings.MessagesMsg02
50620  End If
50630  Set tColl = Nothing
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
50650         RunProgramAfterSaving GetShortName(OutputFilename), GetRunProgramAfterSavingProgramParameters, Options.RunProgramAfterSavingWindowstyle
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
