Attribute VB_Name = "modMain"
Option Explicit

Public InputFilename As String, ShowAnimationWindow As Boolean, _
 ComputerScreenResolution As Long

Private UnLoadFile As Boolean, ClearCacheDir As Boolean, _
 OutputFilename As String, InitSettings As Boolean, frmMainSUP As Long, _
 AddWindowsExplorerIntegration As Boolean, RemoveWindowsExplorerIntegration As Boolean

Private bInstallPrinter As Boolean, InstallPrinterName As String, bUninstallPrinter As Boolean, UnInstallPrinterName As String
Private bInstallWindowsPrinter As Boolean, bUninstallWindowsPrinter As Boolean
Private SetupLogFile As String, bNoMsg As Boolean, OutputSubFormat As String

Public Sub Main()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If App.StartMode = vbSModeStandalone Or IsInIDE Then
50020    InstanceCounter = InstanceCounter + 1
50030    ProgramIsVisible = True
50040    StartProgram
50050   Else
50060    ProgramWindowState = 1
50070  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modMain", "Main")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub StartProgram(Optional Params As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim filename As String, res As Boolean
50020
50030  ' Reduce the working size of used memory
50040  Call SetProcessWorkingSetSize(GetCurrentProcess(), -1, -1)
50050
50060  AnalyzeCommandlineParameters Params
50070
50080  If CheckInstance Then
50090   CheckProgramInstances
50100  End If
50110
50120  InitProgram
50130
50140  IfLoggingWriteLogfile "PDFCreator Program Start"
50150  IfLoggingWriteLogfile "Windowsversion: " & GetWinVersionStr
50160  If UseINI Then
50170    IfLoggingWriteLogfile "UseINI: True"
50180    IfLoggingWriteLogfile "PDFCreatorINIFile: " & PDFCreatorINIFile
50190   Else
50200    IfLoggingWriteLogfile "UseINI: False"
50210  End If
50220  If InstalledAsServer Then
50230    IfLoggingWriteLogfile "InstalledAsServer: True"
50240   Else
50250    IfLoggingWriteLogfile "InstalledAsServer: False"
50260  End If
50270  IfLoggingWriteLogfile "MyAppData: " & GetMyAppData
50280
50290  If UnLoadFile Then
50300   CreateUnloadFile
50310   InstanceCounter = InstanceCounter - 1
50320   Exit Sub
50330  End If
50340
50350  CheckForUpdate
50360
50370  If AddWindowsExplorerIntegration = True And RemoveWindowsExplorerIntegration = False Then
50380   AddExplorerIntegration
50390  End If
50400  If AddWindowsExplorerIntegration = False And RemoveWindowsExplorerIntegration = True Then
50410   RemoveExplorerIntegration
50420  End If
50430
50440  If ClearCacheDir Then
50450   ClearCache
50460  End If
50470
50480  If InitSettings Then
50490   SaveOptions Options
50500  End If
50510
50520  If LenB(Trim(SetupLogFile)) = 0 Then
50530   SetupLogFile = CompletePath(App.Path) & "SetupLog.txt"
50540  End If
50550
50560  If bUninstallPrinter Then
50570   If PrinterIsInstalled(UnInstallPrinterName) Then
50580     res = UnInstallPrinter(UnInstallPrinterName, "")
50590    Else
50600     IfLoggingWriteLogfile "Printer '" & UnInstallPrinterName & "' isn't installed!"
50610     If bNoMsg = True Then
50620      MsgBox "Printer '" & UnInstallPrinterName & "' isn't installed!", vbOKOnly + vbExclamation
50630     End If
50640   End If
50650  End If
50660  If bInstallPrinter Then
50670   If PrinterIsInstalled(InstallPrinterName) Then
50680     IfLoggingWriteLogfile "Printer '" & InstallPrinterName & "' is already installed!"
50690     If bNoMsg = True Then
50700      MsgBox "Printer '" & InstallPrinterName & "' is already installed!", vbOKOnly + vbExclamation
50710     End If
50720    Else
50730     res = InstallPrinter(InstallPrinterName, "PDFCreator", "PDFCreator:", "")
50740   End If
50750  End If
50760
50770  If bUninstallWindowsPrinter Then
50780   If PrinterIsInstalled(UnInstallPrinterName) Then
50790     Call UnInstallWindowsPrinter("PDFCreator", "PDFCreator:", "PDFCreator", InstallPrinterName, SetupLogFile)
50800    Else
50810     IfLoggingWriteLogfile "Printer '" & UnInstallPrinterName & "' isn't installed!"
50820     If bNoMsg = False Then
50830      MsgBox "Printer '" & UnInstallPrinterName & "' isn't installed!", vbOKOnly + vbExclamation
50840     End If
50850   End If
50860  End If
50870  If bInstallWindowsPrinter Then
50880   If PrinterIsInstalled(InstallPrinterName) Then
50890     IfLoggingWriteLogfile "Printer '" & InstallPrinterName & "' is already installed!"
50900     If bNoMsg = False Then
50910      MsgBox "Printer '" & InstallPrinterName & "' is already installed!", vbOKOnly + vbExclamation
50920     End If
50930    Else
50940     Call InstallWindowsPrinter("PDFCreator", "PDFCreator:", "PDFCreator", InstallPrinterName, SetupLogFile, App.Path)
50950   End If
50960  End If
50970
50980  PrintFiles
50990
51000  If ShowOnlyOptions Then
51010   frmOptions.Show vbModal
51020   InstanceCounter = InstanceCounter - 1
51030   Exit Sub
51040  End If
51050
51060  If ShowOnlyLogfile Then
51070   frmLog.Show vbModal
51080   InstanceCounter = InstanceCounter - 1
51090   Exit Sub
51100  End If
51110
51120  LoadGhostscriptDLL
51130
51140  If PDFCreatorPrinter = False Then
51150   If FileExists(InputFilename) = True And LenB(OutputFilename) = 0 Then
51160     filename = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory, "~PS")
51170     KillFile filename
51180     If IsValidGraphicFile(InputFilename) Then
51190       Call Image2PS(InputFilename, filename)
51200      Else
51210       FileCopy InputFilename, filename
51220     End If
51230     If FileExists(InputFilename) And DeleteIF Then
51240      KillFile InputFilename
51250     End If
51260    Else
51270     If IsValidGraphicFile(InputFilename) Then
51280       filename = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory, "~PS")
51290       Call Image2PS(InputFilename, filename)
51300       ConvertFile filename, OutputFilename, OutputSubFormat
51310       If FileExists(filename) Then
51320        KillFile filename
51330       End If
51340      Else
51350       ConvertFile InputFilename, OutputFilename, OutputSubFormat
51360     End If
51370     If FileExists(InputFilename) And DeleteIF Then
51380      KillFile InputFilename
51390     End If
51400     If FileExists(OutputFilename) And OpenOF Then
51410      OpenDocument OutputFilename
51420     End If
51430   End If
51440  End If
51450
51460  If NoStart Then
51470   InstanceCounter = InstanceCounter - 1
51480   Exit Sub
51490  End If
51500
51510  If ProgramIsRunning(PDFCreator_GUID) Then
51520    ' There is a local running instance
51530    If Not NoAbortIfRunning Then
51540     InstanceCounter = InstanceCounter - 1
51550     Exit Sub
51560    End If
51570   Else
51580  ' Create a mutex
51590    Set mutexLocal = New clsMutex
51600    mutexLocal.CreateMutex PDFCreator_GUID
51610    Set mutexGlobal = New clsMutex
51620    ' Check for a global running instance
51630    If mutexGlobal.CheckMutex("Global\" & PDFCreator_GUID) = False Then
51640     mutexGlobal.CreateMutex "Global\" & PDFCreator_GUID
51650    End If
51660  End If
51670
51680  If IsWin9xMe = False And IsWinNT4 = False And IsWin2000 = False Then
51690   InitCommonControls
51700  End If
51710
51720  Load frmMain
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modMain", "StartProgram")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub AnalyzeCommandlineParameters(Optional Params As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim cSwitch As String
50020  If IsMissing(Params) = False Then
50030    CommandLine = Params & " " & VBA.Command$
50040   Else
50050    CommandLine = VBA.Command$
50060  End If
50070  AnalyzeAdditionalParameters
50080  If Len(CommandLine) > 0 Then
50090   If UCase$(CommandSwitch("NO", False)) = "ABORTIFRUNNING" Then
50100     NoAbortIfRunning = True
50110    Else
50120     NoAbortIfRunning = False
50130   End If
50140   If UCase$(CommandSwitch("NO", False)) = "PROCESSING" Then
50150     NoProcessing = True
50160    Else
50170     NoProcessing = False
50180   End If
50190   If UCase$(CommandSwitch("NO", False)) = "PROCESSINGATSTARTUP" Then
50200     NoProcessingAtStartup = True
50210    Else
50220     NoProcessingAtStartup = False
50230   End If
50240   If UCase$(CommandSwitch("L", False)) = "OG" Then
50250     enableSpecialLogging = True
50260    Else
50270     enableSpecialLogging = False
50280   End If
50290   If UCase$(CommandSwitch("SHOW", False)) = "ONLYOPTIONS" Then
50300     ShowOnlyOptions = True
50310     NoAbortIfRunning = True
50320     NoProcessing = True
50330    Else
50340     ShowOnlyOptions = False
50350   End If
50360   If UCase$(CommandSwitch("SHOW", False)) = "ONLYLOGFILE" Then
50370     ShowOnlyLogfile = True
50380     NoAbortIfRunning = True
50390     NoProcessing = True
50400    Else
50410     ShowOnlyLogfile = False
50420   End If
50430   If UCase$(CommandSwitch("P", False)) = "PDFCREATORPRINTER" Then
50440     PDFCreatorPrinter = True
50450    Else
50460     PDFCreatorPrinter = False
50470   End If
50480   ' Initialize unload running program
50490   If UCase$(CommandSwitch("UL", False)) = "TRUE" Then
50500     UnLoadFile = True
50510    Else
50520     UnLoadFile = False
50530   End If
50540
50550   If UCase$(CommandSwitch("NO", False)) = "MSG" Then
50560     bNoMsg = True
50570    Else
50580     bNoMsg = False
50590   End If
50600
50610   ' Clear the cache
50620   If UCase$(CommandSwitch("CLEAR", False)) = "CACHE" Then
50630     ClearCacheDir = True
50640    Else
50650     ClearCacheDir = False
50660   End If
50670
50680   ' Init settings
50690   If UCase$(CommandSwitch("IN", False)) = "IT" Then
50700     InitSettings = True
50710    Else
50720     InitSettings = False
50730   End If
50740
50750   ' Windows-Explorer integration
50760   If UCase$(CommandSwitch("ADD", False)) = "WINDOWSEXPLORERINTEGRATION" Then
50770     AddWindowsExplorerIntegration = True
50780    Else
50790     AddWindowsExplorerIntegration = False
50800   End If
50810   If UCase$(CommandSwitch("REMOVE", False)) = "WINDOWSEXPLORERINTEGRATION" Then
50820     RemoveWindowsExplorerIntegration = True
50830    Else
50840     RemoveWindowsExplorerIntegration = False
50850   End If
50860
50870   ' Check running instance
50880   If UCase$(CommandSwitch("Check", False)) = "INSTANCE" Then
50890     CheckInstance = True
50900    Else
50910     CheckInstance = False
50920   End If
50930
50940   PrintFilename = CommandSwitch("PF", True)
50950   InputFilename = CommandSwitch("IF", True)
50960   OutputFilename = CommandSwitch("OF", True)
50970   OutputSubFormat = CommandSwitch("OutputSubFormat", True)
50980
50990   ' Check delete inputfile after converting
51000   If UCase$(CommandSwitch("Delete", False)) = "IF" Then
51010     DeleteIF = True
51020    Else
51030     DeleteIF = False
51040   End If
51050
51060   ' Open the outputfile after converting
51070   If UCase$(CommandSwitch("Open", False)) = "OF" Then
51080     OpenOF = True
51090    Else
51100     OpenOF = False
51110   End If
51120
51130   ' SetupLogFile
51140   cSwitch = Trim$(CommandSwitch("SETUPLOGFILE", True))
51150   If LenB(cSwitch) > 0 Then
51160    SetupLogFile = cSwitch
51170   End If
51180
51190   ' UnInstallPrinter
51200   cSwitch = Trim$(CommandSwitch("UNINSTALLPRINTER", True))
51210   If LenB(cSwitch) > 0 Then
51220    bUninstallPrinter = True
51230    UnInstallPrinterName = cSwitch
51240   End If
51250   ' InstallPrinter
51260   cSwitch = Trim$(CommandSwitch("INSTALLPRINTER", True))
51270   If LenB(cSwitch) > 0 Then
51280    bInstallPrinter = True
51290    InstallPrinterName = cSwitch
51300   End If
51310   ' UnInstallWindowsPrinter
51320   cSwitch = Trim$(CommandSwitch("UNINSTALLWINDOWSPRINTER", True))
51330   If LenB(cSwitch) > 0 Then
51340    bUninstallWindowsPrinter = True
51350    UnInstallPrinterName = cSwitch
51360   End If
51370   ' InstallWindowsPrinter
51380   cSwitch = Trim$(CommandSwitch("INSTALLWINDOWSPRINTER", True))
51390   If LenB(cSwitch) > 0 Then
51400    bInstallWindowsPrinter = True
51410    InstallPrinterName = cSwitch
51420   End If
51430
51440   cSwitch = CommandSwitch("OPTIONSFILE", True)
51450   If LenB(cSwitch) > 0 Then
51460    If FileExists(cSwitch) = True Then
51470     Optionsfile = cSwitch
51480    End If
51490   End If
51500   If UCase$(CommandSwitch("NO", False)) = "START" Then
51510     NoStart = True
51520    Else
51530     NoStart = False
51540   End If
51550   If UCase$(CommandSwitch("NO", False)) = "PSCHECK" Then
51560     NoPSCheck = True
51570    Else
51580     NoPSCheck = False
51590   End If
51600  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modMain", "AnalyzeCommandlineParameters")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub AnalyzeAdditionalParameters()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010   If UCase$(CommandSwitch(Chr$(83) & Chr$(104) & Chr$(111) & Chr$(119), False)) = Chr$(73) & Chr$(78) & Chr$(70) & Chr$(79) Then
50020    MsgBox Chr$(80) & Chr$(68) & Chr$(70) & Chr$(67) & Chr$(114) & Chr$(101) & Chr$(97) & Chr$(116) & Chr$(111) & Chr$(114) & _
    Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & _
    Chr$(109) & Chr$(97) & Chr$(100) & Chr$(101) & Chr$(32) & Chr$(98) & Chr$(121) & Chr$(32) & Chr$(119) & Chr$(119) & Chr$(119) & Chr$(46) & Chr$(112) & Chr$(100) & Chr$(102) & Chr$(102) & Chr$(111) & Chr$(114) & Chr$(103) & Chr$(101) & Chr$(46) & Chr$(111) & Chr$(114) & Chr$(103)
50050   End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modMain", "AnalyzeAdditionalParameters")
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
50010  Dim INIFilename As String, tStr As String
50020
50030  ShellAndWaitingIsRunning = False
50040  ChangeDefaultprinter = False
50050  PrintSelectedJobs = False
50060  Restart = False
50070  SavePasswordsForThisSession = False
50080  ShowAnimationWindow = False
50090  Init_ExactTimer
50100
50110  PDFCreatorLogfilePath = CompletePath(GetTempPath) & "PDFCreator\"
50120  ErrPtnr.SetProgInfo App.EXEName & " " & GetProgramReleaseStr
50130
50140  If UseINI Then
50150   INIFilename = App.EXEName & ".ini"
50160   If InstalledAsServer = True Then
50170     PDFCreatorINIFile = CompletePath(GetPDFCreatorApplicationPath) & INIFilename
50180    Else
50190     If DirExists(GetMyAppData) = True Then
50200       tStr = CompletePath(GetMyAppData) & "PDFCreator"
50210       If DirExists(tStr) = False Then
50220        MakePath tStr
50230       End If
50240       PDFCreatorINIFile = CompletePath(tStr) & INIFilename
50250      Else
50260       PDFCreatorINIFile = CompletePath(GetPDFCreatorApplicationPath) & INIFilename
50270     End If
50280   End If
50290   If LenB(Optionsfile) > 0 Then
50300    PDFCreatorINIFile = Optionsfile
50310   End If
50320  End If
50330
50340  InitLanguagesStrings
50350  ReadLanguageFromOptions
50360  LanguagePath = CompletePath(GetPDFCreatorApplicationPath) & "Languages\"
50370  If FileExists(CompletePath(GetMyAppData) & "PDFCreator\Languages\" & Options.Language & ".ini") Then
50380    Languagefile = CompletePath(GetMyAppData) & "PDFCreator\Languages\" & Options.Language & ".ini"
50390   Else
50400    Languagefile = LanguagePath & Options.Language & ".ini"
50410  End If
50420  If UCase$(Options.Language) = "ESPANOL" And FileExists(Languagefile) = False And _
   FileExists(LanguagePath & "spanish.ini") = True Then
50440   Languagefile = LanguagePath & "spanish.ini"
50450   Options.Language = "spanish"
50460  End If
50470  If FileExists(Languagefile) = True Then
50480    LoadLanguage Languagefile
50490   Else
50500 '   If Not InstalledAsServer Then
50510 '    MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & ">" & Languagefile & "<"
50520 '   End If
50530    IfLoggingWriteLogfile "Language file >" & Languagefile & _
    "< not found! Error [" & Err.LastDllError & "]"
50550  End If
50560
50570  Options = ReadOptions
50580
50590  If LenB(Optionsfile) > 0 Then
50600   Options = ReadOptionsINI(Options, Optionsfile, False, False)
50610  End If
50620
50630  If IsWin9xMe = False Then
50641   Select Case Options.ProcessPriority
         Case 0: 'Idle
50660     SetProcessPriority Idle
50670    Case 1: 'Normal
50680     SetProcessPriority Normal
50690    Case 2: 'High
50700     SetProcessPriority High
50710    Case 3: 'Realtime
50720     SetProcessPriority RealTime
50730   End Select
50740  End If
50750
50760  CreatePDFCreatorTempfolder
50770  ComputerScreenResolution = ScreenResolution
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modMain", "InitProgram")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub CreateUnloadFile()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim fn As Long, tStr As String
50020  fn = FreeFile
50030  tStr = GetPDFCreatorApplicationPath & "Unload.tmp"
50040  If FileExists(tStr) = False Then
50050   Open tStr For Output As #fn
50060   Close #fn
50070  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modMain", "CreateUnloadFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub PrintFile(filename As String, Optional frm As Form, Optional xpPgb As XP_ProgressBar, _
 Optional lblFilename As Label, Optional lblSize As Label, Optional lblCount As Label)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim PDFCreatorPrintername As String, DefaultPrintername As String, _
  files As Collection, i As Long, tStrf() As String, tFilename As String, _
  kB As Long, MB As Long, GB As Long, tStr As String
50040  kB = 1024: MB = kB * 1024: GB = MB * 1024
50050  If Len(filename) > 0 Then
50060    If UCase$(Printer.DeviceName) <> UCase$(GetPDFCreatorPrintername) Then
50070     If Options.NoConfirmMessageSwitchingDefaultprinter = 0 Then
50080      If ChangeDefaultprinter = False Then
50090       frmSwitchDefaultprinter.Show vbModal
50100       If ChangeDefaultprinter = False Then
50110        Exit Sub
50120       End If
50130      End If
50140     End If
50150     PDFCreatorPrintername = GetPDFCreatorPrintername
50160     If LenB(PDFCreatorPrintername) = 0 Then
50170      MsgBox LanguageStrings.MessagesMsg26 & " [1]"
50180      Exit Sub
50190     End If
50200     DefaultPrintername = Printer.DeviceName
50210     SetDefaultprinterInProg PDFCreatorPrintername
50220    End If
50230    Set files = GetFiles(filename, "", SortedByName)
50240    If files.Count > 0 Then
50250      DoEvents
50260      If Not frm Is Nothing Then
50270       SetTopMost frm, True, True
50280      End If
50290      For i = 1 To files.Count
50300       tStrf = Split(files(i), "|")
50310       SplitPath tStrf(1), , , tFilename
50320       If Not lblFilename Is Nothing Then
50330        lblFilename.Caption = LanguageStrings.ListFilename & ": " & tFilename
50340       End If
50350       If Not lblSize Is Nothing Then
50360        If CLng(tStrf(2)) > GB Then
50370          tStr = Format$(CDbl(tStrf(2)) / GB, "0.00 " & LanguageStrings.ListGBytes)
50380         Else
50390          If CLng(tStrf(2)) > MB Then
50400            tStr = Format$(CDbl(tStrf(2)) / MB, "0.00 " & LanguageStrings.ListMBytes)
50410           Else
50420            If CLng(tStrf(2)) > kB Then
50430              tStr = Format$(CDbl(tStrf(2)) / kB, "0.00 " & LanguageStrings.ListKBytes)
50440             Else
50450              tStr = Format$(tStrf(2), "0 " & LanguageStrings.ListBytes)
50460          End If
50470         End If
50480        End If
50490        lblSize.Caption = LanguageStrings.ListSize & ": " & tStr
50500       End If
50510       If Not xpPgb Is Nothing Then
50520        xpPgb.value = i
50530       End If
50540       If Not lblCount Is Nothing Then
50550        lblCount.Caption = CStr(i) & " (" & CStr(files.Count) & ")"
50560        lblCount.Left = (frm.Width - lblCount.Width) / 2
50570       End If
50580       If CancelPrintfiles = True Then
50590        Exit For
50600       End If
50610       DoEvents
50620       ShellAndWait 0, "print", tStrf(1), "", tStrf(0), wHidden, WCTermination, 0, True
50630       DoEvents
50640      Next i
50650     Else
50660      MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & vbCrLf & "B: " & filename
50670    End If
50680    If DefaultPrintername <> vbNullString Then
50690     SetDefaultprinterInProg DefaultPrintername
50700    End If
50710   End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modMain", "PrintFile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub LoadGhostscriptDLL()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim gsvers As Collection, reg As clsRegistry, tsf() As String, tStr As String, _
  Path As String
50030  GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
50040
50050  If GsDllLoaded = 0 Then
50060    IfLoggingWriteLogfile ("Cannot load " & CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
50070    If FileExists(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll) = False Then
50080      IfLoggingWriteLogfile ("Found " & CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll & " = False")
50090     Else
50100      IfLoggingWriteLogfile ("Found " & CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll & " = True")
50110    End If
50120    Set gsvers = GetAllGhostscriptversions
50130    If gsvers.Count = 0 Then
50140      SetPrinterStop True
50150 '     mnPrinter(0).Checked = True
50160      MsgBox LanguageStrings.MessagesMsg08
50170     Else
50180      Set reg = New clsRegistry: reg.hkey = HKEY_LOCAL_MACHINE
50190      If InStr(gsvers.Item(1), ":") Then
50200        reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50210        Options.DirectoryGhostscriptBinaries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryBinaries"))
50220        Options.DirectoryGhostscriptLibraries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryLibraries"))
50230        Options.DirectoryGhostscriptFonts = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
50240       Else
50250        If InStr(UCase$(gsvers.Item(1)), "AFPL") Then
50260         If InStr(gsvers.Item(1), " ") > 0 Then
50270          tsf = Split(gsvers.Item(1), " ")
50280          reg.KeyRoot = "SOFTWARE\AFPL Ghostscript\" & tsf(UBound(tsf))
50290          tStr = reg.GetRegistryValue("GS_DLL")
50300          SplitPath tStr, , Path
50310          Options.DirectoryGhostscriptBinaries = CompletePath(Path)
50320          Options.DirectoryGhostscriptFonts = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
50330          Options.DirectoryGhostscriptLibraries = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50340          Options.DirectoryGhostscriptFonts = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50350          If tsf(UBound(tsf)) <> "8.00" Then
50360           Options.DirectoryGhostscriptResource = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "Resource"
50370          End If
50380         End If
50390        End If
50400        If InStr(UCase$(gsvers.Item(1)), "GNU") Then
50410         If InStr(gsvers.Item(1), " ") > 0 Then
50420          tsf = Split(gsvers.Item(1), " ")
50430          reg.KeyRoot = "SOFTWARE\GNU Ghostscript\" & tsf(UBound(tsf))
50440          tStr = reg.GetRegistryValue("GS_DLL")
50450          SplitPath tStr, , Path
50460          Options.DirectoryGhostscriptBinaries = CompletePath(Path)
50470          Options.DirectoryGhostscriptLibraries = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50480          Options.DirectoryGhostscriptFonts = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50490         End If
50500        End If
50510        If InStr(UCase$(gsvers.Item(1)), "GPL") Then
50520         If InStr(gsvers.Item(1), " ") > 0 Then
50530          tsf = Split(gsvers.Item(1), " ")
50540          reg.KeyRoot = "SOFTWARE\GPL Ghostscript\" & tsf(UBound(tsf))
50550          tStr = reg.GetRegistryValue("GS_DLL")
50560          SplitPath tStr, , Path
50570          Options.DirectoryGhostscriptBinaries = CompletePath(Path)
50580          Options.DirectoryGhostscriptLibraries = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50590          Options.DirectoryGhostscriptFonts = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50600          Options.DirectoryGhostscriptResource = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "Resource"
50610         End If
50620        End If
50630      End If
50640      Set reg = Nothing
50650      IfLoggingWriteLogfile ("Try to load alternative Ghostscript: " & CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
50660      If FileExists(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll) = False Then
50670        IfLoggingWriteLogfile ("Found alternative " & CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll & " = False")
50680       Else
50690        IfLoggingWriteLogfile ("Found alternative " & CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll & " = True")
50700      End If
50710      GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
50720      If GsDllLoaded = 0 Then
50730        SetPrinterStop True
50740        MsgBox LanguageStrings.MessagesMsg08
50750       Else
50760        GSRevision = GetGhostscriptRevision
50770      End If
50780    End If
50790   Else
50800    GSRevision = GetGhostscriptRevision
50810  End If
50820
50830  SecurityIsPossible = False
50840  If GSRevision.intRevision >= 814 Or FileExists(GetPDFCreatorApplicationPath & "pdfenc.exe") Then
50850   SecurityIsPossible = True
50860  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modMain", "LoadGhostscriptDLL")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SetFont(frm As Form, ByVal Fontname As String, ByVal Charset As Long, ByVal Fontsize As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ctl As Control, eB As isExplorerBar, ts As TabStrip, df As dmFrame, f As StdFont
50020
50030  If LenB(Trim$(Fontname)) = 0 Then
50040   Exit Sub
50050  End If
50060
50070  Set f = New StdFont
50080  f.Name = Fontname
50090  f.Size = Fontsize
50100  f.Charset = Charset
50110
50120  For Each ctl In frm.Controls
50130   If TypeOf ctl Is Label Or _
     TypeOf ctl Is Form Or _
     TypeOf ctl Is ComboBox Or _
     TypeOf ctl Is CheckBox Or _
     TypeOf ctl Is CommandButton Or _
     TypeOf ctl Is ListView Or _
     TypeOf ctl Is StatusBar Or _
     TypeOf ctl Is TextBox Or _
     TypeOf ctl Is Frame Then
50220    With ctl
50230     .Font = Fontname
50240     If Not (TypeOf ctl Is StatusBar) And Not (TypeOf ctl Is ListView) Then
50250      .Fontsize = Fontsize
50260     End If
50270     .Font.Charset = Charset
50280    End With
50290   End If
50300
50310   If TypeOf ctl Is isExplorerBar Then
50320    Set eB = ctl
50330    eB.Font.Name = Fontname
50340    eB.Font.Size = Fontsize
50350    eB.Font.Charset = Charset
50360   End If
50370   If TypeOf ctl Is TabStrip Then
50380    Set ts = ctl
50390    ts.Font.Name = Fontname
50400    ts.Font.Size = Fontsize
50410    ts.Font.Charset = Charset
50420   End If
50430   If TypeOf ctl Is dmFrame Then
50440    Set df = ctl
50450    df.Font.Name = Fontname
50460    df.Font.Size = Fontsize
50470    df.Font.Charset = Charset
50480    Set df.Font = f
50490   End If
50500  Next ctl
50510
50520  Set f = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modMain", "SetFont")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub PrintFiles()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim files As Collection
50020  If LenB(PrintFilename) > 0 Then
50030   Set files = GetFiles(PrintFilename, "", SortedByName)
50040   If files.Count > 0 Then
50050     Load frmPrintfiles
50060    Else
50070     MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & vbCrLf & "C: " & PrintFilename
50080   End If
50090   Set files = Nothing
50100  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modMain", "PrintFiles")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub PrintFile2(PrintFilename As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim files As Collection
50020  If LenB(PrintFilename) > 0 Then
50030   Set files = GetFiles(PrintFilename, "", SortedByName)
50040   If files.Count > 0 Then
50050     Load frmPrintfiles
50060    Else
50070     MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & vbCrLf & "D: " & PrintFilename
50080   End If
50090   Set files = Nothing
50100  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modMain", "PrintFile2")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub PrintURL(ByVal URL As String, Optional ByVal TimeBetweenLoadAndPrint As Long = 0)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim web As InternetExplorer
50020  Set web = New InternetExplorer
50030  web.Navigate2 URL
50040  Do Until web.ReadyState = READYSTATE_COMPLETE
50050   DoEvents
50060   If StopURLPrinting = True Then
50070    Exit Sub
50080   End If
50090  Loop
50100  DoEvents
50110  Sleep TimeBetweenLoadAndPrint
50120  DoEvents
50130  If (web.QueryStatusWB(OLECMDID_PRINT) And OLECMDF_ENABLED) = OLECMDF_ENABLED Then
50140   web.ExecWB OLECMDID.OLECMDID_PRINT, OLECMDEXECOPT.OLECMDEXECOPT_DONTPROMPTUSER
50150  End If
50160  web.Quit
50170  Set web = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modMain", "PrintURL")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
