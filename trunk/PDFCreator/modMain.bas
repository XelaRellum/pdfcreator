Attribute VB_Name = "modMain"
Option Explicit

Public InputFilename As String, ShowAnimationWindow As Boolean, _
 ComputerScreenResolution As Long, In_eActionTimer As Boolean

Private UnLoadFile As Boolean, ClearCacheDir As Boolean, _
 OutputFilename As String, InitSettings As Boolean, frmMainSUP As Long, _
 AddWindowsExplorerIntegration As Boolean, RemoveWindowsExplorerIntegration As Boolean
 
Private bInstallPrinter As Boolean, InstallPrinterName As String, bUninstallPrinter As Boolean, UnInstallPrinterName As String
Private bInstallWindowsPrinter As Boolean, bUninstallWindowsPrinter As Boolean
Private SetupLogFile As String, bNoMsg As Boolean, OutputSubFormat As String
Public IsFrmMainLoaded As Boolean

Public Sub Main()
 CheckInstalledAsServer
 If App.StartMode = vbSModeStandalone Or IsInIDE Then
   InstanceCounter = InstanceCounter + 1
   ProgramIsVisible = True
   StartProgram
  Else
   ProgramWindowState = 1
 End If
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
50160  If InstalledAsServer Then
50170    IfLoggingWriteLogfile "InstalledAsServer: True"
50180   Else
50190    IfLoggingWriteLogfile "InstalledAsServer: False"
50200  End If
50210  IfLoggingWriteLogfile "MyAppData: " & GetMyAppData
50220
50230  If UnLoadFile Then
50240   CreateUnloadFile
50250   InstanceCounter = InstanceCounter - 1
50260   Exit Sub
50270  End If
50280
50290  CheckForUpdate
50300
50310  If AddWindowsExplorerIntegration = True And RemoveWindowsExplorerIntegration = False Then
50320   AddExplorerIntegration
50330  End If
50340  If AddWindowsExplorerIntegration = False And RemoveWindowsExplorerIntegration = True Then
50350   RemoveExplorerIntegration
50360  End If
50370
50380  If ClearCacheDir Then
50390   ClearCache
50400  End If
50410
50420  If InitSettings Then
50430   SaveOptions Options 'Initialize the default settings
50440  End If
50450
50460  If LenB(Trim(SetupLogFile)) = 0 Then
50470   SetupLogFile = CompletePath(App.Path) & "SetupLog.txt"
50480  End If
50490
50500  If bUninstallPrinter Then
50510   If Not IsAdmin Then
50520    MsgBox LanguageStrings.PrintersAdminNotice
50530    Exit Sub
50540   End If
50550   If PrinterIsInstalled(UnInstallPrinterName) Then
50560     res = UnInstallPrinter(UnInstallPrinterName, "")
50570    Else
50580     IfLoggingWriteLogfile "Printer '" & UnInstallPrinterName & "' isn't installed!"
50590     If bNoMsg = True Then
50600      MsgBox "Printer '" & UnInstallPrinterName & "' isn't installed!", vbOKOnly + vbExclamation
50610     End If
50620   End If
50630  End If
50640  If bInstallPrinter Then
50650   If Not IsAdmin Then
50660    MsgBox LanguageStrings.PrintersAdminNotice
50670    Exit Sub
50680   End If
50690   If PrinterIsInstalled(InstallPrinterName) Then
50700     IfLoggingWriteLogfile "Printer '" & InstallPrinterName & "' is already installed!"
50710     If bNoMsg = True Then
50720      MsgBox "Printer '" & InstallPrinterName & "' is already installed!", vbOKOnly + vbExclamation
50730     End If
50740    Else
50750     res = InstallPrinter(InstallPrinterName, "PDFCreator", "PDFCreator:", "")
50760   End If
50770  End If
50780
50790  If bUninstallWindowsPrinter Then
50800   If PrinterIsInstalled(UnInstallPrinterName) Then
50810     Call UnInstallWindowsPrinter("PDFCreator", "PDFCreator:", "PDFCreator", InstallPrinterName, SetupLogFile)
50820    Else
50830     IfLoggingWriteLogfile "Printer '" & UnInstallPrinterName & "' isn't installed!"
50840     If bNoMsg = False Then
50850      MsgBox "Printer '" & UnInstallPrinterName & "' isn't installed!", vbOKOnly + vbExclamation
50860     End If
50870   End If
50880  End If
50890  If bInstallWindowsPrinter Then
50900   If PrinterIsInstalled(InstallPrinterName) Then
50910     IfLoggingWriteLogfile "Printer '" & InstallPrinterName & "' is already installed!"
50920     If bNoMsg = False Then
50930      MsgBox "Printer '" & InstallPrinterName & "' is already installed!", vbOKOnly + vbExclamation
50940     End If
50950    Else
50960     Call InstallWindowsPrinter("PDFCreator", "PDFCreator:", "PDFCreator", InstallPrinterName, SetupLogFile, App.Path)
50970   End If
50980  End If
50990
51000  PrintFiles
51010
51020  If ShowOnlyOptions Then
51030   frmOptions.Show vbModal
51040   InstanceCounter = InstanceCounter - 1
51050   Exit Sub
51060  End If
51070
51080  If ShowOnlyLogfile Then
51090   frmLog.Show vbModal
51100   InstanceCounter = InstanceCounter - 1
51110   Exit Sub
51120  End If
51130
51140  LoadGhostscriptDLL
51150
51160  If PDFCreatorPrinter = False Then
51170   If FileExists(InputFilename) = True And LenB(OutputFilename) = 0 Then
51180     filename = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory, "~PS")
51190     KillFile filename
51200     If IsValidGraphicFile(InputFilename) Then
51210       Call Image2PS(InputFilename, filename)
51220      Else
51230       FileCopy InputFilename, filename
51240     End If
51250     If FileExists(InputFilename) And DeleteIF Then
51260      KillFile InputFilename
51270     End If
51280    Else
51290     If IsValidGraphicFile(InputFilename) Then
51300       filename = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory, "~PS")
51310       Call Image2PS(InputFilename, filename)
51320       ConvertFile filename, OutputFilename, OutputSubFormat
51330       If FileExists(filename) Then
51340        KillFile filename
51350       End If
51360      Else
51370       ConvertFile InputFilename, OutputFilename, OutputSubFormat
51380     End If
51390     If FileExists(InputFilename) And DeleteIF Then
51400      KillFile InputFilename
51410     End If
51420     If FileExists(OutputFilename) And OpenOF Then
51430      OpenDocument OutputFilename
51440     End If
51450   End If
51460  End If
51470
51480  If NoStart Then
51490   InstanceCounter = InstanceCounter - 1
51500   Exit Sub
51510  End If
51520
51530  If mutexLocal Is Nothing Then
51540   Set mutexLocal = New clsMutex
51550  End If
51560  If mutexGlobal Is Nothing Then
51570   Set mutexGlobal = New clsMutex
51580  End If
51590  If ProgramIsRunning(PDFCreator_GUID) Then
51600    ' There is a local running instance
51610    If Not NoAbortIfRunning Then
51620     InstanceCounter = InstanceCounter - 1
51630     Exit Sub
51640    End If
51650   Else
51660  ' Create a mutex
51670    mutexLocal.CreateMutex PDFCreator_GUID
51680    ' Check for a global running instance
51690    If mutexGlobal.CheckMutex("Global\" & PDFCreator_GUID) = False Then
51700     mutexGlobal.CreateMutex "Global\" & PDFCreator_GUID
51710    End If
51720  End If
51730
51740 If IsFrmMainLoaded Then
51750  Exit Sub
51760 End If
51770
51780
51790  If IsWin9xMe = False And IsWinNT4 = False And IsWin2000 = False Then
51800   InitCommonControls
51810  End If
51820
51830  Load frmMain
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
50140  InitLanguagesStrings
50150  ReadLanguageFromOptions
50160  LanguagePath = CompletePath(GetPDFCreatorApplicationPath) & "Languages\"
50170  If FileExists(CompletePath(GetMyAppData) & "PDFCreator\Languages\" & Options.Language & ".ini") Then
50180    Languagefile = CompletePath(GetMyAppData) & "PDFCreator\Languages\" & Options.Language & ".ini"
50190   Else
50200    Languagefile = LanguagePath & Options.Language & ".ini"
50210  End If
50220  If UCase$(Options.Language) = "ESPANOL" And FileExists(Languagefile) = False And _
   FileExists(LanguagePath & "spanish.ini") = True Then
50240   Languagefile = LanguagePath & "spanish.ini"
50250   Options.Language = "spanish"
50260  End If
50270  If FileExists(Languagefile) = True Then
50280    LoadLanguage Languagefile
50290   Else
50300 '   If Not InstalledAsServer Then
50310 '    MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & ">" & Languagefile & "<"
50320 '   End If
50330    IfLoggingWriteLogfile "Language file >" & Languagefile & _
    "< not found! Error [" & Err.LastDllError & "]"
50350  End If
50360
50370  Options = ReadOptions
50380  PrinterTemppath = Options.PrinterTemppath
50390  If LenB(Optionsfile) > 0 Then
50400   Options = ReadOptionsINI(Options, Optionsfile, False, False)
50410  End If
50420
50430  If IsWin9xMe = False Then
50441   Select Case Options.ProcessPriority
         Case 0: 'Idle
50460     SetProcessPriority Idle
50470    Case 1: 'Normal
50480     SetProcessPriority Normal
50490    Case 2: 'High
50500     SetProcessPriority High
50510    Case 3: 'Realtime
50520     SetProcessPriority RealTime
50530   End Select
50540  End If
50550
50560  CreatePDFCreatorTempfolder
50570  ComputerScreenResolution = ScreenResolution
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
50010
50020  Dim PDFCreatorPrintername As String, DefaultPrintername As String, _
  files As Collection, i As Long, tStrf() As String, tFilename As String, _
  kB As Long, MB As Long, GB As Long, tStr As String
50050  kB = 1024: MB = kB * 1024: GB = MB * 1024
50060  If Len(filename) > 0 Then
50070    If UCase$(Printer.DeviceName) <> UCase$(GetPDFCreatorPrintername) Then
50080     If Options.NoConfirmMessageSwitchingDefaultprinter = 0 Then
50090      If ChangeDefaultprinter = False Then
50100       frmSwitchDefaultprinter.Show vbModal
50110       If ChangeDefaultprinter = False Then
50120        Exit Sub
50130       End If
50140      End If
50150     End If
50160     PDFCreatorPrintername = GetPDFCreatorPrintername
50170     If LenB(PDFCreatorPrintername) = 0 Then
50180      MsgBox LanguageStrings.MessagesMsg26 & " [1]"
50190      Exit Sub
50200     End If
50210     DefaultPrintername = Printer.DeviceName
50220     SetDefaultprinterInProg PDFCreatorPrintername
50230    End If
50240    Set files = GetFiles(filename, "", SortedByName)
50250    If files.Count > 0 Then
50260      DoEvents
50270      If Not frm Is Nothing Then
50280       SetTopMost frm, True, True
50290      End If
50300      For i = 1 To files.Count
50310       tStrf = Split(files(i), "|")
50320       SplitPath tStrf(1), , , tFilename
50330       If Not lblFilename Is Nothing Then
50340        lblFilename.Caption = LanguageStrings.ListFilename & ": " & tFilename
50350       End If
50360       If Not lblSize Is Nothing Then
50370        If CLng(tStrf(2)) > GB Then
50380          tStr = Format$(CDbl(tStrf(2)) / GB, "0.00 " & LanguageStrings.ListGBytes)
50390         Else
50400          If CLng(tStrf(2)) > MB Then
50410            tStr = Format$(CDbl(tStrf(2)) / MB, "0.00 " & LanguageStrings.ListMBytes)
50420           Else
50430            If CLng(tStrf(2)) > kB Then
50440              tStr = Format$(CDbl(tStrf(2)) / kB, "0.00 " & LanguageStrings.ListKBytes)
50450             Else
50460              tStr = Format$(tStrf(2), "0 " & LanguageStrings.ListBytes)
50470          End If
50480         End If
50490        End If
50500        lblSize.Caption = LanguageStrings.ListSize & ": " & tStr
50510       End If
50520       If Not xpPgb Is Nothing Then
50530        xpPgb.value = i
50540       End If
50550       If Not lblCount Is Nothing Then
50560        lblCount.Caption = CStr(i) & " (" & CStr(files.Count) & ")"
50570        lblCount.Left = (frm.Width - lblCount.Width) / 2
50580       End If
50590       If CancelPrintfiles = True Then
50600        Exit For
50610       End If
50620       DoEvents
50630       ShellAndWait 0, "print", tStrf(1), "", tStrf(0), wHidden, WCTermination, 0, True
50640       DoEvents
50650      Next i
50660     Else
50670      MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & vbCrLf & filename
50680    End If
50690    If DefaultPrintername <> vbNullString Then
50700     SetDefaultprinterInProg DefaultPrintername
50710    End If
50720   End If
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
50010  Dim ctl As Control, ts As TabStrip, df As dmFrame, f As StdFont, trv As TreeView
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
50310   If TypeOf ctl Is TreeView Then
50320    Set trv = ctl
50330    trv.Font.Name = Fontname
50340    trv.Font.Size = Fontsize
50350    trv.Font.Charset = Charset
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
50010  Const PRINT_WAITFORCOMPLETION = 2
50020  Dim web As InternetExplorer
50030  Set web = New InternetExplorer
50040  web.Navigate2 URL
50050  Do Until web.ReadyState = READYSTATE_COMPLETE
50060   DoEvents
50070   If StopURLPrinting = True Then
50080    Exit Sub
50090   End If
50100  Loop
50110  DoEvents
50120  Sleep TimeBetweenLoadAndPrint
50130  DoEvents
50140  If (web.QueryStatusWB(OLECMDID_PRINT) And OLECMDF_ENABLED) = OLECMDF_ENABLED Then
50150   web.ExecWB OLECMDID.OLECMDID_PRINT, OLECMDEXECOPT.OLECMDEXECOPT_DONTPROMPTUSER, PRINT_WAITFORCOMPLETION
50160  End If
50170  web.Quit
50180  Set web = Nothing
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
