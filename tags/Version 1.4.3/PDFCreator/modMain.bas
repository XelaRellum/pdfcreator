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
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  PDFCreatorApplicationPath = GetPDFCreatorApplicationPath
50020  InstalledAsServer = CheckInstalledAsServer
50030  pdfforgeDllInstalled = pdfforgeDllIsInstalled
50040  If App.StartMode = vbSModeStandalone Or IsInIDE Then
50050    InstanceCounter = InstanceCounter + 1
50060    ProgramIsVisible = True
50070    StartProgram
50080   Else
50090    ProgramWindowState = 1
50100  End If
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
50010  Dim filename As String, res As Boolean, Path As String, File As String, psFileName As String, _
  InfoSpoolFileName As String, strGUID As String, spoolDirectory As String
50030
50040  ' Reduce the working size of used memory
50050  Call SetProcessWorkingSetSize(GetCurrentProcess(), -1, -1)
50060
50070  AnalyzeCommandlineParameters Params
50080
50090  If DelayedStart > 0 Then
50100   Sleep DelayedStart
50110  End If
50120
50130  If CheckInstance Then
50140   CheckProgramInstances
50150  End If
50160
50170  InitProgram
50180
50190  IfLoggingWriteLogfile "PDFCreator Program Start"
50200  IfLoggingWriteLogfile "Windowsversion: " & GetWinVersionStr
50210  If InstalledAsServer Then
50220    IfLoggingWriteLogfile "InstalledAsServer: True"
50230   Else
50240    IfLoggingWriteLogfile "InstalledAsServer: False"
50250  End If
50260  IfLoggingWriteLogfile "MyAppData: " & GetMyAppData
50270
50280  spoolDirectory = GetPDFCreatorSpoolDirectory
50290
50300  If CheckFolderWriteAccess(spoolDirectory) Then
50310    IfLoggingWriteLogfile "PDFCreator spool directory (Write access: true): " & spoolDirectory
50320   Else
50330    IfLoggingWriteLogfile "PDFCreator spool directory (Write access: false): " & spoolDirectory
50340  End If
50350
50360  If UnLoadFile Then
50370   CreateUnloadFile
50380   InstanceCounter = InstanceCounter - 1
50390   Exit Sub
50400  End If
50410
50420  If Options.DisableUpdateCheck <> 1 Then
50430    CheckForUpdateAutomatically False, False, 2000
50440   Else
50450    IfLoggingWriteLogfile "Update check is disabled!"
50460  End If
50470
50480  If AddWindowsExplorerIntegration = True And RemoveWindowsExplorerIntegration = False Then
50490   AddExplorerIntegration
50500  End If
50510  If AddWindowsExplorerIntegration = False And RemoveWindowsExplorerIntegration = True Then
50520   RemoveExplorerIntegration
50530  End If
50540
50550  If ClearCacheDir Then
50560   ClearCache
50570  End If
50580
50590  If InitSettings Then
50600   SaveOptions Options 'Initialize the default settings
50610  End If
50620
50630  If LenB(Trim(SetupLogFile)) = 0 Then
50640   SetupLogFile = CompletePath(App.Path) & "SetupLog.txt"
50650  End If
50660
50670  If bUninstallPrinter Then
50680   If Not IsAdmin Then
50690    MsgBox LanguageStrings.PrintersAdminNotice
50700    Exit Sub
50710   End If
50720   If PrinterIsInstalled(UnInstallPrinterName) Then
50730     res = UnInstallPrinter(UnInstallPrinterName, "")
50740    Else
50750     IfLoggingWriteLogfile "Printer '" & UnInstallPrinterName & "' isn't installed!"
50760     If bNoMsg = True Then
50770      MsgBox "Printer '" & UnInstallPrinterName & "' isn't installed!", vbOKOnly + vbExclamation
50780     End If
50790   End If
50800  End If
50810  If bInstallPrinter Then
50820   If Not IsAdmin Then
50830    MsgBox LanguageStrings.PrintersAdminNotice
50840    Exit Sub
50850   End If
50860   If PrinterIsInstalled(InstallPrinterName) Then
50870     IfLoggingWriteLogfile "Printer '" & InstallPrinterName & "' is already installed!"
50880     If bNoMsg = True Then
50890      MsgBox "Printer '" & InstallPrinterName & "' is already installed!", vbOKOnly + vbExclamation
50900     End If
50910    Else
50920     res = InstallPrinter(InstallPrinterName, "PDFCreator", "PDFCreator:", "")
50930   End If
50940  End If
50950
50960  If bUninstallWindowsPrinter Then
50970   If PrinterIsInstalled(UnInstallPrinterName) Then
50980     Call UnInstallWindowsPrinter("pdfcmon", "pdfcmon", "PDFCreator", InstallPrinterName, SetupLogFile)
50990    Else
51000     IfLoggingWriteLogfile "Printer '" & UnInstallPrinterName & "' isn't installed!"
51010     If bNoMsg = False Then
51020      MsgBox "Printer '" & UnInstallPrinterName & "' isn't installed!", vbOKOnly + vbExclamation
51030     End If
51040   End If
51050  End If
51060  If bInstallWindowsPrinter Then
51070   If PrinterIsInstalled(InstallPrinterName) Then
51080     IfLoggingWriteLogfile "Printer '" & InstallPrinterName & "' is already installed!"
51090     If bNoMsg = False Then
51100      MsgBox "Printer '" & InstallPrinterName & "' is already installed!", vbOKOnly + vbExclamation
51110     End If
51120    Else
51130     Call InstallWindowsPrinter("pdfcmon", "pdfcmon", "PDFCreator", InstallPrinterName, SetupLogFile, App.Path)
51140   End If
51150  End If
51160
51170  PrintFiles
51180
51190  If ShowOnlyOptions Then
51200   frmOptions.Show vbModal
51210   InstanceCounter = InstanceCounter - 1
51220   Exit Sub
51230  End If
51240
51250  If ShowOnlyLogfile Then
51260   frmLog.Show vbModal
51270   InstanceCounter = InstanceCounter - 1
51280   Exit Sub
51290  End If
51300
51310  LoadGhostscriptDLL
51320
51330  If PDFCreatorPrinter = False Then
51340   If FileExists(InputFilename) = True And LenB(OutputFilename) = 0 Then
51350     strGUID = GetGUID
51360     psFileName = spoolDirectory & strGUID & ".ps"
51370     If IsValidGraphicFile(InputFilename) Then
51380       Call Image2PS(InputFilename, psFileName)
51390      Else
51400       FileCopy InputFilename, psFileName
51410     End If
51420     InfoSpoolFileName = CreateInfoSpoolFile(psFileName)
51430     If FileExists(InputFilename) And DeleteIF Then
51440      KillFile InputFilename
51450     End If
51460    Else
51470     If IsValidGraphicFile(InputFilename) Then
51480       strGUID = GetGUID
51490       psFileName = spoolDirectory & strGUID & ".ps"
51500       psFileName = CompletePath(Path) & File & ".ps"
51510       Call Image2PS(InputFilename, psFileName)
51520       ConvertFile psFileName, OutputFilename, OutputSubFormat
51530       If FileExists(psFileName) Then
51540        KillFile psFileName
51550       End If
51560      Else
51570       ConvertFile InputFilename, OutputFilename, OutputSubFormat
51580     End If
51590     If FileExists(InputFilename) And DeleteIF Then
51600      KillFile InputFilename
51610     End If
51620     If FileExists(OutputFilename) And OpenOF Then
51630      OpenDocument OutputFilename
51640     End If
51650   End If
51660  End If
51670
51680  If NoStart Then
51690   InstanceCounter = InstanceCounter - 1
51700   Exit Sub
51710  End If
51720
51730  If mutexLocal Is Nothing Then
51740   Set mutexLocal = New clsMutex
51750  End If
51760  If mutexGlobal Is Nothing Then
51770   Set mutexGlobal = New clsMutex
51780  End If
51790  If ProgramIsRunning(PDFCreator_GUID) Then
51800    ' There is a local running instance
51810    If Not NoAbortIfRunning Then
51820     InstanceCounter = InstanceCounter - 1
51830     Exit Sub
51840    End If
51850   Else
51860  ' Create a mutex
51870    mutexLocal.CreateMutex PDFCreator_GUID
51880    ' Check for a global running instance
51890    If mutexGlobal.CheckMutex("Global\" & PDFCreator_GUID) = False Then
51900     mutexGlobal.CreateMutex "Global\" & PDFCreator_GUID
51910    End If
51920  End If
51930
51940 If IsFrmMainLoaded Then
51950  Exit Sub
51960 End If
51970
51980
51990  If IsWin9xMe = False And IsWinNT4 = False And IsWin2000 = False Then
52000   InitCommonControls
52010  End If
52020
52030  Load frmMain
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
50090   ' Windows-Explorer integration
50100   If UCase$(CommandSwitch("ADD", False)) = "WINDOWSEXPLORERINTEGRATION" Then
50110     AddWindowsExplorerIntegration = True
50120    Else
50130     AddWindowsExplorerIntegration = False
50140   End If
50150   ' Check running instance
50160   If UCase$(CommandSwitch("Check", False)) = "INSTANCE" Then
50170     CheckInstance = True
50180    Else
50190     CheckInstance = False
50200   End If
50210   ' Clear the cache
50220   If UCase$(CommandSwitch("CLEAR", False)) = "CACHE" Then
50230     ClearCacheDir = True
50240    Else
50250     ClearCacheDir = False
50260   End If
50270   ' Delayed start in milliseconds
50280   cSwitch = CommandSwitch("DELAYEDSTART", True)
50290   If LenB(cSwitch) > 0 Then
50300    If IsNumeric(cSwitch) = True Then
50310     DelayedStart = CLng(cSwitch)
50320     If DelayedStart < 0 Then
50330      DelayedStart = 0
50340     End If
50350    End If
50360   End If
50370   ' Check delete inputfile after converting
50380   If UCase$(CommandSwitch("Delete", False)) = "IF" Then
50390     DeleteIF = True
50400    Else
50410     DeleteIF = False
50420   End If
50430   InputFilename = CommandSwitch("IF", True)
50440   ' Init settings
50450   If UCase$(CommandSwitch("INI", False)) = "T" Then
50460     InitSettings = True
50470    Else
50480     InitSettings = False
50490   End If
50500   ' InstallPrinter
50510   cSwitch = Trim$(CommandSwitch("INSTALLPRINTER", True))
50520   If LenB(cSwitch) > 0 Then
50530    bInstallPrinter = True
50540    InstallPrinterName = cSwitch
50550   End If
50560   ' InstallWindowsPrinter
50570   cSwitch = Trim$(CommandSwitch("INSTALLWINDOWSPRINTER", True))
50580   If LenB(cSwitch) > 0 Then
50590    bInstallWindowsPrinter = True
50600    InstallPrinterName = cSwitch
50610   End If
50620   If UCase$(CommandSwitch("L", False)) = "OG" Then
50630     enableSpecialLogging = True
50640    Else
50650     enableSpecialLogging = False
50660   End If
50670   If UCase$(CommandSwitch("NOAB", False)) = "ORTIFRUNNING" Then
50680     NoAbortIfRunning = True
50690    Else
50700     NoAbortIfRunning = False
50710   End If
50720   If UCase$(CommandSwitch("NOM", False)) = "SG" Then
50730     bNoMsg = True
50740    Else
50750     bNoMsg = False
50760   End If
50770   If UCase$(CommandSwitch("NOPR", False)) = "OCESSING" Then
50780     NoProcessing = True
50790    Else
50800     NoProcessing = False
50810   End If
50820   If UCase$(CommandSwitch("NOPROCESSING", False)) = "ATSTARTUP" Then
50830     NoProcessingAtStartup = True
50840    Else
50850     NoProcessingAtStartup = False
50860   End If
50870   If UCase$(CommandSwitch("NOPS", False)) = "CHECK" Then
50880     NoPSCheck = True
50890    Else
50900     NoPSCheck = False
50910   End If
50920   If UCase$(CommandSwitch("NOST", False)) = "ART" Then
50930     NoStart = True
50940    Else
50950     NoStart = False
50960   End If
50970   OutputFilename = CommandSwitch("OF", True)
50980   ' Open the outputfile after converting
50990   If UCase$(CommandSwitch("Open", False)) = "OF" Then
51000     OpenOF = True
51010    Else
51020     OpenOF = False
51030   End If
51040   cSwitch = CommandSwitch("OPTIONSFILE", True)
51050   If LenB(cSwitch) > 0 Then
51060    If FileExists(cSwitch) = True Then
51070     Optionsfile = cSwitch
51080    End If
51090   End If
51100   OutputSubFormat = CommandSwitch("OutputSubFormat", True)
51110   PrintFilename = CommandSwitch("PF", True)
51120   PrinterInfFilename = CommandSwitch("PIF", True)
51130   If UCase$(CommandSwitch("PP", False)) = "DFCREATORPRINTER" Then
51140     PDFCreatorPrinter = True
51150    Else
51160     PDFCreatorPrinter = False
51170   End If
51180   ' Windows-Explorer integration
51190   If UCase$(CommandSwitch("REMOVE", False)) = "WINDOWSEXPLORERINTEGRATION" Then
51200     RemoveWindowsExplorerIntegration = True
51210    Else
51220     RemoveWindowsExplorerIntegration = False
51230   End If
51240   ' SetupLogFile
51250   cSwitch = Trim$(CommandSwitch("SETUPLOGFILE", True))
51260   If LenB(cSwitch) > 0 Then
51270    SetupLogFile = cSwitch
51280   End If
51290   If UCase$(CommandSwitch("SHOWONLYO", False)) = "PTIONS" Then
51300     ShowOnlyOptions = True
51310     NoAbortIfRunning = True
51320     NoProcessing = True
51330    Else
51340     ShowOnlyOptions = False
51350   End If
51360   If UCase$(CommandSwitch("SHOWONLYL", False)) = "OGFILE" Then
51370     ShowOnlyLogfile = True
51380     NoAbortIfRunning = True
51390     NoProcessing = True
51400    Else
51410     ShowOnlyLogfile = False
51420   End If
51430   ' UnInstallPrinter
51440   cSwitch = Trim$(CommandSwitch("UNINSTALLPRINTER", True))
51450   If LenB(cSwitch) > 0 Then
51460    bUninstallPrinter = True
51470    UnInstallPrinterName = cSwitch
51480   End If
51490   ' UnInstallWindowsPrinter
51500   cSwitch = Trim$(CommandSwitch("UNINSTALLWINDOWSPRINTER", True))
51510   If LenB(cSwitch) > 0 Then
51520    bUninstallWindowsPrinter = True
51530    UnInstallPrinterName = cSwitch
51540   End If
51550   ' Initialize unload running program
51560   If UCase$(CommandSwitch("UL", False)) = "TRUE" Then
51570     UnLoadFile = True
51580    Else
51590     UnLoadFile = False
51600   End If
51610  End If
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
50110  PDFCreatorLogfilePath = CompletePath(GetTempPathApi) & "PDFCreator\"
50120  ErrPtnr.SetProgInfo App.EXEName & " " & GetProgramReleaseStr
50130
50140  InitLanguagesStrings
50150  ReadLanguageFromOptions
50160  LanguagePath = CompletePath(PDFCreatorApplicationPath) & "Languages\"
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
50380  PrinterTemppath = GetTempPathApi
50390  If LenB(Optionsfile) > 0 Then
50400   Options = ReadOptionsINI(Options, Optionsfile, False)
50410  End If
50420  If Options.Logging = 1 Then
50430    Logging = True
50440   Else
50450    Logging = False
50460  End If
50470
50480  If IsWin9xMe = False Then
50491   Select Case Options.ProcessPriority
         Case 0: 'Idle
50510     SetProcessPriority Idle
50520    Case 1: 'Normal
50530     SetProcessPriority Normal
50540    Case 2: 'High
50550     SetProcessPriority High
50560    Case 3: 'Realtime
50570     SetProcessPriority RealTime
50580   End Select
50590  End If
50600
50610  'CreatePDFCreatorTempfolder
50620  ComputerScreenResolution = ScreenResolution
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
50030  tStr = PDFCreatorApplicationPath & "Unload.tmp"
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
50060  If LenB(filename) = 0 Then
50070   MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & vbCrLf & "PrintFile: Empty filename!", vbExclamation
50080   Exit Sub
50090  End If
50100  If IsPDFCreatorPrinter(Printer.DeviceName) = False Then
50110   If Options.NoConfirmMessageSwitchingDefaultprinter = 0 Then
50120    If ChangeDefaultprinter = False Then
50130     frmSwitchDefaultprinter.Show vbModal
50140     If ChangeDefaultprinter = False Then
50150      Exit Sub
50160     End If
50170    End If
50180   End If
50190   If UCase$(Printer.DeviceName) <> UCase$(GetPDFCreatorPrintername) Then
50200    PDFCreatorPrintername = GetPDFCreatorPrintername
50210    If LenB(PDFCreatorPrintername) = 0 Then
50220     MsgBox LanguageStrings.MessagesMsg26 & " [1]"
50230     Exit Sub
50240    End If
50250    DefaultPrintername = Printer.DeviceName
50260    SetDefaultprinterInProg PDFCreatorPrintername
50270   End If
50280  End If
50290
50300  Set files = GetFiles(filename, "", SortedByName)
50310  If files.Count > 0 Then
50320    DoEvents
50330    If Not frm Is Nothing Then
50340     SetTopMost frm, True, True
50350    End If
50360    For i = 1 To files.Count
50370     tStrf = Split(files(i), "|")
50380     SplitPath tStrf(1), , , tFilename
50390     If Not lblFilename Is Nothing Then
50400      lblFilename.Caption = LanguageStrings.ListFilename & ": " & tFilename
50410     End If
50420     If Not lblSize Is Nothing Then
50430      If CLng(tStrf(2)) > GB Then
50440        tStr = Format$(CDbl(tStrf(2)) / GB, "0.00 " & LanguageStrings.ListGBytes)
50450       Else
50460        If CLng(tStrf(2)) > MB Then
50470          tStr = Format$(CDbl(tStrf(2)) / MB, "0.00 " & LanguageStrings.ListMBytes)
50480         Else
50490          If CLng(tStrf(2)) > kB Then
50500            tStr = Format$(CDbl(tStrf(2)) / kB, "0.00 " & LanguageStrings.ListKBytes)
50510           Else
50520            tStr = Format$(tStrf(2), "0 " & LanguageStrings.ListBytes)
50530          End If
50540        End If
50550      End If
50560      lblSize.Caption = LanguageStrings.ListSize & ": " & tStr
50570     End If
50580     If Not xpPgb Is Nothing Then
50590      xpPgb.value = i
50600     End If
50610     If Not lblCount Is Nothing Then
50620      lblCount.Caption = CStr(i) & " (" & CStr(files.Count) & ")"
50630      lblCount.Left = (frm.Width - lblCount.Width) / 2
50640     End If
50650     If CancelPrintfiles = True Then
50660      Exit For
50670     End If
50680     DoEvents
50690     ShellAndWait 0, "print", tStrf(1), "", tStrf(0), wHidden, WCTermination, 0, True
50700     DoEvents
50710    Next i
50720  End If
50730  If DefaultPrintername <> vbNullString Then
50740   SetDefaultprinterInProg DefaultPrintername
50750  End If
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
50840  If GSRevision.intRevision >= 814 Or FileExists(PDFCreatorApplicationPath & "pdfenc.exe") Then
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

Public Sub SetFontUserControl(UserControlObject As Variant, ByVal Fontname As String, ByVal Charset As Long, ByVal Fontsize As Long)
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
50120  For Each ctl In UserControlObject.object.Controls
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
Select Case ErrPtnr.OnError("modMain", "SetFontUserControl")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SetFontControls(ctls As Variant, ByVal Fontname As String, ByVal Charset As Long, ByVal Fontsize As Long)
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
50120  For Each ctl In ctls
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
50240     .Font.Italic = False
50250     If Not (TypeOf ctl Is StatusBar) And Not (TypeOf ctl Is ListView) Then
50260      .Fontsize = Fontsize
50270     End If
50280     .Font.Charset = Charset
50290    End With
50300   End If
50310
50320   If TypeOf ctl Is TreeView Then
50330    Set trv = ctl
50340    ctl.Font.Italic = False
50350    trv.Font.Name = Fontname
50360    trv.Font.Size = Fontsize
50370    trv.Font.Charset = Charset
50380   End If
50390   If TypeOf ctl Is TabStrip Then
50400    Set ts = ctl
50410    ctl.Font.Italic = False
50420    ts.Font.Name = Fontname
50430    ts.Font.Size = Fontsize
50440    ts.Font.Charset = Charset
50450   End If
50460   If TypeOf ctl Is dmFrame Then
50470    Set df = ctl
50480    df.Font.Italic = False
50490    df.Font.Name = Fontname
50500    df.Font.Size = Fontsize
50510    df.Font.Charset = Charset
50520    Set df.Font = f
50530   End If
50540  Next ctl
50550
50560  Set f = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modMain", "SetFontControls")
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
