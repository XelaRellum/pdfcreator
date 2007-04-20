Attribute VB_Name = "modMain"
Option Explicit

Public InputFilename As String, ShowAnimationWindow As Boolean, _
 ComputerScreenResolution As Long

Private UnLoadFile As Boolean, ClearCacheDir As Boolean, _
 OutputFilename As String, InitSettings As Boolean, frmMainSUP As Long, _
 AddWindowsExplorerIntegration As Boolean, RemoveWindowsExplorerIntegration As Boolean

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
50010  Dim Filename As String
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
50350  If AddWindowsExplorerIntegration = True And RemoveWindowsExplorerIntegration = False Then
50360   AddExplorerIntegration
50370  End If
50380  If AddWindowsExplorerIntegration = False And RemoveWindowsExplorerIntegration = True Then
50390   RemoveExplorerIntegration
50400  End If
50410
50420  If ClearCacheDir Then
50430   ClearCache
50440  End If
50450
50460  PrintFiles
50470
50480  If InitSettings Then
50490   SaveOptions Options
50500  End If
50510
50520  If NoStart Then
50530   InstanceCounter = InstanceCounter - 1
50540   Exit Sub
50550  End If
50560
50570  If ShowOnlyOptions Then
50580   frmOptions.Show vbModal
50590   InstanceCounter = InstanceCounter - 1
50600   Exit Sub
50610  End If
50620
50630  If ShowOnlyLogfile Then
50640   frmLog.Show vbModal
50650   InstanceCounter = InstanceCounter - 1
50660   Exit Sub
50670  End If
50680
50690  LoadGhostscriptDLL
50700
50710  If PDFCreatorPrinter = False Then
50720   If FileExists(InputFilename) = True And LenB(OutputFilename) = 0 Then
50730     Filename = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory, "~PS")
50740     KillFile Filename
50750     FileCopy InputFilename, Filename
50760     If FileExists(InputFilename) And DeleteIF Then
50770      KillFile InputFilename
50780     End If
50790    Else
50800     ConvertPostscriptFile InputFilename, OutputFilename
50810     If FileExists(InputFilename) And DeleteIF Then
50820      KillFile InputFilename
50830     End If
50840     If FileExists(OutputFilename) And OpenOF Then
50850       OpenDocument OutputFilename
50860     End If
50870   End If
50880  End If
50890
50900  If ProgramIsRunning(PDFCreator_GUID) Then
50910    ' There is a local running instance
50920    If Not NoAbortIfRunning Then
50930     InstanceCounter = InstanceCounter - 1
50940     Exit Sub
50950    End If
50960   Else
50970  ' Create a mutex
50980    Set mutexLocal = New clsMutex
50990    mutexLocal.CreateMutex PDFCreator_GUID
51000    Set mutexGlobal = New clsMutex
51010    ' Check for a global running instance
51020    If mutexGlobal.CheckMutex("Global\" & PDFCreator_GUID) = False Then
51030     mutexGlobal.CreateMutex "Global\" & PDFCreator_GUID
51040    End If
51050  End If
51060
51070  If IsWin9xMe = False And IsWinNT4 = False And IsWin2000 = False Then
51080   InitCommonControls
51090  End If
51100
51110  Load frmMain
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
50070  If Len(CommandLine) > 0 Then
50080   If UCase$(CommandSwitch("NO", False)) = "ABORTIFRUNNING" Then
50090     NoAbortIfRunning = True
50100    Else
50110     NoAbortIfRunning = False
50120   End If
50130   If UCase$(CommandSwitch("NO", False)) = "PROCESSING" Then
50140     NoProcessing = True
50150    Else
50160     NoProcessing = False
50170   End If
50180   If UCase$(CommandSwitch("NO", False)) = "PROCESSINGATSTARTUP" Then
50190     NoProcessingAtStartup = True
50200    Else
50210     NoProcessingAtStartup = False
50220   End If
50230   If UCase$(CommandSwitch("L", False)) = "OG" Then
50240     enableSpecialLogging = True
50250    Else
50260     enableSpecialLogging = False
50270   End If
50280   If UCase$(CommandSwitch("SHOW", False)) = "ONLYOPTIONS" Then
50290     ShowOnlyOptions = True
50300     NoAbortIfRunning = True
50310     NoProcessing = True
50320    Else
50330     ShowOnlyOptions = False
50340   End If
50350   If UCase$(CommandSwitch("SHOW", False)) = "ONLYLOGFILE" Then
50360     ShowOnlyLogfile = True
50370     NoAbortIfRunning = True
50380     NoProcessing = True
50390    Else
50400     ShowOnlyLogfile = False
50410   End If
50420   If UCase$(CommandSwitch("P", False)) = "PDFCREATORPRINTER" Then
50430     PDFCreatorPrinter = True
50440    Else
50450     PDFCreatorPrinter = False
50460   End If
50470   ' Initialize unload running program
50480   If UCase$(CommandSwitch("UL", False)) = "TRUE" Then
50490     UnLoadFile = True
50500    Else
50510     UnLoadFile = False
50520   End If
50530
50540   ' Clear the cache
50550   If UCase$(CommandSwitch("CLEAR", False)) = "CACHE" Then
50560     ClearCacheDir = True
50570    Else
50580     ClearCacheDir = False
50590   End If
50600
50610   ' Init settings
50620   If UCase$(CommandSwitch("IN", False)) = "IT" Then
50630     InitSettings = True
50640    Else
50650     InitSettings = False
50660   End If
50670
50680   ' Windows-Explorer integration
50690   If UCase$(CommandSwitch("ADD", False)) = "WINDOWSEXPLORERINTEGRATION" Then
50700     AddWindowsExplorerIntegration = True
50710    Else
50720     AddWindowsExplorerIntegration = False
50730   End If
50740   If UCase$(CommandSwitch("REMOVE", False)) = "WINDOWSEXPLORERINTEGRATION" Then
50750     RemoveWindowsExplorerIntegration = True
50760    Else
50770     RemoveWindowsExplorerIntegration = False
50780   End If
50790
50800
50810   ' Check running instance
50820   If UCase$(CommandSwitch("Check", False)) = "INSTANCE" Then
50830     CheckInstance = True
50840    Else
50850     CheckInstance = False
50860   End If
50870
50880   PrintFilename = CommandSwitch("PF", True)
50890   InputFilename = CommandSwitch("IF", True)
50900   OutputFilename = CommandSwitch("OF", True)
50910
50920   ' Check delete inputfile after converting
50930   If UCase$(CommandSwitch("Delete", False)) = "IF" Then
50940     DeleteIF = True
50950    Else
50960     DeleteIF = False
50970   End If
50980
50990   ' Open the outputfile after converting
51000   If UCase$(CommandSwitch("Open", False)) = "OF" Then
51010     OpenOF = True
51020    Else
51030     OpenOF = False
51040   End If
51050
51060
51070   cSwitch = CommandSwitch("OPTIONSFILE", True)
51080   If LenB(cSwitch) > 0 Then
51090    If FileExists(cSwitch) = True Then
51100     Optionsfile = cSwitch
51110    End If
51120   End If
51130   If UCase$(CommandSwitch("NO", False)) = "START" Then
51140     NoStart = True
51150    Else
51160     NoStart = False
51170   End If
51180   If UCase$(CommandSwitch("NO", False)) = "PSCHECK" Then
51190     NoPSCheck = True
51200    Else
51210     NoPSCheck = False
51220   End If
51230  End If
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

Public Sub PrintFile(Filename As String, Optional Frm As Form, Optional xpPgb As XP_ProgressBar, _
 Optional lblFilename As Label, Optional lblSize As Label, Optional lblCount As Label)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim PDFCreatorPrintername As String, DefaultPrintername As String, _
  Files As Collection, i As Long, tStrf() As String, tFilename As String, _
  kB As Long, MB As Long, GB As Long, tStr As String
50040  kB = 1024: MB = kB * 1024: GB = MB * 1024
50050  If Len(Filename) > 0 Then
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
50230    Set Files = GetFiles(Filename, "", SortedByName)
50240    If Files.Count > 0 Then
50250      DoEvents
50260      If Not Frm Is Nothing Then
50270       SetTopMost Frm, True, True
50280      End If
50290      For i = 1 To Files.Count
50300       tStrf = Split(Files(i), "|")
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
50520        xpPgb.Value = i
50530       End If
50540       If Not lblCount Is Nothing Then
50550        lblCount.Caption = CStr(i) & " (" & CStr(Files.Count) & ")"
50560        lblCount.Left = (Frm.Width - lblCount.Width) / 2
50570       End If
50580       If CancelPrintfiles = True Then
50590        Exit For
50600       End If
50610       DoEvents
50620       ShellAndWait 0, "print", tStrf(1), "", tStrf(0), wHidden, WCTermination, 0, True
50630       DoEvents
50640      Next i
50650     Else
50660      MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & vbCrLf & "B: " & Filename
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

Public Sub SetFont(Frm As Form, ByVal Fontname As String, ByVal Charset As Long, ByVal Fontsize As Long)
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
50120  For Each ctl In Frm.Controls
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
50010  Dim Files As Collection
50020  If LenB(PrintFilename) > 0 Then
50030   Set Files = GetFiles(PrintFilename, "", SortedByName)
50040   If Files.Count > 0 Then
50050     Load frmPrintfiles
50060    Else
50070     MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & vbCrLf & "C: " & PrintFilename
50080   End If
50090   Set Files = Nothing
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
50010  Dim Files As Collection
50020  If LenB(PrintFilename) > 0 Then
50030   Set Files = GetFiles(PrintFilename, "", SortedByName)
50040   If Files.Count > 0 Then
50050     Load frmPrintfiles
50060    Else
50070     MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & vbCrLf & "D: " & PrintFilename
50080   End If
50090   Set Files = Nothing
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
50160  Set web = Nothing
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