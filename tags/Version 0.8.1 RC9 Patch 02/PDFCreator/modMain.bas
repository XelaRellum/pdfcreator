Attribute VB_Name = "modMain"
Option Explicit

Public InputFilename As String, ShowAnimationWindow As Boolean, _
 ComputerScreenResolution As Long

Private UnLoadFile As Boolean, ClearCacheDir As Boolean, _
 OutputFilename As String, InitSettings As Boolean, frmMainSUP As Long

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
50120
50130  InitProgram
50140
50150  IfLoggingWriteLogfile "PDFCreator Program Start"
50160  IfLoggingWriteLogfile "Windowsversion: " & GetWinVersionStr
50170  If UseINI Then
50180    IfLoggingWriteLogfile "UseINI: True"
50190    IfLoggingWriteLogfile "PDFCreatorINIFile: " & PDFCreatorINIFile
50200   Else
50210    IfLoggingWriteLogfile "UseINI: False"
50220  End If
50230  If InstalledAsServer Then
50240    IfLoggingWriteLogfile "InstalledAsServer: True"
50250   Else
50260    IfLoggingWriteLogfile "InstalledAsServer: False"
50270  End If
50280  IfLoggingWriteLogfile "MyAppData: " & GetMyAppData
50290
50300
50310  If UnLoadFile Then
50320   CreateUnloadFile
50330   InstanceCounter = InstanceCounter - 1
50340   Exit Sub
50350  End If
50360
50370  If ClearCacheDir Then
50380   ClearCache
50390  End If
50400
50410  PrintFiles
50420
50430  If InitSettings Then
50440   SaveOptions Options
50450  End If
50460
50470  If NoStart Then
50480   InstanceCounter = InstanceCounter - 1
50490   Exit Sub
50500  End If
50510
50520  If ShowOnlyOptions Then
50530   frmOptions.Show vbModal
50540   InstanceCounter = InstanceCounter - 1
50550   Exit Sub
50560  End If
50570
50580  If ShowOnlyLogfile Then
50590   frmLog.Show vbModal
50600   InstanceCounter = InstanceCounter - 1
50610   Exit Sub
50620  End If
50630
50640  LoadGhostscriptDLL
50650
50660  If PDFCreatorPrinter = False Then
50670   If FileExists(InputFilename) = True And LenB(OutputFilename) = 0 Then
50680     Filename = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory, "~PS")
50690     KillFile Filename
50700     FileCopy InputFilename, Filename
50710    Else
50720     ConvertPostscriptFile InputFilename, OutputFilename
50730   End If
50740  End If
50750
50760  If ProgramIsRunning(PDFCreator_GUID) Then
50770    ' There is a local running instance
50780    If Not NoAbortIfRunning Then
50790     InstanceCounter = InstanceCounter - 1
50800     Exit Sub
50810    End If
50820   Else
50830  ' Create a mutex
50840    Set mutexLocal = New clsMutex
50850    mutexLocal.CreateMutex PDFCreator_GUID
50860    Set mutexGlobal = New clsMutex
50870    ' Check for a global running instance
50880    If mutexGlobal.CheckMutex("Global\" & PDFCreator_GUID) = False Then
50890     mutexGlobal.CreateMutex "Global\" & PDFCreator_GUID
50900    End If
50910  End If
50920
50930  If IsWin9xMe = False And IsWinNT4 = False And IsWin2000 = False Then
50940   InitCommonControls
50950  End If
50960
50970  Load frmMain
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
50680   ' Check running instance
50690   If UCase$(CommandSwitch("Check", False)) = "INSTANCE" Then
50700     CheckInstance = True
50710    Else
50720     CheckInstance = False
50730   End If
50740
50750   PrintFilename = CommandSwitch("PF", True)
50760   InputFilename = CommandSwitch("IF", True)
50770   OutputFilename = CommandSwitch("OF", True)
50780   cSwitch = CommandSwitch("OPTIONSFILE", True)
50790   If LenB(cSwitch) > 0 Then
50800    If FileExists(cSwitch) = True Then
50810     Optionsfile = cSwitch
50820    End If
50830   End If
50840   If UCase$(CommandSwitch("NO", False)) = "START" Then
50850     NoStart = True
50860    Else
50870     NoStart = False
50880   End If
50890   If UCase$(CommandSwitch("NO", False)) = "PSCHECK" Then
50900     NoPSCheck = True
50910    Else
50920     NoPSCheck = False
50930   End If
50940  End If
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
50090
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
50370  Languagefile = LanguagePath & Options.Language & ".ini"
50380  If UCase$(Options.Language) = "ESPANOL" And FileExists(Languagefile) = False And _
   FileExists(LanguagePath & "spanish.ini") = True Then
50400   Languagefile = LanguagePath & "spanish.ini"
50410   Options.Language = "spanish"
50420  End If
50430  If FileExists(Languagefile) = True Then
50440    LoadLanguage Languagefile
50450   Else
50460    MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & ">" & Languagefile & "<"
50470 '   Options.Language = "english"
50480  End If
50490
50500  Options = ReadOptions
50510
50520  If IsWin9xMe = False Then
50531   Select Case Options.ProcessPriority
         Case 0: 'Idle
50550     SetProcessPriority Idle
50560    Case 1: 'Normal
50570     SetProcessPriority Normal
50580    Case 2: 'High
50590     SetProcessPriority High
50600    Case 3: 'Realtime
50610     SetProcessPriority RealTime
50620   End Select
50630  End If
50640
50650  CreatePDFCreatorTempfolder
50660  ComputerScreenResolution = ScreenResolution
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
50030  tStr = CompletePath(App.Path) & "Unload.tmp"
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
50660      MsgBox LanguageStrings.MessagesMsg14
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
50840  If GSRevision.intRevision >= 814 Or FileExists(CompletePath(App.Path) & "pdfenc.exe") Then
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
50010  Dim ctl As Control, eB As isExplorerBar
50020
50030  If LenB(Trim$(Fontname)) = 0 Then
50040   Exit Sub
50050  End If
50060
50070  For Each ctl In Frm.Controls
50080   If TypeOf ctl Is Label Or _
     TypeOf ctl Is Form Or _
     TypeOf ctl Is ComboBox Or _
     TypeOf ctl Is CheckBox Or _
     TypeOf ctl Is CommandButton Or _
     TypeOf ctl Is ListView Or _
     TypeOf ctl Is StatusBar Or _
     TypeOf ctl Is TextBox Or _
     TypeOf ctl Is Frame Then
50170    With ctl
50180     .Font = Fontname
50190     If Not (TypeOf ctl Is StatusBar) And Not (TypeOf ctl Is ListView) Then
50200      .Fontsize = Fontsize
50210     End If
50220     .Font.Charset = Charset
50230    End With
50240   End If
50250   If TypeOf ctl Is isExplorerBar Then
50260    Set eB = ctl
50270    eB.Font.Name = Fontname
50280    eB.Font.Size = Fontsize
50290    eB.Font.Charset = Charset
50300   End If
50310  Next ctl
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
50070     MsgBox LanguageStrings.MessagesMsg14
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
50070     MsgBox LanguageStrings.MessagesMsg14
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

