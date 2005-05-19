Attribute VB_Name = "modMain"
Option Explicit

Public mutex As clsMutex, LanguagePath As String, Languagefile As String, _
 InputFilename As String, IFIsPS As Boolean, ShowAnimationWindow As Boolean, _
 ComputerScreenResolution As Long

Private UnLoadFile As Boolean, ClearCacheDir As Boolean, NoStart As Boolean, _
 OutputFilename As String, InitSettings As Boolean

Public Sub Main()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim Files As Collection, Filename As String
50050
50060  ' Reduce the working size of used memory
50070  Call SetProcessWorkingSetSize(GetCurrentProcess(), -1, -1)
50080
50090  AnalyzeCommandlineParameters
50100  InitProgram
50110
50120  IfLoggingWriteLogfile "PDFCreator Program Start"
50130  IfLoggingWriteLogfile "MyAppData: " & GetMyAppData
50140  IfLoggingWriteLogfile "PDFCreatorINIFile: " & PDFCreatorINIFile
50150
50160  If UnLoadFile Then
50170   CreateUnloadFile
50180   Exit Sub
50190  End If
50200
50210  If ClearCacheDir Then
50220   ClearCache
50230  End If
50240
50250  If LenB(PrintFilename) > 0 Then
50260   Set Files = GetFiles(PrintFilename, "", SortedByName)
50270   If Files.Count > 0 Then
50280     Load frmPrintfiles
50290    Else
50300     MsgBox LanguageStrings.MessagesMsg14
50310   End If
50320  End If
50330
50340  If InitSettings Then
50350   SaveOptions Options
50360  End If
50370
50380  If NoStart Then
50390   Exit Sub
50400  End If
50410
50420  If ShowOnlyOptions Then
50430   frmOptions.Show vbModal
50440   Exit Sub
50450  End If
50460
50470  If ShowOnlyLogfile Then
50480   frmLog.Show vbModal
50490   Exit Sub
50500  End If
50510
50520  LoadGhostscriptDLL
50530
50540  If PDFCreatorPrinter = False And FileExists(InputFilename) = True And LenB(OutputFilename) = 0 Then
50550    Filename = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory, "~PS")
50560    KillFile Filename
50570    FileCopy InputFilename, Filename
50580   Else
50590    ConvertPostscriptFile InputFilename, OutputFilename
50600  End If
50610
50620  If ProgramIsRunning(PDFCreator_GUID) Then
50630    If Not NoAbortIfRunning Then
50640     Exit Sub
50650    End If
50660   Else
50670  ' Create a mutex
50680    Set mutex = New clsMutex
50690    mutex.CreateMutex PDFCreator_GUID
50700  End If
50710
50720  If IsWin9xMe = False And IsWinNT4 = False And IsWin2000 = False Then
50730   InitCommonControls
50740  End If
50750
50760  Load frmMain
50770 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50780 Exit Sub
ErrPtnr_OnError:
50801 Select Case ErrPtnr.OnError("modMain", "Main")
      Case 0: Resume
50820 Case 1: Resume Next
50830 Case 2: Exit Sub
50840 Case 3: End
50850 End Select
50860 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub AnalyzeCommandlineParameters()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim cSwitch As String
50050  If Len(VBA.Command$) > 0 Then
50060   If UCase$(CommandSwitch("NO", False)) = "ABORTIFRUNNING" Then
50070     NoAbortIfRunning = True
50080    Else
50090     NoAbortIfRunning = False
50100   End If
50110   If UCase$(CommandSwitch("NO", False)) = "PROCESSING" Then
50120     NoProcessing = True
50130    Else
50140     NoProcessing = False
50150   End If
50160   If UCase$(CommandSwitch("L", False)) = "OG" Then
50170     enableSpecialLogging = True
50180    Else
50190     enableSpecialLogging = False
50200   End If
50210   If UCase$(CommandSwitch("SHOW", False)) = "ONLYOPTIONS" Then
50220     ShowOnlyOptions = True
50230     NoAbortIfRunning = True
50240     NoProcessing = True
50250    Else
50260     ShowOnlyOptions = False
50270   End If
50280   If UCase$(CommandSwitch("SHOW", False)) = "ONLYLOGFILE" Then
50290     ShowOnlyLogfile = True
50300     NoAbortIfRunning = True
50310     NoProcessing = True
50320    Else
50330     ShowOnlyLogfile = False
50340   End If
50350   If UCase$(CommandSwitch("P", False)) = "PDFCREATORPRINTER" Then
50360     PDFCreatorPrinter = True
50370    Else
50380     PDFCreatorPrinter = False
50390   End If
50400   ' Initialize unload running program
50410   If UCase$(CommandSwitch("UL", False)) = "TRUE" Then
50420     UnLoadFile = True
50430    Else
50440     UnLoadFile = False
50450   End If
50460
50470   ' Clear the cache
50480   If UCase$(CommandSwitch("CLEAR", False)) = "CACHE" Then
50490     ClearCacheDir = True
50500    Else
50510     ClearCacheDir = False
50520   End If
50530
50540   ' Init settings
50550   If UCase$(CommandSwitch("IN", False)) = "IT" Then
50560     InitSettings = True
50570    Else
50580     InitSettings = False
50590   End If
50600
50610
50620   PrintFilename = CommandSwitch("PF", True)
50630   InputFilename = CommandSwitch("IF", True)
50640   OutputFilename = CommandSwitch("OF", True)
50650   cSwitch = CommandSwitch("OPTIONSFILE", True)
50660   If LenB(cSwitch) > 0 Then
50670    If FileExists(cSwitch) = True Then
50680     Optionsfile = cSwitch
50690    End If
50700   End If
50710   If UCase$(CommandSwitch("NO", False)) = "START" Then
50720     NoStart = True
50730    Else
50740     NoStart = False
50750   End If
50760  End If
50770 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50780 Exit Sub
ErrPtnr_OnError:
50801 Select Case ErrPtnr.OnError("modMain", "AnalyzeCommandlineParameters")
      Case 0: Resume
50820 Case 1: Resume Next
50830 Case 2: Exit Sub
50840 Case 3: End
50850 End Select
50860 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub InitProgram()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim INIFilename As String, tstr As String
50050  ShellAndWaitingIsRunning = False
50060  ChangeDefaultprinter = False
50070  PrintSelectedJobs = False
50080  Restart = False
50090  SavePasswordsForThisSession = False
50100  ShowAnimationWindow = False
50110  PDFCreatorLogfilePath = CompletePath(GetTempPath) & "PDFCreator\"
50120  ErrPtnr.SetProgInfo App.EXEName & " " & GetProgramReleaseStr
50130  INIFilename = App.EXEName & ".ini"
50140  If InstalledAsServer = True Then
50150    PDFCreatorINIFile = CompletePath(GetPDFCreatorApplicationPath) & INIFilename
50160   Else
50170    If DirExists(GetMyAppData) = True Then
50180      tstr = CompletePath(GetMyAppData) & "PDFCreator"
50190      If DirExists(tstr) = False Then
50200       MakePath tstr
50210      End If
50220      PDFCreatorINIFile = CompletePath(tstr) & INIFilename
50230     Else
50240      PDFCreatorINIFile = CompletePath(GetPDFCreatorApplicationPath) & INIFilename
50250    End If
50260  End If
50270
50280  If LenB(Optionsfile) > 0 Then
50290   PDFCreatorINIFile = Optionsfile
50300  End If
50310
50320  Options = ReadOptions
50330
50340  If IsWin9xMe = False Then
50351   Select Case Options.ProcessPriority
         Case 0: 'Idle
50370     SetProcessPriority Idle
50380    Case 1: 'Normal
50390     SetProcessPriority Normal
50400    Case 2: 'High
50410     SetProcessPriority High
50420    Case 3: 'Realtime
50430     SetProcessPriority RealTime
50440   End Select
50450  End If
50460
50470  InitLanguagesStrings
50480  LanguagePath = CompletePath(GetPDFCreatorApplicationPath) & "Languages\"
50490  Languagefile = LanguagePath & Options.Language & ".ini"
50500  If UCase$(Options.Language) = "ESPANOL" And FileExists(Languagefile) = False And _
   FileExists(LanguagePath & "spanish.ini") = True Then
50520   Languagefile = LanguagePath & "spanish.ini"
50530   Options.Language = "spanish"
50540  End If
50550  If FileExists(Languagefile) = True Then
50560    LoadLanguage Languagefile
50570   Else
50580    MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & ">" & Languagefile & "<"
50590 '   Options.Language = "english"
50600  End If
50610  CreatePDFCreatorTempfolder
50620  ComputerScreenResolution = ScreenResolution
50630 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50640 Exit Sub
ErrPtnr_OnError:
50661 Select Case ErrPtnr.OnError("modMain", "InitProgram")
      Case 0: Resume
50680 Case 1: Resume Next
50690 Case 2: Exit Sub
50700 Case 3: End
50710 End Select
50720 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub CreateUnloadFile()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim fn As Long, tstr As String
50050  fn = FreeFile
50060  tstr = CompletePath(App.Path) & "Unload.tmp"
50070  If FileExists(tstr) = False Then
50080   Open tstr For Output As #fn
50090   Close #fn
50100  End If
50110 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50120 Exit Sub
ErrPtnr_OnError:
50141 Select Case ErrPtnr.OnError("modMain", "CreateUnloadFile")
      Case 0: Resume
50160 Case 1: Resume Next
50170 Case 2: Exit Sub
50180 Case 3: End
50190 End Select
50200 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub PrintFile(Frm As Form, Filename As String, xpPgb As XP_ProgressBar, lblFilename As Label, lblSize As Label, lblCount As Label)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim PDFCreatorPrintername As String, DefaultPrintername As String, _
  Files As Collection, i As Long, tStrf() As String, tFilename As String, _
  kB As Long, MB As Long, GB As Long, tstr As String
50070  kB = 1024: MB = kB * 1024: GB = MB * 1024
50080  If Len(Filename) > 0 Then
50090    If UCase$(Printer.DeviceName) <> UCase$(GetPDFCreatorPrintername) Then
50100     If Options.NoConfirmMessageSwitchingDefaultprinter = 0 Then
50110      If ChangeDefaultprinter = False Then
50120       frmSwitchDefaultprinter.Show vbModal
50130       If ChangeDefaultprinter = False Then
50140        Exit Sub
50150       End If
50160      End If
50170     End If
50180     PDFCreatorPrintername = GetPDFCreatorPrintername
50190     If LenB(PDFCreatorPrintername) = 0 Then
50200      MsgBox LanguageStrings.MessagesMsg26 & " [1]"
50210      Exit Sub
50220     End If
50230     DefaultPrintername = Printer.DeviceName
50240     SetDefaultprinterInProg PDFCreatorPrintername
50250    End If
50260    Set Files = GetFiles(Filename, "", SortedByName)
50270    If Files.Count > 0 Then
50280      DoEvents
50290      SetTopMost Frm, True, True
50300      For i = 1 To Files.Count
50310       tStrf = Split(Files(i), "|")
50320       SplitPath tStrf(1), , , tFilename
50330       lblFilename.Caption = LanguageStrings.ListFilename & ": " & tFilename
50340       If CLng(tStrf(2)) > GB Then
50350         tstr = Format$(CDbl(tStrf(2)) / GB, "0.00 " & LanguageStrings.ListGBytes)
50360        Else
50370         If CLng(tStrf(2)) > MB Then
50380           tstr = Format$(CDbl(tStrf(2)) / MB, "0.00 " & LanguageStrings.ListMBytes)
50390          Else
50400           If CLng(tStrf(2)) > kB Then
50410             tstr = Format$(CDbl(tStrf(2)) / kB, "0.00 " & LanguageStrings.ListKBytes)
50420            Else
50430             tstr = Format$(tStrf(2), "0 " & LanguageStrings.ListBytes)
50440         End If
50450        End If
50460       End If
50470       lblSize.Caption = LanguageStrings.ListSize & ": " & tstr
50480       xpPgb.Value = i
50490       lblCount.Caption = CStr(i) & " (" & CStr(Files.Count) & ")"
50500       lblCount.Left = (Frm.Width - lblCount.Width) / 2
50510       If CancelPrintfiles = True Then
50520        Exit For
50530       End If
50540       DoEvents
50550       ShellAndWait 0, "print", tStrf(1), "", tStrf(0), wHidden, WCTermination, 0, True
50560       DoEvents
50570      Next i
50580     Else
50590      MsgBox LanguageStrings.MessagesMsg14
50600    End If
50610    If DefaultPrintername <> vbNullString Then
50620     SetDefaultprinterInProg DefaultPrintername
50630    End If
50640   End If
50650 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50660 Exit Sub
ErrPtnr_OnError:
50681 Select Case ErrPtnr.OnError("modMain", "PrintFile")
      Case 0: Resume
50700 Case 1: Resume Next
50710 Case 2: Exit Sub
50720 Case 3: End
50730 End Select
50740 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
 End Sub

Private Sub LoadGhostscriptDLL()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim gsvers As Collection, reg As clsRegistry, tsf() As String, tstr As String, _
  Path As String
50060  GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
50070
50080  If GsDllLoaded = 0 Then
50090    IfLoggingWriteLogfile ("Cannot load " & CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
50100    If FileExists(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll) = False Then
50110      IfLoggingWriteLogfile ("Found " & CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll & " = False")
50120     Else
50130      IfLoggingWriteLogfile ("Found " & CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll & " = True")
50140    End If
50150    Set gsvers = GetAllGhostscriptversions
50160    If gsvers.Count = 0 Then
50170      SetPrinterStop True
50180 '     mnPrinter(0).Checked = True
50190      MsgBox LanguageStrings.MessagesMsg08
50200     Else
50210      Set reg = New clsRegistry: reg.hkey = HKEY_LOCAL_MACHINE
50220      If InStr(gsvers.Item(1), ":") Then
50230        reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50240        Options.DirectoryGhostscriptBinaries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryBinaries"))
50250        Options.DirectoryGhostscriptLibraries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryLibraries"))
50260        Options.DirectoryGhostscriptFonts = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
50270       Else
50280        If InStr(UCase$(gsvers.Item(1)), "AFPL") Then
50290         If InStr(gsvers.Item(1), " ") > 0 Then
50300          tsf = Split(gsvers.Item(1), " ")
50310          reg.KeyRoot = "SOFTWARE\AFPL Ghostscript\" & tsf(UBound(tsf))
50320          tstr = reg.GetRegistryValue("GS_DLL")
50330          SplitPath tstr, , Path
50340          Options.DirectoryGhostscriptBinaries = CompletePath(Path)
50350          Options.DirectoryGhostscriptFonts = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
50360          Options.DirectoryGhostscriptLibraries = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50370          Options.DirectoryGhostscriptFonts = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50380          If tsf(UBound(tsf)) <> "8.00" Then
50390           Options.DirectoryGhostscriptResource = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "Resource"
50400          End If
50410         End If
50420        End If
50430        If InStr(UCase$(gsvers.Item(1)), "GNU") Then
50440         If InStr(gsvers.Item(1), " ") > 0 Then
50450          tsf = Split(gsvers.Item(1), " ")
50460          reg.KeyRoot = "SOFTWARE\GNU Ghostscript\" & tsf(UBound(tsf))
50470          tstr = reg.GetRegistryValue("GS_DLL")
50480          SplitPath tstr, , Path
50490          Options.DirectoryGhostscriptBinaries = CompletePath(Path)
50500          Options.DirectoryGhostscriptLibraries = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50510          Options.DirectoryGhostscriptFonts = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50520         End If
50530        End If
50540        If InStr(UCase$(gsvers.Item(1)), "GPL") Then
50550         If InStr(gsvers.Item(1), " ") > 0 Then
50560          tsf = Split(gsvers.Item(1), " ")
50570          reg.KeyRoot = "SOFTWARE\GPL Ghostscript\" & tsf(UBound(tsf))
50580          tstr = reg.GetRegistryValue("GS_DLL")
50590          SplitPath tstr, , Path
50600          Options.DirectoryGhostscriptBinaries = CompletePath(Path)
50610          Options.DirectoryGhostscriptLibraries = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50620          Options.DirectoryGhostscriptFonts = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50630          Options.DirectoryGhostscriptResource = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "Resource"
50640         End If
50650        End If
50660      End If
50670      Set reg = Nothing
50680      IfLoggingWriteLogfile ("Try to load alternative Ghostscript: " & CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
50690      If FileExists(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll) = False Then
50700        IfLoggingWriteLogfile ("Found alternative " & CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll & " = False")
50710       Else
50720        IfLoggingWriteLogfile ("Found alternative " & CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll & " = True")
50730      End If
50740      GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
50750      If GsDllLoaded = 0 Then
50760        SetPrinterStop True
50770        MsgBox LanguageStrings.MessagesMsg08
50780       Else
50790        GSRevision = GetGhostscriptRevision
50800      End If
50810    End If
50820   Else
50830    GSRevision = GetGhostscriptRevision
50840  End If
50850
50860  SecurityIsPossible = False
50870  If GSRevision.intRevision >= 814 Or FileExists(CompletePath(App.Path) & "pdfenc.exe") Then
50880   SecurityIsPossible = True
50890  End If
50900 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50910 Exit Sub
ErrPtnr_OnError:
50931 Select Case ErrPtnr.OnError("modMain", "LoadGhostscriptDLL")
      Case 0: Resume
50950 Case 1: Resume Next
50960 Case 2: Exit Sub
50970 Case 3: End
50980 End Select
50990 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
