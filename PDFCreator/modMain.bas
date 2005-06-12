Attribute VB_Name = "modMain"
Option Explicit

Public mutex As clsMutex, LanguagePath As String, Languagefile As String, _
 InputFilename As String, IFIsPS As Boolean, ShowAnimationWindow As Boolean, _
 ComputerScreenResolution As Long

Private UnLoadFile As Boolean, ClearCacheDir As Boolean, NoStart As Boolean, _
 OutputFilename As String, InitSettings As Boolean

Public Sub Main()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Files As Collection, Filename As String
50020
50030  ' Reduce the working size of used memory
50040  Call SetProcessWorkingSetSize(GetCurrentProcess(), -1, -1)
50050
50060  AnalyzeCommandlineParameters
50070  InitProgram
50080
50090  IfLoggingWriteLogfile "PDFCreator Program Start"
50100  IfLoggingWriteLogfile "MyAppData: " & GetMyAppData
50110  IfLoggingWriteLogfile "PDFCreatorINIFile: " & PDFCreatorINIFile
50120
50130  If UnLoadFile Then
50140   CreateUnloadFile
50150   Exit Sub
50160  End If
50170
50180  If ClearCacheDir Then
50190   ClearCache
50200  End If
50210
50220  If LenB(PrintFilename) > 0 Then
50230   Set Files = GetFiles(PrintFilename, "", SortedByName)
50240   If Files.Count > 0 Then
50250     Load frmPrintfiles
50260    Else
50270     MsgBox LanguageStrings.MessagesMsg14
50280   End If
50290  End If
50300
50310  If InitSettings Then
50320   SaveOptions Options
50330  End If
50340
50350  If NoStart Then
50360   Exit Sub
50370  End If
50380
50390  If ShowOnlyOptions Then
50400   frmOptions.Show vbModal
50410   Exit Sub
50420  End If
50430
50440  If ShowOnlyLogfile Then
50450   frmLog.Show vbModal
50460   Exit Sub
50470  End If
50480
50490  LoadGhostscriptDLL
50500
50510  If PDFCreatorPrinter = False And FileExists(InputFilename) = True And LenB(OutputFilename) = 0 Then
50520    Filename = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory, "~PS")
50530    KillFile Filename
50540    FileCopy InputFilename, Filename
50550   Else
50560    ConvertPostscriptFile InputFilename, OutputFilename
50570  End If
50580
50590  If ProgramIsRunning(PDFCreator_GUID) Then
50600
50610     If Not NoAbortIfRunning Then
50620     Exit Sub
50630    End If
50640   Else
50650  ' Create a mutex
50660    Set mutex = New clsMutex
50670    mutex.CreateMutex PDFCreator_GUID
50680  End If
50690
50700  If IsWin9xMe = False And IsWinNT4 = False And IsWin2000 = False Then
50710   InitCommonControls
50720  End If
50730
50740  Load frmMain
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

Private Sub AnalyzeCommandlineParameters()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim cSwitch As String
50020  If Len(VBA.Command$) > 0 Then
50030   If UCase$(CommandSwitch("NO", False)) = "ABORTIFRUNNING" Then
50040     NoAbortIfRunning = True
50050    Else
50060     NoAbortIfRunning = False
50070   End If
50080   If UCase$(CommandSwitch("NO", False)) = "PROCESSING" Then
50090     NoProcessing = True
50100    Else
50110     NoProcessing = False
50120   End If
50130   If UCase$(CommandSwitch("L", False)) = "OG" Then
50140     enableSpecialLogging = True
50150    Else
50160     enableSpecialLogging = False
50170   End If
50180   If UCase$(CommandSwitch("SHOW", False)) = "ONLYOPTIONS" Then
50190     ShowOnlyOptions = True
50200     NoAbortIfRunning = True
50210     NoProcessing = True
50220    Else
50230     ShowOnlyOptions = False
50240   End If
50250   If UCase$(CommandSwitch("SHOW", False)) = "ONLYLOGFILE" Then
50260     ShowOnlyLogfile = True
50270     NoAbortIfRunning = True
50280     NoProcessing = True
50290    Else
50300     ShowOnlyLogfile = False
50310   End If
50320   If UCase$(CommandSwitch("P", False)) = "PDFCREATORPRINTER" Then
50330     PDFCreatorPrinter = True
50340    Else
50350     PDFCreatorPrinter = False
50360   End If
50370   ' Initialize unload running program
50380   If UCase$(CommandSwitch("UL", False)) = "TRUE" Then
50390     UnLoadFile = True
50400    Else
50410     UnLoadFile = False
50420   End If
50430
50440   ' Clear the cache
50450   If UCase$(CommandSwitch("CLEAR", False)) = "CACHE" Then
50460     ClearCacheDir = True
50470    Else
50480     ClearCacheDir = False
50490   End If
50500
50510   ' Init settings
50520   If UCase$(CommandSwitch("IN", False)) = "IT" Then
50530     InitSettings = True
50540    Else
50550     InitSettings = False
50560   End If
50570
50580
50590   PrintFilename = CommandSwitch("PF", True)
50600   InputFilename = CommandSwitch("IF", True)
50610   OutputFilename = CommandSwitch("OF", True)
50620   cSwitch = CommandSwitch("OPTIONSFILE", True)
50630   If LenB(cSwitch) > 0 Then
50640    If FileExists(cSwitch) = True Then
50650     Optionsfile = cSwitch
50660    End If
50670   End If
50680   If UCase$(CommandSwitch("NO", False)) = "START" Then
50690     NoStart = True
50700    Else
50710     NoStart = False
50720   End If
50730  End If
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
50010  Dim INIFilename As String, tstr As String
50020  ShellAndWaitingIsRunning = False
50030  ChangeDefaultprinter = False
50040  PrintSelectedJobs = False
50050  Restart = False
50060  SavePasswordsForThisSession = False
50070  ShowAnimationWindow = False
50080  PDFCreatorLogfilePath = CompletePath(GetTempPath) & "PDFCreator\"
50090  ErrPtnr.SetProgInfo App.EXEName & " " & GetProgramReleaseStr
50100  INIFilename = App.EXEName & ".ini"
50110  If InstalledAsServer = True Then
50120    PDFCreatorINIFile = CompletePath(GetPDFCreatorApplicationPath) & INIFilename
50130   Else
50140    If DirExists(GetMyAppData) = True Then
50150      tstr = CompletePath(GetMyAppData) & "PDFCreator"
50160      If DirExists(tstr) = False Then
50170       MakePath tstr
50180      End If
50190      PDFCreatorINIFile = CompletePath(tstr) & INIFilename
50200     Else
50210      PDFCreatorINIFile = CompletePath(GetPDFCreatorApplicationPath) & INIFilename
50220    End If
50230  End If
50240
50250  If LenB(Optionsfile) > 0 Then
50260   PDFCreatorINIFile = Optionsfile
50270  End If
50280
50290  Options = ReadOptions
50300
50310  If IsWin9xMe = False Then
50321   Select Case Options.ProcessPriority
         Case 0: 'Idle
50340     SetProcessPriority Idle
50350    Case 1: 'Normal
50360     SetProcessPriority Normal
50370    Case 2: 'High
50380     SetProcessPriority High
50390    Case 3: 'Realtime
50400     SetProcessPriority RealTime
50410   End Select
50420  End If
50430
50440  InitLanguagesStrings
50450  LanguagePath = CompletePath(GetPDFCreatorApplicationPath) & "Languages\"
50460  Languagefile = LanguagePath & Options.Language & ".ini"
50470  If UCase$(Options.Language) = "ESPANOL" And FileExists(Languagefile) = False And _
   FileExists(LanguagePath & "spanish.ini") = True Then
50490   Languagefile = LanguagePath & "spanish.ini"
50500   Options.Language = "spanish"
50510  End If
50520  If FileExists(Languagefile) = True Then
50530    LoadLanguage Languagefile
50540   Else
50550    MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & ">" & Languagefile & "<"
50560 '   Options.Language = "english"
50570  End If
50580  CreatePDFCreatorTempfolder
50590  ComputerScreenResolution = ScreenResolution
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
50010  Dim fn As Long, tstr As String
50020  fn = FreeFile
50030  tstr = CompletePath(App.Path) & "Unload.tmp"
50040  If FileExists(tstr) = False Then
50050   Open tstr For Output As #fn
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

Public Sub PrintFile(Frm As Form, Filename As String, xpPgb As XP_ProgressBar, lblFilename As Label, lblSize As Label, lblCount As Label)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim PDFCreatorPrintername As String, DefaultPrintername As String, _
  Files As Collection, i As Long, tStrf() As String, tFilename As String, _
  kB As Long, MB As Long, GB As Long, tstr As String
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
50260      SetTopMost Frm, True, True
50270      For i = 1 To Files.Count
50280       tStrf = Split(Files(i), "|")
50290       SplitPath tStrf(1), , , tFilename
50300       lblFilename.Caption = LanguageStrings.ListFilename & ": " & tFilename
50310       If CLng(tStrf(2)) > GB Then
50320         tstr = Format$(CDbl(tStrf(2)) / GB, "0.00 " & LanguageStrings.ListGBytes)
50330        Else
50340         If CLng(tStrf(2)) > MB Then
50350           tstr = Format$(CDbl(tStrf(2)) / MB, "0.00 " & LanguageStrings.ListMBytes)
50360          Else
50370           If CLng(tStrf(2)) > kB Then
50380             tstr = Format$(CDbl(tStrf(2)) / kB, "0.00 " & LanguageStrings.ListKBytes)
50390            Else
50400             tstr = Format$(tStrf(2), "0 " & LanguageStrings.ListBytes)
50410         End If
50420        End If
50430       End If
50440       lblSize.Caption = LanguageStrings.ListSize & ": " & tstr
50450       xpPgb.Value = i
50460       lblCount.Caption = CStr(i) & " (" & CStr(Files.Count) & ")"
50470       lblCount.Left = (Frm.Width - lblCount.Width) / 2
50480       If CancelPrintfiles = True Then
50490        Exit For
50500       End If
50510       DoEvents
50520       ShellAndWait 0, "print", tStrf(1), "", tStrf(0), wHidden, WCTermination, 0, True
50530       DoEvents
50540      Next i
50550     Else
50560      MsgBox LanguageStrings.MessagesMsg14
50570    End If
50580    If DefaultPrintername <> vbNullString Then
50590     SetDefaultprinterInProg DefaultPrintername
50600    End If
50610   End If
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
50010  Dim gsvers As Collection, reg As clsRegistry, tsf() As String, tstr As String, _
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
50290          tstr = reg.GetRegistryValue("GS_DLL")
50300          SplitPath tstr, , Path
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
50440          tstr = reg.GetRegistryValue("GS_DLL")
50450          SplitPath tstr, , Path
50460          Options.DirectoryGhostscriptBinaries = CompletePath(Path)
50470          Options.DirectoryGhostscriptLibraries = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50480          Options.DirectoryGhostscriptFonts = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50490         End If
50500        End If
50510        If InStr(UCase$(gsvers.Item(1)), "GPL") Then
50520         If InStr(gsvers.Item(1), " ") > 0 Then
50530          tsf = Split(gsvers.Item(1), " ")
50540          reg.KeyRoot = "SOFTWARE\GPL Ghostscript\" & tsf(UBound(tsf))
50550          tstr = reg.GetRegistryValue("GS_DLL")
50560          SplitPath tstr, , Path
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
50030  For Each ctl In Frm.Controls
50040   If TypeOf ctl Is Label Or _
     TypeOf ctl Is Form Or _
     TypeOf ctl Is ComboBox Or _
     TypeOf ctl Is CheckBox Or _
     TypeOf ctl Is CommandButton Or _
     TypeOf ctl Is ListView Or _
     TypeOf ctl Is StatusBar Or _
     TypeOf ctl Is TextBox Or _
     TypeOf ctl Is Frame Then
50130    With ctl
50140     .Font = Fontname
50150     If Not (TypeOf ctl Is StatusBar) And Not (TypeOf ctl Is ListView) Then
50160      .Fontsize = Fontsize
50170     End If
50180     .Font.Charset = Charset
50190    End With
50200   End If
50210   If TypeOf ctl Is isExplorerBar Then
50220    Set eB = ctl
50230    eB.Font.Name = Fontname
50240    eB.Font.Size = Fontsize
50250    eB.Font.Charset = Charset
50260   End If
50270  Next ctl
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
