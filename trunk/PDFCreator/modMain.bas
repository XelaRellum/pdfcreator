Attribute VB_Name = "modMain"
Option Explicit

Public mutex As clsMutex, LanguagePath As String, Languagefile As String, _
 InputFilename As String, IFIsPS As Boolean

Private UnLoadFile As Boolean, ClearCacheDir As Boolean, NoStart As Boolean, _
 OutputFilename As String

Public Sub Main()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Files As Collection
50020
50030  AnalyzeCommandlineParameters
50040  InitProgram
50050
50060  IfLoggingWriteLogfile "PDFCreator Program Start"
50070  IfLoggingWriteLogfile "MyAppData: " & GetMyAppData
50080  IfLoggingWriteLogfile "PDFCreatorINIFile: " & PDFCreatorINIFile
50090
50100  If UnLoadFile Then
50110   CreateUnloadFile
50120   Exit Sub
50130  End If
50140
50150  If ClearCacheDir Then
50160   ClearCache
50170  End If
50180
50190  If LenB(PrintFilename) > 0 Then
50200 '  PrintFile PrintFilename
50210   Set Files = GetFiles(PrintFilename, "")
50220   If Files.Count > 0 Then
50230     frmPrintfiles.Show
50240    Else
50250     MsgBox LanguageStrings.MessagesMsg14
50260   End If
50270  End If
50280
50290  If NoStart Then
50300   Exit Sub
50310  End If
50320
50330  If ShowOnlyOptions Then
50340   frmOptions.Show vbModal
50350   Exit Sub
50360  End If
50370
50380  If ShowOnlyLogfile Then
50390   frmLog.Show vbModal
50400   Exit Sub
50410  End If
50420
50430  LoadGhostscriptDLL
50440  ConvertPostscriptFile InputFilename, OutputFilename
50450
50460  If ProgramIsRunning(PDFCreator_GUID) Then
50470    If Not NoAbortIfRunning Then
50480     Exit Sub
50490    End If
50500   Else
50510  ' Create a mutex
50520    Set mutex = New clsMutex
50530    mutex.CreateMutex PDFCreator_GUID
50540  End If
50550 ' Debug.Print TestConvertPS2PDF("", "")
50560 ' Debug.Print OptimizePDF("", "")
50570 ' Exit Sub
50580  frmMain.Show
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
50510   PrintFilename = CommandSwitch("PF", True)
50520   InputFilename = CommandSwitch("IF", True)
50530   OutputFilename = CommandSwitch("OF", True)
50540   cSwitch = CommandSwitch("OPTIONSFILE", True)
50550   If LenB(cSwitch) > 0 Then
50560    If FileExists(cSwitch) = True Then
50570     Optionsfile = cSwitch
50580    End If
50590   End If
50600   If UCase$(CommandSwitch("NO", False)) = "START" Then
50610     NoStart = True
50620    Else
50630     NoStart = False
50640   End If
50650  End If
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
50020  ChangeDefaultprinter = False
50030  PrintSelectedJobs = False
50040  Restart = False
50050  SavePasswordsForThisSession = False
50060  PDFCreatorLogfilePath = CompletePath(GetTempPath) & "PDFCreator\"
50070  ErrPtnr.SetProgInfo App.EXEName & " " & GetProgramReleaseStr
50080  INIFilename = App.EXEName & ".ini"
50090  If InstalledAsServer = True Then
50100    PDFCreatorINIFile = CompletePath(GetPDFCreatorApplicationPath) & INIFilename
50110   Else
50120    If DirExists(GetMyAppData) = True Then
50130      tStr = CompletePath(GetMyAppData) & "PDFCreator"
50140      If DirExists(tStr) = False Then
50150       MakePath tStr
50160      End If
50170      PDFCreatorINIFile = CompletePath(tStr) & INIFilename
50180     Else
50190      PDFCreatorINIFile = CompletePath(GetPDFCreatorApplicationPath) & INIFilename
50200    End If
50210  End If
50220
50230  If LenB(Optionsfile) > 0 Then
50240   PDFCreatorINIFile = Optionsfile
50250  End If
50260
50270  Options = ReadOptions
50280
50290  If IsWin9xMe = False Then
50301   Select Case Options.ProcessPriority
         Case 0: 'Idle
50320     SetProcessPriority Idle
50330    Case 1: 'Normal
50340     SetProcessPriority Normal
50350    Case 2: 'High
50360     SetProcessPriority High
50370    Case 3: 'Realtime
50380     SetProcessPriority RealTime
50390   End Select
50400  End If
50410
50420  InitLanguagesStrings
50430  LanguagePath = CompletePath(GetPDFCreatorApplicationPath) & "Languages\"
50440  Languagefile = LanguagePath & Options.Language & ".ini"
50450  If FileExists(Languagefile) = True Then
50460    LoadLanguage Languagefile
50470   Else
50480    MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & ">" & Languagefile & "<"
50490 '   Options.Language = "english"
50500  End If
50510  CreatePDFCreatorTempfolder
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

Public Sub PrintFile(Frm As Form, Filename As String, xpPgb As XP_ProgressBar, lblFilename As Label, lblSize As Label, lblCount As Label)
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
50230    Set Files = GetFiles(Filename, "")
50240    If Files.Count > 0 Then
50250      DoEvents
50260      SetTopMost Frm, True, True
50270      For i = 1 To Files.Count
50280       tStrf = Split(Files(i), "|")
50290       SplitPath tStrf(1), , , tFilename
50300       lblFilename.Caption = LanguageStrings.ListFilename & ": " & tFilename
50310       If CLng(tStrf(2)) > GB Then
50320         tStr = Format$(CDbl(tStrf(2)) / GB, "0.00 " & LanguageStrings.ListGBytes)
50330        Else
50340         If CLng(tStrf(2)) > MB Then
50350           tStr = Format$(CDbl(tStrf(2)) / MB, "0.00 " & LanguageStrings.ListMBytes)
50360          Else
50370           If CLng(tStrf(2)) > kB Then
50380             tStr = Format$(CDbl(tStrf(2)) / kB, "0.00 " & LanguageStrings.ListKBytes)
50390            Else
50400             tStr = Format$(tStrf(2), "0 " & LanguageStrings.ListBytes)
50410         End If
50420        End If
50430       End If
50440       lblSize.Caption = LanguageStrings.ListSize & ": " & tStr
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
50190      If InStr(gsvers.item(1), ":") Then
50200        reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50210        Options.DirectoryGhostscriptBinaries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryBinaries"))
50220        Options.DirectoryGhostscriptLibraries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryLibraries"))
50230        Options.DirectoryGhostscriptFonts = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
50240       Else
50250        If InStr(UCase$(gsvers.item(1)), "AFPL") Then
50260         If InStr(gsvers.item(1), " ") > 0 Then
50270          tsf = Split(gsvers.item(1), " ")
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
50400        If InStr(UCase$(gsvers.item(1)), "GNU") Then
50410         If InStr(gsvers.item(1), " ") > 0 Then
50420          tsf = Split(gsvers.item(1), " ")
50430          reg.KeyRoot = "SOFTWARE\GNU Ghostscript\" & tsf(UBound(tsf))
50440          tStr = reg.GetRegistryValue("GS_DLL")
50450          SplitPath tStr, , Path
50460          Options.DirectoryGhostscriptBinaries = CompletePath(Path)
50470          Options.DirectoryGhostscriptFonts = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
50480          Options.DirectoryGhostscriptLibraries = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50490          Options.DirectoryGhostscriptFonts = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50500         End If
50510        End If
50520        If InStr(UCase$(gsvers.item(1)), "GPL") Then
50530         If InStr(gsvers.item(1), " ") > 0 Then
50540          tsf = Split(gsvers.item(1), " ")
50550          reg.KeyRoot = "SOFTWARE\GPL Ghostscript\" & tsf(UBound(tsf))
50560          tStr = reg.GetRegistryValue("GS_DLL")
50570          SplitPath tStr, , Path
50580          Options.DirectoryGhostscriptBinaries = CompletePath(Path)
50590          Options.DirectoryGhostscriptFonts = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
50600          Options.DirectoryGhostscriptLibraries = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
50610          Options.DirectoryGhostscriptFonts = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
50620          Options.DirectoryGhostscriptResource = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "Resource"
50630         End If
50640        End If
50650      End If
50660      Set reg = Nothing
50670      IfLoggingWriteLogfile ("Try to load alternative Ghostscript: " & CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
50680      If FileExists(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll) = False Then
50690        IfLoggingWriteLogfile ("Found alternative " & CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll & " = False")
50700       Else
50710        IfLoggingWriteLogfile ("Found alternative " & CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll & " = True")
50720      End If
50730      GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
50740      If GsDllLoaded = 0 Then
50750        SetPrinterStop True
50760        MsgBox LanguageStrings.MessagesMsg08
50770       Else
50780        GSRevision = GetGhostscriptRevision
50790      End If
50800    End If
50810   Else
50820    GSRevision = GetGhostscriptRevision
50830  End If
50840
50850  SecurityIsPossible = False
50860  If GSRevision.intRevision >= 814 Or FileExists(CompletePath(App.Path) & "pdfenc.exe") Then
50870   SecurityIsPossible = True
50880  End If
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
