Attribute VB_Name = "modMain"
Option Explicit

Public InputFilename As String, ShowAnimationWindow As Boolean, _
 ComputerScreenResolution As Long

Private UnLoadFile As Boolean, ClearCacheDir As Boolean, _
 OutputFilename As String, InitSettings As Boolean, frmMainSUP As Long

Public Sub Main()
 If App.StartMode = vbSModeStandalone Or IsInIDE Then
   InstanceCounter = InstanceCounter + 1
   ProgramIsVisible = True
   StartProgram
  Else
   ProgramWindowState = 1
 End If
End Sub

Public Sub StartProgram(Optional Params As String)
 Dim Filename As String

 ' Reduce the working size of used memory
 Call SetProcessWorkingSetSize(GetCurrentProcess(), -1, -1)

 AnalyzeCommandlineParameters Params

 If CheckInstance Then
  CheckProgramInstances
 End If


 InitProgram

 IfLoggingWriteLogfile "PDFCreator Program Start"
 IfLoggingWriteLogfile "MyAppData: " & GetMyAppData
 IfLoggingWriteLogfile "PDFCreatorINIFile: " & PDFCreatorINIFile

 If UnLoadFile Then
  CreateUnloadFile
  InstanceCounter = InstanceCounter - 1
  Exit Sub
 End If

 If ClearCacheDir Then
  ClearCache
 End If

 PrintFiles

 If InitSettings Then
  SaveOptions Options
 End If

 If NoStart Then
  InstanceCounter = InstanceCounter - 1
  Exit Sub
 End If

 If ShowOnlyOptions Then
  frmOptions.Show vbModal
  InstanceCounter = InstanceCounter - 1
  Exit Sub
 End If

 If ShowOnlyLogfile Then
  frmLog.Show vbModal
  InstanceCounter = InstanceCounter - 1
  Exit Sub
 End If

 LoadGhostscriptDLL

 If PDFCreatorPrinter = False Then
  If FileExists(InputFilename) = True And LenB(OutputFilename) = 0 Then
    Filename = GetTempFile(CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory, "~PS")
    KillFile Filename
    FileCopy InputFilename, Filename
   Else
    ConvertPostscriptFile InputFilename, OutputFilename
  End If
 End If

 If ProgramIsRunning(PDFCreator_GUID) Then
   ' There is a local running instance
   If Not NoAbortIfRunning Then
    InstanceCounter = InstanceCounter - 1
    Exit Sub
   End If
  Else
 ' Create a mutex
   Set mutexLocal = New clsMutex
   mutexLocal.CreateMutex PDFCreator_GUID
   Set mutexGlobal = New clsMutex
   ' Check for a global running instance
   If mutexGlobal.CheckMutex("Global\" & PDFCreator_GUID) = False Then
    mutexGlobal.CreateMutex "Global\" & PDFCreator_GUID
   End If
 End If

 If IsWin9xMe = False And IsWinNT4 = False And IsWin2000 = False Then
  InitCommonControls
 End If

 Load frmMain
End Sub

Public Sub AnalyzeCommandlineParameters(Optional Params As String)
 Dim cSwitch As String
 If IsMissing(Params) = False Then
   CommandLine = Params & " " & VBA.Command$
  Else
   CommandLine = VBA.Command$
 End If
 If Len(CommandLine) > 0 Then
  If UCase$(CommandSwitch("NO", False)) = "ABORTIFRUNNING" Then
    NoAbortIfRunning = True
   Else
    NoAbortIfRunning = False
  End If
  If UCase$(CommandSwitch("NO", False)) = "PROCESSING" Then
    NoProcessing = True
   Else
    NoProcessing = False
  End If
  If UCase$(CommandSwitch("NO", False)) = "PROCESSINGATSTARTUP" Then
    NoProcessingAtStartup = True
   Else
    NoProcessingAtStartup = False
  End If
  If UCase$(CommandSwitch("L", False)) = "OG" Then
    enableSpecialLogging = True
   Else
    enableSpecialLogging = False
  End If
  If UCase$(CommandSwitch("SHOW", False)) = "ONLYOPTIONS" Then
    ShowOnlyOptions = True
    NoAbortIfRunning = True
    NoProcessing = True
   Else
    ShowOnlyOptions = False
  End If
  If UCase$(CommandSwitch("SHOW", False)) = "ONLYLOGFILE" Then
    ShowOnlyLogfile = True
    NoAbortIfRunning = True
    NoProcessing = True
   Else
    ShowOnlyLogfile = False
  End If
  If UCase$(CommandSwitch("P", False)) = "PDFCREATORPRINTER" Then
    PDFCreatorPrinter = True
   Else
    PDFCreatorPrinter = False
  End If
  ' Initialize unload running program
  If UCase$(CommandSwitch("UL", False)) = "TRUE" Then
    UnLoadFile = True
   Else
    UnLoadFile = False
  End If

  ' Clear the cache
  If UCase$(CommandSwitch("CLEAR", False)) = "CACHE" Then
    ClearCacheDir = True
   Else
    ClearCacheDir = False
  End If

  ' Init settings
  If UCase$(CommandSwitch("IN", False)) = "IT" Then
    InitSettings = True
   Else
    InitSettings = False
  End If

  ' Check running instance
  If UCase$(CommandSwitch("Check", False)) = "INSTANCE" Then
    CheckInstance = True
   Else
    CheckInstance = False
  End If

  PrintFilename = CommandSwitch("PF", True)
  InputFilename = CommandSwitch("IF", True)
  OutputFilename = CommandSwitch("OF", True)
  cSwitch = CommandSwitch("OPTIONSFILE", True)
  If LenB(cSwitch) > 0 Then
   If FileExists(cSwitch) = True Then
    Optionsfile = cSwitch
   End If
  End If
  If UCase$(CommandSwitch("NO", False)) = "START" Then
    NoStart = True
   Else
    NoStart = False
  End If
  If UCase$(CommandSwitch("NO", False)) = "PSCHECK" Then
    NoPSCheck = True
   Else
    NoPSCheck = False
  End If
 End If
End Sub

Private Sub InitProgram()
 Dim INIFilename As String, tStr As String

 ShellAndWaitingIsRunning = False
 ChangeDefaultprinter = False
 PrintSelectedJobs = False
 Restart = False
 SavePasswordsForThisSession = False
 ShowAnimationWindow = False


 PDFCreatorLogfilePath = CompletePath(GetTempPath) & "PDFCreator\"
 ErrPtnr.SetProgInfo App.EXEName & " " & GetProgramReleaseStr

 If UseINI Then
  INIFilename = App.EXEName & ".ini"
  If InstalledAsServer = True Then
    PDFCreatorINIFile = CompletePath(GetPDFCreatorApplicationPath) & INIFilename
   Else
    If DirExists(GetMyAppData) = True Then
      tStr = CompletePath(GetMyAppData) & "PDFCreator"
      If DirExists(tStr) = False Then
       MakePath tStr
      End If
      PDFCreatorINIFile = CompletePath(tStr) & INIFilename
     Else
      PDFCreatorINIFile = CompletePath(GetPDFCreatorApplicationPath) & INIFilename
    End If
  End If
  If LenB(Optionsfile) > 0 Then
   PDFCreatorINIFile = Optionsfile
  End If
 End If

 InitLanguagesStrings
 ReadLanguageFromOptions
 LanguagePath = CompletePath(GetPDFCreatorApplicationPath) & "Languages\"
 Languagefile = LanguagePath & Options.Language & ".ini"
 If UCase$(Options.Language) = "ESPANOL" And FileExists(Languagefile) = False And _
   FileExists(LanguagePath & "spanish.ini") = True Then
  Languagefile = LanguagePath & "spanish.ini"
  Options.Language = "spanish"
 End If
 If FileExists(Languagefile) = True Then
   LoadLanguage Languagefile
  Else
   MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & ">" & Languagefile & "<"
'   Options.Language = "english"
 End If

 Options = ReadOptions

 If IsWin9xMe = False Then
  Select Case Options.ProcessPriority
   Case 0: 'Idle
    SetProcessPriority Idle
   Case 1: 'Normal
    SetProcessPriority Normal
   Case 2: 'High
    SetProcessPriority High
   Case 3: 'Realtime
    SetProcessPriority RealTime
  End Select
 End If

 CreatePDFCreatorTempfolder
 ComputerScreenResolution = ScreenResolution
End Sub

Private Sub CreateUnloadFile()
 Dim fn As Long, tStr As String
 fn = FreeFile
 tStr = CompletePath(App.Path) & "Unload.tmp"
 If FileExists(tStr) = False Then
  Open tStr For Output As #fn
  Close #fn
 End If
End Sub

Public Sub PrintFile(Filename As String, Optional Frm As Form, Optional xpPgb As XP_ProgressBar, _
 Optional lblFilename As Label, Optional lblSize As Label, Optional lblCount As Label)
 Dim PDFCreatorPrintername As String, DefaultPrintername As String, _
  Files As Collection, i As Long, tStrf() As String, tFilename As String, _
  kB As Long, MB As Long, GB As Long, tStr As String
 kB = 1024: MB = kB * 1024: GB = MB * 1024
 If Len(Filename) > 0 Then
   If UCase$(Printer.DeviceName) <> UCase$(GetPDFCreatorPrintername) Then
    If Options.NoConfirmMessageSwitchingDefaultprinter = 0 Then
     If ChangeDefaultprinter = False Then
      frmSwitchDefaultprinter.Show vbModal
      If ChangeDefaultprinter = False Then
       Exit Sub
      End If
     End If
    End If
    PDFCreatorPrintername = GetPDFCreatorPrintername
    If LenB(PDFCreatorPrintername) = 0 Then
     MsgBox LanguageStrings.MessagesMsg26 & " [1]"
     Exit Sub
    End If
    DefaultPrintername = Printer.DeviceName
    SetDefaultprinterInProg PDFCreatorPrintername
   End If
   Set Files = GetFiles(Filename, "", SortedByName)
   If Files.Count > 0 Then
     DoEvents
     If Not Frm Is Nothing Then
      SetTopMost Frm, True, True
     End If
     For i = 1 To Files.Count
      tStrf = Split(Files(i), "|")
      SplitPath tStrf(1), , , tFilename
      If Not lblFilename Is Nothing Then
       lblFilename.Caption = LanguageStrings.ListFilename & ": " & tFilename
      End If
      If Not lblSize Is Nothing Then
       If CLng(tStrf(2)) > GB Then
         tStr = Format$(CDbl(tStrf(2)) / GB, "0.00 " & LanguageStrings.ListGBytes)
        Else
         If CLng(tStrf(2)) > MB Then
           tStr = Format$(CDbl(tStrf(2)) / MB, "0.00 " & LanguageStrings.ListMBytes)
          Else
           If CLng(tStrf(2)) > kB Then
             tStr = Format$(CDbl(tStrf(2)) / kB, "0.00 " & LanguageStrings.ListKBytes)
            Else
             tStr = Format$(tStrf(2), "0 " & LanguageStrings.ListBytes)
         End If
        End If
       End If
       lblSize.Caption = LanguageStrings.ListSize & ": " & tStr
      End If
      If Not xpPgb Is Nothing Then
       xpPgb.Value = i
      End If
      If Not lblCount Is Nothing Then
       lblCount.Caption = CStr(i) & " (" & CStr(Files.Count) & ")"
       lblCount.Left = (Frm.Width - lblCount.Width) / 2
      End If
      If CancelPrintfiles = True Then
       Exit For
      End If
      DoEvents
      ShellAndWait 0, "print", tStrf(1), "", tStrf(0), wHidden, WCTermination, 0, True
      DoEvents
     Next i
    Else
     MsgBox LanguageStrings.MessagesMsg14
   End If
   If DefaultPrintername <> vbNullString Then
    SetDefaultprinterInProg DefaultPrintername
   End If
  End If
End Sub

Private Sub LoadGhostscriptDLL()
 Dim gsvers As Collection, reg As clsRegistry, tsf() As String, tStr As String, _
  Path As String
 GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)

 If GsDllLoaded = 0 Then
   IfLoggingWriteLogfile ("Cannot load " & CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
   If FileExists(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll) = False Then
     IfLoggingWriteLogfile ("Found " & CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll & " = False")
    Else
     IfLoggingWriteLogfile ("Found " & CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll & " = True")
   End If
   Set gsvers = GetAllGhostscriptversions
   If gsvers.Count = 0 Then
     SetPrinterStop True
'     mnPrinter(0).Checked = True
     MsgBox LanguageStrings.MessagesMsg08
    Else
     Set reg = New clsRegistry: reg.hkey = HKEY_LOCAL_MACHINE
     If InStr(gsvers.Item(1), ":") Then
       reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
       Options.DirectoryGhostscriptBinaries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryBinaries"))
       Options.DirectoryGhostscriptLibraries = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryLibraries"))
       Options.DirectoryGhostscriptFonts = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
      Else
       If InStr(UCase$(gsvers.Item(1)), "AFPL") Then
        If InStr(gsvers.Item(1), " ") > 0 Then
         tsf = Split(gsvers.Item(1), " ")
         reg.KeyRoot = "SOFTWARE\AFPL Ghostscript\" & tsf(UBound(tsf))
         tStr = reg.GetRegistryValue("GS_DLL")
         SplitPath tStr, , Path
         Options.DirectoryGhostscriptBinaries = CompletePath(Path)
         Options.DirectoryGhostscriptFonts = CompletePath(reg.GetRegistryValue("GhostscriptDirectoryFonts"))
         Options.DirectoryGhostscriptLibraries = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
         Options.DirectoryGhostscriptFonts = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
         If tsf(UBound(tsf)) <> "8.00" Then
          Options.DirectoryGhostscriptResource = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "Resource"
         End If
        End If
       End If
       If InStr(UCase$(gsvers.Item(1)), "GNU") Then
        If InStr(gsvers.Item(1), " ") > 0 Then
         tsf = Split(gsvers.Item(1), " ")
         reg.KeyRoot = "SOFTWARE\GNU Ghostscript\" & tsf(UBound(tsf))
         tStr = reg.GetRegistryValue("GS_DLL")
         SplitPath tStr, , Path
         Options.DirectoryGhostscriptBinaries = CompletePath(Path)
         Options.DirectoryGhostscriptLibraries = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
         Options.DirectoryGhostscriptFonts = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
        End If
       End If
       If InStr(UCase$(gsvers.Item(1)), "GPL") Then
        If InStr(gsvers.Item(1), " ") > 0 Then
         tsf = Split(gsvers.Item(1), " ")
         reg.KeyRoot = "SOFTWARE\GPL Ghostscript\" & tsf(UBound(tsf))
         tStr = reg.GetRegistryValue("GS_DLL")
         SplitPath tStr, , Path
         Options.DirectoryGhostscriptBinaries = CompletePath(Path)
         Options.DirectoryGhostscriptLibraries = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "lib"
         Options.DirectoryGhostscriptFonts = CompletePath(Mid(Mid(Path, 1, InStrRev(Path, "\") - 1), 1, InStrRev(Mid(Path, 1, InStrRev(Path, "\") - 1), "\"))) & "fonts"
         Options.DirectoryGhostscriptResource = CompletePath(Mid(Path, 1, InStrRev(Path, "\") - 1)) & "Resource"
        End If
       End If
     End If
     Set reg = Nothing
     IfLoggingWriteLogfile ("Try to load alternative Ghostscript: " & CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
     If FileExists(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll) = False Then
       IfLoggingWriteLogfile ("Found alternative " & CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll & " = False")
      Else
       IfLoggingWriteLogfile ("Found alternative " & CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll & " = True")
     End If
     GsDllLoaded = LoadDLL(CompletePath(Options.DirectoryGhostscriptBinaries) & GsDll)
     If GsDllLoaded = 0 Then
       SetPrinterStop True
       MsgBox LanguageStrings.MessagesMsg08
      Else
       GSRevision = GetGhostscriptRevision
     End If
   End If
  Else
   GSRevision = GetGhostscriptRevision
 End If

 SecurityIsPossible = False
 If GSRevision.intRevision >= 814 Or FileExists(CompletePath(App.Path) & "pdfenc.exe") Then
  SecurityIsPossible = True
 End If
End Sub

Public Sub SetFont(Frm As Form, ByVal Fontname As String, ByVal Charset As Long, ByVal Fontsize As Long)
 Dim ctl As Control, eB As isExplorerBar

 If LenB(Trim$(Fontname)) = 0 Then
  Exit Sub
 End If

 For Each ctl In Frm.Controls
  If TypeOf ctl Is Label Or _
     TypeOf ctl Is Form Or _
     TypeOf ctl Is ComboBox Or _
     TypeOf ctl Is CheckBox Or _
     TypeOf ctl Is CommandButton Or _
     TypeOf ctl Is ListView Or _
     TypeOf ctl Is StatusBar Or _
     TypeOf ctl Is TextBox Or _
     TypeOf ctl Is Frame Then
   With ctl
    .Font = Fontname
    If Not (TypeOf ctl Is StatusBar) And Not (TypeOf ctl Is ListView) Then
     .Fontsize = Fontsize
    End If
    .Font.Charset = Charset
   End With
  End If
  If TypeOf ctl Is isExplorerBar Then
   Set eB = ctl
   eB.Font.Name = Fontname
   eB.Font.Size = Fontsize
   eB.Font.Charset = Charset
  End If
 Next ctl
End Sub

Public Sub PrintFiles()
 Dim Files As Collection
 If LenB(PrintFilename) > 0 Then
  Set Files = GetFiles(PrintFilename, "", SortedByName)
  If Files.Count > 0 Then
    Load frmPrintfiles
   Else
    MsgBox LanguageStrings.MessagesMsg14
  End If
  Set Files = Nothing
 End If
End Sub

Public Sub PrintFile2(PrintFilename As String)
 Dim Files As Collection
 If LenB(PrintFilename) > 0 Then
  Set Files = GetFiles(PrintFilename, "", SortedByName)
  If Files.Count > 0 Then
    Load frmPrintfiles
   Else
    MsgBox LanguageStrings.MessagesMsg14
  End If
  Set Files = Nothing
 End If
End Sub

