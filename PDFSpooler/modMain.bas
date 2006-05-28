Attribute VB_Name = "modMain"
Option Explicit

Public PDFCreatorPath As String

Public Sub Main()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim UserName As String, SessionID As Long, _
  UserAppData As String, UserLocalTemp As String, tStr As String, _
  hToken As Long, hProfile As Long, AppPath As String, AppParams As String, _
  stdio As clsStdIO, cinStr As String, SpooltimeSeconds As Double, _
  LoggedInConsole As Boolean, Tempfile As String
50060
50070  ' Reduce the working size of used memory
50080  Call SetProcessWorkingSetSize(GetCurrentProcess(), -1, -1)
50090
50100  InitProgram
50110  AnalyzeCommandlineParameters
50120
50130  If CheckInstance Then
50140   CheckProgramInstances
50150  End If
50160
50170  If NoStart Then
50180   Exit Sub
50190  End If
50200
50210 ' Create a local and a global mutex
50220  CreateMutexPDFSpooler
50230
50240  PDFCreatorPath = CompletePath(GetPDFCreatorApplicationPath)
50250  Call WriteInfoToSpecialLogfile
50260
50270  LoggedInConsole = False
50280  UserName = Environ$("Redmon_User")
50290
50300  If Len(Environ$("REDMON_SESSIONID")) > 0 Then
50310   SessionID = Environ$("REDMON_SESSIONID")
50320  End If
50330
50340  hToken = 0: hProfile = 0
50350  Call GetUserLocalDirs(HKEY_CURRENT_USER, UserAppData, UserLocalTemp)
50360  If IsWin9xMe = False Then
50370   If GetUserSessionToken(UserName, SessionID, hToken, enableSpecialLogging) = 0 Then
50380    If LoadProfile(UserName, hToken, hProfile) = 0 Then
50390     Call GetUserLocalDirs(hProfile, UserAppData, UserLocalTemp)
50400     LoggedInConsole = True
50410    End If
50420   End If
50430  End If
50440  WriteToSpecialLogfile "UserAppData=" & UserAppData
50450  WriteToSpecialLogfile "UserLocalTemp=" & UserLocalTemp
50460
50470  If InstalledAsServer Then
50480    PDFCreatorINIFile = CompletePath(GetPDFCreatorApplicationPath) & "PDFCreator.ini"
50490   Else
50500    PDFCreatorINIFile = CompletePath(UserAppData) & "PDFcreator\PDFCreator.ini"
50510  End If
50520  If LenB(Optionsfile) > 0 Then
50530   PDFCreatorINIFile = Optionsfile
50540  End If
50550
50560  WriteToSpecialLogfile "InstalledAsServer=" & InstalledAsServer
50570  WriteToSpecialLogfile "UseINI=" & UseINI
50580  WriteToSpecialLogfile "PDFCreatorINIFile=" & PDFCreatorINIFile
50590
50600  If IsWin9xMe Then
50610    Options = ReadOptions(True)
50620   Else
50630    WriteToSpecialLogfile "Read options: start"
50640    Options = ReadOptions(True, hProfile)
50650    WriteToSpecialLogfile "Read options: ready"
50660  End If
50670
50680  WriteToSpecialLogfile ""
50690  WriteToSpecialLogfile "Options.DirectoryGhostscriptBinaries=" & Options.DirectoryGhostscriptBinaries
50700  WriteToSpecialLogfile "Options.DirectoryGhostscriptFonts=" & Options.DirectoryGhostscriptFonts
50710  WriteToSpecialLogfile "Options.DirectoryGhostscriptLibraries=" & Options.DirectoryGhostscriptLibraries
50720  WriteToSpecialLogfile "Options.DirectoryGhostscriptResource=" & Options.DirectoryGhostscriptResource
50730  WriteToSpecialLogfile "Options.PrinterTemppath=" & Options.PrinterTemppath
50740
50750  WriteToSpecialLogfile ""
50760  WriteToSpecialLogfile "PDFCreatorTempPath=" & Options.PrinterTemppath
50770  WriteToSpecialLogfile "GetPDFCreatorTempfolder=" & GetPDFCreatorTempfolder
50780
50790  If SleepTime > 0 Then
50800   Sleep SleepTime
50810  End If
50820
50830  If StartPDFCreatorProgram And ProgramIsRunning(PDFCreator_GUID) = False Then
50840   Shell """" & PDFCreatorPath & "PDFCreator.exe""", vbNormalFocus
50850   Exit Sub
50860  End If
50870
50880  If PDFCreatorPrinter Then
50890   Set stdio = New clsStdIO
50900   If InstalledAsServer = False Then
50910     cinStr = stdio.StdIn(GetPDFCreatorTempfolder(, UserLocalTemp), SpooltimeSeconds)
50920     WriteToSpecialLogfile "cinStr=" & cinStr
50930     If FileLen(cinStr) > 0 Then
50940       WriteToSpecialLogfile "cinstr-Filelen=" & CStr(FileLen(cinStr))
50950       If IsWin9xMe = True Then
50960         Shell """" & PDFCreatorPath & "PDFCreator.exe"" " & Command$ & " -IF""" & cinStr & """"
50970        Else
50980         AppPath = PDFCreatorPath & "PDFCreator.exe"
50990         AppParams = " " & Command$ & " -IF""" & cinStr & """" ' The leading space is important and cannot be removed.
51000         If LoggedInConsole = True Then
51010          WriteToSpecialLogfile "Start RunAsUser: " & AppPath & ", " & AppParams
51020          RunAsUser hToken, AppPath, AppParams, App.Path
51030         End If
51040       End If
51050      Else
51060       KillFile cinStr
51070     End If
51080    Else
51090     WriteToSpecialLogfile "InstalledAsServer=1"
51100     tStr = GetPDFCreatorTempfolder
51110     cinStr = stdio.StdIn(tStr, SpooltimeSeconds)
51120     WriteToSpecialLogfile "cinStr=" & cinStr
51130   End If
51140   Set stdio = Nothing
51150  End If
51160
51170  WriteToSpecialLogfile ""
51180  WriteEnvironmentToSpecialLogfile
51190
51200  If IsWin9xMe = False Then
51210   UnloadProfile hToken, hProfile
51220   CloseToken hToken
51230  End If
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

Private Sub WriteInfoToSpecialLogfile()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tStr As String, c As Collection
50020  WriteToSpecialLogfile "Start: " & Now, True
50030  WriteToSpecialLogfile "ProgramReleaseStr: " & GetProgramReleaseStr
50040
50050  tStr = GetPDFCreatorApplicationPath & "PDFSpooler.exe"
50060  If FileExists(tStr) = True Then
50070    Set c = GetFileVersion(tStr)
50080    tStr = "Version: " & c(2) & "; Size: " & Format(FileLen(tStr), "###,###,###,### Bytes")
50090   Else
50100    tStr = "-"
50110  End If
50120  WriteToSpecialLogfile "PDFSpooler.exe: " & tStr
50130
50140  tStr = GetPDFCreatorApplicationPath & "PDFCreator.exe"
50150  If FileExists(tStr) = True Then
50160    Set c = GetFileVersion(tStr)
50170    tStr = "Version: " & c(2) & "; Size: " & Format(FileLen(tStr), "###,###,###,### Bytes")
50180   Else
50190    tStr = "-"
50200  End If
50210  WriteToSpecialLogfile "PDFCreator.exe: " & tStr
50220
50230  tStr = GetPDFCreatorApplicationPath & "Languages\TransTool.exe"
50240  If FileExists(tStr) = True Then
50250    Set c = GetFileVersion(tStr)
50260    tStr = "Version: " & c(2) & "; Size: " & Format(FileLen(tStr), "###,###,###,### Bytes")
50270   Else
50280    tStr = "-"
50290  End If
50300  WriteToSpecialLogfile "TransTool.exe: " & tStr
50310
50320
50330  WriteToSpecialLogfile "Redmon settings:"
50340  tStr = "REDMON_PORT"
50350  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50360  tStr = "REDMON_JOB"
50370  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50380  tStr = "REDMON_PRINTER"
50390  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50400  tStr = "REDMON_MACHINE"
50410  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50420  tStr = "REDMON_USER"
50430  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50440  tStr = "REDMON_DOCNAME"
50450  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50460  tStr = "REDMON_FILENAME"
50470  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50480  tStr = "REDMON_SESSIONID"
50490  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50500  WriteToSpecialLogfile vbCrLf & "Computer settings:"
50510  WriteToSpecialLogfile "LoggedOn-UserName=" & GetUsername
50520  WriteToSpecialLogfile "Computer=" & GetComputerName
50530  WriteToSpecialLogfile "PDFCreatorPath=" & PDFCreatorPath
50540  WriteToSpecialLogfile "IsWin9xMe=" & IsWin9xMe
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modMain", "WriteInfoToSpecialLogfile")
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
50010  ' Commandswitches
50020  ' -LOG
50030  '      Enables logging
50040  ' -PPDFCREATORPRINTER
50050  '      Start PDFCreator if PDFSpooler found printer dates
50060  '      Now know PDFCreator that PDFSpooler call PDFCreator
50070  ' -SL<Number>
50080  '      Wait <Number> milliseconds
50090  '      Example: -SL500
50100  '      Wait 500 milliseconds
50110  ' -ST
50120  '  -STTRUE
50130  '      Start pdfcreator, after using SL
50140  Dim cSwitch As String
50150  If Len(VBA.Command$) > 0 Then
50160   If UCase$(CommandSwitch("L", False)) = "OG" Then
50170     enableSpecialLogging = True
50180    Else
50190     enableSpecialLogging = False
50200   End If
50210   cSwitch = CommandSwitch("SL", True)
50220   If LenB(cSwitch) > 0 Then
50230    If IsNumeric(cSwitch) = True Then
50240     SleepTime = CLng(cSwitch)
50250    End If
50260   End If
50270   If UCase$(CommandSwitch("ST", False)) = "TRUE" Then
50280     StartPDFCreatorProgram = True
50290    Else
50300     StartPDFCreatorProgram = False
50310   End If
50320   If UCase$(CommandSwitch("P", False)) = "PDFCREATORPRINTER" Then
50330     PDFCreatorPrinter = True
50340    Else
50350     PDFCreatorPrinter = False
50360   End If
50370   cSwitch = CommandSwitch("OPTIONSFILE", True)
50380   If LenB(cSwitch) > 0 Then
50390    If FileExists(cSwitch) = True Then
50400     Optionsfile = cSwitch
50410    End If
50420   End If
50430   ' Check running instance
50440   If UCase$(CommandSwitch("Check", False)) = "INSTANCE" Then
50450     CheckInstance = True
50460    Else
50470     CheckInstance = False
50480   End If
50490   If UCase$(CommandSwitch("NO", False)) = "START" Then
50500     NoStart = True
50510    Else
50520     NoStart = False
50530   End If
50540  End If
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
50010  SleepTime = -1
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

Public Sub CreateMutexPDFSpooler()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ' Check for a local running instance
50020  Set mutexLocal = New clsMutex
50030  If Not mutexLocal.CheckMutex(PDFSpooler_GUID) Then
50040   mutexLocal.CreateMutex PDFSpooler_GUID
50050  End If
50060  ' Check for a global running instance
50070  Set mutexGlobal = New clsMutex
50080  If Not mutexGlobal.CheckMutex("Global\" & PDFSpooler_GUID) Then
50090   mutexGlobal.CreateMutex "Global\" & PDFSpooler_GUID
50100  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modMain", "CreateMutexPDFSpooler")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
