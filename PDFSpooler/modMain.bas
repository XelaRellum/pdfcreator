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
50360  If IsWin9xMe = True Then
50370    WriteToSpecialLogfile "IsWin9xMe=True"
50380   Else
50390    WriteToSpecialLogfile "IsWin9xMe=False"
50400    If GetUserSessionToken(UserName, SessionID, hToken, enableSpecialLogging) = 0 Then
50410     If LoadProfile(UserName, hToken, hProfile) = 0 Then
50420      Call GetUserLocalDirs(hProfile, UserAppData, UserLocalTemp)
50430      LoggedInConsole = True
50440     End If
50450    End If
50460  End If
50470  WriteToSpecialLogfile "UserAppData=" & UserAppData
50480  WriteToSpecialLogfile "UserLocalTemp=" & UserLocalTemp
50490
50500  If InstalledAsServer Then
50510    PDFCreatorINIFile = CompletePath(GetPDFCreatorApplicationPath) & "PDFCreator.ini"
50520   Else
50530    PDFCreatorINIFile = CompletePath(UserAppData) & "PDFcreator\PDFCreator.ini"
50540  End If
50550  If LenB(Optionsfile) > 0 Then
50560   PDFCreatorINIFile = Optionsfile
50570  End If
50580
50590  WriteToSpecialLogfile "InstalledAsServer=" & InstalledAsServer
50600  WriteToSpecialLogfile "UseINI=" & UseINI
50610  WriteToSpecialLogfile "PDFCreatorINIFile=" & PDFCreatorINIFile
50620
50630  InitLanguagesStrings
50640  If IsWin9xMe Then
50650    ReadLanguageFromOptions
50660   Else
50670    ReadLanguageFromOptions hProfile
50680  End If
50690  WriteToSpecialLogfile "Language=" & Options.Language
50700  LanguagePath = CompletePath(GetPDFCreatorApplicationPath) & "Languages\"
50710  Languagefile = LanguagePath & Options.Language & ".ini"
50720  If UCase$(Options.Language) = "ESPANOL" And FileExists(Languagefile) = False And _
  FileExists(LanguagePath & "spanish.ini") = True Then
50740   Languagefile = LanguagePath & "spanish.ini"
50750   Options.Language = "spanish"
50760  End If
50770  If FileExists(Languagefile) = True Then
50780    LoadLanguage Languagefile
50790   Else
50800    MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & ">" & Languagefile & "<"
50810  '   Options.Language = "english"
50820  End If
50830
50840  If IsWin9xMe Then
50850    Options = ReadOptions(True)
50860   Else
50870    Options = ReadOptions(True, hProfile)
50880  End If
50890
50900  WriteToSpecialLogfile ""
50910  WriteToSpecialLogfile "PDFCreatorTempPath=" & Options.PrinterTemppath
50920  WriteToSpecialLogfile "GetPDFCreatorTempfolder=" & GetPDFCreatorTempfolder
50930
50940  If SleepTime > 0 Then
50950   Sleep SleepTime
50960  End If
50970
50980  If StartPDFCreatorProgram And ProgramIsRunning(PDFCreator_GUID) = False Then
50990   Shell """" & PDFCreatorPath & "PDFCreator.exe""", vbNormalFocus
51000   Exit Sub
51010  End If
51020
51030  If PDFCreatorPrinter Then
51040   Set stdio = New clsStdIO
51050   If InstalledAsServer = False Then
51060     cinStr = stdio.StdIn(GetPDFCreatorTempfolder(, UserLocalTemp), SpooltimeSeconds)
51070     WriteToSpecialLogfile "cinStr=" & cinStr
51080     If FileLen(cinStr) > 0 Then
51090       WriteToSpecialLogfile "cinstr-Filelen=" & CStr(FileLen(cinStr))
51100       If IsWin9xMe = True Then
51110         Shell """" & PDFCreatorPath & "PDFCreator.exe"" " & Command$ & " -IF""" & cinStr & """"
51120        Else
51130         AppPath = PDFCreatorPath & "PDFCreator.exe"
51140         AppParams = " " & Command$ & " -IF""" & cinStr & """" ' The leading space is important and cannot be removed.
51150         If LoggedInConsole = True Then
51160          WriteToSpecialLogfile "Start RunAsUser: " & AppPath & ", " & AppParams
51170          RunAsUser hToken, AppPath, AppParams, App.Path
51180         End If
51190       End If
51200      Else
51210       KillFile cinStr
51220     End If
51230    Else
51240     WriteToSpecialLogfile "InstalledAsServer=1"
51250     tStr = GetPDFCreatorTempfolder
51260     cinStr = stdio.StdIn(tStr, SpooltimeSeconds)
51270     WriteToSpecialLogfile "cinStr=" & cinStr
51280   End If
51290   Set stdio = Nothing
51300  End If
51310
51320  WriteToSpecialLogfile ""
51330  WriteEnvironmentToSpecialLogfile
51340
51350  If IsWin9xMe = False Then
51360   UnloadProfile hToken, hProfile
51370   CloseToken hToken
51380  End If
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
50010  Dim tStr As String
50020  WriteToSpecialLogfile "Start: " & Now, True
50030  WriteToSpecialLogfile "Redmon settings:"
50040  tStr = "REDMON_PORT"
50050  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50060  tStr = "REDMON_JOB"
50070  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50080  tStr = "REDMON_PRINTER"
50090  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50100  tStr = "REDMON_MACHINE"
50110  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50120  tStr = "REDMON_USER"
50130  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50140  tStr = "REDMON_DOCNAME"
50150  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50160  tStr = "REDMON_FILENAME"
50170  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50180  tStr = "REDMON_SESSIONID"
50190  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50200  WriteToSpecialLogfile vbCrLf & "Computer settings:"
50210  WriteToSpecialLogfile "LoggedOn-UserName=" & GetUsername
50220  WriteToSpecialLogfile "Computer=" & GetComputerName
50230  WriteToSpecialLogfile "PDFCreatorPath=" & PDFCreatorPath
50240  WriteToSpecialLogfile vbCrLf
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

