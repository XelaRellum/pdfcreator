Attribute VB_Name = "modMain"
Option Explicit

Public mutex As clsMutex, PDFCreatorPath As String

Sub Main()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim UserName As String, SessionID As Long, _
  UserAppData As String, UserLocalTemp As String, tStr As String, _
  hToken As Long, hProfile As Long, AppPath As String, AppParams As String, _
  stdio As clsStdIO, cinStr As String, SpooltimeSeconds As Double, _
  LoggedInConsole As Boolean, Tempfile As String
50090
50100  ' Reduce the working size of used memory
50110  Call SetProcessWorkingSetSize(GetCurrentProcess(), -1, -1)
50120
50130  InitProgram
50140  AnalyzeCommandlineParameters
50150
50160  PDFCreatorPath = CompletePath(GetPDFCreatorApplicationPath)
50170  WriteInfoToSpecialLogfile
50180
50190  LoggedInConsole = False
50200  UserName = Environ$("Redmon_User")
50210
50220  If Len(Environ$("REDMON_SESSIONID")) > 0 Then
50230   SessionID = Environ$("REDMON_SESSIONID")
50240  End If
50250
50260  hToken = 0: hProfile = 0
50270  Call GetUserLocalDirs(HKEY_CURRENT_USER, UserAppData, UserLocalTemp)
50280  If IsWin9xMe = True Then
50290    WriteToSpecialLogfile "IsWin9xMe=True"
50300   Else
50310    WriteToSpecialLogfile "IsWin9xMe=False"
50320    If GetUserSessionToken(UserName, SessionID, hToken, enableSpecialLogging) = 0 Then
50330     If LoadProfile(UserName, hToken, hProfile) = 0 Then
50340      Call GetUserLocalDirs(hProfile, UserAppData, UserLocalTemp)
50350      LoggedInConsole = True
50360     End If
50370    End If
50380  End If
50390
50400  WriteToSpecialLogfile "UserAppData=" & UserAppData
50410  WriteToSpecialLogfile "UserLocalTemp=" & UserLocalTemp
50420
50430  If InstalledAsServer Then
50440    PDFCreatorINIFile = CompletePath(GetPDFCreatorApplicationPath) & "PDFCreator.ini"
50450   Else
50460    PDFCreatorINIFile = CompletePath(UserAppData) & "PDFcreator\PDFCreator.ini"
50470  End If
50480
50490  If LenB(Optionsfile) > 0 Then
50500   PDFCreatorINIFile = Optionsfile
50510  End If
50520
50530  WriteToSpecialLogfile "PDFCreatorINIFile=" & PDFCreatorINIFile
50540
50550  Options = ReadOptions
50560
50570  If SleepTime > 0 Then
50580   Sleep SleepTime
50590  End If
50600
50610  If StartPDFcreatorProgram And ProgramIsRunning(PDFCreator_GUID) = False Then
50620   Shell """" & PDFCreatorPath & "PDFCreator.exe""", vbNormalFocus
50630   Set mutex = Nothing
50640   Exit Sub
50650  End If
50660
50670  If PDFCreatorPrinter Then
50680   If DirExists(UserAppData) = False Then
50690    MakePath UserAppData & "PDFcreator"
50700   End If
50710
50720   Set stdio = New clsStdIO
50730   If InstalledAsServer = False Then
50740     WriteToSpecialLogfile "InstalledAsServer=0"
50750     cinStr = stdio.StdIn(Options.PrinterTemppath, SpooltimeSeconds)
50760     WriteToSpecialLogfile "cinStr=" & cinStr
50770     If FileLen(cinStr) > 0 Then
50780       WriteToSpecialLogfile "cinstr-Filelen=" & CStr(FileLen(cinStr))
50790       If IsWin9xMe = True Then
50800         Shell """" & PDFCreatorPath & "PDFCreator.exe"" " & Command$ & " -IF""" & cinStr & """"
50810        Else
50820         AppPath = PDFCreatorPath & "PDFCreator.exe"
50830         AppParams = " " & Command$ & " -IF""" & cinStr & """" ' The leading space is important and cannot be removed.
50840         If LoggedInConsole = True Then
50850          WriteToSpecialLogfile "Start RunAsUser: " & AppPath & ", " & AppParams
50860          RunAsUser hToken, AppPath, AppParams, App.Path
50870         End If
50880       End If
50890      Else
50900       KillFile cinStr
50910     End If
50920    Else
50930     WriteToSpecialLogfile "InstalledAsServer=1"
50940     tStr = GetPDFCreatorTempfolder
50950     cinStr = stdio.StdIn(tStr, SpooltimeSeconds)
50960     WriteToSpecialLogfile "cinStr=" & cinStr
50970   End If
50980   Set stdio = Nothing
50990  End If
51000
51010  If IsWin9xMe = False Then
51020   UnloadProfile hToken, hProfile
51030   CloseToken hToken
51040  End If
51050  Set mutex = Nothing
51060 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
51070 Exit Sub
ErrPtnr_OnError:
51091 Select Case ErrPtnr.OnError("modMain", "Main")
      Case 0: Resume
51110 Case 1: Resume Next
51120 Case 2: Exit Sub
51130 Case 3: End
51140 End Select
51150 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub WriteInfoToSpecialLogfile()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim tStr As String
50050  WriteToSpecialLogfile "Start: " & Now, True
50060  WriteToSpecialLogfile "Redmon settings:"
50070  tStr = "REDMON_PORT"
50080  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50090  tStr = "REDMON_JOB"
50100  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50110  tStr = "REDMON_PRINTER"
50120  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50130  tStr = "REDMON_MACHINE"
50140  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50150  tStr = "REDMON_USER"
50160  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50170  tStr = "REDMON_DOCNAME"
50180  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50190  tStr = "REDMON_FILENAME"
50200  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50210  tStr = "REDMON_SESSIONID"
50220  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50230  WriteToSpecialLogfile vbCrLf & "Computer settings:"
50240  WriteToSpecialLogfile "LoggedOn-UserName=" & GetUsername
50250  WriteToSpecialLogfile "Computer=" & GetComputerName
50260  WriteToSpecialLogfile "PDFCreatorPath=" & PDFCreatorPath
50270  WriteToSpecialLogfile vbCrLf
50280 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50290 Exit Sub
ErrPtnr_OnError:
50311 Select Case ErrPtnr.OnError("modMain", "WriteInfoToSpecialLogfile")
      Case 0: Resume
50330 Case 1: Resume Next
50340 Case 2: Exit Sub
50350 Case 3: End
50360 End Select
50370 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub AnalyzeCommandlineParameters()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  ' Commandswitches
50050  ' -LOG
50060  '      Enables logging
50070  ' -PPDFCREATORPRINTER
50080  '      Start PDFCreator if PDFSpooler found printer dates
50090  '      Now know PDFCreator that PDFSpooler call PDFCreator
50100  ' -SL<Number>
50110  '      Wait <Number> milliseconds
50120  '      Example: -SL500
50130  '      Wait 500 milliseconds
50140  ' -ST
50150  '  -STTRUE
50160  '      Start pdfcreator, after using SL
50170  Dim cSwitch As String
50180  If Len(VBA.Command$) > 0 Then
50190   If UCase$(CommandSwitch("L", False)) = "OG" Then
50200     enableSpecialLogging = True
50210    Else
50220     enableSpecialLogging = False
50230     enableSpecialLogging = True
50240   End If
50250   cSwitch = CommandSwitch("SL", True)
50260   If LenB(cSwitch) > 0 Then
50270    If IsNumeric(cSwitch) = True Then
50280     SleepTime = CLng(cSwitch)
50290    End If
50300   End If
50310   If UCase$(CommandSwitch("ST", False)) = "TRUE" Then
50320     StartPDFcreatorProgram = True
50330    Else
50340     StartPDFcreatorProgram = False
50350   End If
50360   If UCase$(CommandSwitch("P", False)) = "PDFCREATORPRINTER" Then
50370     PDFCreatorPrinter = True
50380    Else
50390     PDFCreatorPrinter = False
50400   End If
50410   cSwitch = CommandSwitch("OPTIONSFILE", True)
50420   If LenB(cSwitch) > 0 Then
50430    If FileExists(cSwitch) = True Then
50440     Optionsfile = cSwitch
50450    End If
50460   End If
50470  End If
50480 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50490 Exit Sub
ErrPtnr_OnError:
50511 Select Case ErrPtnr.OnError("modMain", "AnalyzeCommandlineParameters")
      Case 0: Resume
50530 Case 1: Resume Next
50540 Case 2: Exit Sub
50550 Case 3: End
50560 End Select
50570 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub InitProgram()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  SleepTime = -1
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Sub
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("modMain", "InitProgram")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Sub
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
