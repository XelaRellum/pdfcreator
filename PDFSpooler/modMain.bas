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
50100  InitProgram
50110  AnalyzeCommandlineParameters
50120
50130  PDFCreatorPath = CompletePath(GetPDFCreatorApplicationPath)
50140  WriteInfoToSpecialLogfile
50150
50160  LoggedInConsole = False
50170  UserName = Environ$("Redmon_User")
50180
50190  If Len(Environ$("REDMON_SESSIONID")) > 0 Then
50200   SessionID = Environ$("REDMON_SESSIONID")
50210  End If
50220
50230  hToken = 0: hProfile = 0
50240  Call GetUserLocalDirs(HKEY_CURRENT_USER, UserAppData, UserLocalTemp)
50250  If IsWin9xMe = True Then
50260    WriteToSpecialLogfile "IsWin9xMe=True"
50270   Else
50280    WriteToSpecialLogfile "IsWin9xMe=False"
50290    If GetUserSessionToken(UserName, SessionID, hToken, enableSpecialLogging) = 0 Then
50300     If LoadProfile(UserName, hToken, hProfile) = 0 Then
50310      Call GetUserLocalDirs(hProfile, UserAppData, UserLocalTemp)
50320      LoggedInConsole = True
50330     End If
50340    End If
50350  End If
50360
50370  WriteToSpecialLogfile "UserAppData=" & UserAppData
50380  WriteToSpecialLogfile "UserLocalTemp=" & UserLocalTemp
50390
50400  If InstalledAsServer Then
50410    PDFCreatorINIFile = CompletePath(GetPDFCreatorApplicationPath) & "PDFCreator.ini"
50420   Else
50430    PDFCreatorINIFile = CompletePath(UserAppData) & "PDFcreator\PDFCreator.ini"
50440  End If
50450
50460  If LenB(Optionsfile) > 0 Then
50470   PDFCreatorINIFile = Optionsfile
50480  End If
50490
50500  WriteToSpecialLogfile "PDFCreatorINIFile=" & PDFCreatorINIFile
50510  Options = ReadOptions
50520
50530  If SleepTime > 0 Then
50540   Sleep SleepTime
50550  End If
50560
50570  If StartPDFcreatorProgram And ProgramIsRunning(PDFCreator_GUID) = False Then
50580   Shell """" & PDFCreatorPath & "PDFCreator.exe""", vbNormalFocus
50590   Set mutex = Nothing
50600   Exit Sub
50610  End If
50620
50630  If PDFCreatorPrinter Then
50640   If DirExists(UserAppData) = False Then
50650    MakePath UserAppData & "PDFcreator"
50660   End If
50670
50680   Set stdio = New clsStdIO
50690   If InstalledAsServer = False Then
50700     WriteToSpecialLogfile "InstalledAsServer=0"
50710     cinStr = stdio.StdIn(Options.PrinterTemppath, SpooltimeSeconds)
50720     WriteToSpecialLogfile "cinStr=" & cinStr
50730     If FileLen(cinStr) > 0 Then
50740       WriteToSpecialLogfile "cinstr-Filelen=" & CStr(FileLen(cinStr))
50750       If IsWin9xMe = True Then
50760         Shell """" & PDFCreatorPath & "PDFCreator.exe"" " & Command$ & " -IF""" & cinStr & """"
50770        Else
50780         AppPath = PDFCreatorPath & "PDFCreator.exe"
50790         AppParams = " " & Command$ & " -IF""" & cinStr & """" ' The leading space is important and cannot be removed.
50800         If LoggedInConsole = True Then
50810          WriteToSpecialLogfile "Start RunAsUser: " & AppPath & ", " & AppParams
50820          RunAsUser hToken, AppPath, AppParams, App.Path
50830         End If
50840       End If
50850      Else
50860       KillFile cinStr
50870     End If
50880    Else
50890     WriteToSpecialLogfile "InstalledAsServer=1"
50900     tStr = GetPDFCreatorTempfolder
50910     cinStr = stdio.StdIn(tStr, SpooltimeSeconds)
50920     WriteToSpecialLogfile "cinStr=" & cinStr
50930     Tempfile = GetTempFile(CompletePath(tStr) & CompletePath(PDFCreatorSpoolDirectory) & Environ$("REDMON_USER") & "\", "~PD")
50940     KillFile Tempfile
50950     If FileExists(Tempfile) = False Then
50960      WriteToSpecialLogfile "Move the inputfile to " & Tempfile
50970      Name cinStr As Tempfile
50980     End If
50990   End If
51000   Set stdio = Nothing
51010  End If
51020
51030  If IsWin9xMe = False Then
51040   UnloadProfile hToken, hProfile
51050   CloseToken hToken
51060  End If
51070  Set mutex = Nothing
51080 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
51090 Exit Sub
ErrPtnr_OnError:
51111 Select Case ErrPtnr.OnError("modMain", "Main")
      Case 0: Resume
51130 Case 1: Resume Next
51140 Case 2: Exit Sub
51150 Case 3: End
51160 End Select
51170 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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
50230   End If
50240   cSwitch = CommandSwitch("SL", True)
50250   If LenB(cSwitch) > 0 Then
50260    If IsNumeric(cSwitch) = True Then
50270     SleepTime = CLng(cSwitch)
50280    End If
50290   End If
50300   If UCase$(CommandSwitch("ST", False)) = "TRUE" Then
50310     StartPDFcreatorProgram = True
50320    Else
50330     StartPDFcreatorProgram = False
50340   End If
50350   If UCase$(CommandSwitch("P", False)) = "PDFCREATORPRINTER" Then
50360     PDFCreatorPrinter = True
50370    Else
50380     PDFCreatorPrinter = False
50390   End If
50400   cSwitch = CommandSwitch("OPTIONSFILE", True)
50410   If LenB(cSwitch) > 0 Then
50420    If FileExists(cSwitch) = True Then
50430     Optionsfile = cSwitch
50440    End If
50450   End If
50460  End If
50470 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50480 Exit Sub
ErrPtnr_OnError:
50501 Select Case ErrPtnr.OnError("modMain", "AnalyzeCommandlineParameters")
      Case 0: Resume
50520 Case 1: Resume Next
50530 Case 2: Exit Sub
50540 Case 3: End
50550 End Select
50560 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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

