Attribute VB_Name = "modMain"
Option Explicit

Public mutex As clsMutex

Sub Main()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim UserName As String, SessionID As Long, PDFCreatorPath As String, _
  UserAppData As String, UserLocalTemp As String, tStr As String, _
  hToken As Long, hProfile As Long, mSeconds As Long, cSwitch As String, _
  stSwitch As String, AppPath As String, AppParams As String, stdio As clsStdIO, _
  cinStr As String, SpooltimeSeconds As Double, LoggedInConsole As Boolean, _
  LogStr As String, Tempfile As String
50070
50080  ' Commandswitches
50090  ' -LOG
50100  '      Enables logging
50110  ' -PPDFCREATORPRINTER
50120  '      Start PDFCreator if PDFSpooler found printer dates
50130  '      Now know PDFCreator that PDFSpooler call PDFCreator
50140  ' -SL<Number>
50150  '      Wait <Number> milliseconds
50160  '      Example: -SL500
50170  '      Wait 500 milliseconds
50180  ' -ST
50190  '  -STTRUE
50200  '      Start pdfcreator, after using SL
50210
50220  ' LOG switch
50230  If UCase$(CommandSwitch("L", True)) = "OG" Then
50240    enableSpecialLogging = True: LogStr = "-LOG"
50250   Else
50260    enableSpecialLogging = False: LogStr = ""
50270  End If
50280  WriteToSpecialLogfile "Start: " & Now, True
50290
50300  LoggedInConsole = False
50310  UserName = Environ$("Redmon_User")
50320
50330  If Len(Environ$("REDMON_SESSIONID")) > 0 Then
50340   SessionID = Environ$("REDMON_SESSIONID")
50350  End If
50360  PDFCreatorPath = CompletePath(GetPDFCreatorApplicationPath)
50370  PDFCreatorLogfilePath = CompletePath(Environ$("TEMP"))
50380
50390  WriteToSpecialLogfile "Redmon settings:"
50400  tStr = "REDMON_PORT"
50410  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50420  tStr = "REDMON_JOB"
50430  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50440  tStr = "REDMON_PRINTER"
50450  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50460  tStr = "REDMON_MACHINE"
50470  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50480  tStr = "REDMON_USER"
50490  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50500  tStr = "REDMON_DOCNAME"
50510  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50520  tStr = "REDMON_FILENAME"
50530  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50540  tStr = "REDMON_SESSIONID"
50550  WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
50560
50570  WriteToSpecialLogfile vbCrLf & "Computer settings:"
50580  WriteToSpecialLogfile "LoggedOn-UserName=" & GetUsername
50590  WriteToSpecialLogfile "Computer=" & GetComputerName
50600  WriteToSpecialLogfile "PDFCreatorPath=" & PDFCreatorPath
50610  WriteToSpecialLogfile "PDFCreatorLogfilePath=" & PDFCreatorLogfilePath
50620
50630  hToken = 0: hProfile = 0
50640  Call GetUserLocalDirs(HKEY_CURRENT_USER, UserAppData, UserLocalTemp)
50650  If IsWin9xMe = True Then
50660    WriteToSpecialLogfile "IsWin9xMe = True"
50670   Else
50680    WriteToSpecialLogfile "IsWin9xMe = False"
50690    If GetUserSessionToken(UserName, SessionID, hToken, enableSpecialLogging) = 0 Then
50700     If LoadProfile(UserName, hToken, hProfile) = 0 Then
50710      Call GetUserLocalDirs(hProfile, UserAppData, UserLocalTemp)
50720      LoggedInConsole = True
50730     End If
50740    End If
50750  End If
50760
50770  WriteToSpecialLogfile "UserAppData=" & UserAppData
50780  WriteToSpecialLogfile "UserLocalTemp=" & UserLocalTemp
50790
50800  If InstalledAsServer = True Then
50810    PDFCreatorINIFile = CompletePath(GetPDFCreatorApplicationPath) & "PDFCreator.ini"
50820   Else
50830    PDFCreatorINIFile = CompletePath(UserAppData) & "PDFcreator\PDFCreator.ini"
50840  End If
50850
50860  WriteToSpecialLogfile "PDFCreatorINIFile=" & PDFCreatorINIFile
50870
50880  ' SL switch - ask for time
50890  cSwitch = CommandSwitch("SL", True)
50900  If LenB(cSwitch) > 0 Then
50910   If IsNumeric(cSwitch) = True Then
50920    mSeconds = CLng(cSwitch)
50930    Sleep mSeconds
50940   End If
50950  End If
50960
50970  ' ST switch - ask for simple run
50980  stSwitch = CommandSwitch("ST", True)
50990  If LenB(stSwitch) > 0 Then
51000   If UCase$(stSwitch) = "TRUE" And ProgramIsRunning(PDFCreator_GUID) = False Then
51010    Shell """" & PDFCreatorPath & "PDFCreator.exe""", vbNormalFocus
51020    Set mutex = Nothing
51030    End
51040   End If
51050  End If
51060
51070  ' PDFCREATORPRINTER switch
51080  If UCase$(CommandSwitch("P", True)) = "PDFCREATORPRINTER" Then
51090   If DirExists(UserAppData) = False Then
51100    MakePath UserAppData & "PDFcreator"
51110   End If
51120
51130   Set stdio = New clsStdIO
51140   If InstalledAsServer = False Then
51150     WriteToSpecialLogfile "InstalledAsServer=0"
51160     If IsWin9xMe = True Then
51170       cinStr = stdio.StdIn(GetTempPath, SpooltimeSeconds)
51180      Else
51190       cinStr = stdio.StdIn(UserLocalTemp, SpooltimeSeconds)
51200     End If
51210     WriteToSpecialLogfile "cinStr=" & cinStr
51220     If FileLen(cinStr) > 0 Then
51230       If IsWin9xMe = True Then
51240         Shell """" & PDFCreatorPath & "PDFCreator.exe"" -PPDFCREATORPRINTER -IF""" & cinStr & """" & " " & LogStr
51250        Else
51260         AppPath = PDFCreatorPath & "PDFCreator.exe"
51270         AppParams = " -PPDFCREATORPRINTER -IF""" & cinStr & """" & " " & LogStr
51280         If LoggedInConsole = True Then
51290          WriteToSpecialLogfile "Start RunAsUser: " & AppPath & ", " & AppParams
51300          RunAsUser hToken, AppPath, AppParams, App.Path
51310         End If
51320       End If
51330      Else
51340       KillFile cinStr
51350     End If
51360    Else
51370     WriteToSpecialLogfile "InstalledAsServer=1"
51380     Options = ReadOptions
51390     tStr = GetPDFCreatorTempfolder
51400     cinStr = stdio.StdIn(tStr, SpooltimeSeconds)
51410     WriteToSpecialLogfile "cinStr=" & cinStr
51420     Tempfile = GetTempFile(CompletePath(tStr) & PDFCreatorSpoolDirectory & "\" & Environ$("REDMON_USER") & "\", "~PD")
51430     KillFile Tempfile
51440     If FileExists(Tempfile) = False Then
51450      'IfLoggingWriteLogfile "Move the inputfile to " & Tempfile
51460      WriteToSpecialLogfile "Move the inputfile to " & Tempfile
51470      Name cinStr As Tempfile
51480     End If
51490   End If
51500   Set stdio = Nothing
51510  End If
51520
51530  If IsWin9xMe = False Then
51540   UnloadProfile hToken, hProfile
51550   CloseToken hToken
51560  End If
51570  Set mutex = Nothing
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
