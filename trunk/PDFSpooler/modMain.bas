Attribute VB_Name = "modMain"
Option Explicit

Public mutex As clsMutex, enableLogging As Boolean

Sub Main()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim UserName As String, SessionID As Long, PDFCreatorPath As String, _
  UserAppData As String, UserLocalTemp As String, _
  hToken As Long, hProfile As Long, mSeconds As Long, cSwitch As String, _
  stSwitch As String, AppPath As String, AppParams As String, stdio As clsStdIO, _
  cinStr As String, SpooltimeSeconds As Double
50090
50100  ' Commandswitches
50110  ' -LOG
50120  '      Enables logging
50130  ' -PPDFCREATORPRINTER
50140  '      Start PDFCreator if PDFSpooler found printer dates
50150  '      Now know PDFCreator that PDFSpooler call PDFCreator
50160  ' -SL<Number>
50170  '      Wait <Number> milliseconds
50180  '      Example: -SL500
50190  '      Wait 500 milliseconds
50200  ' -ST
50210  '  -STTRUE
50220  '      Start pdfcreator, after using SL
50230
50240  ' LOG switch
50250  If UCase$(CommandSwitch("L", True)) = "OG" Then
50260    enableLogging = True
50270   Else
50280    enableLogging = False
50290  End If
50300  WriteToSpecialLogfile "Start", True
50310  UserName = Environ$("Redmon_User")
50320  If Len(Environ$("REDMON_SESSIONID")) > 0 Then
50330   SessionID = Environ$("REDMON_SESSIONID")
50340  End If
50350  PDFCreatorPath = CompletePath(GetPDFCreatorApplicationPath)
50360  PDFCreatorLogfilePath = Environ$("TEMP") & "\"
50370  IfLoggingWriteLogfile "PDFSpooler Program Start"
50380
50390  WriteToSpecialLogfile "UserName=" & UserName
50400  WriteToSpecialLogfile "SessionID=" & SessionID
50410  WriteToSpecialLogfile "PDFCreatorPath=" & PDFCreatorPath
50420  WriteToSpecialLogfile "PDFCreatorLogfilePath=" & PDFCreatorLogfilePath
50430
50440  If IsWin9xMe = True Then
50450    hToken = 0
50460    hProfile = 0
50470    Call GetUserLocalDirs(HKEY_CURRENT_USER, UserAppData, UserLocalTemp)
50480    WriteToSpecialLogfile "IsWin9xMe = True"
50490   Else
50500    If GetUserSessionToken(UserName, SessionID, hToken) Then
50510     Set mutex = Nothing
50520     Exit Sub
50530    End If
50540    If LoadProfile(UserName, hToken, hProfile) Then
50550     Set mutex = Nothing
50560     CloseToken hToken
50570     Exit Sub
50580    End If
50590    Call GetUserLocalDirs(hProfile, UserAppData, UserLocalTemp)
50600    WriteToSpecialLogfile "IsWin9xMe = False"
50610  End If
50620  WriteToSpecialLogfile "UserAppData=" & UserAppData
50630  WriteToSpecialLogfile "UserLocalTemp=" & UserLocalTemp
50640
50650  PDFCreatorINIFile = UserAppData & "PDFcreator\PDFCreator.ini"
50660
50670  WriteToSpecialLogfile "PDFCreatorINIFile=" & PDFCreatorINIFile
50680
50690  ' SL switch - ask for time
50700  cSwitch = CommandSwitch("SL", True)
50710  If LenB(cSwitch) > 0 Then
50720   If IsNumeric(cSwitch) = True Then
50730    mSeconds = CLng(cSwitch)
50740    Sleep mSeconds
50750   End If
50760  End If
50770
50780  ' ST switch - ask for simple run
50790  stSwitch = CommandSwitch("ST", True)
50800  If LenB(stSwitch) > 0 Then
50810   If UCase$(stSwitch) = "TRUE" Then
50820    Shell """" & PDFCreatorPath & "PDFCreator.exe""", vbNormalFocus
50830    Set mutex = Nothing
50840    End
50850   End If
50860  End If
50870
50880  ' PDFCREATORPRINTER switch
50890  If UCase$(CommandSwitch("P", True)) = "PDFCREATORPRINTER" Then
50900   If CheckPath(UserAppData) = True Then
50910    If Dir(UserAppData & "PDFcreator", vbDirectory) = "" Then
50920     MakePath UserAppData & "PDFcreator"
50930    End If
50940   End If
50950   Options = ReadOptions
50960   If Options.Logging > 0 Then
50970    enableLogging = True
50980   End If
50990
51000   Set stdio = New clsStdIO
51010   cinStr = stdio.StdIn(UserLocalTemp, SpooltimeSeconds)
51020   Set stdio = Nothing
51030
51040   WriteToSpecialLogfile "cinStr=" & cinStr
51050
51060   If FileLen(cinStr) > 0 Then
51070     If IsWin9xMe = True Then
51080       Shell """" & PDFCreatorPath & "PDFCreator.exe"" -PPDFCREATORPRINTER -IF""" & cinStr & """"
51090      Else
51100       AppPath = PDFCreatorPath & "PDFCreator.exe"
51110       AppParams = " -PPDFCREATORPRINTER -IF""" & cinStr & """"
51120       RunAsUser hToken, AppPath, AppParams, App.Path
51130     End If
51140    Else
51150     Kill cinStr
51160   End If
51170  End If
51180
51190  If IsWin9xMe = False Then
51200   UnloadProfile hToken, hProfile
51210   CloseToken hToken
51220  End If
51230  IfLoggingWriteLogfile "PDFSpooler: Spoolfile: " & cinStr
51240  IfLoggingWriteLogfile "PDFSpooler Program End"
51250  Set mutex = Nothing
51260 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
51270 Exit Sub
ErrPtnr_OnError:
51291 Select Case ErrPtnr.OnError("modMain", "Main")
      Case 0: Resume
51310 Case 1: Resume Next
51320 Case 2: Exit Sub
51330 Case 3: End
51340 End Select
51350 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub WriteToSpecialLogfile(StatusText As String, Optional CreateFile As Boolean = False)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim fname As String, fn As Long
50050  If enableLogging = True Then
50060   fname = CompletePath(Environ$("Systemdrive")) & "PDFCreator-Errorlog.txt"
50070   fn = FreeFile
50080   If FileExists(fname) = False Or CreateFile = True Then
50090     Open fname For Output As #fn
50100    Else
50110     Open fname For Append As #fn
50120   End If
50130   Print #fn, StatusText
50140   Close #fn
50150  End If
50160 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50170 Exit Sub
ErrPtnr_OnError:
50191 Select Case ErrPtnr.OnError("modMain", "WriteToSpecialLogfile")
      Case 0: Resume
50210 Case 1: Resume Next
50220 Case 2: Exit Sub
50230 Case 3: End
50240 End Select
50250 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

