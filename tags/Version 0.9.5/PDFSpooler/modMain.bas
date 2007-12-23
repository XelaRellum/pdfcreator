Attribute VB_Name = "modMain"
Option Explicit


Public PDFCreatorPath As String

'Const MAXIMUM_ALLOWED As Long = &H2000000
'Private Declare Function ImpersonateSelf Lib "advapi32.dll" (ByVal ImpersonationLevel As Long) As Long
'Public Declare Function ImpersonateLoggedOnUser Lib "advapi32.dll" (ByVal hToken As Long) As Long
'Private Declare Function DuplicateTokenEx Lib "advapi32" (ByVal hExistingToken As Long, ByVal dwDesiredAcces As Long, lpTokenAttribute As Long, ImpersonatonLevel As Long, ByVal tokenType As Long, Phandle As Long) As Long

Public Sub Main()
 Dim UserName As String, SessionID As Long, _
  UserAppData As String, UserLocalTemp As String, tStr As String, _
  hToken As Long, hProfile As Long, AppPath As String, AppParams As String, _
  stdio As clsStdIO, cinStr As String, SpooltimeSeconds As Double, _
  LoggedInConsole As Boolean, Tempfile As String, res As Long

 ' Reduce the working size of used memory
 Call SetProcessWorkingSetSize(GetCurrentProcess(), -1, -1)

 InitProgram
 AnalyzeCommandlineParameters

 If CheckInstance Then
  CheckProgramInstances
 End If

 If NoStart Then
  Exit Sub
 End If

' Create a local and a global mutex
 CreateMutexPDFSpooler

 PDFCreatorPath = CompletePath(GetPDFCreatorApplicationPath)
 Call WriteInfoToSpecialLogfile

 LoggedInConsole = False
 UserName = Environ$("Redmon_User")

 If Len(Environ$("REDMON_SESSIONID")) > 0 Then
  SessionID = Environ$("REDMON_SESSIONID")
 End If
 
 hToken = 0: hProfile = 0
 If IsWin9xMe = False Then
   WriteToSpecialLogfile "Start GetUserSessionToken"
   res = GetUserSessionToken(UserName, SessionID, hToken, enableSpecialLogging)
'   WriteToSpecialLogfile "Try to ImpersonateSelf"
'   res = ImpersonateLoggedOnUser(hToken)
'   WriteToSpecialLogfile "res ImpersonateSelf = " & res
'   If res <> 0 Then res = 0
'   Dim hNewToken As Long
'   res = DuplicateTokenEx(hToken, MAXIMUM_ALLOWED, 0, 2, 1, hNewToken)
'   WriteToSpecialLogfile "res DuplicateTokenEx = " & res
'   If res <> 0 Then res = 0
   If res = 0 Then
     res = LoadProfile(UserName, hToken, hProfile)
     If res = 0 Then
       WriteToSpecialLogfile "hProfile = " & hProfile
       Call GetUserLocalDirs(hProfile, UserAppData, UserLocalTemp)
       LoggedInConsole = True
      Else
       WriteToSpecialLogfile "LoadProfile -> Error [" & res & "] = " & RaiseAPIError(res)
     End If
    Else
     WriteToSpecialLogfile "GetUserSessionToken -> Error"
   End If
  Else
   Call GetUserLocalDirs(HKEY_CURRENT_USER, UserAppData, UserLocalTemp)
 End If
 WriteToSpecialLogfile "UserAppData=" & UserAppData
 WriteToSpecialLogfile "UserLocalTemp=" & UserLocalTemp

 If InstalledAsServer Then
   PDFCreatorINIFile = CompletePath(GetPDFCreatorApplicationPath) & "PDFCreator.ini"
  Else
   PDFCreatorINIFile = CompletePath(UserAppData) & "PDFcreator\PDFCreator.ini"
 End If
 If LenB(Optionsfile) > 0 Then
  PDFCreatorINIFile = Optionsfile
 End If

 WriteToSpecialLogfile "InstalledAsServer=" & InstalledAsServer
 WriteToSpecialLogfile "UseINI=" & UseINI
 WriteToSpecialLogfile "PDFCreatorINIFile=" & PDFCreatorINIFile

 If IsWin9xMe Then
   Options = ReadOptions(True)
  Else
   WriteToSpecialLogfile "Read options: start"
   Options = ReadOptions(True, hProfile)
   WriteToSpecialLogfile "Read options: ready"
 End If

 WriteToSpecialLogfile ""
 WriteToSpecialLogfile "Options.DirectoryGhostscriptBinaries=" & Options.DirectoryGhostscriptBinaries
 WriteToSpecialLogfile "Options.DirectoryGhostscriptFonts=" & Options.DirectoryGhostscriptFonts
 WriteToSpecialLogfile "Options.DirectoryGhostscriptLibraries=" & Options.DirectoryGhostscriptLibraries
 WriteToSpecialLogfile "Options.DirectoryGhostscriptResource=" & Options.DirectoryGhostscriptResource
 WriteToSpecialLogfile "Options.PrinterTemppath=" & Options.PrinterTemppath

 WriteToSpecialLogfile ""
 WriteToSpecialLogfile "PDFCreatorTempPath=" & Options.PrinterTemppath
 WriteToSpecialLogfile "GetPDFCreatorTempfolder=" & GetPDFCreatorTempfolder(, UserLocalTemp)

 If SleepTime > 0 Then
  Sleep SleepTime
 End If

 If StartPDFCreatorProgram And ProgramIsRunning(PDFCreator_GUID) = False Then
  Shell """" & PDFCreatorPath & "PDFCreator.exe""", vbNormalFocus
  Exit Sub
 End If

 If PDFCreatorPrinter Then
  Set stdio = New clsStdIO
  If InstalledAsServer = False Then
    cinStr = stdio.StdIn(GetPDFCreatorTempfolder(, UserLocalTemp), SpooltimeSeconds)
    WriteToSpecialLogfile "cinStr=" & cinStr
    If FileLen(cinStr) > 0 Then
      WriteToSpecialLogfile "cinstr-Filelen=" & CStr(FileLen(cinStr))
      If IsWin9xMe = True Then
        tStr = """" & PDFCreatorPath & "PDFCreator.exe"" " & Command$ & " -IF""" & cinStr & """"
        WriteToSpecialLogfile "IsWin9xMe: Start shell: " & tStr
        Shell tStr
       Else
        AppPath = PDFCreatorPath & "PDFCreator.exe"
        AppParams = " " & Command$ & " -IF""" & cinStr & """" ' The leading space is important and cannot be removed.
        WriteToSpecialLogfile "Not IsWin9xMe: LoggedInConsole " & CStr(LoggedInConsole)
        If LoggedInConsole = True Then
         WriteToSpecialLogfile "Start RunAsUser: " & AppPath & ", " & AppParams
         RunAsUser hToken, AppPath, AppParams, App.Path
        End If
      End If
     Else
      KillFile cinStr
    End If
   Else
    WriteToSpecialLogfile "InstalledAsServer=1"
    tStr = GetPDFCreatorTempfolder
    cinStr = stdio.StdIn(tStr, SpooltimeSeconds)
    WriteToSpecialLogfile "cinStr=" & cinStr
  End If
  Set stdio = Nothing
 End If

 WriteToSpecialLogfile ""
 WriteEnvironmentToSpecialLogfile

 If IsWin9xMe = False Then
  UnloadProfile hToken, hProfile
  CloseToken hToken
 End If

 WriteToSpecialLogfile vbCrLf & "End: " & Now

End Sub

Private Sub WriteInfoToSpecialLogfile()
 Dim tStr As String, c As Collection

 WriteToSpecialLogfile "ProgramReleaseStr: " & GetProgramReleaseStr, True

 tStr = GetPDFCreatorApplicationPath & "PDFSpooler.exe"
 If FileExists(tStr) = True Then
   Set c = GetFileVersion(tStr)
   tStr = "Version: " & c(2) & "; Size: " & Format(FileLen(tStr), "###,###,###,### Bytes")
  Else
   tStr = "-"
 End If
 WriteToSpecialLogfile "PDFSpooler.exe: " & tStr

 tStr = GetPDFCreatorApplicationPath & "PDFCreator.exe"
 If FileExists(tStr) = True Then
   Set c = GetFileVersion(tStr)
   tStr = "Version: " & c(2) & "; Size: " & Format(FileLen(tStr), "###,###,###,### Bytes")
  Else
   tStr = "-"
 End If
 WriteToSpecialLogfile "PDFCreator.exe: " & tStr

 tStr = GetPDFCreatorApplicationPath & "Languages\TransTool.exe"
 If FileExists(tStr) = True Then
   Set c = GetFileVersion(tStr)
   tStr = "Version: " & c(2) & "; Size: " & Format(FileLen(tStr), "###,###,###,### Bytes")
  Else
   tStr = "-"
 End If
 WriteToSpecialLogfile "TransTool.exe: " & tStr


 WriteToSpecialLogfile "Redmon settings:"
 tStr = "REDMON_PORT"
 WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
 tStr = "REDMON_JOB"
 WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
 tStr = "REDMON_PRINTER"
 WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
 tStr = "REDMON_MACHINE"
 WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
 tStr = "REDMON_USER"
 WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
 tStr = "REDMON_DOCNAME"
 WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
 tStr = "REDMON_FILENAME"
 WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
 tStr = "REDMON_SESSIONID"
 WriteToSpecialLogfile tStr & "=" & Environ$(tStr)
 WriteToSpecialLogfile vbCrLf & "Computer settings:"
 WriteToSpecialLogfile "LoggedOn-UserName=" & GetUsername
 WriteToSpecialLogfile "Computer=" & GetComputerName
 WriteToSpecialLogfile "PDFCreatorPath=" & PDFCreatorPath
 WriteToSpecialLogfile "IsWin9xMe=" & IsWin9xMe
End Sub

Private Sub AnalyzeCommandlineParameters()
 ' Commandswitches
 ' -LOG
 '      Enables logging
 ' -PPDFCREATORPRINTER
 '      Start PDFCreator if PDFSpooler found printer dates
 '      Now know PDFCreator that PDFSpooler call PDFCreator
 ' -SL<Number>
 '      Wait <Number> milliseconds
 '      Example: -SL500
 '      Wait 500 milliseconds
 ' -ST
 '  -STTRUE
 '      Start pdfcreator, after using SL
 Dim cSwitch As String
 If Len(VBA.Command$) > 0 Then
  If UCase$(CommandSwitch("L", False)) = "OG" Then
    enableSpecialLogging = True
   Else
    enableSpecialLogging = False
  End If
  cSwitch = CommandSwitch("SL", True)
  If LenB(cSwitch) > 0 Then
   If IsNumeric(cSwitch) = True Then
    SleepTime = CLng(cSwitch)
   End If
  End If
  If UCase$(CommandSwitch("ST", False)) = "TRUE" Then
    StartPDFCreatorProgram = True
   Else
    StartPDFCreatorProgram = False
  End If
  If UCase$(CommandSwitch("P", False)) = "PDFCREATORPRINTER" Then
    PDFCreatorPrinter = True
   Else
    PDFCreatorPrinter = False
  End If
  cSwitch = CommandSwitch("OPTIONSFILE", True)
  If LenB(cSwitch) > 0 Then
   If FileExists(cSwitch) = True Then
    Optionsfile = cSwitch
   End If
  End If
  ' Check running instance
  If UCase$(CommandSwitch("Check", False)) = "INSTANCE" Then
    CheckInstance = True
   Else
    CheckInstance = False
  End If
  If UCase$(CommandSwitch("NO", False)) = "START" Then
    NoStart = True
   Else
    NoStart = False
  End If
 End If
End Sub

Private Sub InitProgram()
 SleepTime = -1
End Sub

Public Sub CreateMutexPDFSpooler()
 ' Check for a local running instance
 Set mutexLocal = New clsMutex
 If Not mutexLocal.CheckMutex(PDFSpooler_GUID) Then
  mutexLocal.CreateMutex PDFSpooler_GUID
 End If
 ' Check for a global running instance
 Set mutexGlobal = New clsMutex
 If Not mutexGlobal.CheckMutex("Global\" & PDFSpooler_GUID) Then
  mutexGlobal.CreateMutex "Global\" & PDFSpooler_GUID
 End If
End Sub
