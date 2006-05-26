Attribute VB_Name = "modMain"
Option Explicit

Public PDFCreatorPath As String

Public Sub Main()
 Dim UserName As String, SessionID As Long, _
  UserAppData As String, UserLocalTemp As String, tStr As String, _
  hToken As Long, hProfile As Long, AppPath As String, AppParams As String, _
  stdio As clsStdIO, cinStr As String, SpooltimeSeconds As Double, _
  LoggedInConsole As Boolean, Tempfile As String

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
 Call GetUserLocalDirs(HKEY_CURRENT_USER, UserAppData, UserLocalTemp)
 If IsWin9xMe = False Then
  If GetUserSessionToken(UserName, SessionID, hToken, enableSpecialLogging) = 0 Then
   If LoadProfile(UserName, hToken, hProfile) = 0 Then
    Call GetUserLocalDirs(hProfile, UserAppData, UserLocalTemp)
    LoggedInConsole = True
   End If
  End If
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
   Options = ReadOptions(True, hProfile)
 End If

 WriteToSpecialLogfile ""
 WriteToSpecialLogfile "Options.DirectoryGhostscriptBinaries=" & Options.DirectoryGhostscriptBinaries
 WriteToSpecialLogfile "Options.DirectoryGhostscriptFonts=" & Options.DirectoryGhostscriptFonts
 WriteToSpecialLogfile "Options.DirectoryGhostscriptLibraries=" & Options.DirectoryGhostscriptLibraries
 WriteToSpecialLogfile "Options.DirectoryGhostscriptResource=" & Options.DirectoryGhostscriptResource
 WriteToSpecialLogfile "Options.PrinterTemppath=" & Options.PrinterTemppath

 WriteToSpecialLogfile ""
 WriteToSpecialLogfile "PDFCreatorTempPath=" & Options.PrinterTemppath
 WriteToSpecialLogfile "GetPDFCreatorTempfolder=" & GetPDFCreatorTempfolder

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
        Shell """" & PDFCreatorPath & "PDFCreator.exe"" " & Command$ & " -IF""" & cinStr & """"
       Else
        AppPath = PDFCreatorPath & "PDFCreator.exe"
        AppParams = " " & Command$ & " -IF""" & cinStr & """" ' The leading space is important and cannot be removed.
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
End Sub

Private Sub WriteInfoToSpecialLogfile()
 Dim tStr As String, c As Collection
 WriteToSpecialLogfile "Start: " & Now, True
 WriteToSpecialLogfile "ProgramReleaseStr: " & GetProgramReleaseStr
 
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
