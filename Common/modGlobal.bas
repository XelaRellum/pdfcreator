Attribute VB_Name = "modGlobal"
Option Explicit

Public Const Uninstall_GUID = "{0001B4FD-9EA3-4D90-A79E-FD14BA3AB01D}"
Public Const PDFCreator_GUID = "{A7332D94-E8FE-40B2-937F-8515FC0FF52F}"
Public Const TransTool_GUID = "{B7BCA0D2-7305-4318-BA7A-01B028D910EB}"
Public Const PDFSpooler_GUID = "{C387A397-047A-4354-AE89-F75B1B550257}"
Public Const UnInst_GUID = "{D95872D0-0DE7-4C01-859C-1BAE47FB1C6B}"
Public Const Paypal = "http://www.paypal.com/xclick/business=paypal01%40heindoerfer.com&item_name=PDFCreator&no_note=1&tax=0&currency_code=EUR"
Public Const Homepage = "http://www.pdfcreator.de.vu"
Public Const Sourceforge = "http://www.sourceforge.net/projects/pdfcreator"
Public Const UpdateURL = "http://www.pdfcreator.de.vu/update.txt"
Public Const PDFCreatorLogfile = "PDFCreator.log"
Public Const PDFCreatorSpoolDirectory = "PDFCreatorSpool"
Public Const CompatibleLanguageVersion = "0.8.0"

Public PDFCreatorINIFile As String, _
       PrinterStop As Boolean, _
       Printing As Boolean, _
       PDFCreatorLogfilePath As String, _
       SavePasswordsForThisSession As Boolean, _
       OwnerPassword As String, _
       UserPassword As String, _
       ChangeDefaultprinter As Boolean, _
       SecurityIsPossible As Boolean, _
       Restart As Boolean, _
       SaveOpenCancel As Boolean, _
       SaveOpenFilename As Collection, _
       SaveOpenFilterindex As Long, _
       enableSpecialLogging As Boolean, _
       ShowOnlyLogfile As Boolean, _
       ShowOnlyOptions As Boolean, _
       PrintSelectedJobs As Boolean, _
       NoAbortIfRunning As Boolean, _
       NoProcessing As Boolean, _
       PDFCreatorPrinter As Boolean, _
       SleepTime As Long, _
       StartPDFcreatorProgram As Boolean, _
       PrintFilename As String, _
       Optionsfile As String, _
       CancelPrintfiles As Boolean

Public Function GetPDFCreatorApplicationPath() As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim reg As clsRegistry, tStr As String
50050  Set reg = New clsRegistry
50060  With reg
50070   .hkey = HKEY_LOCAL_MACHINE
50080   .KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50090   tStr = .GetRegistryValue("Inno Setup: App Path")
50100  End With
50110  If LenB(LTrim$(tStr)) = 0 Then
50120   tStr = App.Path
50130  End If
50140  GetPDFCreatorApplicationPath = CompletePath(tStr)
50150  Set reg = Nothing
50160 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50170 Exit Function
ErrPtnr_OnError:
50191 Select Case ErrPtnr.OnError("modGlobal", "GetPDFCreatorApplicationPath")
      Case 0: Resume
50210 Case 1: Resume Next
50220 Case 2: Exit Function
50230 Case 3: End
50240 End Select
50250 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetPDFCreatorTempfolder() As String
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  GetPDFCreatorTempfolder = Options.PrinterTemppath
50050 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50060 Exit Function
ErrPtnr_OnError:
50081 Select Case ErrPtnr.OnError("modGlobal", "GetPDFCreatorTempfolder")
      Case 0: Resume
50100 Case 1: Resume Next
50110 Case 2: Exit Function
50120 Case 3: End
50130 End Select
50140 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function InstalledAsServer() As Boolean
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim reg As clsRegistry
50050  InstalledAsServer = False
50060  Set reg = New clsRegistry
50070  With reg
50080   .hkey = HKEY_LOCAL_MACHINE
50090   .KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50100   If .GetRegistryValue("PDFServer") = 1 Then
50110    InstalledAsServer = True
50120   End If
50130  End With
50140  Set reg = Nothing
50150 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50160 Exit Function
ErrPtnr_OnError:
50181 Select Case ErrPtnr.OnError("modGlobal", "InstalledAsServer")
      Case 0: Resume
50200 Case 1: Resume Next
50210 Case 2: Exit Function
50220 Case 3: End
50230 End Select
50240 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function ProgramIsRunning(GUIDStr As String) As Boolean
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim mutex As clsMutex
50050  Set mutex = New clsMutex
50060  ProgramIsRunning = mutex.CheckMutex(GUIDStr)
50070  Set mutex = Nothing
50080 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50090 Exit Function
ErrPtnr_OnError:
50111 Select Case ErrPtnr.OnError("modGlobal", "ProgramIsRunning")
      Case 0: Resume
50130 Case 1: Resume Next
50140 Case 2: Exit Function
50150 Case 3: End
50160 End Select
50170 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub WriteToSpecialLogfile(StatusText As String, Optional CreateFile As Boolean = False)
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim FName As String, fn As Long, Path As String, Drive As String
50050  If enableSpecialLogging = True Then
50060
50070   Path = LTrim(Environ$("Systemdrive"))
50080   If LenB(Path) = 0 Then
50090    SplitPath LTrim(Environ$("Windir")), Drive
50100    Path = Drive
50110    If LenB(Path) < 2 Then
50120     Path = "c:\"
50130    End If
50140   End If
50150   FName = CompletePath(Path) & "PDFCreator-Errorlog.txt"
50160   fn = FreeFile
50170   If FileExists(FName) = False Or CreateFile = True Then
50180     Open FName For Output As #fn
50190     Print #fn, "Windowsversion: " & GetWinVersionStr
50200     Print #fn, "PDFCreator-Revision: " & GetProgramReleaseStr
50210    Else
50220     Open FName For Append As #fn
50230   End If
50240   Print #fn, StatusText
50250   Close #fn
50260  End If
50270 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50280 Exit Sub
ErrPtnr_OnError:
50301 Select Case ErrPtnr.OnError("modGlobal", "WriteToSpecialLogfile")
      Case 0: Resume
50320 Case 1: Resume Next
50330 Case 2: Exit Sub
50340 Case 3: End
50350 End Select
50360 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub


