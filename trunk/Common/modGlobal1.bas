Attribute VB_Name = "modGlobal1"
Option Explicit

Public Const Uninstall_GUID = "{0001B4FD-9EA3-4D90-A79E-FD14BA3AB01D}"
Public Const PDFCreator_GUID = "{A7332D94-E8FE-40B2-937F-8515FC0FF52F}"
Public Const TransTool_GUID = "{B7BCA0D2-7305-4318-BA7A-01B028D910EB}"
Public Const PDFSpooler_GUID = "{C387A397-047A-4354-AE89-F75B1B550257}"
Public Const UnInst_GUID = "{D95872D0-0DE7-4C01-859C-1BAE47FB1C6B}"
Public Const WM_GUID = "{5E59D3A8-BDDD-E0F0-6D3B-A8E147725130}"
Public Const PaypalPDFCreator = "https://www.paypal.com/cgi-bin/webscr?cmd=_xclick&business=paypal01%40heindoerfer%2ecom&item_name=PDFCreator&no_shipping=2&no_note=1&tax=0&currency_code=EUR&lc=US&bn=PP%2dDonationsBF&charset=UTF%2d8"
Public Const PaypalTransTool = "https://www.paypal.com/cgi-bin/webscr?cmd=_xclick&business=paypal01%40heindoerfer%2ecom&item_name=TransTool&no_shipping=2&no_note=1&tax=0&currency_code=EUR&lc=US&bn=PP%2dDonationsBF&charset=UTF%2d8"
Public Const Homepage = "http://www.pdfforge.org/"
Public Const Sourceforge = "http://www.sourceforge.net/projects/pdfcreator"
Public Const UpdateURL = "http://www.pdfforge.org/products/pdfcreator/update.txt"
Public Const PDFCreatorLogfile = "PDFCreator.log"
Public Const PDFCreatorSpoolDirectory = "PDFCreatorSpool"
Public Const CompatibleLanguageVersion = "0.9.5"

Public CancelPrintfiles As Boolean
Public ChangeDefaultprinter As Boolean
Public CheckInstance As Boolean
Public ConvertedOutputFilename As String
Public enableSpecialLogging As Boolean
Public HelpFile As String
Public IFIsPS As Boolean
Public Languagefile As String
Public LanguagePath As String
Public mutexGlobal As clsMutex
Public mutexLocal As clsMutex
Public NoAbortIfRunning As Boolean
Public NoProcessing As Boolean
Public NoProcessingAtStartup As Boolean
Public NoPSCheck As Boolean
Public NoStart As Boolean
Public Optionsfile As String
Public OwnerPassword As String
Public PDFCreatorINIFile As String
Public PDFCreatorLogfilePath As String
Public PDFCreatorPrinter As Boolean
Public PrinterStop As Boolean
Public PrintFilename As String
Public Printing As Boolean
Public PrintSelectedJobs As Boolean
Public ProgramIsStarted As Boolean
Public ProgramIsVisible As Boolean
Public ReadyConverting As Boolean
Public Restart As Boolean
Public SaveOpenCancel As Boolean
Public SaveOpenFilename As Collection
Public SaveOpenFilterindex As Long
Public SavePasswordsForThisSession As Boolean
Public SecurityIsPossible As Boolean
Public SleepTime As Long
Public StartPDFCreatorProgram As Boolean
Public ShowOnlyLogfile As Boolean
Public ShowOnlyOptions As Boolean
Public UserPassword As String
Public InstanceCounter As Long
Public GhostscriptError As Long
Public ProgramWindowState As Long
Public DeleteIF As Boolean
Public OpenOF As Boolean
Public IsConverted As Boolean
Public StopURLPrinting As Boolean

Public Sub CheckProgramInstances()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tStr As String
50020  tStr = "Check program instances" & vbCrLf & vbCrLf
50030  tStr = tStr & "PDFCreator:" & vbTab & GetCheckProgramInstancesStr(PDFCreator_GUID) & vbCrLf
50040  tStr = tStr & "PDFSpooler:" & vbTab & GetCheckProgramInstancesStr(PDFSpooler_GUID) & vbCrLf
50050  tStr = tStr & "TransTool:" & vbTab & GetCheckProgramInstancesStr(TransTool_GUID) & vbCrLf
50060  MsgBox tStr
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGlobal1", "CheckProgramInstances")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function GetCheckProgramInstancesStr(MutexName As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tStr As String
50020  tStr = ""
50030  Set mutexLocal = New clsMutex
50040  If mutexLocal.CheckMutex(MutexName) = True Then
50050   tStr = "Local"
50060  End If
50070  Set mutexGlobal = New clsMutex
50080  If mutexGlobal.CheckMutex("Global\" & MutexName) = True Then
50090   If LenB(tStr) > 0 Then
50100     tStr = tStr & ", Global"
50110    Else
50120     tStr = "Global"
50130   End If
50140  End If
50150  If LenB(tStr) = 0 Then
50160   tStr = "No instances found."
50170  End If
50180  GetCheckProgramInstancesStr = tStr
50190  Set mutexLocal = Nothing
50200  Set mutexGlobal = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGlobal1", "GetCheckProgramInstancesStr")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetPDFCreatorApplicationPath() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry, tStr As String
50020  Set reg = New clsRegistry
50030  With reg
50040   .hkey = HKEY_LOCAL_MACHINE
50050   .KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50060   tStr = .GetRegistryValue("Inno Setup: App Path")
50070  End With
50080  If LenB(LTrim$(tStr)) = 0 Then
50090   tStr = App.Path
50100  End If
50110  GetPDFCreatorApplicationPath = CompletePath(tStr)
50120  Set reg = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGlobal1", "GetPDFCreatorApplicationPath")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub CheckForUpdate()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim lucYear As Long, curYear As Long, lucMonth As Long, curMonth As Long, lucDay As Long, curDay As Long
50020  Dim lucDate As Date, curDate As Date, diff As Long
50030  Dim upd As clsUpdate
50040
50050  If (Options.UpdateInterval) = 0 Then
50060   Exit Sub
50070  End If
50080
50090  curDate = Now
50100  If Len(Options.LastUpdateCheck) <> 8 Then
50110   SetLastUpdateCeck curDate
50120   Exit Sub
50130  End If
50140  If IsNumeric(Options.LastUpdateCheck) = False Then
50150   SetLastUpdateCeck curDate
50160   Exit Sub
50170  End If
50180
50190  curYear = Year(curDate): curMonth = Month(curDate): curDay = Day(curDate)
50200  lucYear = CLng(Mid$(Options.LastUpdateCheck, 1, 4))
50210  lucMonth = CLng(Mid$(Options.LastUpdateCheck, 5, 2))
50220  lucDay = CLng(Mid$(Options.LastUpdateCheck, 7, 2))
50230
50240  lucDate = DateSerial(lucYear, lucMonth, lucDay)
50250
50260  If (curYear < lucYear) Or (curYear = lucYear And curMonth < lucMonth) Or (curYear = lucYear And curMonth = lucMonth And curDay < lucDay) Then
50270   Options.LastUpdateCheck = Format$(curDate, "YYYYMMDD")
50280   SaveOption Options, "LastUpdateCheck"
50290   Exit Sub
50300  End If
50310
50320  If (Options.UpdateInterval) = 1 Then ' Once a day
50330    diff = DateDiff("d", lucDate, curDate)
50340    If diff >= 1 Then
50350     Set upd = New clsUpdate
50360     upd.CheckForUpdates
50370     SetLastUpdateCeck curDate
50380    End If
50390   ElseIf (Options.UpdateInterval) = 2 Then ' Once a week
50400    diff = DateDiff("d", lucDate, curDate)
50410    If diff >= 7 Then
50420     Set upd = New clsUpdate
50430     upd.CheckForUpdates
50440     SetLastUpdateCeck curDate
50450    End If
50460   Else ' Once a month
50470    diff = DateDiff("m", lucDate, curDate)
50480    If diff >= 1 Then
50490     Set upd = New clsUpdate
50500     upd.CheckForUpdates
50510     SetLastUpdateCeck curDate
50520    End If
50530  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGlobal1", "CheckForUpdate")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub SetLastUpdateCeck(d As Date)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Options.LastUpdateCheck = Format$(d, "YYYYMMDD")
50020  SaveOption Options, "LastUpdateCheck"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGlobal1", "SetLastUpdateCeck")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
