Attribute VB_Name = "modGlobal1"
Option Explicit

Public Const Uninstall_GUID = "{0001B4FD-9EA3-4D90-A79E-FD14BA3AB01D}"
Public Const PDFCreator_GUID = "{A7332D94-E8FE-40B2-937F-8515FC0FF52F}"
Public Const TransTool_GUID = "{B7BCA0D2-7305-4318-BA7A-01B028D910EB}"
Public Const PDFSpooler_GUID = "{C387A397-047A-4354-AE89-F75B1B550257}"
Public Const UnInst_GUID = "{D95872D0-0DE7-4C01-859C-1BAE47FB1C6B}"
Public Const WM_GUID = "{5E59D3A8-BDDD-E0F0-6D3B-A8E147725130}"
Public Const PaypalPDFCreator = "http://www.pdfforge.org/donate"
Public Const PaypalTransTool = "http://www.pdfforge.org/donate"
Public Const Homepage = "http://www.pdfforge.org/"
Public Const Sourceforge = "http://www.sourceforge.net/projects/pdfcreator"
Public Const UpdateURL = "http://www.pdfforge.org/products/pdfcreator/update.txt"
Public Const PDFCreatorLogfile = "PDFCreator.log"
Public Const PDFCreatorSpoolDirectory = "PDFCreatorSpool"
Public Const PDFSpoolerExe = "PDFSpool.exe"
Public Const CompatibleLanguageVersion = "1.0.0"

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
Public ShutDown As Boolean
Public CurrentLanguage As String
Public CurrentPrinterProfile As String
Public InstalledAsServer As Boolean
Public PrinterTemppath As String

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
50010  Dim tStr As String, m As clsMutex
50020  tStr = ""
50030  Set m = New clsMutex
50040  If m.CheckMutex(MutexName) = True Then
50050   tStr = "Local"
50060  End If
50070  If mutexGlobal.CheckMutex("Global\" & MutexName) = True Then
50080   If LenB(tStr) > 0 Then
50090     tStr = tStr & ", Global"
50100    Else
50110     tStr = "Global"
50120   End If
50130  End If
50140  If LenB(tStr) = 0 Then
50150   tStr = "No instances found."
50160  End If
50170  GetCheckProgramInstancesStr = tStr
50180  Set m = Nothing
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

'Public Sub Log2File(Optional value As String = "", Optional filename As String = "C:\PDFCreator_cClose.txt")
''---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
'On Error GoTo ErrPtnr_OnError
''---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
'50010  Dim ff As Long
'50020  ff = FreeFile
'50030  Open filename For Append As #ff
'50040  Print #ff, Now2 & ": " & value
'50050  Close #ff
''---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
'Exit Sub
'ErrPtnr_OnError:
'Select Case ErrPtnr.OnError("modGlobal1", "Log2File")
'Case 0: Resume
'Case 1: Resume Next
'Case 2: Exit Sub
'Case 3: End
'End Select
''---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
'End Sub


