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
Public Const CompatibleLanguageVersion = "0.8.1"

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
       StartPDFCreatorProgram As Boolean, _
       PrintFilename As String, _
       Optionsfile As String, _
       CancelPrintfiles As Boolean

Public HelpFile As String, LanguagePath As String, Languagefile As String, IFIsPS As Boolean, _
 mutexLocal As clsMutex, mutexGlobal As clsMutex, CheckInstance As Boolean, NoStart As Boolean

Public Function GetPDFCreatorApplicationPath() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry, tstr As String
50020  Set reg = New clsRegistry
50030  With reg
50040   .hkey = HKEY_LOCAL_MACHINE
50050   .KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50060   tstr = .GetRegistryValue("Inno Setup: App Path")
50070  End With
50080  If LenB(LTrim$(tstr)) = 0 Then
50090   tstr = App.Path
50100  End If
50110  GetPDFCreatorApplicationPath = CompletePath(tstr)
50120  Set reg = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGlobal", "GetPDFCreatorApplicationPath")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetPDFCreatorTempfolder(Optional Preview As Boolean = False, Optional Temppath As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  GetPDFCreatorTempfolder = ResolveEnvironment(GetSubstFilename2(Options.PrinterTemppath, Preview, Temppath))
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGlobal", "GetPDFCreatorTempfolder")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function InstalledAsServer() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry
50020  InstalledAsServer = False
50030  Set reg = New clsRegistry
50040  With reg
50050   .hkey = HKEY_LOCAL_MACHINE
50060   .KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50070   If .GetRegistryValue("PDFServer") = 1 Then
50080    InstalledAsServer = True
50090   End If
50100  End With
50110  Set reg = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGlobal", "InstalledAsServer")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function ProgramIsRunning(GUIDStr As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim mutex As clsMutex
50020  Set mutex = New clsMutex
50030  ProgramIsRunning = mutex.CheckMutex(GUIDStr)
50040  Set mutex = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGlobal", "ProgramIsRunning")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub WriteToSpecialLogfile(StatusText As String, Optional CreateFile As Boolean = False)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim FName As String, fn As Long, Path As String, Drive As String
50020  If enableSpecialLogging = True Then
50030
50040   Path = LTrim(Environ$("Systemdrive"))
50050   If LenB(Path) = 0 Then
50060    SplitPath LTrim(Environ$("Windir")), Drive
50070    Path = Drive
50080    If LenB(Path) < 2 Then
50090     Path = "c:\"
50100    End If
50110   End If
50120   FName = CompletePath(Path) & "PDFCreator-Errorlog.txt"
50130   fn = FreeFile
50140   If FileExists(FName) = False Or CreateFile = True Then
50150     Open FName For Output As #fn
50160     Print #fn, "Windowsversion: " & GetWinVersionStr
50170     Print #fn, "PDFCreator-Revision: " & GetProgramReleaseStr
50180    Else
50190     Open FName For Append As #fn
50200   End If
50210   Print #fn, StatusText
50220   Close #fn
50230  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGlobal", "WriteToSpecialLogfile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub WriteEnvironmentToSpecialLogfile()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim FName As String, fn As Long, Path As String, Drive As String, i As Long
50020  If enableSpecialLogging = True Then
50030   Path = LTrim(Environ$("Systemdrive"))
50040   If LenB(Path) = 0 Then
50050    SplitPath LTrim(Environ$("Windir")), Drive
50060    Path = Drive
50070    If LenB(Path) < 2 Then
50080     Path = "c:\"
50090    End If
50100   End If
50110   FName = CompletePath(Path) & "PDFCreator-Errorlog.txt"
50120   fn = FreeFile
50130   i = 1
50140   Open FName For Append As #fn
50150   Print #fn, "Environment:"
50160   While Environ$(i) <> ""
50170    Print #fn, Environ$(i)
50180    DoEvents
50190    i = i + 1
50200   Wend
50210   Close #fn
50220  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGlobal", "WriteEnvironmentToSpecialLogfile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function GetDocDate(Optional StandardDate As String = "", Optional StandardDateformat As String = "", Optional UseThisdate As String = "") As String
 On Error Resume Next
 Dim tstr As String, DateFormat As String, Usingdate As String

 If LenB(Trim$(StandardDate)) = 0 Then ' No standard date
   Usingdate = UseThisdate
  Else
   If LenB(RemoveLeadingAndTrailingQuotes(Trim$(StandardDate))) = 0 Then 'Empty date
     Usingdate = ""
    Else
     Usingdate = StandardDate
   End If
 End If

 If Len(StandardDateformat) > 0 Then
   DateFormat = StandardDateformat
  Else
   DateFormat = "YYYYMMDDHHNNSS"
 End If

 tstr = Format$(Usingdate, DateFormat)
 If LenB(tstr) = 0 Then
  tstr = Usingdate
 End If
 GetDocDate = tstr
End Function

Public Function GetSubstFilename2(TokenFilename As String, Optional Preview As Boolean = True, Optional Temppath As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim DateTime As String, Author As String, ClientComputer As String, UserName As String, _
  Computername As String, MyFiles As String, MyDesktop As String, Filename As String, _
  Title As String, tstr As String
50040
50050  If Len(TokenFilename) = 0 Then
50060   Exit Function
50070  End If
50080
50090  DateTime = GetDocDate("", Options.StandardDateformat, CStr(Now))
50100  If Preview Then
50110   Author = "'Preview Author'"
50120   ClientComputer = "'Preview ClientComputer'"
50130  End If
50140
50150  UserName = GetUsername
50160
50170  Computername = GetComputerName
50180  MyFiles = GetMyFiles
50190  MyDesktop = GetDesktop
50200
50210  Filename = TokenFilename
50220  Filename = Replace(Filename, "<DateTime>", DateTime, , , vbTextCompare)
50230  Filename = Replace(Filename, "<Computername>", Computername, , , vbTextCompare)
50240
50250  Filename = Replace(Filename, "<ClientComputer>", ClientComputer, , , vbTextCompare)
50260  Filename = Replace(Filename, "<Username>", UserName, , , vbTextCompare)
50270  Filename = Replace(Filename, "<Title>", Title, , , vbTextCompare)
50280
50290  Filename = Replace(Filename, "<Author>", Author, , , vbTextCompare)
50300
50310  Filename = Replace(Filename, "<MyFiles>", CompletePath(MyFiles), , , vbTextCompare)
50320  Filename = Replace(Filename, "<MyDesktop>", CompletePath(MyDesktop), , , vbTextCompare)
50330  If LenB(Temppath) > 0 Then
50340    Filename = Replace(Filename, "<Temp>", CompletePath(Temppath), , , vbTextCompare)
50350   Else
50360    Filename = Replace(Filename, "<Temp>", CompletePath(GetTempPathReg(HKEY_CURRENT_USER)), , , vbTextCompare)
50370  End If
50380
50390  tstr = "DOCNAME"
50400  If Preview Then
50410    Filename = Replace(Filename, "<REDMON_" & tstr & ">", "'Preview REDMON_" & tstr & "'", , , vbTextCompare)
50420   Else
50430    Filename = Replace(Filename, "<REDMON_" & tstr & ">", Environ$("REDMON_DOCNAME"), , , vbTextCompare)
50440  End If
50450  tstr = "JOB"
50460  If Preview Then
50470    Filename = Replace(Filename, "<REDMON_" & tstr & ">", "'Preview REDMON_" & tstr & "'", , , vbTextCompare)
50480   Else
50490    Filename = Replace(Filename, "<REDMON_" & tstr & ">", Environ$("REDMON_JOB"), , , vbTextCompare)
50500  End If
50510  tstr = "MACHINE"
50520  If Preview Then
50530    Filename = Replace(Filename, "<REDMON_" & tstr & ">", "'Preview REDMON_" & tstr & "'", , , vbTextCompare)
50540   Else
50550    Filename = Replace(Filename, "<REDMON_" & tstr & ">", Environ$("REDMON_MACHINE"), , , vbTextCompare)
50560  End If
50570  tstr = "PORT"
50580  If Preview Then
50590    Filename = Replace(Filename, "<REDMON_" & tstr & ">", "'Preview REDMON_" & tstr & "'", , , vbTextCompare)
50600   Else
50610    Filename = Replace(Filename, "<REDMON_" & tstr & ">", Environ$("REDMON_PORT"), , , vbTextCompare)
50620  End If
50630  tstr = "PRINTER"
50640  If Preview Then
50650    Filename = Replace(Filename, "<REDMON_" & tstr & ">", "'Preview REDMON_" & tstr & "'", , , vbTextCompare)
50660   Else
50670    Filename = Replace(Filename, "<REDMON_" & tstr & ">", Environ$("REDMON_PRINTER"), , , vbTextCompare)
50680  End If
50690  tstr = "SESSIONID"
50700  If Preview Then
50710    Filename = Replace(Filename, "<REDMON_" & tstr & ">", "'Preview REDMON_" & tstr & "'", , , vbTextCompare)
50720   Else
50730    Filename = Replace(Filename, "<REDMON_" & tstr & ">", Environ$("REDMON_SESSIONID"), , , vbTextCompare)
50740  End If
50750  tstr = "USER"
50760  If Preview Then
50770    Filename = Replace(Filename, "<REDMON_" & tstr & ">", "'Preview REDMON_" & tstr & "'", , , vbTextCompare)
50780   Else
50790    Filename = Replace(Filename, "<REDMON_" & tstr & ">", Environ$("REDMON_USER"), , , vbTextCompare)
50800  End If
50810
50820  If Options.RemoveSpaces = 1 Then
50830   Filename = Trim$(Filename)
50840  End If
50850  If Len(Filename) > 4 Then
50860   If InStr(2, Filename, "\\") > 0 Then
50870    Filename = Mid$(Filename, 1, 1) & Replace$(Mid(Filename, 2), "\\", "\")
50880   End If
50890  End If
50900  GetSubstFilename2 = Filename
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGlobal", "GetSubstFilename2")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub CheckProgramInstances()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tstr As String
50020  tstr = "Check program instances" & vbCrLf & vbCrLf
50030  tstr = tstr & "PDFCreator:" & vbTab & GetCheckProgramInstancesStr(PDFCreator_GUID) & vbCrLf
50040  tstr = tstr & "PDFSpooler:" & vbTab & GetCheckProgramInstancesStr(PDFSpooler_GUID) & vbCrLf
50050  tstr = tstr & "TransTool:" & vbTab & GetCheckProgramInstancesStr(TransTool_GUID) & vbCrLf
50060  MsgBox tstr
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGlobal", "CheckProgramInstances")
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
50010  Dim tstr As String
50020  tstr = ""
50030  Set mutexLocal = New clsMutex
50040  If mutexLocal.CheckMutex(MutexName) = True Then
50050   tstr = "Local"
50060  End If
50070  Set mutexGlobal = New clsMutex
50080  If mutexGlobal.CheckMutex("Global\" & MutexName) = True Then
50090   If LenB(tstr) > 0 Then
50100     tstr = tstr & ", Global"
50110    Else
50120     tstr = "Global"
50130   End If
50140  End If
50150  If LenB(tstr) = 0 Then
50160   tstr = "No instances found."
50170  End If
50180  GetCheckProgramInstancesStr = tstr
50190  Set mutexLocal = Nothing
50200  Set mutexGlobal = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGlobal", "GetCheckProgramInstancesStr")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

