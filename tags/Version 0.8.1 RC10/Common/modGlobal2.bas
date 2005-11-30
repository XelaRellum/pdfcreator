Attribute VB_Name = "modGlobal2"
Option Explicit

Public Function GetPDFCreatorTempfolder(Optional Preview As Boolean = False, Optional Temppath As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  GetPDFCreatorTempfolder = ResolveEnvironment(GetSubstFilename2(Options.PrinterTemppath, Preview, Temppath))
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGlobal2", "GetPDFCreatorTempfolder")
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
Select Case ErrPtnr.OnError("modGlobal2", "InstalledAsServer")
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
Select Case ErrPtnr.OnError("modGlobal2", "ProgramIsRunning")
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
Select Case ErrPtnr.OnError("modGlobal2", "WriteToSpecialLogfile")
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
Select Case ErrPtnr.OnError("modGlobal2", "WriteEnvironmentToSpecialLogfile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function GetDocDate(Optional StandardDate As String = "", Optional StandardDateformat As String = "", Optional UseThisdate As String = "") As String
 On Error Resume Next
 Dim tStr As String, DateFormat As String, Usingdate As String

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

 tStr = Format$(Usingdate, DateFormat)
 If LenB(tStr) = 0 Then
  tStr = Usingdate
 End If
 GetDocDate = tStr
End Function

Public Function GetSubstFilename2(TokenFilename As String, Optional Preview As Boolean = True, Optional Temppath As String, Optional PostscriptFile As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim DateTime As String, Author As String, ClientComputer As String, UserName As String, _
  Computername As String, MyFiles As String, MyDesktop As String, Filename As String, _
  Title As String, tStr As String, isf As InfoSpoolFile, FilePath As String
50040
50050  If Len(TokenFilename) = 0 Then
50060   Exit Function
50070  End If
50080
50090  If LenB(PostscriptFile) > 0 Then
50100   isf = ReadInfoSpoolfile(PostscriptFile)
50110  End If
50120
50130  DateTime = GetDocDate("", Options.StandardDateformat, CStr(Now))
50140  If Preview Then
50150   Author = "'Preview Author'"
50160   ClientComputer = "'Preview ClientComputer'"
50170  End If
50180
50190  UserName = GetDocUsername("", True)
50200
50210  Computername = GetComputerName
50220  MyFiles = GetMyFiles
50230  MyDesktop = GetDesktop
50240
50250  If LenB(Environ$("REDMON_MACHINE")) > 0 Then
50260    ClientComputer = Environ$("REDMON_MACHINE")
50270   Else
50280    ClientComputer = GetComputerName
50290  End If
50300
50310  Filename = TokenFilename
50320  Filename = Replace(Filename, "<DateTime>", DateTime, , , vbTextCompare)
50330  Filename = Replace(Filename, "<Computername>", Computername, , , vbTextCompare)
50340
50350  Filename = Replace(Filename, "<ClientComputer>", ClientComputer, , , vbTextCompare)
50360  Filename = Replace(Filename, "<Username>", UserName, , , vbTextCompare)
50370  Filename = Replace(Filename, "<Title>", Title, , , vbTextCompare)
50380
50390  Filename = Replace(Filename, "<Author>", Author, , , vbTextCompare)
50400
50410  Filename = Replace(Filename, "<MyFiles>", CompletePath(MyFiles), , , vbTextCompare)
50420  Filename = Replace(Filename, "<MyDesktop>", CompletePath(MyDesktop), , , vbTextCompare)
50430  If LenB(Temppath) > 0 Then
50440    Filename = Replace(Filename, "<Temp>", CompletePath(Temppath), , , vbTextCompare)
50450   Else
50460    Filename = Replace(Filename, "<Temp>", CompletePath(GetTempPathReg(HKEY_CURRENT_USER)), , , vbTextCompare)
50470  End If
50480
50490  tStr = "DOCNAME"
50500  If Preview Then
50510    Filename = Replace(Filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
50520   Else
50530    If LenB(isf.REDMON_DOCNAME) = 0 Then
50540      Filename = Replace(Filename, "<REDMON_" & tStr & ">", Environ$("REDMON_DOCNAME"), , , vbTextCompare)
50550     Else
50560      Filename = Replace(Filename, "<REDMON_" & tStr & ">", isf.REDMON_DOCNAME, , , vbTextCompare)
50570    End If
50580  End If
50590  tStr = "DOCNAME_FILE"
50600  If Preview Then
50610    Filename = Replace(Filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
50620   Else
50630    If LenB(isf.REDMON_DOCNAME) = 0 Then
50640      SplitPath Environ$("REDMON_DOCNAME"), , , , FilePath
50650     Else
50660      SplitPath isf.REDMON_DOCNAME, , , , FilePath
50670    End If
50680    Filename = Replace(Filename, "<REDMON_" & tStr & ">", FilePath, , , vbTextCompare)
50690  End If
50700  tStr = "DOCNAME_PATH"
50710  If Preview Then
50720    Filename = Replace(Filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
50730   Else
50740    If LenB(isf.REDMON_DOCNAME) = 0 Then
50750      SplitPath Environ$("REDMON_DOCNAME"), , FilePath
50760     Else
50770      SplitPath isf.REDMON_DOCNAME, , FilePath
50780    End If
50790    Filename = Replace(Filename, "<REDMON_" & tStr & ">", FilePath, , , vbTextCompare)
50800  End If
50810  tStr = "JOB"
50820  If Preview Then
50830    Filename = Replace(Filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
50840   Else
50850    If LenB(isf.REDMON_JOB) = 0 Then
50860      Filename = Replace(Filename, "<REDMON_" & tStr & ">", Environ$("REDMON_JOB"), , , vbTextCompare)
50870     Else
50880      Filename = Replace(Filename, "<REDMON_" & tStr & ">", isf.REDMON_JOB, , , vbTextCompare)
50890    End If
50900  End If
50910  tStr = "MACHINE"
50920  If Preview Then
50930    Filename = Replace(Filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
50940   Else
50950    If LenB(isf.REDMON_MACHINE) = 0 Then
50960      Filename = Replace(Filename, "<REDMON_" & tStr & ">", Replace$(Environ$("REDMON_MACHINE"), "\\", ""), , , vbTextCompare)
50970     Else
50980      Filename = Replace(Filename, "<REDMON_" & tStr & ">", Replace$(isf.REDMON_MACHINE, "\\", ""), , , vbTextCompare)
50990    End If
51000  End If
51010  tStr = "PORT"
51020  If Preview Then
51030    Filename = Replace(Filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
51040   Else
51050    If LenB(isf.REDMON_PORT) = 0 Then
51060      Filename = Replace(Filename, "<REDMON_" & tStr & ">", Environ$("REDMON_PORT"), , , vbTextCompare)
51070     Else
51080      Filename = Replace(Filename, "<REDMON_" & tStr & ">", isf.REDMON_PORT, , , vbTextCompare)
51090    End If
51100  End If
51110  tStr = "PRINTER"
51120  If Preview Then
51130    Filename = Replace(Filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
51140   Else
51150    If LenB(isf.REDMON_PRINTER) = 0 Then
51160      Filename = Replace(Filename, "<REDMON_" & tStr & ">", Environ$("REDMON_PRINTER"), , , vbTextCompare)
51170     Else
51180      Filename = Replace(Filename, "<REDMON_" & tStr & ">", isf.REDMON_PRINTER, , , vbTextCompare)
51190    End If
51200  End If
51210  tStr = "SESSIONID"
51220  If Preview Then
51230    Filename = Replace(Filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
51240   Else
51250    If LenB(isf.REDMON_SESSIONID) = 0 Then
51260      Filename = Replace(Filename, "<REDMON_" & tStr & ">", Environ$("REDMON_SESSIONID"), , , vbTextCompare)
51270     Else
51280      Filename = Replace(Filename, "<REDMON_" & tStr & ">", isf.REDMON_SESSIONID, , , vbTextCompare)
51290    End If
51300  End If
51310  tStr = "USER"
51320  If Preview Then
51330    Filename = Replace(Filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
51340   Else
51350    If LenB(isf.REDMON_USER) = 0 Then
51360      Filename = Replace(Filename, "<REDMON_" & tStr & ">", Environ$("REDMON_USER"), , , vbTextCompare)
51370     Else
51380      Filename = Replace(Filename, "<REDMON_" & tStr & ">", isf.REDMON_USER, , , vbTextCompare)
51390    End If
51400  End If
51410
51420  If Options.RemoveSpaces = 1 Then
51430   Filename = Trim$(Filename)
51440  End If
51450  If Len(Filename) >= 4 Then
51460   If InStr(2, Filename, "\\") > 0 Then
51470    Filename = Mid$(Filename, 1, 1) & Replace$(Mid(Filename, 2), "\\", "\")
51480   End If
51490  End If
51500  GetSubstFilename2 = Filename
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGlobal2", "GetSubstFilename2")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub PrintTestpage(Optional f As Form)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim TestPSPage As String, fn As Long, Filename As String, tStr As String, _
  c As Collection
50030  If Not f Is Nothing Then
50040   f.Timer1.Enabled = False
50050  End If
50060  TestPSPage = GetTestpageFromRessource
50070  TestPSPage = Replace(TestPSPage, "[INFOTITLE]", LanguageStrings.OptionsTestpage, , 1, vbTextCompare)
50080  TestPSPage = Replace(TestPSPage, "[INFORELEASE]", App.Title & " " & GetProgramReleaseStr, , 1, vbTextCompare)
50090  TestPSPage = Replace(TestPSPage, "[INFODATE]", Now, , 1, vbTextCompare)
50100  TestPSPage = Replace(TestPSPage, "[INFOAUTHORS]", "Philip Chinery, Frank Heind\224rfer", , 1, vbTextCompare)
50110  TestPSPage = Replace(TestPSPage, "[INFOHOMEPAGE]", Homepage, , 1, vbTextCompare)
50120  tStr = GetPDFCreatorApplicationPath & "PDFCreator.exe"
50130  If FileExists(tStr) = True Then
50140    Set c = GetFileVersion(tStr)
50150    tStr = "Version: " & c(2) & "; Size: " & Format(FileLen(tStr), "###,###,###,### Bytes")
50160   Else
50170    tStr = ""
50180  End If
50190  TestPSPage = Replace(TestPSPage, "[INFOPDFCREATOR]", tStr, , 1, vbTextCompare)
50200
50210  tStr = CompletePath(GetSystemDirectory()) & "PDFSpooler.exe"
50220  If FileExists(tStr) = True Then
50230    Set c = GetFileVersion(tStr)
50240    tStr = "Version: " & c(2) & "; Size: " & Format(FileLen(tStr), "###,###,###,### Bytes")
50250   Else
50260    tStr = ""
50270  End If
50280  TestPSPage = Replace(TestPSPage, "[INFOPDFSPOOLER]", tStr, , 1, vbTextCompare)
50290
50300  tStr = GetPDFCreatorApplicationPath & "Languages\Transtool.exe"
50310  If FileExists(tStr) = True Then
50320    Set c = GetFileVersion(tStr)
50330    tStr = "Version: " & c(2) & "; Size: " & Format(FileLen(tStr), "###,###,###,### Bytes")
50340   Else
50350    tStr = ""
50360  End If
50370  TestPSPage = Replace(TestPSPage, "[INFOTRANSTOOL]", tStr, , 1, vbTextCompare)
50380
50390  TestPSPage = Replace(TestPSPage, "[INFOCOMPUTER]", GetComputerName, , 1, vbTextCompare)
50400  tStr = GetWinVersionStr
50410  TestPSPage = Replace(TestPSPage, "[INFOWINDOWS]", _
  Mid(tStr, 1, IIf(InStr(1, tStr, "[") > 0, InStr(1, tStr, "[") - 1, Len(tStr))), 1, vbTextCompare)
50430
50440  fn = FreeFile
50450  tStr = CompletePath(GetPDFCreatorTempfolder) & PDFCreatorSpoolDirectory
50460  If DirExists(tStr) = False Then
50470   MakePath tStr
50480  End If
50490  Filename = GetTempFile(tStr, "~PS")
50500  Open Filename For Output As fn
50510  Print #fn, TestPSPage
50520  Close #fn
50530  If Not f Is Nothing Then
50540   f.CheckPrintJobs
50550   f.Timer1.Enabled = True
50560  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGlobal2", "PrintTestpage")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function IsPrintable(Filename As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Ext As String, reg As clsRegistry, tStr As String
50020  IsPrintable = False
50030  SplitPath Filename, , , , , Ext
50040  Set reg = New clsRegistry
50050  reg.hkey = HKEY_CLASSES_ROOT
50060  reg.KeyRoot = "." & Ext
50070  If reg.KeyExists Then
50080   tStr = reg.GetRegistryValue("")
50090   If LenB(tStr) > 0 Then
50100    reg.KeyRoot = tStr
50110    reg.Subkey = "shell\print\command"
50120    tStr = reg.GetRegistryValue("")
50130    If LenB(tStr) > 0 Then
50140     IsPrintable = True
50150    End If
50160   End If
50170  End If
50180  Set reg = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGlobal2", "IsPrintable")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetTestpageFromRessource() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  GetTestpageFromRessource = _
  Replace(StrConv(LoadResData(101, "TESTPAGE"), vbUnicode), vbCrLf, vbLf, , , vbBinaryCompare)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGlobal2", "GetTestpageFromRessource")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
