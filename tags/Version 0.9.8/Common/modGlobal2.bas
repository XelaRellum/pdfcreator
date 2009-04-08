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
50020
50030  If enableSpecialLogging = True Then
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
50160     Print #fn, "Start: " & Now
50170     Print #fn, "Windowsversion: " & GetWinVersionStr
50180     Print #fn, "PDFCreator-Revision: " & GetProgramReleaseStr
50190    Else
50200     Open FName For Append As #fn
50210   End If
50220   Print #fn, StatusText
50230   Close #fn
50240  End If
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

Public Function GetSubstFilename2(TokenFilename As String, Optional Preview As Boolean = True, Optional Temppath As String, Optional PostscriptFile As String, Optional hkey1 As hkey = HKEY_CURRENT_USER) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim DateTime As String, Author As String, ClientComputer As String, UserName As String, _
  Computername As String, MyFiles As String, MyDesktop As String, filename As String, _
  title As String, tStr As String, isf As InfoSpoolFile, FilePath As String
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
50310  filename = TokenFilename
50320  filename = Replace(filename, "<DateTime>", DateTime, , , vbTextCompare)
50330  filename = Replace(filename, "<Computername>", Computername, , , vbTextCompare)
50340
50350  filename = Replace(filename, "<ClientComputer>", ClientComputer, , , vbTextCompare)
50360  filename = Replace(filename, "<Username>", UserName, , , vbTextCompare)
50370  filename = Replace(filename, "<Title>", title, , , vbTextCompare)
50380
50390  filename = Replace(filename, "<Author>", Author, , , vbTextCompare)
50400
50410  filename = Replace(filename, "<MyFiles>", CompletePath(MyFiles), , , vbTextCompare)
50420  filename = Replace(filename, "<MyDesktop>", CompletePath(MyDesktop), , , vbTextCompare)
50430
50440  If Options.Counter = 922337203685477@ Then
50450   Options.Counter = 0
50460  End If
50470  Options.Counter = Round(Options.Counter)
50480
50490  filename = Replace(filename, "<Counter>", Format$(Options.Counter + 1, String(15, "0")), , , vbTextCompare)
50500
50510  If LenB(Temppath) > 0 Then
50520    filename = Replace(filename, "<Temp>", CompletePath(Temppath), , , vbTextCompare)
50530   Else
50540    filename = Replace(filename, "<Temp>", CompletePath(GetTempPathReg(hkey1)), , , vbTextCompare)
50550  End If
50560
50570  tStr = "DOCNAME"
50580  If Preview Then
50590    filename = Replace(filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
50600   Else
50610    If LenB(isf.REDMON_DOCNAME) = 0 Then
50620      filename = Replace(filename, "<REDMON_" & tStr & ">", Environ$("REDMON_DOCNAME"), , , vbTextCompare)
50630     Else
50640      filename = Replace(filename, "<REDMON_" & tStr & ">", isf.REDMON_DOCNAME, , , vbTextCompare)
50650    End If
50660  End If
50670  tStr = "DOCNAME_FILE"
50680  If Preview Then
50690    filename = Replace(filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
50700   Else
50710    If LenB(isf.REDMON_DOCNAME) = 0 Then
50720      SplitPath Environ$("REDMON_DOCNAME"), , , , FilePath
50730     Else
50740      SplitPath isf.REDMON_DOCNAME, , , , FilePath
50750    End If
50760    filename = Replace(filename, "<REDMON_" & tStr & ">", FilePath, , , vbTextCompare)
50770  End If
50780  tStr = "DOCNAME_PATH"
50790  If Preview Then
50800    filename = Replace(filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
50810   Else
50820    If LenB(isf.REDMON_DOCNAME) = 0 Then
50830      SplitPath Environ$("REDMON_DOCNAME"), , FilePath
50840     Else
50850      SplitPath isf.REDMON_DOCNAME, , FilePath
50860    End If
50870    filename = Replace(filename, "<REDMON_" & tStr & ">", FilePath, , , vbTextCompare)
50880  End If
50890  tStr = "JOB"
50900  If Preview Then
50910    filename = Replace(filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
50920   Else
50930    If LenB(isf.REDMON_JOB) = 0 Then
50940      filename = Replace(filename, "<REDMON_" & tStr & ">", Environ$("REDMON_JOB"), , , vbTextCompare)
50950     Else
50960      filename = Replace(filename, "<REDMON_" & tStr & ">", isf.REDMON_JOB, , , vbTextCompare)
50970    End If
50980  End If
50990  tStr = "MACHINE"
51000  If Preview Then
51010    filename = Replace(filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
51020   Else
51030    If LenB(isf.REDMON_MACHINE) = 0 Then
51040      filename = Replace(filename, "<REDMON_" & tStr & ">", Replace$(Environ$("REDMON_MACHINE"), "\\", ""), , , vbTextCompare)
51050     Else
51060      filename = Replace(filename, "<REDMON_" & tStr & ">", Replace$(isf.REDMON_MACHINE, "\\", ""), , , vbTextCompare)
51070    End If
51080  End If
51090  tStr = "PORT"
51100  If Preview Then
51110    filename = Replace(filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
51120   Else
51130    If LenB(isf.REDMON_PORT) = 0 Then
51140      filename = Replace(filename, "<REDMON_" & tStr & ">", Environ$("REDMON_PORT"), , , vbTextCompare)
51150     Else
51160      filename = Replace(filename, "<REDMON_" & tStr & ">", isf.REDMON_PORT, , , vbTextCompare)
51170    End If
51180  End If
51190  tStr = "PRINTER"
51200  If Preview Then
51210    filename = Replace(filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
51220   Else
51230    If LenB(isf.REDMON_PRINTER) = 0 Then
51240      filename = Replace(filename, "<REDMON_" & tStr & ">", Environ$("REDMON_PRINTER"), , , vbTextCompare)
51250     Else
51260      filename = Replace(filename, "<REDMON_" & tStr & ">", isf.REDMON_PRINTER, , , vbTextCompare)
51270    End If
51280  End If
51290  tStr = "SESSIONID"
51300  If Preview Then
51310    filename = Replace(filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
51320   Else
51330    If LenB(isf.REDMON_SESSIONID) = 0 Then
51340      filename = Replace(filename, "<REDMON_" & tStr & ">", Environ$("REDMON_SESSIONID"), , , vbTextCompare)
51350     Else
51360      filename = Replace(filename, "<REDMON_" & tStr & ">", isf.REDMON_SESSIONID, , , vbTextCompare)
51370    End If
51380  End If
51390  tStr = "USER"
51400  If Preview Then
51410    filename = Replace(filename, "<REDMON_" & tStr & ">", "'Preview REDMON_" & tStr & "'", , , vbTextCompare)
51420   Else
51430    If LenB(isf.REDMON_USER) = 0 Then
51440      filename = Replace(filename, "<REDMON_" & tStr & ">", Environ$("REDMON_USER"), , , vbTextCompare)
51450     Else
51460      filename = Replace(filename, "<REDMON_" & tStr & ">", isf.REDMON_USER, , , vbTextCompare)
51470    End If
51480  End If
51490
51500  If Options.RemoveSpaces = 1 Then
51510   filename = Trim$(filename)
51520  End If
51530  If Len(filename) >= 4 Then
51540   If InStr(2, filename, "\\") > 0 Then
51550    filename = Mid$(filename, 1, 1) & Replace$(Mid(filename, 2), "\\", "\")
51560   End If
51570  End If
51580  GetSubstFilename2 = filename
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
