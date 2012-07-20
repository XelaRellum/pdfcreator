Attribute VB_Name = "modGlobal2"
Option Explicit

Public Function GetPDFCreatorTempfolder() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  GetPDFCreatorTempfolder = CompletePath(GetTempPathApi) & "PDFCreator\"
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

Public Function GetPDFCreatorSpoolDirectory()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim spoolDirectory As String
50020  If InstalledAsServer Then
50030    spoolDirectory = CompletePath(GetWindowsDirectory) & "Temp\PDFCreator\" & PDFCreatorSpoolDirectory
50040   Else
50050    spoolDirectory = GetPDFCreatorTempfolder & PDFCreatorSpoolDirectory
50060  End If
50070  If DirExists(spoolDirectory) = False Then
50080   MakePath spoolDirectory
50090  End If
50100  GetPDFCreatorSpoolDirectory = CompletePath(spoolDirectory)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGlobal2", "GetPDFCreatorSpoolDirectory")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function CheckInstalledAsServer() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry
50020  CheckInstalledAsServer = False
50030  Set reg = New clsRegistry
50040  With reg
50050   .hkey = HKEY_LOCAL_MACHINE
50060   .KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50070   If .GetRegistryValue("PDFServer") = 1 Then
50080    CheckInstalledAsServer = True
50090   End If
50100  End With
50110  Set reg = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGlobal2", "CheckInstalledAsServer")
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
50040   Path = LTrim(Environ$("SystemDrive"))
50050   If LenB(Path) = 0 Then
50060    SplitPath LTrim(Environ$("Windir")), Drive
50070    Path = Drive
50080    If LenB(Path) < 2 Then
50090     Path = "c:\"
50100    End If
50110   End If
50120   Path = CompletePath(Path) & "Temp\"
50130   If Not DirExists(Path) Then
50140    MakePath Path
50150   End If
50160   FName = Path & "PDFCreator-Errorlog.txt"
50170   fn = FreeFile
50180   If FileExists(FName) = False Or CreateFile = True Then
50190     Open FName For Output As #fn
50200     Print #fn, "Start: " & Now
50210     Print #fn, "Windowsversion: " & GetWinVersionStr
50220     Print #fn, "PDFCreator-Revision: " & GetProgramReleaseStr
50230    Else
50240     Open FName For Append As #fn
50250   End If
50260   Print #fn, StatusText
50270   Close #fn
50280  End If
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
50030   Path = LTrim(Environ$("SystemDrive"))
50040   If LenB(Path) = 0 Then
50050    SplitPath LTrim(Environ$("Windir")), Drive
50060    Path = Drive
50070    If LenB(Path) < 2 Then
50080     Path = "c:\"
50090    End If
50100   End If
50110   Path = CompletePath(Path) & "Temp\"
50120   If Not DirExists(Path) Then
50130    MakePath Path
50140   End If
50150   FName = Path & "PDFCreator-Errorlog.txt"
50160   fn = FreeFile
50170   i = 1
50180   Open FName For Append As #fn
50190   Print #fn, "Environment:"
50200   While Environ$(i) <> ""
50210    Print #fn, Environ$(i)
50220    DoEvents
50230    i = i + 1
50240   Wend
50250   Close #fn
50260  End If
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

Public Function GetSubstFilename2(TokenFilename As String, Optional Preview As Boolean = True, Optional Temppath As String, Optional InfoSpoolFileName As String, Optional bReplaceForbiddenChars As Boolean = True) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim DateTime As String, Author As String, ClientComputer As String, ClientUsername As String, UserName As String, _
  Computername As String, MyFiles As String, MyDesktop As String, filename As String, _
  Title As String, tStr As String, FilePath As String
50040  Dim isf As clsInfoSpoolFile
50050
50060  If Len(TokenFilename) = 0 Then
50070   Exit Function
50080  End If
50090
50100  Set isf = New clsInfoSpoolFile
50110  If LenB(InfoSpoolFileName) > 0 Then
50120   isf.ReadInfoFile InfoSpoolFileName
50130  End If
50140
50150  DateTime = GetDocDate("", Options.StandardDateformat, CStr(Now))
50160  Computername = GetComputerName
50170
50180  MyFiles = GetMyFiles
50190  MyDesktop = GetDesktop
50200
50210  filename = TokenFilename
50220  filename = Replace(filename, "<DateTime>", DateTime, , , vbTextCompare)
50230  filename = Replace(filename, "<Computername>", Computername, , , vbTextCompare)
50240
50250  filename = Replace(filename, "<Username>", GetUsername, , , vbTextCompare)
50260
50270  filename = Replace(filename, "<MyFiles>", CompletePath(MyFiles), , , vbTextCompare)
50280  filename = Replace(filename, "<MyDesktop>", CompletePath(MyDesktop), , , vbTextCompare)
50290
50300  If Options.Counter = 922337203685477@ Then
50310   Options.Counter = 0
50320  End If
50330  Options.Counter = Round(Options.Counter)
50340
50350  filename = Replace(filename, "<Counter>", Format$(Options.Counter + 1, String(15, "0")), , , vbTextCompare)
50360
50370  If LenB(Temppath) > 0 Then
50380    filename = Replace(filename, "<Temp>", CompletePath(Temppath), , , vbTextCompare)
50390   Else
50400    filename = Replace(filename, "<Temp>", CompletePath(GetTempPathApi), , , vbTextCompare)
50410  End If
50420
50430  tStr = "Author"
50440  If Options.UseStandardAuthor = 1 Then
50450    Author = Options.StandardAuthor
50460   Else
50470    If Preview = True Then
50480      filename = Replace(filename, "<" & tStr & ">", "'Preview " & tStr & "'", , , vbTextCompare)
50490     Else
50500      filename = Replace(filename, "<" & tStr & ">", isf.FirstUserName, , , vbTextCompare)
50510    End If
50520  End If
50530
50540  tStr = "Title"
50550  If Preview Then
50560    filename = Replace(filename, "<" & tStr & ">", "'Preview " & tStr & "'", , , vbTextCompare)
50570   Else
50580    If LenB(isf.FirstDocumentTitle) > 0 Then
50590     filename = Replace(filename, "<" & tStr & ">", isf.FirstDocumentTitle, , , vbTextCompare)
50600    End If
50610  End If
50620  tStr = "DocumentFilename"
50630  If Preview = True Then
50640    filename = Replace(filename, "<" & tStr & ">", "'Preview " & tStr & "'", , , vbTextCompare)
50650   Else
50660    SplitPath isf.FirstDocumentTitle, , , , FilePath
50670    filename = Replace(filename, "<" & tStr & ">", FilePath, , , vbTextCompare)
50680  End If
50690  tStr = "DocumentPath"
50700  If Preview = True Then
50710    filename = Replace(filename, "<" & tStr & ">", "'Preview " & tStr & "'", , , vbTextCompare)
50720   Else
50730    SplitPath isf.FirstDocumentTitle, , FilePath
50740    filename = Replace(filename, "<" & tStr & ">", FilePath, , , vbTextCompare)
50750  End If
50760  tStr = "JobID"
50770  If Preview Then
50780    filename = Replace(filename, "<" & tStr & ">", "'Preview " & tStr & "'", , , vbTextCompare)
50790   Else
50800    If LenB(isf.FirstJobID) > 0 Then
50810     filename = Replace(filename, "<" & tStr & ">", isf.FirstJobID, , , vbTextCompare)
50820    End If
50830  End If
50840  tStr = "ClientComputer"
50850  If Preview Then
50860    filename = Replace(filename, "<" & tStr & ">", "'Preview " & tStr & "'", , , vbTextCompare)
50870   Else
50880    If LenB(isf.FirstClientComputer) > 0 Then
50890     filename = Replace(filename, "<" & tStr & ">", isf.FirstClientComputer, , , vbTextCompare)
50900    End If
50910  End If
50920  tStr = "PrinterName"
50930  If Preview Then
50940    filename = Replace(filename, "<" & tStr & ">", "'Preview " & tStr & "'", , , vbTextCompare)
50950   Else
50960    If LenB(isf.FirstPrinterName) > 0 Then
50970     filename = Replace(filename, "<" & tStr & ">", isf.FirstPrinterName, , , vbTextCompare)
50980    End If
50990  End If
51000  tStr = "SessionID"
51010  If Preview Then
51020    filename = Replace(filename, "<" & tStr & ">", "'Preview " & tStr & "'", , , vbTextCompare)
51030   Else
51040    If LenB(isf.FirstSessionID) > 0 Then
51050     filename = Replace(filename, "<" & tStr & ">", isf.FirstSessionID, , , vbTextCompare)
51060    End If
51070  End If
51080
51090  If Options.RemoveSpaces = 1 Then
51100   filename = Trim$(filename)
51110  End If
51120  If bReplaceForbiddenChars Then
51130   filename = ReplaceForbiddenChars(filename)
51140  End If
51150
51160  If Len(filename) >= 4 Then
51170   If InStr(2, filename, "\\") > 0 Then
51180    filename = Mid$(filename, 1, 1) & Replace$(Mid(filename, 2), "\\", "\")
51190   End If
51200  End If
51210  GetSubstFilename2 = filename
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

Public Sub SavePrinterProfiles(PrinterProfiles As Collection)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry, i As Long, SubKey As String, value As String
50020
50030  Set reg = New clsRegistry
50040
50050  If InstalledAsServer Then
50060    reg.hkey = HKEY_LOCAL_MACHINE
50070   Else
50080    reg.hkey = HKEY_CURRENT_USER
50090  End If
50100
50110  reg.KeyRoot = "Software\PDFCreator"
50120  SubKey = "Printers"
50130  If reg.KeyExists Then
50140   reg.DeleteKey SubKey
50150  End If
50160  reg.SubKey = SubKey
50170  reg.CreateKey
50180
50190  For i = 1 To PrinterProfiles.Count
50200   If LCase$(PrinterProfiles(i)(1)) = LCase$(LanguageStrings.OptionsProfileDefaultName) Then
50210     value = ""
50220    Else
50230     value = PrinterProfiles(i)(1)
50240   End If
50250   reg.SetRegistryValue PrinterProfiles(i)(0), value, REG_SZ
50260  Next i
50270  Set reg = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGlobal2", "SavePrinterProfiles")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function GetGUID() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim udtGUID As GUID
50020
50030  If (CoCreateGuid(udtGUID) = 0) Then
50040   GetGUID = _
   String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & _
   String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & _
   String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
   IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
   IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
   IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
   IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
   IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
   IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
   IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
   IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
50160  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGlobal2", "GetGUID")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
