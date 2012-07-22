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
50030    spoolDirectory = GetSystemTempPath & "PDFCreator\" & PDFCreatorSpoolDirectory
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

Public Function ExtractTokenParameter(ByVal TokenFilename As String, Token As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim dtS As String, sDT As Long, eDT As Long
50020  dtS = "<" & Token & ":"
50030  sDT = InStr(1, TokenFilename, dtS, vbTextCompare)
50040  If sDT > 0 Then
50050   eDT = InStr(sDT + Len(dtS), TokenFilename, ">", vbTextCompare)
50060   If eDT > sDT Then
50070    ExtractTokenParameter = Mid$(TokenFilename, sDT + Len(dtS), eDT - (sDT + Len(dtS)))
50080   End If
50090  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGlobal2", "ExtractTokenParameter")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function ReplaceToken(ByVal TokenFilename As String, Token As String, value As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim dtS As String, sDT As Long, eDT As Long, dateTimeFormat As String
50020  ReplaceToken = TokenFilename
50030  dtS = "<" & Token
50040  sDT = InStr(1, TokenFilename, dtS, vbTextCompare)
50050  If sDT > 0 Then
50060   eDT = InStr(sDT + Len(dtS), TokenFilename, ">", vbTextCompare)
50070   If eDT > sDT Then
50080    ReplaceToken = Mid$(TokenFilename, 1, sDT - 1) & value & Mid$(TokenFilename, eDT + 1)
50090   End If
50100  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGlobal2", "ReplaceToken")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetSubstFilename2(TokenFilename As String, Optional Preview As Boolean = True, Optional Temppath As String, Optional InfoSpoolFileName As String, Optional bReplaceForbiddenChars As Boolean = True) As String
 Dim DateTime As String, Author As String, ClientComputer As String, ClientUsername As String, UserName As String, _
  Computername As String, MyFiles As String, MyDesktop As String, filename As String, _
  Title As String, tStr As String, FilePath As String, tokenParameter As String
 Dim isf As clsInfoSpoolFile, currentDateTime As Date, oldFileName As String, counterStr As String

 If Len(TokenFilename) = 0 Then
  Exit Function
 End If

 Set isf = New clsInfoSpoolFile
 If LenB(InfoSpoolFileName) > 0 Then
  isf.ReadInfoFile InfoSpoolFileName
 End If

 Computername = GetComputerName

 MyFiles = GetMyFiles
 MyDesktop = GetDesktop

 filename = TokenFilename

 currentDateTime = Now
 tStr = "DateTime"
 Do
  tokenParameter = Trim$(ExtractTokenParameter(filename, tStr))
  If LenB(tokenParameter) = 0 Then
   tokenParameter = Options.StandardDateformat
  End If
  DateTime = GetDocDate("", tokenParameter, CStr(currentDateTime))
  oldFileName = filename
  filename = ReplaceToken(filename, tStr, DateTime)
 Loop Until oldFileName = filename

 filename = Replace(filename, "<Computername>", Replace(Computername, "\\", ""), , , vbTextCompare)

 filename = Replace(filename, "<Username>", GetUsername, , , vbTextCompare)

 filename = Replace(filename, "<MyFiles>", CompletePath(MyFiles), , , vbTextCompare)
 filename = Replace(filename, "<MyDesktop>", CompletePath(MyDesktop), , , vbTextCompare)

 If Options.Counter = 922337203685477@ Then
  Options.Counter = 0
 End If
 Options.Counter = Round(Options.Counter)

 tStr = "Counter"
 Do
  tokenParameter = Trim$(ExtractTokenParameter(filename, tStr))
  If LenB(tokenParameter) = 0 Then
   tokenParameter = String(15, "0")
  End If
  counterStr = Format$(Options.Counter + 1, tokenParameter)
  oldFileName = filename
  filename = ReplaceToken(filename, tStr, counterStr)
 Loop Until oldFileName = filename

 If LenB(Temppath) > 0 Then
   filename = Replace(filename, "<Temp>", CompletePath(Temppath), , , vbTextCompare)
  Else
   filename = Replace(filename, "<Temp>", CompletePath(GetTempPathApi), , , vbTextCompare)
 End If

 tStr = "Author"
 If Preview = True Then
   filename = Replace(filename, "<" & tStr & ">", "'Preview " & tStr & "'", , , vbTextCompare)
  Else
   filename = Replace(filename, "<" & tStr & ">", ReplaceForbiddenChars(isf.FirstUserName), , , vbTextCompare)
 End If

 tStr = "Title"
 If Preview Then
   filename = Replace(filename, "<" & tStr & ">", "'Preview " & tStr & "'", , , vbTextCompare)
  Else
   If LenB(isf.FirstDocumentTitle) > 0 Then
    filename = Replace(filename, "<" & tStr & ">", ReplaceForbiddenChars(isf.FirstDocumentTitle), , , vbTextCompare)
   End If
 End If
 tStr = "DocumentFilename"
 If Preview = True Then
   filename = Replace(filename, "<" & tStr & ">", "'Preview " & tStr & "'", , , vbTextCompare)
  Else
   SplitPath isf.FirstDocumentTitle, , , , FilePath
   filename = Replace(filename, "<" & tStr & ">", FilePath, , , vbTextCompare)
 End If
 tStr = "DocumentPath"
 If Preview = True Then
   filename = Replace(filename, "<" & tStr & ">", "'Preview " & tStr & "'", , , vbTextCompare)
  Else
   SplitPath isf.FirstDocumentTitle, , FilePath
   filename = Replace(filename, "<" & tStr & ">", FilePath, , , vbTextCompare)
 End If
 tStr = "JobID"
 If Preview Then
   filename = Replace(filename, "<" & tStr & ">", "'Preview " & tStr & "'", , , vbTextCompare)
  Else
   If LenB(isf.FirstJobID) > 0 Then
    filename = Replace(filename, "<" & tStr & ">", ReplaceForbiddenChars(isf.FirstJobID), , , vbTextCompare)
   End If
 End If
 tStr = "ClientComputer"
 If Preview Then
   filename = Replace(filename, "<" & tStr & ">", "'Preview " & tStr & "'", , , vbTextCompare)
  Else
   If LenB(isf.FirstClientComputer) > 0 Then
    filename = Replace(filename, "<" & tStr & ">", ReplaceForbiddenChars(Replace(isf.FirstClientComputer, "\\", "")), , , vbTextCompare)
   End If
 End If
 tStr = "PrinterName"
 If Preview Then
   filename = Replace(filename, "<" & tStr & ">", "'Preview " & tStr & "'", , , vbTextCompare)
  Else
   If LenB(isf.FirstPrinterName) > 0 Then
    filename = Replace(filename, "<" & tStr & ">", ReplaceForbiddenChars(isf.FirstPrinterName), , , vbTextCompare)
   End If
 End If
 tStr = "SessionID"
 If Preview Then
   filename = Replace(filename, "<" & tStr & ">", "'Preview " & tStr & "'", , , vbTextCompare)
  Else
   If LenB(isf.FirstSessionID) > 0 Then
    filename = Replace(filename, "<" & tStr & ">", ReplaceForbiddenChars(isf.FirstSessionID), , , vbTextCompare)
   End If
 End If

 If Options.RemoveSpaces = 1 Then
  filename = Trim$(filename)
 End If
 If bReplaceForbiddenChars Then
  filename = ReplaceForbiddenChars(filename)
 End If

 If Len(filename) >= 4 Then
  If InStr(2, filename, "\\") > 0 Then
   filename = Mid$(filename, 1, 1) & Replace$(Mid(filename, 2), "\\", "\")
  End If
 End If
 GetSubstFilename2 = filename
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

