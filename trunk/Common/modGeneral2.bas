Attribute VB_Name = "modGeneral2"
Option Explicit

Public Sub CombineFiles(ByVal filename As String, files As Collection, _
 Optional BufferSize As Long = 65536, Optional stb As StatusBar)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, j As Long, fnSource As Long, fnDest As Long, sBuffer As String, _
  aLen As Double, tLen As Double, bsize As Long, fpos As Long
50030
50040  bsize = BufferSize
50050  filename = Trim$(filename)
50060  If filename = vbNullString Or files.Count = 0 Or Right$(filename, 1) = "\" Then
50070   Exit Sub
50080  End If
50090  If files.Count = 1 Then
50100   Exit Sub
50110  End If
50120  fnDest = FreeFile
50130  aLen = 0: tLen = 0
50140  For i = 1 To files.Count
50150   aLen = aLen + FileLen(files.Item(i))
50160  Next i
50170  Open filename For Binary As #fnDest
50180  For i = 1 To files.Count
50190   DoEvents
50200   If FileExists(files.Item(i)) = False Then
50210    MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & vbCrLf & files.Item(i)
50220   End If
50230   If FileLen(files.Item(i)) > 0 Then
50240    fnSource = FreeFile
50250    Open files.Item(i) For Binary Access Read As #fnSource
50260    If bsize > LOF(fnSource) Then
50270     bsize = LOF(fnSource)
50280    End If
50290    fpos = 1
50300    For j = 1 To LOF(fnSource) \ bsize
50310     fpos = (j - 1) * bsize + 1
50320     Seek #fnSource, fpos
50330     sBuffer = Input(bsize, fnSource)
50340     Put #fnDest, , sBuffer
50350     tLen = tLen + bsize
50360     If Not stb Is Nothing Then
50370      stb.Panels("Percent").Text = Format(CDbl(tLen) / CDbl(aLen), "0.0%")
50380     End If
50390     DoEvents
50400    Next j
50410    If LOF(fnSource) > (j - 1) * bsize Then
50420     fpos = (j - 1) * bsize + 1
50430     Seek #fnSource, fpos
50440     sBuffer = Input(LOF(fnSource) - (j - 1) * bsize, fnSource)
50450     Put #fnDest, , sBuffer
50460     tLen = tLen + (LOF(fnSource) - (j - 1) * bsize)
50470    End If
50480    Close #fnSource
50490   End If
50500   DoEvents
50510  Next i
50520  For i = 1 To files.Count
50530   KillFile files.Item(i)
50540   KillInfoSpoolfile files.Item(i)
50550   DoEvents
50560  Next i
50570  Close #fnDest
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral2", "CombineFiles")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub


Public Sub CombineFilesOld(ByVal filename As String, files As Collection, stb As StatusBar)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, fnSource As Long, fnDest As Long, sBuffer As String, _
  aLen As Double, tLen As Double
50030
50040  filename = Trim$(filename)
50050  If filename = vbNullString Or files.Count = 0 Or Right$(filename, 1) = "\" Then
50060   Exit Sub
50070  End If
50080  If FileExists(filename) = True Then
50090   Exit Sub
50100  End If
50110  If files.Count = 1 Then
50120   Exit Sub
50130  End If
50140  fnDest = FreeFile
50150  aLen = 0
50160  For i = 1 To files.Count
50170   aLen = aLen + FileLen(files.Item(i))
50180  Next i
50190  Open filename For Binary As #fnDest
50200  For i = 1 To files.Count
50210   DoEvents
50220   If FileLen(files.Item(i)) > 0 Then
50230    fnSource = FreeFile
50240    Open files.Item(i) For Binary Access Read As #fnSource
50250    sBuffer = String(LOF(fnSource), Chr$(0))
50260    Get #fnSource, , sBuffer
50270    Put #fnDest, , sBuffer
50280    Close #fnSource
50290   End If
50300   tLen = tLen + FileLen(files.Item(i))
50310   stb.Panels("Percent").Text = Format$(tLen / aLen, "0.0%")
50320   KillFile files.Item(i)
50330   DoEvents
50340  Next i
50350  Close #fnDest
50360  stb.Panels("Percent").Text = vbNullString
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral2", "CombineFilesOld")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SetFrame(ByRef ctl As Control, Optional TempOptionsDesign = -1)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tOD As Long
50020  If TempOptionsDesign > -1 Then
50030    tOD = TempOptionsDesign
50040   Else
50050    tOD = Options.OptionsDesign
50060  End If
50070  ctl.Font.Size = 10
50080  ctl.TextShaddowColor = &HC00000
50090  If ComputerScreenResolution <= 8 Or tOD = 1 Then
50100    ctl.UseGradient = False: ctl.Caption3D = [Flat Caption]
50110    If UCase$(ctl.Name) = "DMFRADESCRIPTION" Then
50120      ctl.BarColorFrom = vbRed
50130     Else
50140      ctl.BarColorFrom = vbBlue
50150    End If
50160   Else
50170    ctl.UseGradient = True: ctl.Caption3D = [Raised Caption]
50180    If UCase$(ctl.Name) = "DMFRADESCRIPTION" Then
50190      ctl.BarColorFrom = &HB0BED  '&HA0A0FF ' &HFB&
50200      ctl.BarColorTo = &H20564  '&HE0& ' &HDDFF&
50210     Else
50220      ctl.BarColorFrom = &HED0B0B '&HFFA0A0
50230      ctl.BarColorTo = &H640502 '&H600000
50240    End If
50250  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral2", "SetFrame")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function GetSaveAutosaveFormatExtension(Index As Long) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50011  Select Case Index
        Case -1, 0, 1, 2
50030    GetSaveAutosaveFormatExtension = ".pdf"
50040   Case 3
50050    GetSaveAutosaveFormatExtension = ".png"
50060   Case 4
50070    GetSaveAutosaveFormatExtension = ".jpg"
50080   Case 5
50090    GetSaveAutosaveFormatExtension = ".bmp"
50100   Case 6
50110    GetSaveAutosaveFormatExtension = ".pcx"
50120   Case 7
50130    GetSaveAutosaveFormatExtension = ".tif"
50140   Case 8
50150    GetSaveAutosaveFormatExtension = ".ps"
50160   Case 9
50170    GetSaveAutosaveFormatExtension = ".eps"
50180   Case 10
50190    GetSaveAutosaveFormatExtension = ".txt"
50200   Case 11
50210    GetSaveAutosaveFormatExtension = ".psd"
50220   Case 12
50230    GetSaveAutosaveFormatExtension = ".pcl"
50240   Case 13
50250    GetSaveAutosaveFormatExtension = ".raw"
50260   Case 14
50270    GetSaveAutosaveFormatExtension = ".svg"
50280  End Select
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral2", "GetSaveAutosaveFormatExtension")
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
50010  Dim TestPSPage As String, fn As Long, filename As String, tStr As String, _
  c As Collection
50030  Dim Path As String, File As String, ini As clsINI, Title As String, TestPSFileName As String, strGUID As String
50040  Dim reg As clsRegistry
50050
50060  If Not f Is Nothing Then
50070   f.Timer1.Enabled = False
50080  End If
50090  TestPSPage = GetTestpageFromRessource
50100  TestPSPage = Replace(TestPSPage, "[INFOTITLE]", LanguageStrings.OptionsTestpage, , 1, vbTextCompare)
50110  TestPSPage = Replace(TestPSPage, "[INFORELEASE]", App.Title & " " & GetProgramReleaseStr, , 1, vbTextCompare)
50120  TestPSPage = Replace(TestPSPage, "[INFODATE]", Now, , 1, vbTextCompare)
50130  TestPSPage = Replace(TestPSPage, "[INFOAUTHORS]", "Philip Chinery, Frank Heind\224rfer", , 1, vbTextCompare)
50140  TestPSPage = Replace(TestPSPage, "[INFOHOMEPAGE]", Homepage, , 1, vbTextCompare)
50150  tStr = PDFCreatorApplicationPath & "PDFCreator.exe"
50160  If FileExists(tStr) = True Then
50170    Set c = GetFileVersion(tStr)
50180    tStr = "Version: " & c(2) & "; Size: " & Format(FileLen(tStr), "###,###,###,### Bytes")
50190   Else
50200    tStr = ""
50210  End If
50220  TestPSPage = Replace(TestPSPage, "[INFOPDFCREATOR]", tStr, , 1, vbTextCompare)
50230
50240  tStr = ""
50250 ' Set reg = New clsRegistry
50260 ' With reg
50270 '  .hkey = HKEY_LOCAL_MACHINE
50280 '  .KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50290 '  If .KeyExists = True Then
50300 '   If IsWin64 Then
50310 '     tStr = Trim$(.GetRegistryValue("pdfcmonVersion64"))
50320 '    Else
50330 '     tStr = Trim$(.GetRegistryValue("pdfcmonVersion32"))
50340 '   End If
50350 '  End If
50360 ' End With
50370 ' Set reg = Nothing
50380  Wow64FsRedirection True
50390  tStr = CompletePath(GetSystemDirectory) & "pdfcmon.dll"
50400  Wow64FsRedirection False
50410
50420  If FileExists(tStr) = True Then
50430    Set c = GetFileVersion(tStr)
50440    tStr = "Version: " & c(2) & "; Size: " & Format(FileLen(tStr), "###,###,###,### Bytes")
50450   Else
50460    tStr = ""
50470  End If
50480
50490
50500
50510  TestPSPage = Replace(TestPSPage, "[INFOPDFMONITOR]", tStr, , 1, vbTextCompare)
50520
50530  tStr = PDFCreatorApplicationPath & "Languages\Transtool.exe"
50540  If FileExists(tStr) = True Then
50550    Set c = GetFileVersion(tStr)
50560    tStr = "Version: " & c(2) & "; Size: " & Format(FileLen(tStr), "###,###,###,### Bytes")
50570   Else
50580    tStr = ""
50590  End If
50600  TestPSPage = Replace(TestPSPage, "[INFOTRANSTOOL]", tStr, , 1, vbTextCompare)
50610
50620  TestPSPage = Replace(TestPSPage, "[INFOCOMPUTER]", GetComputerName, , 1, vbTextCompare)
50630  tStr = GetWinVersionStr
50640  TestPSPage = Replace(TestPSPage, "[INFOWINDOWS]", _
  Mid(tStr, 1, IIf(InStr(1, tStr, "[") > 0, InStr(1, tStr, "[") - 1, Len(tStr))), 1, vbTextCompare)
50660
50670  If IsWin64 Then
50680    TestPSPage = Replace(TestPSPage, "[INFO64BIT]", "true", , 1, vbTextCompare)
50690   Else
50700    TestPSPage = Replace(TestPSPage, "[INFO64BIT]", "false", , 1, vbTextCompare)
50710  End If
50720
50730  tStr = GetPDFCreatorSpoolDirectory
50740  strGUID = GetGUID
50750  File = CompletePath(tStr) & strGUID
50760  filename = File & ".inf"
50770  fn = FreeFile
50780  Open filename For Output As fn
50790  Print #fn, TestPSPage
50800  Close #fn
50810
50820  TestPSFileName = File & ".ps"
50830  Name filename As TestPSFileName
50840
50850  Set ini = New clsINI
50860  ini.filename = CompletePath(Path) & File & ".inf"
50870  ini.Section = "1"
50880  Title = GetPSTitleFromPSString(TestPSPage)
50890  ini.SaveKey GetComputerName, "ClientComputer"
50900  ini.SaveKey Title, "DocumentTitle"
50910  ini.SaveKey "0", "JobId"
50920  ini.SaveKey "", "Printername"
50930  ini.SaveKey "", "SessionID"
50940  ini.SaveKey TestPSFileName, "SpoolFilename"
50950  ini.SaveKey GetUsername, "UserName"
50960  ini.SaveKey Environ$("SESSIONNAME"), "WinStation"
50970
50980  If Not f Is Nothing Then
50990   f.CheckPrintJobs
51000   f.Timer1.Enabled = True
51010  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral2", "PrintTestpage")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function IsPrintable(filename As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Ext As String, reg As clsRegistry, tStr As String
50020  IsPrintable = False
50030  SplitPath filename, , , , , Ext
50040  Set reg = New clsRegistry
50050  reg.hkey = HKEY_CLASSES_ROOT
50060  reg.KeyRoot = "." & Ext
50070  If reg.KeyExists Then
50080   tStr = reg.GetRegistryValue("")
50090   If LenB(tStr) > 0 Then
50100    reg.KeyRoot = tStr
50110    reg.SubKey = "shell\print\command"
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
Select Case ErrPtnr.OnError("modGeneral2", "IsPrintable")
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
Select Case ErrPtnr.OnError("modGeneral2", "GetTestpageFromRessource")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub CheckForUpdateAutomatically(ShowMessageNoNewUpdates As Boolean, ShowErrorMessage As Boolean, TimeOutInMs As Long)
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
50320  diff = DateDiff("d", lucDate, curDate)
50330  If (Options.UpdateInterval) = 1 Then ' Once a day
50340    If diff >= 1 Then
50350     Set upd = New clsUpdate
50360     upd.CheckForUpdates ShowMessageNoNewUpdates, ShowErrorMessage, TimeOutInMs
50370     SetLastUpdateCeck curDate
50380    End If
50390   ElseIf (Options.UpdateInterval) = 2 Then ' Once a week
50400    If diff >= 7 Then
50410     Set upd = New clsUpdate
50420     upd.CheckForUpdates ShowMessageNoNewUpdates, ShowErrorMessage, TimeOutInMs
50430     SetLastUpdateCeck curDate
50440    End If
50450   Else ' Once a month
50460    If diff >= 30 Then
50470     Set upd = New clsUpdate
50480     upd.CheckForUpdates ShowMessageNoNewUpdates, ShowErrorMessage, TimeOutInMs
50490     SetLastUpdateCeck curDate
50500    End If
50510  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral2", "CheckForUpdateAutomatically")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub SetLastUpdateCeck(d As Date)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Options.LastUpdateCheck = Format$(d, "YYYYMMDD")
50020  SaveOption Options, "LastUpdateCheck"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral2", "SetLastUpdateCeck")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function IsPDFArchitectInstalled() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If FileExists(PDFCreatorApplicationPath & "PDFArchitect\PDFArchitect.exe") Then
50020    IsPDFArchitectInstalled = True
50030   Else
50040    IsPDFArchitectInstalled = False
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral2", "IsPDFArchitectInstalled")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
