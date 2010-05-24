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
50260  End Select
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
50030  If Not f Is Nothing Then
50040   f.Timer1.Enabled = False
50050  End If
50060  TestPSPage = GetTestpageFromRessource
50070  TestPSPage = Replace(TestPSPage, "[INFOTITLE]", LanguageStrings.OptionsTestpage, , 1, vbTextCompare)
50080  TestPSPage = Replace(TestPSPage, "[INFORELEASE]", App.title & " " & GetProgramReleaseStr, , 1, vbTextCompare)
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
50210  tStr = GetPDFCreatorApplicationPath & "PDFSpooler.exe"
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
50490  filename = GetTempFile(tStr, "~PS")
50500  Open filename For Output As fn
50510  Print #fn, TestPSPage
50520  Close #fn
50530  If Not f Is Nothing Then
50540   f.CheckPrintJobs
50550   f.Timer1.Enabled = True
50560  End If
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
Select Case ErrPtnr.OnError("modGeneral2", "CheckForUpdate")
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
