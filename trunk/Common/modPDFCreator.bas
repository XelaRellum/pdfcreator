Attribute VB_Name = "modPDFCreator"
Option Explicit

Public Const Uninstall_GUID = "{0001B4FD-9EA3-4D90-A79E-FD14BA3AB01D}"
Public Const PDFCreator_GUID = "{A7332D94-E8FE-40B2-937F-8515FC0FF52F}"
Public Const TransTool_GUID = "{B7BCA0D2-7305-4318-BA7A-01B028D910EB}"
Public Const PDFSpooler_GUID = "{C387A397-047A-4354-AE89-F75B1B550257}"
Public Const UnInst_GUID = "{D95872D0-0DE7-4C01-859C-1BAE47FB1C6B}"
Public Const Paypal = "https://www.paypal.com/xclick/business=paypal01%40heindoerfer.com&item_name=PDFCreator&no_note=1&tax=0&currency_code=EUR"
Public Const Homepage = "http://www.pdfcreator.de.vu"
Public Const Sourceforge = "http://www.sourceforge.net/projects/pdfcreator"
Public Const UpdateURL = "http://www.pdfcreator.de.vu/update.txt"
Public Const PDFCreatorLogfile = "PDFCreator.log"
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
       SaveOpenFilterindex As Long

Public Function ReadLogfile() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim fn As Long, bufStr As String, tStr As String
50020
50030  If LenB(Dir(PDFCreatorLogfilePath & PDFCreatorLogfile)) = 0 Then
50040   Exit Function
50050  End If
50060
50070  fn = FreeFile
50080  Open PDFCreatorLogfilePath & PDFCreatorLogfile For Input As #fn
50090  bufStr = vbNullString
50100  While Not EOF(fn)
50110   If Len(bufStr) = 0 Then
50120     Line Input #fn, bufStr
50130    Else
50140     Line Input #fn, tStr
50150     If Len(Trim$(tStr)) > 0 Then
50160      bufStr = bufStr & vbCrLf & tStr
50170     End If
50180   End If
50190  Wend
50200  Close #fn
50210  ReadLogfile = bufStr
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDFCreator", "ReadLogfile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub ClearLogfile()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim fn As Long
50020  If LenB(Dir(PDFCreatorLogfilePath & PDFCreatorLogfile)) = 0 Then
50030   Exit Sub
50040  End If
50050  fn = FreeFile
50060  Open PDFCreatorLogfilePath & PDFCreatorLogfile For Output As #fn
50070  Close #fn
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDFCreator", "ClearLogfile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub WriteLogfile(Logtext As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim fn As Long, i As Long, bufStr As String, s() As String, tB As Boolean
50020
50030  bufStr = ReadLogfile
50040
50050  If LenB(Dir(PDFCreatorLogfilePath, vbDirectory)) = 0 Then
50060    tB = MakePath(PDFCreatorLogfilePath)
50070   Else
50080    tB = True
50090  End If
50100
50110  If tB = True Then
50120   fn = FreeFile
50130   Open PDFCreatorLogfilePath & PDFCreatorLogfile For Output As #fn
50140
50150   If Len(bufStr) > 0 Then
50160     s = Split(bufStr, vbCrLf)
50170     Print #fn, Now & ": " & Logtext
50180     For i = LBound(s) To UBound(s)
50190      Print #fn, Trim$(Replace(s(i), vbCrLf, ""))
50200     Next i
50210    Else
50220     Print #fn, Now & ": " & Logtext
50230   End If
50240   Close #fn
50250  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDFCreator", "WriteLogfile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub IfLoggingWriteLogfile(Logtext As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim fn As Long, i As Long, bufStr As String, s() As String, _
  tStr As String, tB As Boolean
50030
50040  If Options.Logging = 0 Then
50050   Exit Sub
50060  End If
50070
50080  If LenB(Dir(PDFCreatorLogfilePath, vbDirectory)) = 0 Then
50090    tB = MakePath(PDFCreatorLogfilePath)
50100   Else
50110    tB = True
50120  End If
50130
50140  If tB = True Then
50150
50160   bufStr = ReadLogfile
50170
50180   fn = FreeFile
50190   Open PDFCreatorLogfilePath & PDFCreatorLogfile For Output As #fn
50200
50210   If Len(bufStr) > 0 Then
50220     s = Split(bufStr, vbCrLf)
50230     Print #fn, Now & ": " & Logtext
50240     If Options.LogLines < UBound(s) - 1 Then
50250       For i = 2 To Options.LogLines
50260        tStr = s(i - 2)
50270        Print #fn, s(i - 2)
50280       Next i
50290      Else
50300       For i = LBound(s) To UBound(s)
50310        tStr = s(i)
50320        Print #fn, Trim$(Replace$(s(i), vbCrLf, ""))
50330       Next i
50340     End If
50350    Else
50360     Print #fn, Now & ": " & Logtext
50370   End If
50380   Close #fn
50390  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDFCreator", "IfLoggingWriteLogfile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub IfLoggingShowLogfile(Optional frmLog As Form, Optional frmMain As Form)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Options As tOptions
50020
50030  Options = ReadOptions
50040
50050  If Options.Logging = 0 Then
50060   Exit Sub
50070  End If
50080  If IsMissing(frmLog) = False And IsMissing(frmMain) = False Then
50090   frmLog.Show vbModal, frmMain
50100  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDFCreator", "IfLoggingShowLogfile")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub CreatePDFCreatorTempfolder()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Temppath As String
50020  If DirExists(Options.PrinterTemppath) = False Then
50030   Temppath = CompletePath(GetTempPath)
50040   MakePath Temppath
50050   Options.PrinterTemppath = Temppath
50060   SaveOptions Options
50070   Options = ReadOptions
50080  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDFCreator", "CreatePDFCreatorTempfolder")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function GetPDFCreatorTempfolder() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  GetPDFCreatorTempfolder = Options.PrinterTemppath
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDFCreator", "GetPDFCreatorTempfolder")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
 
Public Sub PsAssociate()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry
50020  Set reg = New clsRegistry
50030  With reg
50040   .hkey = HKEY_CLASSES_ROOT
50050   .CreateKey ".ps"
50060   .KeyRoot = ".ps"
50070   .SetRegistryValue "", "Postscript", REG_SZ
50080   .KeyRoot = ""
50090   .CreateKey "Postscript"
50100   .KeyRoot = "Postscript"
50110   .CreateKey "Shell"
50120   .KeyRoot = "Postscript\Shell"
50130   .CreateKey "Open"
50140   .KeyRoot = "Postscript\Shell\Open"
50150   .CreateKey "Command"
50160   .KeyRoot = "Postscript\Shell\Open\Command"
50170   .SetRegistryValue "", """" & App.Path & "\" & App.EXEName & ".exe"" -IF""%1""", REG_SZ
50180   .KeyRoot = "Postscript"
50190   .CreateKey "DefaultIcon"
50200   .KeyRoot = "PostScript\DefaultIcon"
50210   .SetRegistryValue "", App.Path & "\" & App.EXEName & ".exe,13", REG_SZ
50220  End With
50230  Set reg = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDFCreator", "PsAssociate")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function IsPsAssociate() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry
50020  Set reg = New clsRegistry
50030  IsPsAssociate = False
50040  reg.hkey = HKEY_CLASSES_ROOT
50050  reg.KeyRoot = ".ps"
50060  If reg.KeyExists = True Then
50070   If UCase$(reg.GetRegistryValue("")) = UCase$("Postscript") Then
50080    reg.KeyRoot = "Postscript"
50090    If reg.KeyExists = True Then
50100     reg.KeyRoot = "Postscript\DefaultIcon"
50110     If UCase$(reg.GetRegistryValue("")) = UCase$(App.Path & "\" & App.EXEName & ".exe,13") Then
50120      reg.KeyRoot = "Postscript\Shell\Open\Command"
50130      If UCase$(reg.GetRegistryValue("")) = UCase$("""" & App.Path & "\" & App.EXEName & ".exe"" -IF""%1""") Then
50140       IsPsAssociate = True
50150      End If
50160     End If
50170    End If
50180   End If
50190  End If
50200  Set reg = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDFCreator", "IsPsAssociate")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetProgramReleaseStr() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry, Release As String
50020  Set reg = New clsRegistry
50030  With reg
50040   .hkey = HKEY_LOCAL_MACHINE
50050   .KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50060   Release = .GetRegistryValue("ApplicationVersion")
50070   If Len(Trim$(.GetRegistryValue("BetaVersion"))) > 0 Then
50080    Release = Release & " Beta " & .GetRegistryValue("BetaVersion")
50090   End If
50100  End With
50110  Set reg = Nothing
50120  GetProgramReleaseStr = Release
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDFCreator", "GetProgramReleaseStr")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetProgramRelease(Optional WithBeta As Boolean = True) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry, Release As String
50020  Set reg = New clsRegistry
50030  With reg
50040   .hkey = HKEY_LOCAL_MACHINE
50050   .KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50060   Release = .GetRegistryValue("ApplicationVersion")
50070   If WithBeta = True Then
50080    If Len(Trim$(.GetRegistryValue("BetaVersion"))) > 0 Then
50090      Release = Release & "." & .GetRegistryValue("BetaVersion")
50100     Else
50110      Release = Release & ".0"
50120    End If
50130   End If
50140  End With
50150  Set reg = Nothing
50160  GetProgramRelease = Release
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDFCreator", "GetProgramRelease")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub ClearCache()
 On Error Resume Next
 If Dir(GetPDFCreatorTempfolder, vbDirectory) <> "" Then
  RmDir GetPDFCreatorTempfolder
 End If
End Sub

Public Function CheckPDFCreatorVersion(UpdateVersionsStr As String) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ' return 1 if there is a new version, otherwise 0
50020  Dim ProgRelease() As String, updRelease() As String, progReleaseStr As String, i As Byte
50030  CheckPDFCreatorVersion = 0
50040  progReleaseStr = GetProgramRelease
50050  If Len(progReleaseStr) = 0 Then
50060   Exit Function
50070  End If
50080  If Len(UpdateVersionsStr) = 0 Then
50090   Exit Function
50100  End If
50110  If InStr(1, progReleaseStr, ".") = 0 Then
50120   Exit Function
50130  End If
50140  If InStr(1, UpdateVersionsStr, ".") = 0 Then
50150   Exit Function
50160  End If
50170  ProgRelease = Split(progReleaseStr, ".")
50180  updRelease = Split(UpdateVersionsStr, ".")
50190  If UBound(ProgRelease) <> 2 Then
50200   Exit Function
50210  End If
50220  If UBound(updRelease) <> 2 Then
50230   Exit Function
50240  End If
50250  For i = 0 To 2
50260   If IsNumeric(updRelease(i)) = False Or IsNumeric(ProgRelease(0)) = False Then
50270    Exit Function
50280   End If
50290  Next i
50300  If CLng(updRelease(0)) > CLng(ProgRelease(0)) Then
50310   CheckPDFCreatorVersion = 1
50320   Exit Function
50330  End If
50340  If CLng(updRelease(0)) < CLng(ProgRelease(0)) Then
50350   Exit Function
50360  End If
50370  If CLng(updRelease(1)) > CLng(ProgRelease(1)) Then
50380   CheckPDFCreatorVersion = 1
50390   Exit Function
50400  End If
50410  If CLng(updRelease(1)) < CLng(ProgRelease(1)) Then
50420   Exit Function
50430  End If
50440  If CLng(updRelease(2)) > CLng(ProgRelease(2)) Then
50450   CheckPDFCreatorVersion = 1
50460   Exit Function
50470  End If
50480  If CLng(updRelease(2)) < CLng(ProgRelease(2)) Then
50490   Exit Function
50500  End If
50510  If (CLng(updRelease(3)) > CLng(ProgRelease(3)) And CLng(ProgRelease(3)) > 0) Or _
    (CLng(updRelease(3)) = 0 And CLng(ProgRelease(3)) > 0) Then
50530   CheckPDFCreatorVersion = 1
50540   Exit Function
50550  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDFCreator", "CheckPDFCreatorVersion")
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
50010  Dim reg As clsRegistry
50020  Set reg = New clsRegistry
50030  With reg
50040   .hkey = HKEY_LOCAL_MACHINE
50050   .KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50060   GetPDFCreatorApplicationPath = .GetRegistryValue("Inno Setup: App Path")
50070  End With
50080  Set reg = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPDFCreator", "GetPDFCreatorApplicationPath")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
