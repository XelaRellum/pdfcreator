Attribute VB_Name = "modPDFCreator"
Option Explicit

Public Const Uninstall_GUID = "{0001B4FD-9EA3-4D90-A79E-FD14BA3AB01D}"
Public Const PDFCreator_GUID = "{A7332D94-E8FE-40B2-937F-8515FC0FF52F}"
Public Const TransTool_GUID = "{B7BCA0D2-7305-4318-BA7A-01B028D910EB}"
Public Const PDFSpooler_GUID = "{C387A397-047A-4354-AE89-F75B1B550257}"
Public Const UnInst_GUID = "{D95872D0-0DE7-4C01-859C-1BAE47FB1C6B}"

Public Const PDFCreatorLogfile = "PDFCreator.log"

Public PDFCreatorINIFile As String, PrinterStop As Boolean, Printing As Boolean

Public Function ReadLogfile() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim fn As Long, bufStr As String, tStr As String
50020
50030  If Len(Dir(App.Path & "\" & PDFCreatorLogfile)) = 0 Then
50040   Exit Function
50050  End If
50060
50070  fn = FreeFile
50080  Open App.Path & "\" & PDFCreatorLogfile For Input As #fn
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
50020  fn = FreeFile
50030  Open App.Path & "\" & PDFCreatorLogfile For Output As #fn
50040  Close #fn
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
50010  Dim fn As Long, i As Long, bufStr As String, s() As String
50020
50030  bufStr = ReadLogfile
50040
50050  fn = FreeFile
50060  Open App.Path & "\" & PDFCreatorLogfile For Output As #fn
50070
50080  If Len(bufStr) > 0 Then
50090    s = Split(bufStr, vbCrLf)
50100    Print #fn, Now & ": " & Logtext
50110    For i = LBound(s) To UBound(s)
50120     Print #fn, Trim$(Replace(s(i), vbCrLf, ""))
50130    Next i
50140   Else
50150    Print #fn, Now & ": " & Logtext
50160  End If
50170  Close #fn
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
50010  Dim fn As Long, i As Long, bufStr As String, s() As String, tStr As String
50020
50030  If Options.Logging = 0 Then
50040   Exit Sub
50050  End If
50060
50070  bufStr = ReadLogfile
50080
50090  fn = FreeFile
50100  Open App.Path & "\" & PDFCreatorLogfile For Output As #fn
50110
50120  If Len(bufStr) > 0 Then
50130    s = Split(bufStr, vbCrLf)
50140    Print #fn, Now & ": " & Logtext
50150    If Options.LogLines < UBound(s) - 1 Then
50160      For i = 2 To Options.LogLines
50170       tStr = s(i - 2)
50180       Print #fn, s(i - 2)
50190      Next i
50200     Else
50210      For i = LBound(s) To UBound(s)
50220       tStr = s(i)
50230       Print #fn, Trim$(Replace$(s(i), vbCrLf, ""))
50240      Next i
50250    End If
50260   Else
50270    Print #fn, Now & ": " & Logtext
50280  End If
50290  Close #fn
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
50020  Temppath = GetPDFCreatorTempfolder
50030  If Len(Trim$(Temppath)) = 0 Then
50040   Temppath = CompletePath(GetTempPath) & "PDFCreator"
50050   Options.PrinterTemppath = Temppath
50060   SaveOptions Options
50070   Options = ReadOptions
50080  End If
50090  If Dir(Temppath, vbDirectory) = "" Then
50100   MakePath Temppath
50110  End If
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

Public Function GetProgramRelease() As String
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
50110
50120  Set reg = Nothing
50130  GetProgramRelease = Release
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

