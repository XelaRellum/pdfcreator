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
 Dim fn As Long, bufStr As String, tStr As String

 If LenB(Dir(PDFCreatorLogfilePath & PDFCreatorLogfile)) = 0 Then
  Exit Function
 End If

 fn = FreeFile
 Open PDFCreatorLogfilePath & PDFCreatorLogfile For Input As #fn
 bufStr = vbNullString
 While Not EOF(fn)
  If Len(bufStr) = 0 Then
    Line Input #fn, bufStr
   Else
    Line Input #fn, tStr
    If Len(Trim$(tStr)) > 0 Then
     bufStr = bufStr & vbCrLf & tStr
    End If
  End If
 Wend
 Close #fn
 ReadLogfile = bufStr
End Function

Public Sub ClearLogfile()
 Dim fn As Long
 If LenB(Dir(PDFCreatorLogfilePath & PDFCreatorLogfile)) = 0 Then
  Exit Sub
 End If
 fn = FreeFile
 Open PDFCreatorLogfilePath & PDFCreatorLogfile For Output As #fn
 Close #fn
End Sub

Public Sub WriteLogfile(Logtext As String)
 Dim fn As Long, i As Long, bufStr As String, s() As String, tB As Boolean

 bufStr = ReadLogfile

 If LenB(Dir(PDFCreatorLogfilePath, vbDirectory)) = 0 Then
   tB = MakePath(PDFCreatorLogfilePath)
  Else
   tB = True
 End If

 If tB = True Then
  fn = FreeFile
  Open PDFCreatorLogfilePath & PDFCreatorLogfile For Output As #fn

  If Len(bufStr) > 0 Then
    s = Split(bufStr, vbCrLf)
    Print #fn, Now & ": " & Logtext
    For i = LBound(s) To UBound(s)
     Print #fn, Trim$(Replace(s(i), vbCrLf, ""))
    Next i
   Else
    Print #fn, Now & ": " & Logtext
  End If
  Close #fn
 End If
End Sub

Public Sub IfLoggingWriteLogfile(Logtext As String)
 Dim fn As Long, i As Long, bufStr As String, s() As String, _
  tStr As String, tB As Boolean

 If Options.Logging = 0 Then
  Exit Sub
 End If

 If LenB(Dir(PDFCreatorLogfilePath, vbDirectory)) = 0 Then
   tB = MakePath(PDFCreatorLogfilePath)
  Else
   tB = True
 End If

 If tB = True Then

  bufStr = ReadLogfile

  fn = FreeFile
  Open PDFCreatorLogfilePath & PDFCreatorLogfile For Output As #fn

  If Len(bufStr) > 0 Then
    s = Split(bufStr, vbCrLf)
    Print #fn, Now & ": " & Logtext
    If Options.LogLines < UBound(s) - 1 Then
      For i = 2 To Options.LogLines
       tStr = s(i - 2)
       Print #fn, s(i - 2)
      Next i
     Else
      For i = LBound(s) To UBound(s)
       tStr = s(i)
       Print #fn, Trim$(Replace$(s(i), vbCrLf, ""))
      Next i
    End If
   Else
    Print #fn, Now & ": " & Logtext
  End If
  Close #fn
 End If
End Sub

Public Sub IfLoggingShowLogfile(Optional frmLog As Form, Optional frmMain As Form)
 Dim Options As tOptions

 Options = ReadOptions

 If Options.Logging = 0 Then
  Exit Sub
 End If
 If IsMissing(frmLog) = False And IsMissing(frmMain) = False Then
  frmLog.Show vbModal, frmMain
 End If
End Sub

Public Sub CreatePDFCreatorTempfolder()
 Dim Temppath As String
 'Temppath = GetPDFCreatorTempfolder
 If Len(Trim$(Temppath)) = 0 Then
  Temppath = CompletePath(GetTempPath)
  Options.PrinterTemppath = Temppath
  SaveOptions Options
  Options = ReadOptions
 End If
 If Dir(Mid(Temppath, 1, Len(Temppath) - 1), vbDirectory) = "" Then
  MakePath Mid(Temppath, 1, Len(Temppath) - 1)
 End If
End Sub

Public Function GetPDFCreatorTempfolder() As String
 GetPDFCreatorTempfolder = Options.PrinterTemppath
End Function
 
Public Sub PsAssociate()
 Dim reg As clsRegistry
 Set reg = New clsRegistry
 With reg
  .hkey = HKEY_CLASSES_ROOT
  .CreateKey ".ps"
  .KeyRoot = ".ps"
  .SetRegistryValue "", "Postscript", REG_SZ
  .KeyRoot = ""
  .CreateKey "Postscript"
  .KeyRoot = "Postscript"
  .CreateKey "Shell"
  .KeyRoot = "Postscript\Shell"
  .CreateKey "Open"
  .KeyRoot = "Postscript\Shell\Open"
  .CreateKey "Command"
  .KeyRoot = "Postscript\Shell\Open\Command"
  .SetRegistryValue "", """" & App.Path & "\" & App.EXEName & ".exe"" -IF""%1""", REG_SZ
  .KeyRoot = "Postscript"
  .CreateKey "DefaultIcon"
  .KeyRoot = "PostScript\DefaultIcon"
  .SetRegistryValue "", App.Path & "\" & App.EXEName & ".exe,13", REG_SZ
 End With
 Set reg = Nothing
End Sub

Public Function IsPsAssociate() As Boolean
 Dim reg As clsRegistry
 Set reg = New clsRegistry
 IsPsAssociate = False
 reg.hkey = HKEY_CLASSES_ROOT
 reg.KeyRoot = ".ps"
 If reg.KeyExists = True Then
  If UCase$(reg.GetRegistryValue("")) = UCase$("Postscript") Then
   reg.KeyRoot = "Postscript"
   If reg.KeyExists = True Then
    reg.KeyRoot = "Postscript\DefaultIcon"
    If UCase$(reg.GetRegistryValue("")) = UCase$(App.Path & "\" & App.EXEName & ".exe,13") Then
     reg.KeyRoot = "Postscript\Shell\Open\Command"
     If UCase$(reg.GetRegistryValue("")) = UCase$("""" & App.Path & "\" & App.EXEName & ".exe"" -IF""%1""") Then
      IsPsAssociate = True
     End If
    End If
   End If
  End If
 End If
 Set reg = Nothing
End Function

Public Function GetProgramReleaseStr() As String
 Dim reg As clsRegistry, Release As String
 Set reg = New clsRegistry
 With reg
  .hkey = HKEY_LOCAL_MACHINE
  .KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
  Release = .GetRegistryValue("ApplicationVersion")
  If Len(Trim$(.GetRegistryValue("BetaVersion"))) > 0 Then
   Release = Release & " Beta " & .GetRegistryValue("BetaVersion")
  End If
 End With
 Set reg = Nothing
 GetProgramReleaseStr = Release
End Function

Public Function GetProgramRelease(Optional WithBeta As Boolean = True) As String
 Dim reg As clsRegistry, Release As String
 Set reg = New clsRegistry
 With reg
  .hkey = HKEY_LOCAL_MACHINE
  .KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
  Release = .GetRegistryValue("ApplicationVersion")
  If WithBeta = True Then
   If Len(Trim$(.GetRegistryValue("BetaVersion"))) > 0 Then
     Release = Release & "." & .GetRegistryValue("BetaVersion")
    Else
     Release = Release & ".0"
   End If
  End If
 End With
 Set reg = Nothing
 GetProgramRelease = Release
End Function

Public Sub ClearCache()
 On Error Resume Next
 If Dir(GetPDFCreatorTempfolder, vbDirectory) <> "" Then
  RmDir GetPDFCreatorTempfolder
 End If
End Sub

Public Function CheckPDFCreatorVersion(UpdateVersionsStr As String) As Long
 ' return 1 if there is a new version, otherwise 0
 Dim progRelease() As String, updRelease() As String
 progRelease = Split(GetProgramRelease, ".")
 updRelease = Split(UpdateVersionsStr, ".")
 CheckPDFCreatorVersion = 0
 If CLng(updRelease(0)) > CLng(progRelease(0)) Then
  CheckPDFCreatorVersion = 1
  Exit Function
 End If
 If CLng(updRelease(0)) < CLng(progRelease(0)) Then
  Exit Function
 End If
 If CLng(updRelease(1)) > CLng(progRelease(1)) Then
  CheckPDFCreatorVersion = 1
  Exit Function
 End If
 If CLng(updRelease(1)) < CLng(progRelease(1)) Then
  Exit Function
 End If
 If CLng(updRelease(2)) > CLng(progRelease(2)) Then
  CheckPDFCreatorVersion = 1
  Exit Function
 End If
 If CLng(updRelease(2)) < CLng(progRelease(2)) Then
  Exit Function
 End If
 If (CLng(updRelease(3)) > CLng(progRelease(3)) And CLng(progRelease(3)) > 0) Or _
    (CLng(updRelease(3)) = 0 And CLng(progRelease(3)) > 0) Then
  CheckPDFCreatorVersion = 1
  Exit Function
 End If
End Function

Public Function GetPDFCreatorApplicationPath() As String
 Dim reg As clsRegistry
 Set reg = New clsRegistry
 With reg
  .hkey = HKEY_LOCAL_MACHINE
  .KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
  GetPDFCreatorApplicationPath = .GetRegistryValue("Inno Setup: App Path")
 End With
 Set reg = Nothing
End Function
