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
 Dim fn As Long, bufStr As String, tStr As String

 If Len(Dir(App.Path & "\" & PDFCreatorLogfile)) = 0 Then
  Exit Function
 End If

 fn = FreeFile
 Open App.Path & "\" & PDFCreatorLogfile For Input As #fn
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
 fn = FreeFile
 Open App.Path & "\" & PDFCreatorLogfile For Output As #fn
 Close #fn
End Sub

Public Sub WriteLogfile(Logtext As String)
 Dim fn As Long, i As Long, bufStr As String, s() As String

 bufStr = ReadLogfile

 fn = FreeFile
 Open App.Path & "\" & PDFCreatorLogfile For Output As #fn

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
End Sub

Public Sub IfLoggingWriteLogfile(Logtext As String)
 Dim fn As Long, i As Long, bufStr As String, s() As String, tStr As String

 If Options.Logging = 0 Then
  Exit Sub
 End If

 bufStr = ReadLogfile

 fn = FreeFile
 Open App.Path & "\" & PDFCreatorLogfile For Output As #fn

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
 Temppath = GetPDFCreatorTempfolder
 If Len(Trim$(Temppath)) = 0 Then
  Temppath = CompletePath(GetTempPath) & "PDFCreator"
  Options.PrinterTemppath = Temppath
  SaveOptions Options
  Options = ReadOptions
 End If
 If Dir(Temppath, vbDirectory) = "" Then
  MakePath Temppath
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

Public Function GetProgramRelease() As String
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
 GetProgramRelease = Release
End Function
