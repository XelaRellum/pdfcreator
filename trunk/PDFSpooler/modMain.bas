Attribute VB_Name = "modMain"
Option Explicit

Public mutex As clsMutex

Sub Main()
 Dim fn As Long, stdio As clsStdIO, cinStr As String, Tempfile As String, _
  Spoolfile As String, TimerID As Long, mSeconds As Long, cSwitch As String, _
  stSwitch As String, SpooltimeSeconds As Double, PDFCreatorPath As String
 ' Commandswitches
 ' -PPDFCREATORPRINTER
 '  Start PDFCreator if PDFSpooler found printer dates
 '  Now know PDFCreator that PDFSpooler call PDFCreator
 ' -SL<Number>
 '  Wait <Number> milliseconds
 '  Example: -SL500
 '   Wait 500 milliseconds
 ' -ST
 '  -STTRUE
 '   Start pdfcreator, after using SL

 ReadVersionInfo

 If Environ$("Redmon_User") <> "" And Win95 = False And Win98 = False And WinME = False Then
  LoadprofileUser Environ$("Redmon_User")
 End If


 PDFCreatorPath = CompletePath(GetPDFCreatorApplicationPath)

 PDFCreatorLogfilePath = CompletePath(GetTempPath)
 cSwitch = CommandSwitch("SL", True)
 If LenB(cSwitch) > 0 Then
  If IsNumeric(cSwitch) = True Then
   mSeconds = CLng(cSwitch)
   Sleep mSeconds
  End If
 End If

 stSwitch = CommandSwitch("ST", True)
 If LenB(stSwitch) > 0 Then
  If UCase$(stSwitch) = "TRUE" Then
   Shell """" & PDFCreatorPath & "PDFCreator.exe""", vbNormalFocus
   IfLoggingWriteLogfile "PDFSpooler Program End"
   Set mutex = Nothing
   End
  End If
 End If

 If UCase$(CommandSwitch("P", True)) = "PDFCREATORPRINTER" Then
  If CheckPath(GetMyAppData) = True Then
    If Dir(GetMyAppData & "PDFcreator", vbDirectory) = "" Then
     MakePath GetMyAppData & "PDFcreator"
    End If
    PDFCreatorINIFile = GetMyAppData & "PDFcreator\PDFCreator.ini"
   Else
    PDFCreatorINIFile = PDFCreatorPath & "PDFCreator.ini"
  End If

  Options = ReadOptions

  IfLoggingWriteLogfile "PDFSpooler Program Start"
  CreatePDFCreatorTempfolder
  Set stdio = New clsStdIO
  cinStr = stdio.StdIn(SpooltimeSeconds)
  Set stdio = Nothing
  If FileLen(cinStr) > 0 Then
    If Win95 = True Or Win98 = True Or WinME = True Then
      Shell """" & PDFCreatorPath & "PDFCreator.exe"" -PPDFCREATORPRINTER -IF""" & cinStr & """"
     Else
      RunAsUser """" & PDFCreatorPath & "PDFCreator.exe"" -PPDFCREATORPRINTER -IF""" & cinStr & """", App.Path, Environ("Redmon_User")
    End If
   Else
    Kill cinStr
  End If
 End If
 Set mutex = Nothing
 IfLoggingWriteLogfile "PDFSpooler: Spoolfile: " & Spoolfile
 IfLoggingWriteLogfile "PDFSpooler: Spoolfiletime (seconds): " & SpooltimeSeconds
 IfLoggingWriteLogfile "PDFSpooler Program End"
End Sub
