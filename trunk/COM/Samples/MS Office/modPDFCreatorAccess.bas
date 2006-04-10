Attribute VB_Name = "modPDFCreatorAccess"
Option Compare Database
Option Explicit

' Add a reference to PDFCreator

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const maxTime = 10    ' in seconds
Private Const sleepTime = 250 ' in milliseconds

Public Function Start()
 PrintRep "Report1"
End Function

Public Sub PrintRep(RepName As String)
 Dim PDFCreator1 As PDFCreator.clsPDFCreator, DefaultPrinter As String, c As Long, _
  OutputFilename As String
 Set PDFCreator1 = New clsPDFCreator
 With PDFCreator1
  .cStart "/NoProcessingAtStartup"
  .cOption("UseAutosave") = 1
  .cOption("UseAutosaveDirectory") = 1
  .cOption("AutosaveDirectory") = "C:\"
  .cOption("AutosaveFilename") = RepName
  .cOption("AutosaveFormat") = 0                            ' 0 = PDF
  DefaultPrinter = .cDefaultPrinter
  .cDefaultPrinter = "PDFCreator"
  .cClearCache
  DoCmd.OpenReport RepName, acViewNormal
  .cPrinterStop = False
 End With

 c = 0

 Do While (PDFCreator1.cOutputFilename = "") And (c < (maxTime * 1000 / sleepTime))
  c = c + 1
  Sleep 200
 Loop

 OutputFilename = PDFCreator1.cOutputFilename

 With PDFCreator1
  .cDefaultPrinter = DefaultPrinter
  Sleep 200
  .cClose
 End With

 Sleep 2000 ' Wait until PDFCreator is removed from memory

 If OutputFilename = "" Then
  MsgBox "Creating pdf file." & vbCrLf & vbCrLf & _
   "An error is occured: Time is up!", vbExclamation + vbSystemModal
 End If

End Sub
