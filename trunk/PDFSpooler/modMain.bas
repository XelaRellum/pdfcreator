Attribute VB_Name = "modMain"
Option Explicit

Sub Main()
 On Error Resume Next
 Dim fn As Long, stdio As clsStdIO, cinStr As String, Tempfile As String, _
  Spoolfile As String
 
 If UCase$(CommandSwitch("P", True)) = "PDFCREATORPRINTER" Then
  CreatePDFCreatorTempfolder
  Set stdio = New clsStdIO
  cinStr = stdio.StdIn
  If Len(cinStr) > 0 Then
   Tempfile = GetTempFile(GetTempPath & "PDFCreator\", "~PS")
   fn = FreeFile
   Open Tempfile For Output As #fn
   Print #fn, cinStr
   Close #fn
   Spoolfile = GetTempFile(GetTempPath & "PDFCreator\", "~PD")
   Kill Spoolfile
   Name Tempfile As Spoolfile
   Shell App.Path & "\PDFCreator.exe -PPDFCREATORPRINTER", vbNormalFocus
  End If
  Set stdio = Nothing
 End If
End Sub
