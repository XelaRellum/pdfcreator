Attribute VB_Name = "modMain"
Option Explicit

Public mutex As clsMutex

Sub Main()
 Dim fn As Long, stdio As clsStdIO, cinStr As String, Tempfile As String, _
  Spoolfile As String, TimerID As Long

 TimerID = SetTimer(0, 0, 250, AddressOf TimerProc)

 If UCase$(CommandSwitch("P", True)) = "PDFCREATORPRINTER" Then
  If CheckPath(GetMyAppData) = True Then
    If Dir(GetMyAppData & "PDFcreator", vbDirectory) = "" Then
     MakePath GetMyAppData & "PDFcreator"
    End If
    PDFCreatorINIFile = GetMyAppData & "PDFcreator\PDFCreator.ini"
   Else
    PDFCreatorINIFile = App.Path & "\PDFCreator.ini"
  End If

  Options = ReadOptions
  CreatePDFCreatorTempfolder
  Set stdio = New clsStdIO
  cinStr = stdio.StdIn
  Set stdio = Nothing
  If FileLen(cinStr) > 0 Then
    Spoolfile = GetTempFile(Options.PrinterTemppath, "~PD")
    Kill Spoolfile
    Name cinStr As Spoolfile
    Shell App.Path & "\PDFCreator.exe -PPDFCREATORPRINTER", vbNormalFocus
   Else
    Kill cinStr
  End If
 End If
 TimerID = KillTimer(0, TimerID)
 Set mutex = Nothing
End Sub
