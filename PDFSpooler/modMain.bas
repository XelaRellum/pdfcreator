Attribute VB_Name = "modMain"
Option Explicit

Public mutex As clsMutex

Sub Main()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim fn As Long, stdio As clsStdIO, cinStr As String, Tempfile As String, _
  Spoolfile As String, TimerID As Long
50030
50040  TimerID = SetTimer(0, 0, 250, AddressOf TimerProc)
50050
50060  If UCase$(CommandSwitch("P", True)) = "PDFCREATORPRINTER" Then
50070   If CheckPath(GetMyAppData) = True Then
50080     If Dir(GetMyAppData & "PDFcreator", vbDirectory) = "" Then
50090      MakePath GetMyAppData & "PDFcreator"
50100     End If
50110     PDFCreatorINIFile = GetMyAppData & "PDFcreator\PDFCreator.ini"
50120    Else
50130     PDFCreatorINIFile = App.Path & "\PDFCreator.ini"
50140   End If
50150
50160   Options = ReadOptions
50170   CreatePDFCreatorTempfolder
50180   Set stdio = New clsStdIO
50190   cinStr = stdio.StdIn
50200   Set stdio = Nothing
50210   If FileLen(cinStr) > 0 Then
50220     Spoolfile = GetTempFile(Options.PrinterTemppath, "~PD")
50230     Kill Spoolfile
50240     Name cinStr As Spoolfile
50250     Shell App.Path & "\PDFCreator.exe -PPDFCREATORPRINTER", vbNormalFocus
50260    Else
50270     Kill cinStr
50280   End If
50290  End If
50300  TimerID = KillTimer(0, TimerID)
50310  Set mutex = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modMain", "Main")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
