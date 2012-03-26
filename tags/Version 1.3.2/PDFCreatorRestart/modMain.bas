Attribute VB_Name = "modMain"
Option Explicit

Public Sub Main()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim PDFCreatorPath As String
50020
50030  SleepTime = -1
50040  AnalyzeCommandlineParameters
50050
50060  PDFCreatorPath = CompletePath(GetPDFCreatorApplicationPath)
50070
50080  If SleepTime > 0 Then
50090   Sleep SleepTime
50100  End If
50110
50120  If StartPDFCreatorProgram And ProgramIsRunning(PDFCreator_GUID) = False Then
50130   Shell """" & PDFCreatorPath & "PDFCreator.exe""", vbNormalFocus
50140  End If
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

Private Sub AnalyzeCommandlineParameters()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ' Commandswitches
50020  ' -SL<Number>
50030  '      Wait <Number> milliseconds
50040  '      Example: -SL500
50050  '      Wait 500 milliseconds
50060  ' -ST
50070  '  -STTRUE
50080  '      Start pdfcreator, after using SL
50090  Dim cSwitch As String
50100  If Len(VBA.Command$) > 0 Then
50110   cSwitch = CommandSwitch("SL", True)
50120   If LenB(cSwitch) > 0 Then
50130    If IsNumeric(cSwitch) = True Then
50140     SleepTime = CLng(cSwitch)
50150    End If
50160   End If
50170   If UCase$(CommandSwitch("ST", False)) = "TRUE" Then
50180     StartPDFCreatorProgram = True
50190    Else
50200     StartPDFCreatorProgram = False
50210   End If
50220  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modMain", "AnalyzeCommandlineParameters")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
