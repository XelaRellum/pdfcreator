Attribute VB_Name = "modRunProgram"
Option Explicit


Public Sub RunProgramAfterSaving(hwnd As Long, Docname As String, Parameters As String, Windowstyle As Long, Optional Spoolfile As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tStr As String, Program As String, WorkingFolder As String, InfoParam As String
50020  If FileExists(Docname) = True And FileExists(RemoveLeadingAndTrailingQuotes(Options.RunProgramAfterSavingProgramname)) = True Then
50030   tStr = "Run program after saving: Program:" & Options.RunProgramAfterSavingProgramname & _
   " Parameters:" & Parameters & Docname & "   WaitUntilReady:" & Options.RunProgramAfterSavingWaitUntilReady
50050   IfLoggingWriteLogfile tStr
50060   WriteToSpecialLogfile tStr
50070   Program = RemoveLeadingAndTrailingQuotes(Options.RunProgramAfterSavingProgramname)
50080   SplitPath Program, , WorkingFolder
50090   tStr = Trim$(GetDocUsername(Spoolfile, False))
50100   If LenB(Trim$(tStr)) > 0 Then
50110    InfoParam = """" & Trim$(tStr) & """"
50120   End If
50130   tStr = Trim$(GetClientMachine(Spoolfile, False))
50140   If LenB(Trim$(tStr)) > 0 Then
50150    InfoParam = InfoParam & " """ & Trim$(tStr) & """"
50160   End If
50170   If LenB(Trim$(InfoParam)) > 0 Then
50180    InfoParam = " " & InfoParam
50190   End If
50200   If Options.RunProgramAfterSavingWaitUntilReady = 1 Then
50210     ShellAndWait hwnd, "open", Program, Parameters & """" & Docname & """" & InfoParam, CompletePath(WorkingFolder), Windowstyle, WCTermination
50220    Else
50230     ShellAndWait hwnd, "open", Program, Parameters & """" & Docname & """" & InfoParam, CompletePath(WorkingFolder), Windowstyle, WCNone
50240   End If
50250  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modRunProgram", "RunProgramAfterSaving")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub RunProgramBeforeSaving(hwnd As Long, Docname As String, Parameters As String, Windowstyle As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tStr As String, Program As String, WorkingFolder As String, InfoParam As String
50020  If FileExists(Docname) = True And FileExists(RemoveLeadingAndTrailingQuotes(Options.RunProgramBeforeSavingProgramname)) = True Then
50030   tStr = "Run program before saving: Program:" & Options.RunProgramBeforeSavingProgramname & _
   " Parameters:" & Parameters & Docname
50050   IfLoggingWriteLogfile tStr
50060   WriteToSpecialLogfile tStr
50070   Program = RemoveLeadingAndTrailingQuotes(Options.RunProgramBeforeSavingProgramname)
50080   SplitPath Program, , WorkingFolder
50090   tStr = Trim$(GetDocUsername(Docname, False))
50100   If LenB(Trim$(tStr)) > 0 Then
50110    InfoParam = """" & Trim$(tStr) & """"
50120   End If
50130   tStr = Trim$(GetClientMachine(Docname, False))
50140   If LenB(Trim$(tStr)) > 0 Then
50150    InfoParam = InfoParam & " """ & Trim$(tStr) & """"
50160   End If
50170   If LenB(Trim$(InfoParam)) > 0 Then
50180    InfoParam = " " & InfoParam
50190   End If
50200   ShellAndWait hwnd, "open", Program, Parameters & """" & Docname & """" & InfoParam, WorkingFolder, Windowstyle, WCTermination
50210  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modRunProgram", "RunProgramBeforeSaving")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub


