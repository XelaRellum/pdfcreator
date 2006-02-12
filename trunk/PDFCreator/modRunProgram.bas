Attribute VB_Name = "modRunProgram"
Option Explicit

Public Sub RunProgramAfterSaving(hwnd As Long, ByVal Docname As String, ByVal Parameters As String, Windowstyle As Long, PostscriptFile As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tStr As String, Program As String, WorkingFolder As String, drv As String
50020  Program = RemoveLeadingAndTrailingQuotes(Options.RunProgramAfterSavingProgramname)
50030  SplitPath Program, drv
50040  If LenB(drv) = 0 Then
50050   Program = ResolveRelativePath(Program, GetPDFCreatorApplicationPath)
50060  End If
50070  If FileExists(Docname) = True And FileExists(Program) = True Then
50080   tStr = "Run program after saving: Program:" & Program & _
   " Parameters:" & Parameters & Docname & "   WaitUntilReady:" & Options.RunProgramAfterSavingWaitUntilReady
50100   IfLoggingWriteLogfile tStr
50110   WriteToSpecialLogfile tStr
50120   SplitPath Program, , WorkingFolder
50130   Parameters = GetSubstFilename2(Parameters, False, , PostscriptFile)
50140   Parameters = Replace$(Parameters, "<OutputFilename>", Docname, , , vbTextCompare)
50150   If Options.RunProgramAfterSavingWaitUntilReady = 1 Then
50160     ShellAndWait hwnd, "open", Program, Parameters, CompletePath(WorkingFolder), Windowstyle, WCTermination
50170    Else
50180     ShellAndWait hwnd, "open", Program, Parameters, CompletePath(WorkingFolder), Windowstyle, WCNone
50190   End If
50200  End If
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

Public Sub RunProgramBeforeSaving(hwnd As Long, ByVal Docname As String, ByVal Parameters As String, Windowstyle As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tStr As String, Program As String, WorkingFolder As String, drv As String
50020  Program = RemoveLeadingAndTrailingQuotes(Options.RunProgramBeforeSavingProgramname)
50030  SplitPath Program, drv
50040  If LenB(drv) = 0 Then
50050   Program = ResolveRelativePath(Program, GetPDFCreatorApplicationPath)
50060  End If
50070  If FileExists(Docname) = True And FileExists(Program) = True Then
50080   tStr = "Run program before saving: Program:" & Options.RunProgramBeforeSavingProgramname & _
   " Parameters:" & Parameters & Docname
50100   IfLoggingWriteLogfile tStr
50110   WriteToSpecialLogfile tStr
50120   SplitPath Program, , WorkingFolder
50130   Parameters = GetSubstFilename2(Parameters, False, , Docname)
50140   Parameters = Replace$(Parameters, "<TempFilename>", Docname, , , vbTextCompare)
50150   ShellAndWait hwnd, "open", Program, Parameters, WorkingFolder, Windowstyle, WCTermination
50160  End If
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
