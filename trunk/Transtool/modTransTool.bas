Attribute VB_Name = "modTransTool"
Option Explicit

Public TemplateInifile As String, TranslatedInifile As String, _
 RecentFilesCount As Long, LastSearchstrings As Collection, hLItems As Collection, _
 lsvColor As OLE_COLOR, lsvBold As Boolean

Public Sub Main()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  RecentFilesCount = 10
50020  TemplateInifile = App.Path & "\english.ini"
50030  Set hLItems = New Collection
50040  AnalyzeCommandlineParameters
50050  If CheckTemplate(TemplateInifile) = False Then
50060   TemplateInifile = ""
50070  End If
50080  Set LastSearchstrings = New Collection
50090  frmMain.Show
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modTransTool", "Main")
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
50010  Dim TemplateFile As String, Translatedfile As String, Path As String
50020  ' The program has commandswitches
50030  ' -Templatefile: The template file
50040  ' -Translatedfile: The translated file
50050
50060  If Len(VBA.Command$) > 0 Then
50070   TemplateFile = CommandSwitch("Templatefile=", True)
50080   SplitPath TemplateFile, , Path
50090   If LenB(Path) = 0 Then
50100    TemplateFile = CompletePath(App.Path) & TemplateFile
50110   End If
50120   If LenB(TemplateFile) > 0 Then
50130    If FileExists(TemplateFile) Then
50140      TemplateInifile = TemplateFile
50150     Else
50160      MsgBox "The template file doesn't exists!", vbExclamation
50170    End If
50180   End If
50190   Translatedfile = CommandSwitch("Translatedfile=", True)
50200   SplitPath Translatedfile, , Path
50210   If LenB(Path) = 0 Then
50220    Translatedfile = CompletePath(App.Path) & Translatedfile
50230   End If
50240   If LenB(Translatedfile) > 0 Then
50250    If FileExists(Translatedfile) = True Then
50260      If FileInUse(Translatedfile) = False Then
50270        TranslatedInifile = Translatedfile
50280       Else
50290        MsgBox "The translatedfile is in use!", vbExclamation
50300      End If
50310     Else
50320      MsgBox "The translatedfile doesn't exists!", vbExclamation
50330    End If
50340   End If
50350  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modTransTool", "AnalyzeCommandlineParameters")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function CheckTemplate(IniFile As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ini As clsINI, secs As Collection, keys As Collection, i As Long, j As Long, _
  Filename As String
50030  SplitPath IniFile, , , Filename
50040  CheckTemplate = False
50050  Set ini = New clsINI
50060  ini.Filename = IniFile
50070  If ini.CheckIniFile = False Then
50080   MsgBox "Template file '" & Filename & "' not found!", vbCritical
50090   Exit Function
50100  End If
50110  Set secs = ini.GetAllSectionsFromInifile(, True)
50120  If secs.Count = 0 Then
50130   MsgBox "Template file '" & Filename & "' has no sections!"
50140   Exit Function
50150  End If
50160  For i = 1 To secs.Count
50170   Set keys = ini.GetAllKeysFromSection(secs(i), , , True)
50180   If keys.Count = 0 Then
50190    MsgBox "In Template file '" & Filename & "' are no keys in section [" & secs(i) & "]!", vbCritical
50200    Exit Function
50210   End If
50220   For j = 1 To keys.Count
50230    If Len(Trim$(keys(j)(1))) = 0 Then
50240     MsgBox "In Template file '" & Filename & "' the key '" & keys(j)(0) & "' in section [" & secs(i) & "] has no value!", vbCritical
50250     Exit Function
50260    End If
50270   Next j
50280  Next i
50290  CheckTemplate = True
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modTransTool", "CheckTemplate")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub HighlightListitems(lsv As ListView, Items As Collection, _
 stb As StatusBar, xpPgb As XP_ProgressBar, _
 Optional PanelName As String = "", _
 Optional HighlightForeColor As OLE_COLOR = vbRed, _
 Optional Bold As Boolean = False, _
 Optional StatusText As String = "Marking ...")
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long
50020  If Items.Count > 0 Then
50030   If Len(PanelName) > 0 Then
50040    stb.Panels(PanelName).Text = StatusText: stb.Refresh
50050   End If
50060   With xpPgb
50070    .Min = 0
50080    .Max = lsv.ListItems.Count
50090   End With
50100   For i = 1 To Items.Count
50110    xpPgb.Value = i
50120    If Items(i)(1) = 0 Then
50130      lsv.ListItems(Items(i)(0)).ForeColor = HighlightForeColor
50140      lsv.ListItems(Items(i)(0)).Bold = Bold
50150     Else
50160      lsv.ListItems(Items(i)(0)).ListSubItems(Items(i)(1)).ForeColor = HighlightForeColor
50170      lsv.ListItems(Items(i)(0)).ListSubItems(Items(i)(1)).Bold = Bold
50180    End If
50190   Next i
50200   With xpPgb
50210    .Min = 0
50220    .Max = 0
50230   End With
50240   lsv.Refresh
50250   With lsv.ListItems(Items(1)(0))
50260    .Selected = True
50270    .EnsureVisible
50280    .Selected = False
50290   End With
50300   If Len(PanelName) > 0 Then
50310    stb.Panels(PanelName).Text = "": stb.Refresh
50320   End If
50330  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modTransTool", "HighlightListitems")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub


