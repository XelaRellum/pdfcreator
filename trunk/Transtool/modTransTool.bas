Attribute VB_Name = "modTransTool"
Option Explicit

Public TemplateInifile As String, TranslatedInifile As String, _
 RecentFilesCount As Long, LastSearchstrings As Collection, hLItems As Collection, _
 lsvColor As OLE_COLOR, lsvBold As Boolean

Public Sub Main()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040
50050  ' Reduce the working size of used memory
50060  Call SetProcessWorkingSetSize(GetCurrentProcess(), -1, -1)
50070  RecentFilesCount = 10
50080
50090  TemplateInifile = App.Path & "\english.ini"
50100  Set hLItems = New Collection
50110  AnalyzeCommandlineParameters
50120  If CheckTemplate(TemplateInifile) = False Then
50130   TemplateInifile = ""
50140  End If
50150  Set LastSearchstrings = New Collection
50160  If IsWin9xMe = False And IsWinNT4 = False And IsWin2000 = False Then
50170   InitCommonControls
50180  End If
50190  frmMain.Show
50200 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50210 Exit Sub
ErrPtnr_OnError:
50231 Select Case ErrPtnr.OnError("modTransTool", "Main")
      Case 0: Resume
50250 Case 1: Resume Next
50260 Case 2: Exit Sub
50270 Case 3: End
50280 End Select
50290 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub AnalyzeCommandlineParameters()
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim TemplateFile As String, Translatedfile As String, Path As String
50050  ' The program has commandswitches
50060  ' -Templatefile: The template file
50070  ' -Translatedfile: The translated file
50080
50090  If Len(VBA.Command$) > 0 Then
50100   TemplateFile = CommandSwitch("Templatefile=", True)
50110   SplitPath TemplateFile, , Path
50120   If LenB(Path) = 0 Then
50130    TemplateFile = CompletePath(App.Path) & TemplateFile
50140   End If
50150   If LenB(TemplateFile) > 0 Then
50160    If FileExists(TemplateFile) Then
50170      TemplateInifile = TemplateFile
50180     Else
50190      MsgBox "The template file doesn't exists!", vbExclamation
50200    End If
50210   End If
50220   Translatedfile = CommandSwitch("Translatedfile=", True)
50230   SplitPath Translatedfile, , Path
50240   If LenB(Path) = 0 Then
50250    Translatedfile = CompletePath(App.Path) & Translatedfile
50260   End If
50270   If LenB(Translatedfile) > 0 Then
50280    If FileExists(Translatedfile) = True Then
50290      If FileInUse(Translatedfile) = False Then
50300        TranslatedInifile = Translatedfile
50310       Else
50320        MsgBox "The translatedfile is in use!", vbExclamation
50330      End If
50340     Else
50350      MsgBox "The translatedfile doesn't exists!", vbExclamation
50360    End If
50370   End If
50380  End If
50390 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50400 Exit Sub
ErrPtnr_OnError:
50421 Select Case ErrPtnr.OnError("modTransTool", "AnalyzeCommandlineParameters")
      Case 0: Resume
50440 Case 1: Resume Next
50450 Case 2: Exit Sub
50460 Case 3: End
50470 End Select
50480 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function CheckTemplate(IniFile As String) As Boolean
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim ini As clsINI, secs As Collection, keys As Collection, i As Long, j As Long, _
  Filename As String
50060  SplitPath IniFile, , , Filename
50070  CheckTemplate = False
50080  Set ini = New clsINI
50090  ini.Filename = IniFile
50100  If ini.CheckIniFile = False Then
50110   MsgBox "Template file '" & Filename & "' not found!", vbCritical
50120   Exit Function
50130  End If
50140  Set secs = ini.GetAllSectionsFromInifile(, True)
50150  If secs.Count = 0 Then
50160   MsgBox "Template file '" & Filename & "' has no sections!"
50170   Exit Function
50180  End If
50190  For i = 1 To secs.Count
50200   Set keys = ini.GetAllKeysFromSection(secs(i), , , True)
50210   If keys.Count = 0 Then
50220    MsgBox "In Template file '" & Filename & "' are no keys in section [" & secs(i) & "]!", vbCritical
50230    Exit Function
50240   End If
50250   For j = 1 To keys.Count
50260    If Len(Trim$(keys(j)(1))) = 0 Then
50270     MsgBox "In Template file '" & Filename & "' the key '" & keys(j)(0) & "' in section [" & secs(i) & "] has no value!", vbCritical
50280     Exit Function
50290    End If
50300   Next j
50310  Next i
50320  CheckTemplate = True
50330 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50340 Exit Function
ErrPtnr_OnError:
50361 Select Case ErrPtnr.OnError("modTransTool", "CheckTemplate")
      Case 0: Resume
50380 Case 1: Resume Next
50390 Case 2: Exit Function
50400 Case 3: End
50410 End Select
50420 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub HighlightListitems(lsv As ListView, Items As Collection, _
 stb As StatusBar, pgb As Object, _
 Optional PanelName As String = "", _
 Optional HighlightForeColor As OLE_COLOR = vbRed, _
 Optional Bold As Boolean = False, _
 Optional StatusText As String = "Marking ...")
50010 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50020 On Error GoTo ErrPtnr_OnError
50030 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50040  Dim i As Long
50050  If Items.Count > 0 Then
50060   If Len(PanelName) > 0 Then
50070    stb.Panels(PanelName).Text = StatusText: stb.Refresh
50080   End If
50090   With pgb
50100    .Min = 0
50110    .Max = Items.Count
50120   End With
50130   For i = 1 To Items.Count
50140    pgb.Value = i
50150    If Items(i)(1) = 0 Then
50160      lsv.ListItems(Items(i)(0)).ForeColor = HighlightForeColor
50170      lsv.ListItems(Items(i)(0)).Bold = Bold
50180     Else
50190      lsv.ListItems(Items(i)(0)).ListSubItems(Items(i)(1)).ForeColor = HighlightForeColor
50200      lsv.ListItems(Items(i)(0)).ListSubItems(Items(i)(1)).Bold = Bold
50210    End If
50220   Next i
50230   With pgb
50240    .Min = 0
50250    .Max = 1
50260   End With
50270   lsv.Refresh
50280   With lsv.ListItems(Items(1)(0))
50290    .Selected = True
50300    .EnsureVisible
50310    .Selected = False
50320   End With
50330   If Len(PanelName) > 0 Then
50340    stb.Panels(PanelName).Text = "": stb.Refresh
50350   End If
50360  End If
50370 '---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
50380 Exit Sub
ErrPtnr_OnError:
50401 Select Case ErrPtnr.OnError("modTransTool", "HighlightListitems")
      Case 0: Resume
50420 Case 1: Resume Next
50430 Case 2: Exit Sub
50440 Case 3: End
50450 End Select
50460 '---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub


