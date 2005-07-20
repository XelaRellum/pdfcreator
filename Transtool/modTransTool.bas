Attribute VB_Name = "modTransTool"
Option Explicit

Public TemplateInifile As String, TranslatedInifile As String, _
 RecentFilesCount As Long, LastSearchstrings As Collection, hLItems As Collection, _
 lsvColor As OLE_COLOR, lsvBold As Boolean

Public Sub Main()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020  ' Reduce the working size of used memory
50030  Call SetProcessWorkingSetSize(GetCurrentProcess(), -1, -1)
50040  RecentFilesCount = 10
50050
50060  TemplateInifile = App.Path & "\english.ini"
50070  Set hLItems = New Collection
50080  AnalyzeCommandlineParameters
50090
50100  If CheckInstance Then
50110   CheckProgramInstances
50120  End If
50130
50140  If Not NoStart Then
50150   If CheckTemplate(TemplateInifile) = False Then
50160    TemplateInifile = ""
50170   End If
50180   Set LastSearchstrings = New Collection
50190   If IsWin9xMe = False And IsWinNT4 = False And IsWin2000 = False Then
50200    InitCommonControls
50210   End If
50220   frmMain.Show
50230  End If
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
50010  Dim TemplateFile As String, TranslatedFile As String, Path As String
50020  ' The program has commandswitches
50030  ' -Templatefile: The template file
50040  ' -Translatedfile: The translated file
50050
50060  If Len(VBA.Command$) > 0 Then
50070   ' Check running instance
50080   If UCase$(CommandSwitch("Check", False)) = "INSTANCE" Then
50090     CheckInstance = True
50100    Else
50110     CheckInstance = False
50120   End If
50130   If UCase$(CommandSwitch("NO", False)) = "START" Then
50140     NoStart = True
50150    Else
50160     NoStart = False
50170   End If
50180   TemplateFile = CommandSwitch("TemplateFile", True)
50190   SplitPath TemplateFile, , Path
50200   If LenB(Path) = 0 And LenB(TemplateFile) > 0 Then
50210    TemplateFile = CompletePath(App.Path) & TemplateFile
50220   End If
50230   If LenB(TemplateFile) > 0 Then
50240    If FileExists(TemplateFile) Then
50250      TemplateInifile = TemplateFile
50260     Else
50270      MsgBox "The template file doesn't exists!", vbExclamation
50280    End If
50290   End If
50300   TranslatedFile = CommandSwitch("TranslatedFile", True)
50310   SplitPath TranslatedFile, , Path
50320   If LenB(Path) = 0 And LenB(TranslatedFile) > 0 Then
50330    TranslatedFile = CompletePath(App.Path) & TranslatedFile
50340   End If
50350   If LenB(TranslatedFile) > 0 Then
50360    If FileExists(TranslatedFile) = True Then
50370      If FileInUse(TranslatedFile) = False Then
50380        TranslatedInifile = TranslatedFile
50390       Else
50400        MsgBox "The translatedfile is in use!", vbExclamation
50410      End If
50420     Else
50430      MsgBox "The translatedfile doesn't exists!", vbExclamation
50440    End If
50450   End If
50460  End If
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
 stb As StatusBar, pgb As Object, _
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
50060   With pgb
50070    .Min = 0
50080    .Max = Items.Count
50090   End With
50100   For i = 1 To Items.Count
50110    pgb.Value = i
50120    If Items(i)(1) = 0 Then
50130      lsv.ListItems(Items(i)(0)).ForeColor = HighlightForeColor
50140      lsv.ListItems(Items(i)(0)).Bold = Bold
50150     Else
50160      lsv.ListItems(Items(i)(0)).ListSubItems(Items(i)(1)).ForeColor = HighlightForeColor
50170      lsv.ListItems(Items(i)(0)).ListSubItems(Items(i)(1)).Bold = Bold
50180    End If
50190   Next i
50200   With pgb
50210    .Min = 0
50220    .Max = 1
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


