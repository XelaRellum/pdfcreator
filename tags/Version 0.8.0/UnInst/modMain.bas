Attribute VB_Name = "modMain"
Option Explicit

Sub Main()
 Dim a As Long, LanguagePath As String, Languagefile As String, _
  mutex As clsMutex, PDFCreatorINIPath As String, Temppath As String

 On Error Resume Next

 If UCase$(CommandSwitch("UI", True)) = "TRUE" Then
  Set mutex = New clsMutex
  If mutex.CheckMutex(UnInst_GUID) = False Then
    mutex.CreateMutex UnInst_GUID
   Else
    End
  End If
  ReadVersionInfo
  Temppath = GetTempPath
  
  If CheckPath(GetMyAppData) = True Then
    PDFCreatorINIFile = CompletePath(GetMyAppData) & "PDFcreator\PDFCreator.ini"
    PDFCreatorINIPath = CompletePath(GetMyAppData) & "PDFcreator"
   Else
    PDFCreatorINIFile = App.Path & "\PDFCreator.ini"
    PDFCreatorINIPath = App.Path
  End If
  Options = ReadOptions
  LanguagePath = App.Path & "\Languages\"
  Languagefile = LanguagePath & Options.Language & ".ini"
  LoadLanguage Languagefile
  a = 0
  If UCase$(CommandSwitch("SAVEOPTIONS=", True)) = "YES" Or _
     UCase$(CommandSwitch("SAVEOPTIONS=", True)) = "TRUE" Then
   a = vbYes
  End If
  If UCase$(CommandSwitch("SAVEOPTIONS=", True)) = "NO" Or _
     UCase$(CommandSwitch("SAVEOPTIONS=", True)) = "FALSE" Then
   a = vbNo
  End If
  If a = 0 Then
   a = MsgBox(LanguageStrings.MessagesMsg13, vbYesNo)
  End If
  If a = vbYes Then
   If LenB(Dir(PDFCreatorINIFile)) > 0 Then
    Kill PDFCreatorINIFile
    If UCase$(PDFCreatorINIPath) = UCase$(CompletePath(GetMyAppData) & "PDFcreator") Then
     RemoveCompletePath PDFCreatorINIPath
    End If
   End If
  End If
  Temppath = Mid(Temppath, 1, InStrRev(Temppath, "\", Len(Temppath) - 1) - 1)
  If LenB(Dir(Temppath, vbDirectory)) > 0 Then
   RemoveCompletePath Temppath, False
  End If
  Screen.MousePointer = vbHourglass
  RemoveExplorerIntegration
  Screen.MousePointer = vbNormal
  mutex.CloseMutex
  Set mutex = Nothing
 End If
End Sub
