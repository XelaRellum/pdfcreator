Attribute VB_Name = "modMain"
Option Explicit

Sub Main()
 On Local Error Resume Next
 Dim a As Long, LanguagePath As String, Languagefile As String, _
  mutex As clsMutex, PDFCreatorINIPath As String
 If UCase$(CommandSwitch("UI", True)) = "TRUE" Then
  Set mutex = New clsMutex
  If mutex.CheckMutex(UnInst_GUID) = False Then
    mutex.CreateMutex UnInst_GUID
   Else
    End
  End If
  If CheckPath(GetMyAppData) = True Then
    PDFCreatorINIFile = GetMyAppData & "PDFcreator\PDFCreator.ini"
    PDFCreatorINIPath = GetMyAppData & "PDFcreator"
   Else
    PDFCreatorINIFile = App.Path & "\PDFCreator.ini"
    PDFCreatorINIPath = App.Path
  End If
  Options = ReadOptions
  LanguagePath = App.Path & "\Languages\"
  Languagefile = LanguagePath & Options.Language & ".ini"
  LoadLanguage Languagefile
  a = MsgBox(LanguageStrings.MessagesMsg13, vbYesNo)
  If a = vbYes Then
   If Dir(PDFCreatorINIFile) <> "" Then
    Kill PDFCreatorINIFile
    If UCase$(PDFCreatorINIPath) = UCase$(GetMyAppData & "PDFcreator") Then
     Kill PDFCreatorINIPath & "\*.*"
     RmDir PDFCreatorINIPath
    End If
   End If
  End If
  mutex.CloseMutex
  Set mutex = Nothing
 End If
End Sub
