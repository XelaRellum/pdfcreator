VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents dl As clsDownload
Attribute dl.VB_VarHelpID = -1

Public Sub CheckForUpdates()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim updStr As String, updStrA() As String, aw As Long
50020  Set dl = New clsDownload
50030  updStr = dl.DownloadString(UpdateURL)
50040 '   updStr = dl.DownloadString("http://localhost:8080/update.txt")
50050  If Len(updStr) > 0 Then
50060    If CheckPDFCreatorVersion(updStr) > 0 Then
50070      updStrA = Split(updStr, ".")
50080      If updStrA(3) = 0 Then
50090        updStr = "[" & updStrA(0) & "." & updStrA(1) & "." & updStrA(2) & "]"
50100       Else
50110        updStr = "[" & updStrA(0) & "." & updStrA(1) & "." & updStrA(2) & " Beta " & updStrA(3) & "]"
50120      End If
50130      aw = MsgBox(Replace$(LanguageStrings.MessagesMsg32, "%1", updStr), vbYesNo + vbQuestion)
50140      If aw = vbYes Then
50150       OpenDocument "http://www.pdfforge.org/products/pdfcreator/download"
50160      End If
50170     Else
50180      MsgBox LanguageStrings.MessagesMsg33, vbOKOnly + vbInformation
50190    End If
50200   Else
50210    MsgBox LanguageStrings.MessagesMsg31 & ": " & dl.ErrorDescription & " [" & dl.ErrorNumber & "]", vbOKOnly + vbExclamation
50220  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsUpdate", "CheckForUpdates")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
