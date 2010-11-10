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

Public Sub CheckForUpdates(Optional ShowMessageNoNewUpdates As Boolean = False, Optional ShowErrorMessage As Boolean = True, Optional TimeOutInMs As Long = 10000)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim updStr As String, updStrA() As String, aw As Long, errStr As String, mp As Long
50020  Set dl = New clsDownload
50030  mp = Screen.MousePointer
50040  Screen.MousePointer = vbHourglass
50050  updStr = dl.DownloadString(UpdateURL, TimeOutInMs)
50060  Screen.MousePointer = mp
50070  If Len(updStr) > 0 Then
50080    If CheckPDFCreatorVersion(updStr) > 0 Then
50090      updStrA = Split(updStr, ".")
50100      If updStrA(3) = 0 Then
50110        updStr = "[" & updStrA(0) & "." & updStrA(1) & "." & updStrA(2) & "]"
50120       Else
50130        updStr = "[" & updStrA(0) & "." & updStrA(1) & "." & updStrA(2) & " Beta " & updStrA(3) & "]"
50140      End If
50150      aw = MsgBox(Replace$(LanguageStrings.MessagesMsg32, "%1", updStr), vbYesNo + vbQuestion)
50160      If aw = vbYes Then
50170       OpenDocument "http://www.pdfforge.org/products/pdfcreator/download"
50180      End If
50190     Else
50200      If ShowMessageNoNewUpdates Then
50210       IfLoggingWriteLogfile LanguageStrings.MessagesMsg33
50220       MsgBox LanguageStrings.MessagesMsg33, vbOKOnly + vbInformation
50230      End If
50240    End If
50250   Else
50260    errStr = LanguageStrings.MessagesMsg31 & ": " & dl.ErrorDescription & " [" & dl.ErrorNumber & "]"
50270    IfLoggingWriteLogfile "CheckForUpdates: " & errStr
50280    If ShowErrorMessage Then
50290     MsgBox errStr, vbOKOnly + vbExclamation
50300    End If
50310  End If
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
