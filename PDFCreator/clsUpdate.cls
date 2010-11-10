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

Public Sub CheckForUpdates(Optional ShowMessageNoNewUpdates As Boolean = False, Optional ShowErrorMessage As Boolean = True)
 Dim updStr As String, updStrA() As String, aw As Long, errStr As String
 Set dl = New clsDownload
 updStr = dl.DownloadString(UpdateURL)
'   updStr = dl.DownloadString("http://localhost:8080/update.txt")
 If Len(updStr) > 0 Then
   If CheckPDFCreatorVersion(updStr) > 0 Then
     updStrA = Split(updStr, ".")
     If updStrA(3) = 0 Then
       updStr = "[" & updStrA(0) & "." & updStrA(1) & "." & updStrA(2) & "]"
      Else
       updStr = "[" & updStrA(0) & "." & updStrA(1) & "." & updStrA(2) & " Beta " & updStrA(3) & "]"
     End If
     aw = MsgBox(Replace$(LanguageStrings.MessagesMsg32, "%1", updStr), vbYesNo + vbQuestion)
     If aw = vbYes Then
      OpenDocument "http://www.pdfforge.org/products/pdfcreator/download"
     End If
    Else
     If ShowMessageNoNewUpdates Then
      IfLoggingWriteLogfile LanguageStrings.MessagesMsg33
      MsgBox LanguageStrings.MessagesMsg33, vbOKOnly + vbInformation
     End If
   End If
  Else
   errStr = LanguageStrings.MessagesMsg31 & ": " & dl.ErrorDescription & " [" & dl.ErrorNumber & "]"
   IfLoggingWriteLogfile "CheckForUpdates: " & errStr
   If ShowErrorMessage Then
    MsgBox errStr, vbOKOnly + vbExclamation
   End If
 End If
End Sub
