Attribute VB_Name = "modPDFCreator"
Option Explicit

Public Const PDFCreator_GUID = "{A7332D94-E8FE-40B2-937F-8515FC0FF52F}"

Public Const PDFCreatorLogfile = "PDFCreator.log"

Public PDFCreatorINIFile As String, PrinterStop As Boolean, Printing As Boolean

Public Function ReadLogfile() As String
 Dim fn As Long, bufStr As String, tStr As String
 
 If Dir(App.Path & "\" & PDFCreatorLogfile) = "" Then
  Exit Function
 End If
 
 fn = FreeFile
 Open App.Path & "\" & PDFCreatorLogfile For Input As #fn
 bufStr = ""
 While Not EOF(fn)
  If Len(bufStr) = 0 Then
    Line Input #fn, bufStr
   Else
    Line Input #fn, tStr
    If Trim$(tStr) <> "" Then
     bufStr = bufStr & vbCrLf & tStr
    End If
  End If
 Wend
 Close #fn
 ReadLogfile = bufStr
End Function

Public Sub ClearLogfile()
 Dim fn As Long
 fn = FreeFile
 Open App.Path & "\" & PDFCreatorLogfile For Output As #fn
 Close #fn
End Sub

Public Sub WriteLogfile(Logtext As String)
 Dim fn As Long, i As Long, bufStr As String, s() As String
 
 bufStr = ReadLogfile
 
 fn = FreeFile
 Open App.Path & "\" & PDFCreatorLogfile For Output As #fn
 
 If Len(bufStr) > 0 Then
   s = Split(bufStr, vbCrLf)
   Print #fn, Now & ": " & Logtext
   For i = LBound(s) To UBound(s)
    Print #fn, Trim$(Replace(s(i), vbCrLf, ""))
   Next i
  Else
   Print #fn, Now & ": " & Logtext
 End If
 Close #fn
End Sub

Public Sub IfLoggingWriteLogfile(Logtext As String)
 Dim fn As Long, i As Long, bufStr As String, s() As String, _
  Options As tOptions, tStr As String
 
 Options = ReadOptions
 
 If Options.Logging = 0 Then
  Exit Sub
 End If
 
 bufStr = ReadLogfile
 
 fn = FreeFile
 Open App.Path & "\" & PDFCreatorLogfile For Output As #fn
 
 If Len(bufStr) > 0 Then
   s = Split(bufStr, vbCrLf)
   Print #fn, Now & ": " & Logtext
   If Options.LogLines < UBound(s) - 1 Then
     For i = 2 To Options.LogLines
      tStr = s(i - 2)
      Print #fn, s(i - 2)
     Next i
    Else
     For i = LBound(s) To UBound(s)
      tStr = s(i)
      Print #fn, Trim$(Replace$(s(i), vbCrLf, ""))
     Next i
   End If
  Else
   Print #fn, Now & ": " & Logtext
 End If
 Close #fn
End Sub

Public Sub IfLoggingShowLogfile(Optional frmLog As Form, Optional frmMain As Form)
 Dim Options As tOptions
 
 Options = ReadOptions
 
 If Options.Logging = 0 Then
  Exit Sub
 End If
 If IsMissing(frmLog) = False And IsMissing(frmMain) = False Then
  frmLog.Show vbModal, frmMain
 End If
End Sub

Public Sub CreatePDFCreatorTempfolder()
 Dim Temppath As String
 Temppath = GetPDFCreatorTempfolder
 If Dir(Temppath, vbDirectory) = "" Then
  MkDir Temppath
 End If
End Sub

Public Function GetPDFCreatorTempfolder() As String
 Dim Temppath As String
 Temppath = GetTempPath
 GetPDFCreatorTempfolder = Temppath & "PDFCreator\"
End Function
