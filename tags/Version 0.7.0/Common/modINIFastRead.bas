Attribute VB_Name = "modINIFastRead"
Option Explicit

Public Sub ReadINISection(ByVal FileName As String, ByRef Section As String, ByRef hHash As clsHash)
 Dim strINI() As String, curSection As String, i As Integer, hSplit() As String, _
  strTmp As String

 Section = UCase$(Section)
 strINI = ReadToLines(FileName, , , True)

 For i = 0 To UBound(strINI())
  strTmp = Trim$(strINI(i))
  If strTmp <> vbNullString Then
   Select Case Left$(strTmp, 1)
    Case "["
     curSection = UCase$(Mid$(strTmp, 2, Len(strTmp) - 2))
    Case ";"
    Case Else
     If Section = curSection Then
      hSplit() = Split(strTmp, "=", 2)
      If UBound(hSplit) >= 1 Then
       hHash.Add hSplit(0), hSplit(1)
      End If
     End If
   End Select
  End If
 Next i
End Sub

Private Function ReadToLines(FileName As String, Optional ErrNumber As Long, _
 Optional ErrDescription As String, Optional ShowHourGlass As Boolean = True) As String()

 Dim s As String, s1() As String

 ReDim s1(0)
 s = ReadToString(FileName)
 If ErrNumber = 0 Then
  s1() = Split(s, vbCrLf)
 End If

 ReadToLines = s1()
End Function

Private Function ReadToString(FileName As String) As String
 Dim FNr As Integer, s As String

 FNr = FreeFile
 Open FileName For Binary As #FNr
 s = Space$(LOF(FNr)): Get #FNr, , s
 Close #FNr
 ReadToString = s

 Exit Function
End Function
