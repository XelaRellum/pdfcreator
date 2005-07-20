Attribute VB_Name = "modINIFastRead"
Option Explicit

Public Sub ReadINISection(ByVal Filename As String, ByRef Section As String, ByRef hHash As clsHash)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim strINI() As String, curSection As String, i As Integer, hSplit() As String, _
  strTmp As String
50030
50040  Section = UCase$(Section)
50050  strINI = ReadToLines(Filename)
50060
50070  For i = 0 To UBound(strINI())
50080   strTmp = Trim$(strINI(i))
50090   If strTmp <> vbNullString Then
50101    Select Case Left$(strTmp, 1)
          Case "["
50120      curSection = UCase$(Mid$(strTmp, 2, Len(strTmp) - 2))
50130     Case ";"
50140     Case Else
50150      If Section = curSection Then
50160       hSplit() = Split(strTmp, "=", 2)
50170       If UBound(hSplit) >= 1 Then
50180        hHash.Add hSplit(0), hSplit(1)
50190       End If
50200      End If
50210    End Select
50220   End If
50230  Next i
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modINIFastRead", "ReadINISection")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function ReadToLines(Filename As String) As String()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim s As String, s1() As String
50020  ReDim s1(0)
50030  s = ReadToString(Filename)
50040  If InStr(s, vbCrLf) Then
50050    s1() = Split(s, vbCrLf)
50060   Else
50070    s1(0) = s
50080  End If
50090  ReadToLines = s1()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modINIFastRead", "ReadToLines")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function ReadToString(Filename As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim fn As Integer, fl As Long, s As String
50020  s = ""
50030  If FileExists(Filename) Then
50040   fn = FreeFile
50050   Open Filename For Binary Access Read Shared As #fn
50060   fl = LOF(fn)
50070   If fl > 0 Then
50080    s = Space$(fl)
50090    Get #fn, , s
50100   End If
50110   Close #fn
50120  End If
50130  ReadToString = s
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modINIFastRead", "ReadToString")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
