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
50050  strINI = ReadToLines(Filename, , , True)
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

Private Function ReadToLines(Filename As String, Optional ErrNumber As Long, _
 Optional ErrDescription As String, Optional ShowHourGlass As Boolean = True) As String()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020  Dim s As String, s1() As String
50030
50040  ReDim s1(0)
50050  s = ReadToString(Filename)
50060  If ErrNumber = 0 Then
50070   s1() = Split(s, vbCrLf)
50080  End If
50090
50100  ReadToLines = s1()
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
50010  Dim FNr As Integer, s As String
50020
50030  FNr = FreeFile
50040  Open Filename For Binary As #FNr
50050  s = Space$(LOF(FNr)): Get #FNr, , s
50060  Close #FNr
50070  ReadToString = s
50080
50090  Exit Function
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
