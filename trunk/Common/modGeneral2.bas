Attribute VB_Name = "modGeneral2"
Option Explicit

Public Sub CombineFiles(ByVal Filename As String, Files As Collection, _
 Optional BufferSize As Long = 65536, Optional stb As StatusBar)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, j As Long, fnSource As Long, fnDest As Long, sBuffer As String, _
  aLen As Double, tLen As Double, bsize As Long, fpos As Long
50030
50040  bsize = BufferSize
50050  Filename = Trim$(Filename)
50060  If Filename = vbNullString Or Files.Count = 0 Or Right$(Filename, 1) = "\" Then
50070   Exit Sub
50080  End If
50090  If Files.Count = 1 Then
50100   Exit Sub
50110  End If
50120  fnDest = FreeFile
50130  aLen = 0: tLen = 0
50140  For i = 1 To Files.Count
50150   aLen = aLen + FileLen(Files.Item(i))
50160  Next i
50170  Open Filename For Binary As #fnDest
50180  For i = 1 To Files.Count
50190   DoEvents
50200   If FileExists(Files.Item(i)) = False Then
50210    MsgBox LanguageStrings.MessagesMsg14 & vbCrLf & vbCrLf & Files.Item(i)
50220   End If
50230   If FileLen(Files.Item(i)) > 0 Then
50240    fnSource = FreeFile
50250    Open Files.Item(i) For Binary Access Read As #fnSource
50260    If bsize > LOF(fnSource) Then
50270     bsize = LOF(fnSource)
50280    End If
50290    fpos = 1
50300    For j = 1 To LOF(fnSource) \ bsize
50310     fpos = (j - 1) * bsize + 1
50320     Seek #fnSource, fpos
50330     sBuffer = Input(bsize, fnSource)
50340     Put #fnDest, , sBuffer
50350     tLen = tLen + bsize
50360     If Not stb Is Nothing Then
50370      stb.Panels("Percent").Text = Format(CDbl(tLen) / CDbl(aLen), "0.0%")
50380     End If
50390     DoEvents
50400    Next j
50410    If LOF(fnSource) > (j - 1) * bsize Then
50420     fpos = (j - 1) * bsize + 1
50430     Seek #fnSource, fpos
50440     sBuffer = Input(LOF(fnSource) - (j - 1) * bsize, fnSource)
50450     Put #fnDest, , sBuffer
50460     tLen = tLen + (LOF(fnSource) - (j - 1) * bsize)
50470    End If
50480    Close #fnSource
50490   End If
50500   DoEvents
50510  Next i
50520  For i = 1 To Files.Count
50530   KillFile Files.Item(i)
50540   KillInfoSpoolfile Files.Item(i)
50550   DoEvents
50560  Next i
50570  Close #fnDest
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral2", "CombineFiles")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub


Public Sub CombineFilesOld(ByVal Filename As String, Files As Collection, stb As StatusBar)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, fnSource As Long, fnDest As Long, sBuffer As String, _
  aLen As Double, tLen As Double
50030
50040  Filename = Trim$(Filename)
50050  If Filename = vbNullString Or Files.Count = 0 Or Right$(Filename, 1) = "\" Then
50060   Exit Sub
50070  End If
50080  If FileExists(Filename) = True Then
50090   Exit Sub
50100  End If
50110  If Files.Count = 1 Then
50120   Exit Sub
50130  End If
50140  fnDest = FreeFile
50150  aLen = 0
50160  For i = 1 To Files.Count
50170   aLen = aLen + FileLen(Files.Item(i))
50180  Next i
50190  Open Filename For Binary As #fnDest
50200  For i = 1 To Files.Count
50210   DoEvents
50220   If FileLen(Files.Item(i)) > 0 Then
50230    fnSource = FreeFile
50240    Open Files.Item(i) For Binary Access Read As #fnSource
50250    sBuffer = String(LOF(fnSource), Chr$(0))
50260    Get #fnSource, , sBuffer
50270    Put #fnDest, , sBuffer
50280    Close #fnSource
50290   End If
50300   tLen = tLen + FileLen(Files.Item(i))
50310   stb.Panels("Percent").Text = Format$(tLen / aLen, "0.0%")
50320   KillFile Files.Item(i)
50330   DoEvents
50340  Next i
50350  Close #fnDest
50360  stb.Panels("Percent").Text = vbNullString
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral2", "CombineFilesOld")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub


