Attribute VB_Name = "modEllipsis"
Option Explicit

Public Function ShortenPath(ByVal hDC As Long, ByVal Path As String, ByVal WidthInPixel As Long) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim nRect As Rect, nPos As Long, nPath As String
50020
50030  nRect.Right = WidthInPixel
50040  Path = Path & Chr$(0)
50050  DrawText hDC, Path, -1, nRect, DT_MODIFYSTRING Or DT_SINGLELINE Or DT_PATH_ELLIPSIS
50060  nPos = InStr(Path, Chr$(0))
50070  If nPos Then
50080    ShortenPath = Left$(Path, nPos - 1)
50090   Else
50100    ShortenPath = Path
50110  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modEllipsis", "ShortenPath")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function ShortenString(ByVal hDC As Long, ByVal Str As String, ByVal WidthInPixel As Long) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim nRect As Rect, nPos As Long, nPath As String
50020
50030  nRect.Right = WidthInPixel
50040  Str = Str & Chr$(0)
50050  DrawText hDC, Str, -1, nRect, DT_MODIFYSTRING Or DT_SINGLELINE Or DT_END_ELLIPSIS
50060  nPos = InStr(Str, Chr$(0))
50070  If nPos > 0 Then
50080    ShortenString = Left$(Str, nPos - 1)
50090   Else
50100    ShortenString = Str
50110  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modEllipsis", "ShortenString")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function


