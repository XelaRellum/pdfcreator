Attribute VB_Name = "modJPEG"
Option Explicit

Public Type tJPEGInfo
 Width As Long
 Height As Long
 BitsPerPixel As Long
End Type

Public Function GetJPEGInfo(sourceFileName As String) As tJPEGInfo
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim fn As Long, buffer() As Byte, chunkB(0 To 1) As Byte, lChunkB(0 To 1) As Byte, lChunk As Long, SOIMarker(0 To 1) As Byte
50020
50030  fn = FreeFile()
50040  Open sourceFileName For Binary As #fn
50050  Get #fn, , SOIMarker
50060  If SOIMarker(0) = &HFF And SOIMarker(1) = &HD8 Then
50070   Do Until EOF(fn)
50080    Get #fn, , chunkB
50090    Get #fn, , lChunkB
50100    lChunk = lChunkB(0) * 256& + lChunkB(1) - 2
50110    ReDim buffer(0 To lChunk - 1)
50120    Get #fn, , buffer
50130    If chunkB(0) = &HFF And chunkB(1) = &HC0 Then ' SOF chunk
50140     GetJPEGInfo.Height = buffer(1) * 256 + buffer(2)
50150     GetJPEGInfo.Width = buffer(3) * 256 + buffer(4)
50160     GetJPEGInfo.BitsPerPixel = buffer(5) * 8
50170     Close #fn
50180     Exit Function
50190    End If
50200    DoEvents
50210   Loop
50220  End If
50230  Close #fn
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modJPEG", "GetJPEGInfo")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function


