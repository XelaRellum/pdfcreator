VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function cOpenFileDialog(ByRef files As Variant, Optional ByVal InitFilename As String = "", _
 Optional ByVal Filter As String, Optional ByVal DefaultFileExtension As String = "*.*", _
 Optional ByVal InitDir As String = "", Optional ByVal DialogTitle As String = "", _
 Optional ByVal Flags As Long, Optional ByVal hwnd As Long = 0, Optional ByVal FilterIndex As Long = 1) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010
50020  Dim res As Long, filesCol As Collection, i
50030
50040  res = OpenFileDialog(filesCol, InitFilename, Filter, DefaultFileExtension, InitDir, DialogTitle, Flags, hwnd, FilterIndex)
50050
50060  cOpenFileDialog = res
50070
50080  If res > 0 Then
50090   files = filesCol(1)
50100   For i = 2 To filesCol.Count
50110    files = files & "|" & filesCol(i)
50120   Next i
50130  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsTools", "cOpenFileDialog")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function


Public Function cSaveFileDialog(filename As Variant, Optional InitFilename As String = "", _
 Optional Filter As String, Optional DefaultFileExtension As String = "*.*", _
 Optional InitDir As String = "", Optional DialogTitle As String = "", _
 Optional Flags As Long, Optional hwnd As Long = 0, Optional FilterIndex As Long = 1) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim res As Long, filenameStr As String, i
50020
50030  cSaveFileDialog = SaveFileDialog(filenameStr, InitFilename, Filter, DefaultFileExtension, InitDir, DialogTitle, Flags, hwnd, FilterIndex)
50040  filename = filenameStr
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsTools", "cSaveFileDialog")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function


