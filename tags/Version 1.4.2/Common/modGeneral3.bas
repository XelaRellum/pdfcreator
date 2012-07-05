Attribute VB_Name = "modGeneral3"
Option Explicit
                    
Public Function CompletePath(Path As String) As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If Len(Path) = 0 Then
50020   Exit Function
50030  End If
50040  Path = Trim$(Path)
50050  If Right$(Path, 1) = "\" Then
50060    CompletePath = Path
50070   Else
50080    CompletePath = Path & "\"
50090  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral3", "CompletePath")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function InCollection(colTest As Collection, sKey As String) As Boolean
'
' Check to see if item [sKey] is in collection [colTest].
' Return True if it is, false if not
'
 On Error GoTo ErrorHandler
 If VarType(colTest.Item(sKey)) = vbObject Then
  '
  ' This test will indicate if the item actually exists in the
  ' collection. No further checking is needed.
  '
 End If
 InCollection = True
Exit Function
ErrorHandler:
   InCollection = False
End Function

Public Function ProgramIsRunning(GUIDStr As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim mutex As clsMutex
50020  Set mutex = New clsMutex
50030  ProgramIsRunning = mutex.CheckMutex(GUIDStr)
50040  Set mutex = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral3", "ProgramIsRunning")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub QuickSortSpoolFiles(ByRef vSort() As clsSpoolFile, Optional ByVal lngStart As Variant, Optional ByVal lngEnd As Variant)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, j As Long, h As clsSpoolFile, x As String
50020
50030  If IsMissing(lngStart) Then lngStart = LBound(vSort)
50040  If IsMissing(lngEnd) Then lngEnd = UBound(vSort)
50050
50060  i = lngStart: j = lngEnd
50070  x = vSort((lngStart + lngEnd) / 2).FileDateTimeJobIdKey
50080
50090  Do
50100   While (StrComp(vSort(i).FileDateTimeJobIdKey, x) < 0): i = i + 1: Wend
50110   While (StrComp(vSort(j).FileDateTimeJobIdKey, x) > 0): j = j - 1: Wend
50120
50130   If (i <= j) Then
50140    Set h = vSort(i)
50150    Set vSort(i) = vSort(j)
50160    Set vSort(j) = h
50170    i = i + 1: j = j - 1
50180   End If
50190  Loop Until (i > j)
50200
50210  If (lngStart < j) Then QuickSortSpoolFiles vSort, lngStart, j
50220  If (i < lngEnd) Then QuickSortSpoolFiles vSort, i, lngEnd
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral3", "QuickSortSpoolFiles")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function GetFileLength(filename As String) As Variant
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim hFile As Long, Lng() As Long
50020  ReDim Lng(3)
50030  hFile = CreateFile(filename, GENERIC_WRITE Or GENERIC_READ, 0, 0, OPEN_ALWAYS, _
                    FILE_ATTRIBUTE_NORMAL Or FILE_FLAG_SEQUENTIAL_SCAN, 0)
50050  Lng(0) = 14
50060  Lng(1) = 0
50070  Lng(2) = GetFileSize(hFile, Lng(3))
50080  CloseHandle hFile
50090
50100  MoveMemory GetFileLength, Lng(0), 16
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modGeneral3", "GetFileLength")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
