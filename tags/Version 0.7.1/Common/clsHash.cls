VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---[ Associative Array for VB, v1.0 ]---------------------'
'                                                          '
' Associative Array for VB is Copyright (C) 2000 Ashley    '
' Ross, aka Eldren.                                        '
' It is under the GNU GPL, version 2 or greater. This      '
' means that you may use and modify it freely, within the  '
' terms of the GPL. See http://www.fsf.org/ for more info. '
'                                                          '
' If you make any changes to this, please tell me about    '
' them, possibly including the newer copy in your email.   '
' Thanks.                                                  '
'                                                          '
'---------[ Ashley Ross aka Eldren ~ eldrenr@yahoo.com ]---'

Option Explicit
Option Compare Text

Private aKey() As String, aVal() As Variant, Upper As Long, Last As Long

Public Sub Add(Key As String, Value As Variant)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, thisKey As String, found As Boolean
50020
50030  ReDim Preserve aKey(Upper)
50040  ReDim Preserve aVal(Upper)
50050
50060  For i = 0 To Upper
50070   thisKey = aKey(i)
50080   If StrComp(thisKey, Key) = 0 Then
50090    found = True
50100    Exit For
50110   End If
50120  Next i
50130
50140  If found Then
50150    aKey(i) = Key
50160    aVal(i) = Value
50170    Last = i
50180   Else
50190    aKey(Upper) = Key
50200    aVal(Upper) = Value
50210    Last = Upper
50220    Upper = Upper + 1
50230  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsHash", "Add")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function Retrieve(Key As String, Optional Default As Variant = 0) As Variant
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, thisKey As String, found As Boolean
50020
50030  If Upper > 0 Then
50040   For i = 0 To Upper - 1
50050    thisKey = aKey(i)
50060    If StrComp(thisKey, Key) = 0 Then
50070     found = True
50080     Exit For
50090    End If
50100   Next i
50110
50120   If found Then
50130    Retrieve = aVal(i)
50140    Last = i
50150    Exit Function
50160   End If
50170  End If
50180
50190  If Default = 0 Then
50200    Retrieve = -1
50210   Else
50220    Retrieve = Default
50230  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsHash", "Retrieve")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub Delete(Key As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, i2 As Long, thisKey As String, found As Boolean
50020
50030  If Upper = 0 Then Exit Sub
50040  For i = 0 To Upper - 1
50050   thisKey = aKey(i)
50060   If StrComp(thisKey, Key) = 0 Then
50070    found = True
50080    Exit For
50090   End If
50100  Next i
50110
50120  If found = False Then Exit Sub
50130
50140  Last = i
50150
50160  i = i + 1
50170  For i2 = i To Upper - 1
50180   aKey(i2 - 1) = aKey(i2)
50190   aVal(i2 - 1) = aVal(i2)
50200  Next i2
50210
50220  Upper = Upper - 1
50230  ReDim Preserve aKey(Upper)
50240  ReDim Preserve aVal(Upper)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsHash", "Delete")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function Keys() As String()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, Key() As String
50020  ReDim Key(0 To Upper - 1)
50030
50040  For i = 0 To Upper - 1
50050   Key(i) = aKey(i)
50060  Next i
50070
50080  Keys = Key
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsHash", "Keys")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function Values() As Variant()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, Val() As Variant
50020  ReDim Val(0 To Upper - 1)
50030
50040  For i = 0 To Upper - 1
50050   Val(i) = aVal(i)
50060  Next i
50070
50080  Values = Val
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsHash", "Values")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function Every() As Variant()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim KeyVal() As Variant
50020  ReDim KeyVal(0 To 1)
50030
50040  Last = Last + 1
50050  If Last > Upper - 1 Then Last = 0
50060  KeyVal(0) = aKey(Last)
50070  KeyVal(1) = aVal(Last)
50080  Every = KeyVal
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsHash", "Every")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
