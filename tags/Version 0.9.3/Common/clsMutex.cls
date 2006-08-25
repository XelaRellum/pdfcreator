VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMutex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private hMutex As Long

Public Function CheckMutex(ByVal Mutexname As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If IDERunning = True Then
50020    CheckMutex = App.PrevInstance
50030   Else
50040    hMutex = CreateMutexA(ByVal 0&, 1, Mutexname)
50050    If (Err.LastDllError = ERROR_ALREADY_EXISTS) Or (Err.LastDllError = ERROR_ACCESS_DENIED) Then
50060      CheckMutex = True
50070     Else
50080      CheckMutex = False
50090    End If
50100    If (hMutex <> 0) Then
50110     CloseHandle hMutex
50120    End If
50130  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsMutex", "CheckMutex")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function CreateMutex(ByVal Mutexname As String) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  CreateMutex = 0
50020  hMutex = CreateMutexA(ByVal 0&, 1, Mutexname)
50030  If (Err.LastDllError = ERROR_ALREADY_EXISTS) Then
50040    CreateMutex = 0
50050   Else
50060    CreateMutex = hMutex
50070  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsMutex", "CreateMutex")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function IDERunning() As Boolean
 On Error GoTo ErrorHandler
 Debug.Print 1 / 0
 Exit Function
ErrorHandler:
 IDERunning = True
End Function

Public Function CloseMutex() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim res As Long
50020  If (hMutex <> 0) Then
50030   res = CloseHandle(hMutex)
50040  End If
50050  If res = 0 Then
50060    CloseMutex = 0
50070    hMutex = 0
50080   Else
50090    CloseMutex = res
50100  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsMutex", "CloseMutex")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

