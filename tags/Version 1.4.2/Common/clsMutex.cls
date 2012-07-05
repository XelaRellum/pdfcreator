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

Public Property Get MutexHandle() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  MutexHandle = hMutex
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsMutex", "MutexHandle [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Function CheckMutex(ByVal MutexName As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim res As Long, hM As Long
50020  If IDERunning = True Then
50030    CheckMutex = App.PrevInstance
50040   Else
50050    hM = CreateMutexA(ByVal 0&, 1, MutexName)
50060    If (Err.LastDllError = ERROR_ALREADY_EXISTS) Or (Err.LastDllError = ERROR_ACCESS_DENIED) Then
50070      CheckMutex = True
50080     Else
50090      CheckMutex = False
50100    End If
50110    If (hM <> 0) Then
50120     res = ReleaseMutex(hM)
50130     res = CloseHandle(hM)
50140    End If
50150  End If
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

Public Function CreateMutex(ByVal MutexName As String) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim res As Long
50020  CreateMutex = 0
50030  hMutex = CreateMutexA(ByVal 0&, 1, MutexName)
50040  If (Err.LastDllError = ERROR_ALREADY_EXISTS) Then
50050    If (hMutex <> 0) Then
50060     res = ReleaseMutex(hMutex)
50070     res = CloseHandle(hMutex)
50080    End If
50090    CreateMutex = 0
50100   Else
50110    CreateMutex = hMutex
50120  End If
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

Public Function CloseMutex() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim res As Long
50020  If (hMutex <> 0) Then
50030   res = ReleaseMutex(hMutex)
50040   res = CloseHandle(hMutex)
50050  End If
50060  If res = 1 Then
50070    CloseMutex = False
50080    hMutex = 0
50090   Else
50100    CloseMutex = True
50110  End If
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

