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

Private Declare Function CreateMutexA Lib "Kernel32" _
 (ByVal lpMutexAttributes As Long, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function CloseHandle Lib "Kernel32" _
 (ByVal hObject As Long) As Long

Private Const ERROR_ALREADY_EXISTS = 183&

Private hMutex As Long

Public Function CheckMutex(ByVal Mutexname As String) As Boolean
 If IDERunning = True Then
   CheckMutex = App.PrevInstance
  Else
   hMutex = CreateMutexA(ByVal 0&, 1, Mutexname)
   If (Err.LastDllError = ERROR_ALREADY_EXISTS) Then
     CheckMutex = True
    Else
     CheckMutex = False
   End If
   If (hMutex <> 0) Then
    CloseHandle hMutex
   End If
 End If
End Function

Public Function CreateMutex(ByVal Mutexname As String) As Long
 CreateMutex = 0
 hMutex = CreateMutexA(ByVal 0&, 1, Mutexname)
 If (Err.LastDllError = ERROR_ALREADY_EXISTS) Then
   CreateMutex = 0
  Else
   CreateMutex = hMutex
 End If
End Function

Private Function IDERunning() As Boolean
 On Error GoTo ErrorHandler
 Debug.Print 1 / 0
 Exit Function
ErrorHandler:
 IDERunning = True
End Function

Public Function CloseMutex() As Long
 Dim res As Long
 If (hMutex <> 0) Then
  res = CloseHandle(hMutex)
 End If
 If res = 0 Then
   CloseMutex = 0
   hMutex = 0
  Else
   CloseMutex = res
 End If
End Function

