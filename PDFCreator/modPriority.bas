Attribute VB_Name = "modPriority"
Option Explicit

Const THREAD_BASE_PRIORITY_IDLE = -15
Const THREAD_BASE_PRIORITY_LOWRT = 15
Const THREAD_BASE_PRIORITY_MIN = -2
Const THREAD_BASE_PRIORITY_MAX = 2
Const THREAD_PRIORITY_LOWEST = THREAD_BASE_PRIORITY_MIN
Const THREAD_PRIORITY_HIGHEST = THREAD_BASE_PRIORITY_MAX
Const THREAD_PRIORITY_BELOW_NORMAL = (THREAD_PRIORITY_LOWEST + 1)
Const THREAD_PRIORITY_ABOVE_NORMAL = (THREAD_PRIORITY_HIGHEST - 1)
Const THREAD_PRIORITY_IDLE = THREAD_BASE_PRIORITY_IDLE
Const THREAD_PRIORITY_NORMAL = 0
Const THREAD_PRIORITY_TIME_CRITICAL = THREAD_BASE_PRIORITY_LOWRT
Const HIGH_PRIORITY_CLASS = &H80
Const IDLE_PRIORITY_CLASS = &H40
Const NORMAL_PRIORITY_CLASS = &H20
Const REALTIME_PRIORITY_CLASS = &H100
Private Declare Function SetThreadPriority Lib "Kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Private Declare Function SetPriorityClass Lib "Kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare Function GetThreadPriority Lib "Kernel32" (ByVal hThread As Long) As Long
Private Declare Function GetPriorityClass Lib "Kernel32" (ByVal hProcess As Long) As Long
Private Declare Function GetCurrentThread Lib "Kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "Kernel32" () As Long

Public Enum tProcessPriority
 RealTime = REALTIME_PRIORITY_CLASS
 High = HIGH_PRIORITY_CLASS
 Normal = NORMAL_PRIORITY_CLASS
 Idle = IDLE_PRIORITY_CLASS
End Enum

Public Function GetProcessPriority() As Long
 If IsWin9xMe = False Then
  GetPriorityClass GetCurrentProcess
 End If
End Function

Public Sub SetProcessPriority(ProccessPriority As tProcessPriority)
 If IsWin9xMe = False Then
  SetPriorityClass GetCurrentProcess, ProccessPriority
 End If
End Sub
