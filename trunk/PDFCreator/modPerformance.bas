Attribute VB_Name = "modPerformance"
Option Explicit

Private Declare Function QueryPerformanceFrequency Lib _
      "kernel32" (ByRef Frequency As Currency) As Long

Private Declare Function QueryPerformanceCounter Lib _
      "kernel32" (ByRef TimerValue As Currency) As Long

Private Const i_Frequency As Currency = 100

Public Sub Init_ExactTimer()
  QueryPerformanceFrequency i_Frequency
End Sub

Public Function ExactTimer_Value() As Currency
  QueryPerformanceCounter ExactTimer_Value
  ExactTimer_Value = ExactTimer_Value / i_Frequency
End Function
