Attribute VB_Name = "modTimer"
Option Explicit

Public Declare Function SetTimer Lib "user32" _
 (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" _
 (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Public Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
End Sub


