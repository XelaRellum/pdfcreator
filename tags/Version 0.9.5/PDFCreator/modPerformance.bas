Attribute VB_Name = "modPerformance"
Option Explicit

Public PerformanceTimer As Boolean
Private i_Frequency As Currency

Public Sub Init_ExactTimer()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010   Dim res As Long
50020   res = QueryPerformanceFrequency(i_Frequency)
50030   If res = 0 Then
50040     PerformanceTimer = False
50050    Else
50060     PerformanceTimer = True
50070   End If
50080   If i_Frequency = 0 Then
50090    i_Frequency = 1
50100   End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPerformance", "Init_ExactTimer")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function ExactTimer_Value() As Currency
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010   QueryPerformanceCounter ExactTimer_Value
50020   ExactTimer_Value = ExactTimer_Value / i_Frequency
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPerformance", "ExactTimer_Value")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
