Attribute VB_Name = "modTimer"
Option Explicit

Public Declare Function SetTimer Lib "user32" _
 (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" _
 (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Public Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ' Create a mutex if possible
50020  Set mutex = New clsMutex
50030  If mutex.CheckMutex(PDFSpooler_GUID) = False Then
50040   mutex.CreateMutex PDFSpooler_GUID
50050  End If
50060  DoEvents
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modTimer", "TimerProc")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
