Attribute VB_Name = "modTopMost"
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal Order As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal Flags As Long) As Long

Public Function SetTopMost(AnyForm As Form, Optional ByVal State As Boolean = True, Optional ByVal Activate As Boolean = True) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Dim nFlags As Long
50020     Dim nTopMode As Long
50030
50040     Const SWP_NOMOVE = 2
50050     Const SWP_NOSIZE = 1
50060     Const HWND_TOPMOST = -1
50070     Const HWND_NOTOPMOST = -2
50080     Const SWP_NOACTIVATE = &H10
50090
50100     nFlags = SWP_NOMOVE Or SWP_NOSIZE
50110     If Not Activate Then
50120         nFlags = nFlags Or SWP_NOACTIVATE
50130     End If
50141     Select Case State
              Case True
50160             nTopMode = HWND_TOPMOST
50170         Case False
50180             nTopMode = HWND_NOTOPMOST
50190     End Select
50200     SetTopMost = CBool(SetWindowPos(AnyForm.hwnd, nTopMode, 0, 0, 0, 0, nFlags))
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modTopMost", "SetTopMost")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

