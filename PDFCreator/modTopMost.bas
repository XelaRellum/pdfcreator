Attribute VB_Name = "modTopMost"
Option Explicit

Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal Order As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal Flags As Long) As Long

Public Function SetTopMost(AnyForm As Form, Optional ByVal State As Boolean = True, Optional ByVal Activate As Boolean = True) As Boolean
    Dim nFlags As Long
    Dim nTopMode As Long
    
    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    Const SWP_NOACTIVATE = &H10

    nFlags = SWP_NOMOVE Or SWP_NOSIZE
    If Not Activate Then
        nFlags = nFlags Or SWP_NOACTIVATE
    End If
    Select Case State
        Case True
            nTopMode = HWND_TOPMOST
        Case False
            nTopMode = HWND_NOTOPMOST
    End Select
    SetTopMost = CBool(SetWindowPos(AnyForm.hWnd, nTopMode, 0, 0, 0, 0, nFlags))
End Function

