Attribute VB_Name = "modResize"
Option Explicit

Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub CopyMemory1 Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub CopyMemory2 Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Any, Source As Any, ByVal Length As Long)

Private Type POINTAPI
 x As Long
 y As Long
End Type

Private Type MINMAXINFO
 ptReserved As POINTAPI
 ptMaxSize As POINTAPI
 ptMaxPosition As POINTAPI
 ptMinTrackSize As POINTAPI
 ptMaxTrackSize As POINTAPI
End Type

Public Type SIZEPAR
 xMin As Long
 yMin As Long
 xMax As Long
 yMax As Long
End Type

Private Const GWL_WNDPROC = -4
Private Const WM_GETMINMAXINFO = &H24

Private WinOldProc&
Private spR As SIZEPAR
Private Frm As Form

Public Sub InitResize(F As Form, R As SIZEPAR)
 spR = R: Set Frm = F
 WinOldProc = SetWindowLong(Frm.hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub UnHookResize()
 Call SetWindowLong(Frm.hWnd, GWL_WNDPROC, WinOldProc)
End Sub

Private Function WindowProc(ByVal hWnd&, ByVal uMsg&, ByVal wParam&, ByVal lParam&) As Long
 Dim res As Long, MM As MINMAXINFO
 If uMsg = WM_GETMINMAXINFO And Frm.WindowState = 0 Then
   Call CopyMemory1(MM, lParam, Len(MM))
   MM.ptMaxPosition.x = 0
   MM.ptMaxPosition.y = 0
   MM.ptMaxSize.x = Screen.Width / Screen.TwipsPerPixelX
   MM.ptMaxSize.y = Screen.Height / Screen.TwipsPerPixelY
      
   MM.ptMinTrackSize.x = spR.xMin
   MM.ptMinTrackSize.y = spR.yMin
   MM.ptMaxTrackSize.x = spR.xMax
   MM.ptMaxTrackSize.y = spR.yMax
      
   Call CopyMemory2(lParam&, MM, Len(MM))
   res = DefWindowProc(hWnd, uMsg, wParam, lParam)
  Else
   res = CallWindowProc(WinOldProc, hWnd, uMsg, wParam, lParam)
 End If
 WindowProc = res
End Function
