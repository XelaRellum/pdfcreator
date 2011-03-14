Attribute VB_Name = "modTaskbar"
Option Explicit

Public Enum tBarAlign
 tbaLeft = 0
 tbaRight = 1
 tbaTop = 2
 tbaBottom = 3
End Enum

Public Function IsTaskBarOnTop() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim AppBar As APPBARDATA, State As Long
50020  AppBar.cbSize = Len(AppBar)
50030  State = SHAppBarMessage(ABM_GETSTATE, AppBar)
50040  IsTaskBarOnTop = (State = ABS_ALWAYSONTOP)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modTaskbar", "IsTaskBarOnTop")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function TaskBarAlign() As tBarAlign
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim AppBar As APPBARDATA, lResult As Long, tbHeight As Long, tbWidth As Long
50020  lResult = SHAppBarMessage(ABM_GETTASKBARPOS, AppBar)
50030  With AppBar.rc
50040   tbHeight = .Bottom - .Top
50050   tbWidth = .Right - .Left
50060  End With
50070
50080  If tbHeight < Screen.Height / Screen.TwipsPerPixelY Then
50090    If AppBar.rc.Top <= 0 Then
50100      TaskBarAlign = tbaTop
50110     Else
50120      TaskBarAlign = tbaBottom
50130     End If
50140   Else
50150    If AppBar.rc.Left <= 0 Then
50160      TaskBarAlign = tbaLeft
50170     Else
50180      TaskBarAlign = tbaRight
50190     End If
50200   End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modTaskbar", "TaskBarAlign")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function TaskBarHeight() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim AppBar As APPBARDATA, lResult As Long
50020
50030  lResult = SHAppBarMessage(ABM_GETTASKBARPOS, AppBar)
50040  With AppBar.rc
50050   TaskBarHeight = .Bottom - .Top
50060  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modTaskbar", "TaskBarHeight")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function TaskBarWidth() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim AppBar As APPBARDATA, lResult As Long
50020  lResult = SHAppBarMessage(ABM_GETTASKBARPOS, AppBar)
50030  With AppBar.rc
50040   TaskBarWidth = .Right - .Left
50050  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modTaskbar", "TaskBarWidth")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

