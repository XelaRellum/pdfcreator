Attribute VB_Name = "modPanelControl"
Option Explicit

' *** Den Artikel zu diesem Modul finden Sie unter http://www.aboutvb.de/khw/artikel/khwpanelcontrol.htm ***

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long

Private Const GWL_USERDATA = (-21)

Public Enum PanelControlErrorConstants
    pcErrInvalidPanelObject = 10001
    pcErrInvalidPanelIndexKey = 10002
End Enum

Public Function SetPanelControl(Control As Control, StatusBar As StatusBar, PanelKey As Variant, Optional ByVal AdjustStatusBarToControl As Boolean) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Dim nParent As Object
50020     Dim nControl As Control
50030     Dim nOldParentWnd As Long
50040     Dim nPanel As Panel
50050
50060     With Control
50070         If TypeOf .Parent Is MDIForm Then
50080             Set nParent = .Container
50090         Else
50100             Set nParent = .Parent
50110         End If
50120     End With
50130     If IsObject(PanelKey) Then
50140         If TypeOf PanelKey Is Panel Then
50150             Set nPanel = PanelKey
50160         End If
50170     End If
50180     If nPanel Is Nothing Then
50190         Set nPanel = StatusBar.Panels(PanelKey)
50200     End If
50210     With Control
50220         nOldParentWnd = SetParent(.hwnd, StatusBar.hwnd)
50230         If GetWindowLong(.hwnd, GWL_USERDATA) = 0 Then
50240             SetWindowLong .hwnd, GWL_USERDATA, nOldParentWnd
50250         Else
50260             SetPanelControl = nOldParentWnd
50270         End If
50280     End With
50290     zAdjust Control, nParent, nPanel, StatusBar, AdjustStatusBarToControl
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPanelControl", "SetPanelControl")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub RemovePanelControl(Control As Control, Optional OldParentWnd As Long)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     With Control
50020         .Visible = False
50030         If OldParentWnd Then
50040             SetParent .hwnd, OldParentWnd
50050         Else
50060             SetParent .hwnd, GetWindowLong(.hwnd, GWL_USERDATA)
50070         End If
50080     End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPanelControl", "RemovePanelControl")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub AdjustControlToPanel(Control As Control, StatusBar As StatusBar, PanelKey As Variant, Optional ByVal AdjustStatusBarToControl As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     Dim nParent As Object
50020     Dim nPanel As Panel
50030
50040     If IsObject(PanelKey) Then
50050         If TypeOf PanelKey Is Panel Then
50060             Set nPanel = PanelKey
50070         End If
50080     End If
50090     If nPanel Is Nothing Then
50100         Set nPanel = StatusBar.Panels(PanelKey)
50110     End If
50120     With Control
50130         If TypeOf .Parent Is MDIForm Then
50140             Set nParent = .Container
50150         Else
50160             Set nParent = .Parent
50170         End If
50180     End With
50190     zAdjust Control, nParent, nPanel, StatusBar, AdjustStatusBarToControl
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPanelControl", "AdjustControlToPanel")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub zAdjust(Control As Control, Parent As Object, Panel As Panel, StatusBar As StatusBar, ByVal AdjustStatusBarToControl As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010     On Error Resume Next
50020     With Parent
50030         If AdjustStatusBarToControl Then
50040             StatusBar.Height = Control.Height + .ScaleY(6, vbPixels)
50050             If Panel.Index = 1 Then
50060                 Control.Move Panel.Left + .ScaleX(1, vbPixels), .ScaleY(4, vbPixels), Panel.Width - .ScaleX(3, vbPixels)
50070             Else
50080                 Control.Move Panel.Left + .ScaleX(3, vbPixels), .ScaleY(4, vbPixels), Panel.Width - .ScaleX(4, vbPixels)
50090             End If
50100         Else
50110             If Panel.Index = 1 Then
50120                 Control.Move Panel.Left + .ScaleX(1, vbPixels), .ScaleY(4, vbPixels), Panel.Width - .ScaleX(3, vbPixels), StatusBar.Height - .ScaleY(6, vbPixels)
50130             Else
50140                 Control.Move Panel.Left + .ScaleX(3, vbPixels), .ScaleY(4, vbPixels), Panel.Width - .ScaleX(4, vbPixels), StatusBar.Height - .ScaleY(6, vbPixels)
50150             End If
50160         End If
50170     End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPanelControl", "zAdjust")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

