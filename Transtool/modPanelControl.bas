Attribute VB_Name = "modPanelControl"
Option Explicit

' *** Den Artikel zu diesem Modul finden Sie unter http://www.aboutvb.de/khw/artikel/khwpanelcontrol.htm ***

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long

Private Const GWL_USERDATA = (-21)

Public Enum PanelControlErrorConstants
    pcErrInvalidPanelObject = 10001
    pcErrInvalidPanelIndexKey = 10002
End Enum

Public Function SetPanelControl(Control As Control, StatusBar As StatusBar, PanelKey As Variant, Optional ByVal AdjustStatusBarToControl As Boolean) As Long
    Dim nParent As Object
    Dim nControl As Control
    Dim nOldParentWnd As Long
    Dim nPanel As Panel

    With Control
        If TypeOf .Parent Is MDIForm Then
            Set nParent = .Container
        Else
            Set nParent = .Parent
        End If
    End With
    If IsObject(PanelKey) Then
        If TypeOf PanelKey Is Panel Then
            Set nPanel = PanelKey
        End If
    End If
    If nPanel Is Nothing Then
        Set nPanel = StatusBar.Panels(PanelKey)
    End If
    With Control
        nOldParentWnd = SetParent(.hWnd, StatusBar.hWnd)
        If GetWindowLong(.hWnd, GWL_USERDATA) = 0 Then
            SetWindowLong .hWnd, GWL_USERDATA, nOldParentWnd
        Else
            SetPanelControl = nOldParentWnd
        End If
    End With
    zAdjust Control, nParent, nPanel, StatusBar, AdjustStatusBarToControl
End Function

Public Sub RemovePanelControl(Control As Control, Optional OldParentWnd As Long)
    With Control
        .Visible = False
        If OldParentWnd Then
            SetParent .hWnd, OldParentWnd
        Else
            SetParent .hWnd, GetWindowLong(.hWnd, GWL_USERDATA)
        End If
    End With
End Sub

Public Sub AdjustControlToPanel(Control As Control, StatusBar As StatusBar, PanelKey As Variant, Optional ByVal AdjustStatusBarToControl As Boolean)
    Dim nParent As Object
    Dim nPanel As Panel

    If IsObject(PanelKey) Then
        If TypeOf PanelKey Is Panel Then
            Set nPanel = PanelKey
        End If
    End If
    If nPanel Is Nothing Then
        Set nPanel = StatusBar.Panels(PanelKey)
    End If
    With Control
        If TypeOf .Parent Is MDIForm Then
            Set nParent = .Container
        Else
            Set nParent = .Parent
        End If
    End With
    zAdjust Control, nParent, nPanel, StatusBar, AdjustStatusBarToControl
End Sub

Private Sub zAdjust(Control As Control, Parent As Object, Panel As Panel, StatusBar As StatusBar, ByVal AdjustStatusBarToControl As Boolean)
    On Error Resume Next
    With Parent
        If AdjustStatusBarToControl Then
            StatusBar.Height = Control.Height + .ScaleY(6, vbPixels)
            If Panel.Index = 1 Then
                Control.Move Panel.Left + .ScaleX(1, vbPixels), .ScaleY(4, vbPixels), Panel.Width - .ScaleX(3, vbPixels)
            Else
                Control.Move Panel.Left + .ScaleX(3, vbPixels), .ScaleY(4, vbPixels), Panel.Width - .ScaleX(4, vbPixels)
            End If
        Else
            If Panel.Index = 1 Then
                Control.Move Panel.Left + .ScaleX(1, vbPixels), .ScaleY(4, vbPixels), Panel.Width - .ScaleX(3, vbPixels), StatusBar.Height - .ScaleY(6, vbPixels)
            Else
                Control.Move Panel.Left + .ScaleX(3, vbPixels), .ScaleY(4, vbPixels), Panel.Width - .ScaleX(4, vbPixels), StatusBar.Height - .ScaleY(6, vbPixels)
            End If
        End If
    End With
End Sub

