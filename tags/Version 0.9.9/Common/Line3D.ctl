VERSION 5.00
Begin VB.UserControl Line3D 
   AutoRedraw      =   -1  'True
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5430
   ScaleHeight     =   34
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   362
   ToolboxBitmap   =   "Line3D.ctx":0000
End
Attribute VB_Name = "Line3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/* Author: Morgan Haueisen (morganh@hartcom.net)
'/* Copyright (c) 2004
'/* Version 1.0.0

Option Explicit

Public Enum enuLineTypes
 [Flat Line] = 0
 [Inserted Line] = 1
 [Raised Line] = 2
End Enum

Private mudtLineType      As enuLineTypes
Private mlng3DHighlight   As OLE_COLOR
Private mlng3DShadow      As OLE_COLOR

Public Property Get DrawStyle() As DrawStyleConstants
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  DrawStyle = UserControl.DrawStyle
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("Line3D", "DrawStyle [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let DrawStyle(ByVal vNewValue As DrawStyleConstants)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  UserControl.DrawStyle = vNewValue
50020  PropertyChanged "DrawStyle"
50030  UserControl.Cls
50040  Call UserControl_Resize
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("Line3D", "DrawStyle [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get Border3DHighlight() As OLE_COLOR
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Border3DHighlight = mlng3DHighlight
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("Line3D", "Border3DHighlight [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let Border3DHighlight(ByVal vNewValue As OLE_COLOR)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  mlng3DHighlight = vNewValue
50020  PropertyChanged "Border3DHighlight"
50030  Call UserControl_Resize
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("Line3D", "Border3DHighlight [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get Border3DShadow() As OLE_COLOR
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Border3DShadow = mlng3DShadow
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("Line3D", "Border3DShadow [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let Border3DShadow(ByVal vNewValue As OLE_COLOR)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  mlng3DShadow = vNewValue
50020  PropertyChanged "Border3DShadow"
50030  Call UserControl_Resize
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("Line3D", "Border3DShadow [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get LineTypes() As enuLineTypes
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  LineTypes = mudtLineType
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("Line3D", "LineTypes [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let LineTypes(ByVal vNewValue As enuLineTypes)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  mudtLineType = vNewValue
50020  PropertyChanged "LineType"
50030  Call UserControl_Resize
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("Line3D", "LineTypes [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Private Sub UserControl_InitProperties()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  mudtLineType = 1
50020  mlng3DHighlight = vb3DHighlight
50030  mlng3DShadow = vb3DShadow
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("Line3D", "UserControl_InitProperties")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 On Error Resume Next
 With PropBag
  mudtLineType = .ReadProperty("LineType", 1)
  mlng3DHighlight = .ReadProperty("3DHighlight", vb3DHighlight)
  mlng3DShadow = .ReadProperty("3DShadow", vb3DShadow)
  UserControl.DrawStyle = .ReadProperty("DrawStyle", UserControl.DrawStyle)
 End With
End Sub

Private Sub UserControl_Resize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  With UserControl
50020   If .Width >= .Height Then
50030     '/* Horizontal Line
50041     Select Case mudtLineType
           Case 2 'Raised
50060       UserControl.Line (0, 0)-(.ScaleWidth, 0), mlng3DHighlight
50070       UserControl.Line (0, 1)-(.ScaleWidth, 1), mlng3DShadow
50080       .Height = 2 * Screen.TwipsPerPixelY
50090      Case 1 'Inserted
50100       UserControl.Line (0, 0)-(.ScaleWidth, 0), mlng3DShadow
50110       UserControl.Line (0, 1)-(.ScaleWidth, 1), mlng3DHighlight
50120       .Height = 2 * Screen.TwipsPerPixelY
50130      Case Else ' Flat
50140       UserControl.Line (0, 0)-(.ScaleWidth, 0), mlng3DShadow
50150       .Height = Screen.TwipsPerPixelY
50160     End Select
50170    Else
50180     '/* Vertical Line
50191     Select Case mudtLineType
           Case 2 'Raised
50210       UserControl.Line (0, 0)-(0, .ScaleHeight), mlng3DHighlight
50220       UserControl.Line (1, 0)-(1, .ScaleHeight), mlng3DShadow
50230       .Width = 2 * Screen.TwipsPerPixelX
50240      Case 1 'Inserted
50250       UserControl.Line (0, 0)-(0, .ScaleHeight), mlng3DShadow
50260       UserControl.Line (1, 0)-(1, .ScaleHeight), mlng3DHighlight
50270       .Width = 2 * Screen.TwipsPerPixelX
50280      Case Else 'Flat
50290       UserControl.Line (0, 0)-(0, .ScaleHeight), mlng3DShadow
50300       .Width = Screen.TwipsPerPixelX
50310     End Select
50320   End If
50330   .Refresh
50340  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("Line3D", "UserControl_Resize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 On Error Resume Next
  With PropBag
   .WriteProperty "LineType", mudtLineType
   .WriteProperty "3DHighlight", mlng3DHighlight
   .WriteProperty "3DShadow", mlng3DShadow
   .WriteProperty "DrawStyle", UserControl.DrawStyle
  End With
End Sub

