VERSION 5.00
Begin VB.UserControl dmFrame 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000B&
   ClientHeight    =   1365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2010
   ControlContainer=   -1  'True
   FillColor       =   &H8000000F&
   ForeColor       =   &H00FFFFFF&
   ForwardFocus    =   -1  'True
   ScaleHeight     =   91
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   134
   ToolboxBitmap   =   "dmFrame.ctx":0000
End
Attribute VB_Name = "dmFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Based on the source code from dreamvb
' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=58966&lngWId=1
'
' Added some new features
' BarColorFrom, BarColorTo, Caption3D, TextShaddowColor, BarHeight, Enabled

Private Const InitBarColorFrom = &HFFCCBF
Private Const InitBarColorTo = &H823326

Private Const m_def_Caption As String = "CoolXPFrame"
Private Const m_def_Caption3D As String = 0

Public Enum eCaption3D
 [Flat Caption] = 0
 [Inserted Caption] = 1
 [Raised Caption] = 2
End Enum


'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private m_Caption As String
Private m_Caption3D As eCaption3D
Private n_BarHeight As Long
Private m_GradEn As Boolean
Private m_BarColorFrom As OLE_COLOR
Private m_BarColorTo As OLE_COLOR
Private m_OutLineColor As OLE_COLOR
Private m_OutLineStyle As DrawStyleConstants
Private m_Alignment As dmAlignment
Private m_UseTextShaddow As Boolean
Private m_TextShaddowColor As OLE_COLOR
Private m_Relocatable As Boolean

Private OldY As Single, OnBar As Boolean

Private Type dmRgb
 Red As Long
 Green As Long
 Blue As Long
End Type

Enum dmAlignment
 dmLeft = 0
 dmCenter
 dmRight
End Enum

Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event BarMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event Click()
Event BarClick()
Event DblClick()
Event BarDblClick()

Private Function LongToRGB(nLongVal As Long) As dmRgb
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim b1 As Byte, b2 As Byte, b3 As Byte, b4 As Byte
50020 ' b1 = (nLongVal And &HFF000000) \ &H1000000 And &HFF
50030  b2 = (nLongVal And &HFF0000) \ &H10000
50040  b3 = (nLongVal And &HFF00&) \ &H100
50050  b4 = nLongVal And &HFF&
50060  LongToRGB.Red = b4: LongToRGB.Green = b3: LongToRGB.Blue = b2
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "LongToRGB")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function SysColor(ByVal Color As Long) As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If Color And &H80000000 Then
50020    SysColor = GetSysColor(Color And &H7FFFFFFF)
50030   Else
50040    SysColor = Color
50050  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "SysColor")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function RGBToGray(RGBValue As dmRgb) As dmRgb
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim gray As Long
50020  With RGBValue
50030   gray = 0.299 * .Red + 0.587 * .Green + 0.114 * .Blue
50040  End With
50050  With RGBToGray
50060   .Red = gray
50070   .Green = gray
50080   .Blue = gray
50090  End With
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "RGBToGray")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Sub DrawdmFrame()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Integer, dm_RgbFrom As dmRgb, dm_RgbTo As dmRgb, tOC As OLE_COLOR, _
  tX As Long, tY As Long, tD As Long
50030
50040  UserControl.Cls
50050  n_BarHeight = UserControl.TextHeight("Xz") + 1
50060  dm_RgbFrom = LongToRGB(m_BarColorFrom)
50070  dm_RgbTo = LongToRGB(m_BarColorTo)
50080
50090  If UserControl.Enabled = False Then
50100   dm_RgbFrom = LongToRGB(SysColor(vbButtonFace))
50110   dm_RgbTo = LongToRGB(SysColor(vbButtonFace))
50120  End If
50130
50140  If m_GradEn Then
50150    For i = 0 To n_BarHeight
50160     UserControl.Line (0, i + (0 * n_BarHeight))-(UserControl.ScaleWidth, i + (0 * _
     n_BarHeight)), _
     RGB( _
     dm_RgbFrom.Red + (dm_RgbTo.Red - dm_RgbFrom.Red) / n_BarHeight * i, _
     dm_RgbFrom.Green + (dm_RgbTo.Green - dm_RgbFrom.Green) / n_BarHeight * i, _
     dm_RgbFrom.Blue + (dm_RgbTo.Blue - dm_RgbFrom.Blue) / n_BarHeight * i)
50220    Next i
50230   Else
50240    UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, n_BarHeight), RGB(dm_RgbFrom.Red, dm_RgbFrom.Green, dm_RgbFrom.Blue), BF
50250  End If
50260
50270  If Enabled = False Then
50280    UserControl.Line (0, n_BarHeight + 1)-(UserControl.ScaleWidth - 1, n_BarHeight + 1), SysColor(vbButtonShadow), BF
50290  End If
50300
50310  UserControl.CurrentY = 1.3
50320
50331  Select Case m_Alignment
        Case dmLeft
50350    UserControl.CurrentX = 3
50360   Case dmCenter
50370    UserControl.CurrentX = (UserControl.ScaleWidth - TextWidth(m_Caption) + 3) / 2
50380   Case dmRight
50390    UserControl.CurrentX = (UserControl.ScaleWidth - TextWidth(m_Caption) - 3)
50400  End Select
50410
50421  Select Case m_Caption3D
        Case 1
50440    tOC = UserControl.ForeColor
50450    If Enabled = True Then
50460      UserControl.ForeColor = m_TextShaddowColor
50470     Else
50480      UserControl.ForeColor = vbWhite
50490    End If
50500    tX = UserControl.CurrentX
50510    tY = UserControl.CurrentY
50520    tD = CDbl(UserControl.Fontsize) / 8#
50530    UserControl.CurrentY = tY - IIf(Int(tD) = tD, tD, Int(tD) + 1)
50540    UserControl.CurrentX = tX - IIf(Int(tD) = tD, tD, Int(tD) + 1)
50550    UserControl.Print m_Caption
50560    UserControl.ForeColor = tOC
50570    UserControl.CurrentY = tY
50580    UserControl.CurrentX = tX
50590   Case 2
50600    tOC = UserControl.ForeColor
50610    If Enabled = True Then
50620      UserControl.ForeColor = m_TextShaddowColor
50630     Else
50640      UserControl.ForeColor = vbWhite
50650    End If
50660    tX = UserControl.CurrentX
50670    tY = UserControl.CurrentY
50680    tD = CDbl(UserControl.Fontsize) / 8#
50690    UserControl.CurrentY = tY + IIf(Int(tD) = tD, tD, Int(tD) + 1)
50700    UserControl.CurrentX = tX + IIf(Int(tD) = tD, tD, Int(tD) + 1)
50710    UserControl.Print m_Caption
50720    UserControl.ForeColor = tOC
50730    UserControl.CurrentY = tY
50740    UserControl.CurrentX = tX
50750  End Select
50760
50770
50780  If UserControl.Enabled = False Then
50790    tOC = UserControl.ForeColor
50800    UserControl.ForeColor = vbButtonShadow
50810    UserControl.Print m_Caption
50820    UserControl.ForeColor = tOC
50830   Else
50840    UserControl.Print m_Caption
50850  End If
50860
50870  UserControl.DrawStyle = m_OutLineStyle
50880  If UserControl.Enabled = True Then
50890    UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), m_OutLineColor, B
50900   Else
50910    UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), SysColor(vbButtonShadow), B
50920  End If
50930  UserControl.DrawStyle = vbSolid
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "DrawdmFrame")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Property Get Alignment() As dmAlignment
Attribute Alignment.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Alignment = m_Alignment
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "Alignment [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let Alignment(ByVal New_Alignment As dmAlignment)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  m_Alignment = New_Alignment
50020  PropertyChanged "Alignment"
50030  DrawdmFrame
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "Alignment [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get Caption() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Caption = m_Caption
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "Caption [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let Caption(ByVal New_Caption As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  m_Caption = New_Caption
50020  PropertyChanged "Caption"
50030  DrawdmFrame
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "Caption [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get Caption3D() As eCaption3D
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Caption3D = m_Caption3D
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "Caption3D [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let Caption3D(ByVal newValue As eCaption3D)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  m_Caption3D = newValue
50020  PropertyChanged "Caption3D"
50030  DrawdmFrame
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "Caption3D [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Private Sub UserControl_Initialize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  m_Caption = m_def_Caption
50020  m_Alignment = dmLeft
50030  m_OutLineStyle = vbSolid
50040  m_OutLineColor = &H80000010
50050  m_BarColorFrom = InitBarColorFrom
50060  m_BarColorTo = InitBarColorTo
50070  m_GradEn = True
50080  m_UseTextShaddow = True
50090  m_TextShaddowColor = RGB(32, 32, 32)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "UserControl_Initialize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 ' RaiseEvent MouseMove(Button, Shift, x, Y)
50020 ' If Button <> vbLeftButton Then
50030 '  Exit Sub
50040 ' End If
50050 ' If Not (Y > (n_BarHeight - 1)) And m_Relocatable = True Then
50060 '  Call ReleaseCapture
50070 '  Call SendMessage(UserControl.hwnd, &HA1, 2, 0&)
50080 ' End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "UserControl_MouseMove")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
50020  m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
50030  m_Caption3D = PropBag.ReadProperty("Caption3D", m_def_Caption3D)
50040  UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
50050  UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &HFFFFFF)
50060  m_OutLineColor = PropBag.ReadProperty("OutLineColor", &H80000010)
50070  m_Alignment = PropBag.ReadProperty("Alignment", 0)
50080  m_BarColorFrom = PropBag.ReadProperty("BarColorFrom", InitBarColorFrom)
50090  m_BarColorTo = PropBag.ReadProperty("BarColorTo", InitBarColorTo)
50100  m_GradEn = PropBag.ReadProperty("UseGradient", True)
50110  Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
50120  m_OutLineStyle = PropBag.ReadProperty("OutLineStyle", 0)
50130  Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
50140  UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
50150  m_TextShaddowColor = PropBag.ReadProperty("TextShaddowColor", RGB(32, 32, 32))
50160  UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
50170  m_Relocatable = PropBag.ReadProperty("Relocatable", False)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "UserControl_ReadProperties")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub UserControl_Resize()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  DrawdmFrame
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "UserControl_Resize")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub UserControl_Show()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  DrawdmFrame
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "UserControl_Show")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
50020  Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
50030  Call PropBag.WriteProperty("Caption3D", m_Caption3D, m_def_Caption3D)
50040  Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
50050  Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &HFFFFFF)
50060  Call PropBag.WriteProperty("OutLineColor", m_OutLineColor, &H80000010)
50070  Call PropBag.WriteProperty("BarColorFrom", m_BarColorFrom, InitBarColorFrom)
50080  Call PropBag.WriteProperty("BarColorTo", m_BarColorTo, InitBarColorTo)
50090  Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
50100  Call PropBag.WriteProperty("UseGradient", m_GradEn, True)
50110  Call PropBag.WriteProperty("OutLineStyle", m_OutLineStyle, 0)
50120  Call PropBag.WriteProperty("Alignment", m_Alignment, 0)
50130  Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
50140  Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
50150  Call PropBag.WriteProperty("TextShaddowColor", m_TextShaddowColor, RGB(32, 32, 32))
50160  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
50170  Call PropBag.WriteProperty("Relocatable", m_Relocatable, False)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "UserControl_WriteProperties")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Property Get BarHeight() As Long
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  BarHeight = n_BarHeight
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "BarHeight [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  BackColor = UserControl.BackColor
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "BackColor [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  UserControl.BackColor() = New_BackColor
50020  PropertyChanged "BackColor"
50030  DrawdmFrame
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "BackColor [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  ForeColor = UserControl.ForeColor
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "ForeColor [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let OutLineColor(ByVal New_OutlineColor As OLE_COLOR)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  m_OutLineColor = New_OutlineColor
50020  PropertyChanged "OutLineColor"
50030  DrawdmFrame
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "OutLineColor [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get OutLineColor() As OLE_COLOR
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  OutLineColor = m_OutLineColor
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "OutLineColor [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  UserControl.ForeColor() = New_ForeColor
50020  PropertyChanged "ForeColor"
50030  DrawdmFrame
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "ForeColor [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get BarColorFrom() As OLE_COLOR
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  BarColorFrom = m_BarColorFrom
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "BarColorFrom [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let BarColorFrom(ByVal New_Color As OLE_COLOR)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  m_BarColorFrom = New_Color
50020  PropertyChanged "BarColorFrom"
50030  DrawdmFrame
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "BarColorFrom [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get BarColorTo() As OLE_COLOR
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  BarColorTo = m_BarColorTo
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "BarColorTo [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let BarColorTo(ByVal New_Color As OLE_COLOR)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  m_BarColorTo = New_Color
50020  PropertyChanged "BarColorTo"
50030  DrawdmFrame
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "BarColorTo [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get TextShaddowColor() As OLE_COLOR
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  TextShaddowColor = m_TextShaddowColor
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "TextShaddowColor [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let TextShaddowColor(ByVal New_Color As OLE_COLOR)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  m_TextShaddowColor = New_Color
50020  PropertyChanged "TextShaddowColor"
50030  DrawdmFrame
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "TextShaddowColor [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let Relocatable(ByVal vNewValue As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  m_Relocatable = vNewValue
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "Relocatable [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get Relocatable() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Relocatable = m_Relocatable
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "Relocatable [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property
 
Public Property Get UseGradient() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  UseGradient = m_GradEn
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "UseGradient [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let UseGradient(ByVal vNewValue As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  m_GradEn = vNewValue
50020  DrawdmFrame
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "UseGradient [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Set Font = UserControl.Font
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "Font [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Set Font(ByVal New_Font As StdFont)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Set UserControl.Font = New_Font
50020  PropertyChanged "Font"
50030  DrawdmFrame
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "Font [SET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get Enabled() As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Enabled = UserControl.Enabled
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "Enabled [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ctl As Control
50020  UserControl.Enabled = New_Enabled
50030  PropertyChanged "Enabled"
50040  For Each ctl In Controls
50050   ctl.Enabled = New_Enabled
50060  Next ctl
50070  DrawdmFrame
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "Enabled [LET]")
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
50010  Set UserControl.Font = Ambient.Font
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "UserControl_InitProperties")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Property Get OutLineStyle() As DrawStyleConstants
Attribute OutLineStyle.VB_Description = "Determines the line style for output from graphics methods."
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  OutLineStyle = m_OutLineStyle
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "OutLineStyle [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let OutLineStyle(ByVal New_OutLineStyle As DrawStyleConstants)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  m_OutLineStyle = New_OutLineStyle
50020  PropertyChanged "OutLineStyle"
50030  DrawdmFrame
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "OutLineStyle [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  RaiseEvent MouseDown(Button, Shift, x, Y)
50020  OnBar = False
50030  If Not (Y > (n_BarHeight - 1)) Then
50040   OldY = Y: OnBar = True
50050   RaiseEvent BarMouseDown(Button, Shift, x, Y)
50060  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "UserControl_MouseDown")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Set MouseIcon = UserControl.MouseIcon
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "MouseIcon [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Set UserControl.MouseIcon = New_MouseIcon
50020  PropertyChanged "MouseIcon"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "MouseIcon [SET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  MousePointer = UserControl.MousePointer
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "MousePointer [GET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  UserControl.MousePointer() = New_MousePointer
50020  PropertyChanged "MousePointer"
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Property
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "MousePointer [LET]")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Property
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Property

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  RaiseEvent MouseUp(Button, Shift, x, Y)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "UserControl_MouseUp")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub UserControl_Click()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If Not OnBar Then
50020   RaiseEvent Click
50030  End If
50040  If Not (Y > (n_BarHeight - 1)) And OnBar Then
50050   RaiseEvent BarClick
50060  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "UserControl_Click")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub UserControl_DblClick()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If Not OnBar Then
50020   RaiseEvent DblClick
50030  End If
50040  If Not (Y > (n_BarHeight - 1)) And OnBar Then
50050   RaiseEvent BarDblClick
50060  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("dmFrame", "UserControl_DblClick")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub
