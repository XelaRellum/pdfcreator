VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function SystemParametersInfo Lib _
        "user32" Alias "SystemParametersInfoA" (ByVal _
        uAction As Long, ByVal uParam As Long, _
        lpvParam As Any, ByVal fuWinIni As Long) As Long
               
Private Declare Function GetDeviceCaps Lib "gdi32" _
        (ByVal hDC As Long, ByVal nIndex As Long) As _
        Long

Private Const SPI_GETICONTITLELOGFONT = 31
Private Const SPI_GETNONCLIENTMETRICS = 41
Private Const LF_FACESIZE = 32
Private Const LOGPIXELSY = 90

Private Type NMLOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName(LF_FACESIZE - 4) As Byte
End Type

Private Type NONCLIENTMETRICS
  cbSize As Long
  iBorderWidth As Long
  iScrollWidth As Long
  iScrollHeight As Long
  iCaptionWidth As Long
  iCaptionHeight As Long
  lfCaptionFont As NMLOGFONT
  iSMCaptionWidth As Long
  iSMCaptionHeight As Long
  lfSMCaptionFont As NMLOGFONT
  iMenuWidth As Long
  iMenuHeight As Long
  lfMenuFont As NMLOGFONT
  lfStatusFont As NMLOGFONT
  lfMessageFont As NMLOGFONT
End Type

Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName(LF_FACESIZE) As Byte
End Type

Public Enum Systemfont
 Caption = 0
 SMCaption = 1
 Menu = 2
 Status = 3
 Message = 4
 Icon = 5
End Enum

Public Function GetSystemFont(Frm As Form, Font As Systemfont) As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim NCM As NONCLIENTMETRICS, ITLF As LOGFONT, res As Long, FName As String, _
 sFont(2) As String, FSize As Long, tNMLF As NMLOGFONT
50030
50040  Set GetSystemFont = New Collection
50050
50060  NCM.cbSize = 340
50070  res = SystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0, NCM, 0)
50080  If res <> 0 Then
50090   If Font = Icon Then
50100    Call SystemParametersInfo(SPI_GETICONTITLELOGFONT, 0, ITLF, 0)
50110    FName = StrConv(ITLF.lfFaceName, vbUnicode)
50120    FName = Left$(FName, InStr(1, FName, Chr$(0)) - 1)
50130    FSize = ITLF.lfHeight
50140    If FSize < 1 Then
50150     FSize = Abs((72 / GetDeviceCaps(Frm.hDC, LOGPIXELSY)) * FSize)
50160    End If
50170    sFont(0) = FName: sFont(1) = FSize: sFont(2) = ITLF.lfCharSet
50180    GetSystemFont.Add sFont
50190    Exit Function
50200   End If
50210   Select Case Font
   Case 0:
50230     tNMLF = NCM.lfCaptionFont
50240    Case 1:
50250     tNMLF = NCM.lfSMCaptionFont
50260    Case 2:
50270     tNMLF = NCM.lfMenuFont
50280    Case 3:
50290     tNMLF = NCM.lfStatusFont
50300    Case 4:
50310     tNMLF = NCM.lfMessageFont
50320   End Select
50330   FName = StrConv(tNMLF.lfFaceName, vbUnicode)
50340   FName = Left$(FName, InStr(1, FName, Chr$(0)) - 1)
50350   FSize = tNMLF.lfHeight
50360   If FSize < 1 Then
50370    FSize = Abs((72 / GetDeviceCaps(Frm.hDC, LOGPIXELSY)) * FSize)
50380   End If
50390   sFont(0) = FName: sFont(1) = FSize: sFont(2) = tNMLF.lfCharSet
50400   GetSystemFont.Add sFont
50410  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("clsSystem", "GetSystemFont")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

