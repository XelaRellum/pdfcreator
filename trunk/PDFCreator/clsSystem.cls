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

Public Function GetSystemFont(frm As Form, Font As Systemfont) As Collection
 Dim NCM As NONCLIENTMETRICS, ITLF As LOGFONT, res As Long, FName As String, _
 sFont(2) As String, FSize As Long, tNMLF As NMLOGFONT
 
 Set GetSystemFont = New Collection
 
 NCM.cbSize = 340
 res = SystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0, NCM, 0)
 If res <> 0 Then
  If Font = Icon Then
   Call SystemParametersInfo(SPI_GETICONTITLELOGFONT, 0, ITLF, 0)
   FName = StrConv(ITLF.lfFaceName, vbUnicode)
   FName = Left$(FName, InStr(1, FName, Chr$(0)) - 1)
   FSize = ITLF.lfHeight
   If FSize < 1 Then
    FSize = Abs((72 / GetDeviceCaps(frm.hDC, LOGPIXELSY)) * FSize)
   End If
   sFont(0) = FName: sFont(1) = FSize: sFont(2) = ITLF.lfCharSet
   GetSystemFont.Add sFont
   Exit Function
  End If
  Select Case Font
   Case 0:
    tNMLF = NCM.lfCaptionFont
   Case 1:
    tNMLF = NCM.lfSMCaptionFont
   Case 2:
    tNMLF = NCM.lfMenuFont
   Case 3:
    tNMLF = NCM.lfStatusFont
   Case 4:
    tNMLF = NCM.lfMessageFont
  End Select
  FName = StrConv(tNMLF.lfFaceName, vbUnicode)
  FName = Left$(FName, InStr(1, FName, Chr$(0)) - 1)
  FSize = tNMLF.lfHeight
  If FSize < 1 Then
   FSize = Abs((72 / GetDeviceCaps(frm.hDC, LOGPIXELSY)) * FSize)
  End If
  sFont(0) = FName: sFont(1) = FSize: sFont(2) = tNMLF.lfCharSet
  GetSystemFont.Add sFont
 End If
End Function

