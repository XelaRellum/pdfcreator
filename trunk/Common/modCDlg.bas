Attribute VB_Name = "modCDlg"
Option Explicit

Const FW_NORMAL = 400
Const DEFAULT_CHARSET = 1
Const OUT_DEFAULT_PRECIS = 0
Const CLIP_DEFAULT_PRECIS = 0
Const DEFAULT_QUALITY = 0
Const DEFAULT_PITCH = 0
Const FF_ROMAN = 16
Const CF_PRINTERFONTS = &H2
Const CF_SCREENFONTS = &H1
Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Const CF_EFFECTS = &H100&
Const CF_FORCEFONTEXIST = &H10000
Const CF_INITTOLOGFONTSTRUCT = &H40&
Const CF_LIMITSIZE = &H2000&
Const REGULAR_FONTTYPE = &H400
Const LF_FACESIZE = 32
Const CCHDEVICENAME = 32
Const CCHFORMNAME = 32
Const GMEM_MOVEABLE = &H2
Const GMEM_ZEROINIT = &H40
Const DM_DUPLEX = &H1000&
Const DM_ORIENTATION = &H1&
Const PD_PRINTSETUP = &H40
Const PD_DISABLEPRINTTOFILE = &H80000

Public Enum OpenSaveFlags
 OFN_ALLOWMULTISELECT = &H200
 OFN_CREATEPROMPT = &H2000
 OFN_ENABLEHOOK = &H20
 OFN_ENABLETEMPLATE = &H40
 OFN_ENABLETEMPLATEHANDLE = &H80
 OFN_EXPLORER = &H80000
 OFN_EXTENSIONDIFFERENT = &H400
 OFN_FILEMUSTEXIST = &H1000
 OFN_HIDEREADONLY = &H4
 OFN_LONGNAMES = &H200000
 OFN_NOCHANGEDIR = &H8
 OFN_NODEREFERENCELINKS = &H100000
 OFN_NOLONGNAMES = &H40000
 OFN_NONETWORKBUTTON = &H20000
 OFN_NOREADONLYRETURN = &H8000&
 OFN_NOTESTFILECREATE = &H10000
 OFN_NOVALIDATE = &H100
 OFN_OVERWRITEPROMPT = &H2
 OFN_PATHMUSTEXIST = &H800
 OFN_READONLY = &H1
 OFN_SHAREAWARE = &H4000
 OFN_SHAREFALLTHROUGH = 2
 OFN_SHAREWARN = 0
 OFN_SHARENOWARN = 1
 OFN_SHOWHELP = &H10
 OFS_MAXPATHNAME = 260
End Enum

Public Type POINTAPI
 X As Long
 Y As Long
End Type

Public Type RECT
 Left As Long
 Top As Long
 Right As Long
 Bottom As Long
End Type

Public Type OPENFILENAME
 nStructSize    As Long
 hWndOwner      As Long
 hInstance      As Long
 sFilter        As String
 sCustomFilter  As String
 nMaxCustFilter As Long
 nFilterIndex   As Long
 sFile          As String
 nMaxFile       As Long
 sFileTitle     As String
 nMaxTitle      As Long
 sInitialDir    As String
 sDialogTitle   As String
 Flags          As Long
 nFileOffset    As Integer
 nFileExtension As Integer
 sDefFileExt    As String
 nCustData      As Long
 fnHook         As Long
 sTemplateName  As String
End Type

Public Type PAGESETUPDLG
 lStructSize As Long
 hWndOwner As Long
 hDevMode As Long
 hDevNames As Long
 Flags As Long
 ptPaperSize As POINTAPI
 rtMinMargin As RECT
 rtMargin As RECT
 hInstance As Long
 lCustData As Long
 lpfnPageSetupHook As Long
 lpfnPagePaintHook As Long
 lpPageSetupTemplateName As String
 hPageSetupTemplate As Long
End Type

Public Type CHOOSECOLOR
 lStructSize As Long
 hWndOwner As Long
 hInstance As Long
 rgbResult As Long
 lpCustColors As String
 Flags As Long
 lCustData As Long
 lpfnHook As Long
 lpTemplateName As String
End Type

Public Type LOGFONT
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
 lfFaceName As String * 31
End Type

Public Type CHOOSEFONT
 lStructSize As Long
 hWndOwner As Long ' caller's window handle
 hDC As Long ' printer DC/IC or NULL
 lpLogFont As Long ' ptr. to a LOGFONT struct
 iPointSize As Long ' 10 * size in points of Selected Font
 Flags As Long ' enum. type flags
 rgbColors As Long ' returned text color
 lCustData As Long ' data passed to hook fn.
 lpfnHook As Long ' ptr. to hook function
 lpTemplateName As String ' custom template name
 hInstance As Long ' instance handle of.EXE that
 ' contains cust. dlg. template
 lpszStyle As String ' return the style field here
 ' must be LF_FACESIZE or bigger
 nFontType As Integer ' same value reported to the EnumFonts
 ' call back with the extra FONTTYPE_
 ' bits added
 MISSING_ALIGNMENT As Integer
 nSizeMin As Long ' minimum pt size allowed&
 nSizeMax As Long ' max pt size allowed if
 ' CF_LIMITSIZE is used
End Type

Public Type PRINTDLG_TYPE
 lStructSize As Long
 hWndOwner As Long
 hDevMode As Long
 hDevNames As Long
 hDC As Long
 Flags As Long
 nFromPage As Integer
 nToPage As Integer
 nMinPage As Integer
 nMaxPage As Integer
 nCopies As Integer
 hInstance As Long
 lCustData As Long
 lpfnPrintHook As Long
 lpfnSetupHook As Long
 lpPrintTemplateName As String
 lpSetupTemplateName As String
 hPrintTemplate As Long
 hSetupTemplate As Long
End Type

Public Type DEVNAMES_TYPE
 wDriverOffset As Integer
 wDeviceOffset As Integer
 wOutputOffset As Integer
 wDefault As Integer
 extra As String * 100
End Type

Public Type DEVMODE_TYPE
 dmDeviceName As String * CCHDEVICENAME
 dmSpecVersion As Integer
 dmDriverVersion As Integer
 dmSize As Integer
 dmDriverExtra As Integer
 dmFields As Long
 dmOrientation As Integer
 dmPaperSize As Integer
 dmPaperLength As Integer
 dmPaperWidth As Integer
 dmScale As Integer
 dmCopies As Integer
 dmDefaultSource As Integer
 dmPrintQuality As Integer
 dmColor As Integer
 dmDuplex As Integer
 dmYResolution As Integer
 dmTTOption As Integer
 dmCollate As Integer
 dmFormName As String * CCHFORMNAME
 dmUnusedPadding As Integer
 dmBitsPerPel As Integer
 dmPelsWidth As Long
 dmPelsHeight As Long
 dmDisplayFlags As Long
 dmDisplayFrequency As Long
End Type

Public Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long
Public Declare Function PAGESETUPDLG Lib "comdlg32.dll" Alias "PageSetupDlgA" (pPagesetupdlg As PAGESETUPDLG) As Long
Public Declare Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT) As Long
Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function GlobalLock Lib "Kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "Kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalAlloc Lib "Kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "Kernel32" (ByVal hMem As Long) As Long


Public Function OpenFileDialog(Files As Collection, Optional InitFilename As String = "", _
 Optional Filter As String, Optional DefaultFileExtension As String = "*.*", _
 Optional InitDir As String = "", Optional DialogTitle As String = "", _
 Optional Flags As OpenSaveFlags, Optional hwnd As Long = 0) As Long
 'res = -1 : UserCancel
 'res > 0  : Count of Files

 Dim buff As String, tFil As String, buffA() As String, i As Long, ofn As OPENFILENAME
 If Len(Filter) > 0 Then
   If InStr(Filter, "|") > 0 Then
     tFil = Replace(Filter, "|", vbNullChar)
     tFil = tFil & vbNullChar & vbNullChar
    Else
     tFil = Filter & vbNullChar & vbNullChar
   End If
  Else
   tFil = vbNullChar & vbNullChar
 End If

 With ofn
  .nStructSize = Len(ofn)
  .hWndOwner = hwnd
  .sFilter = tFil
  .nFilterIndex = 0
  .sFile = InitFilename & Space$(1024) & vbNullChar & vbNullChar
  .nMaxFile = Len(.sFile)
  .sDefFileExt = DefaultFileExtension & vbNullChar & vbNullChar
  .sFileTitle = vbNullChar & Space$(512) & vbNullChar & vbNullChar
  .nMaxTitle = Len(ofn.sFileTitle)
  If InitDir = "" Then
    .sInitialDir = App.Path & vbNullChar & vbNullChar
   Else
    .sInitialDir = InitDir & vbNullChar & vbNullChar
  End If
  .sDialogTitle = DialogTitle
  .Flags = Flags
 End With

 Set Files = New Collection
 If GetOpenFileName(ofn) <> 0 Then
   buff = Trim$(Replace$(Left$(ofn.sFile, Len(ofn.sFile) - 2), vbNullChar & vbNullChar, ""))
   Do While Right$(buff, 1) = vbNullChar
    buff = Mid(buff, 1, Len(buff) - 1)
    DoEvents
   Loop
   If Len(buff) > 3 Then
    If InStr(buff, vbNullChar) > 0 Then
      buffA = Split(buff, vbNullChar)
      For i = LBound(buffA) + 1 To UBound(buffA)
       If Len(buffA(i)) > 0 Then
        Files.Add CompletePath(buffA(0)) & buffA(i)
       End If
      Next i
     Else
      Files.Add buff
    End If
   End If
   OpenFileDialog = Files.Count
  Else
   OpenFileDialog = -1
 End If
End Function

Public Function SaveFileDialog(Filename As String, Optional InitFilename As String = "", _
 Optional Filter As String, Optional DefaultFileExtension As String = "*.*", _
 Optional InitDir As String = "", Optional DialogTitle As String = "", _
 Optional Flags As OpenSaveFlags, Optional hwnd As Long = 0) As Long
 'res = -1 : UserCancel
 'res > 0  : Ok

 Dim buff As String, tFil As String, buffA() As String, i As Long, ofn As OPENFILENAME
 If Len(Filter) > 0 Then
   If InStr(Filter, "|") > 0 Then
     tFil = Replace(Filter, "|", vbNullChar)
     tFil = tFil & vbNullChar & vbNullChar
    Else
     tFil = Filter & vbNullChar & vbNullChar
   End If
  Else
   tFil = vbNullChar & vbNullChar
 End If

 With ofn
  .nStructSize = Len(ofn)
  .hWndOwner = hwnd
  .sFilter = tFil
  .nFilterIndex = 0
  .sFile = InitFilename & Space$(1024) & vbNullChar & vbNullChar
  .nMaxFile = Len(.sFile)
  .sDefFileExt = DefaultFileExtension & vbNullChar & vbNullChar
  .sFileTitle = vbNullChar & Space$(512) & vbNullChar & vbNullChar
  .nMaxTitle = Len(ofn.sFileTitle)
  If InitDir = "" Then
    .sInitialDir = App.Path & vbNullChar & vbNullChar
   Else
    .sInitialDir = InitDir & vbNullChar & vbNullChar
  End If
  .sDialogTitle = DialogTitle
  .Flags = Flags
 End With

 If GetSaveFileName(ofn) <> 0 Then
   buff = Trim$(Replace$(Left$(ofn.sFile, Len(ofn.sFile) - 2), vbNullChar & vbNullChar, ""))
   Do While Right$(buff, 1) = vbNullChar
    buff = Mid(buff, 1, Len(buff) - 1)
    DoEvents
   Loop
   If Len(buff) > 3 Then
    Filename = buff
   End If
   SaveFileDialog = ofn.nFilterIndex
  Else
   SaveFileDialog = -1
 End If
End Function


