Attribute VB_Name = "modPrinterEnums"
Option Explicit

Private Const PRINTER_ATTRIBUTE_DEFAULT = &H4
Private Const PRINTER_ATTRIBUTE_DIRECT = &H2
Private Const PRINTER_ATTRIBUTE_ENABLE_BIDI = &H800&
Private Const PRINTER_ATTRIBUTE_LOCAL = &H40
Private Const PRINTER_ATTRIBUTE_NETWORK = &H10
Private Const PRINTER_ATTRIBUTE_QUEUED = &H1
Private Const PRINTER_ATTRIBUTE_SHARED = &H8
Private Const PRINTER_ATTRIBUTE_WORK_OFFLINE = &H400
Private Const PRINTER_ENUM_CONNECTIONS = &H4
Private Const PRINTER_ENUM_CONTAINER = &H8000&
Private Const PRINTER_ENUM_DEFAULT = &H1
Private Const PRINTER_ENUM_EXPAND = &H4000
Private Const PRINTER_ENUM_LOCAL = &H2
Private Const PRINTER_ENUM_ICON1 = &H10000
Private Const PRINTER_ENUM_ICON2 = &H20000
Private Const PRINTER_ENUM_ICON3 = &H40000
Private Const PRINTER_ENUM_ICON4 = &H80000
Private Const PRINTER_ENUM_ICON5 = &H100000
Private Const PRINTER_ENUM_ICON6 = &H200000
Private Const PRINTER_ENUM_ICON7 = &H400000
Private Const PRINTER_ENUM_ICON8 = &H800000
Private Const PRINTER_ENUM_NAME = &H8
Private Const PRINTER_ENUM_NETWORK = &H40
Private Const PRINTER_ENUM_REMOTE = &H10
Private Const PRINTER_ENUM_SHARED = &H20
Private Const PRINTER_LEVEL1 = &H1
Private Const PRINTER_LEVEL4 = &H4
Private Const SIZEOFMONITOR_INFO_1 = 4
Private Const SIZEOFPORT_INFO_2 = 20
Private Const SIZEOFPRINTER_INFO_1 = 16
Private Const SIZEOFPRINTER_INFO_4 = 12

Private Type DRIVER_INFO_1
 pName As Long
End Type

Private Type DRIVER_INFO_3
 cVersion As Long
 pName As String
 pEnvironment As String
 pDriverPath As String
 pDataFile As String
 pConfigFile As String
 pHelpFile As String
 pDependentFiles As String
 pMonitorName As String
 pDefaultDataType As String
End Type

Private Type MONITOR_INFO_1
 pName As Long
End Type

Private Type MONITOR_INFO_2
 pName As Long
 pEnvironment As Long
 pDLLName As Long
End Type

Private Type PORT_INFO_2
 pPortName    As Long
 pMonitorName As Long
 pDescription As Long
 fPortType    As Long
 Reserved     As Long
End Type

Private Enum PortTypes
 PORT_TYPE_WRITE = &H1
 PORT_TYPE_READ = &H2
 PORT_TYPE_REDIRECTED = &H4
 PORT_TYPE_NET_ATTACHED = &H8
End Enum

Public Type PRINTER_INFO_1
 Flags As Long
 prescription As Long
 Pane As Long
 Comment As Long
End Type

Public Type PRINTER_INFO_4
 pPrinterName As Long
 pServerName As Long
 Attributes As Long
End Type

Private Declare Function EnumMonitors Lib "winspool.drv" Alias "EnumMonitorsA" (ByVal pName As String, ByVal Level As Long, pMonitors As Any, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare Function EnumPorts Lib "winspool.drv" Alias "EnumPortsA" (ByVal pName As String, ByVal nLevel As Long, lpbPorts As Any, ByVal cbBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersA" (ByVal Flags As Long, ByVal Name As String, ByVal Level As Long, pPrinterEnum As Any, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare Function EnumPrinterDrivers Lib "winspool.drv" Alias "EnumPrinterDriversA" (ByVal pName As String, ByVal pEnvironment As String, ByVal Level As Long, pDriverInfo As Any, ByVal cdBuf As Long, pcbNeeded As Long, pcRetruned As Long) As Long

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (lpString As Any) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (lpString1 As Any, lpString2 As Any) As Long
Private Declare Sub MoveMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Function GetAvailableMonitors() As Collection
 Dim pcbNeeded As Long, pcReturned As Long, mi1() As MONITOR_INFO_1, _
  i As Integer, sPortType As String, tColl As Collection
   
 Set tColl = New Collection
 EnumMonitors vbNullString, 1, 0, 0, pcbNeeded, pcReturned
 
 If pcbNeeded Then
  ReDim mi1((pcbNeeded / SIZEOFMONITOR_INFO_1))
  If EnumMonitors(vbNullString, 1, mi1(0), pcbNeeded, pcbNeeded, pcReturned) Then
   For i = 0 To (pcReturned - 1)
    tColl.Add GetStrFromPtrA(mi1(i).pName)
   Next i
  End If
 End If
 
 Set GetAvailableMonitors = tColl
End Function

Public Function GetAvailablePorts() As Collection
 Dim pcbNeeded As Long, pcReturned As Long, pi2() As PORT_INFO_2, _
  i As Integer, sPortType As String, tColl As Collection
   
 Set tColl = New Collection
 Call EnumPorts(vbNullString, 2, 0, 0, pcbNeeded, pcReturned)
 If pcbNeeded Then
  ReDim pi2((pcbNeeded / SIZEOFPORT_INFO_2))
  If EnumPorts(vbNullString, 2, pi2(0), pcbNeeded, pcbNeeded, pcReturned) Then
   For i = 0 To (pcReturned - 1)
    tColl.Add GetStrFromPtrA(pi2(i).pPortName)
   Next i
  End If
 End If
 
 Set GetAvailablePorts = tColl
End Function

Public Function GetAvailablePrinterdrivers() As Collection
 Dim lngDriverInfo1Level As Long, lngDriverInfo1Needed As Long, _
  lngDriverInfo1Returned As Long, bytDriverInfo1Buffer() As Byte, _
  udtDriverInfo1() As DRIVER_INFO_1, lngDriverInfo1Count As Long, _
  strDriverInfo1Name As String * 128, lngWin32apiResultCode As Long, _
  tColl As Collection
 
 Set tColl = New Collection
 lngDriverInfo1Level = 1
 lngWin32apiResultCode = EnumPrinterDrivers(vbNullString, vbNullString, _
  lngDriverInfo1Level, ByVal vbNullString, 0, lngDriverInfo1Needed, lngDriverInfo1Returned)
 If lngDriverInfo1Needed > 0 Then
  ReDim bytDriverInfo1Buffer(lngDriverInfo1Needed - 1)
  lngWin32apiResultCode = EnumPrinterDrivers(vbNullString, vbNullString, _
   lngDriverInfo1Level, bytDriverInfo1Buffer(0), lngDriverInfo1Needed, _
   lngDriverInfo1Needed, lngDriverInfo1Returned)
  ReDim udtDriverInfo1(lngDriverInfo1Returned - 1)
  MoveMemory udtDriverInfo1(0), bytDriverInfo1Buffer(0), Len(udtDriverInfo1(0)) * lngDriverInfo1Returned
  For lngDriverInfo1Count = 0 To lngDriverInfo1Returned - 1
   lngWin32apiResultCode = lstrcpy(ByVal strDriverInfo1Name, ByVal udtDriverInfo1 _
    (lngDriverInfo1Count).pName)
   tColl.Add Left(strDriverInfo1Name, InStr(strDriverInfo1Name, vbNullChar) - 1)
  Next lngDriverInfo1Count
 End If
 Set GetAvailablePrinterdrivers = tColl
End Function

Public Function GetAvailablePrinters() As Collection
 If IsWin9xMe = True Then
   Set GetAvailablePrinters = EnumPrintersWin9x
  Else
   Set GetAvailablePrinters = EnumPrintersWinNT
 End If
End Function

Private Function EnumPrintersWinNT() As Collection
 Dim Success As Boolean, cbRequired As Long, cbBuffer As Long, nEntries As Long, _
  pntr() As PRINTER_INFO_4, c As Long, tColl As Collection
 
 Set tColl = New Collection
   
 Call EnumPrinters(PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, vbNullString, _
  PRINTER_LEVEL4, 0, 0, cbRequired, nEntries)
            
 ReDim pntr((cbRequired \ SIZEOFPRINTER_INFO_4))
 cbBuffer = cbRequired
 If EnumPrinters(PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, vbNullString, _
  PRINTER_LEVEL4, pntr(0), cbBuffer, cbRequired, nEntries) Then
  For c = 0 To nEntries - 1
   tColl.Add GetStrFromPtrA(pntr(c).pPrinterName)
  Next c
 End If
 Set EnumPrintersWinNT = tColl
End Function

Public Function EnumPrintersWin9x() As Collection
 Dim cbRequired As Long, cbBuffer As Long, pntr() As PRINTER_INFO_1, nEntries As Long, _
  c As Long, sFlags As String, tColl As Collection
   
 Set tColl = New Collection
   
 Call EnumPrinters(PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, vbNullString, _
  PRINTER_LEVEL1, 0, 0, cbRequired, nEntries)
                     
 ReDim pntr((cbRequired \ SIZEOFPRINTER_INFO_1))
 cbBuffer = cbRequired
    
 If EnumPrinters(PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, vbNullString, _
  PRINTER_LEVEL1, pntr(0), cbBuffer, cbRequired, nEntries) Then
 
  For c = 0 To nEntries - 1
   tColl.Add GetStrFromPtrA(pntr(c).Pane)
  Next c
 End If
 Set EnumPrintersWin9x = tColl
End Function

Public Function GetStrFromPtrA(lpszA As Long) As String
 GetStrFromPtrA = String$(lstrlen(ByVal lpszA), 0)
 Call lstrcpy(ByVal GetStrFromPtrA, ByVal lpszA)
End Function

