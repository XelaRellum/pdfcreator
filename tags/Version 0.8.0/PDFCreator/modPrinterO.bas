Attribute VB_Name = "modPrinter"
Option Explicit

Public Monitorname As String, Portname As String, Drivername As String, PrinterName As String

Private Enum eInstall
 Install = 0
 UnInstall = 1
End Enum

Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200

Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const PRINTER_ACCESS_ADMINISTER = &H4
Private Const PRINTER_ACCESS_USE = &H8
Private Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

Private Type PRINTER_DEFAULTS
 pDatatype As Long
 pDevMode As Long
 DesiredAccess As Long
End Type

Private Type PRINTER_INFO_2
 pServerName As String
 pPrinterName As String
 pShareName As String
 pPortName As String
 pDriverName As String
 pComment As String
 pLocation As String
 pDevMode As Long
 pSepFile As String
 pPrintProcessor As String
 pDatatype As String
 pParameters As String
 pSecurityDescriptor As Long
 Attributes As Long
 Priority As Long
 DefaultPriority As Long
 StartTime As Long
 UntilTime As Long
 Status As Long
 cJobs As Long
 AveragePPM As Long
End Type

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

Private Type PORT_INFO_2
 pPortName    As Long
 pMonitorName As Long
 pDescription As Long
 fPortType    As Long
 Reserved     As Long
End Type
Private Const SIZEOFPORT_INFO_2 = 20
    
Public Type MONITOR_INFO_1
  pName As Long
End Type
Public Const SIZEOFMONITOR_INFO_1 = 4

Private Type MONITOR_INFO_2
 pName As String
 pEnvironment As String
 pDLLName As String
End Type


'----
Private Type PRINTER_INFO_1
 Flags As Long
 prescription As Long
 Pane As Long
 Comment As Long
End Type

Private Type PRINTER_INFO_4
 pPrinterName As Long
 pServerName As Long
 Attributes As Long
End Type

Private Const SIZEOFPRINTER_INFO_1 = 16
Private Const SIZEOFPRINTER_INFO_4 = 12

Private Const PRINTER_LEVEL1 = &H1
Private Const PRINTER_LEVEL4 = &H4

Private Const PRINTER_ENUM_DEFAULT = &H1
Private Const PRINTER_ENUM_LOCAL = &H2
Private Const PRINTER_ENUM_CONNECTIONS = &H4
Private Const PRINTER_ENUM_NAME = &H8
Private Const PRINTER_ENUM_REMOTE = &H10
Private Const PRINTER_ENUM_SHARED = &H20
Private Const PRINTER_ENUM_NETWORK = &H40

Private Const PRINTER_ATTRIBUTE_DEFAULT = &H4
Private Const PRINTER_ATTRIBUTE_DIRECT = &H2
Private Const PRINTER_ATTRIBUTE_ENABLE_BIDI = &H800&
Private Const PRINTER_ATTRIBUTE_LOCAL = &H40
Private Const PRINTER_ATTRIBUTE_NETWORK = &H10
Private Const PRINTER_ATTRIBUTE_QUEUED = &H1
Private Const PRINTER_ATTRIBUTE_SHARED = &H8
Private Const PRINTER_ATTRIBUTE_WORK_OFFLINE = &H400

Private Const PRINTER_ENUM_CONTAINER = &H8000&
Private Const PRINTER_ENUM_EXPAND = &H4000
Private Const PRINTER_ENUM_ICON1 = &H10000
Private Const PRINTER_ENUM_ICON2 = &H20000
Private Const PRINTER_ENUM_ICON3 = &H40000
Private Const PRINTER_ENUM_ICON4 = &H80000
Private Const PRINTER_ENUM_ICON5 = &H100000
Private Const PRINTER_ENUM_ICON6 = &H200000
Private Const PRINTER_ENUM_ICON7 = &H400000
Private Const PRINTER_ENUM_ICON8 = &H800000


Private Declare Function AddMonitor Lib "winspool.drv" Alias "AddMonitorA" (ByVal pName As String, ByVal Level As Long, pMonitors As Any) As Long
Private Declare Function AddPrinter Lib "winspool.drv" Alias "AddPrinterA" (ByVal pName As String, ByVal Level As Long, pPrinter As Any) As Long
Private Declare Function AddPrinterDriver Lib "winspool.drv" Alias "AddPrinterDriverA" (ByVal pName As String, ByVal Level As Long, pDriverInfo As Any) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function DeleteMonitor Lib "winspool.drv" Alias "DeleteMonitorA" (ByVal pName As String, ByVal pEnviroment As String, ByVal pMonitorName As String) As Long
Private Declare Function DeletePort Lib "winspool.drv" Alias "DeletePortA" (ByVal pName As String, ByVal hwnd As Long, ByVal pPortName As String) As Long
Private Declare Function DeletePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function DeletePrinterDriver Lib "winspool.drv" Alias "DeletePrinterDriverA" (ByVal pName As String, ByVal pEnviroment As String, ByVal pDriverName As String) As Long
Private Declare Function GetPrinterDriverDirectory Lib "winspool.drv" Alias "GetPrinterDriverDirectoryA" (ByVal pName As String, ByVal pEnvironment As String, ByVal Level As Long, pDriverDirectory As Any, ByVal cdBuf As Long, pcbNeeded As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long

Private Declare Function EnumMonitors Lib "winspool.drv" Alias "EnumMonitorsA" (ByVal pName As String, ByVal Level As Long, pMonitors As Any, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare Function EnumPorts Lib "winspool.drv" Alias "EnumPortsA" (ByVal pName As String, ByVal nLevel As Long, lpbPorts As Any, ByVal cbBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare Function EnumPrinterDrivers Lib "winspool.drv" Alias "EnumPrinterDriversA" (ByVal pName As String, ByVal pEnvironment As String, ByVal Level As Long, pDriverInfo As Any, ByVal cdBuf As Long, pcbNeeded As Long, pcRetruned As Long) As Long
Private Declare Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersA" (ByVal Flags As Long, ByVal Name As String, ByVal Level As Long, pPrinterEnum As Any, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long

Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (lpString1 As Any, lpString2 As Any) As Long
Private Declare Sub MoveMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Public Sub InstallCompletePrinter()
 PrinterMonitor Install
 GetAvailableMonitors
 PrinterPort Install
 GetAvailablePorts
 PrinterDriver Install
 WindowsPrinter Install
End Sub

Public Sub UnInstallCompletePrinter()
 WindowsPrinter UnInstall
 GetAvailablePrinters
 PrinterDriver UnInstall: GetAvailablePrinterdrivers
 MsgBox "Uninstall Monitor"
 PrinterMonitor UnInstall: GetAvailableMonitors
 MsgBox "Uninstall Port"
 PrinterPort UnInstall: GetAvailablePorts
 MsgBox "Uninstall Ready"
End Sub

Private Sub PrinterMonitor(InstallTyp As eInstall)
 Dim res As Long, Monitor2 As MONITOR_INFO_2, tStr As String
 Select Case InstallTyp
  Case 0: ' Install
   tStr = "Install"
   If IsPrinterMonitorInstalled(Monitorname) = False Then
     With Monitor2
      .pName = Monitorname & Chr$(0)
      If IsWin9xMe = True Then
        .pEnvironment = "Windows 4.0" & Chr$(0)
        .pDLLName = "redmon95.dll" & Chr$(0)
       Else
        .pEnvironment = "Windows NT x86" & Chr$(0)
        .pDLLName = "redmonnt.dll" & Chr$(0)
      End If
     End With
     res = AddMonitor(vbNullString, 2, Monitor2)
    Else
     WriteLogfile "PrinterMonitor [" & tStr & "]: Printermonitor is already installed."
     Exit Sub
    End If
  Case 1: ' UnInstall
   tStr = "UnInstall"
   MsgBox "Monitorname:" & Monitorname
   If IsPrinterMonitorInstalled(Monitorname) = False Then
     MsgBox "Unistall the monitor"
     res = DeleteMonitor(vbNullString, vbNullString, Monitorname & vbNullString)
    Else
     WriteLogfile "PrinterMonitor [" & tStr & "]: Printermonitor is not installed."
     Exit Sub
   End If
 End Select
 If res = 0 Then
   WriteLogfile "PrinterMonitor [" & tStr & "]: Error -> " & RaiseAPIError
   MsgBox "PrinterMonitor [" & tStr & "]: Error -> " & RaiseAPIError
  Else
   WriteLogfile "PrinterMonitor [" & tStr & "]: Success"
   MsgBox "PrinterMonitor [" & tStr & "]: Success"
 End If
End Sub

Private Sub PrinterPort(InstallTyp As eInstall)
 Dim reg As clsRegistry, b() As Byte, tStr As String, i As Integer, res As Long
 Set reg = New clsRegistry
 Select Case InstallTyp
  Case 0: ' Install
   tStr = "Install"
   If IsPrinterPortInstalled(Portname) = False Then
     reg.hkey = HKEY_LOCAL_MACHINE
     reg.KeyRoot = "System\CurrentControlSet\Control\Print\Monitors"
     reg.CreateKey Monitorname & "\Ports\" & Portname
     reg.KeyRoot = "System\CurrentControlSet\Control\Print\Monitors\" & Monitorname & "\Ports\" & Portname
     reg.SetRegistryValue "Arguments", "-PPDFCREATORPRINTER", REG_SZ
     reg.SetRegistryValue "Command", App.Path & "\pdfcreator.exe", REG_SZ
     i = 300
     reg.SetRegistryValue "Delay", i, REG_DWORD
     reg.SetRegistryValue "Description", "Redirected Port", REG_SZ
     i = 0
     reg.SetRegistryValue "LogFileDebug", i, REG_DWORD
     reg.SetRegistryValue "LogFileName", vbNullString, REG_SZ
     reg.SetRegistryValue "LogFileUse", i, REG_DWORD
     reg.SetRegistryValue "Output", i, REG_DWORD
     reg.SetRegistryValue "Printer", "PDFCreator", REG_SZ
     reg.SetRegistryValue "PrintError", i, REG_DWORD
     reg.SetRegistryValue "RunUser", i, REG_DWORD
     i = 0
     reg.SetRegistryValue "ShowWindow", i, REG_DWORD
     reg.KeyRoot = vbNullString
     reg.CreateKey "System\CurrentControlSet\Control\Print\Ports\" & Portname
    Else
     WriteLogfile "PrinterPort [" & tStr & "]: Printerport is already installed."
     Exit Sub
   End If
   res = 1
  Case 1: ' UnInstall
   tStr = "UnInstall"
   If IsPrinterPortInstalled(Portname) = False Then
     res = DeletePort(vbNullString, 0, Portname & vbNullString)
    Else
     WriteLogfile "PrinterPort [" & tStr & "]: Printerport is not installed."
     Exit Sub
   End If
'   reg.hkey = HKEY_LOCAL_MACHINE
'   reg.DeleteKey "System\CurrentControlSet\Control\Print\Monitors\" & Monitorname & "\Ports\" & Portname
'   reg.DeleteKey "System\CurrentControlSet\Control\Print\Ports\" & Portname
 End Select
 Set reg = Nothing
 If res = 0 Then
   WriteLogfile "PrinterPort [" & tStr & "]: Error -> " & RaiseAPIError
  Else
   WriteLogfile "PrinterPort [" & tStr & "]: Success"
 End If
End Sub

Private Sub PrinterDriver(InstallTyp As eInstall)
 Dim DI3 As DRIVER_INFO_3, Driverpath As String, res As Long
 Dim lngDriverDirectoryLevel    As Long
 Dim lngDriverDirectoryNeeded   As Long
 Dim bytDriverDirectoryBuffer() As Byte
 Dim strDriverDirectory         As String * 512
 Dim lngWin32apiResultCode      As Long
 Dim tStr As String
 Select Case InstallTyp
  Case 0: ' Install
   tStr = "Install"
   If IsPrinterDriverInstalled(Drivername) = False Then
     lngDriverDirectoryLevel = 1
     lngWin32apiResultCode = GetPrinterDriverDirectory(vbNullString, vbNullString, lngDriverDirectoryLevel, ByVal vbNullString, 0, lngDriverDirectoryNeeded)
     ReDim bytDriverDirectoryBuffer(lngDriverDirectoryNeeded - 1)
     lngWin32apiResultCode = GetPrinterDriverDirectory(vbNullString, vbNullString, lngDriverDirectoryLevel, bytDriverDirectoryBuffer(0), lngDriverDirectoryNeeded, lngDriverDirectoryNeeded)
     lngWin32apiResultCode = lstrcpy(ByVal strDriverDirectory, bytDriverDirectoryBuffer(0))

     Driverpath = Left(strDriverDirectory, InStr(strDriverDirectory, vbNullChar) - 1)
     If Right(Driverpath, 1) <> "\" Then Driverpath = Driverpath & "\"

     With DI3
      .pName = Drivername & vbNullString
      If IsWin9xMe = True Then
        .cVersion = 0
        .pConfigFile = "PSCRIPT.DRV" & vbNullString
        .pDataFile = "APLWCSB1.SPD" & vbNullString
        .pDriverPath = "PSCRIPT.DRV" & vbNullString
        .pEnvironment = "Windows 4.0" & vbNullString
        .pHelpFile = "PSCRIPT.HLP" & vbNullString
        .pDependentFiles = "PSCRIPT.DRV" & Chr$(0) & "PSCRIPT.HLP" & Chr$(0) & Chr$(0)
        .pMonitorName = Monitorname & vbNullString
       Else
        .cVersion = 1
        .pConfigFile = Driverpath & "PSCRPTUI.DLL" & vbNullString
        .pDataFile = Driverpath & "ADIST5.PPD" & vbNullString
        .pDriverPath = Driverpath & "PSCRIPT.DLL" & vbNullString
        .pEnvironment = vbNullString
        .pHelpFile = Driverpath & "PSCRIPT.HLP" & vbNullString
        .pDependentFiles = Driverpath & "PSCRIPT.DLL" & vbNullString & vbNullString
        .pMonitorName = vbNullString
      End If
      .pDefaultDataType = "RAW" & vbNullString
     End With

     res = AddPrinterDriver(vbNullString, 3, DI3)
    Else
     WriteLogfile "PrinterDriver [" & tStr & "]: Printerdriver is already installed."
     Exit Sub
   End If
  Case 1: ' UnInstall
   tStr = "UnInstall"
   If IsPrinterDriverInstalled(Drivername) = True Then
     res = DeletePrinterDriver(vbNullString, vbNullString, Drivername & vbNullString)
    Else
     WriteLogfile "PrinterDriver [" & tStr & "]: Printerdriver is not installed."
     Exit Sub
   End If
 End Select
 If res = 0 Then
   WriteLogfile "PrinterDriver [" & tStr & "]: Error -> " & RaiseAPIError
  Else
   WriteLogfile "PrinterDriver [" & tStr & "]: Success"
 End If
End Sub

Private Sub WindowsPrinter(InstallTyp As eInstall)
 Dim res As Long, PI As PRINTER_INFO_2, pHandle As Long, _
  pd As PRINTER_DEFAULTS, tStr As String, ini As clsINI
 Select Case InstallTyp
  Case 0: ' Install
   tStr = "Install"
   If IsPrinterInstalled(PrinterName) = False Then
     With PI
      .pPrinterName = PrinterName & vbNullString
      .pDriverName = Drivername & vbNullString
      .pPortName = Portname & vbNullString
      .pServerName = vbNullString
      .pComment = PrinterName & vbNullString
      .pPrintProcessor = "WinPrint" & vbNullString
      .Priority = 1
      .DefaultPriority = 1
      .pDatatype = "RAW" & vbNullString
      .AveragePPM = 0
      .cJobs = 0
      .pDevMode = 0
      .pLocation = vbNullString
      .pParameters = 0
      .pSecurityDescriptor = 0
      .pShareName = PrinterName & vbNullString
      .StartTime = 0
      .Status = 0
      .UntilTime = 0
     End With

     res = AddPrinter(vbNullString, 2, PI)
     If res <> 0 Then
      ClosePrinter res
     End If
     If IsWin9xMe = True Then
      Set ini = New clsINI
      ini.Filename = GetWindowsDirectory & "\win.ini"
      ini.Key = "PDFCreator"
      ini.Section = "Devices"
      ini.SaveKey "PSCRIPT,PDFCreator:"
      ini.Section = "PrinterPorts"
      ini.SaveKey "PSCRIPT,PDFCreator:,15,45"
      Set ini = Nothing
     End If
    Else
     WriteLogfile "Printer [" & tStr & "]: Printer is already installed."
     Exit Sub
   End If
  Case 1: ' UnInstall
   tStr = "UnInstall"
   If IsPrinterInstalled(PrinterName) = True Then
     With pd
      .pDatatype = 0
      .pDevMode = 0
      .DesiredAccess = PRINTER_ALL_ACCESS
     End With

     res = OpenPrinter(PrinterName, pHandle, pd)

     If res <> 0 Then
      res = DeletePrinter(pHandle)
     End If
     If res <> 0 Then
      res = ClosePrinter(pHandle)
     End If
    Else
     WriteLogfile "Printer [" & tStr & "]: Printer is not installed."
     Exit Sub
   End If
 End Select
 If res = 0 Then
   WriteLogfile "Printer [" & tStr & "]: Error -> " & RaiseAPIError
  Else
   WriteLogfile "Printer [" & tStr & "]: Success"
 End If
End Sub

Private Function IsPrinterMonitorInstalled(PrinterDriver As String) As Boolean
 Dim tColl As Collection, i As Long
 Set tColl = GetAvailableMonitors
 IsPrinterMonitorInstalled = False
 For i = 1 To tColl.Count
  If UCase$(PrinterDriver) = UCase$(tColl.item(i)) Then
   IsPrinterMonitorInstalled = True
   Exit For
  End If
 Next i
End Function

Private Function IsPrinterPortInstalled(PrinterDriver As String) As Boolean
 Dim tColl As Collection, i As Long
 Set tColl = GetAvailablePorts
 IsPrinterPortInstalled = False
 For i = 1 To tColl.Count
  If UCase$(PrinterDriver) = UCase$(tColl.item(i)) Then
   IsPrinterPortInstalled = True
   Exit For
  End If
 Next i
End Function

Private Function IsPrinterDriverInstalled(PrinterDriver As String) As Boolean
 Dim tColl As Collection, i As Long
 Set tColl = GetAvailablePrinterdrivers
 IsPrinterDriverInstalled = False
 For i = 1 To tColl.Count
  If UCase$(PrinterDriver) = UCase$(tColl.item(i)) Then
   IsPrinterDriverInstalled = True
   Exit For
  End If
 Next i
End Function

Public Function IsPrinterInstalled(PrinterName As String) As Boolean
 Dim tColl As Collection, i As Long
 Set tColl = GetAvailablePrinters
 IsPrinterInstalled = False
 For i = 1 To tColl.Count
  MsgBox i & ": " & tColl.item(i)
  If UCase$(PrinterName) = UCase$(tColl.item(i)) Then
   IsPrinterInstalled = True
   Exit For
  End If
 Next i
End Function

Public Function RaiseAPIError() As String
 Dim ErrorMsg As String, ErrNum As Long
 ErrNum = Err.LastDllError
 ErrorMsg = String(256, 0)
 ErrorMsg = Left$(ErrorMsg, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, ErrNum, 0&, ErrorMsg, Len(ErrorMsg), ByVal 0))
 If Mid(ErrorMsg, Len(ErrorMsg) - 1) = vbCrLf Then
  ErrorMsg = Mid(ErrorMsg, 1, Len(ErrorMsg) - 2)
 End If
 RaiseAPIError = ErrorMsg
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
 If lngDriverInfo1Needed <= 0 Then
   tColl.Add "Found no printerdriver."
  Else
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

Public Function GetAvailableMonitors() As Collection
 Dim pcbNeeded As Long, pcReturned As Long, mi1() As MONITOR_INFO_1, i As Long, _
  sPortType As String, tColl As Collection
   
 Set tColl = New Collection
 EnumMonitors vbNullString, 1, 0, 0, pcbNeeded, pcReturned
 
 If pcbNeeded <> 0 Then
  ReDim mi1(pcbNeeded / SIZEOFMONITOR_INFO_1)
  If EnumMonitors(vbNullString, 1, mi1(0), pcbNeeded, pcbNeeded, pcReturned) <> 0 Then
   For i = 0 To (pcReturned - 1)
     tColl.Add GetStrFromPtrA(mi1(i).pName)
   Next i
  End If
 End If
 
 Set GetAvailableMonitors = tColl
End Function

Public Function GetAvailablePorts() As Collection
 Dim pcbNeeded As Long, pcReturned As Long, pi2() As PORT_INFO_2, i As Long, _
  sPortType As String, tColl As Collection
   
 Set tColl = New Collection
 Call EnumPorts(vbNullString, 2, 0, 0, pcbNeeded, pcReturned)
 If pcbNeeded <> 0 Then
  ReDim pi2((pcbNeeded / SIZEOFPORT_INFO_2))
  If EnumPorts(vbNullString, 2, pi2(0), pcbNeeded, pcbNeeded, pcReturned) Then
   For i = 0 To (pcReturned - 1)
    tColl.Add GetStrFromPtrA(pi2(i).pPortName)
   Next i
  End If
 End If
 Set GetAvailablePorts = tColl
End Function

Public Function GetAvailablePrinters() As Collection
 If IsWin9xMe = True Then
   Set GetAvailablePrinters = EnumPrintersWin9x
  Else
   Set GetAvailablePrinters = EnumPrintersWinNT
 End If
End Function

Private Function EnumPrintersWinNT() As Collection
 Dim Success As Boolean, cbRequired As Long, cbBuffer As Long, _
  pntr() As PRINTER_INFO_4, nEntries As Long, i As Long, sAttr As String
   
 Dim tColl As Collection
 Set tColl = New Collection
   
 Call EnumPrinters(PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, vbNullString, _
                   PRINTER_LEVEL4, 0, 0, cbRequired, nEntries)
 ReDim pntr((cbRequired \ SIZEOFPRINTER_INFO_4))
 cbBuffer = cbRequired
 If EnumPrinters(PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, vbNullString, _
                 PRINTER_LEVEL4, pntr(0), cbBuffer, cbRequired, nEntries) Then
   For i = 0 To nEntries - 1
    With pntr(i)
     sAttr = ""
     If (.Attributes And PRINTER_ATTRIBUTE_DEFAULT) Then sAttr = "default "
     If (.Attributes And PRINTER_ATTRIBUTE_DIRECT) Then sAttr = sAttr & "direct "
     If (.Attributes And PRINTER_ATTRIBUTE_ENABLE_BIDI) Then sAttr = sAttr & "bidirectional "
     If (.Attributes And PRINTER_ATTRIBUTE_LOCAL) Then sAttr = sAttr & "local "
     If (.Attributes And PRINTER_ATTRIBUTE_NETWORK) Then sAttr = sAttr & "net "
     If (.Attributes And PRINTER_ATTRIBUTE_QUEUED) Then sAttr = sAttr & "queued "
     If (.Attributes And PRINTER_ATTRIBUTE_SHARED) Then sAttr = sAttr & "shared "
     If (.Attributes And PRINTER_ATTRIBUTE_WORK_OFFLINE) Then sAttr = sAttr & "offline "
     tColl.Add GetStrFromPtrA(.pPrinterName) ' & "; " & GetStrFromPtrA(.pServerName) & "; " & sAttr
    End With
   Next i
  Else
   tColl.Add "Error enumerating printers."
 End If
 Set EnumPrintersWinNT = tColl
End Function

Public Function EnumPrintersWin9x() As Collection
 Dim cbRequired As Long, cbBuffer As Long, pntr() As PRINTER_INFO_1, nEntries As Long, _
  i As Long, sFlags As String
    
 Dim tColl As Collection
 Set tColl = New Collection
   
 Call EnumPrinters(PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, vbNullString, PRINTER_LEVEL1, 0, 0, cbRequired, nEntries)
 ReDim pntr((cbRequired \ SIZEOFPRINTER_INFO_1))
 cbBuffer = cbRequired
 If EnumPrinters(PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, vbNullString, PRINTER_LEVEL1, _
                 pntr(0), cbBuffer, cbRequired, nEntries) <> 0 Then
   For i = 0 To nEntries - 1
    With pntr(i)
     sFlags = ""
     If (.Flags And PRINTER_ENUM_CONTAINER) Then sFlags = "enumerable "
     If (.Flags And PRINTER_ENUM_EXPAND) Then sFlags = sFlags & "expand "
     If (.Flags And PRINTER_ENUM_ICON1) Then sFlags = sFlags & "icon1 "
     If (.Flags And PRINTER_ENUM_ICON2) Then sFlags = sFlags & "icon2 "
     If (.Flags And PRINTER_ENUM_ICON3) Then sFlags = sFlags & "icon3 "
     If (.Flags And PRINTER_ENUM_ICON8) Then sFlags = sFlags & "icon8 "
            
     tColl.Add GetStrFromPtrA(.Pane) & "; " & sFlags & "; " & GetStrFromPtrA(.prescription)
    End With
   Next i
  Else
   tColl.Add "Error enumerating printers."
 End If
 Set EnumPrintersWin9x = tColl
End Function
