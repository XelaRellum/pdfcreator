Attribute VB_Name = "modPrinter"
Option Explicit

Public Monitorname As String, Portname As String, DriverName As String, Printername As String

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

    
Private Type MONITOR_INFO_2
 pName As String
 pEnvironment As String
 pDLLName As String
End Type


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

Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (lpString1 As Any, lpString2 As Any) As Long
Private Declare Function FormatMessage Lib "Kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Public Sub InstallCompletePrinter()
 PrinterMonitor Install
 PrinterPort Install
 PrinterDriver Install
 WindowsPrinter Install
End Sub

Public Sub UnInstallCompletePrinter()
 WindowsPrinter UnInstall
 GetAvailablePrinters

 PrinterDriver UnInstall
 GetAvailablePrinterdrivers

 PrinterPort UnInstall
 GetAvailablePorts

 PrinterMonitor UnInstall
 GetAvailableMonitors
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
   If IsPrinterMonitorInstalled(Monitorname) = True Then
     res = DeleteMonitor(vbNullString, vbNullString, Monitorname & vbNullString)
    Else
     WriteLogfile "PrinterMonitor [" & tStr & "]: Printermonitor is not installed."
     Exit Sub
   End If
 End Select
 If res = 0 Then
   WriteLogfile "PrinterMonitor [" & tStr & "]: Error -> " & RaiseAPIError
  Else
   WriteLogfile "PrinterMonitor [" & tStr & "]: Success"
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
   If IsPrinterPortInstalled(Portname) = True Then
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
   If IsPrinterDriverInstalled(DriverName) = False Then
     lngDriverDirectoryLevel = 1
     lngWin32apiResultCode = GetPrinterDriverDirectory(vbNullString, vbNullString, lngDriverDirectoryLevel, ByVal vbNullString, 0, lngDriverDirectoryNeeded)
     ReDim bytDriverDirectoryBuffer(lngDriverDirectoryNeeded - 1)
     lngWin32apiResultCode = GetPrinterDriverDirectory(vbNullString, vbNullString, lngDriverDirectoryLevel, bytDriverDirectoryBuffer(0), lngDriverDirectoryNeeded, lngDriverDirectoryNeeded)
     lngWin32apiResultCode = lstrcpy(ByVal strDriverDirectory, bytDriverDirectoryBuffer(0))

     Driverpath = Left(strDriverDirectory, InStr(strDriverDirectory, vbNullChar) - 1)
     If Right(Driverpath, 1) <> "\" Then Driverpath = Driverpath & "\"

     With DI3
      .pName = DriverName & vbNullString
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
   If IsPrinterDriverInstalled(DriverName) = True Then
     res = DeletePrinterDriver(vbNullString, vbNullString, DriverName & vbNullString)
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
   If IsPrinterInstalled(Printername) = False Then
     With PI
      .pPrinterName = Printername & vbNullString
      .pDriverName = DriverName & vbNullString
      .pPortName = Portname & vbNullString
      .pServerName = vbNullString
      .pComment = Printername & vbNullString
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
      .pShareName = Printername & vbNullString
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
   If IsPrinterInstalled(Printername) = True Then
     With pd
      .pDatatype = 0
      .pDevMode = 0
      .DesiredAccess = PRINTER_ALL_ACCESS
     End With

     res = OpenPrinter(Printername, pHandle, pd)

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

Private Function IsPrinterMonitorInstalled(PrinterMonitor As String) As Boolean
 Dim tColl As Collection, i As Long
 Set tColl = GetAvailableMonitors
 IsPrinterMonitorInstalled = False
 For i = 1 To tColl.Count
  If UCase$(PrinterMonitor) = UCase$(tColl.item(i)) Then
   IsPrinterMonitorInstalled = True
   Exit For
  End If
 Next i
End Function

Private Function IsPrinterPortInstalled(PrinterPort As String) As Boolean
 Dim tColl As Collection, i As Long
 Set tColl = GetAvailablePorts
 IsPrinterPortInstalled = False
 For i = 1 To tColl.Count
  If UCase$(PrinterPort) = UCase$(tColl.item(i)) Then
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

Private Function IsPrinterInstalled(Printername As String) As Boolean
 Dim tColl As Collection, i As Long
 Set tColl = GetAvailablePrinters
 IsPrinterInstalled = False
 For i = 1 To tColl.Count
  If UCase$(Printername) = UCase$(tColl.item(i)) Then
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
