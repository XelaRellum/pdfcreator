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
Private Declare Function DeletePort Lib "winspool.drv" Alias "DeletePortA" (ByVal pName As String, ByVal hWnd As Long, ByVal pPortName As String) As Long
Private Declare Function DeletePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function DeletePrinterDriver Lib "winspool.drv" Alias "DeletePrinterDriverA" (ByVal pName As String, ByVal pEnviroment As String, ByVal pDriverName As String) As Long
Private Declare Function GetPrinterDriverDirectory Lib "winspool.drv" Alias "GetPrinterDriverDirectoryA" (ByVal pName As String, ByVal pEnvironment As String, ByVal Level As Long, pDriverDirectory As Any, ByVal cdBuf As Long, pcbNeeded As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long

Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (lpString1 As Any, lpString2 As Any) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Public Sub InstallCompletePrinter()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  PrinterMonitor Install
50020  PrinterPort Install
50030  PrinterDriver Install
50040  WindowsPrinter Install
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "InstallCompletePrinter")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub UnInstallCompletePrinter()
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  WindowsPrinter UnInstall
50020  PrinterDriver UnInstall
50030  PrinterPort UnInstall
50040  PrinterMonitor UnInstall
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "UnInstallCompletePrinter")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub PrinterMonitor(InstallTyp As eInstall)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim res As Long, Monitor2 As MONITOR_INFO_2, tStr As String
50020  Select Case InstallTyp
  Case 0: ' Install
50040    tStr = "Install"
50050    If IsPrinterMonitorInstalled(Monitorname) = False Then
50060      With Monitor2
50070       .pName = Monitorname & Chr$(0)
50080       If IsWin9xMe = True Then
50090         .pEnvironment = "Windows 4.0" & Chr$(0)
50100         .pDLLName = "redmon95.dll" & Chr$(0)
50110        Else
50120         .pEnvironment = "Windows NT x86" & Chr$(0)
50130         .pDLLName = "redmonnt.dll" & Chr$(0)
50140       End If
50150      End With
50160      res = AddMonitor(vbNullString, 2, Monitor2)
50170     Else
50180      WriteLogfile "PrinterMonitor [" & tStr & "]: Printermonitor is already installed."
50190      Exit Sub
50200     End If
50210   Case 1: ' UnInstall
50220    tStr = "UnInstall"
50230    If IsPrinterMonitorInstalled(Monitorname) = False Then
50240      res = DeleteMonitor(vbNullString, vbNullString, Monitorname & vbNullString)
50250     Else
50260      WriteLogfile "PrinterMonitor [" & tStr & "]: Printermonitor is not installed."
50270      Exit Sub
50280    End If
50290  End Select
50300  If res = 0 Then
50310    WriteLogfile "PrinterMonitor [" & tStr & "]: Error -> " & RaiseAPIError
50320   Else
50330    WriteLogfile "PrinterMonitor [" & tStr & "]: Success"
50340  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "PrinterMonitor")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub PrinterPort(InstallTyp As eInstall)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry, b() As Byte, tStr As String, i As Integer, res As Long
50020  Set reg = New clsRegistry
50030  Select Case InstallTyp
  Case 0: ' Install
50050    tStr = "Install"
50060    If IsPrinterPortInstalled(Portname) = False Then
50070      reg.hkey = HKEY_LOCAL_MACHINE
50080      reg.KeyRoot = "System\CurrentControlSet\Control\Print\Monitors"
50090      reg.CreateKey Monitorname & "\Ports\" & Portname
50100      reg.KeyRoot = "System\CurrentControlSet\Control\Print\Monitors\" & Monitorname & "\Ports\" & Portname
50110      reg.SetRegistryValue "Arguments", "-PPDFCREATORPRINTER", REG_SZ
50120      reg.SetRegistryValue "Command", App.Path & "\pdfcreator.exe", REG_SZ
50130      i = 300
50140      reg.SetRegistryValue "Delay", i, REG_DWORD
50150      reg.SetRegistryValue "Description", "Redirected Port", REG_SZ
50160      i = 0
50170      reg.SetRegistryValue "LogFileDebug", i, REG_DWORD
50180      reg.SetRegistryValue "LogFileName", vbNullString, REG_SZ
50190      reg.SetRegistryValue "LogFileUse", i, REG_DWORD
50200      reg.SetRegistryValue "Output", i, REG_DWORD
50210      reg.SetRegistryValue "Printer", "PDFCreator", REG_SZ
50220      reg.SetRegistryValue "PrintError", i, REG_DWORD
50230      reg.SetRegistryValue "RunUser", i, REG_DWORD
50240      i = 0
50250      reg.SetRegistryValue "ShowWindow", i, REG_DWORD
50260      reg.KeyRoot = vbNullString
50270      reg.CreateKey "System\CurrentControlSet\Control\Print\Ports\" & Portname
50280     Else
50290      WriteLogfile "PrinterPort [" & tStr & "]: Printerport is already installed."
50300      Exit Sub
50310    End If
50320    res = 1
50330   Case 1: ' UnInstall
50340    tStr = "UnInstall"
50350    If IsPrinterPortInstalled(Portname) = False Then
50360      res = DeletePort(vbNullString, 0, Portname & vbNullString)
50370     Else
50380      WriteLogfile "PrinterPort [" & tStr & "]: Printerport is not installed."
50390      Exit Sub
50400    End If
50410 '   reg.hkey = HKEY_LOCAL_MACHINE
50420 '   reg.DeleteKey "System\CurrentControlSet\Control\Print\Monitors\" & Monitorname & "\Ports\" & Portname
50430 '   reg.DeleteKey "System\CurrentControlSet\Control\Print\Ports\" & Portname
50440  End Select
50450  Set reg = Nothing
50460  If res = 0 Then
50470    WriteLogfile "PrinterPort [" & tStr & "]: Error -> " & RaiseAPIError
50480   Else
50490    WriteLogfile "PrinterPort [" & tStr & "]: Success"
50500  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "PrinterPort")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Sub PrinterDriver(InstallTyp As eInstall)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim DI3 As DRIVER_INFO_3, Driverpath As String, res As Long
50020  Dim lngDriverDirectoryLevel    As Long
50030  Dim lngDriverDirectoryNeeded   As Long
50040  Dim bytDriverDirectoryBuffer() As Byte
50050  Dim strDriverDirectory         As String * 512
50060  Dim lngWin32apiResultCode      As Long
50070  Dim tStr As String
50080  Select Case InstallTyp
  Case 0: ' Install
50100    tStr = "Install"
50110    If IsPrinterDriverInstalled(Drivername) = False Then
50120      lngDriverDirectoryLevel = 1
50130      lngWin32apiResultCode = GetPrinterDriverDirectory(vbNullString, vbNullString, lngDriverDirectoryLevel, ByVal vbNullString, 0, lngDriverDirectoryNeeded)
50140      ReDim bytDriverDirectoryBuffer(lngDriverDirectoryNeeded - 1)
50150      lngWin32apiResultCode = GetPrinterDriverDirectory(vbNullString, vbNullString, lngDriverDirectoryLevel, bytDriverDirectoryBuffer(0), lngDriverDirectoryNeeded, lngDriverDirectoryNeeded)
50160      lngWin32apiResultCode = lstrcpy(ByVal strDriverDirectory, bytDriverDirectoryBuffer(0))
50170
50180      Driverpath = Left(strDriverDirectory, InStr(strDriverDirectory, vbNullChar) - 1)
50190      If Right(Driverpath, 1) <> "\" Then Driverpath = Driverpath & "\"
50200
50210      With DI3
50220       .pName = Drivername & vbNullString
50230       If IsWin9xMe = True Then
50240         .cVersion = 0
50250         .pConfigFile = "PSCRIPT.DRV" & vbNullString
50260         .pDataFile = "APLWCSB1.SPD" & vbNullString
50270         .pDriverPath = "PSCRIPT.DRV" & vbNullString
50280         .pEnvironment = "Windows 4.0" & vbNullString
50290         .pHelpFile = "PSCRIPT.HLP" & vbNullString
50300         .pDependentFiles = "PSCRIPT.DRV" & Chr$(0) & "PSCRIPT.HLP" & Chr$(0) & Chr$(0)
50310         .pMonitorName = Monitorname & vbNullString
50320        Else
50330         .cVersion = 1
50340         .pConfigFile = Driverpath & "PSCRPTUI.DLL" & vbNullString
50350         .pDataFile = Driverpath & "ADIST5.PPD" & vbNullString
50360         .pDriverPath = Driverpath & "PSCRIPT.DLL" & vbNullString
50370         .pEnvironment = vbNullString
50380         .pHelpFile = Driverpath & "PSCRIPT.HLP" & vbNullString
50390         .pDependentFiles = Driverpath & "PSCRIPT.DLL" & vbNullString & vbNullString
50400         .pMonitorName = vbNullString
50410       End If
50420       .pDefaultDataType = "RAW" & vbNullString
50430      End With
50440
50450      res = AddPrinterDriver(vbNullString, 3, DI3)
50460     Else
50470      WriteLogfile "PrinterDriver [" & tStr & "]: Printerdriver is already installed."
50480      Exit Sub
50490    End If
50500   Case 1: ' UnInstall
50510    tStr = "UnInstall"
50520    If IsPrinterDriverInstalled(Drivername) = True Then
50530      res = DeletePrinterDriver(vbNullString, vbNullString, Drivername & vbNullString)
50540     Else
50550      WriteLogfile "PrinterDriver [" & tStr & "]: Printerdriver is not installed."
50560      Exit Sub
50570    End If
50580  End Select
50590  If res = 0 Then
50600    WriteLogfile "PrinterDriver [" & tStr & "]: Error -> " & RaiseAPIError
50610   Else
50620    WriteLogfile "PrinterDriver [" & tStr & "]: Success"
50630  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "PrinterDriver")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
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
      ini.FileName = GetWindowsDirectory & "\win.ini"
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
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tColl As Collection, i As Long
50020  Set tColl = GetAvailableMonitors
50030  IsPrinterMonitorInstalled = False
50040  For i = 1 To tColl.Count
50050   If UCase$(PrinterDriver) = UCase$(tColl.item(i)) Then
50060    IsPrinterMonitorInstalled = True
50070    Exit For
50080   End If
50090  Next i
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "IsPrinterMonitorInstalled")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function IsPrinterPortInstalled(PrinterDriver As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tColl As Collection, i As Long
50020  Set tColl = GetAvailablePorts
50030  IsPrinterPortInstalled = False
50040  For i = 1 To tColl.Count
50050   If UCase$(PrinterDriver) = UCase$(tColl.item(i)) Then
50060    IsPrinterPortInstalled = True
50070    Exit For
50080   End If
50090  Next i
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "IsPrinterPortInstalled")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function IsPrinterDriverInstalled(PrinterDriver As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tColl As Collection, i As Long
50020  Set tColl = GetAvailablePrinterdrivers
50030  IsPrinterDriverInstalled = False
50040  For i = 1 To tColl.Count
50050   If UCase$(PrinterDriver) = UCase$(tColl.item(i)) Then
50060    IsPrinterDriverInstalled = True
50070    Exit For
50080   End If
50090  Next i
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "IsPrinterDriverInstalled")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function IsPrinterInstalled(PrinterName As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim tColl As Collection, i As Long
50020  Set tColl = GetAvailablePrinters
50030  IsPrinterInstalled = False
50040  For i = 1 To tColl.Count
50050   If UCase$(PrinterName) = UCase$(tColl.item(i)) Then
50060    IsPrinterInstalled = True
50070    Exit For
50080   End If
50090  Next i
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "IsPrinterInstalled")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function RaiseAPIError() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ErrorMsg As String, ErrNum As Long
50020  ErrNum = Err.LastDllError
50030  ErrorMsg = String(256, 0)
50040  ErrorMsg = Left$(ErrorMsg, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, ErrNum, 0&, ErrorMsg, Len(ErrorMsg), ByVal 0))
50050  If Mid(ErrorMsg, Len(ErrorMsg) - 1) = vbCrLf Then
50060   ErrorMsg = Mid(ErrorMsg, 1, Len(ErrorMsg) - 2)
50070  End If
50080  RaiseAPIError = ErrorMsg
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "RaiseAPIError")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
