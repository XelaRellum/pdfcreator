Attribute VB_Name = "modPrinter"
Option Explicit

Const Win9xDLLName = "pdfcmn95.dll", WinNTDLLName = "pdfcmnnt.dll", _
 Win9xEnvironment = "Windows 4.0", WinNTEnvironment = "Windows NT x86", _
 PrintSystem = "windows"

Public Type PRINTER_INFO_2a
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
 status As Long
 cJobs As Long
 AveragePPM As Long
End Type

Public DLLName As String, Environment As String

Public Sub InstallWindowsPrinter(Monitorname As String, Portname As String, Drivername As String, _
 Printername As String, LogFile As String, AppDir As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim RedmonCommand As String, reg As clsRegistry, value As String, orgValue As String
50020  WriteToLog "Start: Installation printer """ & Printername & """", LogFile, True
50030  RedmonCommand = CompletePath(AppDir) & "PDFSpooler.exe"
50040  If IsWin9xMe Then
50050    Environment = Win9xEnvironment
50060    DLLName = Win9xDLLName
50070   Else
50080    Environment = WinNTEnvironment
50090    DLLName = WinNTDLLName
50100  End If
50110  GetPorts
50120  If PortIsInstalled(Portname) = False Then
50130   InstallPort Portname, Monitorname, RedmonCommand, Printername
50140  End If
50150  GetMonitors
50160  GetPorts
50170  If MonitorIsInstalled(Monitorname) = False Then
50180   InstallMonitor Monitorname, Environment, DLLName, LogFile
50190  End If
50200  GetMonitors
50210  GetDrivers "Windows 4.0"
50220  GetDrivers "Windows NT x86"
50230  If DriverIsInstalled(Drivername) = False Then
50240   InstallDriver Drivername, LogFile
50250  End If
50260  GetPrinters
50270  If PrinterIsInstalled(Printername) = False Then
50280   InstallPrinter Printername, Drivername, Portname, LogFile
50290  End If
50300  GetPrinters
50310  If IsWinVista Then
50320   Set reg = New clsRegistry
50330   With reg
50340    .hkey = HKEY_LOCAL_MACHINE
50350    .KeyRoot = "SYSTEM\CurrentControlSet\Services"
50360    .SubKey = "Spooler"
50370    orgValue = .GetRegistryValue("RequiredPrivileges")
50380    value = orgValue
50390    If InStr(1, Replace(value, vbNullChar, " "), "SeBackupPrivilege", vbTextCompare) = 0 Then
50400     If Asc(Mid(value, Len(value), 1)) <> 0 Then
50410      value = value & vbNullChar
50420     End If
50430     value = value & "SeBackupPrivilege" & vbNullChar
50440    End If
50450    If InStr(1, Replace(value, vbNullChar, " "), "SeRestorePrivilege", vbTextCompare) = 0 Then
50460     If Asc(Mid(value, Len(value), 1)) <> 0 Then
50470      value = value & vbNullChar
50480     End If
50490     value = value & "SeRestorePrivilege" & vbNullChar
50500    End If
50510    If orgValue <> value Then
50520     Call .SetRegistryValue("RequiredPrivileges", value, REG_MULTI_SZ)
50530     'Debug.Print StopService("Spooler")
50540     'Debug.Print StartService("Spooler")
50550    End If
50560   End With
50570  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "InstallWindowsPrinter")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Sub UnInstallWindowsPrinter(Monitorname As String, Portname As String, Drivername As String, Printername As String, LogFile As String, Optional OnlyPrinter As Boolean = False)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  WriteToLog "Start: Uninstallation printer ""PDFCreator""", LogFile, True
50020  UnInstallPrinter Printername, LogFile
50030  GetPrinters
50040  If OnlyPrinter Then
50050   Exit Sub
50060  End If
50070  UnInstallDriver Drivername, "", LogFile
50080  GetDrivers
50090  UnInstallPort Portname, LogFile
50100  GetPorts
50110  UnInstallMonitor Monitorname, LogFile
50120  GetMonitors
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "UnInstallWindowsPrinter")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function InstallMonitor(Monitorname As String, Environment As String, DLLName As String, LogFile As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim mi As MONITOR_INFO_2
50020  With mi
50030   .pName = Monitorname
50040   .pEnvironment = Environment
50050   .pDLLName = DLLName
50060  End With
50070  If AddMonitor(vbNullString, 2, mi) = 0 Then
50080    WriteToLog "InstallMonitor: " & RaiseAPIError, LogFile
50090    InstallMonitor = False
50100   Else
50110    WriteToLog "InstallMonitor: success", LogFile, True
50120    InstallMonitor = True
50130  End If
50140  If IsWin9xMe = True Then
50150   Call SendMessage(65535, 26, 0, PrintSystem)
50160  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "InstallMonitor")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function UnInstallMonitor(Monitorname As String, LogFile As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If DeleteMonitor(vbNullString, vbNullString, Monitorname) = 0 Then
50020    WriteToLog "UnInstallMonitor: " & RaiseAPIError, LogFile
50030    UnInstallMonitor = False
50040   Else
50050    WriteToLog "UnInstallMonitor: success", LogFile
50060    UnInstallMonitor = True
50070  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "UnInstallMonitor")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Sub InstallPort(Portname As String, Monitorname As String, RedmonCommand As String, Printername As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010 ' Über die Registry, weil die Api einen Benutzereingriff erfordert
50020  Dim reg As clsRegistry
50030  Set reg = New clsRegistry
50040  With reg
50050   .hkey = HKEY_LOCAL_MACHINE
50060   .KeyRoot = "System\CurrentControlSet\Control\Print\Monitors"
50070   .SubKey = Monitorname & "\Ports\" & Portname
50080   If .KeyExists = False Then
50090    .CreateKey
50100   End If
50110   .SetRegistryValue "Arguments", "-PPDFCREATORPRINTER", REG_SZ
50120   .SetRegistryValue "Command", RedmonCommand, REG_SZ
50130   .SetRegistryValue "Delay", 300, REG_DWORD
50140   .SetRegistryValue "Description", "PDFCreator Redirected Port", REG_SZ
50150   .SetRegistryValue "LogFileDebug", 0, REG_DWORD
50160   .SetRegistryValue "LogFileUse", 0, REG_DWORD
50170   .SetRegistryValue "Output", 0, REG_DWORD
50180   .SetRegistryValue "Printer", Printername, REG_SZ
50190   .SetRegistryValue "PrintError", 0, REG_DWORD
50200   .SetRegistryValue "RunUser", 0, REG_DWORD
50210   .SetRegistryValue "ShowWindow", 0, REG_DWORD
50220  End With
50230  Set reg = Nothing
50240  If IsWin9xMe = True Then
50250   Call SendMessage(65535, 26, 0, PrintSystem)
50260  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "InstallPort")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function UnInstallPort(Portname As String, LogFile As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If DeletePort(vbNullString, 0, Portname) = 0 Then
50020    WriteToLog "UnInstallPort: " & RaiseAPIError, LogFile
50030    UnInstallPort = False
50040   Else
50050    WriteToLog "UnInstallPort: success", LogFile
50060    UnInstallPort = True
50070  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "UnInstallPort")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function InstallDriver(Drivername As String, LogFile As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim di As DRIVER_INFO_3
50020  With di
50030   .pName = Drivername
50040   .pDefaultDataType = "RAW"
50050   .pMonitorName = ""
50060   If IsWin9xMe = True Then
50070    .cVersion = 0
50080    .pDependentFiles = "ADOBEPS4.HLP" & vbNullString & _
                    "ICONLIB.DLL" & vbNullString & _
                    "PSMON.DLL" & vbNullString & _
                    "ADFONTS.MFM" & vbNullString & _
                    "ADOBEPS4.HLP" & vbNullString & _
                    "ADOBEPS4.DRV" & vbNullString & _
                    "ADIST5.PPD" & vbNullString & vbNullString
50150    .pConfigFile = "ADOBEPS4.DRV"
50160    .pDriverPath = "ADOBEPS4.DRV"
50170    .pEnvironment = Win9xEnvironment
50180    .pHelpFile = "ADOBEPS4.HLP"
50190    .pDataFile = "ADIST5.PPD"
50200 ' ???
50210 '   .cVersion = 3474436
50220   End If
50230   If IsWinNT4 = True Then
50240    .cVersion = 2
50250    .pDependentFiles = "PDFCREAT.PPD" & vbNullString & _
                    "ADOBEPS5.DLL" & vbNullString & _
                    "ADOBEPSU.DLL" & vbNullString & _
                    "ADOBEPS5.NTF" & vbNullString & _
                    "ADOBEPSU.HLP" & vbNullString & vbNullString
50300    .pConfigFile = "ADOBEPSU.DLL"
50310    .pDriverPath = "ADOBEPS5.DLL"
50320    .pEnvironment = WinNTEnvironment
50330    .pHelpFile = "ADOBEPSU.HLP"
50340    .pDataFile = "PDFCREAT.PPD"
50350   End If
50360   If IsWin2000 = True Then
50370    .cVersion = 3
50380    .pDependentFiles = "PSCRIPT.NTF" & vbNullString & vbNullString
50390    .pConfigFile = "PS5UI.DLL"
50400    .pDriverPath = "PSCRIPT5.DLL"
50410    .pEnvironment = WinNTEnvironment
50420    .pHelpFile = "PSCRIPT.HLP"
50430    .pDataFile = "PDFCREAT.PPD"
50440   End If
50450   If IsWinXPPlus = True Or IsWinVista = True Then 'WinXP and above
50460    .cVersion = 3
50470    .pDependentFiles = "PSCRIPT.NTF" & vbNullString & vbNullString
50480    .pConfigFile = "PS5UI.DLL"
50490    .pDriverPath = "PSCRIPT5.DLL"
50500    .pEnvironment = WinNTEnvironment
50510    .pHelpFile = "PSCRIPT.HLP"
50520    .pDataFile = "PDFCREAT.PPD"
50530   End If
50540   If AddPrinterDriver("", .cVersion, di) = 0 Then
50550     WriteToLog "InstallDriver: " & RaiseAPIError, LogFile
50560     InstallDriver = False
50570    Else
50580     WriteToLog "InstallDriver: success", LogFile
50590     InstallDriver = True
50600   End If
50610  End With
50620  If IsWin9xMe = True Then
50630   Call SendMessage(65535, 26, 0, PrintSystem)
50640  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "InstallDriver")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function UnInstallDriver(Drivername As String, Environment, LogFile As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  If DeletePrinterDriver(vbNullString, Environment, Drivername) = 0 Then
50020    WriteToLog "UnInstallDriver: " & RaiseAPIError, LogFile
50030    UnInstallDriver = False
50040   Else
50050    WriteToLog "UnInstallDriver: success", LogFile
50060    UnInstallDriver = True
50070  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "UnInstallDriver")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function InstallPrinter(Printername As String, Drivername As String, Portname As String, LogFile As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim pi As PRINTER_INFO_2a, sServer As String
50020  With pi
50030 '  .pPrinterName = StrPtr(Printername & vbNullString)
50040 '  .pDriverName = StrPtr(Drivername & vbNullString)
50050 '  .pPrintProcessor = StrPtr("WinPrint" & vbNullString)
50060 '  .pPortName = StrPtr(Portname & vbNullString)
50070 '  .pComment = StrPtr("eDoc Printer" & vbNullString)
50080 '  .pShareName = StrPtr(Printername & vbNullString)
50090 '  .pDatatype = StrPtr("RAW" & vbNullString)
50100
50110   .pPrinterName = Printername
50120   .pDriverName = Drivername
50130   .pPrintProcessor = "WinPrint"
50140   .pPortName = Portname
50150 '  .pPortName = "LPT1:"
50160   .pComment = "eDoc Printer"
50170   .pShareName = Printername
50180   .pDatatype = "RAW"
50190
50200   .Priority = 1
50210   .DefaultPriority = 1
50220   If GetPrinters.Count = 0 Then
50230     .Attributes = &H4 ' Set as defaultprinter
50240    Else
50250     .Attributes = &H0
50260   End If
50270  End With
50280  LogFile = Trim$(LogFile)
50290  If AddPrinter(vbNullString, 2, pi) = 0 Then
50300    If LenB(LogFile) > 0 Then
50310     WriteToLog "InstallPrinter [" & Printername & "]: " & RaiseAPIError, LogFile
50320    End If
50330    IfLoggingWriteLogfile "InstallPrinter [" & Printername & "]: " & RaiseAPIError
50340    InstallPrinter = False
50350   Else
50360    If LenB(LogFile) > 0 Then
50370     WriteToLog "InstallPrinter: success", LogFile
50380    End If
50390    IfLoggingWriteLogfile "InstallPrinter: success"
50400    InstallPrinter = True
50410  End If
50420  If IsWin9xMe = True Then
50430   Call SendMessage(65535, 26, 0, PrintSystem)
50440  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "InstallPrinter")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function UnInstallPrinter(Printername As String, LogFile As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim pd As PRINTER_DEFAULTS, pHandle As Long
50020  With pd
50030   .pDatatype = 0
50040   .pDevMode = 0
50050   .DesiredAccess = PRINTER_ALL_ACCESS
50060  End With
50070  LogFile = Trim$(LogFile)
50080  If OpenPrinter(Printername, pHandle, pd) <> 0 Then
50090    If DeletePrinter(pHandle) <> 0 Then
50100      If ClosePrinter(pHandle) <> 0 Then
50110        If LenB(LogFile) > 0 Then
50120         WriteToLog "UnInstallPrinter: success", LogFile
50130        End If
50140        IfLoggingWriteLogfile "UnInstallPrinter: success"
50150        UnInstallPrinter = True
50160       Else
50170        If LenB(LogFile) > 0 Then
50180         WriteToLog "UnInstallPrinter: " & RaiseAPIError, LogFile
50190        End If
50200        IfLoggingWriteLogfile "UnInstallPrinter: " & RaiseAPIError
50210        UnInstallPrinter = False
50220      End If
50230     Else
50240      If LenB(LogFile) > 0 Then
50250       WriteToLog "UnInstallPrinter: " & RaiseAPIError, LogFile
50260      End If
50270      IfLoggingWriteLogfile "UnInstallPrinter: " & RaiseAPIError
50280      UnInstallPrinter = False
50290    End If
50300   Else
50310    If LenB(LogFile) > 0 Then
50320     WriteToLog "UnInstallPrinter: " & RaiseAPIError, LogFile
50330    End If
50340    IfLoggingWriteLogfile "UnInstallPrinter: " & RaiseAPIError
50350    UnInstallPrinter = False
50360  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "UnInstallPrinter")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function GetMonitors() As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim cdBuf As Long, pcbNeeded As Long, pcReturned As Long, _
  res As Long, tL() As Long, i As Long
50030  Set GetMonitors = New Collection
50040  ReDim tL(0)
50050  res = EnumMonitors("", 1, tL(0), cdBuf, pcbNeeded, pcReturned)
50060  If pcbNeeded > 0 Then
50070   cdBuf = pcbNeeded
50080   ReDim tL(pcbNeeded / Len(tL(0)))
50090   res = EnumMonitors("", 1, tL(0), cdBuf, pcbNeeded, pcReturned)
50100   For i = 0 To pcReturned - 1
50110    GetMonitors.Add GetStrFromPtrA(tL(i))
50120   Next i
50130  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "GetMonitors")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function GetPorts() As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim cdBuf As Long, pcbNeeded As Long, pcReturned As Long, _
  res As Long, tL() As Long, i As Long
50030  Set GetPorts = New Collection
50040  ReDim tL(0)
50050  res = EnumPorts("", 1, tL(0), cdBuf, pcbNeeded, pcReturned)
50060  If pcbNeeded > 0 Then
50070   cdBuf = pcbNeeded
50080   ReDim tL(pcbNeeded / Len(tL(0)))
50090   res = EnumPorts("", 1, tL(0), cdBuf, pcbNeeded, pcReturned)
50100   For i = 0 To pcReturned - 1
50110    GetPorts.Add GetStrFromPtrA(tL(i))
50120   Next i
50130  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "GetPorts")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function GetDrivers(Optional Environment As String = "") As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim cdBuf As Long, pcbNeeded As Long, pcReturned As Long, _
  res As Long, tL() As Long, i As Long
50030  Set GetDrivers = New Collection
50040  ReDim tL(0)
50050  res = EnumPrinterDrivers("", Environment, 1, tL(0), cdBuf, pcbNeeded, pcReturned)
50060  If pcbNeeded > 0 Then
50070   cdBuf = pcbNeeded
50080   ReDim tL(pcbNeeded / Len(tL(0)))
50090   res = EnumPrinterDrivers("", Environment, 1, tL(0), cdBuf, pcbNeeded, pcReturned)
50100   For i = 0 To pcReturned - 1
50110    GetDrivers.Add GetStrFromPtrA(tL(i))
50120   Next i
50130  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "GetDrivers")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function GetPrinters() As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim cdBuf As Long, pcbNeeded As Long, pcReturned As Long, _
  res As Long, tL() As Long, i As Long, PRINTER_LEVEL As Long
50030  Set GetPrinters = New Collection
50040  ReDim tL(0)
50050  PRINTER_LEVEL = 2
50060  res = EnumPrinters(PRINTER_ENUM_LOCAL, "", PRINTER_LEVEL, tL(0), cdBuf, pcbNeeded, pcReturned)
50070  If pcbNeeded > 0 Then
50080   cdBuf = pcbNeeded
50090   ReDim tL(pcbNeeded / Len(tL(0)))
50100   res = EnumPrinters(PRINTER_ENUM_LOCAL, "", PRINTER_LEVEL, tL(0), cdBuf, pcbNeeded, pcReturned)
50110   For i = 0 To pcReturned - 1
50120 '   Debug.Print GetStrFromPtrA(tL(i * 21 + 1)) & "   " & GetStrFromPtrA(tL(i * 21 + 3))
50130    GetPrinters.Add GetStrFromPtrA(tL(i * 21 + 1)) & Chr$(0) & GetStrFromPtrA(tL(i * 21 + 3))
50140   Next i
50150  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "GetPrinters")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub WriteToLog(Str1 As String, LogFile As String, Optional CreateNewFile As Boolean = False)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim fn As Long
50020  If LenB(LogFile) > 0 Then
50030   fn = FreeFile
50040   If FileExists(LogFile) = False Or CreateNewFile = True Then
50050     Open LogFile For Output As fn
50060     Print #fn, "Windowsversion: " & GetWinVersionStr & vbCrLf & _
     "PDFCreator-Revision: " & GetProgramReleaseStr & vbCrLf
50080    Else
50090     Open LogFile For Append As fn
50100     If FileLen(LogFile) = 0 Then
50110      Print #fn, "Windowsversion: " & GetWinVersionStr & vbCrLf & _
      "PDFCreator-Revision: " & GetProgramReleaseStr & vbCrLf
50130     End If
50140   End If
50150   Print #fn, Str1
50160   Close fn
50170  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "WriteToLog")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Private Function PortIsInstalled(Portname As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim coll As Collection, i As Long
50020  Set coll = GetPorts
50030  PortIsInstalled = False
50040  For i = 1 To coll.Count
50050   If UCase$(Portname) = UCase$(coll.Item(i)) Then
50060    PortIsInstalled = True
50070    Exit For
50080   End If
50090  Next i
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "PortIsInstalled")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function MonitorIsInstalled(Monitorname As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim coll As Collection, i As Long
50020  Set coll = GetMonitors
50030  MonitorIsInstalled = False
50040  For i = 1 To coll.Count
50050   If UCase$(Monitorname) = UCase$(coll.Item(i)) Then
50060    MonitorIsInstalled = True
50070    Exit For
50080   End If
50090  Next i
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "MonitorIsInstalled")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Private Function DriverIsInstalled(Drivername As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim coll As Collection, i As Long
50020  Set coll = GetDrivers
50030  DriverIsInstalled = False
50040  For i = 1 To coll.Count
50050   If UCase$(Drivername) = UCase$(coll.Item(i)) Then
50060    DriverIsInstalled = True
50070    Exit For
50080   End If
50090  Next i
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "DriverIsInstalled")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function PrinterIsInstalled(Printername As String) As Boolean
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Printers As Collection, i As Long, pf() As String
50020  Set Printers = GetPrinters
50030  PrinterIsInstalled = False
50040  For i = 1 To Printers.Count
50050   pf = Split(Printers(i), Chr$(0))
50060   If UCase$(Printername) = UCase$(pf(0)) Then
50070    PrinterIsInstalled = True
50080    Exit For
50090   End If
50100  Next i
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "PrinterIsInstalled")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetFreePDFCreatorPort() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim ports As Collection, tStr As String, foundFreePort As Boolean, _
  i As Long, j As Long
50030  Set ports = GetPorts
50040  If ports.Count = 0 Then
50050   GetFreePDFCreatorPort = "PDFCreator001:"
50060   Exit Function
50070  End If
50080  For i = 1 To 999
50090   tStr = "PDFCreator" & Format$(i, "000") & ":"
50100   foundFreePort = False
50110   For j = 1 To ports.Count
50120    If UCase$(tStr) <> UCase$(ports(j)) Then
50130     foundFreePort = True
50140     Exit For
50150    End If
50160   Next j
50170   If foundFreePort = True Then
50180    Exit For
50190   End If
50200  Next i
50210  If foundFreePort = True Then
50220    GetFreePDFCreatorPort = tStr
50230   Else
50240    MsgBox "Cannot find a free printer port!", vbExclamation
50250  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "GetFreePDFCreatorPort")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Sub InstallAdditionalWindowsPrinter(Printername As String, LogFile As String, AppDir As String)
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim i As Long, c As Collection
50020  If PrinterIsInstalled(Printername) = True Then
50030   MsgBox "Printer '" & Printername & "' is already installed!", vbExclamation
50040   Exit Sub
50050  End If
50060  InstallWindowsPrinter "PDFCreator", GetFreePDFCreatorPort, "PDFCreator", Printername, LogFile, AppDir
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Sub
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "InstallAdditionalWindowsPrinter")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Sub
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Sub

Public Function GetPDFCreatorPrintername() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim PDFCreatorPrinters As Collection, InstalledPDFCreatorPrinter As String, i As Long
50020  InstalledPDFCreatorPrinter = LCase$(GetInstalledPDFCreatorPrinter)
50030  Set PDFCreatorPrinters = GetPDFCreatorPrinters
50040  For i = 1 To PDFCreatorPrinters.Count
50050   If LCase$(PDFCreatorPrinters(i)) = InstalledPDFCreatorPrinter Then
50060    GetPDFCreatorPrintername = PDFCreatorPrinters(i)
50070    Exit Function
50080   End If
50090  Next i
50100  If PDFCreatorPrinters.Count > 0 Then
50110   GetPDFCreatorPrintername = PDFCreatorPrinters(1)
50120  End If
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "GetPDFCreatorPrintername")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetInstalledPDFCreatorPrinter() As String
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim reg As clsRegistry
50020  Set reg = New clsRegistry
50030  reg.hkey = HKEY_LOCAL_MACHINE
50040  reg.KeyRoot = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Uninstall_GUID
50050  GetInstalledPDFCreatorPrinter = reg.GetRegistryValue("Printername")
50060  Set reg = Nothing
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "GetInstalledPDFCreatorPrinter")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function

Public Function GetPDFCreatorPrinters() As Collection
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
On Error GoTo ErrPtnr_OnError
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
50010  Dim Printers1 As Collection, PDFCreatorPrinters As Collection, _
  reg As clsRegistry, SubKeys As Collection, i As Long, j As Long
50030  Set GetPDFCreatorPrinters = New Collection
50040  Set Printers1 = GetAvailablePrinters2
50050  Set PDFCreatorPrinters = New Collection
50060  Set reg = New clsRegistry
50070  Set SubKeys = reg.EnumRegistryKeys(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Print\Monitors\PDFCreator\Ports")
50080  For i = 1 To Printers1.Count
50090   For j = 1 To SubKeys.Count
50100    If SubKeys(j) = Printers1(i)(1) Then
50110     AddSortedStr PDFCreatorPrinters, CStr(Printers1(i)(0))
50120    End If
50130   Next j
50140  Next i
50150  Set GetPDFCreatorPrinters = PDFCreatorPrinters
'---ErrPtnr-OnError-START--- DO NOT MODIFY ! ---
Exit Function
ErrPtnr_OnError:
Select Case ErrPtnr.OnError("modPrinter", "GetPDFCreatorPrinters")
Case 0: Resume
Case 1: Resume Next
Case 2: Exit Function
Case 3: End
End Select
'---ErrPtnr-OnError-END--- DO NOT MODIFY ! ---
End Function
