Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function AddPortEx Lib "winspool.drv" Alias "AddPortExA" (ByVal pName As String, ByVal pLevel As Long, lpBuffer As Any, ByVal pMonitorName As String) As Long
Public Declare Function AddPort Lib "winspool.drv" Alias "AddPortA" (ByVal pName As String, ByVal hwnd As Long, ByVal pMonitorName As String) As Long
Public Declare Function AddPrinter Lib "winspool.drv" Alias "AddPrinterA" (ByVal pName As String, ByVal Level As Long, pPrinter As Any) As Long
Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function AddPrinterDriver Lib "winspool.drv" Alias "AddPrinterDriverA" (ByVal pName As String, ByVal Level As Long, pDriverInfo As Any) As Long
Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Public Declare Function GetPrinterDriverDirectory Lib "winspool.drv" Alias "GetPrinterDriverDirectoryA" (ByVal pName As String, ByVal pEnvironment As String, ByVal Level As Long, pDriverDirectory As Any, ByVal cdBuf As Long, pcbNeeded As Long) As Long
Public Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (lpString1 As Any, lpString2 As Any) As Long

Public Enum PortTypes
  PORT_TYPE_WRITE = &H1
  PORT_TYPE_READ = &H2
  PORT_TYPE_REDIRECTED = &H4
  PORT_TYPE_NET_ATTACHED = &H8
End Enum

Public Type PORT_INFO_2S
   pPortName As String
   pMonitorName As String
   pDescription As String
   fPortType As Long
   Reserved As Long
End Type

Public Type PRINTER_INFO_2
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
 
Public Type DRIVER_INFO_3
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
 

Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF



' http://www.vb2themax.com/Item.asp?PageID=CodeBank&ID=388
' Install a new printer on the system
'   sPrinterName is the name to assign to identify the printer
'   sDriver is a string identifying the driver for the printer
'   sPort is the printer's COM port
'   sComment is a comment to associate to the printer item
'
' Example:
'    Dim bOK As Boolean
'    bOK = InstallPrinter("Epson", "Epson Stylus COLOR 440", "LPT1:", ,
'  "My favourite printer")
'    MsgBox "Printer added: " & bOK

Public Function InstallPrinter(ByVal sPrinterName As String, ByVal sDriver As String, _
    Optional ByVal sPort As String = "LPT1:", Optional sServer As String, _
    Optional sComment As String) As Boolean
    
    Dim hPrinter As Long
    Dim PI As PRINTER_INFO_2
    
    ' fill the PRINTER_INFO_2 struct
    With PI
        .pPrinterName = sPrinterName
        .pDriverName = sDriver
        .pPortName = sPort
        .pServerName = sServer
        .pComment = sComment
        .pPrintProcessor = "WinPrint"
        .Priority = 1
        .DefaultPriority = 1
        .pDatatype = "RAW"
    End With
    
    ' add the printer
    hPrinter = AddPrinter(sServer, 2, PI)
    ' if successful close the printer and return True
    If hPrinter <> 0 Then
        ClosePrinter hPrinter
        InstallPrinter = True
    Else
        RaiseAPIError
    End If

End Function

Public Function AddDriver() As Boolean
  Dim DI3 As DRIVER_INFO_3
  Dim DriverPath As String
  Dim result As Long

  Dim lngDriverDirectoryLevel    As Long
  Dim lngDriverDirectoryNeeded   As Long
  Dim bytDriverDirectoryBuffer() As Byte
  Dim strDriverDirectory         As String * 512
  Dim lngWin32apiResultCode      As Long

  lngDriverDirectoryLevel = 1
  lngWin32apiResultCode = GetPrinterDriverDirectory(vbNullString, vbNullString, lngDriverDirectoryLevel, ByVal vbNullString, 0, lngDriverDirectoryNeeded)
  ReDim bytDriverDirectoryBuffer(lngDriverDirectoryNeeded - 1)
  lngWin32apiResultCode = GetPrinterDriverDirectory(vbNullString, vbNullString, lngDriverDirectoryLevel, bytDriverDirectoryBuffer(0), lngDriverDirectoryNeeded, lngDriverDirectoryNeeded)
  lngWin32apiResultCode = lstrcpy(ByVal strDriverDirectory, bytDriverDirectoryBuffer(0))

  DriverPath = Left(strDriverDirectory, InStr(strDriverDirectory, vbNullChar) - 1)
  If Right(DriverPath, 1) <> "\" Then DriverPath = DriverPath & "\"
  'DriverPath = "C:\WINNT\system32\spool\drivers\w32x86\"
  
  With DI3
    .cVersion = 1
    .pName = "PDFCreator"
    .pEnvironment = vbNullString
    .pDriverPath = DriverPath & "PSCRIPT.DLL" & vbNullString
    .pDataFile = DriverPath & "ADIST5.PPD" & vbNullString
    .pConfigFile = DriverPath & "PSCRPTUI.DLL" & vbNullString
    .pHelpFile = DriverPath & "PSCRIPT.HLP"
    .pDependentFiles = DriverPath & "ADIST5.BPD" & vbNullString & DriverPath & "PSCRIPT.DLL" & vbNullString & DriverPath & "HP4500.PPD" & vbNullString & DriverPath & "PSCRPTUI.DLL" & vbNullString & DriverPath & "PSCRIPT.HLP" & vbNullString & vbNullString
    .pMonitorName = vbNullString
    .pDefaultDataType = vbNullString
  End With

    Dim Path As String, cNeeded

    Path = String(256, 0)

  result = AddPrinterDriver(vbNullString, 3, DI3)

If result = 0 Then
  RaiseAPIError
Else
  AddDriver = True
  'MsgBox "Treiber erfolgreich installiert"
End If
End Function

Public Function AddMyPort() As Boolean
   Dim success As Boolean
   Dim Port2 As PORT_INFO_2S
   Dim isWindowsNT As Boolean

    With Port2
       .fPortType = &H1 Or &H2
       .pMonitorName = "Local Port"
       .pPortName = App.Path & "\spool.ps"
    End With
    
    isWindowsNT = Environ$("OS") <> ""
    
    If isWindowsNT Then
      success = AddPortEx(vbNullString, 1, Port2, "Local Port")
    Else
      success = AddPort(vbNullString, 0&, "Local Port")
    End If
        
    If success Then
        AddMyPort = True
        'MsgBox "AddPort war erfolgreich"
    Else
        RaiseAPIError
    End If
End Function

Public Function RaiseAPIError()
Dim ErrorMsg As String, ErrNum As Long

    ErrNum = Err.LastDllError
    
    ErrorMsg = String(256, 0)
    ErrorMsg = Left$(ErrorMsg, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, ErrNum, 0&, ErrorMsg, Len(ErrorMsg), ByVal 0))
    
    'MsgBox ErrorMsg

End Function
