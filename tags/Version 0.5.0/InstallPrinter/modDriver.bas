Attribute VB_Name = "modDriver"
Option Explicit

Public Declare Function AddPrinterDriver Lib "winspool.drv" Alias "AddPrinterDriverA" (ByVal pName As String, ByVal Level As Long, pDriverInfo As Any) As Long
Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Public Declare Function GetPrinterDriverDirectory Lib "winspool.drv" Alias "GetPrinterDriverDirectoryA" (ByVal pName As String, ByVal pEnvironment As String, ByVal Level As Long, pDriverDirectory As Any, ByVal cdBuf As Long, pcbNeeded As Long) As Long
Public Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (lpString1 As Any, lpString2 As Any) As Long
Public Declare Function DeletePrinterDriver Lib "winspool.drv" Alias "DeletePrinterDriverA" (ByVal pName As String, ByVal pEnvironment As String, ByVal pDriverName As String) As Long

Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF

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
    .pHelpFile = vbNullString 'DriverPath & "PSCRIPT.HLP"
    .pDependentFiles = DriverPath & "ADIST5.BPD" & vbNullString & DriverPath & "PSCRIPT.DLL" & vbNullString & DriverPath & "HP4500.PPD" & vbNullString & DriverPath & "PSCRPTUI.DLL" & vbNullString & DriverPath & "PSCRIPT.HLP" & vbNullString & vbNullString
    .pMonitorName = vbNullString '"PJL Language Monitor"
    .pDefaultDataType = vbNullString
  End With

    Dim Path As String, cNeeded

    Path = String(256, 0)

  result = AddPrinterDriver(vbNullString, 3, DI3)

MsgBox result
RaiseAPIError
End Function

Public Function RaiseAPIError()
Dim ErrorMsg As String, ErrNum As Long

    ErrNum = Err.LastDllError
    
    ErrorMsg = String(256, 0)
    ErrorMsg = Left$(ErrorMsg, FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, ErrNum, 0&, ErrorMsg, Len(ErrorMsg), ByVal 0))
    
    MsgBox ErrorMsg

End Function
